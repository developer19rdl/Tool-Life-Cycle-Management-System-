Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class BladeMicrometer75
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = ""
    Private _category As String = "Instrument"
    Private _selectedMastersObservation As New List(Of MasterSelectorItem)
    Private _selectedMastersGeometric As New List(Of MasterSelectorItem)



    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _activeMinRange As Decimal = 0
    Private _activeMaxRange As Decimal = 25
    Private _isPopulating As Boolean = False

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    ' 30 Nominal Ranges for Blade Micrometer (up to 75mm)
    Private ReadOnly nominalRanges As Decimal() = {
        2.5, 5.1, 7.7, 10.3, 12.9, 15.0, 17.6, 20.2, 22.8, 25.0,
        27.5, 30.1, 32.7, 35.3, 37.9, 40.0, 42.6, 45.2, 47.8, 50.0,
        52.5, 55.1, 57.7, 60.3, 62.9, 65.0, 67.6, 70.2, 72.8, 75.0
    }

    ' Geometric Factors (Only Parallelism for Blade)
    Private ReadOnly geometricFactors As String() = {"Parallel"}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = "Blade Micrometer"
        Try
            LoadDynamicLC()
        Catch ex As Exception
        End Try

        ' Add handlers for all trial textboxes (Accuracy)
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                tb.IsEnabled = False
                AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                AddHandler tb.TextChanged, AddressOf TrialTextChanged
                AddHandler tb.LostFocus, AddressOf ObservationLostFocus
            Next
        Next

        ' Add handlers for Geometric trials
        For Each gf In geometricFactors
            Dim tbs = GetTextBoxesForGeometric(gf)
            For Each tb In tbs
                tb.IsEnabled = False ' Initially disabled until LC is selected
                AddHandler tb.TextChanged, AddressOf GeometricTextChanged
                AddHandler tb.KeyDown, AddressOf GeometricKeyDown
                AddHandler tb.LostFocus, AddressOf GeometricLostFocus
            Next
        Next

        ' Add handlers for Environment
        AddHandler TextBox30.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox28.TextChanged, AddressOf EnvironmentTextChanged

        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            LoadInstrumentDetails(TextBox1.Text.Trim())
        End If

        If Not IsEditMode Then
            TxtCalibrationDate.SelectedDate = DateTime.Today
        End If

        ' Add handlers for Duration
        AddHandler TextBox24.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.GotFocus, Sub() _isTimeOutManual = True
        AddHandler TextBox26.PreviewKeyDown, Sub() _isTimeOutManual = True

        ' Setup Real-time timer
        _timerRealTime = New System.Windows.Threading.DispatcherTimer()
        _timerRealTime.Interval = TimeSpan.FromSeconds(5)
        AddHandler _timerRealTime.Tick, AddressOf TimerRealTime_Tick
        _timerRealTime.Start()

        ' Initialize LC and disable fields
        ComboBox5.SelectedIndex = -1
        DisableAllObservationFields()
        ResetAllTrials()
        UpdateSizeLockState()

        If IsEditMode Then
            Try
                PopulateCalibrationData()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub TxtCalibrationDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles TxtCalibrationDate.SelectedDateChanged
        If IsEditMode Then Return
        If TxtCalibrationDate.SelectedDate.HasValue Then
            Dim selectedDateStr = TxtCalibrationDate.SelectedDate.Value.ToShortDateString()
            Dim env = EnvironmentCache.LoadEnvironment(Me.GetType().Name, selectedDateStr)
            If env IsNot Nothing Then
                If String.IsNullOrWhiteSpace(TextBox30.Text) Then TextBox30.Text = env.Temperature
                If String.IsNullOrWhiteSpace(TextBox28.Text) Then TextBox28.Text = env.Humidity
                If String.IsNullOrWhiteSpace(TextBox20.Text) Then TextBox20.Text = env.DepthError
                If String.IsNullOrWhiteSpace(TextBox24.Text) Then TextBox24.Text = env.TimeIn
            End If
        End If
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse String.IsNullOrWhiteSpace(TargetCycle) Then Return

        Try
            Dim dt = MySqlCls.GetBladeMicrometer75Data(TextBox1.Text.Trim(), TargetCycle)
            If dt.Rows.Count > 0 Then
                Dim dr = dt.Rows(0)

                _isPopulating = True

                ComboBox2.Text = dr("Location").ToString()
                TextBox21.Text = dr("Size").ToString()
                ComboBox5.Text = dr("LC").ToString()
                TextBoxColor.Text = dr("Color").ToString()
                TextBox30.Text = dr("Temperature").ToString()
                TextBox28.Text = dr("Humidity").ToString()
                TextBox29.Text = dr("TMU").ToString()
                TextBox20.Text = If(IsDBNull(dr("DepthError")) OrElse dr("DepthError").ToString() = "-", "", dr("DepthError").ToString())
                TextBox19.Text = dr("Remark").ToString()

                If Not IsDBNull(dr("Date")) Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(dr("Date"))
                TextBox24.Text = dr("TimeIn").ToString()
                TextBox26.Text = dr("TimeOut").ToString()
                TextBox25.Text = dr("TotalTime").ToString()

                ' Geometric Factors (Parallelism)
                Dim parallel = {dr("Parallelism_1"), dr("Parallelism_2"), dr("Parallelism_3")}
                PopulateGeometricFields("Parallel", parallel)

                ' Observations
                For Each r In nominalRanges
                    Dim nomStr = r.ToString("0.0").Replace(".", "_")
                    Dim tbs = GetTextBoxesForRange(r)
                    If tbs.Count >= 3 Then
                        tbs(0).Text = If(IsDBNull(dr($"Obs{nomStr}_1")), "", dr($"Obs{nomStr}_1").ToString())
                        tbs(1).Text = If(IsDBNull(dr($"Obs{nomStr}_2")), "", dr($"Obs{nomStr}_2").ToString())
                        tbs(2).Text = If(IsDBNull(dr($"Obs{nomStr}_3")), "", dr($"Obs{nomStr}_3").ToString())
                    End If
                Next

                ' Restore Masters from Cache
                Dim cached = CalibrationMasterCache.LoadMastersDual(TextBox1.Text.Trim(), TargetCycle)
                If cached.Observation.Count > 0 Then
                    _selectedMastersObservation = MySqlCls.GetMastersByDescriptions(cached.Observation)
                    UpdateMasterUIObservation()
                End If
                If cached.Geometric.Count > 0 Then
                    _selectedMastersGeometric = MySqlCls.GetMastersByDescriptions(cached.Geometric)
                    UpdateMasterUIGeometric()
                End If

                _isPopulating = False
                TriggerAllValidations()
                CalculateTMU()
            End If
        Catch ex As Exception
            Console.WriteLine("PopulateCalibrationData Error: " & ex.Message)
            _isPopulating = False
        End Try
    End Sub

    Private Sub PopulateGeometricFields(factor As String, vals As Object())
        Dim tbs = GetTextBoxesForGeometric(factor)
        For i As Integer = 0 To Math.Min(tbs.Count, vals.Length) - 1
            tbs(i).Text = If(IsDBNull(vals(i)), "", vals(i).ToString())
        Next
    End Sub

    Private Sub AdjustVisibilityBySize(sizeStr As String)
        If TextBox21 Is Nothing Then Return
        Try
            _activeMinRange = 0 : _activeMaxRange = 25
            If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
                Dim parts = sizeStr.Split("-"c)
                If parts.Length > 1 Then
                    Dim minPart = parts(0).ToLower().Replace("mm", "").Trim()
                    Dim maxPart = parts(1).ToLower().Replace("mm", "").Trim()
                    Decimal.TryParse(minPart, _activeMinRange)
                    Decimal.TryParse(maxPart, _activeMaxRange)
                End If
            End If

            ' Tab visibility
            If TabItem0_25 IsNot Nothing Then TabItem0_25.IsEnabled = (_activeMinRange <= 0 AndAlso _activeMaxRange >= 25)
            If TabItem25_50 IsNot Nothing Then TabItem25_50.IsEnabled = (_activeMinRange <= 25 AndAlso _activeMaxRange >= 50)
            If TabItem50_75 IsNot Nothing Then TabItem50_75.IsEnabled = (_activeMinRange <= 50 AndAlso _activeMaxRange >= 75)

            ' Switch to visible tab
            If TabControlObservations IsNot Nothing Then
                If TabItem0_25 IsNot Nothing AndAlso TabItem0_25.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem0_25
                ElseIf TabItem25_50 IsNot Nothing AndAlso TabItem25_50.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem25_50
                ElseIf TabItem50_75 IsNot Nothing AndAlso TabItem50_75.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem50_75
                End If
            End If

            For Each r In nominalRanges
                Dim inRange = IsNominalInRange(r)
                Dim tbs = GetTextBoxesForRange(r)
                For Each tb In tbs
                    If Not inRange Then
                        tb.IsEnabled = False : tb.Text = "" : tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                    Else
                        If ComboBox5.SelectedItem IsNot Nothing Then
                            tb.IsEnabled = True
                            If String.IsNullOrWhiteSpace(tb.Text) OrElse tb.Text = "-" Then
                                tb.Background = Brushes.White
                            End If
                        Else
                            tb.IsEnabled = False
                            tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                        End If
                    End If
                Next
            Next

            ' Geometric re-enabling
            If ComboBox5.SelectedItem IsNot Nothing Then
                For Each gf In geometricFactors
                    For Each tb In GetTextBoxesForGeometric(gf)
                        tb.IsEnabled = True
                        If String.IsNullOrWhiteSpace(tb.Text) OrElse tb.Text = "-" Then
                            tb.Background = Brushes.White
                        End If
                    Next
                Next
                Dim firstInRange = nominalRanges.FirstOrDefault(Function(r) IsNominalInRange(r))
                If firstInRange > 0 Then EnableRange(firstInRange)
            End If
            UpdateSizeLockState()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox21_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles TextBox21.SelectionChanged
        If Not _isPopulating Then AdjustVisibilityBySize(TextBox21.Text)
    End Sub

    Private Sub UpdateObservationLabel(nominal As Decimal)
        Try
            Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
            Dim lblName = "Label" & nomStr
            Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)
            If lbl Is Nothing Then Return

            Dim tbs = GetTextBoxesForRange(nominal)
            Dim isFilled = tbs.All(Function(t) Not String.IsNullOrWhiteSpace(t.Text))
            Dim isAnyFilled = tbs.Any(Function(t) Not String.IsNullOrWhiteSpace(t.Text))

            If isFilled Then
                lbl.Foreground = New SolidColorBrush(Color.FromRgb(100, 116, 139)) ' Gray
            ElseIf isAnyFilled OrElse (nominal = nominalRanges.FirstOrDefault(Function(r) IsNominalInRange(r) AndAlso GetTextBoxesForRange(r).Any(Function(t) String.IsNullOrWhiteSpace(t.Text)))) Then
                lbl.Foreground = Brushes.RoyalBlue
            Else
                lbl.Foreground = Brushes.Black
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UpdateGeometricLabel(gf As String)
        Try
            Dim lblName = "Label" & gf
            Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)
            If lbl Is Nothing Then Return

            Dim tbs = GetTextBoxesForGeometric(gf)
            Dim isFilled = tbs.All(Function(t) Not String.IsNullOrWhiteSpace(t.Text))

            If isFilled Then
                lbl.Foreground = New SolidColorBrush(Color.FromRgb(100, 116, 139))
            Else
                lbl.Foreground = Brushes.RoyalBlue
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UpdateOverallStatus()
        If _isPopulating OrElse ComboBox4 Is Nothing Then Return

        Dim hasNG = False

        ' Check Environment
        If IsNGColor(TextBox30.Background) OrElse IsNGColor(TextBox28.Background) Then hasNG = True

        ' Check Accuracy
        For Each r In nominalRanges
            If IsNominalInRange(r) Then
                If GetTextBoxesForRange(r).Any(Function(tb) IsNGColor(tb.Background)) Then
                    hasNG = True : Exit For
                End If
            End If
        Next

        ' Check Geometric
        If Not hasNG Then
            For Each gf In geometricFactors
                If GetTextBoxesForGeometric(gf).Any(Function(tb) IsNGColor(tb.Background)) Then
                    hasNG = True : Exit For
                End If
            Next
        End If

        ComboBox4.Text = If(hasNG, "NG", "OK")
    End Sub

    Private Function IsNGColor(brush As Brush) As Boolean
        If brush Is Nothing Then Return False
        If TypeOf brush Is SolidColorBrush Then
            Dim scb = DirectCast(brush, SolidColorBrush)
            Return scb.Color = Color.FromRgb(254, 226, 226) OrElse scb.Color = Colors.LightPink
        End If
        Return False
    End Function

    Private Sub EnvironmentTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim tb = DirectCast(sender, TextBox)
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            Dim isOK = True
            If tb.Name = "TextBox30" Then ' Temp 20 +/- 1
                If val < 19 OrElse val > 21 Then isOK = False
            ElseIf tb.Name = "TextBox28" Then ' Humidity 50 +/- 10
                If val < 40 OrElse val > 60 Then isOK = False
            End If

            If Not isOK Then
                tb.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' Light Pink
                tb.Foreground = Brushes.Red
            Else
                tb.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' Light Green
                tb.Foreground = Brushes.Black
            End If
        Else
            tb.Background = Brushes.White : tb.Foreground = Brushes.Black
        End If
        UpdateOverallStatus()
    End Sub

    Private Function IsNominalInRange(nominal As Decimal) As Boolean
        Return nominal >= _activeMinRange AndAlso nominal <= _activeMaxRange
    End Function

    Private Sub GeometricKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)
            Dim nameParts = tb.Name.Split("_"c) ' txtParallel_1
            If nameParts.Length = 2 Then
                Dim factor = nameParts(0).Replace("txt", "")
                Dim trialNum = Integer.Parse(nameParts(1))
                If trialNum < 3 Then
                    Dim nextTbName = "txt" & factor & "_" & (trialNum + 1)
                    Dim nextTb = DirectCast(Me.FindName(nextTbName), TextBox)
                    If nextTb IsNot Nothing Then nextTb.Focus()
                Else
                    ' End of Geometry, move to Observations
                    If ComboBox5.SelectedItem IsNot Nothing Then
                        Dim firstInRange = nominalRanges.FirstOrDefault(Function(r) IsNominalInRange(r))
                        If firstInRange > 0 Then
                            Dim obsTbs = GetTextBoxesForRange(firstInRange)
                            If obsTbs.Count > 0 Then obsTbs(0).Focus()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObservationKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)
            Dim nameParts = tb.Name.Split("_"c)
            If nameParts.Length = 3 Then
                Dim trialNum = Integer.Parse(nameParts(2))
                If trialNum < 3 Then
                    Dim nextTbName = nameParts(0) & "_" & nameParts(1) & "_" & (trialNum + 1)
                    Dim nextTb = DirectCast(Me.FindName(nextTbName), TextBox)
                    If nextTb IsNot Nothing Then nextTb.Focus()
                Else
                    Dim nominal = GetNominalFromTextBox(tb)
                    EnableNextRange(nominal)
                End If
            End If
        End If
    End Sub

    Private Sub ObservationLostFocus(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        If Not String.IsNullOrWhiteSpace(tb.Text) Then
            Dim val As Decimal
            If Decimal.TryParse(tb.Text, val) Then tb.Text = val.ToString("0.000")
        End If
    End Sub

    Private Sub GeometricLostFocus(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        If Not String.IsNullOrWhiteSpace(tb.Text) Then
            Dim val As Decimal
            If Decimal.TryParse(tb.Text, val) Then
                ' Allow up to 5 decimal places without forcing trailing zeros
                tb.Text = val.ToString("0.#####")
            End If
        End If
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim tb = DirectCast(sender, TextBox)
        If _masterLimits IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
            Dim val As Decimal
            If Decimal.TryParse(tb.Text, val) Then
                Dim nominal = GetNominalFromTextBox(tb)
                If nominal > 0 Then
                    Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
                    Dim ul = Convert.ToDecimal(_masterLimits($"Obs{nomStr}_UL"))
                    Dim ll = Convert.ToDecimal(_masterLimits($"Obs{nomStr}_LL"))

                    If val > ul OrElse val < ll Then
                        tb.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' Light Pink
                        tb.Foreground = Brushes.Red
                    Else
                        tb.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' Light Green
                        tb.Foreground = Brushes.Black
                    End If
                    UpdateObservationLabel(nominal)
                End If
            Else
                tb.Background = Brushes.White : tb.Foreground = Brushes.Black
            End If
        Else
            tb.Background = Brushes.White : tb.Foreground = Brushes.Black
        End If

        CalculateTMU()
        UpdateOverallStatus()
        UpdateSizeLockState()
    End Sub

    Private Sub GeometricTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim tb = DirectCast(sender, TextBox)
        If _masterLimits IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
            Dim val As Decimal
            If Decimal.TryParse(tb.Text, val) Then
                Dim nameParts = tb.Name.Split("_"c) ' txtParallel_1
                Dim factor = nameParts(0).Replace("txt", "")
                Dim ulCol = If(factor = "Parallel", "Parallelism_UL", $"{factor}_UL")
                Dim llCol = If(factor = "Parallel", "Parallelism_LL", $"{factor}_LL")

                If _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    Dim ul = Convert.ToDecimal(_masterLimits(ulCol))
                    Dim ll = Convert.ToDecimal(_masterLimits(llCol))

                    If val > ul OrElse val < ll Then
                        tb.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' Light Pink
                        tb.Foreground = Brushes.Red
                    Else
                        tb.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' Light Green
                        tb.Foreground = Brushes.Black
                    End If
                End If
                UpdateGeometricLabel(factor)
            Else
                tb.Background = Brushes.White : tb.Foreground = Brushes.Black
            End If
        Else
            tb.Background = Brushes.White : tb.Foreground = Brushes.Black
        End If

        CalculateTMU()
        UpdateOverallStatus()
        UpdateSizeLockState()
    End Sub

    Private Sub EnableNextRange(currentNominal As Decimal)
        Dim nominalIdx = Array.IndexOf(nominalRanges, currentNominal)
        For i = nominalIdx + 1 To nominalRanges.Length - 1
            If IsNominalInRange(nominalRanges(i)) Then
                EnableRange(nominalRanges(i))
                Exit For
            End If
        Next
    End Sub

    Private Sub EnableRange(nominal As Decimal)
        If ComboBox5.SelectedItem Is Nothing Then Return
        If Not IsNominalInRange(nominal) Then Return
        Dim tbs = GetTextBoxesForRange(nominal)
        For Each tb In tbs
            tb.IsEnabled = True
            If String.IsNullOrWhiteSpace(tb.Text) Then tb.Background = Brushes.White
        Next
        If tbs.Count > 0 Then tbs(0).Focus()
    End Sub

    Private Sub DisableAllObservationFields()
        For Each r In nominalRanges
            For Each tb In GetTextBoxesForRange(r)
                tb.IsEnabled = False : tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
        Next
        For Each gf In geometricFactors
            For Each tb In GetTextBoxesForGeometric(gf)
                tb.IsEnabled = False : tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
        Next
    End Sub

    Private Sub ResetAllTrials()
        For Each r In nominalRanges
            For Each tb In GetTextBoxesForRange(r)
                tb.Clear()
            Next
        Next
        For Each gf In geometricFactors
            For Each tb In GetTextBoxesForGeometric(gf)
                tb.Clear()
            Next
        Next
        If TextBox21 IsNot Nothing Then AdjustVisibilityBySize(TextBox21.Text)
        CalculateTMU()
        UpdateSizeLockState()
    End Sub

    Private Function GetTextBoxesForRange(nominal As Decimal) As List(Of TextBox)
        Dim list As New List(Of TextBox)
        Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
        For i = 1 To 3
            Dim name = $"txt{nomStr}_{i}"
            Dim tb = DirectCast(Me.FindName(name), TextBox)
            If tb IsNot Nothing Then list.Add(tb)
        Next
        Return list
    End Function

    Private Function GetTextBoxesForGeometric(gf As String) As List(Of TextBox)
        Dim list As New List(Of TextBox)
        For i = 1 To 3
            Dim name = $"txt{gf}_{i}"
            Dim tb = DirectCast(Me.FindName(name), TextBox)
            If tb IsNot Nothing Then list.Add(tb)
        Next
        Return list
    End Function

    Private Function GetNominalFromTextBox(tb As TextBox) As Decimal
        Dim name = tb.Name.Replace("txt", "")
        Dim parts = name.Split("_"c)
        If parts.Length >= 2 Then
            Dim nominal As Decimal
            If Decimal.TryParse(parts(0) & "." & parts(1), nominal) Then Return nominal
        End If
        Return -1
    End Function

    Private Sub LoadInstrumentDetails(controlNo As String)
        Try
            Dim dtSettings = MySqlCls.ReadDatatable("SELECT TypeName, Category FROM type_details")
            Dim found As Boolean = False
            For Each r As DataRow In dtSettings.Rows
                Dim tbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString())
                Dim dt = MySqlCls.ReadDatatable($"SELECT * FROM `{tbl}` WHERE ControlNo = '{controlNo.Replace("'", "''")}' LIMIT 1")
                If dt.Rows.Count > 0 Then
                    Dim row = dt.Rows(0)
                    Dim cycleName = MySqlCls.GetActiveCycleName()
                    Dim dtDept = MySqlCls.ReadDatatable($"SELECT Department, Color, SizeandRange FROM department_list WHERE `Control No` = '{controlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")

                    _category = r("Category").ToString()
                    Dim prefix = MySQLClass.ExtractPrefix(controlNo)
                    Dim sizes = MySqlCls.GetSizesForGroup(prefix, _category)
                    
                    ' Fallback if no sizes found for group
                    If sizes.Count = 0 Then
                        sizes = New List(Of String) From {"0-25 mm", "25-50 mm", "50-75 mm"}
                    End If
                    
                    TextBox21.ItemsSource = sizes

                    ' Enhanced Size Matching
                    Dim instrumentSize = ""
                    If dtDept.Rows.Count > 0 AndAlso Not IsDBNull(dtDept.Rows(0)("SizeandRange")) Then
                        instrumentSize = dtDept.Rows(0)("SizeandRange").ToString()
                    ElseIf row.Table.Columns.Contains("Size") AndAlso Not IsDBNull(row("Size")) Then
                        instrumentSize = row("Size").ToString()
                    End If

                    If Not String.IsNullOrEmpty(instrumentSize) Then
                        Dim matchedSize = TryMatchSize(instrumentSize, sizes)
                        If matchedSize IsNot Nothing Then
                            TextBox21.Text = matchedSize
                        Else
                            TextBox21.SelectedIndex = If(sizes.Count > 0, 0, -1)
                        End If
                    Else
                        TextBox21.SelectedIndex = If(sizes.Count > 0, 0, -1)
                    End If

                    AdjustVisibilityBySize(TextBox21.Text)
                    UpdateSizeLockState()

                    If dtDept.Rows.Count > 0 Then
                        ComboBox2.Text = dtDept.Rows(0)("Department").ToString()
                        TextBoxColor.Text = dtDept.Rows(0)("Color").ToString()
                    Else
                        ComboBox2.Text = If(row.Table.Columns.Contains("Location"), row("Location").ToString(), "")
                        TextBoxColor.Text = If(row.Table.Columns.Contains("Color"), row("Color").ToString(), "")
                    End If
                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), If(row.Table.Columns.Contains("InstrumentName"), row("InstrumentName").ToString(), ""), If(row.Table.Columns.Contains("GaugeName"), row("GaugeName").ToString(), ""))
                    found = True : Exit For
                End If
            Next
            If Not found Then
                TextBox21.ItemsSource = New List(Of String) From {"0-25 mm", "25-50 mm", "50-75 mm"}
                TextBox21.Text = "0-25 mm" : UpdateSizeLockState()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function TryMatchSize(inputSize As String, availableSizes As List(Of String)) As String
        If String.IsNullOrEmpty(inputSize) Then Return Nothing
        Dim normalizedInput = inputSize.Replace(" ", "").Replace("mm", "").ToLower()

        ' Try exact normalized match
        For Each s In availableSizes
            If s.Replace(" ", "").Replace("mm", "").ToLower() = normalizedInput Then Return s
        Next

        ' Try range matching
        If normalizedInput.Contains("-") Then
            Dim parts = normalizedInput.Split("-"c)
            If parts.Length = 2 Then
                Dim min = parts(0).Trim(), max = parts(1).Trim()
                For Each s In availableSizes
                    Dim sNorm = s.Replace(" ", "").Replace("mm", "").ToLower()
                    If sNorm.StartsWith(min & "-") AndAlso sNorm.EndsWith("-" & max) Then Return s
                Next
            End If
        End If
        Return Nothing
    End Function

    Private Sub UpdateSizeLockState()
        Dim hasData = False
        For Each r In nominalRanges
            If GetTextBoxesForRange(r).Any(Function(tb) Not String.IsNullOrWhiteSpace(tb.Text)) Then
                hasData = True : Exit For
            End If
        Next
        If TextBox21 IsNot Nothing Then TextBox21.IsEnabled = Not hasData
    End Sub

    Private Sub BtnBrowseMasterObservation_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMastersObservation.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMastersObservation = selector.SelectedMasters
            UpdateMasterUIObservation()
            CalculateTMU()
        End If
    End Sub

    Private Sub BtnBrowseMasterGeometric_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMastersGeometric.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMastersGeometric = selector.SelectedMasters
            UpdateMasterUIGeometric()
            CalculateTMU()
        End If
    End Sub

    Private Sub UpdateMasterUIObservation()
        ItemsControlSelectedMastersObservation.ItemsSource = Nothing
        ItemsControlSelectedMastersObservation.ItemsSource = _selectedMastersObservation
        MasterGroupBoxObservation.Visibility = If(_selectedMastersObservation.Count > 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub UpdateMasterUIGeometric()
        ItemsControlSelectedMastersGeometric.ItemsSource = Nothing
        ItemsControlSelectedMastersGeometric.ItemsSource = _selectedMastersGeometric
        MasterGroupBoxGeometric.Visibility = If(_selectedMastersGeometric.Count > 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub TimerRealTime_Tick(sender As Object, e As EventArgs)
        If Not _isTimeOutManual AndAlso Not String.IsNullOrWhiteSpace(TextBox24.Text) Then
            TextBox26.Text = DateTime.Now.ToString("HH:mm")
        End If
    End Sub

    Private Sub LoadDynamicLC()
        Try
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM blade_micrometer_75 WHERE RowType='MASTER' ORDER BY LC ASC")
            ComboBox5.Items.Clear()
            For Each row As DataRow In dt.Rows
                ComboBox5.Items.Add(row("LC").ToString())
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox5_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboBox5.SelectionChanged
        If ComboBox5 Is Nothing OrElse ComboBox5.SelectedItem Is Nothing Then Return
        Dim selectedLC As String = ComboBox5.SelectedItem.ToString()
        _masterLimits = MySqlCls.GetBladeMicrometer75MasterLimits(selectedLC)
        AdjustVisibilityBySize(TextBox21.Text)
        CalculateTMU()
    End Sub

    Private Sub CalculateTMU()
        If _isPopulating OrElse TextBox29 Is Nothing OrElse ComboBox5.SelectedItem Is Nothing Then Return
        Try
            Dim masterU_Geometric As Decimal = 0D
            If _selectedMastersGeometric.Count > 0 Then
                masterU_Geometric = CDec(Math.Sqrt(_selectedMastersGeometric.Sum(Function(m) m.MasterUncertainty ^ 2)))
            Else
                masterU_Geometric = 0.000072D ' Fallback
            End If

            Dim masterU_Observation As Decimal = 0D
            If _selectedMastersObservation.Count > 0 Then
                masterU_Observation = CDec(Math.Sqrt(_selectedMastersObservation.Sum(Function(m) m.MasterUncertainty ^ 2)))
            Else
                masterU_Observation = 0.00012D ' Fallback
            End If
            Dim totalSumSq = 0D

            ' Geometric (Parallelism)
            For Each gf In geometricFactors
                Dim tbs = GetTextBoxesForGeometric(gf)
                If tbs.Any(Function(t) Not String.IsNullOrWhiteSpace(t.Text)) Then
                    Dim vals = tbs.Select(Function(t) If(Decimal.TryParse(t.Text, Nothing), Decimal.Parse(t.Text), 0D)).ToArray()
                    Dim sigmaB1 = calcStdev(vals) / Math.Sqrt(3)
                    totalSumSq += (sigmaB1 ^ 2) + (masterU_Geometric ^ 2)
                End If
            Next

            ' Accuracy
            For Each r In nominalRanges
                Dim tbs = GetTextBoxesForRange(r)
                If tbs.Count > 0 AndAlso tbs(0).IsEnabled AndAlso tbs.Any(Function(t) Not String.IsNullOrWhiteSpace(t.Text)) Then
                    Dim vals = tbs.Select(Function(t) If(Decimal.TryParse(t.Text, Nothing), Decimal.Parse(t.Text), 0D)).ToArray()
                    Dim sigmaB1 = calcStdev(vals) / Math.Sqrt(3)
                    totalSumSq += (sigmaB1 ^ 2) + (masterU_Observation ^ 2)
                End If
            Next

            TextBox29.Text = (2 * Math.Sqrt(totalSumSq)).ToString("F9")
        Catch ex As Exception
        End Try
    End Sub

    Private Function calcStdev(arr() As Decimal) As Decimal
        If arr.Length <= 1 Then Return 0
        Dim avg = arr.Average()
        Dim sumSq = arr.Sum(Function(x) (x - avg) ^ 2)
        Return Math.Sqrt(sumSq / (arr.Length - 1))
    End Function

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs)
        ' Validation checks
        For Each r In nominalRanges
            If IsNominalInRange(r) Then
                Dim tbs = GetTextBoxesForRange(r)
                Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
                If count > 0 AndAlso count < 3 Then
                    MessageBox.Show($"Please complete all 3 trials for {r} mm.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return
                End If
            End If
        Next
        For Each gf In geometricFactors
            Dim tbs = GetTextBoxesForGeometric(gf)
            If tbs.Any(Function(t) String.IsNullOrWhiteSpace(t.Text)) Then
                MessageBox.Show($"Please complete all 3 trials for {gf}.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return
            End If
        Next

        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse ComboBox5.SelectedItem Is Nothing Then
            MessageBox.Show("Control No and LC must be filled.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return
        End If

        Dim type = "Blade Micrometer"
        Dim controlNo = TextBox1.Text.Trim()
        Dim cycleName = MySqlCls.GetActiveCycleName()
        Dim size = TextBox21.Text
        Dim lc = ComboBox5.Text
        Dim color = TextBoxColor.Text
        Dim location = ComboBox2.Text
        Dim temp = TextBox30.Text
        Dim humidity = TextBox28.Text
        Dim tmu = TextBox29.Text
        Dim remark = TextBox19.Text
        Dim timeIn = TextBox24.Text
        Dim timeOut = TextBox26.Text
        Dim totalTime = TextBox25.Text
        Dim status = ComboBox4.Text

        Dim timeInStr = TextBox24.Text.Trim(), timeOutStr = TextBox26.Text.Trim()
        If String.IsNullOrEmpty(timeOutStr) Then timeOutStr = DateTime.Now.ToString("HH:mm") : TextBox26.Text = timeOutStr
        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then MessageBox.Show("Min 2 hours duration required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return

        ' Persist environment data for next instance
        Dim dateStr = If(TxtCalibrationDate.SelectedDate, DateTime.Today).ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, TextBox20.Text)

        Dim obsArr As New Dictionary(Of Decimal, Object())
        For Each r In nominalRanges
            obsArr.Add(r, GetTextBoxesForRange(r).Select(Function(t) If(String.IsNullOrEmpty(t.Text), "-", t.Text)).Cast(Of Object)().ToArray())
        Next

        Dim parallelTbs = GetTextBoxesForGeometric("Parallel")
        Dim parallelism = parallelTbs.Select(Function(t) If(String.IsNullOrEmpty(t.Text), "-", t.Text)).Cast(Of Object)().ToArray()

        Dim success As Boolean = False
        Dim depthErrorVal = If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text)
        If IsEditMode Then
            success = MySqlCls.UpdateBladeMicrometer75(type, controlNo, cycleName, size, lc, color, location, temp, humidity, tmu, parallelism, obsArr, depthErrorVal, timeIn, timeOut, totalTime, status, remark, TxtCalibrationDate.SelectedDate)
        Else
            success = MySqlCls.InsertBladeMicrometer75(type, controlNo, cycleName, size, lc, color, location, temp, humidity, tmu, parallelism, obsArr, depthErrorVal, timeIn, timeOut, totalTime, status, remark, TxtCalibrationDate.SelectedDate)
        End If

        If success Then
            Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
            Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)
            MySqlCls.UpdateRegularCalibrationStatus(controlNo, calibDate, status, targetCyc)
            MySqlCls.InsertResultRecord(controlNo, _category, _instrumentName, targetCyc)

            ' Save masters to cache
            CalibrationMasterCache.SaveMastersDual(controlNo, targetCyc, _selectedMastersObservation, _selectedMastersGeometric)
            
            ' NG/WOP Logic if status is NG
            If status = "NG" Then
                ProcessNGFlow(controlNo, targetCyc, calibDate, status)
            End If
            
            MessageBox.Show("Record saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.Close()
        Else
            MessageBox.Show("Failed to save record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub ProcessNGFlow(ctrlNo As String, cycleName As String, calibDate As DateTime, statusVal As String)
        Dim dtDept = MySqlCls.ReadDatatable($"SELECT * FROM department_list WHERE `Control No` = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")
        If dtDept.Rows.Count > 0 Then
            Dim row = dtDept.Rows(0)
            Dim dept = row("Department").ToString(), instName = row("InstrumentName").ToString()
            Dim color = row("Color").ToString(), size = row("SizeandRange").ToString()
            Dim ngReason = If(Not String.IsNullOrWhiteSpace(TextBox19.Text), TextBox19.Text.Trim(), "Calibration NG")
            Dim wopReason = "Calibration NG" & If(Not String.IsNullOrWhiteSpace(TextBox19.Text), ": " & TextBox19.Text.Trim(), "")
            Dim dueDate = calibDate.AddYears(1).AddDays(-1)

            MySqlCls.InsertNGRecord(ctrlNo, instName, dept, cycleName, ngReason, calibDate, dueDate, statusVal)
            MySqlCls.InsertWOPRecord(ctrlNo, cycleName, dept, instName, color, wopReason)
            MySqlCls.InsertInterchangeRecord(cycleName, ctrlNo, dept, instName, size, color, "WOP", wopReason)
            MySqlCls.ClearRFIDTag(ctrlNo)
        End If
    End Sub

    Private Sub UpdateDuration(sender As Object, e As TextChangedEventArgs)
        Dim timeInStr = TextBox24.Text.Trim(), timeOutStr = TextBox26.Text.Trim()
        Dim dummyTimeOut = If(String.IsNullOrEmpty(timeOutStr), DateTime.Now.ToString("HH:mm"), timeOutStr)
        Dim totalMinutes = CalculateDuration(timeInStr, dummyTimeOut)
        If totalMinutes >= 0 Then
            Dim hours = totalMinutes \ 60, mins = totalMinutes Mod 60
            TextBox25.Text = $"{hours:D2}:{mins:D2}"
            TextBox25.Background = If(totalMinutes < 120, New SolidColorBrush(Color.FromRgb(255, 182, 193)), New SolidColorBrush(Color.FromRgb(144, 238, 144)))
        Else
            TextBox25.Text = "" : TextBox25.Background = Brushes.White
        End If
    End Sub

    Private Sub TriggerAllValidations()
        Try
            _isPopulating = False ' Ensure it's false so handlers run
            
            ' Environment
            EnvironmentTextChanged(TextBox30, Nothing)
            EnvironmentTextChanged(TextBox28, Nothing)
            
            ' Geometric
            For Each gf In geometricFactors
                For Each tb In GetTextBoxesForGeometric(gf)
                    GeometricTextChanged(tb, Nothing)
                Next
            Next
            
            ' Accuracy
            For Each r In nominalRanges
                If IsNominalInRange(r) Then
                    For Each tb In GetTextBoxesForRange(r)
                        TrialTextChanged(tb, Nothing)
                    Next
                End If
            Next
            
            UpdateOverallStatus()
        Catch ex As Exception
        End Try
    End Sub

    Private Function CalculateDuration(timeInStr As String, timeOutStr As String) As Integer
        Try
            Dim timeIn, timeOut As DateTime
            Dim formats As String() = {"HH:mm", "H:mm"}
            If DateTime.TryParseExact(timeInStr.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, timeIn) AndAlso
               DateTime.TryParseExact(timeOutStr.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, timeOut) Then
                If timeOut < timeIn Then timeOut = timeOut.AddDays(1)
                Return CInt((timeOut - timeIn).TotalMinutes)
            End If
        Catch ex As Exception
        End Try
        Return -1
    End Function
End Class

