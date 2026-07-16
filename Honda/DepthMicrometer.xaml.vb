Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class DepthMicrometer
    Inherits Window

    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = ""
    Private _category As String = "Instrument"
    Private _selectedMastersObservation As New List(Of MasterSelectorItem)
    Private _selectedMastersGeometric As New List(Of MasterSelectorItem)



    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _activeMinRange As Decimal = 0
    Private _activeMaxRange As Decimal = 150

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""
    Private _isPopulating As Boolean = False

    ' 31 Nominal Ranges
    Private ReadOnly nominalRanges As Decimal() = {
        0.0D, 5.0D, 10.0D, 15.0D, 20.0D, 25.0D, 30.0D, 35.0D, 40.0D, 45.0D, 50.0D,
        55.0D, 60.0D, 65.0D, 70.0D, 75.0D, 80.0D, 85.0D, 90.0D, 95.0D, 100.0D,
        105.0D, 110.0D, 115.0D, 120.0D, 125.0D, 130.0D, 135.0D, 140.0D, 145.0D, 150.0D
    }

    ' Geometric Factors
    Private ReadOnly geometricFactors As String() = {"Anvil", "Spindle", "Parallel", "Zero"}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = "Depth Micrometer 150"
        LoadDynamicLC()

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
                tb.IsEnabled = False
                AddHandler tb.TextChanged, AddressOf GeometricTextChanged
                AddHandler tb.KeyDown, AddressOf GeometricKeyDown
                AddHandler tb.LostFocus, AddressOf GeometricLostFocus
            Next
        Next

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

        ' Add handlers for Environmental conditions
        AddHandler TextBox30.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox28.TextChanged, AddressOf EnvironmentTextChanged

        ' Initialize LC to none and disable all observations initially
        ComboBox5.SelectedIndex = -1
        DisableAllObservationFields()
        ResetAllTrials()
        UpdateSizeLockState()

        If IsEditMode Then
            PopulateCalibrationData()
        End If
    End Sub

    Private Sub TxtCalibrationDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles TxtCalibrationDate.SelectedDateChanged
        If _isPopulating OrElse IsEditMode Then Return
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
        Try
            Dim dt = MySqlCls.GetDepthMicrometer150Data(TextBox1.Text.Trim(), TargetCycle)
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)

                _isPopulating = True

                TxtCalibrationDate.SelectedDate = Convert.ToDateTime(row("Date"))
                TextBox21.Text = row("Size").ToString()

                ' Set LC by matching the item
                Dim lcVal = row("LC").ToString()
                For i = 0 To ComboBox5.Items.Count - 1
                    If ComboBox5.Items(i).ToString() = lcVal Then
                        ComboBox5.SelectedIndex = i
                        Exit For
                    End If
                Next

                TextBoxColor.Text = row("Color").ToString()
                ComboBox2.Text = row("Location").ToString()
                TextBox30.Text = row("Temperature").ToString()
                TextBox28.Text = row("Humidity").ToString()

                ' Enable geometric fields for population
                For Each gf In geometricFactors
                    For Each tb In GetTextBoxesForGeometric(gf)
                        tb.IsEnabled = True
                        tb.Background = Brushes.White
                    Next
                Next

                ' Enable observation fields based on size
                AdjustVisibilityBySize(TextBox21.Text)

                ' Populate Accuracy Trials
                For Each r In nominalRanges
                    Dim tbs = GetTextBoxesForRange(r)
                    If tbs.Count > 0 AndAlso tbs(0).IsEnabled Then
                        SetTrialsFromDB(row, $"Obs{r.ToString("0")}_", tbs)
                    End If
                Next

                ' Populate Geometric Trials
                SetGeometricFromDB(row, "Anvil", GetTextBoxesForGeometric("Anvil"))
                SetGeometricFromDB(row, "Spindle", GetTextBoxesForGeometric("Spindle"))
                SetGeometricFromDB(row, "Parallel", GetTextBoxesForGeometric("Parallel"))
                SetGeometricFromDB(row, "Zero", GetTextBoxesForGeometric("Zero"))

                TextBox20.Text = If(row("DepthError").ToString() = "-", "", row("DepthError").ToString())
                TextBox24.Text = row("TimeIn").ToString()
                TextBox26.Text = row("TimeOut").ToString()
                TextBox25.Text = row("TotalTime").ToString()
                TextBox19.Text = row("Remark").ToString()

                Dim status = row("Status").ToString()
                For Each item As ComboBoxItem In ComboBox4.Items
                    If item.Content.ToString() = status Then
                        ComboBox4.SelectedItem = item
                        Exit For
                    End If
                Next

                ' Restore cached calibration masters
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
                TextBox29.Text = row("TMU").ToString()
            End If
        Catch ex As Exception
            _isPopulating = False
            Console.WriteLine("PopulateCalibrationData Error: " & ex.Message)
        End Try
    End Sub

    Private Sub SetTrialsFromDB(row As DataRow, prefix As String, tbs As List(Of TextBox))
        For i = 0 To 2
            If i < tbs.Count Then
                Dim val = row($"{prefix}{i + 1}").ToString()
                If val <> "-" Then tbs(i).Text = val
            End If
        Next
    End Sub

    Private Sub SetGeometricFromDB(row As DataRow, gf As String, tbs As List(Of TextBox))
        For i = 0 To 2
            If i < tbs.Count Then
                Dim colName = ""
                If gf = "Parallel" Then
                    colName = $"Parallelism_{i + 1}"
                ElseIf gf = "Zero" Then
                    colName = $"Zero_Error_{i + 1}"
                ElseIf gf = "Anvil" Then
                    colName = $"Flatness_Meas_{i + 1}"
                ElseIf gf = "Spindle" Then
                    colName = $"Flatness_Ref_{i + 1}"
                End If
                If row.Table.Columns.Contains(colName) Then
                    Dim val = row(colName).ToString()
                    If val <> "-" Then tbs(i).Text = val
                End If
            End If
        Next
    End Sub

    Private Sub ComboBox21_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles TextBox21.SelectionChanged
        AdjustVisibilityBySize(TextBox21.Text)
    End Sub

    Private Sub ComboBox21_DropDownClosed(sender As Object, e As EventArgs) Handles TextBox21.DropDownClosed
        AdjustVisibilityBySize(TextBox21.Text)
    End Sub

    Private Sub ComboBox21_LostFocus(sender As Object, e As RoutedEventArgs) Handles TextBox21.LostFocus
        AdjustVisibilityBySize(TextBox21.Text)
    End Sub

    Private Function IsNominalInRange(nominal As Decimal) As Boolean
        Return nominal >= _activeMinRange AndAlso nominal <= _activeMaxRange
    End Function

    Private Sub AdjustVisibilityBySize(sizeStr As String)
        If TextBox21 Is Nothing Then Return
        Try
            _activeMinRange = 0 : _activeMaxRange = 150
            If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
                Dim parts = sizeStr.Split("-"c)
                If parts.Length > 1 Then
                    Decimal.TryParse(parts(0).Replace("mm", "").Trim(), _activeMinRange)
                    Decimal.TryParse(parts(1).Replace("mm", "").Trim(), _activeMaxRange)
                End If
            End If

            If TabItem0_25 IsNot Nothing Then TabItem0_25.IsEnabled = (_activeMinRange <= 0 AndAlso _activeMaxRange >= 25)
            If TabItem25_50 IsNot Nothing Then TabItem25_50.IsEnabled = (_activeMinRange <= 25 AndAlso _activeMaxRange >= 50)
            If TabItem50_75 IsNot Nothing Then TabItem50_75.IsEnabled = (_activeMinRange <= 50 AndAlso _activeMaxRange >= 75)
            If TabItem75_100 IsNot Nothing Then TabItem75_100.IsEnabled = (_activeMinRange <= 75 AndAlso _activeMaxRange >= 100)
            If TabItem100_125 IsNot Nothing Then TabItem100_125.IsEnabled = (_activeMinRange <= 100 AndAlso _activeMaxRange >= 125)
            If TabItem125_150 IsNot Nothing Then TabItem125_150.IsEnabled = (_activeMinRange <= 125 AndAlso _activeMaxRange >= 150)
            
            ' Automatically select the first enabled tab
            If TabControlObservations IsNot Nothing Then
                If TabItem0_25 IsNot Nothing AndAlso TabItem0_25.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem0_25
                ElseIf TabItem25_50 IsNot Nothing AndAlso TabItem25_50.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem25_50
                ElseIf TabItem50_75 IsNot Nothing AndAlso TabItem50_75.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem50_75
                ElseIf TabItem75_100 IsNot Nothing AndAlso TabItem75_100.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem75_100
                ElseIf TabItem100_125 IsNot Nothing AndAlso TabItem100_125.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem100_125
                ElseIf TabItem125_150 IsNot Nothing AndAlso TabItem125_150.IsEnabled Then
                    TabControlObservations.SelectedItem = TabItem125_150
                End If
            End If

            For Each r In nominalRanges
                Dim inRange = IsNominalInRange(r)
                Dim tbs = GetTextBoxesForRange(r)
                Dim nomStr = r.ToString("0.0").Replace(".", "_")
                Dim lblName = $"Label{nomStr}"
                Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)

                For Each tb In tbs
                    If Not inRange Then
                        tb.IsEnabled = False
                        tb.Text = "-"
                        tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                        If lbl IsNot Nothing Then lbl.Foreground = Brushes.LightGray
                    Else
                        If ComboBox5.SelectedItem IsNot Nothing Then
                            tb.IsEnabled = True
                            If String.IsNullOrWhiteSpace(tb.Text) OrElse tb.Text = "-" Then 
                                tb.Text = ""
                                tb.Background = Brushes.White
                            ElseIf tb.Background Is Nothing OrElse (TypeOf tb.Background Is SolidColorBrush AndAlso DirectCast(tb.Background, SolidColorBrush).Color = Color.FromRgb(241, 245, 249)) Then
                                tb.Background = Brushes.White
                            End If
                            If lbl IsNot Nothing Then UpdateObservationLabel(r)
                        Else
                            tb.IsEnabled = False
                            tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                            If lbl IsNot Nothing Then lbl.Foreground = New SolidColorBrush(Color.FromRgb(51, 65, 85))
                        End If
                    End If
                Next
            Next

            If ComboBox5.SelectedItem IsNot Nothing Then
                Dim firstInRange = nominalRanges.FirstOrDefault(Function(r) IsNominalInRange(r))
                If firstInRange >= 0 Then EnableRange(firstInRange)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub GeometricKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)
            Dim nameParts = tb.Name.Split("_"c) ' txtAnvil_1
            If nameParts.Length = 2 Then
                Dim factor = nameParts(0).Replace("txt", "")
                Dim trialNum = Integer.Parse(nameParts(1))
                If trialNum < 3 Then
                    Dim nextTbName = "txt" & factor & "_" & (trialNum + 1)
                    Dim nextTb = DirectCast(Me.FindName(nextTbName), TextBox)
                    If nextTb IsNot Nothing Then nextTb.Focus()
                Else
                    ' Next Factor
                    Dim factorIdx = Array.IndexOf(geometricFactors, factor)
                    If factorIdx < geometricFactors.Length - 1 Then
                        Dim nextFactor = geometricFactors(factorIdx + 1)
                        Dim nextTbName = "txt" & nextFactor & "_1"
                        Dim nextTb = DirectCast(Me.FindName(nextTbName), TextBox)
                        If nextTb IsNot Nothing Then nextTb.Focus()
                    Else
                        ' End of Geometry, move to Observations
                        If ComboBox5.SelectedItem IsNot Nothing Then
                            Dim firstInRange = nominalRanges.FirstOrDefault(Function(r) IsNominalInRange(r))
                            If firstInRange >= 0 Then
                                Dim obsTbs = GetTextBoxesForRange(firstInRange)
                                If obsTbs.Count > 0 Then obsTbs(0).Focus()
                            End If
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
            If nameParts.Length >= 3 Then
                Dim trialNumStr = nameParts(2)
                If trialNumStr.Contains("T") Then
                    trialNumStr = nameParts(2).Split("T"c)(0).Replace("_", "").Trim()
                End If
                Dim trialNum = Integer.Parse(trialNumStr)
                If trialNum < 3 Then
                    Dim suffix = ""
                    If tb.Name.Contains("_T2") Then suffix = "_T2"
                    If tb.Name.Contains("_T3") Then suffix = "_T3"
                    If tb.Name.Contains("_T4") Then suffix = "_T4"
                    If tb.Name.Contains("_T5") Then suffix = "_T5"
                    If tb.Name.Contains("_T6") Then suffix = "_T6"
                    Dim nextTbName = nameParts(0) & "_" & nameParts(1) & "_" & (trialNum + 1) & suffix
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
            If Decimal.TryParse(tb.Text, val) Then
                tb.Text = val.ToString("0.000")
            End If
        End If
    End Sub

    Private Sub GeometricLostFocus(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        If Not String.IsNullOrWhiteSpace(tb.Text) Then
            Dim val As Decimal
            If Decimal.TryParse(tb.Text, val) Then
                tb.Text = val.ToString("0.#####")
            End If
        End If
    End Sub

    Private Sub EnvironmentTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        If String.IsNullOrWhiteSpace(txtBox.Text) Then
            txtBox.Background = Brushes.White
            UpdateOverallStatus()
            Return
        End If

        Dim val As Decimal
        If Decimal.TryParse(txtBox.Text, val) Then
            If txtBox.Name = "TextBox30" Then ' Temperature (20 ± 2 °C)
                txtBox.Background = If(val >= 18D AndAlso val <= 22D, New SolidColorBrush(Color.FromRgb(144, 238, 144)), New SolidColorBrush(Color.FromRgb(255, 182, 193)))
            ElseIf txtBox.Name = "TextBox28" Then ' Humidity (40 - 60 %)
                txtBox.Background = If(val >= 40D AndAlso val <= 60D, New SolidColorBrush(Color.FromRgb(144, 238, 144)), New SolidColorBrush(Color.FromRgb(255, 182, 193)))
            Else
                txtBox.Background = Brushes.White
            End If
        Else
            txtBox.Background = Brushes.White
        End If
        UpdateOverallStatus()
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim txtBox As TextBox = DirectCast(sender, TextBox)

        ' Sync overlap textboxes first
        SyncOverlapTextBoxes(txtBox)

        If _masterLimits Is Nothing Then Return

        Dim nominal As Decimal = GetNominalFromTextBox(txtBox)
        If nominal = -1 Then Return

        Dim nomDbStr = nominal.ToString("0")
        Dim ulCol As String = $"Obs{nomDbStr}_UL"
        Dim llCol As String = $"Obs{nomDbStr}_LL"

        Try
            UpdateObservationLabel(nominal)
            CalculateTMU()
            UpdateSizeLockState()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If

            Dim val As Decimal
            If Decimal.TryParse(txtBox.Text, val) Then
                If _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    Dim ul = Convert.ToDecimal(_masterLimits(ulCol)), ll = Convert.ToDecimal(_masterLimits(llCol))
                    txtBox.Background = If(val >= ll AndAlso val <= ul, New SolidColorBrush(Color.FromRgb(144, 238, 144)), New SolidColorBrush(Color.FromRgb(255, 182, 193)))
                End If
                UpdateOverallStatus()
            Else
                txtBox.Background = Brushes.White
                UpdateOverallStatus()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SyncOverlapTextBoxes(tb As TextBox)
        If _isPopulating Then Return
        Dim name = tb.Name
        Dim targetName As String = ""
        If name.EndsWith("_T2") Then
            targetName = name.Replace("_T2", "")
        ElseIf name.StartsWith("txt25_0_") AndAlso Not name.Contains("_T") Then
            targetName = name & "_T2"
        ElseIf name.EndsWith("_T3") Then
            targetName = name.Replace("_T3", "")
        ElseIf name.StartsWith("txt50_0_") AndAlso Not name.Contains("_T") Then
            targetName = name & "_T3"
        ElseIf name.EndsWith("_T4") Then
            targetName = name.Replace("_T4", "")
        ElseIf name.StartsWith("txt75_0_") AndAlso Not name.Contains("_T") Then
            targetName = name & "_T4"
        ElseIf name.EndsWith("_T5") Then
            targetName = name.Replace("_T5", "")
        ElseIf name.StartsWith("txt100_0_") AndAlso Not name.Contains("_T") Then
            targetName = name & "_T5"
        ElseIf name.EndsWith("_T6") Then
            targetName = name.Replace("_T6", "")
        ElseIf name.StartsWith("txt125_0_") AndAlso Not name.Contains("_T") Then
            targetName = name & "_T6"
        End If

        If Not String.IsNullOrEmpty(targetName) Then
            Dim targetTb = DirectCast(Me.FindName(targetName), TextBox)
            If targetTb IsNot Nothing AndAlso targetTb.Text <> tb.Text Then
                _isPopulating = True
                targetTb.Text = tb.Text
                targetTb.Background = tb.Background
                _isPopulating = False
            End If
        End If
    End Sub

    Private Sub GeometricTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        If _masterLimits Is Nothing Then Return

        Dim gf = geometricFactors.FirstOrDefault(Function(g) txtBox.Name.Contains(g))
        If gf Is Nothing Then Return

        Try
            CalculateTMU()
            UpdateSizeLockState()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If

            Dim val As Decimal
            If Decimal.TryParse(txtBox.Text, val) Then
                Dim ulCol = "", llCol = ""
                If gf = "Anvil" Then
                    ulCol = "Flatness_Meas_UL" : llCol = "Flatness_Meas_LL"
                ElseIf gf = "Spindle" Then
                    ulCol = "Flatness_Ref_UL" : llCol = "Flatness_Ref_LL"
                ElseIf gf = "Parallel" Then
                    ulCol = "Parallelism_UL" : llCol = "Parallelism_LL"
                ElseIf gf = "Zero" Then
                    ulCol = "Zero_Error_UL" : llCol = "Zero_Error_LL"
                End If

                If _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    Dim ul = Convert.ToDecimal(_masterLimits(ulCol)), ll = Convert.ToDecimal(_masterLimits(llCol))
                    txtBox.Background = If(val >= ll AndAlso val <= ul, New SolidColorBrush(Color.FromRgb(144, 238, 144)), New SolidColorBrush(Color.FromRgb(255, 182, 193)))
                End If
                UpdateOverallStatus()
            End If
        Catch ex As Exception
        End Try
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
                tb.IsEnabled = False
                tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
        Next
        For Each gf In geometricFactors
            For Each tb In GetTextBoxesForGeometric(gf)
                tb.IsEnabled = False
                tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
        Next
    End Sub

    Private Sub ResetAllTrials()
        _isPopulating = True
        For Each r In nominalRanges
            Dim tbs1 = GetTextBoxesForRange(r)
            For Each tb In tbs1
                tb.Clear()
                tb.Background = Brushes.White
            Next
            Dim nomStr = r.ToString("0.0").Replace(".", "_")
            For Each suffix In {"_T2", "_T3", "_T4", "_T5", "_T6"}
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txt{nomStr}_{i}{suffix}"), TextBox)
                    If tb IsNot Nothing Then
                        tb.Clear()
                        tb.Background = Brushes.White
                    End If
                Next
            Next
        Next
        For Each gf In geometricFactors
            For Each tb In GetTextBoxesForGeometric(gf)
                tb.Clear()
                tb.Background = Brushes.White
            Next
        Next
        _isPopulating = False

        For Each gf In geometricFactors
            UpdateGeometricLabel(gf)
        Next
        If TextBox21 IsNot Nothing Then AdjustVisibilityBySize(TextBox21.Text)
        CalculateTMU()
        UpdateSizeLockState()
    End Sub

    Private Function GetTextBoxesForRange(nominal As Decimal) As List(Of TextBox)
        Dim list As New List(Of TextBox)
        Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
        
        Dim suffix = ""
        If nominal = 25.0D AndAlso Not (_activeMinRange <= 0 AndAlso _activeMaxRange >= 25) Then
            suffix = "_T2"
        ElseIf nominal = 50.0D AndAlso Not (_activeMinRange <= 25 AndAlso _activeMaxRange >= 50) Then
            suffix = "_T3"
        ElseIf nominal = 75.0D AndAlso Not (_activeMinRange <= 50 AndAlso _activeMaxRange >= 75) Then
            suffix = "_T4"
        ElseIf nominal = 100.0D AndAlso Not (_activeMinRange <= 75 AndAlso _activeMaxRange >= 100) Then
            suffix = "_T5"
        ElseIf nominal = 125.0D AndAlso Not (_activeMinRange <= 100 AndAlso _activeMaxRange >= 125) Then
            suffix = "_T6"
        End If

        For i = 1 To 3
            Dim name = $"txt{nomStr}_{i}{suffix}"
            Dim tb = DirectCast(Me.FindName(name), TextBox)
            If tb IsNot Nothing Then list.Add(tb)
        Next
        Return list
    End Function

    Private Sub UpdateObservationLabel(nominal As Decimal)
        Dim tbs = GetTextBoxesForRange(nominal)
        Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
        Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
        Dim lblName = $"Label{nomStr}"
        Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)

        If lbl IsNot Nothing Then
            lbl.Text = $"{nominal} mm:"
            If ComboBox5.SelectedItem Is Nothing Then
                lbl.Foreground = New SolidColorBrush(Color.FromRgb(51, 65, 85))
            Else
                lbl.Foreground = If(count >= 3, Brushes.Gray, Brushes.DodgerBlue)
            End If
        End If
    End Sub

    Private Sub UpdateGeometricLabel(gf As String)
        Dim tbs = GetTextBoxesForGeometric(gf)
        Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
        Dim lblName = $"Label{gf}"
        Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)

        If lbl IsNot Nothing Then
            If ComboBox5.SelectedItem Is Nothing Then
                lbl.Foreground = New SolidColorBrush(Color.FromRgb(51, 65, 85))
            Else
                lbl.Foreground = If(count >= 3, Brushes.Gray, Brushes.DodgerBlue)
            End If
        End If
    End Sub

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
        For Each suffix In {"_T2", "_T3", "_T4", "_T5", "_T6"}
            If name.EndsWith(suffix) Then
                name = name.Substring(0, name.Length - suffix.Length)
                Exit For
            End If
        Next
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
                        ComboBox2.Text = row("Location").ToString()
                        TextBoxColor.Text = row("Color").ToString()
                    End If
                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), row("InstrumentName").ToString(), row("GaugeName").ToString())
                    found = True : Exit For
                End If
            Next
            If Not found Then
                _instrumentName = "" : ComboBox2.Text = "" : TextBoxColor.Text = "" : TextBox21.Text = "0-25 mm" : UpdateSizeLockState()
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
            Dim tbs = GetTextBoxesForRange(r)
            If tbs.Any(Function(tb) Not String.IsNullOrWhiteSpace(tb.Text)) Then
                hasData = True : Exit For
            End If
        Next
        If Not hasData Then
            For Each gf In geometricFactors
                Dim tbs = GetTextBoxesForGeometric(gf)
                If tbs.Any(Function(tb) Not String.IsNullOrWhiteSpace(tb.Text)) Then
                    hasData = True : Exit For
                End If
            Next
        End If
        If TextBox21 IsNot Nothing Then TextBox21.IsEnabled = Not hasData
    End Sub

    Private Sub BtnBrowseMasterObservation_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMastersObservation.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMastersObservation = selector.SelectedMasters
            UpdateMasterUIObservation() : CalculateTMU()
        End If
    End Sub

    Private Sub BtnBrowseMasterGeometric_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMastersGeometric.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMastersGeometric = selector.SelectedMasters
            UpdateMasterUIGeometric() : CalculateTMU()
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
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM depth_micrometer_150 WHERE RowType='MASTER' ORDER BY LC ASC")
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
        _masterLimits = MySqlCls.GetDepthMicrometer150MasterLimits(selectedLC)

        ' Enable geometric factors
        For Each gf In geometricFactors
            For Each tb In GetTextBoxesForGeometric(gf)
                tb.IsEnabled = True
                If String.IsNullOrWhiteSpace(tb.Text) Then tb.Background = Brushes.White
            Next
        Next

        If _isPopulating Then Return

        ResetAllTrials()
        If TextBox21 IsNot Nothing Then AdjustVisibilityBySize(TextBox21.Text)
        Dim firstInRange = nominalRanges.FirstOrDefault(Function(r) IsNominalInRange(r))
        If firstInRange >= 0 Then EnableRange(firstInRange)
    End Sub

    Private Sub UpdateOverallStatus()
        If ComboBox4 Is Nothing OrElse TextBox30 Is Nothing OrElse TextBox28 Is Nothing Then Return
        Dim isNG = False
        For Each r In nominalRanges
            For Each tb In GetTextBoxesForRange(r)
                If IsNGColor(tb.Background) Then isNG = True : Exit For
            Next
            if isNG then exit for
        Next
        If Not isNG Then
            For Each gf In geometricFactors
                For Each tb In GetTextBoxesForGeometric(gf)
                    If IsNGColor(tb.Background) Then isNG = True : Exit For
                Next
            Next
        End If
        If Not isNG Then
            For Each tb In {TextBox30, TextBox28}
                If IsNGColor(tb.Background) Then isNG = True : Exit For
            Next
        End If
        ComboBox4.SelectedIndex = If(isNG, 1, 0)
    End Sub

    Private Function IsNGColor(brush As Brush) As Boolean
        If brush IsNot Nothing AndAlso TypeOf brush Is SolidColorBrush Then
            Return DirectCast(brush, SolidColorBrush).Color = Color.FromRgb(255, 182, 193)
        End If
        Return False
    End Function

    Private Sub CalculateTMU()
        If TextBox29 Is Nothing Then Return
        Try
            ' 1. Dynamic Master Uncertainties
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

            ' 2. Geometric Repeatability
            For Each gf In geometricFactors
                Dim tbs = GetTextBoxesForGeometric(gf)
                If tbs.Count > 0 Then
                    Dim vals = tbs.Select(Function(t) If(Decimal.TryParse(t.Text, Nothing), Decimal.Parse(t.Text), 0D)).ToArray()
                    Dim sigmaB1 = calcStdev(vals) / Math.Sqrt(3)
                    totalSumSq += (sigmaB1 ^ 2) + (masterU_Geometric ^ 2)
                End If
            Next

            ' 3. Accuracy Repeatability
            For Each r In nominalRanges
                Dim tbs = GetTextBoxesForRange(r)
                If tbs.Count > 0 AndAlso tbs(0).IsEnabled Then
                    Dim vals = tbs.Select(Function(t) If(Decimal.TryParse(t.Text, Nothing), Decimal.Parse(t.Text), 0D)).ToArray()
                    Dim sigmaB1 = calcStdev(vals) / Math.Sqrt(3)
                    totalSumSq += (sigmaB1 ^ 2) + (masterU_Observation ^ 2)
                End If
            Next

            ' 4. Final TMU Calculation (k=2)
            TextBox29.Text = (2 * Math.Sqrt(totalSumSq)).ToString("F9")
            TextBox29.Background = Brushes.White
        Catch ex As Exception
            TextBox29.Text = "Error"
        End Try
    End Sub

    Private Function calcStdev(arr() As Decimal) As Decimal
        Dim n = arr.Length
        If n <= 1 Then Return 0
        Dim mean = arr.Average(), sumSqDiff = arr.Sum(Function(v) (v - mean) ^ 2)
        Return Math.Sqrt(sumSqDiff / (n - 1))
    End Function

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs)
        ' Validation checks
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
            If count > 0 AndAlso count < 3 Then
                MessageBox.Show($"Please complete all 3 trials for {r} mm.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return
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

        Dim timeInStr = TextBox24.Text.Trim(), timeOutStr = TextBox26.Text.Trim()
        If String.IsNullOrEmpty(timeOutStr) Then timeOutStr = DateTime.Now.ToString("HH:mm") : TextBox26.Text = timeOutStr
        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then MessageBox.Show("Min 2 hours duration required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return

        ' Persist environment data for next instance

        ' Preparation
        Dim cycleName = MySqlCls.GetActiveCycleName(), calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim obsDict As New Dictionary(Of Decimal, Object())
        For Each r In nominalRanges
            obsDict.Add(r, GetTextBoxesForRange(r).Select(Function(t) If(String.IsNullOrEmpty(t.Text), "-", t.Text)).Cast(Of Object)().ToArray())
        Next

        Dim anvil = GetTextBoxesForGeometric("Anvil").Select(Function(t) DirectCast(t.Text, Object)).ToArray()
        Dim spindle = GetTextBoxesForGeometric("Spindle").Select(Function(t) DirectCast(t.Text, Object)).ToArray()
        Dim parallel = GetTextBoxesForGeometric("Parallel").Select(Function(t) DirectCast(t.Text, Object)).ToArray()
        Dim zero = GetTextBoxesForGeometric("Zero").Select(Function(t) DirectCast(t.Text, Object)).ToArray()

        Dim statusVal = DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString()


        Dim dateStr = calibDate.ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, TextBox20.Text)

        Dim saved As Boolean = False
        If IsEditMode Then
            saved = MySqlCls.UpdateDepthMicrometer150(
                "Depth Micrometer", TextBox1.Text.Trim(), TargetCycle, TextBox21.Text, ComboBox5.Text, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                anvil, spindle, parallel, zero, obsDict,
                If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text), TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text, calibDate)
        Else
            saved = MySqlCls.InsertDepthMicrometer150(
                "Depth Micrometer", TextBox1.Text.Trim(), cycleName, TextBox21.Text, ComboBox5.Text, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                anvil, spindle, parallel, zero, obsDict,
                If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text), TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text, calibDate)
        End If

        If saved Then
            Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)
            MySqlCls.UpdateRegularCalibrationStatus(TextBox1.Text.Trim(), calibDate, statusVal, targetCyc)
            MySqlCls.InsertResultRecord(TextBox1.Text.Trim(), _category, _instrumentName, targetCyc)

            ' Save masters to cache
            CalibrationMasterCache.SaveMastersDual(TextBox1.Text.Trim(), targetCyc, _selectedMastersObservation, _selectedMastersGeometric)
            If statusVal = "NG" Then
                ProcessNGFlow(TextBox1.Text.Trim(), targetCyc, calibDate, statusVal)
            End If
            CalibrationMasterCache.SaveMastersDual(TextBox1.Text.Trim(), targetCyc, _selectedMastersObservation, _selectedMastersGeometric)
            MessageBox.Show("Record Saved Successfully!") : Me.Close()
        Else
            MessageBox.Show("Failed to save. " & MySqlCls.LastError)
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
            _isPopulating = False
            EnvironmentTextChanged(TextBox30, Nothing)
            EnvironmentTextChanged(TextBox28, Nothing)
            
            For Each gf In geometricFactors
                For Each tb In GetTextBoxesForGeometric(gf)
                    GeometricTextChanged(tb, Nothing)
                Next
            Next
            
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

