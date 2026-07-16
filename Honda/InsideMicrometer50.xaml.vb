Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class InsideMicrometer50
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = ""
    Private _category As String = "Instrument"
    Private _selectedMasters As New List(Of MasterSelectorItem)



    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _activeMinRange As Decimal = 25
    Private _activeMaxRange As Decimal = 50

    Private ReadOnly nominalRanges As Decimal() = {27.5, 30.1, 32.7, 35.3, 37.9, 40.0, 42.6, 45.2, 47.8, 50.0}

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = "Inside Micrometer"
        LoadDynamicLC()
        LoadCalibrationMasters()

        ' Add handlers for all trial textboxes
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                tb.IsEnabled = False
                AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                AddHandler tb.TextChanged, AddressOf TrialTextChanged
                AddHandler tb.LostFocus, AddressOf ObservationLostFocus
            Next
        Next

        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            LoadInstrumentDetails(TextBox1.Text.Trim())
        End If

        If Not IsEditMode Then TxtCalibrationDate.SelectedDate = DateTime.Today

        ' Add handlers for Duration calculation
        AddHandler TextBox24.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.GotFocus, Sub() _isTimeOutManual = True
        AddHandler TextBox26.PreviewKeyDown, Sub() _isTimeOutManual = True

        ' Setup Real-time timer
        _timerRealTime = New System.Windows.Threading.DispatcherTimer()
        _timerRealTime.Interval = TimeSpan.FromSeconds(5)
        AddHandler _timerRealTime.Tick, AddressOf TimerRealTime_Tick
        _timerRealTime.Start()

        ' Initialize LC to none and disable all observations initially
        ComboBox5.SelectedIndex = -1
        DisableAllObservationFields()
        ResetAllTrials()
        UpdateSizeLockState()

        If IsEditMode Then
            PopulateCalibrationData()
        End If

        ' Add handlers for Environmental conditions
        AddHandler TextBox30.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox28.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox20.TextChanged, AddressOf EnvironmentTextChanged
    End Sub

    Private Sub PopulateCalibrationData()
        Try
            Dim dt = MySqlCls.GetInsideMicrometer50Data(TextBox1.Text.Trim(), TargetCycle)
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                
                TxtCalibrationDate.SelectedDate = Convert.ToDateTime(row("Date"))
                TextBox21.Text = row("Size").ToString()
                ComboBox5.Text = row("LC").ToString()
                TextBoxColor.Text = row("Color").ToString()
                ComboBox2.Text = row("Location").ToString()
                TextBox30.Text = row("Temperature").ToString()
                TextBox28.Text = row("Humidity").ToString()
                
                SetTrialsFromDB(row, "Obs27_5_", GetTextBoxesForRange(27.5))
                SetTrialsFromDB(row, "Obs30_1_", GetTextBoxesForRange(30.1))
                SetTrialsFromDB(row, "Obs32_7_", GetTextBoxesForRange(32.7))
                SetTrialsFromDB(row, "Obs35_3_", GetTextBoxesForRange(35.3))
                SetTrialsFromDB(row, "Obs37_9_", GetTextBoxesForRange(37.9))
                SetTrialsFromDB(row, "Obs40_0_", GetTextBoxesForRange(40.0))
                SetTrialsFromDB(row, "Obs42_6_", GetTextBoxesForRange(42.6))
                SetTrialsFromDB(row, "Obs45_2_", GetTextBoxesForRange(45.2))
                SetTrialsFromDB(row, "Obs47_8_", GetTextBoxesForRange(47.8))
                SetTrialsFromDB(row, "Obs50_0_", GetTextBoxesForRange(50.0))

                TextBox20.Text = If(row("DepthError").ToString() = "-", "", row("DepthError").ToString())
                TextBox24.Text = row("TimeIn").ToString()
                TextBox26.Text = row("TimeOut").ToString()
                TextBox25.Text = row("TotalTime").ToString()
                TextBox19.Text = row("Remark").ToString()

                Dim status = row("Status").ToString()
                ComboBox4.Text = status

                ' Calculate TMU but override with saved value
                CalculateTMU()
                TriggerAllValidations()
                TextBox29.Text = row("TMU").ToString()
            End If
        Catch ex As Exception
            Console.WriteLine("PopulateCalibrationData Error: " & ex.Message)
        End Try
    End Sub

    Private Sub TriggerAllValidations()
        Try
            ' Environment
            EnvTextChanged(TextBox30, Nothing)
            EnvTextChanged(TextBox28, Nothing)
            EnvTextChanged(TextBox20, Nothing) ' Depth Error
            
            ' Observations
            For Each r In nominalRanges
                For Each tb In GetTextBoxesForRange(r)
                    TrialTextChanged(tb, Nothing)
                Next
            Next
            
            UpdateOverallStatus()
        Catch ex As Exception
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
        Return nominal > _activeMinRange AndAlso nominal <= _activeMaxRange
    End Function

    Private Sub AdjustVisibilityBySize(sizeStr As String)
        Try
            _activeMinRange = 25 : _activeMaxRange = 50
            If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
                Dim parts = sizeStr.Split("-"c)
                If parts.Length > 1 Then
                    Decimal.TryParse(parts(0).Replace("mm", "").Trim(), _activeMinRange)
                    Decimal.TryParse(parts(1).Replace("mm", "").Trim(), _activeMaxRange)
                End If
            End If

            For Each r In nominalRanges
                Dim inRange = IsNominalInRange(r)
                Dim tbs = GetTextBoxesForRange(r)
                Dim nomStr = r.ToString("0.0").Replace(".", "_")
                Dim lblName = $"Label{nomStr}"
                Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)

                For Each tb In tbs
                    tb.IsEnabled = inRange
                    tb.Background = If(inRange, Brushes.White, New SolidColorBrush(Color.FromRgb(241, 245, 249)))
                    If Not inRange Then tb.Text = ""
                Next

                If lbl IsNot Nothing Then
                    lbl.Foreground = If(inRange, New SolidColorBrush(Color.FromRgb(51, 65, 85)), Brushes.LightSlateGray)
                End If
            Next
            UpdateSizeLockState()
        Catch ex As Exception
        End Try
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

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        If _masterLimits Is Nothing Then
            txtBox.Background = Brushes.White
            Return
        End If

        Dim nominal As Decimal = GetNominalFromTextBox(txtBox)
        If nominal = -1 Then Return

        Dim prefix As String = "Obs"
        Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
        Dim ulCol As String = $"{prefix}{nomStr}_UL"
        Dim llCol As String = $"{prefix}{nomStr}_LL"

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
                ' For Inside Micrometer, use the raw value as the error comparison
                Dim errorVal As Decimal = val

                If _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    Dim ul As Decimal = Convert.ToDecimal(_masterLimits(ulCol))
                    Dim ll As Decimal = Convert.ToDecimal(_masterLimits(llCol))

                    If errorVal >= ll AndAlso errorVal <= ul Then
                        txtBox.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
                    Else
                        txtBox.Background = New SolidColorBrush(Color.FromRgb(255, 182, 193))
                    End If
                Else
                    txtBox.Background = Brushes.White
                End If
                UpdateOverallStatus()
            Else
                txtBox.Background = Brushes.White
                UpdateOverallStatus()
            End If
        Catch ex As Exception
            txtBox.Background = Brushes.White
            UpdateOverallStatus()
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

    Private Sub TxtCalibrationDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles TxtCalibrationDate.SelectedDateChanged
        If TxtCalibrationDate.SelectedDate.HasValue Then
            Dim selectedDateStr = TxtCalibrationDate.SelectedDate.Value.ToShortDateString()
            Dim env = EnvironmentCache.LoadEnvironment(Me.GetType().Name, selectedDateStr)
            If env IsNot Nothing Then
                TextBox30.Text = env.Temperature
                TextBox28.Text = env.Humidity
                TextBox20.Text = env.DepthError
                TextBox24.Text = env.TimeIn
            Else
                TextBox30.Text = "" : TextBox28.Text = "" : TextBox20.Text = "" : TextBox24.Text = ""
            End If
        End If
    End Sub

    Private Sub DisableAllObservationFields()
        For Each r In nominalRanges
            For Each tb In GetTextBoxesForRange(r)
                tb.IsEnabled = False
                tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
        Next
    End Sub

    Private Sub ResetAllTrials()
        For Each r In nominalRanges
            For Each tb In GetTextBoxesForRange(r)
                tb.Clear()
            Next
            UpdateObservationLabel(r)
        Next
        AdjustVisibilityBySize(TextBox21.Text)
        CalculateTMU()
        UpdateSizeLockState()
    End Sub

    Private Function GetTextBoxesForRange(nominal As Decimal) As List(Of TextBox)
        Dim list As New List(Of TextBox)
        Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
        For i = 1 To 3
            Dim name = $"txt{nomStr}_{i}"
            Dim tb = DirectCast(Me.FindName(name), TextBox)
            if tb IsNot Nothing Then list.Add(tb)
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

    Private Sub UpdateObservationLabel(nominal As Decimal)
        Dim tbs = GetTextBoxesForRange(nominal)
        Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
        Dim nomStr = nominal.ToString("0.0").Replace(".", "_")
        Dim lblName = $"Label{nomStr}"
        Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)

        If lbl IsNot Nothing Then
            lbl.Text = $"{nominal.ToString("0.000")} mm:"
            If ComboBox5.SelectedItem Is Nothing Then
                lbl.Foreground = New SolidColorBrush(Color.FromRgb(51, 65, 85))
            Else
                lbl.Foreground = If(count >= 3, Brushes.Gray, Brushes.DodgerBlue)
            End If
        End If
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As RoutedEventArgs) Handles TextBox1.LostFocus
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            LoadInstrumentDetails(TextBox1.Text.Trim())
        End If
    End Sub

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
                        if matchedSize IsNot Nothing Then
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
                _instrumentName = "" : ComboBox2.Text = "" : TextBoxColor.Text = "" : TextBox21.Text = "25-50 mm" : UpdateSizeLockState()
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
        TextBox21.IsEnabled = Not hasData
    End Sub

    Private Sub LoadCalibrationMasters()
        _selectedMasters.Clear()
    End Sub

    Private Sub BtnBrowseMaster_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMasters.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMasters = selector.SelectedMasters
            UpdateMasterUI() : CalculateTMU()
        End If
    End Sub

    Private Sub UpdateMasterUI()
        ItemsControlSelectedMasters.ItemsSource = Nothing
        ItemsControlSelectedMasters.ItemsSource = _selectedMasters
        MasterGroupBox.Visibility = If(_selectedMasters.Count > 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub TimerRealTime_Tick(sender As Object, e As EventArgs)
        If Not _isTimeOutManual AndAlso Not String.IsNullOrWhiteSpace(TextBox24.Text) Then
            TextBox26.Text = DateTime.Now.ToString("HH:mm")
        End If
    End Sub

    Private Sub LoadDynamicLC()
        Try
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM inside_micrometer_50 WHERE RowType='MASTER' ORDER BY LC ASC")
            ComboBox5.Items.Clear()
            For Each row As DataRow In dt.Rows
                ComboBox5.Items.Add(row("LC").ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("LoadDynamicLC Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ComboBox5_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboBox5.SelectionChanged
        If ComboBox5.SelectedItem IsNot Nothing Then
            Dim selectedLC As String = ComboBox5.SelectedItem.ToString()
            _masterLimits = MySqlCls.GetInsideMicrometer50MasterLimits(selectedLC)
            ResetAllTrials()
            AdjustVisibilityBySize(TextBox21.Text)
            EnableRange(25)
        Else
            _masterLimits = Nothing :  ResetAllTrials() : DisableAllObservationFields()
        End If
    End Sub

    Private Sub EnvTextChanged(sender As Object, e As TextChangedEventArgs) Handles TextBox30.TextChanged, TextBox28.TextChanged, TextBox20.TextChanged
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        If _masterLimits Is Nothing Then Return
        Try
            If String.IsNullOrWhiteSpace(txtBox.Text) Then txtBox.Background = Brushes.White : Return
            Dim val As Decimal
            If Decimal.TryParse(txtBox.Text, val) Then
                If txtBox.Name = "TextBox20" Then
                    txtBox.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
                Else
                    Dim llCol = "", ulCol = ""
                    If txtBox.Name = "TextBox30" Then : llCol = "Env_Temp_LL" : ulCol = "Env_Temp_UL"
                    ElseIf txtBox.Name = "TextBox28" Then : llCol = "Env_Hum_LL" : ulCol = "Env_Hum_UL"
                    End If
                    If _masterLimits.Table.Columns.Contains(llCol) AndAlso _masterLimits.Table.Columns.Contains(ulCol) Then
                        Dim ll = Convert.ToDecimal(_masterLimits(llCol)), ul = Convert.ToDecimal(_masterLimits(ulCol))
                        txtBox.Background = If(val >= ll AndAlso val <= ul, New SolidColorBrush(Color.FromRgb(144, 238, 144)), New SolidColorBrush(Color.FromRgb(255, 182, 193)))
                    End If
                End If
                UpdateOverallStatus()
            Else
                txtBox.Background = Brushes.White : UpdateOverallStatus()
            End If
        Catch ex As Exception
            txtBox.Background = Brushes.White : UpdateOverallStatus()
        End Try
    End Sub

    Private Sub UpdateOverallStatus()
        Dim isNG = False
        For Each r In nominalRanges
            For Each tb In GetTextBoxesForRange(r)
                If tb.Background IsNot Nothing AndAlso TypeOf tb.Background Is SolidColorBrush Then
                    If DirectCast(tb.Background, SolidColorBrush).Color = Color.FromRgb(255, 182, 193) Then
                        isNG = True : Exit For
                    End If
                End If
            Next
            if isNG Then Exit For
        Next

        If Not isNG Then
            For Each tb In {TextBox30, TextBox28, TextBox20}
                If tb.Background IsNot Nothing AndAlso TypeOf tb.Background Is SolidColorBrush Then
                    If DirectCast(tb.Background, SolidColorBrush).Color = Color.FromRgb(255, 182, 193) Then isNG = True
                End If
            Next
        End If
        ComboBox4.SelectedIndex = If(isNG, 1, 0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs) Handles Button2.Click
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
            If count > 0 AndAlso count < 3 Then
                MessageBox.Show($"Please complete all 3 trials for {r}mm.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return
            End If
        Next

        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse ComboBox5.SelectedItem Is Nothing Then
            MessageBox.Show("Control No and LC must be filled.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return
        End If
        If ComboBox4.SelectedItem Is Nothing Then MessageBox.Show("Please select Status.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return

        Dim timeInStr = TextBox24.Text.Trim(), timeOutStr = TextBox26.Text.Trim()
        If String.IsNullOrEmpty(timeOutStr) Then timeOutStr = DateTime.Now.ToString("HH:mm") : TextBox26.Text = timeOutStr
        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then MessageBox.Show("Calibration duration must be at least 2 hours.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning) : Return

        Dim dateStr = If(TxtCalibrationDate.SelectedDate, DateTime.Today).ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, TextBox20.Text)

        Dim statusVal = DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString()
        Dim cycleName = MySqlCls.GetActiveCycleName(), calibDateToSave = If(TxtCalibrationDate.SelectedDate, DateTime.Today)

        Dim obs27_5 = GetTrialsArrayFromUI(nominalRanges(0))
        Dim obs30_1 = GetTrialsArrayFromUI(nominalRanges(1))
        Dim obs32_7 = GetTrialsArrayFromUI(nominalRanges(2))
        Dim obs35_3 = GetTrialsArrayFromUI(nominalRanges(3))
        Dim obs37_9 = GetTrialsArrayFromUI(nominalRanges(4))
        Dim obs40_0 = GetTrialsArrayFromUI(nominalRanges(5))
        Dim obs42_6 = GetTrialsArrayFromUI(nominalRanges(6))
        Dim obs45_2 = GetTrialsArrayFromUI(nominalRanges(7))
        Dim obs47_8 = GetTrialsArrayFromUI(nominalRanges(8))
        Dim obs50_0 = GetTrialsArrayFromUI(nominalRanges(9))

        Dim lcVal = ""
        If TypeOf ComboBox5.SelectedItem Is ComboBoxItem Then
            lcVal = DirectCast(ComboBox5.SelectedItem, ComboBoxItem).Content.ToString()
        Else
            lcVal = ComboBox5.SelectedItem.ToString()
        End If

        Dim saved As Boolean
        If IsEditMode Then
            saved = MySqlCls.UpdateInsideMicrometer50(
                _instrumentName, TextBox1.Text.Trim(), cycleName, TextBox21.Text, lcVal, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                obs27_5, obs30_1, obs32_7, obs35_3, obs37_9, obs40_0, obs42_6, obs45_2, obs47_8, obs50_0,
                If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text), TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text,
                calibDateToSave)
        Else
            saved = MySqlCls.InsertInsideMicrometer50(
                _instrumentName, TextBox1.Text.Trim(), cycleName, TextBox21.Text, lcVal, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                obs27_5, obs30_1, obs32_7, obs35_3, obs37_9, obs40_0, obs42_6, obs45_2, obs47_8, obs50_0,
                If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text), TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text,
                calibDateToSave)
        End If

        If saved Then
            MySqlCls.UpdateRegularCalibrationStatus(TextBox1.Text.Trim(), calibDateToSave, statusVal, cycleName)
            If Not String.IsNullOrEmpty(cycleName) Then MySqlCls.InsertResultRecord(TextBox1.Text.Trim(), _category, _instrumentName, cycleName)

            If statusVal = "NG" Then
                Dim ctrlNo = TextBox1.Text.Trim()
                Dim dtDept = MySqlCls.ReadDatatable($"SELECT * FROM department_list WHERE `Control No` = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")
                If dtDept.Rows.Count > 0 Then
                    Dim row = dtDept.Rows(0)
                    Dim dept = row("Department").ToString()
                    Dim instName = row("InstrumentName").ToString()
                    Dim color = row("Color").ToString()
                    Dim size = row("SizeandRange").ToString()

                    Dim ngReason = If(Not String.IsNullOrWhiteSpace(TextBox19.Text), TextBox19.Text.Trim(), "Calibration NG")
                    Dim wopReason = "Calibration NG"
                    If Not String.IsNullOrWhiteSpace(TextBox19.Text) Then
                        wopReason &= ": " & TextBox19.Text.Trim()
                    End If

                    Dim dueDate = calibDateToSave.AddYears(1).AddDays(-1)
                    MySqlCls.InsertNGRecord(ctrlNo, instName, dept, cycleName, ngReason, calibDateToSave, dueDate, statusVal)

                    MySqlCls.InsertWOPRecord(ctrlNo, cycleName, dept, instName, color, wopReason)
                    MySqlCls.InsertInterchangeRecord(cycleName, ctrlNo, dept, instName, size, color, "WOP", wopReason)
                    MySqlCls.ClearRFIDTag(ctrlNo)
                End If
            End If

            MessageBox.Show("Record Saved Successfully!") : Me.Close()
        Else
            MessageBox.Show("Failed to save record. " & MySqlCls.LastError)
        End If
    End Sub

    Private Function GetTrialsArrayFromUI(nominal As Decimal) As Object()
        Dim tbs = GetTextBoxesForRange(nominal), trials(2) As Object
        For i = 0 To 2
            If i < tbs.Count Then
                Dim val As Decimal = 0D
                If Decimal.TryParse(tbs(i).Text, val) Then
                    ' Save the raw value as the error
                    trials(i) = val.ToString("0.000")
                Else
                    trials(i) = "-"
                End If
            Else
                trials(i) = "-"
            End If
        Next
        Return trials
    End Function

    Private Sub CalculateTMU()
        Try
            If _selectedMasters Is Nothing OrElse _selectedMasters.Count = 0 Then
                TextBox29.Text = ""
                TextBox29.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                Return
            End If
            TextBox29.Background = Brushes.White
            Dim sumSqMaster = _selectedMasters.Sum(Function(m) m.MasterUncertainty ^ 2)
            Dim masterU = Math.Sqrt(sumSqMaster), sumSqMU = 0D
            For Each r In nominalRanges
                Dim vals = GetTextBoxesForRange(r).Select(Function(t) If(Decimal.TryParse(t.Text, Nothing), Decimal.Parse(t.Text), 0D)).ToArray()
                sumSqMU += (calcStdev(vals) / Math.Sqrt(3)) ^ 2 + (masterU ^ 2)
            Next
            TextBox29.Text = (Math.Sqrt(sumSqMU) * 2).ToString("F9")
        Catch ex As Exception
        End Try
    End Sub

    Private Function calcStdev(arr() As Decimal) As Decimal
        Dim n = arr.Length
        If n <= 1 Then Return 0
        Dim mean = arr.Average(), sumSqDiff = arr.Sum(Function(v) (v - mean) ^ 2)
        Return Math.Sqrt(sumSqDiff / (n - 1))
    End Function

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

    Private Sub EnvironmentTextChanged(sender As Object, e As TextChangedEventArgs)
        If _masterLimits Is Nothing Then Return
        Dim tb = DirectCast(sender, TextBox)
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            Dim isOK = True
            If tb.Name = "TextBox30" Then ' Temp
                If _masterLimits.Table.Columns.Contains("Env_Temp_LL") AndAlso _masterLimits.Table.Columns.Contains("Env_Temp_UL") Then
                    Dim ll = Convert.ToDecimal(_masterLimits("Env_Temp_LL"))
                    Dim ul = Convert.ToDecimal(_masterLimits("Env_Temp_UL"))
                    If val < ll OrElse val > ul Then isOK = False
                End If
            ElseIf tb.Name = "TextBox28" Then ' Humidity
                If _masterLimits.Table.Columns.Contains("Env_Hum_LL") AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_UL") Then
                    Dim ll = Convert.ToDecimal(_masterLimits("Env_Hum_LL"))
                    Dim ul = Convert.ToDecimal(_masterLimits("Env_Hum_UL"))
                    If val < ll OrElse val > ul Then isOK = False
                End If
            ElseIf tb.Name = "TextBox20" Then ' Depth Error
                If _masterLimits.Table.Columns.Contains("Depth_LL") AndAlso _masterLimits.Table.Columns.Contains("Depth_UL") Then
                    Dim ll = Convert.ToDecimal(_masterLimits("Depth_LL"))
                    Dim ul = Convert.ToDecimal(_masterLimits("Depth_UL"))
                    If val < ll OrElse val > ul Then isOK = False
                End If
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
End Class



