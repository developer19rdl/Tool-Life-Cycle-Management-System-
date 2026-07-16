Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class Dialtestindicator100
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = "Dialtestindicator100"
    Private _category As String = "Instrument"
    Private _selectedMastersObservation As New List(Of MasterSelectorItem)

    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _isPopulating As Boolean = False

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    Private ReadOnly observationParameters As String() = {"Obs_10_Scale_Div", "Obs_Over_1_Rev", "Obs_Over_Meas_Range", "Obs_Hysteresis", "Obs_Repeatability"}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = _instrumentName

        ' Add handlers for observations
        For Each param In observationParameters
            For i = 1 To 3
                Dim tb = DirectCast(Me.FindName($"txtObs_{param}_{i}"), TextBox)
                If tb IsNot Nothing Then
                    AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                    AddHandler tb.TextChanged, AddressOf TrialTextChanged
                    AddHandler tb.LostFocus, AddressOf ObservationLostFocus
                End If
            Next
        Next

        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            Try
                LoadInstrumentDetails(TextBox1.Text.Trim())
            Catch ex As Exception
            End Try
        End If

        If Not IsEditMode Then
            TxtCalibrationDate.SelectedDate = DateTime.Today
        End If

        ' Add handlers for Duration
        AddHandler TextBox24.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.GotFocus, Sub() _isTimeOutManual = True
        AddHandler TextBox26.PreviewKeyDown, Sub() _isTimeOutManual = True
        
        ' Add handlers for limit inputs
        AddHandler TextBox21.TextChanged, AddressOf TextBox21_TextChanged
        AddHandler ComboBox5.SelectionChanged, AddressOf FetchMasterLimitsCombo
        AddHandler ComboBoxRevolution.SelectionChanged, AddressOf FetchMasterLimitsCombo
        AddHandler ComboBoxProbeSize.SelectionChanged, AddressOf FetchMasterLimitsCombo

        _timerRealTime = New System.Windows.Threading.DispatcherTimer()
        _timerRealTime.Interval = TimeSpan.FromSeconds(5)
        AddHandler _timerRealTime.Tick, AddressOf TimerRealTime_Tick
        _timerRealTime.Start()

        If IsEditMode Then
            Try
                PopulateCalibrationData()
            Catch ex As Exception
            End Try
        End If

        ' Add handlers for Environmental conditions
        AddHandler TextBox30.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox28.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox20.TextChanged, AddressOf EnvironmentTextChanged
    End Sub

    Private Sub TxtCalibrationDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles TxtCalibrationDate.SelectedDateChanged
        If TxtCalibrationDate.SelectedDate.HasValue Then
            Dim selectedDateStr = TxtCalibrationDate.SelectedDate.Value.ToShortDateString()
            Dim env = EnvironmentCache.LoadEnvironment(Me.GetType().Name, selectedDateStr)
            If env IsNot Nothing Then
                TextBox30.Text = env.Temperature
                TextBox28.Text = env.Humidity
                TextBox24.Text = env.TimeIn
            Else
                TextBox30.Text = "" : TextBox28.Text = "" : TextBox24.Text = ""
            End If
        End If
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As TextChangedEventArgs)
        FetchMasterLimits()
    End Sub

    Private Sub FetchMasterLimitsCombo(sender As Object, e As SelectionChangedEventArgs)
        FetchMasterLimits()
    End Sub

    Private Sub FetchMasterLimits()
        If _isPopulating Then Return
        Dim sizeText = TextBox21.Text
        Dim selectedLC As String = ""
        Dim selectedRev As String = ""
        Dim selectedProbe As String = ""

        If ComboBox5.SelectedItem IsNot Nothing Then
            selectedLC = DirectCast(ComboBox5.SelectedItem, ComboBoxItem).Content.ToString()
        End If
        If ComboBoxRevolution.SelectedItem IsNot Nothing Then
            selectedRev = DirectCast(ComboBoxRevolution.SelectedItem, ComboBoxItem).Content.ToString()
        End If
        If ComboBoxProbeSize.SelectedItem IsNot Nothing Then
            selectedProbe = DirectCast(ComboBoxProbeSize.SelectedItem, ComboBoxItem).Content.ToString()
        End If

        ' Dynamically enable/disable ProbeSize combobox based on size
        Dim sizeVal As Decimal = 0
        Dim cleanSize = sizeText.ToLower().Replace("mm", "").Replace(" ", "").Trim()
        If cleanSize.Contains("-") Then
            Dim parts = cleanSize.Split("-"c)
            If parts.Length > 1 Then Decimal.TryParse(parts(1), sizeVal)
        Else
            Decimal.TryParse(cleanSize, sizeVal)
        End If

        If sizeVal > 0.5D AndAlso selectedLC = "0.01" AndAlso selectedRev = "One Revolution" Then
            ComboBoxProbeSize.IsEnabled = True
            If ComboBoxProbeSize.SelectedIndex = -1 OrElse ComboBoxProbeSize.Text = "N/A" Then
                ComboBoxProbeSize.Text = "L1 <= 35"
                selectedProbe = "L1 <= 35"
            End If
        Else
            ComboBoxProbeSize.IsEnabled = False
            ComboBoxProbeSize.Text = "N/A"
            selectedProbe = "N/A"
        End If

        If Not String.IsNullOrWhiteSpace(sizeText) AndAlso Not String.IsNullOrWhiteSpace(selectedLC) AndAlso Not String.IsNullOrWhiteSpace(selectedRev) Then
            Dim mappedSize = MapSizeToMasterRange(sizeText, selectedLC, selectedRev)
            _masterLimits = MySqlCls.GetDialtestindicator100MasterLimits(selectedLC, mappedSize, selectedProbe)
            TriggerAllValidations()
        End If
    End Sub

    Private Function MapSizeToMasterRange(sizeText As String, lc As String, revolution As String) As String
        Dim numericSize As Decimal = 0
        Dim cleanSize = sizeText.ToLower().Replace("mm", "").Replace(" ", "").Trim()
        
        If cleanSize.Contains("-") Then
            Dim parts = cleanSize.Split("-"c)
            If parts.Length > 1 Then Decimal.TryParse(parts(1), numericSize)
        Else
            Decimal.TryParse(cleanSize, numericSize)
        End If

        ' Scale Interval: 0.001 or 0.002
        If lc = "0.001" OrElse lc = "0.002" Then
            If revolution = "One Revolution" Then
                Return "0.3 and under"
            Else ' Greater than one Revolution
                If numericSize <= 0.5D Then
                    Return "Over 0.3 upto and inc 0.5"
                Else
                    Return "Over 0.5 upto and inc 0.6"
                End If
            End If
        End If

        ' Scale Interval: 0.01
        If lc = "0.01" Then
            If revolution = "One Revolution" Then
                If numericSize <= 0.5D Then
                    Return "0.5 and Under"
                Else
                    Return "Over 0.5 to 1.0"
                End If
            Else ' Greater than one Revolution
                Return "Over 1.0 to 1.6"
            End If
        End If

        Return "0.3 and under"
    End Function

    Private Sub LoadInstrumentDetails(controlNo As String)
        Try
            Dim dtSettings = MySqlCls.ReadDatatable("SELECT TypeName, Category FROM type_details")
            For Each r As DataRow In dtSettings.Rows
                Dim tbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString())
                Dim query = $"SELECT * FROM `{tbl}` WHERE ControlNo = '{controlNo.Replace("'", "''")}' LIMIT 1"
                Dim dt = MySqlCls.ReadDatatable(query)
                If dt.Rows.Count > 0 Then
                    Dim row = dt.Rows(0)
                    Dim cycleName = MySqlCls.GetActiveCycleName()
                    Dim dtDept = MySqlCls.ReadDatatable($"SELECT Department, Color, SizeandRange FROM department_list WHERE `Control No` = '{controlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")

                    _category = r("Category").ToString()
                    Dim prefix = MySQLClass.ExtractPrefix(controlNo)
                    Dim sizes = MySqlCls.GetSizesForGroup(prefix, _category)

                    Dim matchedSize As String = ""
                    If row.Table.Columns.Contains("Size") AndAlso Not IsDBNull(row("Size")) Then
                        matchedSize = row("Size").ToString()
                    End If

                    If Not String.IsNullOrEmpty(matchedSize) Then
                        TextBox21.Text = matchedSize
                    ElseIf sizes.Count = 1 Then
                        TextBox21.Text = sizes(0)
                    ElseIf sizes.Count > 0 Then
                        TextBox21.Text = sizes(0)
                    End If

                    If dtDept.Rows.Count > 0 Then
                        Dim deptRow = dtDept.Rows(0)
                        ComboBox2.Text = deptRow("Department").ToString()
                        TextBoxColor.Text = deptRow("Color").ToString()
                    Else
                        ComboBox2.Text = row("Location").ToString()
                        TextBoxColor.Text = row("Color").ToString()
                    End If

                    _category = r("Category").ToString()
                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), row("InstrumentName").ToString(), row("GaugeName").ToString())

                    FetchMasterLimits()
                    Exit For
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("LoadInstrumentDetails Error: " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse String.IsNullOrWhiteSpace(TargetCycle) Then Return

        Try
            Dim dt = MySqlCls.GetDialtestindicator100Data(TextBox1.Text.Trim(), TargetCycle)
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
                TextBox19.Text = dr("Remark").ToString()

                ComboBoxRevolution.Text = dr("Type").ToString()
                ComboBoxProbeSize.Text = dr("ProbeSize_L1").ToString()

                If Not IsDBNull(dr("Date")) Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(dr("Date"))
                TextBox24.Text = dr("TimeIn").ToString()
                TextBox26.Text = dr("TimeOut").ToString()
                TextBox25.Text = dr("TotalTime").ToString()
                
                If dr("Status").ToString() = "NG" Then
                    ComboBox4.SelectedIndex = 1
                Else
                    ComboBox4.SelectedIndex = 0
                End If

                ' Observations
                For Each param In observationParameters
                    For i = 1 To 3
                        Dim tb = DirectCast(Me.FindName($"txtObs_{param}_{i}"), TextBox)
                        If tb IsNot Nothing Then tb.Text = If(IsDBNull(dr($"{param}_{i}")), "", dr($"{param}_{i}").ToString())
                    Next
                Next

                ' Restore cached calibration masters
                Dim cached = CalibrationMasterCache.LoadMastersDual(TextBox1.Text.Trim(), TargetCycle)
                If cached.Observation.Count > 0 Then
                    _selectedMastersObservation = MySqlCls.GetMastersByDescriptions(cached.Observation)
                    UpdateMasterUIObservation()
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

    Private Sub TriggerAllValidations()
        Try
            ' Observations
            For Each param In observationParameters
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txtObs_{param}_{i}"), TextBox)
                    If tb IsNot Nothing Then TrialTextChanged(tb, Nothing)
                Next
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ObservationKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)
            Dim parts = tb.Name.Split("_"c)
            If parts.Length >= 3 Then
                Dim trialNumStr = parts(parts.Length - 1)
                Dim trialNum As Integer
                If Integer.TryParse(trialNumStr, trialNum) Then
                    Dim paramParts = parts.Skip(1).Take(parts.Length - 2)
                    Dim param = String.Join("_", paramParts)
                    If trialNum < 3 Then
                        Dim nextTb = DirectCast(Me.FindName($"txtObs_{param}_{trialNum + 1}"), TextBox)
                        If nextTb IsNot Nothing Then nextTb.Focus()
                    Else
                        Dim ptIdx = Array.IndexOf(observationParameters, param)
                        If ptIdx < observationParameters.Length - 1 Then
                            Dim nextTb = DirectCast(Me.FindName($"txtObs_{observationParameters(ptIdx + 1)}_1"), TextBox)
                            If nextTb IsNot Nothing Then nextTb.Focus()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObservationLostFocus(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        If Not String.IsNullOrWhiteSpace(tb.Text) Then
            Dim val As Decimal
            If Decimal.TryParse(tb.Text, val) Then
                tb.Text = val.ToString("0.0000")
            End If
        End If
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        Try
            CalculateTMU()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If
            
            Dim parts = txtBox.Name.Split("_"c)
            If parts.Length >= 3 AndAlso _masterLimits IsNot Nothing Then
                Dim paramParts = parts.Skip(1).Take(parts.Length - 2)
                Dim param = String.Join("_", paramParts)
                
                Dim ulCol = $"{param}_UL"
                Dim llCol = $"{param}_LL"
                
                If _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    If _masterLimits(ulCol) IsNot DBNull.Value AndAlso _masterLimits(llCol) IsNot DBNull.Value Then
                        Dim ul = Convert.ToDecimal(_masterLimits(ulCol))
                        Dim ll = Convert.ToDecimal(_masterLimits(llCol))
                        Dim val As Decimal
                        If Decimal.TryParse(txtBox.Text, val) Then
                            If val > ul OrElse val < ll Then
                                txtBox.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226))
                                txtBox.Foreground = Brushes.Red
                            Else
                                txtBox.Background = Brushes.White
                                txtBox.Foreground = Brushes.Black
                            End If
                        End If
                    Else
                        txtBox.Background = Brushes.White
                        txtBox.Foreground = Brushes.Black
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
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

    Private Sub UpdateMasterUIObservation()
        ItemsControlSelectedMastersObservation.ItemsSource = Nothing
        ItemsControlSelectedMastersObservation.ItemsSource = _selectedMastersObservation
        MasterGroupBoxObservation.Visibility = If(_selectedMastersObservation.Count > 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub TimerRealTime_Tick(sender As Object, e As EventArgs)
        If Not _isTimeOutManual AndAlso Not String.IsNullOrWhiteSpace(TextBox24.Text) Then
            TextBox26.Text = DateTime.Now.ToString("HH:mm")
        End If
    End Sub

    Private Sub CalculateTMU()
        If _isPopulating OrElse TextBox29 Is Nothing Then Return
        Try
            Dim masterU_Observation As Decimal = 0D
            If _selectedMastersObservation.Count > 0 Then
                ' RSS-combine all selected masters: u = sqrt(sum(ui^2))
                masterU_Observation = Math.Sqrt(_selectedMastersObservation.Sum(Function(m) m.MasterUncertainty ^ 2))
            Else
                masterU_Observation = 0.00012D ' Fallback
            End If

            Dim totalSumSq = 0D

            ' Observations Repeatability
            For Each param In observationParameters
                Dim vals As New List(Of Decimal)
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txtObs_{param}_{i}"), TextBox)
                    If tb IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
                        Dim val As Decimal
                        If Decimal.TryParse(tb.Text, val) Then vals.Add(val)
                    End If
                Next
                If vals.Count > 0 Then
                    Dim sigmaB1 = calcStdev(vals.ToArray()) / Math.Sqrt(3)
                    totalSumSq += (sigmaB1 ^ 2) + (masterU_Observation ^ 2)
                End If
            Next

            ' Corrected: Use coverage factor k=2 instead of 1.2D
            TextBox29.Text = (2 * Math.Sqrt(totalSumSq)).ToString("F9")
            TextBox29.Background = Brushes.White
        Catch ex As Exception
            TextBox29.Text = "Error"
        End Try
    End Sub

    Private Function calcStdev(arr() As Decimal) As Decimal
        If arr.Length <= 1 Then Return 0
        Dim avg = arr.Average()
        Dim sumSq = arr.Sum(Function(x) (x - avg) ^ 2)
        Return Math.Sqrt(sumSq / (arr.Length - 1))
    End Function

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MessageBox.Show("Control No cannot be empty.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Duration Validation - must be at least 2 hours
        Dim timeInStr = TextBox24.Text.Trim()
        Dim timeOutStr = TextBox26.Text.Trim()
        If String.IsNullOrEmpty(timeOutStr) Then timeOutStr = DateTime.Now.ToString("HH:mm")
        Dim durationCheck = CalculateDuration(timeInStr, timeOutStr)
        If durationCheck >= 0 AndAlso durationCheck < 120 Then
            MessageBox.Show("Calibration duration must be at least 2 hours.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim controlNo = TextBox1.Text.Trim()
        Dim cycleName = MySqlCls.GetActiveCycleName()
        Dim size = TextBox21.Text
        Dim lc = ComboBox5.Text
        Dim probeSizeL1 = ComboBoxProbeSize.Text
        Dim color = TextBoxColor.Text
        Dim location = ComboBox2.Text
        Dim temp = TextBox30.Text
        Dim humidity = TextBox28.Text
        Dim tmu = TextBox29.Text
        Dim remark = TextBox19.Text
        Dim timeIn = TextBox24.Text
        Dim timeOut = TextBox26.Text
        Dim totalTime = TextBox25.Text
        Dim status = If(ComboBox4.SelectedItem IsNot Nothing, DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString(), "PASS")
        Dim type = ComboBoxRevolution.Text

        Dim dateStr = If(TxtCalibrationDate.SelectedDate, DateTime.Today).ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, "")

        Dim obs_10_Scale_Div = {DirectCast(Me.FindName("txtObs_Obs_10_Scale_Div_1"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_10_Scale_Div_2"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_10_Scale_Div_3"), TextBox).Text}
        Dim obs_Over_1_Rev = {DirectCast(Me.FindName("txtObs_Obs_Over_1_Rev_1"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Over_1_Rev_2"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Over_1_Rev_3"), TextBox).Text}
        Dim obs_Over_Meas_Range = {DirectCast(Me.FindName("txtObs_Obs_Over_Meas_Range_1"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Over_Meas_Range_2"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Over_Meas_Range_3"), TextBox).Text}
        Dim obs_Hysteresis = {DirectCast(Me.FindName("txtObs_Obs_Hysteresis_1"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Hysteresis_2"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Hysteresis_3"), TextBox).Text}
        Dim obs_Repeatability = {DirectCast(Me.FindName("txtObs_Obs_Repeatability_1"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Repeatability_2"), TextBox).Text, DirectCast(Me.FindName("txtObs_Obs_Repeatability_3"), TextBox).Text}

        Dim success As Boolean = False
        If IsEditMode Then
            success = MySqlCls.UpdateDialtestindicator100(type, controlNo, TargetCycle, size, lc, probeSizeL1, color, location, temp, humidity, tmu, obs_10_Scale_Div, obs_Over_1_Rev, obs_Over_Meas_Range, obs_Hysteresis, obs_Repeatability, timeIn, timeOut, totalTime, status, remark, TxtCalibrationDate.SelectedDate)
        Else
            success = MySqlCls.InsertDialtestindicator100(type, controlNo, cycleName, TxtCalibrationDate.SelectedDate, size, lc, probeSizeL1, color, location, temp, humidity, tmu, obs_10_Scale_Div, obs_Over_1_Rev, obs_Over_Meas_Range, obs_Hysteresis, obs_Repeatability, timeIn, timeOut, totalTime, status, remark)
        End If

        If success Then
            Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
            Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)
            MySqlCls.UpdateRegularCalibrationStatus(controlNo, calibDate, status, targetCyc)
            MySqlCls.InsertResultRecord(controlNo, _category, _instrumentName, targetCyc)

            CalibrationMasterCache.SaveMastersDual(controlNo, targetCyc, _selectedMastersObservation, New List(Of MasterSelectorItem)())

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
            If totalMinutes < 120 Then
                TextBox25.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' Light Pink
            Else
                TextBox25.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' Light Green
            End If
        Else
            TextBox25.Text = ""
            TextBox25.Background = Brushes.White
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
        If _isPopulating OrElse _masterLimits Is Nothing Then Return
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
    End Sub
End Class

