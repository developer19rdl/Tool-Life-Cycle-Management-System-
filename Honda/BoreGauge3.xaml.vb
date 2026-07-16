Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Linq

Public Class BoreGauge3
    Inherits Window

    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = "Bore Gauge 3"
    Private _category As String = "Instrument"
    Private _selectedMastersObservation As New List(Of MasterSelectorItem)

    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _isPopulating As Boolean = False

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    Private ReadOnly observationParameters As String() = {"WideRange", "Adjacent", "Repeatability"}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = _instrumentName

        ' Add handlers for observations
        For Each param In observationParameters
            For i = 1 To 3
                Dim tb = DirectCast(Me.FindName($"txt{param}_{i}"), TextBox)
                If tb IsNot Nothing Then
                    AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                    AddHandler tb.TextChanged, AddressOf TrialTextChanged
                    AddHandler tb.LostFocus, AddressOf ObservationLostFocus
                End If
            Next
        Next

        ' Add environmental handlers
        AddHandler TextBox30.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox28.TextChanged, AddressOf EnvironmentTextChanged

        ' Add handlers for Duration
        AddHandler TextBox24.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.GotFocus, Sub() _isTimeOutManual = True
        AddHandler TextBox26.PreviewKeyDown, Sub() _isTimeOutManual = True

        ' Load settings / details
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            LoadInstrumentDetails(TextBox1.Text.Trim())
        End If

        LoadDynamicLC()

        ' Pre-fill last saved values if not in edit mode
        If Not IsEditMode Then
            TxtCalibrationDate.SelectedDate = DateTime.Today
        End If

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
    End Sub

    Private Sub TxtCalibrationDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles TxtCalibrationDate.SelectedDateChanged
        If _isPopulating OrElse IsEditMode Then Return
        If TxtCalibrationDate.SelectedDate.HasValue Then
            Dim selectedDateStr = TxtCalibrationDate.SelectedDate.Value.ToShortDateString()
            Dim env = EnvironmentCache.LoadEnvironment(Me.GetType().Name, selectedDateStr)
            If env IsNot Nothing Then
                If String.IsNullOrWhiteSpace(TextBox30.Text) Then TextBox30.Text = env.Temperature
                If String.IsNullOrWhiteSpace(TextBox28.Text) Then TextBox28.Text = env.Humidity
                If String.IsNullOrWhiteSpace(TextBox24.Text) Then TextBox24.Text = env.TimeIn
            End If
        End If
    End Sub

    Private Sub TextBox21_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles TextBox21.SelectionChanged
        LoadMasterLimits()
    End Sub

    Private Sub ComboBox5_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboBox5.SelectionChanged
        LoadMasterLimits()
    End Sub

    Private Sub LoadMasterLimits()
        If ComboBox5.SelectedItem Is Nothing Then Return
        Dim selectedLC As String = ComboBox5.SelectedItem.ToString()
        _masterLimits = MySqlCls.GetBoreGauge3MasterLimits(selectedLC)
        TriggerAllValidations()
        CalculateTMU()
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse String.IsNullOrWhiteSpace(TargetCycle) Then Return

        Try
            Dim dt = MySqlCls.GetBoreGauge3Data(TextBox1.Text.Trim(), TargetCycle)
            If dt.Rows.Count > 0 Then
                Dim dr = dt.Rows(0)
                _isPopulating = True

                ComboBox2.Text = dr("Location").ToString()
                
                ' Set Size
                Dim sizeVal = dr("Size").ToString()
                TextBox21.Text = sizeVal
                
                ComboBox5.Text = dr("LC").ToString()
                TextBoxColor.Text = dr("Color").ToString()
                TextBox30.Text = dr("Temperature").ToString()
                TextBox28.Text = dr("Humidity").ToString()
                TextBox29.Text = dr("TMU").ToString()
                TextBox20.Text = If(IsDBNull(dr("DepthError")), "", dr("DepthError").ToString())
                TextBox19.Text = dr("Remark").ToString()

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
                    Dim dbParam = param
                    If param = "WideRange" Then
                        dbParam = "Wide_Range"
                    ElseIf param = "Adjacent" Then
                        dbParam = "Adjacent_Error"
                    End If

                    For i = 1 To 3
                        Dim tb = DirectCast(Me.FindName($"txt{param}_{i}"), TextBox)
                        If tb IsNot Nothing Then
                            tb.Text = If(IsDBNull(dr($"{dbParam}_{i}")), "", dr($"{dbParam}_{i}").ToString())
                        End If
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
            ' Environment
            EnvironmentTextChanged(TextBox30, Nothing)
            EnvironmentTextChanged(TextBox28, Nothing)

            ' Observations
            For Each param In observationParameters
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txt{param}_{i}"), TextBox)
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
            If parts.Length = 2 Then
                Dim trialNumStr = parts(1)
                Dim trialNum As Integer
                If Integer.TryParse(trialNumStr, trialNum) Then
                    Dim param = tb.Name.Replace("txt", "").Replace($"_{trialNumStr}", "")
                    If trialNum < 3 Then
                        Dim nextTb = DirectCast(Me.FindName($"txt{param}_{trialNum + 1}"), TextBox)
                        If nextTb IsNot Nothing Then nextTb.Focus()
                    Else
                        Dim ptIdx = Array.IndexOf(observationParameters, param)
                        If ptIdx < observationParameters.Length - 1 Then
                            Dim nextTb = DirectCast(Me.FindName($"txt{observationParameters(ptIdx + 1)}_1"), TextBox)
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
        If _isPopulating Then Return
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        Try
            CalculateTMU()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If
            
            Dim parts = txtBox.Name.Split("_"c)
            If parts.Length = 2 AndAlso _masterLimits IsNot Nothing Then
                Dim trialNumStr = parts(1)
                Dim param = txtBox.Name.Replace("txt", "").Replace($"_{trialNumStr}", "")
                
                Dim dbParam = param
                If param = "WideRange" Then
                    dbParam = "Wide_Range"
                ElseIf param = "Adjacent" Then
                    dbParam = "Adjacent_Error"
                End If
                
                Dim ulCol = $"{dbParam}_UL"
                Dim llCol = $"{dbParam}_LL"
                
                If _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    Dim ul = Convert.ToDecimal(_masterLimits(ulCol))
                    Dim ll = Convert.ToDecimal(_masterLimits(llCol))
                    Dim val As Decimal
                    If Decimal.TryParse(txtBox.Text, val) Then
                        If val > ul OrElse val < ll Then
                            txtBox.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226))
                            txtBox.Foreground = Brushes.Red
                        Else
                            txtBox.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231))
                            txtBox.Foreground = Brushes.Black
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
        UpdateOverallStatus()
    End Sub

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

    Private Sub UpdateOverallStatus()
        If _isPopulating OrElse ComboBox4 Is Nothing Then Return

        Dim hasNG = False

        ' Check Environment
        If IsNGColor(TextBox30.Background) OrElse IsNGColor(TextBox28.Background) Then hasNG = True

        ' Check Observations
        For Each param In observationParameters
            For i = 1 To 3
                Dim tb = DirectCast(Me.FindName($"txt{param}_{i}"), TextBox)
                If tb IsNot Nothing AndAlso IsNGColor(tb.Background) Then
                    hasNG = True : Exit For
                End If
            Next
            If hasNG Then Exit For
        Next

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

    Private Sub BtnBrowseMaster_Click(sender As Object, e As RoutedEventArgs)
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
        ItemsControlSelectedMasters.ItemsSource = Nothing
        ItemsControlSelectedMasters.ItemsSource = _selectedMastersObservation
        MasterGroupBox.Visibility = If(_selectedMastersObservation.Count > 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub TimerRealTime_Tick(sender As Object, e As EventArgs)
        If Not _isTimeOutManual AndAlso Not String.IsNullOrWhiteSpace(TextBox24.Text) Then
            TextBox26.Text = DateTime.Now.ToString("HH:mm")
        End If
    End Sub

    Private Sub LoadDynamicLC()
        Try
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM bore_gauge_3 WHERE RowType='MASTER' ORDER BY LC ASC")
            ComboBox5.Items.Clear()
            For Each row As DataRow In dt.Rows
                ComboBox5.Items.Add(row("LC").ToString())
            Next
            If ComboBox5.Items.Count > 0 Then
                ComboBox5.SelectedIndex = 0
            Else
                ComboBox5.Items.Add("0.001")
                ComboBox5.Items.Add("0.01")
                ComboBox5.SelectedIndex = 0
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CalculateTMU()
        If _isPopulating OrElse TextBox29 Is Nothing OrElse ComboBox5.SelectedItem Is Nothing Then Return
        Try
            Dim sumSqMaster As Decimal = 0D
            If _selectedMastersObservation.Count > 0 Then
                For Each m In _selectedMastersObservation
                    sumSqMaster += (m.MasterUncertainty ^ 2)
                Next
            Else
                sumSqMaster = (0.00012D ^ 2) ' Fallback
            End If
            Dim masterU_Observation As Decimal = Math.Sqrt(sumSqMaster)
            
            Dim totalSumSq = 0D

            ' Observations Repeatability
            For Each param In observationParameters
                Dim vals As New List(Of Decimal)
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txt{param}_{i}"), TextBox)
                    If tb IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
                        Dim val As Decimal
                        If Decimal.TryParse(tb.Text, val) Then vals.Add(val)
                    End If
                Next
                If vals.Count > 0 Then
                    Dim sigmaB1 = calcStdev(vals.ToArray()) / Math.Sqrt(3)
                    ' Excel logic: add master uncertainty square for EACH parameter that has data
                    totalSumSq += (sigmaB1 ^ 2) + (masterU_Observation ^ 2)
                End If
            Next

            TextBox29.Text = (2 * Math.Sqrt(totalSumSq)).ToString("F9")
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
        ' Validation checks
        For Each param In observationParameters
            Dim tbs As New List(Of TextBox)
            For i = 1 To 3
                Dim tb = DirectCast(Me.FindName($"txt{param}_{i}"), TextBox)
                If tb IsNot Nothing Then tbs.Add(tb)
            Next
            Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
            If count > 0 AndAlso count < 3 Then
                MessageBox.Show($"Please complete all 3 trials for {param}.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
            If count = 0 Then
                MessageBox.Show($"Please complete trials for {param}.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
        Next

        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse ComboBox5.SelectedItem Is Nothing Then
            MessageBox.Show("Control No and LC must be filled.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim type = ComboBox1.Text
        Dim controlNo = TextBox1.Text.Trim()
        Dim cycleName = MySqlCls.GetActiveCycleName()
        Dim size = If(TextBox21.SelectedItem IsNot Nothing, TextBox21.SelectedItem.ToString(), TextBox21.Text)
        Dim lc = ComboBox5.Text
        Dim color = TextBoxColor.Text
        Dim location = ComboBox2.Text
        Dim temp = TextBox30.Text
        Dim humidity = TextBox28.Text
        Dim tmu = TextBox29.Text
        Dim depthError = If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text)
        Dim remark = TextBox19.Text
        Dim timeIn = TextBox24.Text
        Dim timeOut = TextBox26.Text
        Dim totalTime = TextBox25.Text
        Dim status = If(ComboBox4.SelectedItem IsNot Nothing, DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString(), "OK")

        Dim timeInStr = TextBox24.Text.Trim(), timeOutStr = TextBox26.Text.Trim()
        If String.IsNullOrEmpty(timeOutStr) Then timeOutStr = DateTime.Now.ToString("HH:mm") : TextBox26.Text = timeOutStr
        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then
            MessageBox.Show("Min 2 hours duration required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Persist environment data for next instance
        Dim dateStr = If(TxtCalibrationDate.SelectedDate, DateTime.Today).ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, "")

        Dim wideRange = {txtWideRange_1.Text, txtWideRange_2.Text, txtWideRange_3.Text}
        Dim adjacent = {txtAdjacent_1.Text, txtAdjacent_2.Text, txtAdjacent_3.Text}
        Dim repeatability = {txtRepeatability_1.Text, txtRepeatability_2.Text, txtRepeatability_3.Text}

        Dim success As Boolean = False
        If IsEditMode Then
            success = MySqlCls.UpdateBoreGauge3(type, controlNo, TargetCycle, size, lc, color, location, temp, humidity, tmu, wideRange, adjacent, repeatability, depthError, timeIn, timeOut, totalTime, status, remark, TxtCalibrationDate.SelectedDate)
        Else
            success = MySqlCls.InsertBoreGauge3(type, controlNo, cycleName, TxtCalibrationDate.SelectedDate, size, lc, color, location, temp, humidity, tmu, wideRange, adjacent, repeatability, depthError, timeIn, timeOut, totalTime, status, remark)
        End If

        If success Then
            Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
            Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)
            MySqlCls.UpdateRegularCalibrationStatus(controlNo, calibDate, status, targetCyc)
            MySqlCls.InsertResultRecord(controlNo, _category, _instrumentName, targetCyc)

            ' Save masters to cache
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
                        sizes = New List(Of String) From {"General"}
                    End If
                    
                    TextBox21.ItemsSource = sizes

                    ' Enhanced Size Matching
                    Dim instrumentSize = ""
                    if dtDept.Rows.Count > 0 AndAlso Not IsDBNull(dtDept.Rows(0)("SizeandRange")) Then
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
                TextBox21.ItemsSource = New List(Of String) From {"General"}
                TextBox21.Text = "General"
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

End Class
