Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class VernierCaliperLowForce
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = ""
    Private _category As String = "Instrument"
    Private _selectedMasters As New List(Of MasterSelectorItem)
    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""


    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _activeMinRange As Decimal = 0
    Private _activeMaxRange As Decimal = 190

    Private ReadOnly nominalRanges As Decimal() = {0, 20, 50, 100, 150, 190}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = "Low Force"
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

        TxtCalibrationDate.SelectedDate = DateTime.Today

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

    Private Sub TextBox21_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles TextBox21.SelectionChanged
        If TextBox21.SelectedItem IsNot Nothing Then
            AdjustVisibilityBySize(TextBox21.SelectedItem.ToString())
        End If
    End Sub

    Private Function IsNominalInRange(nominal As Decimal) As Boolean
        Return nominal >= _activeMinRange AndAlso nominal <= _activeMaxRange
    End Function

    Private Sub AdjustVisibilityBySize(sizeStr As String)
        Try
            Dim maxRange As Integer = 300
            If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
                Dim parts = sizeStr.Split("-"c)
                If parts.Length > 1 Then
                    Dim secondPart = parts(1).Replace("mm", "").Trim()
                    Integer.TryParse(secondPart, maxRange)
                End If
            End If

            For Each r In nominalRanges
                Dim inRange = (r <= maxRange)
                SetRangeEnabledState(r, inRange)
            Next

            UpdateSizeLockState()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SetRangeEnabledState(nominal As Decimal, enabled As Boolean)
        Dim tbs = GetTextBoxesForRange(nominal)
        For Each tb In tbs
            tb.IsEnabled = enabled
            tb.Background = If(enabled, Brushes.White, New SolidColorBrush(Color.FromRgb(241, 245, 249)))
            If Not enabled Then tb.Clear()
        Next
        
        Dim lblName = "Label" & nominal.ToString().Replace(".", "_")
        Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)
        If lbl IsNot Nothing Then
            lbl.Foreground = If(enabled, New SolidColorBrush(Color.FromRgb(51, 65, 85)), Brushes.LightSlateGray)
        End If
    End Sub

    Private Sub ObservationKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)
            Dim nameParts = tb.Name.Split("_"c) 
            If nameParts.Length = 2 Then
                Dim trialNum = Integer.Parse(nameParts(1))
                If trialNum < 3 Then
                    Dim nextTbName = nameParts(0) & "_" & (trialNum + 1)
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
                tb.Text = val.ToString("0.00")
            End If
        End If
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)

        UpdateSizeLockState()

        If _masterLimits Is Nothing Then
            txtBox.Background = Brushes.White
            Return
        End If

        Dim nominal As Decimal = GetNominalFromTextBox(txtBox)
        If nominal = -1 Then Return

        Dim prefix As String = "Obs"
        Dim ulCol As String = $"{prefix}{nominal}_UL"
        Dim llCol As String = $"{prefix}{nominal}_LL"

        Try
            UpdateObservationLabel(nominal)
            CalculateTMU()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If

            Dim val As Decimal
            If Decimal.TryParse(txtBox.Text, val) Then
                Dim error1 As Decimal = val
                Dim error2 As Decimal = val - nominal
                Dim errorVal As Decimal = If(Math.Abs(error1) < Math.Abs(error2), error1, error2)

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

    Private Sub UpdateSizeLockState()
        Try
            Dim anyFilled As Boolean = False
            For Each r In nominalRanges
                For Each tb In GetTextBoxesForRange(r)
                    If Not String.IsNullOrEmpty(tb.Text) Then
                        anyFilled = True
                        Exit For
                    End If
                Next
                If anyFilled Then Exit For
            Next
            
            TextBox21.IsEnabled = Not anyFilled
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EnableNextRange(currentNominal As Decimal)
        Dim nominalIdx = Array.IndexOf(nominalRanges, currentNominal)
        If nominalIdx < nominalRanges.Length - 1 Then
            Dim nextNominal = nominalRanges(nominalIdx + 1)
            EnableRange(nextNominal)
        End If
    End Sub

    Private Sub EnableRange(nominal As Decimal)
        Dim tbs = GetTextBoxesForRange(nominal)
        For Each tb In tbs
            tb.IsEnabled = True
            If String.IsNullOrWhiteSpace(tb.Text) Then tb.Background = Brushes.White
        Next
        if tbs.Count > 0 then tbs(0).Focus()
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
    End Sub

    Private Function GetTextBoxesForRange(nominal As Decimal) As List(Of TextBox)
        Dim list As New List(Of TextBox)
        For i = 1 To 3
            Dim name = $"txt{nominal}_{i}"
            Dim tb = DirectCast(Me.FindName(name), TextBox)
            If tb IsNot Nothing Then list.Add(tb)
        Next
        Return list
    End Function

    Private Function GetNominalFromTextBox(tb As TextBox) As Decimal
        Dim name = tb.Name.Replace("txt", "")
        Dim parts = name.Split("_"c)
        If parts.Length > 0 Then
            Dim nominal As Decimal
            If Decimal.TryParse(parts(0), nominal) Then Return nominal
        End If
        Return -1
    End Function

    Private Sub UpdateObservationLabel(nominal As Decimal)
        Dim tbs = GetTextBoxesForRange(nominal)
        Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
        Dim lblName = $"Label{nominal}"
        Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)

        If lbl IsNot Nothing Then
            lbl.Text = $"{nominal} mm:"
            lbl.Foreground = If(count >= 3, Brushes.Gray, Brushes.DodgerBlue)
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
                    
                    TextBox21.IsEnabled = True
                    TextBox21.ItemsSource = sizes

                    ' Size Matching Logic: Try to match size from database with combobox items
                    Dim matchedSize As String = ""
                    If row.Table.Columns.Contains("Size") AndAlso Not IsDBNull(row("Size")) Then
                        matchedSize = TryMatchSize(row("Size").ToString(), sizes)
                    End If

                    If Not String.IsNullOrEmpty(matchedSize) Then
                        TextBox21.SelectedItem = matchedSize
                    ElseIf sizes.Count = 1 Then
                        TextBox21.SelectedIndex = 0
                    Else
                        TextBox21.SelectedIndex = -1
                        TextBox21.Text = ""
                    End If

                    AdjustVisibilityBySize(TextBox21.Text)

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
                _instrumentName = "" : ComboBox2.Text = "" : TextBoxColor.Text = "" : TextBox21.Text = ""
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function TryMatchSize(inputSize As String, sizes As List(Of String)) As String
        If String.IsNullOrWhiteSpace(inputSize) Then Return ""

        ' 1. Normalization
        Dim cleanInput = inputSize.ToLower().Replace("mm", "").Replace(" ", "").Trim()

        ' 2. Direct clean match
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanInput = cleanS Then Return s
        Next

        ' 3. Endpoint match (e.g. "100" or "100mm" matches "0-100 mm")
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanS.Contains("-") Then
                Dim parts = cleanS.Split("-"c)
                If parts.Length > 1 AndAlso parts(1).Trim() = cleanInput Then
                    Return s
                End If
            End If
        Next

        ' 4. Partial match
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanS.Contains(cleanInput) OrElse cleanInput.Contains(cleanS) Then
                Return s
            End If
        Next

        Return ""
    End Function

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
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM vernier_caliper_low_force WHERE RowType='MASTER' ORDER BY LC ASC")
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
            _masterLimits = MySqlCls.GetVernierMasterLimitsLowForce(selectedLC)
            ResetAllTrials() : EnableRange(0)
        Else
            _masterLimits = Nothing : ResetAllTrials() : DisableAllObservationFields()
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
                If tb.Background IsNot Nothing andalso TypeOf tb.Background Is SolidColorBrush then
                    if DirectCast(tb.Background, SolidColorBrush).Color = Color.FromRgb(255, 182, 193) then
                        isNG = True : Exit For
                    End If
                End If
            Next
            If isNG then Exit For
        Next
        If Not isNG then
            For Each tb In {TextBox30, TextBox28, TextBox20}
                if tb.Background IsNot Nothing andalso TypeOf tb.Background Is SolidColorBrush then
                    if DirectCast(tb.Background, SolidColorBrush).Color = Color.FromRgb(255, 182, 193) then isNG = True
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
            MessageBox.Show("Control No and LC must be filled.") : Return
        End If
        if ComboBox4.SelectedItem Is Nothing then MessageBox.Show("Please select Status.") : Return

        Dim timeInStr = TextBox24.Text.Trim(), timeOutStr = TextBox26.Text.Trim()
        If String.IsNullOrEmpty(timeOutStr) Then timeOutStr = DateTime.Now.ToString("HH:mm") : TextBox26.Text = timeOutStr
        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then MessageBox.Show("Calibration duration must be at least 2 hours.") : Return

        Dim dateStr = If(TxtCalibrationDate.SelectedDate, DateTime.Today).ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, TextBox20.Text)

        Dim statusVal = DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString()
        Dim cycleName = MySqlCls.GetActiveCycleName(), calibDateToSave = If(TxtCalibrationDate.SelectedDate, DateTime.Today)

        ' Prepare arrays for database (using single observation points)
        Dim obs0 = GetTrialsArrayFromUI(0), obs20 = GetTrialsArrayFromUI(20), obs50 = GetTrialsArrayFromUI(50), obs100 = GetTrialsArrayFromUI(100), obs150 = GetTrialsArrayFromUI(150), obs190 = GetTrialsArrayFromUI(190)

        Dim saved As Boolean = False
        If IsEditMode Then
            saved = MySqlCls.UpdateVernierCalibrationLowForce(
                _instrumentName, TextBox1.Text.Trim(), TargetCycle, "Low Force", ComboBox5.Text, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                obs0, obs20, obs50, obs100, obs150, obs190,
                If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text), TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text,
                calibDateToSave)
        Else
            saved = MySqlCls.InsertVernierCalibrationLowForce(
                _instrumentName, TextBox1.Text.Trim(), cycleName, "Low Force", ComboBox5.Text, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                obs0, obs20, obs50, obs100, obs150, obs190,
                If(String.IsNullOrEmpty(TextBox20.Text), "-", TextBox20.Text), TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text,
                calibDateToSave)
        End If

        If saved Then
            MySqlCls.UpdateRegularCalibrationStatus(TextBox1.Text.Trim(), calibDateToSave, statusVal, cycleName)
            If Not String.IsNullOrEmpty(cycleName) Then MySqlCls.InsertResultRecord(TextBox1.Text.Trim(), _category, _instrumentName, cycleName)

            ' Save masters to cache
            CalibrationMasterCache.SaveMasters(TextBox1.Text.Trim(), cycleName, _selectedMasters)

            If statusVal = "NG" Then
                Dim ctrlNo = TextBox1.Text.Trim()
                ' Fetch department and other details from department_list to populate WOP
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
            if i < tbs.Count then
                Dim val As Decimal = 0D
                if Decimal.TryParse(tbs(i).Text, val) then
                    Dim errorVal As Decimal = If(Math.Abs(val) < Math.Abs(val - nominal), val, val - nominal)
                    trials(i) = errorVal.ToString("0.00")
                else
                    trials(i) = "-"
                end if
            else
                trials(i) = "-"
            end if
        Next
        Return trials
    End Function

    Private Function calcSigmaB1(arr() As Decimal) As Decimal
        Dim n = arr.Length
        If n <= 1 Then Return 0
        Dim mean = arr.Average()
        Dim sumSqDiff = arr.Sum(Function(v) (v - mean) ^ 2)
        Dim divisor = CDec(n) * CDec(n - 1)
        If divisor = 0 Then Return 0
        Return CDec(Math.Sqrt(sumSqDiff / divisor))
    End Function

    Private Sub CalculateTMU()
        Try
            If _selectedMasters Is Nothing OrElse _selectedMasters.Count = 0 Then
                TextBox29.Text = ""
                TextBox29.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                Return
            End If
            TextBox29.Background = Brushes.White

            ' 1. Combined Master Uncertainty
            Dim sumSqMaster As Decimal = 0
            If _selectedMasters IsNot Nothing Then
                For Each m In _selectedMasters
                    sumSqMaster += (m.MasterUncertainty ^ 2)
                Next
            End If
            Dim masterU As Decimal = CDec(Math.Sqrt(sumSqMaster))

            ' 2. Low Force has one single series of observations (no Ext/Int split)
            ' Formula: SigmaB1_r = SigmaR1_r / SQRT(n)
            '          MU_r      = SQRT(SigmaB1_r^2 + masterU^2)
            '          PreTMU    = SQRT(SUM MU_r^2)
            '          TMU       = PreTMU * 2
            Dim sumSqMU As Decimal = 0

            For Each r In nominalRanges
                Dim valsList As New List(Of Decimal)
                For Each t In GetTextBoxesForRange(r)
                    Dim val As Decimal
                    If Decimal.TryParse(t.Text, val) Then
                        Dim err1 = val : Dim err2 = val - r
                        valsList.Add(If(Math.Abs(err1) < Math.Abs(err2), err1, err2))
                    Else
                        valsList.Add(0D)
                    End If
                Next

                If valsList.Count > 0 Then
                    Dim sigB1 = calcSigmaB1(valsList.ToArray())
                    sumSqMU += (sigB1 ^ 2) + (masterU ^ 2)
                End If
            Next

            Dim finalTMU = CDec(Math.Sqrt(sumSqMU)) * 2
            TextBox29.Text = finalTMU.ToString("F9")
        Catch ex As Exception
        End Try
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

    Private Sub PopulateCalibrationData()
        If String.IsNullOrEmpty(TargetCycle) Then TargetCycle = MySqlCls.GetActiveCycleName()
        Dim dt = MySqlCls.GetVernierCalibrationLowForceData(TextBox1.Text.Trim(), TargetCycle)
        If dt.Rows.Count = 0 Then Return

        Dim row = dt.Rows(0)

        ' Metadata
        If Not IsDBNull(row("Date")) Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(row("Date"))
        If Not IsDBNull(row("LC")) Then ComboBox5.Text = row("LC").ToString()
        If Not IsDBNull(row("Color")) Then TextBoxColor.Text = row("Color").ToString()
        If Not IsDBNull(row("Location")) Then ComboBox2.Text = row("Location").ToString()
        If Not IsDBNull(row("Temperature")) Then TextBox30.Text = row("Temperature").ToString()
        If Not IsDBNull(row("Humidity")) Then TextBox28.Text = row("Humidity").ToString()
        
        ' Observations
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For i = 1 To 3
                Dim col = $"Obs{r}_{i}"
                If dt.Columns.Contains(col) AndAlso Not IsDBNull(row(col)) AndAlso row(col).ToString() <> "-" Then
                    tbs(i-1).Text = row(col).ToString()
                End If
            Next
        Next

        ' Depth Error
        If Not IsDBNull(row("DepthError")) AndAlso row("DepthError").ToString() <> "-" Then
            TextBox20.Text = row("DepthError").ToString()
        End If

        ' Footer
        If Not IsDBNull(row("TimeIn")) Then TextBox24.Text = FormatTimeForUI(row("TimeIn"))
        If Not IsDBNull(row("TimeOut")) Then 
            TextBox26.Text = FormatTimeForUI(row("TimeOut"))
            _isTimeOutManual = True
        End If
        If Not IsDBNull(row("Remark")) Then TextBox19.Text = row("Remark").ToString()
        
        ' Status 
        If Not IsDBNull(row("Status")) Then
            Dim stat = row("Status").ToString()
            For Each item As ComboBoxItem In ComboBox4.Items
                If item.Content.ToString() = stat Then
                    ComboBox4.SelectedItem = item
                    Exit For
                End If
            Next
        End If

        ' Trigger calculations
        UpdateDuration(Nothing, Nothing)
        TriggerAllValidations()
        CalculateTMU()
        UpdateOverallStatus()
        If Not IsDBNull(row("TMU")) Then
            TextBox29.Text = row("TMU").ToString()
            TextBox29.Background = Brushes.White
        End If

        ' Restore cached calibration masters
        Dim cachedMasters = CalibrationMasterCache.LoadMasters(TextBox1.Text.Trim(), TargetCycle)
        If cachedMasters.Count > 0 Then
            _selectedMasters = MySqlCls.GetMastersByDescriptions(cachedMasters)
            UpdateMasterUI()
        End If
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

    Private Function FormatTimeForUI(timeObj As Object) As String
        If IsDBNull(timeObj) Then Return ""
        Dim s = timeObj.ToString()
        If s.Contains(":") Then
            Dim parts = s.Split(":"c)
            If parts.Length >= 2 Then
                Return parts(0).PadLeft(2, "0"c) & ":" & parts(1).PadLeft(2, "0"c)
            End If
        End If
        Return s
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

