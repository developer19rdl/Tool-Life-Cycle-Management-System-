Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class HeightGauge600
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = ""
    Private _category As String = "Instrument"
    Private _selectedMasters As New List(Of MasterSelectorItem)
    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""



    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False

    Private ReadOnly nominalRanges As Decimal() = {0, 20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = "0-600"
        LoadDynamicLC()
        LoadCalibrationMasters()

        ' Add handlers for all trial textboxes
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                tb.IsEnabled = False
                AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                AddHandler tb.TextChanged, AddressOf TrialTextChanged
                AddHandler tb.LostFocus, AddressOf FormatDecimal
            Next
        Next

        ' Auto-populate master card data if Control No was supplied programmatically
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

        If IsEditMode Then
            PopulateCalibrationData()
        End If

        ' Add handlers for Environmental conditions
        AddHandler TextBox30.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox28.TextChanged, AddressOf EnvironmentTextChanged
    End Sub

    Private Sub ObservationKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)

            ' Move focus to the next textbox in the triad, or the next range
            Dim nameParts = tb.Name.Split("_"c) ' e.g., txt0_1 -> [txt0, 1]
            If nameParts.Length = 2 Then
                Dim trialNum = Integer.Parse(nameParts(1))
                If trialNum < 3 Then
                    ' Move to next trial in same range
                    Dim nextTbName = nameParts(0) & "_" & (trialNum + 1)
                    Dim nextTb = DirectCast(Me.FindName(nextTbName), TextBox)
                    If nextTb IsNot Nothing Then
                        nextTb.Focus()
                    End If
                Else
                    ' Trial 3 finished, enable next range logic
                    Dim nominal = GetNominalFromTextBox(tb)
                    EnableNextRange(nominal)
                End If
            End If
        End If
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)

        ' Size Locking
        UpdateSizeLockState()

        If _masterLimits Is Nothing Then
            txtBox.Background = Brushes.White
            Return
        End If

        Dim nominal As Decimal = GetNominalFromTextBox(txtBox)
        if nominal = -1 Then Return

        ' DB columns are like Obs0_0_UL, Obs20_0_UL
        Dim pointPart = nominal.ToString().Replace(".", "_")
        Dim ulCol As String = $"Obs{pointPart}_0_UL"
        Dim llCol As String = $"Obs{pointPart}_0_LL"

        Try
            UpdateObservationLabel(nominal)
            CalculateTMU()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If

            Dim val As Decimal
            If Decimal.TryParse(txtBox.Text, val) Then
                ' DUAL MODE logic: val is either raw or error
                Dim error1 As Decimal = val
                Dim error2 As Decimal = val - nominal
                Dim errorVal As Decimal = If(Math.Abs(error1) < Math.Abs(error2), error1, error2)

                Dim ul As Decimal = Convert.ToDecimal(_masterLimits(ulCol))
                Dim ll As Decimal = Convert.ToDecimal(_masterLimits(llCol))

                If errorVal >= ll AndAlso errorVal <= ul Then
                    txtBox.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144)) ' Light Green
                Else
                    txtBox.Background = New SolidColorBrush(Color.FromRgb(255, 182, 193)) ' Light Red
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
                Dim tbs = GetTextBoxesForRange(r)
                If tbs.Any(Function(tb) Not String.IsNullOrWhiteSpace(tb.Text)) Then
                    anyFilled = True
                    Exit For
                End If
            Next
            TextBox21.IsEnabled = Not anyFilled
        Catch ex As Exception
        End Try
    End Sub

    Private Function GetMaxRangeFromSize(sizeStr As String) As Integer
        Dim maxRange As Integer = 600
        If Not String.IsNullOrEmpty(sizeStr) Then
            Try
                Dim matches = System.Text.RegularExpressions.Regex.Matches(sizeStr, "\d+")
                If matches.Count > 0 Then
                    Integer.TryParse(matches(matches.Count - 1).Value, maxRange)
                End If
            Catch
                ' Fallback
            End Try
        End If
        Return maxRange
    End Function

    Private Sub EnableNextRange(currentNominal As Decimal)
        Dim maxRange As Integer = GetMaxRangeFromSize(TextBox21.Text)

        Dim nominalIdx = Array.IndexOf(nominalRanges, currentNominal)
        If nominalIdx < nominalRanges.Length - 1 Then
            Dim nextNominal = nominalRanges(nominalIdx + 1)
            If nextNominal <= maxRange Then
                EnableRange(nextNominal)
            End If
        End If
    End Sub

    Private Sub EnableRange(nominal As Decimal)
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
                TextBox24.Text = env.TimeIn
            Else
                TextBox30.Text = ""
                TextBox28.Text = ""
                TextBox24.Text = ""
            End If
        End If
    End Sub

    Private Sub DisableAllObservationFields()
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                tb.IsEnabled = False
                tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
        Next
    End Sub

    Private Sub TextBox21_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles TextBox21.SelectionChanged
        If TextBox21.SelectedItem IsNot Nothing Then
            AdjustVisibilityBySize(TextBox21.SelectedItem.ToString())
        End If
    End Sub

    Private Sub AdjustVisibilityBySize(sizeStr As String)
        Try
            Dim maxRange As Integer = GetMaxRangeFromSize(sizeStr)

            For Each r In nominalRanges
                Dim inRange = (r <= maxRange)
                SetRangeEnabledState(r, inRange)
            Next

            DisableAllObservationFields()
            EnableRange(0)
            UpdateSizeLockState()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SetRangeEnabledState(nominal As Decimal, inRange As Boolean)
        Dim bg = If(inRange, Brushes.White, New SolidColorBrush(Color.FromRgb(241, 245, 249)))
        Dim fg = If(inRange, New SolidColorBrush(Color.FromRgb(51, 65, 85)), Brushes.Gray)

        Dim lbl = DirectCast(Me.FindName("Label" & nominal.ToString().Replace(".", "_")), TextBlock)
        If lbl IsNot Nothing Then lbl.Foreground = fg
        For Each tb In GetTextBoxesForRange(nominal)
            tb.IsEnabled = inRange
            tb.Background = bg
            If Not inRange Then tb.Clear()
        Next
    End Sub

    Private Function GetTextBoxesForRange(nominal As Decimal) As List(Of TextBox)
        Dim list As New List(Of TextBox)
        Dim namePart = nominal.ToString().Replace(".", "_")
        For i = 1 To 3
            Dim name = $"txt{namePart}_{i}"
            Dim tb = DirectCast(Me.FindName(name), TextBox)
            If tb IsNot Nothing Then list.Add(tb)
        Next
        Return list
    End Function

    Private Function GetNominalFromTextBox(tb As TextBox) As Decimal
        Dim name = tb.Name.Replace("txt", "")
        Dim parts = name.Split("_"c)
        If parts.Length > 0 Then
            Dim nominalStr = parts(0).Replace("_", ".")
            Dim nominal As Decimal
            If Decimal.TryParse(nominalStr, nominal) Then Return nominal
        End If
        Return -1
    End Function

    Private Sub UpdateObservationLabel(nominal As Decimal)
        Dim tbs = GetTextBoxesForRange(nominal)
        Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()

        Dim lblName = "Label" & nominal.ToString().Replace(".", "_")
        Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)
        If lbl IsNot Nothing Then
            lbl.Text = $"{nominal} mm:"
            If ComboBox5.SelectedItem Is Nothing Then
                lbl.Foreground = New SolidColorBrush(Color.FromRgb(51, 65, 85))
            Else
                If count >= 3 Then
                    lbl.Foreground = Brushes.Gray
                Else
                    lbl.Foreground = Brushes.DodgerBlue
                End If
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
                Dim query = $"SELECT * FROM `{tbl}` WHERE ControlNo = '{controlNo.Replace("'", "''")}' LIMIT 1"
                Dim dt = MySqlCls.ReadDatatable(query)
                If dt.Rows.Count > 0 Then
                    Dim row = dt.Rows(0)
                    Dim cycleName = MySqlCls.GetActiveCycleName()
                    Dim dtDept = MySqlCls.ReadDatatable($"SELECT Department, Color, SizeandRange FROM department_list WHERE `Control No` = '{controlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")

                    _category = r("Category").ToString()
                    Dim prefix = MySQLClass.ExtractPrefix(controlNo)
                    Dim sizes = MySqlCls.GetSizesForGroup(prefix, _category)
                    TextBox21.ItemsSource = sizes

                    Dim matchedSize As String = ""
                    If row.Table.Columns.Contains("Size") AndAlso Not IsDBNull(row("Size")) Then
                        matchedSize = TryMatchSize(row("Size").ToString(), sizes)
                    End If

                    If Not String.IsNullOrEmpty(matchedSize) Then
                        TextBox21.Text = matchedSize
                    ElseIf sizes.Count = 1 Then
                        TextBox21.Text = sizes(0)
                    Else
                        TextBox21.Text = ""
                    End If

                    If dtDept.Rows.Count > 0 Then
                        Dim deptRow = dtDept.Rows(0)
                        ComboBox2.Text = deptRow("Department").ToString()
                        TextBoxColor.Text = deptRow("Color").ToString()
                    Else
                        ComboBox2.Text = If(row.Table.Columns.Contains("Location"), row("Location").ToString(), "")
                        TextBoxColor.Text = If(row.Table.Columns.Contains("Color"), row("Color").ToString(), "")
                    End If

                    AdjustVisibilityBySize(TextBox21.Text)
                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), row("InstrumentName").ToString(), row("GaugeName").ToString())

                    found = True
                    Exit For
                End If
            Next

            If Not found Then
                _instrumentName = ""
                ClearForm()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function TryMatchSize(inputSize As String, sizes As List(Of String)) As String
        If String.IsNullOrWhiteSpace(inputSize) Then Return ""
        Dim cleanInput = inputSize.ToLower().Replace("mm", "").Replace(" ", "").Trim()
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanInput = cleanS Then Return s
        Next
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanS.Contains("-") Then
                Dim parts = cleanS.Split("-"c)
                If parts.Length > 1 AndAlso parts(1).Trim() = cleanInput Then Return s
            End If
        Next
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanS.Contains(cleanInput) OrElse cleanInput.Contains(cleanS) Then Return s
        Next
        Return ""
    End Function

    Private Sub LoadCalibrationMasters()
        _selectedMasters.Clear()
    End Sub

    Private Sub BtnBrowseMaster_Click(sender As Object, e As RoutedEventArgs) Handles BtnBrowseMaster.Click
        Dim currentSelection = _selectedMasters.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMasters = selector.SelectedMasters
            UpdateMasterUI()
            CalculateTMU()
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

    Private Sub ClearForm()
        ComboBox2.Text = ""
        TextBoxColor.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub LoadDynamicLC()
        Try
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM height_gauge_600 WHERE RowType='MASTER' ORDER BY LC ASC")
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
            _masterLimits = MySqlCls.GetHeightGauge600MasterLimits(selectedLC)
            ResetAllTrials()
        Else
            _masterLimits = Nothing
            ResetAllTrials()
            DisableAllObservationFields()
        End If
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


    Private Sub UpdateOverallStatus()
        Dim isNG As Boolean = False
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                If tb.Background IsNot Nothing AndAlso TypeOf tb.Background Is SolidColorBrush Then
                    Dim color = DirectCast(tb.Background, SolidColorBrush).Color
                    If color = Color.FromRgb(255, 182, 193) Then
                        isNG = True
                        Exit For
                    End If
                End If
            Next
            If isNG Then Exit For
        Next
        If Not isNG Then
            Dim envTbs = {TextBox30, TextBox28}
            For Each tb In envTbs
                If tb.Background IsNot Nothing AndAlso TypeOf tb.Background Is SolidColorBrush Then
                    Dim color = DirectCast(tb.Background, SolidColorBrush).Color
                    If color = Color.FromRgb(255, 182, 193) Then
                        isNG = True
                        Exit For
                    End If
                End If
            Next
        End If
        ComboBox4.SelectedIndex = If(isNG, 1, 0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs) Handles Button2.Click
        Dim maxRange As Integer = GetMaxRangeFromSize(TextBox21.Text)

        For Each r In nominalRanges
            If r <= maxRange Then
                Dim tbs = GetTextBoxesForRange(r)
                Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
                If count > 0 AndAlso count < 3 Then
                    MessageBox.Show($"Please complete all 3 trials for {r}mm.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                    Return
                End If
            End If
        Next

        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse ComboBox5.SelectedItem Is Nothing Then
            MessageBox.Show("Control No and LC must be filled.")
            Return
        End If
        If ComboBox4.SelectedItem Is Nothing Then
            MessageBox.Show("Please select OK/NG Status.")
            Return
        End If

        Dim timeInStr = TextBox24.Text.Trim()
        Dim timeOutStr = TextBox26.Text.Trim()
        If String.IsNullOrEmpty(timeOutStr) Then
            timeOutStr = DateTime.Now.ToString("HH:mm")
            TextBox26.Text = timeOutStr
        End If

        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then
            MessageBox.Show("Calibration duration must be at least 2 hours.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim dateStr = If(TxtCalibrationDate.SelectedDate, DateTime.Today).ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, "")

        Dim statusVal = DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString()
        Dim lcStr = ComboBox5.SelectedItem.ToString()
        Dim calibDateToSave = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim cycleName = MySqlCls.GetActiveCycleName()

        Dim saved As Boolean = False
        If IsEditMode Then
            saved = MySqlCls.UpdateHeightGauge600Calibration(
                _instrumentName, TextBox1.Text.Trim(), TargetCycle, TextBox21.Text, lcStr, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                GetTrialsArrayFromUI(0, maxRange), GetTrialsArrayFromUI(20, maxRange), GetTrialsArrayFromUI(50, maxRange), GetTrialsArrayFromUI(100, maxRange),
                GetTrialsArrayFromUI(150, maxRange), GetTrialsArrayFromUI(200, maxRange), GetTrialsArrayFromUI(250, maxRange), GetTrialsArrayFromUI(300, maxRange),
                GetTrialsArrayFromUI(350, maxRange), GetTrialsArrayFromUI(400, maxRange), GetTrialsArrayFromUI(450, maxRange), GetTrialsArrayFromUI(500, maxRange),
                GetTrialsArrayFromUI(550, maxRange), GetTrialsArrayFromUI(600, maxRange),
                TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text, calibDateToSave)
        Else
            saved = MySqlCls.InsertHeightGauge600Calibration(
                _instrumentName, TextBox1.Text.Trim(), cycleName, TextBox21.Text, lcStr, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                GetTrialsArrayFromUI(0, maxRange), GetTrialsArrayFromUI(20, maxRange), GetTrialsArrayFromUI(50, maxRange), GetTrialsArrayFromUI(100, maxRange),
                GetTrialsArrayFromUI(150, maxRange), GetTrialsArrayFromUI(200, maxRange), GetTrialsArrayFromUI(250, maxRange), GetTrialsArrayFromUI(300, maxRange),
                GetTrialsArrayFromUI(350, maxRange), GetTrialsArrayFromUI(400, maxRange), GetTrialsArrayFromUI(450, maxRange), GetTrialsArrayFromUI(500, maxRange),
                GetTrialsArrayFromUI(550, maxRange), GetTrialsArrayFromUI(600, maxRange),
                TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text, calibDateToSave)
        End If

        If saved Then
            MySqlCls.UpdateRegularCalibrationStatus(TextBox1.Text.Trim(), calibDateToSave, statusVal, cycleName)
            MySqlCls.InsertResultRecord(TextBox1.Text.Trim(), _category, _instrumentName, cycleName)
            CalibrationMasterCache.SaveMasters(TextBox1.Text.Trim(), cycleName, _selectedMasters)

            If statusVal = "NG" Then
                Dim ctrlNo = TextBox1.Text.Trim()
                Dim dtDept = MySqlCls.ReadDatatable($"SELECT * FROM department_list WHERE `Control No` = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")
                If dtDept.Rows.Count > 0 Then
                    Dim row = dtDept.Rows(0)
                    Dim dept = row("Department").ToString()
                    Dim instName = row("InstrumentName").ToString()
                    Dim color = row("Color").ToString()
                    Dim size = ""
                    If row.Table.Columns.Contains("SizeandRange") AndAlso Not IsDBNull(row("SizeandRange")) Then size = row("SizeandRange").ToString()

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

            MessageBox.Show("Record Saved Successfully!")
            Me.Close()
        Else
            MessageBox.Show("Error saving record: " & MySqlCls.LastError)
        End If
    End Sub

    Private Function GetTrialsArrayFromUI(nominal As Decimal, maxRange As Integer) As Object()
        If nominal > maxRange Then Return {DBNull.Value, DBNull.Value, DBNull.Value}
        Dim tbs = GetTextBoxesForRange(nominal)
        Return tbs.Select(Function(t) If(String.IsNullOrWhiteSpace(t.Text), DirectCast(DBNull.Value, Object), DirectCast(t.Text, Object))).ToArray()
    End Function

    Private Sub UpdateDuration(sender As Object, e As TextChangedEventArgs)
        Dim timeInStr = TextBox24.Text.Trim()
        Dim timeOutStr = TextBox26.Text.Trim()

        Dim dummyTimeOut = timeOutStr
        If String.IsNullOrEmpty(dummyTimeOut) Then
            dummyTimeOut = DateTime.Now.ToString("HH:mm")
        End If

        Dim totalMinutes = CalculateDuration(timeInStr, dummyTimeOut)
        
        If totalMinutes >= 0 Then
            Dim hours = totalMinutes \ 60
            Dim mins = totalMinutes Mod 60
            TextBox25.Text = $"{hours:D2}:{mins:D2}"
            
            If totalMinutes < 120 Then
                TextBox25.Background = New SolidColorBrush(Color.FromRgb(255, 182, 193)) ' Red-ish
            Else
                TextBox25.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144)) ' Green-ish
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
            For Each m In _selectedMasters
                sumSqMaster += (m.MasterUncertainty ^ 2)
            Next
            Dim masterU As Decimal = Math.Sqrt(sumSqMaster)

            ' 2. Calculate Repeatability Uncertainty over visible ranges
            Dim totalSumSq As Decimal = 0
            Dim maxRange As Integer = GetMaxRangeFromSize(TextBox21.Text)

            For Each r In nominalRanges
                If r <= maxRange Then
                    Dim tbs = GetTextBoxesForRange(r)
                    Dim vals(2) As Decimal
                    Dim hasData As Boolean = False

                    For i = 1 To 3
                        Dim textVal = If(tbs.Count >= i, tbs(i - 1).Text, "")
                        If Decimal.TryParse(textVal, vals(i - 1)) Then
                            hasData = True
                        Else
                            vals(i - 1) = 0D
                        End If
                    Next

                    If hasData Then
                        Dim stdev = calcStdev(vals)
                        Dim sigB1 = stdev / Math.Sqrt(3)
                        Dim mu_sq = (sigB1 ^ 2) + (masterU ^ 2)
                        totalSumSq += mu_sq
                    End If
                End If
            Next

            ' 3. Final TMU Calculation (k=2)
            Dim finalTMU = Math.Sqrt(totalSumSq) * 2
            TextBox29.Text = finalTMU.ToString("F9")
        Catch ex As Exception
            TextBox29.Text = "Error"
        End Try
    End Sub

    Private Function calcStdev(arr() As Decimal) As Decimal
        Dim n = arr.Length
        If n <= 1 Then Return 0
        Dim mean = arr.Average()
        Dim sumSqDiff = arr.Sum(Function(v) (v - mean) ^ 2)
        Return Math.Sqrt(sumSqDiff / (n - 1))
    End Function

    Private Sub FormatDecimal(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            tb.Text = val.ToString("F3")
        End If
    End Sub

    Private Sub TriggerAllValidations()
        Try
            ' Environment
            EnvironmentTextChanged(TextBox30, Nothing)
            EnvironmentTextChanged(TextBox28, Nothing)
            
            ' Observations
            Dim maxRange As Integer = GetMaxRangeFromSize(TextBox21.Text)
            For Each r In nominalRanges
                If r <= maxRange Then
                    For Each tb In GetTextBoxesForRange(r)
                        TrialTextChanged(tb, Nothing)
                    Next
                End If
            Next
            
            UpdateOverallStatus()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PopulateCalibrationData()
        Try
            Dim dt = MySqlCls.GetHeightGauge600CalibrationData(TextBox1.Text, TargetCycle)
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                If row("Date") IsNot DBNull.Value Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(row("Date"))
                If row("LC") IsNot DBNull.Value Then ComboBox5.Text = row("LC").ToString()
                If row("Size") IsNot DBNull.Value Then TextBox21.Text = row("Size").ToString()
                If row("Temperature") IsNot DBNull.Value Then TextBox30.Text = row("Temperature").ToString()
                If row("Humidity") IsNot DBNull.Value Then TextBox28.Text = row("Humidity").ToString()
                If row("TMU") IsNot DBNull.Value Then TextBox29.Text = row("TMU").ToString()
                If row("TimeIn") IsNot DBNull.Value Then TextBox24.Text = row("TimeIn").ToString()
                If row("TimeOut") IsNot DBNull.Value Then TextBox26.Text = row("TimeOut").ToString()
                If row("TotalTime") IsNot DBNull.Value Then TextBox25.Text = row("TotalTime").ToString()
                If row("Remark") IsNot DBNull.Value Then TextBox19.Text = row("Remark").ToString()
                If row("Status") IsNot DBNull.Value Then
                    For Each item As ComboBoxItem In ComboBox4.Items
                        If item.Content.ToString() = row("Status").ToString() Then
                            ComboBox4.SelectedItem = item
                            Exit For
                        End If
                    Next
                End If

                ' Fill trials
                For Each r In nominalRanges
                    Dim pointPart = r.ToString().Replace(".", "_")
                    Dim tbs = GetTextBoxesForRange(r)
                    For i = 1 To 3
                        Dim col = $"Obs{pointPart}_0_{i}"
                        If row(col) IsNot DBNull.Value Then
                            tbs(i - 1).Text = Convert.ToDecimal(row(col)).ToString("F3")
                        End If
                    Next
                Next

                Dim cachedMasters = CalibrationMasterCache.LoadMasters(TextBox1.Text, TargetCycle)
                If cachedMasters.Count > 0 Then
                    _selectedMasters = MySqlCls.GetMastersByDescriptions(cachedMasters)
                    UpdateMasterUI()
                End If
                AdjustVisibilityBySize(TextBox21.Text)

                ' Trigger calculations and validation coloring
                UpdateDuration(Nothing, Nothing)
                TriggerAllValidations()
                CalculateTMU()
                UpdateOverallStatus()
            End If
        Catch ex As Exception
        End Try
    End Sub

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
            End If

            If Not isOK Then
                tb.Background = New SolidColorBrush(Color.FromRgb(255, 182, 193)) ' Light Red
            Else
                tb.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144)) ' Light Green
            End If
        Else
            tb.Background = Brushes.White
        End If
        UpdateOverallStatus()
    End Sub

End Class
