Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class VernierCaliper300
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = ""
    Private _category As String = "Instrument"
    Private _selectedMasters As New List(Of MasterSelectorItem)
    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

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
                TextBox30.Text = ""
                TextBox28.Text = ""
                TextBox20.Text = ""
                TextBox24.Text = ""
            End If
        End If
    End Sub

    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False

    Private ReadOnly nominalRanges As Decimal() = {0, 20, 50, 100, 150, 200, 250, 300}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = "0-300"
        LoadDynamicLC()
        LoadCalibrationMasters()

        ' Add handlers for all trial textboxes
        For Each r In nominalRanges
            ' External
            Dim tbsExt = GetTextBoxesForRange(r, True)
            For Each tb In tbsExt
                tb.IsEnabled = False
                AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                AddHandler tb.TextChanged, AddressOf TrialTextChanged
                AddHandler tb.LostFocus, AddressOf FormatDecimal
            Next
            ' Internal
            Dim tbsInt = GetTextBoxesForRange(r, False)
            For Each tb In tbsInt
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

        ' Handled by TxtCalibrationDate_SelectedDateChanged which triggers when date is set to Today below

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
        AddHandler TextBox20.TextChanged, AddressOf EnvironmentTextChanged
    End Sub

    Private Sub ObservationKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)

            ' Move focus to the next textbox in the triad, or the next range
            Dim nameParts = tb.Name.Split("_"c) ' e.g., txtExt0_1 -> [txtExt0, 1]
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
                    Dim isExternal = tb.Name.Contains("Ext")
                    EnableNextRange(nominal, isExternal)
                End If
            End If
        End If
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)

        ' Size Locking: Disable Size selection if any observation is entered
        UpdateSizeLockState()

        If _masterLimits Is Nothing Then
            txtBox.Background = Brushes.White
            Return
        End If

        Dim nominal As Decimal = GetNominalFromTextBox(txtBox)
        If nominal = -1 Then Return

        Dim isExternal As Boolean = txtBox.Name.Contains("Ext")
        Dim prefix As String = If(isExternal, "Ext", "Int")
        Dim ulCol As String = $"{prefix}{nominal}_UL"
        Dim llCol As String = $"{prefix}{nominal}_LL"

        Try
            ' Real-time calculation should trigger even on clear/empty/non-numeric values
            UpdateObservationLabel(nominal, isExternal)
            CalculateTMU()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If

            Dim val As Decimal
            If Decimal.TryParse(txtBox.Text, val) Then
                ' DUAL MODE: User might type raw reading (20.01) OR error offset (0.01)
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
                Dim tbsExt = GetTextBoxesForRange(r, True)
                Dim tbsInt = GetTextBoxesForRange(r, False)
                If tbsExt.Any(Function(tb) Not String.IsNullOrWhiteSpace(tb.Text)) OrElse
                   tbsInt.Any(Function(tb) Not String.IsNullOrWhiteSpace(tb.Text)) Then
                    anyFilled = True
                    Exit For
                End If
            Next
            TextBox21.IsEnabled = Not anyFilled
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EnableNextRange(currentNominal As Decimal, isExternal As Boolean)
        Dim maxRange As Integer = 300
        Dim sizeStr = TextBox21.Text
        If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
            Dim parts = sizeStr.Split("-"c)
            If parts.Length > 1 Then
                Dim secondPart = parts(1).Replace("mm", "").Trim()
                Integer.TryParse(secondPart, maxRange)
            End If
        End If

        Dim nominalIdx = Array.IndexOf(nominalRanges, currentNominal)

        If isExternal Then
            If nominalIdx < nominalRanges.Length - 1 Then
                Dim nextNominal = nominalRanges(nominalIdx + 1)
                If nextNominal <= maxRange Then
                    EnableRange(nextNominal, True)
                    Return
                End If
            End If
            ' External finished, start Internal 0mm
            EnableRange(0, False)
        Else
            If nominalIdx < nominalRanges.Length - 1 Then
                Dim nextNominal = nominalRanges(nominalIdx + 1)
                If nextNominal <= maxRange Then
                    EnableRange(nextNominal, False)
                    Return
                End If
            End If
        End If
    End Sub

    Private Sub EnableRange(nominal As Decimal, isExternal As Boolean)
        Dim tbs = GetTextBoxesForRange(nominal, isExternal)
        For Each tb In tbs
            tb.IsEnabled = True
            ' Keep background as is (might be colored if already edited)
            If String.IsNullOrWhiteSpace(tb.Text) Then tb.Background = Brushes.White
        Next
        If tbs.Count > 0 Then tbs(0).Focus()
    End Sub


    Private Sub DisableAllObservationFields()
        For Each r In nominalRanges
            Dim tbsExt = GetTextBoxesForRange(r, True)
            For Each tb In tbsExt
                tb.IsEnabled = False
                tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
            Dim tbsInt = GetTextBoxesForRange(r, False)
            For Each tb In tbsInt
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

            ' Removed sequential lock to allow editing of all valid observation ranges
            ' DisableAllObservationFields()

            ' Enable first range
            ' EnableRange(0, True)
            
            ' Update Size Lock state (it might be filled already if reloading)
            UpdateSizeLockState()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SetRangeEnabledState(nominal As Decimal, inRange As Boolean)
        Dim bg = If(inRange, Brushes.White, New SolidColorBrush(Color.FromRgb(241, 245, 249)))
        Dim fg = If(inRange, New SolidColorBrush(Color.FromRgb(51, 65, 85)), Brushes.Gray)

        Dim lblExt = DirectCast(Me.FindName("LabelExt" & nominal), TextBlock)
        If lblExt IsNot Nothing Then lblExt.Foreground = fg
        For Each tb In GetTextBoxesForRange(nominal, True)
            tb.IsEnabled = inRange
            tb.Background = bg
            ' If not in range, clear text
            If Not inRange Then tb.Clear()
        Next

        Dim lblInt = DirectCast(Me.FindName("LabelInt" & nominal), TextBlock)
        If lblInt IsNot Nothing Then lblInt.Foreground = fg
        For Each tb In GetTextBoxesForRange(nominal, False)
            tb.IsEnabled = inRange
            tb.Background = bg
            ' If not in range, clear text
            If Not inRange Then tb.Clear()
        Next
    End Sub

    Private Function GetTextBoxesForRange(nominal As Decimal, isExternal As Boolean) As List(Of TextBox)
        Dim list As New List(Of TextBox)
        Dim prefix = If(isExternal, "txtExt", "txtInt")
        For i = 1 To 3
            Dim name = $"{prefix}{nominal}_{i}"
            Dim tb = DirectCast(Me.FindName(name), TextBox)
            If tb IsNot Nothing Then list.Add(tb)
        Next
        Return list
    End Function

    Private Function GetNominalFromTextBox(tb As TextBox) As Decimal
        Dim name = tb.Name.Replace("txtExt", "").Replace("txtInt", "")
        Dim parts = name.Split("_"c)
        If parts.Length > 0 Then
            Dim nominal As Decimal
            If Decimal.TryParse(parts(0), nominal) Then Return nominal
        End If
        Return -1
    End Function

    Private Sub UpdateObservationLabel(nominal As Decimal, isExternal As Boolean)
        Dim tbs = GetTextBoxesForRange(nominal, isExternal)
        Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()

        Dim lblName = (If(isExternal, "LabelExt", "LabelInt")) & nominal.ToString()
        Dim lbl = DirectCast(Me.FindName(lblName), TextBlock)
        If lbl IsNot Nothing Then
            lbl.Text = $"{nominal} mm:"

            ' Dynamic color based on LC selection
            If ComboBox5.SelectedItem Is Nothing Then
                ' Default text color (Black/Dark Gray)
                lbl.Foreground = New SolidColorBrush(Color.FromRgb(51, 65, 85)) ' Matching #334155
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

                    ' Size Matching Logic: Try to match size from database with combobox items
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
                        ' Fallback to source table if not in department_list
                        ComboBox2.Text = row("Location").ToString()
                        TextBoxColor.Text = row("Color").ToString()
                    End If

                    AdjustVisibilityBySize(TextBox21.Text)

                    _category = r("Category").ToString()
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

    Private Function GetFreshInstrumentDetails(controlNo As String) As (Category As String, Name As String)
        Try
            Dim dtSettings = MySqlCls.ReadDatatable("SELECT TypeName, Category FROM type_details")
            For Each r As DataRow In dtSettings.Rows
                Dim tbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString())
                Dim query = $"SELECT * FROM `{tbl}` WHERE ControlNo = '{controlNo.Replace("'", "''")}' LIMIT 1"
                Dim dt = MySqlCls.ReadDatatable(query)
                If dt.Rows.Count > 0 Then
                    Dim row = dt.Rows(0)
                    Dim cat = r("Category").ToString()
                    Dim name = If(cat.Equals("Instrument", StringComparison.OrdinalIgnoreCase), row("InstrumentName").ToString(), row("GaugeName").ToString())
                    Return (cat, name)
                End If
            Next
        Catch ex As Exception
        End Try
        Return ("Instrument", "")
    End Function

    Private Sub LoadCalibrationMasters()
        Try
            ' 1. Get instrument's registration date (created_at) - Kept for reference if needed, but no auto-selection
            _selectedMasters.Clear()
            ' Update: No longer auto-selecting masters on load.
        Catch ex As Exception
            Console.WriteLine("LoadCalibrationMasters Error: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnBrowseMaster_Click(sender As Object, e As RoutedEventArgs)
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
            ' UpdateDuration will be called automatically via TextChanged handler
        End If
    End Sub

    Private Sub ClearForm()
        ComboBox2.Text = ""
        TextBoxColor.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub LoadDynamicLC()
        Try
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM vernier_caliper_300 WHERE RowType='MASTER' ORDER BY LC ASC")
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
            _masterLimits = MySqlCls.GetVernierMasterLimits(selectedLC)
            ' Reset trials if LC changes as limits are different
            ResetAllTrials()
        Else
            _masterLimits = Nothing
            ResetAllTrials()
            DisableAllObservationFields()
        End If
    End Sub

    Private Sub ResetAllTrials()
        For Each r In nominalRanges
            For Each tb In GetTextBoxesForRange(r, True)
                tb.Clear()
            Next
            UpdateObservationLabel(r, True)

            For Each tb In GetTextBoxesForRange(r, False)
                tb.Clear()
            Next
            UpdateObservationLabel(r, False)
        Next

        AdjustVisibilityBySize(TextBox21.Text)
        CalculateTMU()
    End Sub

    Private Sub EnvTextChanged(sender As Object, e As TextChangedEventArgs) Handles TextBox30.TextChanged, TextBox28.TextChanged, TextBox20.TextChanged
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        If _masterLimits Is Nothing Then Return

        Try
            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If
            Dim val As Decimal
            If Decimal.TryParse(txtBox.Text, val) Then
                Dim llCol As String = ""
                Dim ulCol As String = ""
                If txtBox.Name = "TextBox30" Then ' Temp
                    llCol = "Env_Temp_LL" : ulCol = "Env_Temp_UL"
                ElseIf txtBox.Name = "TextBox28" Then ' Humidity
                    llCol = "Env_Hum_LL" : ulCol = "Env_Hum_UL"
                ElseIf txtBox.Name = "TextBox20" Then ' DepthError
                    llCol = "Depth_LL" : ulCol = "Depth_UL"
                End If

                Dim ll As Decimal = Convert.ToDecimal(_masterLimits(llCol))
                Dim ul As Decimal = Convert.ToDecimal(_masterLimits(ulCol))

                If val >= ll AndAlso val <= ul Then
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

    Private Sub UpdateOverallStatus()
        Dim isNG As Boolean = False

        ' 1. Check all observation fields
        For Each r In nominalRanges
            Dim tbsExt = GetTextBoxesForRange(r, True)
            Dim tbsInt = GetTextBoxesForRange(r, False)
            For Each tb In tbsExt.Concat(tbsInt)
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

        ' 2. Check environmental fields
        If Not isNG Then
            Dim envTbs = {TextBox30, TextBox28, TextBox20}
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

        If isNG Then
            ComboBox4.SelectedIndex = 1 ' NG
        Else
            ' Only set to OK if at least one field is filled or just default to OK?
            ' The user said "if any one is red -> NG", so if none are red, it should be OK.
            ComboBox4.SelectedIndex = 0 ' OK
        End If
    End Sub

    Private Function GetDecimalOrNull(txt As String) As Decimal?
        Dim val As Decimal
        If Decimal.TryParse(txt, val) Then Return val
        Return Nothing
    End Function

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs) Handles Button2.Click
        ' Optional: Check if at least one Calibration Master is selected (Commented out as per user request to not associate it for now)
        ' If _selectedMasters.Count = 0 Then
        '     MessageBox.Show("Please select at least one Calibration Master.")
        '     Return
        ' End If

        ' Validation: Check if all VISIBLE ranges have 3 trials
        Dim maxRange As Integer = 300
        Dim sizeStr = TextBox21.Text
        If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
            Dim parts = sizeStr.Split("-"c)
            If parts.Length > 1 Then
                Dim secondPart = parts(1).Replace("mm", "").Trim()
                Integer.TryParse(secondPart, maxRange)
            End If
        End If

        ' Validation: Check if all VISIBLE ranges have 3 trials (Removed as per user request to allow empty trials)
        ' For Each r In nominalRanges
        '     If r <= maxRange Then
        '         Dim tbsExt = GetTextBoxesForRange(r, True)
        '         Dim tbsInt = GetTextBoxesForRange(r, False)
        '         If tbsExt.Any(Function(t) String.IsNullOrWhiteSpace(t.Text)) OrElse tbsInt.Any(Function(t) String.IsNullOrWhiteSpace(t.Text)) Then
        '             MessageBox.Show($"Please complete all 3 trials for {r}mm (both External and Internal).")
        '             Return
        '         End If
        '     End If
        ' Next

        ' New Validation: Partial triad check (If user starts typing, they must complete all 3)
        For Each r In nominalRanges
            If r <= maxRange Then
                ' External
                Dim tbsExt = GetTextBoxesForRange(r, True)
                Dim countExt = tbsExt.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
                If countExt > 0 AndAlso countExt < 3 Then
                    MessageBox.Show($"Please complete all 3 trials for {r}mm (External Observation).", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                    Return
                End If

                ' Internal
                Dim tbsInt = GetTextBoxesForRange(r, False)
                Dim countInt = tbsInt.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
                If countInt > 0 AndAlso countInt < 3 Then
                    MessageBox.Show($"Please complete all 3 trials for {r}mm (Internal Observation).", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
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

        ' Duration Validation
        Dim timeInStr = TextBox24.Text.Trim()
        Dim timeOutStr = TextBox26.Text.Trim()

        If String.IsNullOrEmpty(timeOutStr) Then
            timeOutStr = DateTime.Now.ToString("HH:mm")
            TextBox26.Text = timeOutStr
        End If

        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then ' Less than 2 hours (120 minutes)
            MessageBox.Show("Calibration duration must be at least 2 hours.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If



        Dim statusVal = DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString()

        ' Prepare arrays
        Dim ext0 = GetTrialsArrayFromUI(0, True, maxRange)
        Dim ext20 = GetTrialsArrayFromUI(20, True, maxRange)
        Dim ext50 = GetTrialsArrayFromUI(50, True, maxRange)
        Dim ext100 = GetTrialsArrayFromUI(100, True, maxRange)
        Dim ext150 = GetTrialsArrayFromUI(150, True, maxRange)
        Dim ext200 = GetTrialsArrayFromUI(200, True, maxRange)
        Dim ext250 = GetTrialsArrayFromUI(250, True, maxRange)
        Dim ext300 = GetTrialsArrayFromUI(300, True, maxRange)

        Dim int0 = GetTrialsArrayFromUI(0, False, maxRange)
        Dim int20 = GetTrialsArrayFromUI(20, False, maxRange)
        Dim int50 = GetTrialsArrayFromUI(50, False, maxRange)
        Dim int100 = GetTrialsArrayFromUI(100, False, maxRange)
        Dim int150 = GetTrialsArrayFromUI(150, False, maxRange)
        Dim int200 = GetTrialsArrayFromUI(200, False, maxRange)
        Dim int250 = GetTrialsArrayFromUI(250, False, maxRange)
        Dim int300 = GetTrialsArrayFromUI(300, False, maxRange)

        Dim depthErrorVal As Object = GetDecimalOrNull(TextBox20.Text)
        If depthErrorVal Is Nothing Then depthErrorVal = "-"

        Dim lcStr = ComboBox5.SelectedItem.ToString()
        Dim calibDateToSave = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim cycleName = MySqlCls.GetActiveCycleName()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, calibDateToSave.ToShortDateString(), TextBox30.Text, TextBox28.Text, TextBox24.Text, TextBox20.Text)

        Dim freshDetails = GetFreshInstrumentDetails(TextBox1.Text.Trim())
        Dim finalCat = freshDetails.Category
        Dim finalName = If(String.IsNullOrEmpty(freshDetails.Name), _instrumentName, freshDetails.Name)

        Dim mName = "" ' Reserved for future use
        Dim mUnc As Decimal = 0 ' Reserved for future use
        Dim mUncList = "" ' Reserved for future use

        Dim saved As Boolean = False
        If IsEditMode Then
            saved = MySqlCls.UpdateVernierCalibration300(
                finalName, TextBox1.Text.Trim(), TargetCycle, TextBox21.Text, lcStr, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                ext0, ext20, ext50, ext100, ext150, ext200, ext250, ext300,
                int0, int20, int50, int100, int150, int200, int250, int300,
                depthErrorVal, TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text,
                mName, mUnc, mUncList, calibDateToSave)
        Else
            saved = MySqlCls.InsertVernierCalibration300(
                finalName, TextBox1.Text.Trim(), cycleName, TextBox21.Text, lcStr, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                ext0, ext20, ext50, ext100, ext150, ext200, ext250, ext300,
                int0, int20, int50, int100, int150, int200, int250, int300,
                depthErrorVal, TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text,
                mName, mUnc, mUncList, calibDateToSave)
        End If

        If saved Then
            MySqlCls.UpdateRegularCalibrationStatus(TextBox1.Text.Trim(), calibDateToSave, statusVal, cycleName)
            If Not String.IsNullOrEmpty(cycleName) Then
                MySqlCls.InsertResultRecord(TextBox1.Text.Trim(), finalCat, finalName, cycleName)
            End If

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
            
            MessageBox.Show("Record Saved Successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.Close()
        Else
            Dim errorMsg = "Failed to save record."
            If Not String.IsNullOrEmpty(MySqlCls.LastError) Then
                errorMsg &= Environment.NewLine & "Details: " & MySqlCls.LastError
            End If
            MessageBox.Show(errorMsg, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Function GetTrialsArrayFromUI(nominal As Decimal, isExternal As Boolean, maxRange As Decimal) As Object()
        If nominal > maxRange Then
            Return {"-", "-", "-"}
        End If

        Dim tbs = GetTextBoxesForRange(nominal, isExternal)
        Dim trials(2) As Object
        For i = 0 To 2
            If i < tbs.Count Then
                Dim val As Decimal
                If Decimal.TryParse(tbs(i).Text, val) Then
                    ' Store as error value
                    Dim error1 As Decimal = val
                    Dim error2 As Decimal = val - nominal
                    trials(i) = If(Math.Abs(error1) < Math.Abs(error2), error1, error2)
                Else
                    trials(i) = "-"
                End If
            Else
                trials(i) = "-"
            End If
        Next
        Return trials
    End Function

    ''' <summary>
    ''' TMU Step 1+2: SigmaB1 = SampleSTDEV / SQRT(n) = SQRT(sumSqDiff / (n*(n-1)))
    ''' This is the standard measurement uncertainty from repeatability.
    ''' </summary>
    Private Function calcSigmaB1(arr() As Decimal) As Decimal
        Dim n = arr.Length
        If n <= 1 Then Return 0
        Dim mean = arr.Average()
        Dim sumSqDiff = arr.Sum(Function(v) (v - mean) ^ 2)
        Dim divisor = CDec(n) * CDec(n - 1)  ' n*(n-1) = 3*2 = 6 for 3 trials
        If divisor = 0 Then Return 0
        Return CDec(Math.Sqrt(sumSqDiff / divisor))
    End Function

    Private Sub CalculateTMU()
        Try
            If _selectedMasters Is Nothing OrElse _selectedMasters.Count = 0 Then
                TextBox29.Text = ""
                TextBox29.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249)) ' Light grey/slate
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

            ' 2. Calculate MU for External and Internal across all visible ranges
            Dim maxRange As Integer = 300
            Dim sizeStr = TextBox21.Text
            If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
                Dim parts = sizeStr.Split("-"c)
                If parts.Length > 1 Then
                    Dim secondPart = parts(1).Replace("mm", "").Trim()
                    Integer.TryParse(secondPart, maxRange)
                End If
            End If

            ' Reference formula steps (per range i):
            '   Step 1: SigmaR1 = Sample STDEV of 3 trials
            '   Step 2: SigmaB1 = SigmaR1 / SQRT(n) = SQRT(sumSqDiff/(n*(n-1)))
            '   Step 3: MU_i    = SQRT(SigmaB1² + masterU²)
            '   PreTMU_ext = SQRT(Σ MU_ext_i²)
            '   PreTMU_int = SQRT(Σ MU_int_i²)
            '   TMU = SQRT(PreTMU_ext² + PreTMU_int²) × 2
            '       = SQRT(Σ MU_ext² + Σ MU_int²) × 2
            Dim sumSqMU_Ext As Decimal = 0
            Dim sumSqMU_Int As Decimal = 0

            For Each r In nominalRanges
                If r <= maxRange Then
                    ' External
                    Dim extVals As New List(Of Decimal)
                    For Each t In GetTextBoxesForRange(r, True)
                        Dim val As Decimal
                        If Decimal.TryParse(t.Text, val) Then
                            Dim err1 = val : Dim err2 = val - r
                            extVals.Add(If(Math.Abs(err1) < Math.Abs(err2), err1, err2))
                        Else
                            extVals.Add(0D)
                        End If
                    Next
                    Dim sigmaB1_e = calcSigmaB1(extVals.ToArray())
                    Dim mu_e = CDec(Math.Sqrt((sigmaB1_e ^ 2) + (masterU ^ 2)))
                    sumSqMU_Ext += (mu_e ^ 2)

                    ' Internal
                    Dim intVals As New List(Of Decimal)
                    For Each t In GetTextBoxesForRange(r, False)
                        Dim val As Decimal
                        If Decimal.TryParse(t.Text, val) Then
                            Dim err1 = val : Dim err2 = val - r
                            intVals.Add(If(Math.Abs(err1) < Math.Abs(err2), err1, err2))
                        Else
                            intVals.Add(0D)
                        End If
                    Next
                    Dim sigmaB1_i = calcSigmaB1(intVals.ToArray())
                    Dim mu_i = CDec(Math.Sqrt((sigmaB1_i ^ 2) + (masterU ^ 2)))
                    sumSqMU_Int += (mu_i ^ 2)
                End If
            Next

            ' TMU = SQRT(PreTMU_ext² + PreTMU_int²) × 2
            Dim finalTMU = CDec(Math.Sqrt(sumSqMU_Ext + sumSqMU_Int)) * 2
            TextBox29.Text = finalTMU.ToString("F9")

        Catch ex As Exception
            ' Silently handle calculation errors (e.g. during partial entry)
        End Try
    End Sub

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

    Private Sub FormatDecimal(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            tb.Text = val.ToString("F2")
        End If
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrEmpty(TargetCycle) Then TargetCycle = MySqlCls.GetActiveCycleName()
        Dim dt = MySqlCls.GetVernierCalibration300Data(TextBox1.Text.Trim(), TargetCycle)
        If dt.Rows.Count = 0 Then Return

        Dim row = dt.Rows(0)

        ' Metadata & Header
        If Not IsDBNull(row("Date")) Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(row("Date"))
        If Not IsDBNull(row("LC")) Then ComboBox5.Text = row("LC").ToString()
        If Not IsDBNull(row("Color")) Then TextBoxColor.Text = row("Color").ToString()
        If Not IsDBNull(row("Location")) Then ComboBox2.Text = row("Location").ToString()
        If Not IsDBNull(row("Temperature")) Then TextBox30.Text = row("Temperature").ToString()
        If Not IsDBNull(row("Humidity")) Then TextBox28.Text = row("Humidity").ToString()
        
        ' Observations - External
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r, True)
            For i = 1 To 3
                Dim col = $"Ext{r}_{i}"
                If dt.Columns.Contains(col) AndAlso Not IsDBNull(row(col)) AndAlso row(col).ToString() <> "-" Then
                    tbs(i - 1).Text = row(col).ToString()
                End If
            Next
        Next

        ' Observations - Internal
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r, False)
            For i = 1 To 3
                Dim col = $"Int{r}_{i}"
                If dt.Columns.Contains(col) AndAlso Not IsDBNull(row(col)) AndAlso row(col).ToString() <> "-" Then
                    tbs(i - 1).Text = row(col).ToString()
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
            _isTimeOutManual = True ' Prevent timer from overwriting
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

        ' Trigger calculations and validation coloring
        UpdateDuration(Nothing, Nothing)
        TriggerAllValidations()
        CalculateTMU()
        UpdateOverallStatus()

        ' Final TMU pre-fill (after CalculateTMU potentially cleared it)
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
                ' External
                For Each tb In GetTextBoxesForRange(r, True)
                    TrialTextChanged(tb, Nothing)
                Next
                ' Internal
                For Each tb In GetTextBoxesForRange(r, False)
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
