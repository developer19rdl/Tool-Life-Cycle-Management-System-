Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class VernierCaliper600
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = ""
    Private _category As String = "Instrument"
    Private _selectedMasters As New List(Of MasterSelectorItem)
    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    Private _isInternalMode As Boolean = False
    Private _externalData As New Dictionary(Of String, String)
    Private _internalData As New Dictionary(Of String, String)
    Private _textBoxMap As New Dictionary(Of Decimal, List(Of TextBox))
    Private _labelMap As New Dictionary(Of Decimal, TextBlock)



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

        ' Build cache for all trial textboxes and labels
        For Each r In nominalRanges
            Dim tbs As New List(Of TextBox)
            For i = 1 To 3
                Dim name = $"txt{r}_{i}"
                Dim tb = DirectCast(Me.FindName(name), TextBox)
                If tb IsNot Nothing Then
                    tbs.Add(tb)
                    tb.IsEnabled = False
                    AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                    AddHandler tb.TextChanged, AddressOf TrialTextChanged
                    AddHandler tb.LostFocus, AddressOf FormatDecimal
                End If
            Next
            _textBoxMap(r) = tbs
            
            Dim lbl = DirectCast(Me.FindName("Label" & r), TextBlock)
            If lbl IsNot Nothing Then _labelMap(r) = lbl
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

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        if _masterLimits IS Nothing Then
            txtBox.Background = Brushes.White
            Return
        End if

        Dim nominal As Decimal = GetNominalFromTextBox(txtBox)
        If nominal = -1 Then Return

        Dim prefix As String = If(_isInternalMode, "Int", "Ext")
        Dim ulCol As String = $"{prefix}{nominal}_UL"
        Dim llCol As String = $"{prefix}{nominal}_LL"

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

    Private Sub EnableNextRange(currentNominal As Decimal)
        Dim maxRange As Integer = 600
        Dim sizeStr = TextBox21.Text
        If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
            Dim parts = sizeStr.Split("-"c)
            If parts.Length > 1 Then
                Dim secondPart = parts(1).Replace("mm", "").Trim()
                Integer.TryParse(secondPart, maxRange)
            End If
        End If

        If currentNominal >= maxRange Then
            If Not _isInternalMode Then
                ToggleMode.IsChecked = True
            End If
            UpdateOverallStatus()
        End If

        Dim nominalIdx = Array.IndexOf(nominalRanges, currentNominal)

        If nominalIdx < nominalRanges.Length - 1 Then
            Dim nextNominal = nominalRanges(nominalIdx + 1)
            If nextNominal <= maxRange Then
                EnableRange(nextNominal)
            End If
        End If
    End Sub

    Private Sub ToggleMode_CheckedChanged(sender As Object, e As RoutedEventArgs) Handles ToggleMode.Checked, ToggleMode.Unchecked
        ' 1. Save current UI state to the active dictionary
        Dim sourceDict = If(_isInternalMode, _internalData, _externalData)
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For i = 0 To tbs.Count - 1
                sourceDict($"txt{r}_{i + 1}") = tbs(i).Text
            Next
        Next

        ' 2. Switch the mode variable
        _isInternalMode = If(ToggleMode.IsChecked.HasValue, ToggleMode.IsChecked.Value, False)

        ' 3. Batch UI load with simplified validation
        Dim targetDict = If(_isInternalMode, _internalData, _externalData)
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For i = 0 To tbs.Count - 1
                Dim key = $"txt{r}_{i + 1}"
                Dim tb = tbs(i)
                
                RemoveHandler tb.TextChanged, AddressOf TrialTextChanged
                tb.Text = If(targetDict.ContainsKey(key), targetDict(key), "")
                tb.Background = Brushes.White ' Reset initial background
                AddHandler tb.TextChanged, AddressOf TrialTextChanged
            Next
        Next

        RestoreEnabledState(targetDict)
        UpdateOverallStatus()
        
        ' Batch re-validation for visible fields (Fast)
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                If tb.Visibility = Visibility.Visible AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
                    ' Mimic TextChanged for coloring only - logic inside TrialTextChanged but without recursion
                    ValidateFieldColor(tb, r)
                End If
            Next
        Next
        
        CalculateTMU()
    End Sub

    Private Sub UpdateSizeLockState()
        Try
            Dim hasData As Boolean = False
            For Each r In nominalRanges
                For Each tb In GetTextBoxesForRange(r)
                    If Not String.IsNullOrWhiteSpace(tb.Text) Then
                        hasData = True
                        Exit For
                    End If
                Next
                If hasData Then Exit For
            Next
            TextBox21.IsEnabled = Not hasData
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ValidateFieldColor(tb As TextBox, nominal As Decimal)
        If _masterLimits Is Nothing Then Return
        Dim prefix As String = If(_isInternalMode, "Int", "Ext")
        Dim ulCol As String = $"{prefix}{nominal}_UL"
        Dim llCol As String = $"{prefix}{nominal}_LL"
        
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            Dim error1 As Decimal = val
            Dim error2 As Decimal = val - nominal
            Dim errorVal As Decimal = If(Math.Abs(error1) < Math.Abs(error2), error1, error2)

            Dim ul As Decimal = Convert.ToDecimal(_masterLimits(ulCol))
            Dim ll As Decimal = Convert.ToDecimal(_masterLimits(llCol))

            If errorVal >= ll AndAlso errorVal <= ul Then
                tb.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
            Else
                tb.Background = New SolidColorBrush(Color.FromRgb(255, 182, 193))
            End If
        Else
            tb.Background = Brushes.White
        End If
    End Sub

    Private Sub RestoreEnabledState(dict As Dictionary(Of String, String))
        Dim maxRange As Integer = 600
        Dim sizeStr = TextBox21.Text
        If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
            Dim parts = sizeStr.Split("-"c)
            If parts.Length > 1 Then
                Dim secondPart = parts(1).Replace("mm", "").Trim()
                Integer.TryParse(secondPart, maxRange)
            End If
        End If

        For Each r In nominalRanges
            Dim visible = (r <= maxRange)
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                tb.IsEnabled = visible
                If visible AndAlso String.IsNullOrWhiteSpace(tb.Text) Then tb.Background = Brushes.White
            Next
        Next
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

    Private Sub DisableAllObservationFields()
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For Each tb In tbs
                tb.IsEnabled = False
                tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
            Next
        Next
    End Sub

    Private Sub AdjustVisibilityBySize(sizeStr As String)
        Try
            Dim maxRange As Integer = 600
            If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
                Dim parts = sizeStr.Split("-"c)
                If parts.Length > 1 Then
                    Dim secondPart = parts(1).Replace("mm", "").Trim()
                    Integer.TryParse(secondPart, maxRange)
                End If
            End If

            For Each r In nominalRanges
                Dim visible = (r <= maxRange)
                Dim tbs = GetTextBoxesForRange(r)
                For Each tb In tbs
                    tb.IsEnabled = visible
                    tb.Background = If(visible, Brushes.White, New SolidColorBrush(Color.FromRgb(241, 245, 249)))
                    If Not visible Then tb.Clear()
                Next
                
                If _labelMap.ContainsKey(r) Then
                    _labelMap(r).Foreground = If(visible, New SolidColorBrush(Color.FromRgb(51, 65, 85)), Brushes.LightSlateGray)
                End If
            Next

            UpdateSizeLockState()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SetRangeVisibility(nominal As Decimal, visible As Boolean)
        ' Logic handled in AdjustVisibilityBySize
    End Sub

    Private Function GetTextBoxesForRange(nominal As Decimal) As List(Of TextBox)
        If _textBoxMap.ContainsKey(nominal) Then
            Return _textBoxMap(nominal)
        End If
        Return New List(Of TextBox)
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

        If _labelMap.ContainsKey(nominal) Then
            Dim lbl = _labelMap(nominal)
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
                        ComboBox2.Text = row("Location").ToString()
                        TextBoxColor.Text = row("Color").ToString()
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
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM vernier_caliper_600 WHERE RowType='MASTER' ORDER BY LC ASC")
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
            _masterLimits = MySqlCls.GetVernierMasterLimits600(selectedLC)
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
                If txtBox.Name = "TextBox30" Then
                    llCol = "Env_Temp_LL" : ulCol = "Env_Temp_UL"
                ElseIf txtBox.Name = "TextBox28" Then
                    llCol = "Env_Hum_LL" : ulCol = "Env_Hum_UL"
                ElseIf txtBox.Name = "TextBox20" Then
                    llCol = "Depth_LL" : ulCol = "Depth_UL"
                End If

                Dim ll As Decimal = Convert.ToDecimal(_masterLimits(llCol))
                Dim ul As Decimal = Convert.ToDecimal(_masterLimits(ulCol))

                If val >= ll AndAlso val <= ul Then
                    txtBox.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
                Else
                    txtBox.Background = New SolidColorBrush(Color.FromRgb(255, 182, 193))
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
            ComboBox4.SelectedIndex = 1
        Else
            ComboBox4.SelectedIndex = 0
        End If
    End Sub

    Private Function GetDecimalOrNull(txt As String) As Decimal?
        Dim val As Decimal
        If Decimal.TryParse(txt, val) Then Return val
        Return Nothing
    End Function

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs) Handles Button2.Click
        Dim maxRange As Integer = 600
        Dim sizeStr = TextBox21.Text
        If Not String.IsNullOrEmpty(sizeStr) AndAlso sizeStr.Contains("-") Then
            Dim parts = sizeStr.Split("-"c)
            If parts.Length > 1 Then
                Dim secondPart = parts(1).Replace("mm", "").Trim()
                Integer.TryParse(secondPart, maxRange)
            End If
        End If

        ' Validation: Partial triad check (If user starts typing, they must complete all 3)
        For Each r In nominalRanges
            If r <= maxRange Then
                Dim tbs = GetTextBoxesForRange(r)
                Dim count = tbs.Where(Function(t) Not String.IsNullOrWhiteSpace(t.Text)).Count()
                If count > 0 AndAlso count < 3 Then
                    MessageBox.Show($"Please complete all 3 trials for {r}mm observation.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
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



        Dim statusVal = DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString()

        ' Sync current UI to active dictionary before generating arrays
        Dim sourceDict = If(_isInternalMode, _internalData, _externalData)
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For i = 0 To 2
                If i < tbs.Count Then
                    Dim key = $"txt{r}_{i + 1}"
                    sourceDict(key) = tbs(i).Text
                End If
            Next
        Next

        Dim ext0 = GetTrialsArrayFromDict(0, maxRange, _externalData)
        Dim ext20 = GetTrialsArrayFromDict(20, maxRange, _externalData)
        Dim ext50 = GetTrialsArrayFromDict(50, maxRange, _externalData)
        Dim ext100 = GetTrialsArrayFromDict(100, maxRange, _externalData)
        Dim ext150 = GetTrialsArrayFromDict(150, maxRange, _externalData)
        Dim ext200 = GetTrialsArrayFromDict(200, maxRange, _externalData)
        Dim ext250 = GetTrialsArrayFromDict(250, maxRange, _externalData)
        Dim ext300 = GetTrialsArrayFromDict(300, maxRange, _externalData)
        Dim ext350 = GetTrialsArrayFromDict(350, maxRange, _externalData)
        Dim ext400 = GetTrialsArrayFromDict(400, maxRange, _externalData)
        Dim ext450 = GetTrialsArrayFromDict(450, maxRange, _externalData)
        Dim ext500 = GetTrialsArrayFromDict(500, maxRange, _externalData)
        Dim ext550 = GetTrialsArrayFromDict(550, maxRange, _externalData)
        Dim ext600 = GetTrialsArrayFromDict(600, maxRange, _externalData)

        Dim int0 = GetTrialsArrayFromDict(0, maxRange, _internalData)
        Dim int20 = GetTrialsArrayFromDict(20, maxRange, _internalData)
        Dim int50 = GetTrialsArrayFromDict(50, maxRange, _internalData)
        Dim int100 = GetTrialsArrayFromDict(100, maxRange, _internalData)
        Dim int150 = GetTrialsArrayFromDict(150, maxRange, _internalData)
        Dim int200 = GetTrialsArrayFromDict(200, maxRange, _internalData)
        Dim int250 = GetTrialsArrayFromDict(250, maxRange, _internalData)
        Dim int300 = GetTrialsArrayFromDict(300, maxRange, _internalData)
        Dim int350 = GetTrialsArrayFromDict(350, maxRange, _internalData)
        Dim int400 = GetTrialsArrayFromDict(400, maxRange, _internalData)
        Dim int450 = GetTrialsArrayFromDict(450, maxRange, _internalData)
        Dim int500 = GetTrialsArrayFromDict(500, maxRange, _internalData)
        Dim int550 = GetTrialsArrayFromDict(550, maxRange, _internalData)
        Dim int600 = GetTrialsArrayFromDict(600, maxRange, _internalData)

        Dim depthErrorVal As Object = GetDecimalOrNull(TextBox20.Text)
        If depthErrorVal Is Nothing Then depthErrorVal = "-"

        Dim lcStr = ComboBox5.SelectedItem.ToString()
        Dim calibDateToSave = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim cycleName = MySqlCls.GetActiveCycleName()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, calibDateToSave.ToShortDateString(), TextBox30.Text, TextBox28.Text, TextBox24.Text, TextBox20.Text)

        Dim freshDetails = GetFreshInstrumentDetails(TextBox1.Text.Trim())
        Dim finalCat = freshDetails.Category
        Dim finalName = If(String.IsNullOrEmpty(freshDetails.Name), _instrumentName, freshDetails.Name)

        Dim mName = ""
        Dim mUnc As Decimal = 0
        Dim mUncList = ""

        Dim saved As Boolean = False
        If IsEditMode Then
            saved = MySqlCls.UpdateVernierCalibration600(
                finalName, TextBox1.Text.Trim(), TargetCycle, TextBox21.Text, lcStr, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                ext0, ext20, ext50, ext100, ext150, ext200, ext250, ext300, ext350, ext400, ext450, ext500, ext550, ext600,
                int0, int20, int50, int100, int150, int200, int250, int300, int350, int400, int450, int500, int550, int600,
                depthErrorVal, TextBox24.Text, TextBox26.Text, TextBox25.Text, statusVal, TextBox19.Text,
                mName, mUnc, mUncList, calibDateToSave)
        Else
            saved = MySqlCls.InsertVernierCalibration600(
                finalName, TextBox1.Text.Trim(), cycleName, TextBox21.Text, lcStr, TextBoxColor.Text, ComboBox2.Text, TextBox30.Text, TextBox28.Text, TextBox29.Text,
                ext0, ext20, ext50, ext100, ext150, ext200, ext250, ext300, ext350, ext400, ext450, ext500, ext550, ext600,
                int0, int20, int50, int100, int150, int200, int250, int300, int350, int400, int450, int500, int550, int600,
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

    Private Function GetTrialsArrayFromUI(nominal As Decimal, maxRange As Decimal) As Object()
        If nominal > maxRange Then Return {"-", "-", "-"}
        Dim tbs = GetTextBoxesForRange(nominal)
        Dim trials(2) As Object
        For i = 0 To 2
            If i < tbs.Count Then
                Dim val As Decimal
                If Decimal.TryParse(tbs(i).Text, val) Then
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

    Private Function GetTrialsArrayFromDict(nominal As Decimal, maxRange As Decimal, dict As Dictionary(Of String, String)) As Object()
        If nominal > maxRange Then Return {"-", "-", "-"}
        Dim trials(2) As Object
        For i = 1 To 3
            Dim key = $"txt{nominal}_{i}"
            Dim textVal = If(dict.ContainsKey(key), dict(key), "")
            Dim val As Decimal
            If Decimal.TryParse(textVal, val) Then
                Dim error1 As Decimal = val
                Dim error2 As Decimal = val - nominal
                trials(i - 1) = If(Math.Abs(error1) < Math.Abs(error2), error1, error2)
            Else
                trials(i - 1) = "-"
            End If
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
            For Each m In _selectedMasters
                sumSqMaster += (m.MasterUncertainty ^ 2)
            Next
            Dim masterU As Decimal = Math.Sqrt(sumSqMaster)

            ' 2. Per-range MU, accumulated for Ext and Int separately
            ' Formula: SigmaB1 = SigmaR1 / SQRT(n) = SQRT(SumSqDiff / (n*(n-1)))
            '          MU_i    = SQRT(SigmaB1_i^2 + masterU^2)
            '          PreTMU_ext = SQRT(SUM MU_ext_i^2)
            '          PreTMU_int = SQRT(SUM MU_int_i^2)
            '          TMU = SQRT(PreTMU_ext^2 + PreTMU_int^2) * 2
            Dim sumSqMU_Ext As Decimal = 0
            Dim sumSqMU_Int As Decimal = 0
            Dim maxRange As Integer = 600

            For Each r In nominalRanges
                If r > maxRange Then Exit For

                ' --- External ---
                Dim extVals As New List(Of Decimal)
                For i = 1 To 3
                    Dim textVal As String = ""
                    If Not _isInternalMode Then
                        Dim tbs = GetTextBoxesForRange(r)
                        If tbs.Count >= i Then textVal = tbs(i - 1).Text
                    Else
                        Dim key = $"txt{r}_{i}"
                        If _externalData.ContainsKey(key) Then textVal = _externalData(key)
                    End If
                    Dim val As Decimal
                    If Decimal.TryParse(textVal, val) Then
                        Dim err1 = val : Dim err2 = val - r
                        extVals.Add(If(Math.Abs(err1) < Math.Abs(err2), err1, err2))
                    Else
                        extVals.Add(0D)
                    End If
                Next
                Dim sigB1_e = calcSigmaB1(extVals.ToArray())
                sumSqMU_Ext += (sigB1_e ^ 2) + (masterU ^ 2)

                ' --- Internal ---
                Dim intVals As New List(Of Decimal)
                For i = 1 To 3
                    Dim textVal As String = ""
                    If _isInternalMode Then
                        Dim tbs = GetTextBoxesForRange(r)
                        If tbs.Count >= i Then textVal = tbs(i - 1).Text
                    Else
                        Dim key = $"txt{r}_{i}"
                        If _internalData.ContainsKey(key) Then textVal = _internalData(key)
                    End If
                    Dim val As Decimal
                    If Decimal.TryParse(textVal, val) Then
                        Dim err1 = val : Dim err2 = val - r
                        intVals.Add(If(Math.Abs(err1) < Math.Abs(err2), err1, err2))
                    Else
                        intVals.Add(0D)
                    End If
                Next
                Dim sigB1_i = calcSigmaB1(intVals.ToArray())
                sumSqMU_Int += (sigB1_i ^ 2) + (masterU ^ 2)
            Next

            Dim finalTMU = CDec(Math.Sqrt(sumSqMU_Ext + sumSqMU_Int)) * 2
            TextBox29.Text = finalTMU.ToString("F9")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FormatDecimal(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            tb.Text = val.ToString("F2")
        End If
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrEmpty(TargetCycle) Then TargetCycle = MySqlCls.GetActiveCycleName()
        Dim dt = MySqlCls.GetVernierCalibration600Data(TextBox1.Text.Trim(), TargetCycle)
        If dt.Rows.Count = 0 Then Return

        Dim row = dt.Rows(0)

        ' Metadata
        If Not IsDBNull(row("Date")) Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(row("Date"))
        If Not IsDBNull(row("LC")) Then ComboBox5.Text = row("LC").ToString()
        If Not IsDBNull(row("Color")) Then TextBoxColor.Text = row("Color").ToString()
        If Not IsDBNull(row("Location")) Then ComboBox2.Text = row("Location").ToString()
        If Not IsDBNull(row("Temperature")) Then TextBox30.Text = row("Temperature").ToString()
        If Not IsDBNull(row("Humidity")) Then TextBox28.Text = row("Humidity").ToString()
        
        ' Populating dictionaries for dual-mode switching
        For Each r In nominalRanges
            For i = 1 To 3
                Dim extCol = $"Ext{r}_{i}"
                If dt.Columns.Contains(extCol) AndAlso Not IsDBNull(row(extCol)) AndAlso row(extCol).ToString() <> "-" Then
                    _externalData($"txt{r}_{i}") = row(extCol).ToString()
                End If

                Dim intCol = $"Int{r}_{i}"
                If dt.Columns.Contains(intCol) AndAlso Not IsDBNull(row(intCol)) AndAlso row(intCol).ToString() <> "-" Then
                    _internalData($"txt{r}_{i}") = row(intCol).ToString()
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

        ' Trigger load current mode (External by default)
        _isInternalMode = False
        RestoreEnabledState(_externalData)
        
        ' Load external data into textboxes
        For Each r In nominalRanges
            Dim tbs = GetTextBoxesForRange(r)
            For i = 1 To 3
                Dim key = $"txt{r}_{i}"
                If _externalData.ContainsKey(key) Then
                    tbs(i - 1).Text = _externalData(key)
                End If
            Next
        Next

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

    Private Sub UpdateDuration(sender As Object, e As TextChangedEventArgs)
        Dim timeInStr = TextBox24.Text.Trim()
        Dim timeOutStr = TextBox26.Text.Trim()
        Dim dummyTimeOut = timeOutStr
        If String.IsNullOrEmpty(dummyTimeOut) Then dummyTimeOut = DateTime.Now.ToString("HH:mm")
        
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

    Private Function CalculateDuration(timeIn As String, timeOut As String) As Integer
        Try
            Dim tIn, tOut As DateTime
            Dim formats As String() = {"HH:mm", "H:mm"}
            If DateTime.TryParseExact(timeIn.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, tIn) AndAlso
               DateTime.TryParseExact(timeOut.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, tOut) Then
                If tOut < tIn Then tOut = tOut.AddDays(1)
                Return CInt((tOut - tIn).TotalMinutes)
            End If
        Catch ex As Exception
        End Try
        Return -1
    End Function

    Private Sub TriggerAllValidations()
        Try
            ' Environment
            EnvTextChanged(TextBox30, Nothing)
            EnvTextChanged(TextBox28, Nothing)
            EnvTextChanged(TextBox20, Nothing) ' Depth Error
            
            ' Observations - DUAL MODE: We need to trigger for the CURRENTLY visible mode (External by default on load)
            ' However, if we want to be thorough, we should trigger for the controls themselves.
            For Each r In nominalRanges
                For Each tb In GetTextBoxesForRange(r)
                    TrialTextChanged(tb, Nothing)
                Next
            Next
            
            UpdateOverallStatus()
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

