Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq
Imports System
Imports System.Windows.Input


Public Class ThreadPlugGaugeCalibration
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = "Thread Plug Gauge"
    Private _category As String = "Gauge"
    Private _selectedMasters As New System.Collections.Generic.List(Of MasterSelectorItem)()
    Private _lc As String = "0.000001"

    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _isPopulating As Boolean = False

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    Public Sub New()
        InitializeComponent()

        ' Wire up soaking time duration calculation
        AddHandler TxtTimeIn.TextChanged, AddressOf UpdateDuration
        AddHandler TxtTimeOut.TextChanged, AddressOf UpdateDuration
        AddHandler TxtTimeOut.GotFocus, Sub() _isTimeOutManual = True
        AddHandler TxtTimeOut.PreviewKeyDown, Sub() _isTimeOutManual = True
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ' Populate Locations & Colors from database for suggestions
        Try
            ComboLocation.Items.Clear()
            Dim dtDepts = MySqlCls.GetDepartments()
            For Each row As DataRow In dtDepts.Rows
                ComboLocation.Items.Add(row("DepartmentName").ToString())
            Next
        Catch ex As Exception
        End Try

        Try
            ComboColour.Items.Clear()
            Dim dtColors = MySqlCls.ReadDatatable("SELECT DISTINCT Color FROM department_list WHERE Color IS NOT NULL AND Color != ''")
            For Each row As DataRow In dtColors.Rows
                ComboColour.Items.Add(row("Color").ToString())
            Next
        Catch ex As Exception
        End Try

        ' Wire up Category
        Try
            ComboCategory.Items.Clear()
            ComboCategory.Items.Add("Instrument")
            ComboCategory.Items.Add("Gauge")
            ComboCategory.Text = _category
        Catch ex As Exception
        End Try

        ' Populate LC from database
        Try
            ComboLC.Items.Clear()
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM thread_plug_gauge_calibration WHERE RowType='MASTER' ORDER BY LC ASC")
            For Each row As DataRow In dt.Rows
                ComboLC.Items.Add(row("LC").ToString())
            Next
            ComboLC.Text = _lc
        Catch ex As Exception
        End Try

        ' Attach TextChanged validation handlers
        AddHandler TxtTemp.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TxtHumidity.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TxtMU.TextChanged, Sub() CalculateTMU()

        ' Observations TextChanged Handlers
        Dim trialBoxes = {TxtGo1, TxtGo2, TxtGo3, TxtGo4, TxtNoGo1, TxtNoGo2, TxtNoGo3, TxtNoGo4}
        For Each tb In trialBoxes
            If tb IsNot Nothing Then
                AddHandler tb.TextChanged, AddressOf TrialTextChanged
            End If
        Next

        ' Limits TextChanged Handlers to re-trigger validation on observation fields
        AddHandler TxtGoMin.TextChanged, Sub()
                                             TrialTextChanged(TxtGo1, Nothing)
                                             TrialTextChanged(TxtGo2, Nothing)
                                             TrialTextChanged(TxtGo3, Nothing)
                                             TrialTextChanged(TxtGo4, Nothing)
                                         End Sub
        AddHandler TxtGoMax.TextChanged, Sub()
                                             TrialTextChanged(TxtGo1, Nothing)
                                             TrialTextChanged(TxtGo2, Nothing)
                                             TrialTextChanged(TxtGo3, Nothing)
                                             TrialTextChanged(TxtGo4, Nothing)
                                         End Sub
        AddHandler TxtNoGoMin.TextChanged, Sub()
                                               TrialTextChanged(TxtNoGo1, Nothing)
                                               TrialTextChanged(TxtNoGo2, Nothing)
                                               TrialTextChanged(TxtNoGo3, Nothing)
                                               TrialTextChanged(TxtNoGo4, Nothing)
                                           End Sub
        AddHandler TxtNoGoMax.TextChanged, Sub()
                                               TrialTextChanged(TxtNoGo1, Nothing)
                                               TrialTextChanged(TxtNoGo2, Nothing)
                                               TrialTextChanged(TxtNoGo3, Nothing)
                                               TrialTextChanged(TxtNoGo4, Nothing)
                                           End Sub

        ' Attach LostFocus for Control No
        AddHandler TxtControlNo.LostFocus, AddressOf TxtControlNo_LostFocus

        If Not String.IsNullOrWhiteSpace(TxtControlNo.Text) Then
            LoadInstrumentDetails(TxtControlNo.Text.Trim())
        End If

        If Not IsEditMode Then
            TxtCalibrationDate.SelectedDate = DateTime.Today
        End If

        ' Attach Enter-key navigation handlers
        Dim navTextBoxes = {TxtGoSize, TxtGoMin, TxtGoMax, TxtGo1, TxtGo2, TxtGo3, TxtGo4,
                            TxtNoGoSize, TxtNoGoMin, TxtNoGoMax, TxtNoGo1, TxtNoGo2, TxtNoGo3, TxtNoGo4}
        For Each tb In navTextBoxes
            If tb IsNot Nothing Then
                AddHandler tb.PreviewKeyDown, AddressOf TextBox_PreviewKeyDown
            End If
        Next

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

    Private Sub ComboLC_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboLC.SelectionChanged
        If _isPopulating Then Return
        If ComboLC.SelectedItem IsNot Nothing Then
            _lc = ComboLC.SelectedItem.ToString()
            FetchMasterLimits()
        End If
    End Sub

    Private Sub TxtControlNo_LostFocus(sender As Object, e As RoutedEventArgs)
        If Not String.IsNullOrWhiteSpace(TxtControlNo.Text) Then
            LoadInstrumentDetails(TxtControlNo.Text.Trim())
        End If
    End Sub

    Private Sub LoadInstrumentDetails(controlNo As String)
        _isPopulating = True
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
                    ComboCategory.Text = _category

                    Dim matchedSize As String = ""
                    If row.Table.Columns.Contains("Size") AndAlso Not IsDBNull(row("Size")) Then
                        matchedSize = row("Size").ToString()
                    End If

                    If MatchedSize <> "" AndAlso IsEditMode Then
                        TxtGoSize.Text = matchedSize
                        TxtNoGoSize.Text = matchedSize
                    ElseIf Not IsEditMode Then
                        TxtGoSize.Text = matchedSize
                        TxtNoGoSize.Text = matchedSize
                    End If

                    If row.Table.Columns.Contains("LC") AndAlso Not IsDBNull(row("LC")) Then
                        _lc = row("LC").ToString()
                    ElseIf row.Table.Columns.Contains("LeastCount") AndAlso Not IsDBNull(row("LeastCount")) Then
                        _lc = row("LeastCount").ToString()
                    End If
                    ComboLC.Text = _lc

                    If dtDept.Rows.Count > 0 Then
                        Dim deptRow = dtDept.Rows(0)
                        ComboLocation.Text = deptRow("Department").ToString()
                        ComboColour.Text = deptRow("Color").ToString()
                    Else
                        If row.Table.Columns.Contains("Location") Then ComboLocation.Text = row("Location").ToString()
                        If row.Table.Columns.Contains("Color") Then ComboColour.Text = row("Color").ToString()
                    End If

                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), If(row.Table.Columns.Contains("InstrumentName"), row("InstrumentName").ToString(), "Thread Plug Gauge"), If(row.Table.Columns.Contains("GaugeName"), row("GaugeName").ToString(), "Thread Plug Gauge"))

                    FetchMasterLimits()
                    Exit For
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("LoadInstrumentDetails Error: " & ex.Message)
        End Try
        _isPopulating = False
    End Sub

    Private Sub FetchMasterLimits()
        Try
            _masterLimits = MySqlCls.GetThreadPlugGaugeMasterLimits(_lc)
            If _masterLimits IsNot Nothing AndAlso IsEditMode Then
                ' Populate limits UI if controls exist
                If _masterLimits.Table.Columns.Contains("Go_Min_Limit") Then
                    TxtGoMin.Text = _masterLimits("Go_Min_Limit").ToString()
                End If
                If _masterLimits.Table.Columns.Contains("Go_Max_Limit") Then
                    TxtGoMax.Text = _masterLimits("Go_Max_Limit").ToString()
                End If
                If _masterLimits.Table.Columns.Contains("NoGo_Min_Limit") Then
                    TxtNoGoMin.Text = _masterLimits("NoGo_Min_Limit").ToString()
                End If
                If _masterLimits.Table.Columns.Contains("NoGo_Max_Limit") Then
                    TxtNoGoMax.Text = _masterLimits("NoGo_Max_Limit").ToString()
                End If
            ElseIf Not IsEditMode Then
                If _masterLimits IsNot Nothing Then
                    If _masterLimits.Table.Columns.Contains("Go_Min_Limit") Then TxtGoMin.Text = _masterLimits("Go_Min_Limit").ToString()
                    If _masterLimits.Table.Columns.Contains("Go_Max_Limit") Then TxtGoMax.Text = _masterLimits("Go_Max_Limit").ToString()
                    If _masterLimits.Table.Columns.Contains("NoGo_Min_Limit") Then TxtNoGoMin.Text = _masterLimits("NoGo_Min_Limit").ToString()
                    If _masterLimits.Table.Columns.Contains("NoGo_Max_Limit") Then TxtNoGoMax.Text = _masterLimits("NoGo_Max_Limit").ToString()
                Else
                    TxtGoMin.Text = ""
                    TxtGoMax.Text = ""
                    TxtNoGoMin.Text = ""
                    TxtNoGoMax.Text = ""
                End If
            End If

            ' Refresh validation coloring
            EnvironmentTextChanged(TxtTemp, Nothing)
            EnvironmentTextChanged(TxtHumidity, Nothing)

            Dim trialBoxes = {TxtGo1, TxtGo2, TxtGo3, TxtGo4, TxtNoGo1, TxtNoGo2, TxtNoGo3, TxtNoGo4}
            For Each tb In trialBoxes
                If tb IsNot Nothing Then TrialTextChanged(tb, Nothing)
            Next
            CalculateTMU()
        Catch ex As Exception
            Console.WriteLine("FetchMasterLimits Error: " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrWhiteSpace(TxtControlNo.Text) OrElse String.IsNullOrWhiteSpace(TargetCycle) Then Return

        _isPopulating = True
        Try
            Dim dt = MySqlCls.GetThreadPlugGaugeCalibrationData(TxtControlNo.Text.Trim(), TargetCycle)
            If dt.Rows.Count > 0 Then
                Dim dr = dt.Rows(0)
                If Not IsDBNull(dr("Date")) Then
                    TxtCalibrationDate.SelectedDate = Convert.ToDateTime(dr("Date"))
                End If

                If Not IsDBNull(dr("Type")) Then ComboCategory.Text = dr("Type").ToString()
                If Not IsDBNull(dr("Color")) Then ComboColour.Text = dr("Color").ToString()
                If Not IsDBNull(dr("Location")) Then ComboLocation.Text = dr("Location").ToString()
                If Not IsDBNull(dr("Temperature")) Then TxtTemp.Text = dr("Temperature").ToString()
                If Not IsDBNull(dr("Humidity")) Then TxtHumidity.Text = dr("Humidity").ToString()
                If Not IsDBNull(dr("TMU")) Then TxtTMU.Text = dr("TMU").ToString()
                If dr.Table.Columns.Contains("MU") AndAlso Not IsDBNull(dr("MU")) Then
                    TxtMU.Text = dr("MU").ToString()
                End If

                If dr.Table.Columns.Contains("LC") AndAlso Not IsDBNull(dr("LC")) Then
                    _lc = dr("LC").ToString()
                    ComboLC.Text = _lc
                End If

                If Not IsDBNull(dr("Size")) Then
                    TxtGoSize.Text = dr("Size").ToString()
                    TxtNoGoSize.Text = dr("Size").ToString()
                End If
                If Not IsDBNull(dr("Go_Min_Limit")) Then TxtGoMin.Text = dr("Go_Min_Limit").ToString()
                If Not IsDBNull(dr("Go_Max_Limit")) Then TxtGoMax.Text = dr("Go_Max_Limit").ToString()
                If Not IsDBNull(dr("Go_Obs_1")) Then TxtGo1.Text = dr("Go_Obs_1").ToString()
                If Not IsDBNull(dr("Go_Obs_2")) Then TxtGo2.Text = dr("Go_Obs_2").ToString()
                If Not IsDBNull(dr("Go_Obs_3")) Then TxtGo3.Text = dr("Go_Obs_3").ToString()
                If Not IsDBNull(dr("Go_Obs_4")) Then TxtGo4.Text = dr("Go_Obs_4").ToString()

                If Not IsDBNull(dr("NoGo_Min_Limit")) Then TxtNoGoMin.Text = dr("NoGo_Min_Limit").ToString()
                If Not IsDBNull(dr("NoGo_Max_Limit")) Then TxtNoGoMax.Text = dr("NoGo_Max_Limit").ToString()
                If Not IsDBNull(dr("NoGo_Obs_1")) Then TxtNoGo1.Text = dr("NoGo_Obs_1").ToString()
                If Not IsDBNull(dr("NoGo_Obs_2")) Then TxtNoGo2.Text = dr("NoGo_Obs_2").ToString()
                If Not IsDBNull(dr("NoGo_Obs_3")) Then TxtNoGo3.Text = dr("NoGo_Obs_3").ToString()
                If Not IsDBNull(dr("NoGo_Obs_4")) Then TxtNoGo4.Text = dr("NoGo_Obs_4").ToString()

                If Not IsDBNull(dr("TimeIn")) Then TxtTimeIn.Text = dr("TimeIn").ToString()
                If Not IsDBNull(dr("TimeOut")) Then
                    TxtTimeOut.Text = dr("TimeOut").ToString()
                    _isTimeOutManual = True
                End If
                If Not IsDBNull(dr("TotalTime")) Then TxtTotalTime.Text = dr("TotalTime").ToString()
                If Not IsDBNull(dr("Remark")) Then TxtRemark.Text = dr("Remark").ToString()

                If Not IsDBNull(dr("Status")) Then
                    ComboJudgement.Text = dr("Status").ToString()
                End If

                ' Load selected masters
                Dim cached = CalibrationMasterCache.LoadMasters(TxtControlNo.Text.Trim(), TargetCycle)
                If cached.Count > 0 Then
                    _selectedMasters = MySqlCls.GetMastersByDescriptions(cached)
                    UpdateMasterUI()
                End If
                CalculateTMU()
            End If
        Catch ex As Exception
            Console.WriteLine("PopulateCalibrationData Error: " & ex.Message)
        Finally
            _isPopulating = False
        End Try

        ' Trigger coloring and validation after _isPopulating is set to False
        Try
            EnvironmentTextChanged(TxtTemp, Nothing)
            EnvironmentTextChanged(TxtHumidity, Nothing)

            Dim trialBoxes = {TxtGo1, TxtGo2, TxtGo3, TxtGo4, TxtNoGo1, TxtNoGo2, TxtNoGo3, TxtNoGo4}
            For Each tb In trialBoxes
                If tb IsNot Nothing Then TrialTextChanged(tb, Nothing)
            Next
        Catch ex As Exception
            Console.WriteLine("PopulateCalibrationData Revalidation Error: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnBrowseMaster_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMasters.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Window.GetWindow(Me)
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
        If Not _isTimeOutManual AndAlso Not IsEditMode Then
            TxtTimeOut.Text = DateTime.Now.ToString("HH:mm")
            UpdateDuration(Nothing, Nothing)
        End If
    End Sub

    Private Sub UpdateDuration(sender As Object, e As TextChangedEventArgs)
        Dim timeInStr = TxtTimeIn.Text.Trim()
        Dim timeOutStr = TxtTimeOut.Text.Trim()

        Dim dummyTimeOut = timeOutStr
        If String.IsNullOrEmpty(dummyTimeOut) Then
            dummyTimeOut = DateTime.Now.ToString("HH:mm")
        End If

        Dim totalMinutes = CalculateDuration(timeInStr, dummyTimeOut)
        
        If totalMinutes >= 0 Then
            Dim hours = totalMinutes \ 60
            Dim mins = totalMinutes Mod 60
            TxtTotalTime.Text = $"{hours:D2}:{mins:D2}"
            
            If totalMinutes < 120 Then
                TxtTotalTime.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' Light Pink
            Else
                TxtTotalTime.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' Light Green
            End If
        Else
            TxtTotalTime.Text = ""
            TxtTotalTime.Background = Brushes.White
        End If

        UpdateJudgement()
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
        If _isPopulating Then Return
        Dim tb = TryCast(sender, TextBox)
        If tb Is Nothing Then Return

        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            Dim isOK = True
            If tb.Name = "TxtTemp" Then
                Dim ll = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Temp_LL") AndAlso Not IsDBNull(_masterLimits("Env_Temp_LL")), Convert.ToDecimal(_masterLimits("Env_Temp_LL")), 18.0)
                Dim ul = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Temp_UL") AndAlso Not IsDBNull(_masterLimits("Env_Temp_UL")), Convert.ToDecimal(_masterLimits("Env_Temp_UL")), 22.0)
                If val < ll OrElse val > ul Then isOK = False
            ElseIf tb.Name = "TxtHumidity" Then
                Dim ll = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_LL") AndAlso Not IsDBNull(_masterLimits("Env_Hum_LL")), Convert.ToDecimal(_masterLimits("Env_Hum_LL")), 40.0)
                Dim ul = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_UL") AndAlso Not IsDBNull(_masterLimits("Env_Hum_UL")), Convert.ToDecimal(_masterLimits("Env_Hum_UL")), 60.0)
                If val < ll OrElse val > ul Then isOK = False
            End If

            If Not isOK Then
                tb.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226))
                tb.Foreground = Brushes.Red
            Else
                tb.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231))
                tb.Foreground = Brushes.Black
            End If
        Else
            tb.Background = Brushes.White
            tb.Foreground = Brushes.Black
        End If

        UpdateJudgement()
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim tb = TryCast(sender, TextBox)
        If tb Is Nothing Then Return

        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            Dim isOK = True
            Dim ul As Decimal = 0
            Dim ll As Decimal = 0
            
            If tb.Name.StartsWith("TxtGo") Then
                Dim minVal As Decimal
                Dim maxVal As Decimal
                ll = If(Decimal.TryParse(TxtGoMin.Text, minVal), minVal, -5.0)
                ul = If(Decimal.TryParse(TxtGoMax.Text, maxVal), maxVal, 5.0)
            ElseIf tb.Name.StartsWith("TxtNoGo") Then
                Dim minVal As Decimal
                Dim maxVal As Decimal
                ll = If(Decimal.TryParse(TxtNoGoMin.Text, minVal), minVal, -5.0)
                ul = If(Decimal.TryParse(TxtNoGoMax.Text, maxVal), maxVal, 5.0)
            End If
            
            If val < ll OrElse val > ul Then
                tb.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226))
                tb.Foreground = Brushes.Red
            Else
                tb.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231))
                tb.Foreground = Brushes.Black
            End If
        Else
            tb.Background = Brushes.White
            tb.Foreground = Brushes.Black
        End If

        UpdateJudgement()
        CalculateTMU()
    End Sub

    Private Sub UpdateJudgement()
        If _isPopulating Then Return

        Dim isAnyNG As Boolean = False

        ' Check Temp
        Dim tempVal As Decimal
        If Decimal.TryParse(TxtTemp.Text, tempVal) Then
            Dim ll = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Temp_LL") AndAlso Not IsDBNull(_masterLimits("Env_Temp_LL")), Convert.ToDecimal(_masterLimits("Env_Temp_LL")), 18.0)
            Dim ul = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Temp_UL") AndAlso Not IsDBNull(_masterLimits("Env_Temp_UL")), Convert.ToDecimal(_masterLimits("Env_Temp_UL")), 22.0)
            If tempVal < ll OrElse tempVal > ul Then isAnyNG = True
        Else
            isAnyNG = True
        End If

        ' Check Humidity
        Dim humVal As Decimal
        If Decimal.TryParse(TxtHumidity.Text, humVal) Then
            Dim ll = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_LL") AndAlso Not IsDBNull(_masterLimits("Env_Hum_LL")), Convert.ToDecimal(_masterLimits("Env_Hum_LL")), 40.0)
            Dim ul = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_UL") AndAlso Not IsDBNull(_masterLimits("Env_Hum_UL")), Convert.ToDecimal(_masterLimits("Env_Hum_UL")), 60.0)
            If humVal < ll OrElse humVal > ul Then isAnyNG = True
        Else
            isAnyNG = True
        End If

        ' Check Go/Nogo Observations
        Dim trialBoxes = {TxtGo1, TxtGo2, TxtGo3, TxtGo4, TxtNoGo1, TxtNoGo2, TxtNoGo3, TxtNoGo4}
        For Each tb In trialBoxes
            If tb IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
                Dim val As Decimal
                If Decimal.TryParse(tb.Text, val) Then
                    Dim ul As Decimal = 0
                    Dim ll As Decimal = 0
                    If tb.Name.StartsWith("TxtGo") Then
                        Dim minVal As Decimal
                        Dim maxVal As Decimal
                        ll = If(Decimal.TryParse(TxtGoMin.Text, minVal), minVal, -5.0)
                        ul = If(Decimal.TryParse(TxtGoMax.Text, maxVal), maxVal, 5.0)
                    Else
                        Dim minVal As Decimal
                        Dim maxVal As Decimal
                        ll = If(Decimal.TryParse(TxtNoGoMin.Text, minVal), minVal, -5.0)
                        ul = If(Decimal.TryParse(TxtNoGoMax.Text, maxVal), maxVal, 5.0)
                    End If
                    If val < ll OrElse val > ul Then
                        isAnyNG = True
                    End If
                Else
                    isAnyNG = True
                End If
            End If
        Next

        ' Check Duration (Time)
        Dim durationText = TxtTotalTime.Text.Trim()
        If Not String.IsNullOrEmpty(durationText) AndAlso durationText.Contains(":") Then
            Dim parts = durationText.Split(":"c)
            Dim hrs, mns As Integer
            If Integer.TryParse(parts(0), hrs) AndAlso Integer.TryParse(parts(1), mns) Then
                Dim totalMins = hrs * 60 + mns
                If totalMins < 120 Then
                    isAnyNG = True
                End If
            End If
        Else
            isAnyNG = True
        End If

        If isAnyNG Then
            ComboJudgement.SelectedIndex = 1 ' NG
        Else
            ComboJudgement.SelectedIndex = 0 ' OK
        End If
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(TxtControlNo.Text) Then
            MessageBox.Show("Control No cannot be empty.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Duration Validation - must be at least 2 hours
        Dim _timeInChk = TxtTimeIn.Text.Trim()
        Dim _timeOutChk = TxtTimeOut.Text.Trim()
        If String.IsNullOrEmpty(_timeOutChk) Then _timeOutChk = DateTime.Now.ToString("HH:mm")
        Dim _durChk = CalculateDuration(_timeInChk, _timeOutChk)
        If _durChk >= 0 AndAlso _durChk < 120 Then
            MessageBox.Show("Calibration duration must be at least 2 hours.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim type = ComboCategory.Text
        Dim controlNo = TxtControlNo.Text.Trim()
        Dim cycleName = MySqlCls.GetActiveCycleName()
        Dim size = TxtGoSize.Text
        Dim lc = If(ComboLC.SelectedItem IsNot Nothing, ComboLC.SelectedItem.ToString(), ComboLC.Text)
        Dim color = ComboColour.Text
        Dim location = ComboLocation.Text
        Dim temp = TxtTemp.Text
        Dim humidity = TxtHumidity.Text
        Dim tmu = TxtTMU.Text
        Dim mu = TxtMU.Text
        Dim remark = TxtRemark.Text
        Dim timeIn = TxtTimeIn.Text
        Dim timeOut = TxtTimeOut.Text
        Dim totalTime = TxtTotalTime.Text
        Dim status = ComboJudgement.Text
        Dim parameter = "Diameter"
        Dim locationObs = ""

        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, calibDate.ToShortDateString(), TxtTemp.Text, TxtHumidity.Text, TxtTimeIn.Text, "")

        Dim tempGoMin, tempGoMax, tempNoGoMin, tempNoGoMax As Decimal
        Dim go_min As Object = If(Decimal.TryParse(TxtGoMin.Text, tempGoMin), tempGoMin, DBNull.Value)
        Dim go_max As Object = If(Decimal.TryParse(TxtGoMax.Text, tempGoMax), tempGoMax, DBNull.Value)
        Dim goDetails = {go_min, go_max, TxtGo1.Text, TxtGo2.Text, TxtGo3.Text, TxtGo4.Text}

        Dim nogo_min As Object = If(Decimal.TryParse(TxtNoGoMin.Text, tempNoGoMin), tempNoGoMin, DBNull.Value)
        Dim nogo_max As Object = If(Decimal.TryParse(TxtNoGoMax.Text, tempNoGoMax), tempNoGoMax, DBNull.Value)
        Dim noGoDetails = {nogo_min, nogo_max, TxtNoGo1.Text, TxtNoGo2.Text, TxtNoGo3.Text, TxtNoGo4.Text}

        Dim success As Boolean = False
        Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)

        If IsEditMode Then
            success = MySqlCls.UpdateThreadPlugGaugeCalibration(type, controlNo, TargetCycle, size, lc, color, location, temp, humidity, tmu, mu, parameter, locationObs, goDetails, noGoDetails, timeIn, timeOut, totalTime, status, remark, TxtCalibrationDate.SelectedDate)
        Else
            success = MySqlCls.InsertThreadPlugGaugeCalibration(type, controlNo, cycleName, TxtCalibrationDate.SelectedDate, size, lc, color, location, temp, humidity, tmu, mu, parameter, locationObs, goDetails, noGoDetails, timeIn, timeOut, totalTime, status, remark)
        End If

        If success Then
            MySqlCls.UpdateRegularCalibrationStatus(controlNo, calibDate, status, targetCyc)
            MySqlCls.InsertResultRecord(controlNo, _category, _instrumentName, targetCyc)
            CalibrationMasterCache.SaveMasters(controlNo, targetCyc, _selectedMasters)

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
        Try
            Dim dtDept = MySqlCls.ReadDatatable($"SELECT * FROM department_list WHERE `Control No` = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")
            If dtDept.Rows.Count > 0 Then
                Dim row = dtDept.Rows(0)
                Dim dept = row("Department").ToString(), instName = row("InstrumentName").ToString()
                Dim color = row("Color").ToString()
                Dim ngReason = If(Not String.IsNullOrWhiteSpace(TxtRemark.Text), TxtRemark.Text.Trim(), "Calibration NG")
                Dim wopReason = "Calibration NG" & If(Not String.IsNullOrWhiteSpace(TxtRemark.Text), ": " & TxtRemark.Text.Trim(), "")
                Dim dueDate = calibDate.AddYears(1).AddDays(-1)

                MySqlCls.InsertNGRecord(ctrlNo, instName, dept, cycleName, ngReason, calibDate, dueDate, statusVal)
                MySqlCls.InsertWOPRecord(ctrlNo, cycleName, dept, instName, color, wopReason)
                Dim sizeStr = ""
                If row.Table.Columns.Contains("SizeandRange") AndAlso Not IsDBNull(row("SizeandRange")) Then sizeStr = row("SizeandRange").ToString()
                MySqlCls.InsertInterchangeRecord(cycleName, ctrlNo, dept, instName, sizeStr, color, "WOP", wopReason)
                MySqlCls.ClearRFIDTag(ctrlNo)
            End If
        Catch ex As Exception
            Console.WriteLine("ProcessNGFlow Error: " & ex.Message)
        End Try
    End Sub

    Private Sub CalculateTMU()
        Try
            Dim masterU As Decimal = 0
            Dim manualMU As Decimal
            
            If Decimal.TryParse(TxtMU.Text, manualMU) AndAlso manualMU > 0 Then
                masterU = manualMU
            ElseIf _selectedMasters IsNot Nothing AndAlso _selectedMasters.Count > 0 Then
                Dim sumSqMaster As Decimal = 0
                For Each m In _selectedMasters
                    sumSqMaster += (m.MasterUncertainty ^ 2)
                Next
                masterU = Math.Sqrt(sumSqMaster)
            End If

            If masterU <= 0 Then
                TxtTMU.Text = ""
                TxtTMU.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                Return
            End If
            TxtTMU.Background = Brushes.White

            ' 2. Calculate stdev for GO observations
            Dim goVals As New System.Collections.Generic.List(Of Decimal)()
            Dim goTbs = {TxtGo1, TxtGo2, TxtGo3, TxtGo4}
            For Each tb In goTbs
                Dim val As Decimal
                If Decimal.TryParse(tb.Text, val) Then
                    goVals.Add(val)
                End If
            Next

            ' Calculate stdev for NOGO observations
            Dim nogoVals As New System.Collections.Generic.List(Of Decimal)()
            Dim nogoTbs = {TxtNoGo1, TxtNoGo2, TxtNoGo3, TxtNoGo4}
            For Each tb In nogoTbs
                Dim val As Decimal
                If Decimal.TryParse(tb.Text, val) Then
                    nogoVals.Add(val)
                End If
            Next

            ' Only calculate if we have all 4 values for GO and NOGO
            If goVals.Count = 4 AndAlso nogoVals.Count = 4 Then
                Dim stdev_go = calcStdev(goVals.ToArray())
                Dim sigB1_go = stdev_go / Math.Sqrt(4)
                Dim mu_go_sq = (sigB1_go ^ 2) + (masterU ^ 2)

                Dim stdev_nogo = calcStdev(nogoVals.ToArray())
                Dim sigB1_nogo = stdev_nogo / Math.Sqrt(4)
                Dim mu_nogo_sq = (sigB1_nogo ^ 2) + (masterU ^ 2)

                ' Final TMU
                Dim preTMU = Math.Sqrt(mu_go_sq + mu_nogo_sq)
                Dim finalTMU = (preTMU * 2) / 1000.0
                TxtTMU.Text = finalTMU.ToString("F9")
            Else
                TxtTMU.Text = ""
            End If

        Catch ex As Exception
            ' Silently handle calculation errors (e.g. during partial entry)
        End Try
    End Sub

    Private Function calcStdev(arr() As Decimal) As Decimal
        Dim n = arr.Length
        If n <= 1 Then Return 0
        Dim mean = arr.Average()
        Dim sumSqDiff = arr.Sum(Function(v) (v - mean) ^ 2)
        Return Math.Sqrt(sumSqDiff / (n - 1))
    End Function

    Private Sub TextBox_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = TryCast(sender, TextBox)
            If tb IsNot Nothing Then
                Dim navTextBoxes = {TxtGoSize, TxtGoMin, TxtGoMax, TxtGo1, TxtGo2, TxtGo3, TxtGo4,
                                    TxtNoGoSize, TxtNoGoMin, TxtNoGoMax, TxtNoGo1, TxtNoGo2, TxtNoGo3, TxtNoGo4}
                Dim index = Array.IndexOf(navTextBoxes, tb)
                If index >= 0 AndAlso index < navTextBoxes.Length - 1 Then
                    navTextBoxes(index + 1).Focus()
                    e.Handled = True
                ElseIf index = navTextBoxes.Length - 1 Then
                    TxtTimeIn.Focus()
                    e.Handled = True
                End If
            End If
        End If
    End Sub
End Class
