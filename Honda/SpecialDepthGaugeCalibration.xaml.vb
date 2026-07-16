Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq
Imports System
Imports System.Windows.Input
Imports System.IO
Imports System.Text.Json

Public Class DynParamData
    Public Property ParamType As String ' "Dia", "Distance", "Angle"
    Public Property ParamIndex As Integer ' 1, 2, 3...

    Public Property ChkEnabled As CheckBox
    Public Property TxtNominal As TextBox
    Public Property TxtPermErr As TextBox
    Public Property TxtMin As TextBox
    Public Property TxtMax As TextBox
    Public Property TxtObs1 As TextBox
    Public Property TxtObs2 As TextBox
    Public Property TxtObs3 As TextBox
End Class

Public Class SpecialDepthGaugeCalibration
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = "Special Depth Gauge"
    Private _category As String = "Gauge"
    Private _selectedMasters As New System.Collections.Generic.List(Of MasterSelectorItem)()
    Private _lc As String = "0.0001"

    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _isPopulating As Boolean = False

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""
    Private _dynParams As New List(Of DynParamData)()
    Private _diaCount As Integer = 2
    Private _distCount As Integer = 1
    Private _angCount As Integer = 2

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
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM special_depth_gauge_calibration WHERE RowType='MASTER' ORDER BY LC ASC")
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

        ' Attach LostFocus for Control No
        AddHandler TxtControlNo.LostFocus, AddressOf TxtControlNo_LostFocus

        If Not String.IsNullOrWhiteSpace(TxtControlNo.Text) Then
            LoadInstrumentDetails(TxtControlNo.Text.Trim())
        End If

        If Not IsEditMode Then
            TxtCalibrationDate.SelectedDate = DateTime.Today
            If _dynParams.Count = 0 Then
                Dim cached = SpecialDepthConfigCache.LoadConfig()
                If cached IsNot Nothing Then
                    _diaCount = cached.DiaCount
                    _distCount = cached.DistCount
                    _angCount = cached.AngCount
                End If

                _isPopulating = True
                If TxtDiaCount IsNot Nothing Then TxtDiaCount.Text = _diaCount.ToString()
                If TxtDistCount IsNot Nothing Then TxtDistCount.Text = _distCount.ToString()
                If TxtAngCount IsNot Nothing Then TxtAngCount.Text = _angCount.ToString()
                _isPopulating = False

                GenerateObservationRows(_diaCount, _distCount, _angCount)
            End If
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

    Private Sub ComboLC_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboLC.SelectionChanged
        If _isPopulating Then Return
        If ComboLC.SelectedItem IsNot Nothing Then
            _lc = ComboLC.SelectedItem.ToString()
            FetchMasterLimits()
        End If
    End Sub

    Private Sub RevalidateAllObs()
        If _isPopulating Then Return
        For Each dp In _dynParams
            If dp.ChkEnabled.IsChecked.GetValueOrDefault() Then
                ValidateBox(dp.TxtObs1, dp.TxtMin.Text, dp.TxtMax.Text)
                ValidateBox(dp.TxtObs2, dp.TxtMin.Text, dp.TxtMax.Text)
                ValidateBox(dp.TxtObs3, dp.TxtMin.Text, dp.TxtMax.Text)
            End If
        Next
    End Sub

    Private Sub ParamCount_Changed(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        If TxtDiaCount Is Nothing OrElse TxtDistCount Is Nothing OrElse TxtAngCount Is Nothing OrElse ObsStackPanel Is Nothing Then Return

        If Integer.TryParse(TxtDiaCount.Text, _diaCount) AndAlso
           Integer.TryParse(TxtDistCount.Text, _distCount) AndAlso
           Integer.TryParse(TxtAngCount.Text, _angCount) Then
            GenerateObservationRows(_diaCount, _distCount, _angCount)
        End If
    End Sub

    Private Sub GenerateObservationRows(diaCount As Integer, distCount As Integer, angCount As Integer)
        If ObsStackPanel Is Nothing Then Return

        _isPopulating = True
        ObsStackPanel.Children.Clear()
        _dynParams.Clear()

        Dim AddRow = Sub(type As String, count As Integer)
                         For i As Integer = 1 To count
                             Dim dp As New DynParamData() With {.ParamType = type, .ParamIndex = i}
                             Dim prefix = If(type = "Distance", $"Distance_{i}", $"{type}_{i}")
                             Dim labelTxt = $"{type} {i} :"
                             Dim g = CreateDynParamRow(labelTxt, prefix, dp, AddressOf TxtParam_TextChanged, AddressOf ChkParam_CheckedChanged, AddressOf TxtObs_TextChanged, AddressOf TextBox_PreviewKeyDown)
                             _dynParams.Add(dp)
                             ObsStackPanel.Children.Add(g)
                         Next
                     End Sub

        AddRow("Dia", diaCount)
        AddRow("Distance", distCount)
        AddRow("Angle", angCount)

        _isPopulating = False
        CalculateTMU()
        UpdateJudgement()

        ' Save config for state persistence during this day
        SpecialDepthConfigCache.SaveConfig(diaCount, distCount, angCount)
    End Sub

    Private Function CreateDynParamRow(labelTxt As String, propPrefix As String, dp As DynParamData,
                                    paramHandler As TextChangedEventHandler, chkHandler As RoutedEventHandler,
                                    obsHandler As TextChangedEventHandler, keyHandler As KeyEventHandler) As Grid
        Dim g As New Grid()
        g.Margin = New Thickness(0, 4, 0, 4)

        Dim cd(19) As ColumnDefinition
        cd(0) = New ColumnDefinition() With {.Width = New GridLength(0)}
        cd(1) = New ColumnDefinition() With {.Width = New GridLength(70)}
        cd(2) = New ColumnDefinition() With {.Width = New GridLength(55)}
        cd(3) = New ColumnDefinition() With {.Width = New GridLength(5)}
        cd(4) = New ColumnDefinition() With {.Width = GridLength.Auto}
        cd(5) = New ColumnDefinition() With {.Width = New GridLength(65)}
        cd(6) = New ColumnDefinition() With {.Width = New GridLength(5)}
        cd(7) = New ColumnDefinition() With {.Width = GridLength.Auto}
        cd(8) = New ColumnDefinition() With {.Width = New GridLength(65)}
        cd(9) = New ColumnDefinition() With {.Width = New GridLength(5)}
        cd(10) = New ColumnDefinition() With {.Width = GridLength.Auto}
        cd(11) = New ColumnDefinition() With {.Width = New GridLength(65)}
        cd(12) = New ColumnDefinition() With {.Width = New GridLength(15)}
        cd(13) = New ColumnDefinition() With {.Width = GridLength.Auto}
        cd(14) = New ColumnDefinition() With {.Width = New GridLength(65)}
        cd(15) = New ColumnDefinition() With {.Width = New GridLength(10)}
        cd(16) = New ColumnDefinition() With {.Width = GridLength.Auto}
        cd(17) = New ColumnDefinition() With {.Width = New GridLength(65)}
        cd(18) = New ColumnDefinition() With {.Width = New GridLength(10)}
        cd(19) = New ColumnDefinition() With {.Width = GridLength.Auto}
        Dim cd20 = New ColumnDefinition() With {.Width = New GridLength(65)}

        For i = 0 To 19
            g.ColumnDefinitions.Add(cd(i))
        Next
        g.ColumnDefinitions.Add(cd20)

        Dim chk As New CheckBox() With {.IsChecked = True, .VerticalAlignment = VerticalAlignment.Center, .Visibility = Visibility.Collapsed}
        AddHandler chk.Checked, chkHandler
        AddHandler chk.Unchecked, chkHandler
        g.Children.Add(chk)
        Grid.SetColumn(chk, 0)

        Dim lbl As New TextBlock() With {.Text = labelTxt, .FontWeight = FontWeights.SemiBold}
        g.Children.Add(lbl)
        Grid.SetColumn(lbl, 1)

        Dim tNom As New TextBox() With {.Name = "Txt" & propPrefix & "Nominal"}
        AddHandler tNom.TextChanged, paramHandler
        AddHandler tNom.PreviewKeyDown, keyHandler
        g.Children.Add(tNom)
        Grid.SetColumn(tNom, 2)

        Dim lblP As New TextBlock() With {.Text = "± PermErr:", .FontSize = 12}
        g.Children.Add(lblP)
        Grid.SetColumn(lblP, 4)

        Dim tPerm As New TextBox() With {.Name = "Txt" & propPrefix & "PermErr"}
        AddHandler tPerm.TextChanged, paramHandler
        AddHandler tPerm.PreviewKeyDown, keyHandler
        g.Children.Add(tPerm)
        Grid.SetColumn(tPerm, 5)

        Dim lblMin As New TextBlock() With {.Text = "Min:", .FontSize = 12}
        g.Children.Add(lblMin)
        Grid.SetColumn(lblMin, 7)

        Dim tMin As New TextBox() With {.Name = "Txt" & propPrefix & "Min"}
        AddHandler tMin.TextChanged, paramHandler
        AddHandler tMin.PreviewKeyDown, keyHandler
        g.Children.Add(tMin)
        Grid.SetColumn(tMin, 8)

        Dim lblMax As New TextBlock() With {.Text = "Max:", .FontSize = 12}
        g.Children.Add(lblMax)
        Grid.SetColumn(lblMax, 10)

        Dim tMax As New TextBox() With {.Name = "Txt" & propPrefix & "Max"}
        AddHandler tMax.TextChanged, paramHandler
        AddHandler tMax.PreviewKeyDown, keyHandler
        g.Children.Add(tMax)
        Grid.SetColumn(tMax, 11)

        Dim lbl1 As New TextBlock() With {.Text = "1:"}
        g.Children.Add(lbl1)
        Grid.SetColumn(lbl1, 13)

        Dim t1 As New TextBox() With {.Name = "Txt" & propPrefix & "Obs1"}
        AddHandler t1.TextChanged, obsHandler
        AddHandler t1.PreviewKeyDown, keyHandler
        g.Children.Add(t1)
        Grid.SetColumn(t1, 14)

        Dim lbl2 As New TextBlock() With {.Text = "2:"}
        g.Children.Add(lbl2)
        Grid.SetColumn(lbl2, 16)

        Dim t2 As New TextBox() With {.Name = "Txt" & propPrefix & "Obs2"}
        AddHandler t2.TextChanged, obsHandler
        AddHandler t2.PreviewKeyDown, keyHandler
        g.Children.Add(t2)
        Grid.SetColumn(t2, 17)

        Dim lbl3 As New TextBlock() With {.Text = "3:"}
        g.Children.Add(lbl3)
        Grid.SetColumn(lbl3, 19)

        Dim t3 As New TextBox() With {.Name = "Txt" & propPrefix & "Obs3"}
        AddHandler t3.TextChanged, obsHandler
        AddHandler t3.PreviewKeyDown, keyHandler
        g.Children.Add(t3)
        Grid.SetColumn(t3, 20)

        chk.Tag = New Object() {tPerm, tMin, tMax, t1, t2, t3}
        t1.Tag = New Object() {tMin, tMax, propPrefix}
        t2.Tag = New Object() {tMin, tMax, propPrefix}
        t3.Tag = New Object() {tMin, tMax, propPrefix}

        tMin.Tag = propPrefix
        tMax.Tag = propPrefix

        dp.ChkEnabled = chk
        dp.TxtNominal = tNom
        dp.TxtPermErr = tPerm
        dp.TxtMin = tMin
        dp.TxtMax = tMax
        dp.TxtObs1 = t1
        dp.TxtObs2 = t2
        dp.TxtObs3 = t3

        AddHandler tNom.TextChanged, Sub(s, ev)
                                         If Not _isPopulating Then
                                             Dim val = tNom.Text.Trim()
                                             If val = "-" OrElse val = "" Then
                                                 chk.IsChecked = False
                                             Else
                                                 chk.IsChecked = True
                                             End If
                                         End If
                                     End Sub

        Return g
    End Function

    Private Sub ChkParam_CheckedChanged(sender As Object, e As RoutedEventArgs)
        Dim chk = TryCast(sender, CheckBox)
        If chk IsNot Nothing AndAlso chk.Tag IsNot Nothing Then
            Dim arr = DirectCast(chk.Tag, Object())
            Dim isChecked = chk.IsChecked.GetValueOrDefault()
            For Each obj In arr
                Dim tb = TryCast(obj, TextBox)
                If tb IsNot Nothing Then
                    tb.IsEnabled = isChecked
                    If Not isChecked Then
                        tb.Text = "-"
                        tb.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                        tb.Foreground = Brushes.Gray
                    Else
                        If tb.Text = "-" Then tb.Text = ""
                        tb.Background = Brushes.White
                        tb.Foreground = Brushes.Black
                    End If
                End If
            Next
        End If
        If Not _isPopulating Then
            RevalidateAllObs()
            CalculateTMU()
            UpdateJudgement()
        End If
    End Sub

    Private Sub TextBox_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        Dim tb = TryCast(sender, TextBox)
        If tb Is Nothing Then Return

        If e.Key = Key.Right OrElse e.Key = Key.Left OrElse e.Key = Key.Up OrElse e.Key = Key.Down Then
            Dim request = New TraversalRequest(If(e.Key = Key.Right OrElse e.Key = Key.Down, FocusNavigationDirection.Next, FocusNavigationDirection.Previous))
            request.Wrapped = True
            tb.MoveFocus(request)
            e.Handled = True
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
                    Dim dtDept = MySqlCls.ReadDatatable($"SELECT Department, Color FROM department_list WHERE `Control No` = '{controlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")

                    _category = r("Category").ToString()
                    ComboCategory.Text = _category

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

                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), If(row.Table.Columns.Contains("InstrumentName"), row("InstrumentName").ToString(), "Special Depth Gauge"), If(row.Table.Columns.Contains("GaugeName"), row("GaugeName").ToString(), "Special Depth Gauge"))

                    FetchMasterLimits()
                    Exit For
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("LoadInstrumentDetails Error: " & ex.Message)
        Finally
            _isPopulating = False
        End Try
    End Sub

    Private Sub FetchMasterLimits()
        Try
            _masterLimits = MySqlCls.GetSpecialDepthGaugeMasterLimits(_lc)

            ' Refresh validation coloring
            EnvironmentTextChanged(TxtTemp, Nothing)
            EnvironmentTextChanged(TxtHumidity, Nothing)

            RevalidateAllObs()
            CalculateTMU()
        Catch ex As Exception
            Console.WriteLine("FetchMasterLimits Error: " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrWhiteSpace(TxtControlNo.Text) OrElse String.IsNullOrWhiteSpace(TargetCycle) Then Return

        _isPopulating = True
        Try
            Dim masterDt = MySqlCls.GetSpecialDepthGaugeMasterData(TxtControlNo.Text.Trim(), TargetCycle)
            If masterDt IsNot Nothing AndAlso masterDt.Table.Columns.Count > 0 Then
                Dim dr = masterDt
                If Not IsDBNull(dr("Date")) Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(dr("Date"))
                If Not IsDBNull(dr("Type")) Then ComboCategory.Text = dr("Type").ToString()
                If Not IsDBNull(dr("Color")) Then ComboColour.Text = dr("Color").ToString()
                If Not IsDBNull(dr("Location")) Then ComboLocation.Text = dr("Location").ToString()
                If Not IsDBNull(dr("Temperature")) Then TxtTemp.Text = dr("Temperature").ToString()
                If Not IsDBNull(dr("Humidity")) Then TxtHumidity.Text = dr("Humidity").ToString()
                If Not IsDBNull(dr("TMU")) Then TxtTMU.Text = dr("TMU").ToString()
                If dr.Table.Columns.Contains("MU") AndAlso Not IsDBNull(dr("MU")) Then TxtMU.Text = dr("MU").ToString()
                If dr.Table.Columns.Contains("LC") AndAlso Not IsDBNull(dr("LC")) Then
                    _lc = dr("LC").ToString()
                    ComboLC.Text = _lc
                End If
                If Not IsDBNull(dr("TimeIn")) Then TxtTimeIn.Text = dr("TimeIn").ToString()
                If Not IsDBNull(dr("TimeOut")) Then
                    TxtTimeOut.Text = dr("TimeOut").ToString()
                    _isTimeOutManual = True
                End If
                If Not IsDBNull(dr("TotalTime")) Then TxtTotalTime.Text = dr("TotalTime").ToString()
                If Not IsDBNull(dr("Remark")) Then TxtRemarks.Text = dr("Remark").ToString()
                If Not IsDBNull(dr("Status")) Then TxtJudgement.Text = dr("Status").ToString()

                If dr.Table.Columns.Contains("ParameterConfig") AndAlso Not IsDBNull(dr("ParameterConfig")) Then
                    Dim config = dr("ParameterConfig").ToString()
                    Dim parts = config.Split(","c)
                    For Each p In parts
                        Dim kv = p.Split(":"c)
                        If kv.Length = 2 Then
                            Dim k = kv(0).Trim(), v = kv(1).Trim()
                            If k = "Dia" Then _diaCount = Integer.Parse(v)
                            If k = "Dist" Then _distCount = Integer.Parse(v)
                            If k = "Angle" Then _angCount = Integer.Parse(v)
                        End If
                    Next
                Else
                    _diaCount = 2 : _distCount = 1 : _angCount = 2
                End If
                TxtDiaCount.Text = _diaCount.ToString()
                TxtDistCount.Text = _distCount.ToString()
                TxtAngCount.Text = _angCount.ToString()

                ' Fallback if Distance was just 'Distance' instead of 'Distance_1' from legacy
                If _distCount = 1 Then
                    Dim dDt = MySqlCls.GetSpecialDepthGaugeRecordData(TxtControlNo.Text.Trim(), TargetCycle)
                    If dDt.Rows.Count > 0 AndAlso Not dDt.Columns.Contains("Distance_1_Nominal") AndAlso dDt.Columns.Contains("Distance_Nominal") Then
                        For Each c As DataColumn In dDt.Columns
                            c.ColumnName = c.ColumnName.Replace("Distance_", "Distance_1_")
                        Next
                    End If
                End If

                GenerateObservationRows(_diaCount, _distCount, _angCount)
            End If

            Dim recordsDt = MySqlCls.GetSpecialDepthGaugeRecordData(TxtControlNo.Text.Trim(), TargetCycle)
            If recordsDt.Rows.Count > 0 Then
                Dim r = recordsDt.Rows(0)
                If r.Table.Columns.Contains("Distance_Nominal") AndAlso Not r.Table.Columns.Contains("Distance_1_Nominal") Then
                    For Each c As DataColumn In r.Table.Columns
                        c.ColumnName = c.ColumnName.Replace("Distance_", "Distance_1_")
                    Next
                End If

                For Each dp In _dynParams
                    Dim prefix = If(dp.ParamType = "Distance", $"Distance_{dp.ParamIndex}", $"{dp.ParamType}_{dp.ParamIndex}")
                    PopulateParamRow(dp.ChkEnabled, dp.TxtNominal, dp.TxtPermErr, dp.TxtMin, dp.TxtMax, dp.TxtObs1, dp.TxtObs2, dp.TxtObs3, r, prefix)
                Next
            End If

            ' Load selected masters
            Dim cached = CalibrationMasterCache.LoadMasters(TxtControlNo.Text.Trim(), TargetCycle)
            If cached.Count > 0 Then
                _selectedMasters = MySqlCls.GetMastersByDescriptions(cached)
                UpdateMasterUI()
            End If

        Catch ex As Exception
            Console.WriteLine("PopulateCalibrationData Error: " & ex.Message)
        Finally
            _isPopulating = False
        End Try

        Try
            EnvironmentTextChanged(TxtTemp, Nothing)
            EnvironmentTextChanged(TxtHumidity, Nothing)
            RevalidateAllObs()
            CalculateTMU()
        Catch ex As Exception
            Console.WriteLine("PopulateCalibrationData Revalidation Error: " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateParamRow(chk As CheckBox, tNom As TextBox, tPerm As TextBox, tMin As TextBox, tMax As TextBox, t1 As TextBox, t2 As TextBox, t3 As TextBox, dr As DataRow, prefix As String)
        If dr.Table.Columns.Contains(prefix & "_Nominal") AndAlso Not IsDBNull(dr(prefix & "_Nominal")) Then
            Dim val = dr(prefix & "_Nominal").ToString()
            tNom.Text = val
            If val = "-" Then
                chk.IsChecked = False
                ChkParam_CheckedChanged(chk, Nothing)
            End If
        End If
        If chk.IsChecked Then
            If dr.Table.Columns.Contains(prefix & "_Permissible_Error") AndAlso Not IsDBNull(dr(prefix & "_Permissible_Error")) Then tPerm.Text = dr(prefix & "_Permissible_Error").ToString()
            If dr.Table.Columns.Contains(prefix & "_Min_Limit") AndAlso Not IsDBNull(dr(prefix & "_Min_Limit")) Then tMin.Text = dr(prefix & "_Min_Limit").ToString()
            If dr.Table.Columns.Contains(prefix & "_Max_Limit") AndAlso Not IsDBNull(dr(prefix & "_Max_Limit")) Then tMax.Text = dr(prefix & "_Max_Limit").ToString()
            If dr.Table.Columns.Contains(prefix & "_Obs_1") AndAlso Not IsDBNull(dr(prefix & "_Obs_1")) Then t1.Text = dr(prefix & "_Obs_1").ToString()
            If dr.Table.Columns.Contains(prefix & "_Obs_2") AndAlso Not IsDBNull(dr(prefix & "_Obs_2")) Then t2.Text = dr(prefix & "_Obs_2").ToString()
            If dr.Table.Columns.Contains(prefix & "_Obs_3") AndAlso Not IsDBNull(dr(prefix & "_Obs_3")) Then t3.Text = dr(prefix & "_Obs_3").ToString()
        End If
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

    Private Sub Time_TextChanged(sender As Object, e As TextChangedEventArgs)
        UpdateDuration(sender, e)
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
                isOK = (val >= ll AndAlso val <= ul)
            ElseIf tb.Name = "TxtHumidity" Then
                Dim ll = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_LL") AndAlso Not IsDBNull(_masterLimits("Env_Hum_LL")), Convert.ToDecimal(_masterLimits("Env_Hum_LL")), 40.0)
                Dim ul = If(_masterLimits IsNot Nothing AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_UL") AndAlso Not IsDBNull(_masterLimits("Env_Hum_UL")), Convert.ToDecimal(_masterLimits("Env_Hum_UL")), 60.0)
                isOK = (val >= ll AndAlso val <= ul)
            End If

            If isOK Then
                tb.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' Light Green
            Else
                tb.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' Light Pink
            End If
        Else
            tb.Background = Brushes.White
        End If
        UpdateJudgement()
    End Sub

    Private Sub TxtParam_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim tb = TryCast(sender, TextBox)
        If tb IsNot Nothing AndAlso tb.Tag IsNot Nothing Then
            Dim prefix = tb.Tag.ToString()
            RevalidateAllObs()
        End If
        UpdateJudgement()
    End Sub

    Private Sub TxtObs_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating Then Return
        Dim tb = TryCast(sender, TextBox)
        If tb IsNot Nothing AndAlso tb.Tag IsNot Nothing Then
            Dim arr = DirectCast(tb.Tag, Object())
            Dim tMin = DirectCast(arr(0), TextBox)
            Dim tMax = DirectCast(arr(1), TextBox)
            ValidateBox(tb, tMin.Text, tMax.Text)
        End If
    End Sub

    Private Sub ValidateBox(tb As TextBox, minStr As String, maxStr As String)
        If _isPopulating OrElse tb Is Nothing OrElse Not tb.IsEnabled Then Return
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            Dim ll, ul As Decimal
            Dim hasMin = Decimal.TryParse(minStr, ll)
            Dim hasMax = Decimal.TryParse(maxStr, ul)

            Dim isOK = True
            If hasMin AndAlso val < ll Then isOK = False
            If hasMax AndAlso val > ul Then isOK = False

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
        CalculateTMU()
        UpdateJudgement()
    End Sub

    Private Sub CalculateTMU()
        If _isPopulating Then Return

        ' Step 1: Determine masterU (half-width σ for RSS)
        ' Manual MU is entered as half-width σ directly (same as $Q$4 in Excel) — use as-is
        ' Selected masters store expanded U = 2σ, so use MU/2 as half-width
        Dim masterU As Double = 0
        Dim manualMU As Double
        If Not String.IsNullOrWhiteSpace(TxtMU.Text) AndAlso Double.TryParse(TxtMU.Text, manualMU) AndAlso manualMU > 0 Then
            masterU = manualMU
        ElseIf _selectedMasters.Count > 0 Then
            Dim sumSqMaster As Double = 0
            For Each m In _selectedMasters
                sumSqMaster += (m.MasterUncertainty / 2) ^ 2
            Next
            masterU = Math.Sqrt(sumSqMaster)
        End If

        If masterU <= 0 Then
            TxtTMU.Text = ""
            Return
        End If

        ' Step 2: Per-parameter per-tab RSS
        ' Only Dia1, Dia2, Distance contribute — Angles excluded
        ' Each (parameter × tab) = one independent "size" in the reference formula
        Dim muSumSq As Double = 0
        Dim calculatedParamsCount As Integer = 0

        For Each dp In _dynParams
            If dp.ParamType = "Angle" Then Continue For
            If Not dp.ChkEnabled.IsChecked.GetValueOrDefault() Then Continue For

            Dim vals As New System.Collections.Generic.List(Of Double)()
            Dim tbs = {dp.TxtObs1, dp.TxtObs2, dp.TxtObs3}
            For Each tb In tbs
                If tb IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) AndAlso tb.Text <> "-" Then
                    Dim d As Double
                    If Double.TryParse(tb.Text, d) Then vals.Add(d)
                End If
            Next

            If vals.Count = 3 Then
                Dim sigmaR1 = CalcStdevDepth(vals.ToArray())
                Dim sigmaB1 = sigmaR1 / Math.Sqrt(3)
                Dim tmuParam_sq = (sigmaB1 ^ 2) + (masterU ^ 2)
                muSumSq += tmuParam_sq
                calculatedParamsCount += 1
            End If
        Next

        ' Step 3: Final TMU = 2 × SQRT(SUM(M.U_param²))
        ' Matches reference image: Pre_TMU = M.U × 2, TMU = SQRT(SUM(Pre_TMU²))
        If calculatedParamsCount > 0 Then
            Dim finalTMU = Math.Sqrt(muSumSq) * 2
            TxtTMU.Text = finalTMU.ToString("F9")
        Else
            TxtTMU.Text = ""
        End If
    End Sub

    Private Sub UpdateJudgement()
        If _isPopulating Then Return

        Dim envBoxes = {TxtTemp, TxtHumidity, TxtTotalTime}
        Dim anyEnvNG = False
        For Each box In envBoxes
            If box IsNot Nothing AndAlso TypeOf box.Background Is SolidColorBrush Then
                Dim color = DirectCast(box.Background, SolidColorBrush).Color
                If color = Color.FromRgb(254, 226, 226) Then
                    anyEnvNG = True
                    Exit For
                End If
            End If
        Next

        Dim anyObsNG = False
        Dim hasAnyObs = False

        For Each dp In _dynParams
            If dp.ChkEnabled.IsChecked.GetValueOrDefault() Then
                Dim tbs = {dp.TxtObs1, dp.TxtObs2, dp.TxtObs3}
                For Each tb In tbs
                    If tb IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) AndAlso tb.Text <> "-" Then
                        hasAnyObs = True
                        If TypeOf tb.Background Is SolidColorBrush Then
                            Dim color = DirectCast(tb.Background, SolidColorBrush).Color
                            If color = Color.FromRgb(254, 226, 226) Then
                                anyObsNG = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If anyObsNG Then Exit For
            End If
        Next

        If anyEnvNG OrElse anyObsNG Then
            TxtJudgement.Text = "NG"
            TxtJudgement.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226))
        ElseIf hasAnyObs Then
            TxtJudgement.Text = "OK"
            TxtJudgement.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231))
        Else
            TxtJudgement.Text = ""
            TxtJudgement.Background = Brushes.White
        End If
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As RoutedEventArgs)
        Dim cycleName = If(IsEditMode, TargetCycle, MySqlCls.GetActiveCycleName())
        Dim controlNo = TxtControlNo.Text.Trim()
        If String.IsNullOrWhiteSpace(controlNo) Then
            MessageBox.Show("Control No cannot be empty.")
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

        Dim tmu = TxtTMU.Text.Replace("±", "").Trim()
        Dim status = TxtJudgement.Text

        Dim calibDate As DateTime? = Nothing
        If TxtCalibrationDate.SelectedDate.HasValue Then
            calibDate = TxtCalibrationDate.SelectedDate.Value
        End If

        Dim llT As Object = DBNull.Value
        Dim ulT As Object = DBNull.Value
        Dim llH As Object = DBNull.Value
        Dim ulH As Object = DBNull.Value

        If _masterLimits IsNot Nothing Then
            If _masterLimits.Table.Columns.Contains("Env_Temp_LL") Then llT = _masterLimits("Env_Temp_LL")
            If _masterLimits.Table.Columns.Contains("Env_Temp_UL") Then ulT = _masterLimits("Env_Temp_UL")
            If _masterLimits.Table.Columns.Contains("Env_Hum_LL") Then llH = _masterLimits("Env_Hum_LL")
            If _masterLimits.Table.Columns.Contains("Env_Hum_UL") Then ulH = _masterLimits("Env_Hum_UL")
        End If

        Dim paramConfig = $"Dia:{_diaCount},Dist:{_distCount},Angle:{_angCount}"
        MySqlCls.EnsureParameterConfigColumn()
        MySqlCls.EnsureSpecialDepthColumns(_diaCount, _distCount, _angCount)

        MySqlCls.DeleteSpecialDepthGaugeCalibration(controlNo, cycleName)

        Dim success = MySqlCls.InsertSpecialDepthGaugeMaster(controlNo, cycleName, calibDate, _category, "", _lc, ComboColour.Text, ComboLocation.Text, TxtTemp.Text, TxtHumidity.Text, tmu, TxtMU.Text, llT, ulT, llH, ulH, status, TxtRemarks.Text, TxtTimeIn.Text, TxtTimeOut.Text, TxtTotalTime.Text, paramConfig)

        If success Then
            Dim paramList As New System.Collections.Generic.List(Of SpecialDepthDynParam)()
            For Each dp In _dynParams
                If Not dp.ChkEnabled.IsChecked.GetValueOrDefault() Then Continue For
                Dim pfx = If(dp.ParamType = "Distance", $"Distance_{dp.ParamIndex}", $"{dp.ParamType}_{dp.ParamIndex}")
                paramList.Add(New SpecialDepthDynParam With {
                    .ParamPrefix = pfx,
                    .Nominal = dp.TxtNominal.Text,
                    .PermErr = dp.TxtPermErr.Text,
                    .MinLimit = dp.TxtMin.Text,
                    .MaxLimit = dp.TxtMax.Text,
                    .Obs1 = dp.TxtObs1.Text,
                    .Obs2 = dp.TxtObs2.Text,
                    .Obs3 = dp.TxtObs3.Text
                })
            Next
            MySqlCls.InsertSpecialDepthGaugeDynamicRecord(controlNo, cycleName, paramList)

            Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)
            MySqlCls.UpdateRegularCalibrationStatus(controlNo, calibDate.Value, status, targetCyc)
            MySqlCls.InsertResultRecord(controlNo, _category, _instrumentName, targetCyc)

            ' Save selected masters to cache
            CalibrationMasterCache.SaveMasters(controlNo, targetCyc, _selectedMasters)

            If status = "NG" Then
                ProcessNGFlow(controlNo, targetCyc, calibDate.Value, status)
            End If

            MessageBox.Show("Calibration record saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.Close()
        Else
            MessageBox.Show("Failed to save calibration record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub ProcessNGFlow(ctrlNo As String, cycleName As String, calibDate As DateTime, statusVal As String)
        Try
            Dim dtDept = MySqlCls.ReadDatatable($"SELECT * FROM department_list WHERE `Control No` = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")
            If dtDept.Rows.Count > 0 Then
                Dim row = dtDept.Rows(0)
                Dim dept = row("Department").ToString(), instName = row("InstrumentName").ToString()
                Dim color = row("Color").ToString()
                Dim ngReason = If(Not String.IsNullOrWhiteSpace(TxtRemarks.Text), TxtRemarks.Text.Trim(), "Calibration NG")
                Dim wopReason = "Calibration NG" & If(Not String.IsNullOrWhiteSpace(TxtRemarks.Text), ": " & TxtRemarks.Text.Trim(), "")
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

    Private Function CalcStdevDepth(arr() As Double) As Double
        Dim n = arr.Length
        If n <= 1 Then Return 0
        Dim mean = arr.Average()
        Dim sumSqDiff = arr.Sum(Function(v) (v - mean) ^ 2)
        Return Math.Sqrt(sumSqDiff / (n - 1))
    End Function

End Class

Public Class SpecialDepthConfigCache
    Private Shared ReadOnly CacheFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "special_depth_config_cache.json")

    Public Class ConfigEntry
        Public Property SavedDate As DateTime
        Public Property DiaCount As Integer
        Public Property DistCount As Integer
        Public Property AngCount As Integer
    End Class

    Public Shared Function LoadConfig() As ConfigEntry
        Try
            If File.Exists(CacheFilePath) Then
                Dim json = File.ReadAllText(CacheFilePath)
                If Not String.IsNullOrWhiteSpace(json) Then
                    Dim config = JsonSerializer.Deserialize(Of ConfigEntry)(json)
                    If config IsNot Nothing AndAlso config.SavedDate.Date = DateTime.Today Then
                        Return config
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
        Return Nothing
    End Function

    Public Shared Sub SaveConfig(dia As Integer, dist As Integer, ang As Integer)
        Try
            Dim config As New ConfigEntry With {
                .SavedDate = DateTime.Today,
                .DiaCount = dia,
                .DistCount = dist,
                .AngCount = ang
            }
            Dim json = JsonSerializer.Serialize(config)
            File.WriteAllText(CacheFilePath, json)
        Catch ex As Exception
        End Try
    End Sub
End Class
