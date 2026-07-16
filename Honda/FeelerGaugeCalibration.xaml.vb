Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq
Imports System
Imports System.Collections.Generic
Imports System.Windows.Input


Public Class FeelerGaugeCalibration
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = "Feeler Gauge"
    Private _category As String = "Gauge"
    Private _selectedMasters As New List(Of MasterSelectorItem)()
    Private _lc As String = "0.01"

    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _isPopulating As Boolean = False

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    Private _observationRows As New List(Of FeelerObservationRow)()
    Private _navTextBoxes As New List(Of TextBox)()


    Public Class FeelerObservationRow
        Public SizeBox As TextBox
        Public MinBox As TextBox
        Public MaxBox As TextBox
        Public ThickBoxes As TextBox()
        Public CamberBoxes As TextBox()
        Public RowGrid As Grid
    End Class

    Public Sub New()
        InitializeComponent()

        ' Wire up soaking time duration calculation
        AddHandler TxtTimeIn.TextChanged, AddressOf UpdateDuration
        AddHandler TxtTimeOut.TextChanged, AddressOf UpdateDuration
        AddHandler TxtTimeOut.GotFocus, Sub() _isTimeOutManual = True
        AddHandler TxtTimeOut.PreviewKeyDown, Sub() _isTimeOutManual = True
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ' Populate Locations & Colors from database suggestions
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
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM feeler_gauge_calibration WHERE RowType='MASTER' ORDER BY LC ASC")
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
        End If

        _timerRealTime = New System.Windows.Threading.DispatcherTimer()
        _timerRealTime.Interval = TimeSpan.FromSeconds(5)
        AddHandler _timerRealTime.Tick, AddressOf TimerRealTime_Tick
        _timerRealTime.Start()

        AddHandler TxtSize.PreviewKeyDown, AddressOf TextBox_PreviewKeyDown

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
                    TxtSize.Text = matchedSize

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

                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), If(row.Table.Columns.Contains("InstrumentName"), row("InstrumentName").ToString(), "Feeler Gauge"), If(row.Table.Columns.Contains("GaugeName"), row("GaugeName").ToString(), "Feeler Gauge"))

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
            _masterLimits = MySqlCls.GetFeelerGaugeMasterLimits(_lc)
            ' Refresh validation coloring
            EnvironmentTextChanged(TxtTemp, Nothing)
            EnvironmentTextChanged(TxtHumidity, Nothing)
            For Each row In _observationRows
                ValidateRowBoxes(row)
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
            Dim masterRow = MySqlCls.GetFeelerGaugeMasterData(TxtControlNo.Text.Trim(), TargetCycle)
            If masterRow IsNot Nothing Then
                If Not IsDBNull(masterRow("Date")) Then
                    TxtCalibrationDate.SelectedDate = Convert.ToDateTime(masterRow("Date"))
                End If

                If Not IsDBNull(masterRow("Type")) Then ComboCategory.Text = masterRow("Type").ToString()
                If Not IsDBNull(masterRow("Color")) Then ComboColour.Text = masterRow("Color").ToString()
                If Not IsDBNull(masterRow("Location")) Then ComboLocation.Text = masterRow("Location").ToString()
                If Not IsDBNull(masterRow("Temperature")) Then TxtTemp.Text = masterRow("Temperature").ToString()
                If Not IsDBNull(masterRow("Humidity")) Then TxtHumidity.Text = masterRow("Humidity").ToString()
                If Not IsDBNull(masterRow("TMU")) Then TxtTMU.Text = masterRow("TMU").ToString()
                If masterRow.Table.Columns.Contains("MU") AndAlso Not IsDBNull(masterRow("MU")) Then
                    TxtMU.Text = masterRow("MU").ToString()
                End If

                If masterRow.Table.Columns.Contains("LC") AndAlso Not IsDBNull(masterRow("LC")) Then
                    _lc = masterRow("LC").ToString()
                    ComboLC.Text = _lc
                End If

                If masterRow.Table.Columns.Contains("Size") AndAlso Not IsDBNull(masterRow("Size")) Then
                    TxtSize.Text = masterRow("Size").ToString()
                End If

                If masterRow.Table.Columns.Contains("TimeIn") AndAlso Not IsDBNull(masterRow("TimeIn")) Then TxtTimeIn.Text = masterRow("TimeIn").ToString()
                If masterRow.Table.Columns.Contains("TimeOut") AndAlso Not IsDBNull(masterRow("TimeOut")) Then
                    TxtTimeOut.Text = masterRow("TimeOut").ToString()
                    _isTimeOutManual = True
                End If
                If masterRow.Table.Columns.Contains("TotalTime") AndAlso Not IsDBNull(masterRow("TotalTime")) Then TxtTotalTime.Text = masterRow("TotalTime").ToString()
                If Not IsDBNull(masterRow("Remark")) Then TxtRemark.Text = masterRow("Remark").ToString()
                If masterRow.Table.Columns.Contains("Status") AndAlso Not IsDBNull(masterRow("Status")) Then
                    ComboJudgement.Text = masterRow("Status").ToString()
                ElseIf masterRow.Table.Columns.Contains("Judgement") AndAlso Not IsDBNull(masterRow("Judgement")) Then
                    ComboJudgement.Text = masterRow("Judgement").ToString()
                End If

                Dim dtRecords = MySqlCls.GetFeelerGaugeRecordData(TxtControlNo.Text.Trim(), TargetCycle)
                GenerateObservationRows(dtRecords.Rows.Count)

                For idx As Integer = 0 To dtRecords.Rows.Count - 1
                    Dim recordRow = dtRecords.Rows(idx)
                    Dim rowObj = _observationRows(idx)

                    If Not IsDBNull(recordRow("Measurement_Range")) Then rowObj.SizeBox.Text = recordRow("Measurement_Range").ToString()
                    If Not IsDBNull(recordRow("Min_Limit")) Then rowObj.MinBox.Text = recordRow("Min_Limit").ToString()
                    If Not IsDBNull(recordRow("Max_Limit")) Then rowObj.MaxBox.Text = recordRow("Max_Limit").ToString()

                    Dim obs1Col = If(recordRow.Table.Columns.Contains("Obs_1"), "Obs_1", "Thickness_O1")
                    Dim obs2Col = If(recordRow.Table.Columns.Contains("Obs_2"), "Obs_2", "Thickness_O2")
                    Dim obs3Col = If(recordRow.Table.Columns.Contains("Obs_3"), "Obs_3", "Thickness_O3")
                    Dim camber1Col = If(recordRow.Table.Columns.Contains("Camber_Width_Obs_1"), "Camber_Width_Obs_1", "Camber_O1")
                    Dim camber2Col = If(recordRow.Table.Columns.Contains("Camber_Width_Obs_2"), "Camber_Width_Obs_2", "Camber_O2")
                    Dim camber3Col = If(recordRow.Table.Columns.Contains("Camber_Width_Obs_3"), "Camber_Width_Obs_3", "Camber_O3")

                    If Not IsDBNull(recordRow(obs1Col)) Then rowObj.ThickBoxes(0).Text = recordRow(obs1Col).ToString()
                    If Not IsDBNull(recordRow(obs2Col)) Then rowObj.ThickBoxes(1).Text = recordRow(obs2Col).ToString()
                    If Not IsDBNull(recordRow(obs3Col)) Then rowObj.ThickBoxes(2).Text = recordRow(obs3Col).ToString()

                    If Not IsDBNull(recordRow(camber1Col)) Then rowObj.CamberBoxes(0).Text = recordRow(camber1Col).ToString()
                    If Not IsDBNull(recordRow(camber2Col)) Then rowObj.CamberBoxes(1).Text = recordRow(camber2Col).ToString()
                    If Not IsDBNull(recordRow(camber3Col)) Then rowObj.CamberBoxes(2).Text = recordRow(camber3Col).ToString()
                Next

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

        Try
            EnvironmentTextChanged(TxtTemp, Nothing)
            EnvironmentTextChanged(TxtHumidity, Nothing)
            For Each row In _observationRows
                UpdateCamberRow(row)
                ValidateRowBoxes(row)
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

    Private Sub ValidateBox(tb As TextBox, minStr As String, maxStr As String)
        If _isPopulating OrElse tb Is Nothing Then Return
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
        UpdateJudgement()
        CalculateTMU()
    End Sub

    Private Sub UpdateCamberRow(row As FeelerObservationRow)
        If _isPopulating Then Return

        Dim sizeVal As Decimal
        Dim isCamberActive As Boolean = False
        If Decimal.TryParse(row.SizeBox.Text, sizeVal) AndAlso sizeVal > 0.35D Then
            isCamberActive = True
        End If

        If isCamberActive Then
            ' Calculate camber 1 from thick boxes
            Dim vals As New List(Of Decimal)()
            Dim decVal As Decimal
            For Each tb In row.ThickBoxes
                If Decimal.TryParse(tb.Text, decVal) Then
                    vals.Add(decVal)
                End If
            Next

            If vals.Count > 0 Then
                row.CamberBoxes(0).Text = (vals.Max() - vals.Min()).ToString("F3")
            Else
                row.CamberBoxes(0).Text = ""
            End If
            row.CamberBoxes(1).Text = "-"
            row.CamberBoxes(2).Text = "-"
        Else
            row.CamberBoxes(0).Text = "-"
            row.CamberBoxes(1).Text = "-"
            row.CamberBoxes(2).Text = "-"
        End If

        For Each cb In row.CamberBoxes
            cb.IsReadOnly = True
        Next
    End Sub

    Private Sub ValidateRowBoxes(row As FeelerObservationRow)
        If _isPopulating Then Return
        For Each tb In row.ThickBoxes
            ValidateBox(tb, row.MinBox.Text, row.MaxBox.Text)
        Next
        For Each tb In row.CamberBoxes
            ValidateBox(tb, row.MinBox.Text, row.MaxBox.Text)
        Next
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

        ' Check Observations
        For Each row In _observationRows
            Dim minVal, maxVal As Decimal
            Dim hasMin = Decimal.TryParse(row.MinBox.Text, minVal)
            Dim hasMax = Decimal.TryParse(row.MaxBox.Text, maxVal)

            For Each tb In row.ThickBoxes
                Dim val As Decimal
                If Decimal.TryParse(tb.Text, val) Then
                    If hasMin AndAlso val < minVal Then isAnyNG = True
                    If hasMax AndAlso val > maxVal Then isAnyNG = True
                ElseIf Not String.IsNullOrWhiteSpace(tb.Text) AndAlso tb.Text.Trim() <> "-" Then
                    isAnyNG = True
                End If
            Next

            For Each tb In row.CamberBoxes
                Dim val As Decimal
                If Decimal.TryParse(tb.Text, val) Then
                    If hasMin AndAlso val < minVal Then isAnyNG = True
                    If hasMax AndAlso val > maxVal Then isAnyNG = True
                ElseIf Not String.IsNullOrWhiteSpace(tb.Text) AndAlso tb.Text.Trim() <> "-" Then
                    isAnyNG = True
                End If
            Next
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

    Private Sub BtnConfigureSizes_Click(sender As Object, e As RoutedEventArgs)
        Dim prompt As New Window() With {
            .Title = "Configure Size Measurements",
            .Width = 350,
            .Height = 180,
            .WindowStartupLocation = WindowStartupLocation.CenterOwner,
            .Owner = Window.GetWindow(Me),
            .ResizeMode = ResizeMode.NoResize,
            .Background = New SolidColorBrush(Color.FromRgb(248, 250, 252)),
            .FontFamily = New FontFamily("Segoe UI")
        }

        Dim grid As New Grid() With {.Margin = New Thickness(20)}
        grid.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})
        grid.RowDefinitions.Add(New RowDefinition() With {.Height = New GridLength(1, GridUnitType.Star)})
        grid.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})

        Dim lblPrompt As New TextBlock() With {
            .Text = "Enter number of sizes to measure:",
            .Margin = New Thickness(0, 0, 0, 10),
            .FontWeight = FontWeights.SemiBold
        }
        Grid.SetRow(lblPrompt, 0)
        grid.Children.Add(lblPrompt)

        Dim txtInput As New TextBox() With {
            .Height = 32,
            .VerticalContentAlignment = VerticalAlignment.Center,
            .Padding = New Thickness(5)
        }
        Grid.SetRow(txtInput, 1)
        grid.Children.Add(txtInput)

        Dim btnStack As New StackPanel() With {
            .Orientation = Orientation.Horizontal,
            .HorizontalAlignment = HorizontalAlignment.Right,
            .Margin = New Thickness(0, 10, 0, 0)
        }
        Grid.SetRow(btnStack, 2)

        Dim btnCancel As New Button() With {
            .Content = "Cancel",
            .Width = 80,
            .Height = 30,
            .Margin = New Thickness(0, 0, 10, 0),
            .Background = Brushes.Transparent,
            .Foreground = New SolidColorBrush(Color.FromRgb(71, 85, 105))
        }
        AddHandler btnCancel.Click, Sub() prompt.Close()
        btnStack.Children.Add(btnCancel)

        Dim btnOk As New Button() With {
            .Content = "OK",
            .Width = 80,
            .Height = 30,
            .Background = New SolidColorBrush(Color.FromRgb(59, 130, 246)),
            .Foreground = Brushes.White
        }
        AddHandler btnOk.Click, Sub()
                                    Dim num As Integer
                                    If Integer.TryParse(txtInput.Text.Trim(), num) AndAlso num > 0 AndAlso num <= 100 Then
                                        GenerateObservationRows(num)
                                        prompt.DialogResult = True
                                        prompt.Close()
                                    Else
                                        MessageBox.Show("Please enter a valid positive integer (max 100).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning)
                                    End If
                                End Sub
        btnStack.Children.Add(btnOk)
        grid.Children.Add(btnStack)
        prompt.Content = grid

        prompt.ShowDialog()
    End Sub

    Private Sub GenerateObservationRows(numRows As Integer)
        _isPopulating = True
        ObservationRowsContainer.Children.Clear()
        _observationRows.Clear()
        _navTextBoxes.Clear()
        _navTextBoxes.Add(TxtSize)

        For i As Integer = 1 To numRows
            Dim rowGrid As New Grid()
            rowGrid.Margin = New Thickness(0, 4, 0, 4)
            rowGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(1.2, GridUnitType.Star)})
            For col As Integer = 1 To 8
                rowGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(1, GridUnitType.Star)})
            Next

            ' Create Size Box
            Dim txtSize As New TextBox() With {.HorizontalAlignment = HorizontalAlignment.Stretch, .Margin = New Thickness(2)}
            Grid.SetColumn(txtSize, 0)
            rowGrid.Children.Add(txtSize)

            ' Create Min Box
            Dim txtMin As New TextBox() With {.HorizontalAlignment = HorizontalAlignment.Stretch, .Margin = New Thickness(2)}
            Grid.SetColumn(txtMin, 1)
            rowGrid.Children.Add(txtMin)

            ' Create Max Box
            Dim txtMax As New TextBox() With {.HorizontalAlignment = HorizontalAlignment.Stretch, .Margin = New Thickness(2)}
            Grid.SetColumn(txtMax, 2)
            rowGrid.Children.Add(txtMax)

            ' Create Thickness Obs 1, 2, 3
            Dim thickBoxes(2) As TextBox
            For j As Integer = 0 To 2
                thickBoxes(j) = New TextBox() With {.HorizontalAlignment = HorizontalAlignment.Stretch, .Margin = New Thickness(2)}
                Grid.SetColumn(thickBoxes(j), 3 + j)
                rowGrid.Children.Add(thickBoxes(j))
            Next

            ' Create Camber Obs 1, 2, 3
            Dim camberBoxes(2) As TextBox
            For j As Integer = 0 To 2
                camberBoxes(j) = New TextBox() With {.HorizontalAlignment = HorizontalAlignment.Stretch, .Margin = New Thickness(2)}
                Grid.SetColumn(camberBoxes(j), 6 + j)
                rowGrid.Children.Add(camberBoxes(j))
            Next

            Dim rowObj As New FeelerObservationRow() With {
                .SizeBox = txtSize,
                .MinBox = txtMin,
                .MaxBox = txtMax,
                .ThickBoxes = thickBoxes,
                .CamberBoxes = camberBoxes,
                .RowGrid = rowGrid
            }

            ' Wire up event handlers for validation
            AddHandler txtMin.TextChanged, Sub() ValidateRowBoxes(rowObj)
            AddHandler txtMax.TextChanged, Sub() ValidateRowBoxes(rowObj)

            AddHandler txtSize.TextChanged, Sub() UpdateCamberRow(rowObj)

            For j As Integer = 0 To 2
                AddHandler thickBoxes(j).TextChanged, Sub(sender, args)
                                                          ValidateBox(TryCast(sender, TextBox), txtMin.Text, txtMax.Text)
                                                      End Sub
                AddHandler thickBoxes(j).TextChanged, Sub() UpdateCamberRow(rowObj)
                AddHandler camberBoxes(j).TextChanged, Sub(sender, args)
                                                           ValidateBox(TryCast(sender, TextBox), txtMin.Text, txtMax.Text)
                                                       End Sub
            Next

            ' Register for Enter-key navigation
            _navTextBoxes.Add(txtSize)
            _navTextBoxes.Add(txtMin)
            _navTextBoxes.Add(txtMax)
            For Each tb In thickBoxes
                _navTextBoxes.Add(tb)
            Next

            AddHandler txtSize.PreviewKeyDown, AddressOf TextBox_PreviewKeyDown
            AddHandler txtMin.PreviewKeyDown, AddressOf TextBox_PreviewKeyDown
            AddHandler txtMax.PreviewKeyDown, AddressOf TextBox_PreviewKeyDown
            For Each tb In thickBoxes
                AddHandler tb.PreviewKeyDown, AddressOf TextBox_PreviewKeyDown
            Next

            ObservationRowsContainer.Children.Add(rowGrid)
            _observationRows.Add(rowObj)
        Next
        _isPopulating = False
        UpdateJudgement()
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(TxtControlNo.Text) Then
            MessageBox.Show("Control No cannot be empty.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        If String.IsNullOrWhiteSpace(TxtSize.Text) Then
            MessageBox.Show("Measuring Range cannot be empty.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        If _observationRows.Count = 0 Then
            MessageBox.Show("Please configure size measurements first.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning)
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
        Dim size = TxtSize.Text.Trim()
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

        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, calibDate.ToShortDateString(), TxtTemp.Text, TxtHumidity.Text, TxtTimeIn.Text, "")

        ' Check environmental limits to save in MASTER row
        Dim tempMin, tempMax, humMin, humMax As Decimal?
        If _masterLimits IsNot Nothing Then
            If Not IsDBNull(_masterLimits("Env_Temp_LL")) Then tempMin = Convert.ToDecimal(_masterLimits("Env_Temp_LL"))
            If Not IsDBNull(_masterLimits("Env_Temp_UL")) Then tempMax = Convert.ToDecimal(_masterLimits("Env_Temp_UL"))
            If Not IsDBNull(_masterLimits("Env_Hum_LL")) Then humMin = Convert.ToDecimal(_masterLimits("Env_Hum_LL"))
            If Not IsDBNull(_masterLimits("Env_Hum_UL")) Then humMax = Convert.ToDecimal(_masterLimits("Env_Hum_UL"))
        End If

        Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)

        ' Delete old MASTER and RECORD calibration entries for this gauge
        MySqlCls.DeleteFeelerGaugeCalibration(controlNo, targetCyc)

        ' Insert MASTER row
        Dim success = MySqlCls.InsertFeelerGaugeMaster(controlNo, targetCyc, calibDate, type, size, lc, color, location, temp, humidity, tmu, mu, tempMin, tempMax, humMin, humMax, status, remark, ProjectSettings.Current.Username, timeIn, timeOut, totalTime)

        If success Then
            ' Insert each observation row as a RECORD
            For Each row In _observationRows
                Dim measRange = row.SizeBox.Text
                Dim minLimit As Decimal? = Nothing
                Dim maxLimit As Decimal? = Nothing
                Dim decVal As Decimal

                If Decimal.TryParse(row.MinBox.Text, decVal) Then minLimit = decVal
                If Decimal.TryParse(row.MaxBox.Text, decVal) Then maxLimit = decVal

                Dim t1 As Decimal? = Nothing
                Dim t2 As Decimal? = Nothing
                Dim t3 As Decimal? = Nothing
                If Decimal.TryParse(row.ThickBoxes(0).Text, decVal) Then t1 = decVal
                If Decimal.TryParse(row.ThickBoxes(1).Text, decVal) Then t2 = decVal
                If Decimal.TryParse(row.ThickBoxes(2).Text, decVal) Then t3 = decVal

                Dim c1 As Decimal? = Nothing
                Dim c2 As Decimal? = Nothing
                Dim c3 As Decimal? = Nothing
                If Decimal.TryParse(row.CamberBoxes(0).Text, decVal) Then c1 = decVal
                If Decimal.TryParse(row.CamberBoxes(1).Text, decVal) Then c2 = decVal
                If Decimal.TryParse(row.CamberBoxes(2).Text, decVal) Then c3 = decVal

                ' Check judgement for this specific blade/row
                Dim rowJudgement = "OK"
                For Each tb In row.ThickBoxes
                    Dim v As Decimal
                    If Decimal.TryParse(tb.Text, v) Then
                        If minLimit.HasValue AndAlso v < minLimit.Value Then rowJudgement = "NG"
                        If maxLimit.HasValue AndAlso v > maxLimit.Value Then rowJudgement = "NG"
                    End If
                Next
                For Each tb In row.CamberBoxes
                    Dim v As Decimal
                    If Decimal.TryParse(tb.Text, v) Then
                        If minLimit.HasValue AndAlso v < minLimit.Value Then rowJudgement = "NG"
                        If maxLimit.HasValue AndAlso v > maxLimit.Value Then rowJudgement = "NG"
                    End If
                Next

                MySqlCls.InsertFeelerGaugeRecord(controlNo, targetCyc, calibDate, type, size, lc, color, location, measRange, minLimit, maxLimit, t1, t2, t3, c1, c2, c3, rowJudgement, "", ProjectSettings.Current.Username)
            Next

            ' Run standard post-calibration flows
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

            ' Row-by-Row TMU calculation based on Thickness Readings
            Dim sumSqMU As Decimal = 0
            Dim hasValidRow As Boolean = False

            For Each row In _observationRows
                Dim rowVals As New List(Of Decimal)()
                For Each tb In row.ThickBoxes
                    Dim val As Decimal
                    If Decimal.TryParse(tb.Text, val) Then rowVals.Add(val)
                Next

                If rowVals.Count > 0 Then
                    ' 1. Calculate Standard Deviation in Microns (for precision)
                    Dim microns = rowVals.Select(Function(v) v * 1000D).ToArray()
                    Dim stdev_microns As Decimal = If(rowVals.Count > 1, calcStdev(microns), 0)
                    
                    ' 2. Convert back to mm and calculate Sigma B1
                    Dim stdev_mm As Decimal = stdev_microns / 1000D
                    Dim sigB1_row As Decimal = stdev_mm / Math.Sqrt(rowVals.Count)
                    
                    ' 3. Calculate Row M.U Squared
                    Dim mu_row_sq As Decimal = (sigB1_row ^ 2) + (masterU ^ 2)
                    
                    ' 4. Add to running total
                    sumSqMU += mu_row_sq
                    hasValidRow = True
                End If
            Next

            If hasValidRow Then
                Dim preTMU = Math.Sqrt(sumSqMU)
                Dim finalTMU = preTMU * 2
                TxtTMU.Text = finalTMU.ToString("F9")
            Else
                TxtTMU.Text = ""
            End If
        Catch ex As Exception
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
                Dim index = _navTextBoxes.IndexOf(tb)
                If index >= 0 AndAlso index < _navTextBoxes.Count - 1 Then
                    _navTextBoxes(index + 1).Focus()
                    e.Handled = True
                ElseIf index = _navTextBoxes.Count - 1 Then
                    TxtTimeIn.Focus()
                    e.Handled = True
                End If
            End If
        End If
    End Sub
End Class
