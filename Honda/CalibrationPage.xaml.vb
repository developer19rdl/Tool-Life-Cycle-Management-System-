Imports System.Windows
Imports System.Data
Imports System.Reflection
Imports ExcelDataReader
Imports Microsoft.Win32
Imports System.IO

Partial Class CalibrationPage
    Private _mySql As New MySQLClass()
    Private _previousTab As UIElement
    Private _currentInstrumentName As String

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        LoadCycleDropdown()
        LoadCalibrationMasterYears()
    End Sub

    Private Function ParseCycleForSort(cycle As String) As Long
        Try
            Dim safeCycle = cycle.Replace("""", "'")
            Dim parts = safeCycle.Split(" "c)
            Dim tag = parts(0) ' e.g. Jan'26, Jul'26, Apr'26
            Dim mPart = tag.Split("'"c)(0)
            Dim yPart = Integer.Parse(tag.Split("'"c)(1))
            ' Parse any 3-letter month abbreviation (not just Jan/Jul)
            Dim mNum As Integer = 1
            Try
                mNum = DateTime.ParseExact(mPart, "MMM", System.Globalization.CultureInfo.InvariantCulture).Month
            Catch
                mNum = If(mPart.Equals("Jan", StringComparison.OrdinalIgnoreCase), 1, 7)
            End Try
            Return (CLng(yPart) * 12L) + mNum
        Catch
            Return 0
        End Try
    End Function

    Private Sub LoadCycleDropdown()
        Try
            Dim combos As ComboBox() = {ComboCycle, ComboNewEditionCycle, ComboReintroCycle, ComboTempIssuanceCycle}

            For Each cmb In combos
                cmb.Items.Clear()
            Next

            Dim activeCycle = GetCurrentCycleName()
            Dim cycleList As New List(Of String)()
            If Not String.IsNullOrEmpty(activeCycle) Then cycleList.Add(activeCycle)

            ' Get unique cycles from both tables to ensure no cycles are omitted
            Try
                Dim dtInter = _mySql.ReadDatatable("SELECT DISTINCT CycleName FROM interchangeability WHERE CycleName IS NOT NULL AND CycleName != ''")
                For Each row As DataRow In dtInter.Rows
                    Dim cName = row("CycleName").ToString()
                    If Not cycleList.Contains(cName) Then
                        cycleList.Add(cName)
                    End If
                Next
            Catch ex As Exception
            End Try

            Try
                Dim dtCal = _mySql.ReadDatatable("SELECT DISTINCT CycleName FROM regular_calibration WHERE CycleName IS NOT NULL AND CycleName != ''")
                For Each row As DataRow In dtCal.Rows
                    Dim cName = row("CycleName").ToString()
                    If Not cycleList.Contains(cName) Then
                        cycleList.Add(cName)
                    End If
                Next
            Catch ex As Exception
            End Try

            ' Sort Chronologically Latest First
            cycleList.Sort(Function(a, b) ParseCycleForSort(b).CompareTo(ParseCycleForSort(a)))

            For Each c In cycleList
                For Each cmb In combos
                    cmb.Items.Add(c)
                Next
            Next

            ' Select newest cycle by default (sync with interchange tab)
            For Each cmb In combos
                If cmb.Items.Count > 0 Then
                    cmb.SelectedIndex = 0
                End If
            Next

            RefreshAllGrids()
        Catch ex As Exception
            MessageBox.Show("Error loading cycles: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Function GetCurrentCycleName() As String
        ' Delegate to MySQLClass which reads Admin-configured date ranges + ActiveCycle override
        Return _mySql.GetActiveCycleName()
    End Function


    Private Sub ComboCycle_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If ComboCycle IsNot Nothing AndAlso ComboCycle.Items.Count > 0 AndAlso ComboCycle.SelectedItem IsNot Nothing AndAlso BtnImportRegularCalib IsNot Nothing Then
            Dim latestCycle = ComboCycle.Items(0).ToString()
            BtnImportRegularCalib.IsEnabled = (ComboCycle.SelectedItem.ToString() = latestCycle)
        End If
        RefreshAllGrids()
    End Sub

    Private Sub LoadCalibrationData(tableName As String, comboCycle As ComboBox, dataGrid As DataGrid, instrumentControl As ItemsControl, gaugeControl As ItemsControl)
        If comboCycle.SelectedItem Is Nothing Then Return
        Dim selectedCycle = comboCycle.SelectedItem.ToString()

        Try
            ' 1. Load DataGrid data
            ' Only new_addition_calibration has the RequestNo column — keep other tables on the original 7 columns
            Dim extraCol As String = If(tableName = "new_addition_calibration", ", RequestNo", "")
            Dim queryGrid = $"SELECT instrument_name, control_no, status, is_calibrated, calibrated_date, due_date, calibration_status{extraCol} " &
                            $"FROM {tableName} " &
                            $"WHERE CycleName = '{selectedCycle.Replace("'", "''")}'"
            Dim dtGrid = _mySql.ReadDatatable(queryGrid)
            dataGrid.ItemsSource = dtGrid.DefaultView

            ' 2. Load and Partition Card Data
            Dim queryCards = $"SELECT DISTINCT r.instrument_name, m.category " &
                             $"FROM {tableName} r " &
                             $"LEFT JOIN calibrationmapping m ON r.instrument_name = m.type_name " &
                             $"WHERE r.CycleName = '{selectedCycle.Replace("'", "''")}' " &
                             $"ORDER BY r.instrument_name ASC"
            Dim dtCards = _mySql.ReadDatatable(queryCards)

            Dim instrumentList As New DataTable()
            instrumentList.Columns.Add("instrument_name")

            Dim gaugeList As New DataTable()
            gaugeList.Columns.Add("instrument_name")

            For Each row As DataRow In dtCards.Rows
                Dim name = row("instrument_name").ToString()
                Dim category = row("category").ToString().ToLower()

                If category = "gauge" OrElse name.ToLower().Contains("gauge") Then
                    gaugeList.Rows.Add(name)
                Else
                    instrumentList.Rows.Add(name)
                End If
            Next

            instrumentControl.ItemsSource = instrumentList.DefaultView
            gaugeControl.ItemsSource = gaugeList.DefaultView

        Catch ex As Exception
            MessageBox.Show($"Error loading calibration data for {tableName}: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub LoadRegularCalibration()
        LoadCalibrationData("regular_calibration", ComboCycle, RegularCalibDataGrid, InstrumentCardsControl, GaugeCardsControl)
    End Sub

    Private Sub LoadNewEditionCalibration()
        LoadCalibrationData("new_addition_calibration", ComboNewEditionCycle, NewEditionCalibDataGrid, NewEditionInstrumentCardsControl, NewEditionGaugeCardsControl)

        ' Populate the Request No filter combo
        Try
            Dim dv = TryCast(NewEditionCalibDataGrid.ItemsSource, DataView)
            If dv Is Nothing Then Return

            Dim currentFilter = If(ComboNewEditionRequestFilter.SelectedItem IsNot Nothing, ComboNewEditionRequestFilter.SelectedItem.ToString(), "All")

            ' Suppress filter handler while repopulating
            _suppressNewEditionFilter = True
            ComboNewEditionRequestFilter.Items.Clear()
            ComboNewEditionRequestFilter.Items.Add("All")

            Dim seen As New System.Collections.Generic.HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each row As DataRowView In dv
                Dim reqNo = row("RequestNo").ToString().Trim()
                If Not String.IsNullOrEmpty(reqNo) AndAlso seen.Add(reqNo) Then
                    ComboNewEditionRequestFilter.Items.Add(reqNo)
                End If
            Next

            ' Re-select previous value or default to All
            Dim matchIdx = ComboNewEditionRequestFilter.Items.IndexOf(currentFilter)
            ComboNewEditionRequestFilter.SelectedIndex = If(matchIdx >= 0, matchIdx, 0)

            _suppressNewEditionFilter = False

            ' Re-apply filter in case it was active
            ApplyNewEditionRequestFilter()
        Catch ex As Exception
            _suppressNewEditionFilter = False
            Console.WriteLine("LoadNewEditionCalibration filter error: " & ex.Message)
        End Try
    End Sub

    Private _suppressNewEditionFilter As Boolean = False

    Private Sub ComboNewEditionRequestFilter_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _suppressNewEditionFilter Then Return
        ApplyNewEditionRequestFilter()
    End Sub

    Private Sub ApplyNewEditionRequestFilter()
        Dim dv = TryCast(NewEditionCalibDataGrid.ItemsSource, DataView)
        If dv Is Nothing Then Return

        If ComboNewEditionRequestFilter.SelectedItem Is Nothing OrElse
           ComboNewEditionRequestFilter.SelectedItem.ToString() = "All" Then
            dv.RowFilter = ""
        Else
            Dim selectedReq = ComboNewEditionRequestFilter.SelectedItem.ToString().Replace("'", "''")
            dv.RowFilter = $"RequestNo = '{selectedReq}'"
        End If
    End Sub

    Private Sub LoadReintroductionCalibration()
        LoadCalibrationData("reintroduction_calibration", ComboReintroCycle, ReintroCalibDataGrid, ReintroInstrumentCardsControl, ReintroGaugeCardsControl)
    End Sub

    Private Sub LoadTempIssuanceCalibration()
        LoadCalibrationData("temp_issuance_calibration", ComboTempIssuanceCycle, TempIssuanceCalibDataGrid, TempIssuanceInstrumentCardsControl, TempIssuanceGaugeCardsControl)
    End Sub

    Private Sub RefreshAllGrids()
        LoadRegularCalibration()
        LoadNewEditionCalibration()
        LoadReintroductionCalibration()
        LoadTempIssuanceCalibration()
        LoadNGList()

        ' If we are currently in the detail view, refresh it as well
        If Not String.IsNullOrEmpty(_currentInstrumentName) Then
            If CardDetailTabView.Visibility = Visibility.Visible Then
                LoadCardDetailData("regular_calibration", ComboCycle, DetailCalibDataGrid, _currentInstrumentName)
            ElseIf NewEditionCardDetailTabView.Visibility = Visibility.Visible Then
                LoadCardDetailData("new_addition_calibration", ComboNewEditionCycle, NewEditionDetailCalibDataGrid, _currentInstrumentName)
            ElseIf ReintroCardDetailTabView.Visibility = Visibility.Visible Then
                LoadCardDetailData("reintroduction_calibration", ComboReintroCycle, ReintroDetailCalibDataGrid, _currentInstrumentName)
            ElseIf TempIssuanceCardDetailTabView.Visibility = Visibility.Visible Then
                LoadCardDetailData("temp_issuance_calibration", ComboTempIssuanceCycle, TempIssuanceDetailCalibDataGrid, _currentInstrumentName)
            End If
        End If
    End Sub

    Private Sub Button_Vernier300_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New VernierCaliper300()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Vernier600_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New VernierCaliper600()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_VernierLowForce_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New VernierCaliperLowForce()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_DepthVernier_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New DepthVernierCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_TPM25_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ThreadPitchMicrometer25()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_TPM50_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ThreadPitchMicrometer50()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_TPM70_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ThreadPitchMicrometer70()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_TPM100_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ThreadPitchMicrometer100()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Point25_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New PointMicrometer25()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Point50_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New PointMicrometer50()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Inside25_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New InsideMicrometer25()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Inside50_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New InsideMicrometer50()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Groove25_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New GrooveMicrometer25()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Groove50_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New GrooveMicrometer50()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Geartooth50_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New GeartoothMicrometer50()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_ExtMic100_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ExternalMicrometer100()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_DiscMic_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New DiscMicrometer75()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_DepthMic_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New DepthMicrometer()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_BladeMic_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New BladeMicrometer75()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_BoreGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New BoreGauge3()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_HeightMaster_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New HeightMastergaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Passameter60_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New Passameter60()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_DialGuage100_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New DialGuage100()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_Dialtestindicator100_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New Dialtestindicator100()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_HeightGauge600_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New HeightGauge600()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_KeyGrooveGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New KeyGrooveGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_PlainPlugGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New PlainPlugGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_PlainRingGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New PlainRingGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_SnapGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New SnapGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_SpecialHeightGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New SpecialHeightGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_SplineRingGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New SplineRingGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_StraightPlugGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New StraightPlugGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_ThreadRingGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ThreadRingGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_ThreadPlugGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ThreadPlugGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_TaperPlugGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New TaperPlugGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_TaperRingGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New TaperRingGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_FeelerGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New FeelerGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_RadiusGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New RadiusGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_ThreadPitchGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ThreadPitchGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub Button_SpecialDepthGauge_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New SpecialDepthGaugeCalibration()
        AddHandler frm.Closed, AddressOf OnCalibrationWindowClosed
        frm.Show()
    End Sub

    Private Sub OnCalibrationWindowClosed(sender As Object, e As EventArgs)
        RefreshAllGrids()
    End Sub

    Private Sub TabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton As Button = DirectCast(sender, Button)

        ' Reset all tab buttons
        BtnRegularCalib.Tag = ""
        BtnNewEditionCalib.Tag = ""
        BtnReintroduction.Tag = ""
        BtnTempIssuance.Tag = ""
        BtnCalibrationMaster.Tag = ""
        BtnNGList.Tag = ""
        BtnTest.Tag = ""

        ' Set selected tab button
        clickedButton.Tag = "Selected"

        ' Hide all tab views
        RegularCalibTabView.Visibility = Visibility.Collapsed
        NewEditionCalibTabView.Visibility = Visibility.Collapsed
        ReintroTabView.Visibility = Visibility.Collapsed
        TempIssuanceTabView.Visibility = Visibility.Collapsed
        CalibrationMasterTabView.Visibility = Visibility.Collapsed
        NGListTabView.Visibility = Visibility.Collapsed
        TestTabView.Visibility = Visibility.Collapsed

        ' Show corresponding tab view
        If clickedButton Is BtnRegularCalib Then
            RegularCalibTabView.Visibility = Visibility.Visible
            _mySql.CurrentTargetTrackingTable = "regular_calibration"
            ' Default to Status Tab if in Regular Calibration
            If BtnStatusTab.Tag.ToString() = "Selected" Then
                RefreshAllGrids()
            End If
        ElseIf clickedButton Is BtnNewEditionCalib Then
            NewEditionCalibTabView.Visibility = Visibility.Visible
            _mySql.CurrentTargetTrackingTable = "new_addition_calibration"
        ElseIf clickedButton Is BtnReintroduction Then
            ReintroTabView.Visibility = Visibility.Visible
            _mySql.CurrentTargetTrackingTable = "reintroduction_calibration"
        ElseIf clickedButton Is BtnTempIssuance Then
            TempIssuanceTabView.Visibility = Visibility.Visible
            _mySql.CurrentTargetTrackingTable = "temp_issuance_calibration"
        ElseIf clickedButton Is BtnCalibrationMaster Then
            CalibrationMasterTabView.Visibility = Visibility.Visible
            ' Default to Details sub-tab
            BtnMasterDetails.Tag = "Selected"
            BtnMasterDatabase.Tag = ""
            MasterDetailsView.Visibility = Visibility.Visible
            MasterDatabaseView.Visibility = Visibility.Collapsed
            LoadCalibrationMasterDetails() ' Initial load
        ElseIf clickedButton Is BtnNGList Then
            NGListTabView.Visibility = Visibility.Visible
            LoadNGList()
        ElseIf clickedButton Is BtnTest Then
            TestTabView.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub MasterSubTabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton As Button = DirectCast(sender, Button)

        ' Reset all master sub-tab buttons
        BtnMasterDetails.Tag = ""
        BtnMasterDatabase.Tag = ""

        ' Set selected sub-tab button
        clickedButton.Tag = "Selected"

        ' Hide all master sub-tab views
        MasterDetailsView.Visibility = Visibility.Collapsed
        MasterDatabaseView.Visibility = Visibility.Collapsed

        ' Show corresponding sub-tab view
        If clickedButton Is BtnMasterDetails Then
            MasterDetailsView.Visibility = Visibility.Visible
            LoadCalibrationMasterDetails()
        ElseIf clickedButton Is BtnMasterDatabase Then
            MasterDatabaseView.Visibility = Visibility.Visible
            LoadCalibrationMasterTypesCards()
        End If
    End Sub

    Private Sub LoadCalibrationMasterYears()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT YEAR(CalDate) as CalYear FROM calibrationmaster_details WHERE CalDate IS NOT NULL ORDER BY CalYear DESC")
            ComboYear.Items.Clear()
            For Each row As DataRow In dt.Rows
                ComboYear.Items.Add(row("CalYear").ToString())
            Next

            If ComboYear.Items.Count > 0 Then
                ComboYear.SelectedIndex = 0
            Else
                ' If no records, just add current year as placeholder
                ComboYear.Items.Add(DateTime.Now.Year.ToString())
                ComboYear.SelectedIndex = 0
            End If
        Catch ex As Exception
            Console.WriteLine("Error loading master years: " & ex.Message)
        End Try
    End Sub

    Private Sub ComboYear_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        LoadCalibrationMasterDetails()
    End Sub

    Private Sub LoadCalibrationMasterDetails()
        If ComboYear.SelectedItem Is Nothing Then Return
        Dim selectedYear = ComboYear.SelectedItem.ToString()

        Try
            Dim query = $"SELECT * FROM calibrationmaster_details WHERE YEAR(CalDate) = {selectedYear} ORDER BY CalDate DESC"
            Dim dt = _mySql.ReadDatatable(query)
            MasterDetailsDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading calibration master details: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnViewMasterDocument_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim row = TryCast(btn.DataContext, DataRowView)
        If row Is Nothing Then Return

        Dim uploadedDoc = row("uploaded_doc").ToString()
        Dim masterName = row("Description").ToString()

        If String.IsNullOrEmpty(uploadedDoc) Then
            MessageBox.Show("No document uploaded for this record.", "Information", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        Try
            Dim appPath = AppDomain.CurrentDomain.BaseDirectory
            Dim filePath = System.IO.Path.Combine(appPath, uploadedDoc)

            If System.IO.File.Exists(filePath) Then
                Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
            Else
                MessageBox.Show("The document file could not be found." & vbCrLf & filePath, "File Not Found", MessageBoxButton.OK, MessageBoxImage.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show("Error opening document: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub LoadCalibrationMasterTypesCards()
        Try
            Dim dt = _mySql.GetCalibrationMasterTypes()
            If dt IsNot Nothing AndAlso dt.Columns.Contains("CalibrationMasterName") Then
                ' Map column so the common card binding (instrument_name) works.
                dt.Columns("CalibrationMasterName").ColumnName = "instrument_name"
                CalibrationMasterCardsControl.ItemsSource = dt.DefaultView
            Else
                CalibrationMasterCardsControl.ItemsSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading master cards: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub SubTabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton As Button = DirectCast(sender, Button)

        ' Which tab group is this button in?
        Dim parentStack = TryCast(clickedButton.Parent, StackPanel)
        If parentStack Is Nothing Then Return

        ' Reset tags for all buttons in this group
        For Each child In parentStack.Children
            If TypeOf child Is Button Then
                DirectCast(child, Button).Tag = ""
            End If
        Next

        ' Set selected tab tag
        clickedButton.Tag = "Selected"

        ' Handle Regular Calibration Sub-tabs
        If clickedButton Is BtnStatusTab OrElse clickedButton Is BtnInstrumentTab OrElse clickedButton Is BtnGaugeTab Then
            StatusTabView.Visibility = If(clickedButton Is BtnStatusTab, Visibility.Visible, Visibility.Collapsed)
            InstrumentTabView.Visibility = If(clickedButton Is BtnInstrumentTab, Visibility.Visible, Visibility.Collapsed)
            GaugeTabView.Visibility = If(clickedButton Is BtnGaugeTab, Visibility.Visible, Visibility.Collapsed)
            CardDetailTabView.Visibility = Visibility.Collapsed
            If clickedButton Is BtnStatusTab Then RefreshAllGrids()

            ' Handle New Edition Sub-tabs
        ElseIf clickedButton Is BtnNewEditionStatusTab OrElse clickedButton Is BtnNewEditionInstrumentTab OrElse clickedButton Is BtnNewEditionGaugeTab Then
            NewEditionStatusTabView.Visibility = If(clickedButton Is BtnNewEditionStatusTab, Visibility.Visible, Visibility.Collapsed)
            NewEditionInstrumentTabView.Visibility = If(clickedButton Is BtnNewEditionInstrumentTab, Visibility.Visible, Visibility.Collapsed)
            NewEditionGaugeTabView.Visibility = If(clickedButton Is BtnNewEditionGaugeTab, Visibility.Visible, Visibility.Collapsed)
            NewEditionCardDetailTabView.Visibility = Visibility.Collapsed
            If clickedButton Is BtnNewEditionStatusTab Then RefreshAllGrids()

            ' Handle Reintroduction Sub-tabs
        ElseIf clickedButton Is BtnReintroStatusTab OrElse clickedButton Is BtnReintroInstrumentTab OrElse clickedButton Is BtnReintroGaugeTab Then
            ReintroStatusTabView.Visibility = If(clickedButton Is BtnReintroStatusTab, Visibility.Visible, Visibility.Collapsed)
            ReintroInstrumentTabView.Visibility = If(clickedButton Is BtnReintroInstrumentTab, Visibility.Visible, Visibility.Collapsed)
            ReintroGaugeTabView.Visibility = If(clickedButton Is BtnReintroGaugeTab, Visibility.Visible, Visibility.Collapsed)
            ReintroCardDetailTabView.Visibility = Visibility.Collapsed
            If clickedButton Is BtnReintroStatusTab Then RefreshAllGrids()

            ' Handle Temp Issuance Sub-tabs
        ElseIf clickedButton Is BtnTempIssuanceStatusTab OrElse clickedButton Is BtnTempIssuanceInstrumentTab OrElse clickedButton Is BtnTempIssuanceGaugeTab Then
            TempIssuanceStatusTabView.Visibility = If(clickedButton Is BtnTempIssuanceStatusTab, Visibility.Visible, Visibility.Collapsed)
            TempIssuanceInstrumentTabView.Visibility = If(clickedButton Is BtnTempIssuanceInstrumentTab, Visibility.Visible, Visibility.Collapsed)
            TempIssuanceGaugeTabView.Visibility = If(clickedButton Is BtnTempIssuanceGaugeTab, Visibility.Visible, Visibility.Collapsed)
            TempIssuanceCardDetailTabView.Visibility = Visibility.Collapsed
            If clickedButton Is BtnTempIssuanceStatusTab Then RefreshAllGrids()
        End If
    End Sub

    Private Sub TestTabSelector_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton As Button = DirectCast(sender, Button)

        ' Reset test tab selector tags
        BtnTestInstruments.Tag = ""
        BtnTestGauges.Tag = ""

        ' Set selected tab tag
        clickedButton.Tag = "Selected"

        ' Toggle visibility
        If clickedButton Is BtnTestInstruments Then
            TestInstrumentsPanel.Visibility = Visibility.Visible
            TestGaugesPanel.Visibility = Visibility.Collapsed
        ElseIf clickedButton Is BtnTestGauges Then
            TestGaugesPanel.Visibility = Visibility.Visible
            TestInstrumentsPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub Card_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Dim border = DirectCast(sender, Border)
        Dim row = TryCast(border.DataContext, DataRowView)
        If row Is Nothing Then Return

        Dim instrumentName = row("instrument_name").ToString()
        ShowCardDetails(instrumentName)
    End Sub

    Private Sub MasterCard_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Dim border = DirectCast(sender, Border)
        Dim row = TryCast(border.DataContext, DataRowView)
        If row Is Nothing Then Return

        Dim masterName = row("instrument_name").ToString()

        ' Navigate to generic calibration master management page entirely
        Me.NavigationService.Navigate(New CalibrationMasterManagementPage(masterName))
    End Sub

    Private Sub ShowCardDetails(instrumentName As String)
        ' Store current tab to return to it later
        If InstrumentTabView.Visibility = Visibility.Visible Then
            _previousTab = InstrumentTabView
        ElseIf GaugeTabView.Visibility = Visibility.Visible Then
            _previousTab = GaugeTabView
        ElseIf NewEditionInstrumentTabView.Visibility = Visibility.Visible Then
            _previousTab = NewEditionInstrumentTabView
        ElseIf NewEditionGaugeTabView.Visibility = Visibility.Visible Then
            _previousTab = NewEditionGaugeTabView
        ElseIf ReintroInstrumentTabView.Visibility = Visibility.Visible Then
            _previousTab = ReintroInstrumentTabView
        ElseIf ReintroGaugeTabView.Visibility = Visibility.Visible Then
            _previousTab = ReintroGaugeTabView
        ElseIf TempIssuanceInstrumentTabView.Visibility = Visibility.Visible Then
            _previousTab = TempIssuanceInstrumentTabView
        ElseIf TempIssuanceGaugeTabView.Visibility = Visibility.Visible Then
            _previousTab = TempIssuanceGaugeTabView
        End If

        ' Hide card views
        InstrumentTabView.Visibility = Visibility.Collapsed
        GaugeTabView.Visibility = Visibility.Collapsed
        NewEditionInstrumentTabView.Visibility = Visibility.Collapsed
        NewEditionGaugeTabView.Visibility = Visibility.Collapsed
        ReintroInstrumentTabView.Visibility = Visibility.Collapsed
        ReintroGaugeTabView.Visibility = Visibility.Collapsed
        TempIssuanceInstrumentTabView.Visibility = Visibility.Collapsed
        TempIssuanceGaugeTabView.Visibility = Visibility.Collapsed

        _currentInstrumentName = instrumentName

        ' Show detail view based on active major tab
        If RegularCalibTabView.Visibility = Visibility.Visible Then
            CardDetailTabView.Visibility = Visibility.Visible
            TxtDetailTitle.Text = $"Calibration: {instrumentName}"
            LoadCardDetailData("regular_calibration", ComboCycle, DetailCalibDataGrid, instrumentName)
        ElseIf NewEditionCalibTabView.Visibility = Visibility.Visible Then
            NewEditionCardDetailTabView.Visibility = Visibility.Visible
            TxtNewEditionDetailTitle.Text = $"Calibration: {instrumentName}"
            LoadCardDetailData("new_addition_calibration", ComboNewEditionCycle, NewEditionDetailCalibDataGrid, instrumentName)
        ElseIf ReintroTabView.Visibility = Visibility.Visible Then
            ReintroCardDetailTabView.Visibility = Visibility.Visible
            TxtReintroDetailTitle.Text = $"Calibration: {instrumentName}"
            LoadCardDetailData("reintroduction_calibration", ComboReintroCycle, ReintroDetailCalibDataGrid, instrumentName)
        ElseIf TempIssuanceTabView.Visibility = Visibility.Visible Then
            TempIssuanceCardDetailTabView.Visibility = Visibility.Visible
            TxtTempIssuanceDetailTitle.Text = $"Calibration: {instrumentName}"
            LoadCardDetailData("temp_issuance_calibration", ComboTempIssuanceCycle, TempIssuanceDetailCalibDataGrid, instrumentName)
        End If
    End Sub

    Private Sub BtnViewGroupResult_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrEmpty(_currentInstrumentName) Then Return

        Dim selectedCycle = ""
        If RegularCalibTabView.Visibility = Visibility.Visible AndAlso ComboCycle.SelectedItem IsNot Nothing Then
            selectedCycle = ComboCycle.SelectedItem.ToString()
        ElseIf NewEditionCalibTabView.Visibility = Visibility.Visible AndAlso ComboNewEditionCycle.SelectedItem IsNot Nothing Then
            selectedCycle = ComboNewEditionCycle.SelectedItem.ToString()
        ElseIf ReintroTabView.Visibility = Visibility.Visible AndAlso ComboReintroCycle.SelectedItem IsNot Nothing Then
            selectedCycle = ComboReintroCycle.SelectedItem.ToString()
        ElseIf TempIssuanceTabView.Visibility = Visibility.Visible AndAlso ComboTempIssuanceCycle.SelectedItem IsNot Nothing Then
            selectedCycle = ComboTempIssuanceCycle.SelectedItem.ToString()
        Else
            Return
        End If

        ' Open the new summary window
        Dim frm As New CalibrationResultSummaryWindow(_currentInstrumentName, selectedCycle)
        frm.Owner = Window.GetWindow(Me)
        frm.ShowDialog()
    End Sub

    Private Sub LoadCardDetailData(tableName As String, comboCycle As ComboBox, dataGrid As DataGrid, instrumentName As String)
        If comboCycle.SelectedItem Is Nothing Then Return
        Dim selectedCycle = comboCycle.SelectedItem.ToString()

        Try
            Dim query = $"SELECT instrument_name, control_no, status, is_calibrated, calibrated_date, due_date, calibration_status " &
                        $"FROM {tableName} " &
                        $"WHERE CycleName = '{selectedCycle.Replace("'", "''")}' " &
                        $"AND instrument_name = '{instrumentName.Replace("'", "''")}'"
            Dim dt = _mySql.ReadDatatable(query)
            dataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show($"Error loading details for {tableName}: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnBackToCards_Click(sender As Object, e As RoutedEventArgs)
        CardDetailTabView.Visibility = Visibility.Collapsed
        NewEditionCardDetailTabView.Visibility = Visibility.Collapsed
        ReintroCardDetailTabView.Visibility = Visibility.Collapsed
        TempIssuanceCardDetailTabView.Visibility = Visibility.Collapsed

        If _previousTab IsNot Nothing Then
            _previousTab.Visibility = Visibility.Visible
        Else
            ' Fallback to status if no previous tab
            If RegularCalibTabView.Visibility = Visibility.Visible Then
                StatusTabView.Visibility = Visibility.Visible
            ElseIf NewEditionCalibTabView.Visibility = Visibility.Visible Then
                NewEditionStatusTabView.Visibility = Visibility.Visible
            ElseIf ReintroTabView.Visibility = Visibility.Visible Then
                ReintroStatusTabView.Visibility = Visibility.Visible
            ElseIf TempIssuanceTabView.Visibility = Visibility.Visible Then
                TempIssuanceStatusTabView.Visibility = Visibility.Visible
            End If
        End If
    End Sub

    Private Sub BtnViewResult_Click(sender As Object, e As RoutedEventArgs)
        Dim frm As New ResultWindow()
        frm.Owner = Window.GetWindow(Me)
        frm.ShowDialog()
    End Sub

    Private Sub ControlNo_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim row = TryCast(btn.DataContext, DataRowView)
        If row Is Nothing Then Return

        Dim controlNo = row("control_no").ToString()

        ' Check if already calibrated
        Dim isCalib = row("is_calibrated").ToString()
        Dim selectedCycle = ComboCycle.SelectedItem.ToString()
        Dim activeCycle = GetCurrentCycleName()

        If isCalib.Equals("YES", StringComparison.OrdinalIgnoreCase) Then
            ' Only allow editing if it's the active cycle
            If selectedCycle <> activeCycle Then
                MessageBox.Show("This instrument has already been calibrated in a historical cycle and cannot be modified.", "Historical Record Protected", MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            ' If active cycle, ask for confirmation to edit
            Dim result = MessageBox.Show("This instrument has already been calibrated in the current cycle. Do you want to edit the existing record?", "Edit Calibration", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result <> MessageBoxResult.Yes Then Return
        End If

        ' 1. Get Size from inventory (interchangeability/department_list)
        Dim size = _mySql.GetInstrumentSizeByControlNo(controlNo)

        ' 2. Resolve Form Name from calibrationmapping
        Dim formName = _mySql.GetCalibrationMappingForm(controlNo, size)

        If formName = "DialGauge" Then
            formName = "DialGuage100"
        ElseIf formName = "DialTestIndicator" Then
            formName = "Dialtestindicator100"
        ElseIf formName = "BoreGauge" Then
            formName = "BoreGauge3"
        ElseIf formName = "HeightGauge300" Then
            formName = "HeightGauge600"
        End If

        If Not String.IsNullOrEmpty(formName) Then
            Try
                ' 3. Open Window using reflection
                ' RootNamespace is "Honda" as seen in .vbproj
                Dim typeName = "Honda." & formName
                Dim winType = Type.GetType(typeName)

                ' Fallback if Honda prefix isn't needed or different
                If winType Is Nothing Then
                    winType = Assembly.GetExecutingAssembly().GetType(typeName)
                End If
                If winType Is Nothing Then
                    winType = Assembly.GetExecutingAssembly().GetType(formName)
                End If

                If winType IsNot Nothing Then
                    Dim win = DirectCast(Activator.CreateInstance(winType), Window)
                    win.Owner = Window.GetWindow(Me)

                    ' Set Edit mode if applicable
                    If isCalib.Equals("YES", StringComparison.OrdinalIgnoreCase) Then
                        Dim propEdit = winType.GetProperty("IsEditMode")
                        If propEdit IsNot Nothing Then propEdit.SetValue(win, True)

                        Dim propCycle = winType.GetProperty("TargetCycle")
                        If propCycle IsNot Nothing Then propCycle.SetValue(win, selectedCycle)
                    End If

                    ' Auto-populate Control No, and directly populate Size ComboBox
                    Try
                        ' Set Control No
                        Dim txtCtrl = TryCast(win.FindName("TextBox1"), TextBox)
                        If txtCtrl Is Nothing Then txtCtrl = TryCast(win.FindName("TxtControlNo"), TextBox)
                        If txtCtrl IsNot Nothing Then txtCtrl.Text = controlNo

                        ' Determine category for size lookup
                        Dim mappingCat = _mySql.ReadDatatable(
                                $"SELECT category FROM calibrationmapping WHERE type_name = '{formName.Replace("'", "''")}' LIMIT 1")
                        Dim sizeCat As String = "Instrument"
                        If mappingCat.Rows.Count > 0 AndAlso mappingCat.Rows(0)("category").ToString().ToLower() = "gauge" Then
                            sizeCat = "Gauge"
                        End If

                        ' TextBox21 is a ComboBox in instrument forms — populate its ItemsSource directly
                        Dim cmbSize = TryCast(win.FindName("TextBox21"), ComboBox)
                        If cmbSize IsNot Nothing Then
                            Dim prefix = MySQLClass.ExtractPrefix(controlNo)
                            Dim sizes = _mySql.GetSizesForGroup(prefix, sizeCat)
                            cmbSize.ItemsSource = sizes
                            ' Auto-select matching size
                            If Not String.IsNullOrEmpty(size) AndAlso sizes.Count > 0 Then
                                Dim cleanTarget = size.ToLower().Replace("mm", "").Replace(" ", "").Trim()
                                Dim matched = sizes.FirstOrDefault(
                                        Function(s) s.ToLower().Replace("mm", "").Replace(" ", "").Trim() = cleanTarget)
                                If matched IsNot Nothing Then
                                    cmbSize.SelectedItem = matched
                                ElseIf sizes.Count = 1 Then
                                    cmbSize.SelectedIndex = 0
                                End If
                            End If
                        Else
                            ' Fallback for gauge forms that use plain TextBox (TxtSize / TxtGoSize)
                            Dim txtSizeFallback = TryCast(win.FindName("TxtSize"), TextBox)
                            If txtSizeFallback Is Nothing Then txtSizeFallback = TryCast(win.FindName("TxtGoSize"), TextBox)
                            If txtSizeFallback IsNot Nothing Then txtSizeFallback.Text = size
                        End If
                    Catch
                        ' Ignore if controls don't exist
                    End Try

                    win.ShowDialog()

                    ' Refresh grid after potential calibration
                    RefreshAllGrids()
                Else
                    MessageBox.Show($"Calibration form class '{formName}' could not be found in the application.{vbCrLf}Please check if the form name in settings matches the actual window name.", "Mapping Resolution Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                End If
            Catch ex As Exception
                MessageBox.Show("Error opening calibration form: " & ex.Message, "System Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        Else
            MessageBox.Show($"No calibration form mapping found for this instrument group/prefix and calibration category.{vbCrLf}{vbCrLf}Control No: {controlNo}{vbCrLf}Size: {If(String.IsNullOrEmpty(size), "Unknown", size)}{vbCrLf}{vbCrLf}You can configure this in Settings -> Calibration Mapping.", "Mapping Not Found", MessageBoxButton.OK, MessageBoxImage.Information)
        End If
    End Sub

    Private Sub LoadNGList()
        If ComboCycle.SelectedItem Is Nothing Then Return
        Dim selectedCycle = ComboCycle.SelectedItem.ToString()

        Try
            Dim query = $"SELECT * FROM ng_list WHERE CycleName = '{selectedCycle.Replace("'", "''")}' ORDER BY id DESC"
            Dim dt = _mySql.ReadDatatable(query)
            NGListDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading NG list: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnImportRegularCalib_Click(sender As Object, e As RoutedEventArgs)
        If ComboCycle.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a cycle first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim selectedCycle = ComboCycle.SelectedItem.ToString()
        Dim latestCycle = ComboCycle.Items(0).ToString()
        If selectedCycle <> latestCycle Then
            MessageBox.Show("You can only import data to the latest cycle.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Excel Files (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv"
        If ofd.ShowDialog() = True Then
            Try
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

                Dim dtExcel As DataTable = Nothing

                Using stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read)
                    Using reader = If(ofd.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase),
                                     ExcelReaderFactory.CreateCsvReader(stream),
                                     ExcelReaderFactory.CreateReader(stream))

                        Dim conf = New ExcelDataSetConfiguration() With {
                            .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                .UseHeaderRow = True
                            }
                        }

                        Dim dataSet = reader.AsDataSet(conf)
                        If dataSet.Tables.Count > 0 Then
                            dtExcel = dataSet.Tables(0)
                        End If
                    End Using
                End Using

                If dtExcel Is Nothing OrElse dtExcel.Rows.Count = 0 Then
                    MessageBox.Show("No data found in the selected file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return
                End If

                Dim importedCount As Integer = 0
                For Each row As DataRow In dtExcel.Rows
                    Try
                        ' Extract values using explicit column names
                        Dim instName As String = If(dtExcel.Columns.Contains("instrument_name"), row("instrument_name").ToString(), "")
                        Dim ctrlNo As String = If(dtExcel.Columns.Contains("control_no"), row("control_no").ToString(), "").Trim()
                        Dim status As String = If(dtExcel.Columns.Contains("status"), row("status").ToString(), "")
                        Dim isCalib As String = If(dtExcel.Columns.Contains("is_calibrated"), row("is_calibrated").ToString(), "")
                        Dim calibDate As String = If(dtExcel.Columns.Contains("calibrated_date"), row("calibrated_date").ToString(), "")
                        Dim dueDate As String = If(dtExcel.Columns.Contains("due_date"), row("due_date").ToString(), "")
                        Dim calibStatus As String = If(dtExcel.Columns.Contains("calibration_status"), row("calibration_status").ToString(), "")

                        If String.IsNullOrEmpty(ctrlNo) Then Continue For

                        Dim calibDateParsed As String = "NULL"
                        Dim dueDateParsed As String = "NULL"

                        Dim tempDate As DateTime
                        If DateTime.TryParse(calibDate, tempDate) Then
                            calibDateParsed = $"'{tempDate.ToString("yyyy-MM-dd HH:mm:ss")}'"
                        End If
                        If DateTime.TryParse(dueDate, tempDate) Then
                            dueDateParsed = $"'{tempDate.ToString("yyyy-MM-dd HH:mm:ss")}'"
                        End If

                        ' Delete existing record with same control_no and CycleName
                        Dim deleteQuery = $"DELETE FROM regular_calibration WHERE control_no = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{selectedCycle.Replace("'", "''")}'"
                        _mySql.ExecuteQuery(deleteQuery)

                        ' Insert new record (id is omitted so it auto-increments in MySQL)
                        Dim insertQuery = $"INSERT INTO regular_calibration (instrument_name, control_no, status, is_calibrated, calibrated_date, due_date, calibration_status, CycleName) " &
                                          $"VALUES ('{instName.Replace("'", "''")}', '{ctrlNo.Replace("'", "''")}', '{status.Replace("'", "''")}', " &
                                          $"'{isCalib.Replace("'", "''")}', {calibDateParsed}, {dueDateParsed}, '{calibStatus.Replace("'", "''")}', '{selectedCycle.Replace("'", "''")}')"

                        _mySql.ExecuteQuery(insertQuery)
                        importedCount += 1
                    Catch ex As Exception
                        ' Skip on error for a specific row
                        Console.WriteLine($"Error importing row: {ex.Message}")
                    End Try
                Next

                MessageBox.Show($"Successfully imported {importedCount} records.", "Import Success", MessageBoxButton.OK, MessageBoxImage.Information)
                RefreshAllGrids()

            Catch ex As Exception
                MessageBox.Show("Error importing file: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

End Class
