Imports System.Windows
Imports System.Windows.Controls
Imports System.Data
Imports System.Reflection
Imports System.Linq

Public Class InventoryManagementPage
    Inherits Page

    Private _mySql As New MySQLClass()
    Private _selectedDept As String = "All"
    Private _isLoading As Boolean = False
    
    Private _bulkScanItems As New Collections.ObjectModel.ObservableCollection(Of SearchItem)()
    Private _controlWiseItems As New Collections.ObjectModel.ObservableCollection(Of SearchItem)()
    Private _deptWiseItems As New Collections.ObjectModel.ObservableCollection(Of SearchItem)()
    Private _databaseWiseItems As New Collections.ObjectModel.ObservableCollection(Of SearchItem)()
    Private _isBulkScanning As Boolean = False
    Private _gunClient As RFIDGunClient
    Private _usbScanCts As System.Threading.CancellationTokenSource
    Private _selectedBulkDepts As New List(Of String)()
    Private _selectedBulkTypes As New List(Of String)()
    Private _calibAuditItems As New Collections.ObjectModel.ObservableCollection(Of SearchItem)()
    Private _calibratedCount As Integer = 0
    Private _notCalibratedCount As Integer = 0
    Private _exclusionAddItems As New Collections.ObjectModel.ObservableCollection(Of ExclusionItem)()
    Private _exclusionRevokeItems As New Collections.ObjectModel.ObservableCollection(Of ExclusionItem)()

    Public Class ExclusionItem
        Implements System.ComponentModel.INotifyPropertyChanged
        Public Event PropertyChanged As System.ComponentModel.PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

        Protected Sub OnPropertyChanged(name As String)
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(name))
        End Sub

        Private _rfidTag As String
        Public Property RfidTag As String
            Get
                Return _rfidTag
            End Get
            Set(value As String)
                _rfidTag = value
                OnPropertyChanged("RfidTag")
            End Set
        End Property

        Private _controlNo As String
        Public Property ControlNo As String
            Get
                Return _controlNo
            End Get
            Set(value As String)
                _controlNo = value
                OnPropertyChanged("ControlNo")
            End Set
        End Property

        Private _instrumentName As String
        Public Property InstrumentName As String
            Get
                Return _instrumentName
            End Get
            Set(value As String)
                _instrumentName = value
                OnPropertyChanged("InstrumentName")
            End Set
        End Property

        Private _isInInterchange As Boolean = False
        Public Property IsInInterchange As Boolean
            Get
                Return _isInInterchange
            End Get
            Set(value As Boolean)
                _isInInterchange = value
                OnPropertyChanged("IsInInterchange")
            End Set
        End Property

        Private _isSelected As Boolean = True ' Checked by default
        Public Property IsSelected As Boolean
            Get
                Return _isSelected
            End Get
            Set(value As Boolean)
                _isSelected = value
                OnPropertyChanged("IsSelected")
            End Set
        End Property
    End Class

    Public Class SearchItem
        Implements System.ComponentModel.INotifyPropertyChanged
        Public Event PropertyChanged As System.ComponentModel.PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

        Protected Sub OnPropertyChanged(name As String)
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(name))
        End Sub

        Private _rfidTag As String
        Public Property RfidTag As String
            Get
                Return _rfidTag
            End Get
            Set(value As String)
                _rfidTag = value
                OnPropertyChanged("RfidTag")
            End Set
        End Property

        Private _controlNo As String
        Public Property ControlNo As String
            Get
                Return _controlNo
            End Get
            Set(value As String)
                _controlNo = value
                OnPropertyChanged("ControlNo")
            End Set
        End Property

        Private _instrumentName As String
        Public Property InstrumentName As String
            Get
                Return _instrumentName
            End Get
            Set(value As String)
                _instrumentName = value
                OnPropertyChanged("InstrumentName")
            End Set
        End Property

        Private _cycleName As String
        Public Property CycleName As String
            Get
                Return _cycleName
            End Get
            Set(value As String)
                _cycleName = value
                OnPropertyChanged("CycleName")
            End Set
        End Property

        Private _department As String
        Public Property Department As String
            Get
                Return _department
            End Get
            Set(value As String)
                _department = value
                OnPropertyChanged("Department")
            End Set
        End Property

        Private _actionDate As String
        Public Property ActionDate As String
            Get
                Return _actionDate
            End Get
            Set(value As String)
                _actionDate = value
                OnPropertyChanged("ActionDate")
            End Set
        End Property

        Private _actionTime As String
        Public Property ActionTime As String
            Get
                Return _actionTime
            End Get
            Set(value As String)
                _actionTime = value
                OnPropertyChanged("ActionTime")
            End Set
        End Property

        Private _color As String
        Public Property Color As String
            Get
                Return _color
            End Get
            Set(value As String)
                _color = value
                OnPropertyChanged("Color")
            End Set
        End Property

        Private _status As String
        Public Property Status As String
            Get
                Return _status
            End Get
            Set(value As String)
                _status = value
                OnPropertyChanged("Status")
            End Set
        End Property

        Private _isCalibrated As String
        Public Property IsCalibrated As String
            Get
                Return _isCalibrated
            End Get
            Set(value As String)
                _isCalibrated = value
                OnPropertyChanged("IsCalibrated")
            End Set
        End Property

        Private _calibratedDate As String
        Public Property CalibratedDate As String
            Get
                Return _calibratedDate
            End Get
            Set(value As String)
                _calibratedDate = value
                OnPropertyChanged("CalibratedDate")
            End Set
        End Property

        Private _calibrationStatus As String
        Public Property CalibrationStatus As String
            Get
                Return _calibrationStatus
            End Get
            Set(value As String)
                _calibrationStatus = value
                OnPropertyChanged("CalibrationStatus")
            End Set
        End Property

        Private _isSelected As Boolean = False
        Public Property IsSelected As Boolean
            Get
                Return _isSelected
            End Get
            Set(value As Boolean)
                _isSelected = value
                OnPropertyChanged("IsSelected")
            End Set
        End Property

        Private _found As String = "No"
        Public Property Found As String
            Get
                Return _found
            End Get
            Set(value As String)
                _found = value
                OnPropertyChanged("Found")
            End Set
        End Property
    End Class

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _isLoading = True
        LoadCycles()
        PopulateYearFilter()
        PopulateColorFilter()
        PopulateStatusFilter()
        PopulateInterchangeStatusFilter()
        LoadDepartments()
        _isLoading = False

        BulkScanDataGrid.ItemsSource = _bulkScanItems
        ControlWiseDataGrid.ItemsSource = _controlWiseItems
        DeptWiseDataGrid.ItemsSource = _deptWiseItems
        DatabaseWiseDataGrid.ItemsSource = _databaseWiseItems
        ExclusionAddDataGrid.ItemsSource = _exclusionAddItems
        ExclusionRevokeDataGrid.ItemsSource = _exclusionRevokeItems
        CalibAuditDataGrid.ItemsSource = _calibAuditItems

        LoadAuditData()

        AddHandler _controlWiseItems.CollectionChanged, AddressOf OnControlItemsChanged
    End Sub

    Private Sub OnControlItemsChanged(sender As Object, e As Collections.Specialized.NotifyCollectionChangedEventArgs)
        If e.NewItems IsNot Nothing Then
            For Each item As SearchItem In e.NewItems
                AddHandler item.PropertyChanged, AddressOf OnSearchItemPropertyChanged
            Next
        End If
        ' Update remove button status for any list change
        UpdateRemoveButtonStatus()
    End Sub

    Private Sub OnSearchItemPropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs)
        If e.PropertyName = "IsSelected" Then
            UpdateRemoveButtonStatus()
        End If
    End Sub

    Private Sub UpdateRemoveButtonStatus()
        If BtnRemoveControl IsNot Nothing Then
            BtnRemoveControl.IsEnabled = _controlWiseItems.Any(Function(x) x.IsSelected)
        End If
    End Sub

    Private Sub LoadCycles()
        Try
            ' Get unique cycles from regular_calibration
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT CycleName FROM regular_calibration WHERE CycleName IS NOT NULL AND CycleName != '' ORDER BY id DESC")
            ComboCycle.Items.Clear()
            ComboCycle.Items.Add("All")

            Dim activeCycle = GetCurrentCycleName()

            For Each row As DataRow In dt.Rows
                Dim cName = row("CycleName").ToString()
                ComboCycle.Items.Add(cName)
            Next

            ' Select "All" by default as per user request
            ComboCycle.SelectedIndex = 0
        Catch ex As Exception
            MessageBox.Show("Error loading cycles: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Function GetCurrentCycleName() As String
        ' Delegate to MySQLClass which reads Admin-configured date ranges + ActiveCycle override
        Return _mySql.GetActiveCycleName()
    End Function


    Private Sub FilterComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isLoading Then Return
        LoadAuditData()
    End Sub

    Private Sub PopulateYearFilter()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT YEAR(ActionDate) as YearVal FROM interchangeability WHERE ActionDate IS NOT NULL ORDER BY YearVal DESC")
            ComboYear.Items.Clear()
            ComboYear.Items.Add("All")
            For Each row As DataRow In dt.Rows
                ComboYear.Items.Add(row("YearVal").ToString())
            Next
            ComboYear.SelectedIndex = 0
        Catch ex As Exception
            Console.WriteLine("Error loading years: " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateColorFilter()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT Color FROM interchangeability WHERE Color IS NOT NULL AND Color != '' ORDER BY Color")
            ComboColor.Items.Clear()
            ComboColor.Items.Add("All")
            For Each row As DataRow In dt.Rows
                ComboColor.Items.Add(row("Color").ToString())
            Next
            ComboColor.SelectedIndex = 0
        Catch ex As Exception
            Console.WriteLine("Error loading colors: " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateStatusFilter()
        Try
            ComboCalibrationStatus.Items.Clear()
            ComboCalibrationStatus.Items.Add("All")
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT calibration_status FROM regular_calibration WHERE calibration_status IS NOT NULL AND calibration_status != '' ORDER BY calibration_status")
            For Each row As DataRow In dt.Rows
                ComboCalibrationStatus.Items.Add(row("calibration_status").ToString())
            Next
            ComboCalibrationStatus.SelectedIndex = 0
        Catch ex As Exception
            Console.WriteLine("Error loading statuses: " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateInterchangeStatusFilter()
        Try
            ComboStatus.Items.Clear()
            ComboStatus.Items.Add("All")
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT Status FROM interchangeability WHERE Status IS NOT NULL AND Status != '' ORDER BY Status")
            For Each row As DataRow In dt.Rows
                ComboStatus.Items.Add(row("Status").ToString())
            Next
            ComboStatus.SelectedIndex = 0
        Catch ex As Exception
            Console.WriteLine("Error loading interchange statuses: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadDepartments()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DepartmentName FROM departmentmaster ORDER BY DepartmentName")
            Dim deptList As New List(Of String)()
            deptList.Add("All")
            For Each row As DataRow In dt.Rows
                deptList.Add(row("DepartmentName").ToString())
            Next
            DeptFiltersItemsControl.ItemsSource = deptList
        Catch ex As Exception
            Console.WriteLine("Error loading departments: " & ex.Message)
        End Try
    End Sub

    Private Sub TxtSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            LoadAuditData()
        End If
    End Sub

    Private Sub LoadAuditData_Click(sender As Object, e As RoutedEventArgs)
        LoadAuditData()
    End Sub

    Private Sub LoadAuditData()
        If ComboCycle.SelectedItem Is Nothing Then Return
        Dim selectedCycle = ComboCycle.SelectedItem.ToString()
        Dim selectedYear = If(ComboYear.SelectedItem IsNot Nothing, ComboYear.SelectedItem.ToString(), "All")
        Dim selectedColor = If(ComboColor.SelectedItem IsNot Nothing, ComboColor.SelectedItem.ToString(), "All")
        Dim selectedCalibStatus = If(ComboCalibrationStatus.SelectedItem IsNot Nothing, ComboCalibrationStatus.SelectedItem.ToString(), "All")
        Dim selectedStatus = If(ComboStatus.SelectedItem IsNot Nothing, ComboStatus.SelectedItem.ToString(), "All")

        Try
            Dim searchText = TxtSearch.Text.Trim()

            ' Build the query joining interchangeability and regular_calibration
            ' Only show instruments that are currently 'Received'
            Dim query = "SELECT COALESCE(r.instrument_name, i.InstrumentName) as instrument_name, " &
                        "i.ControlNo as control_no, i.CycleName as cycle_name, i.Status as status, i.Department, i.ActionDate, i.ActionTime, i.Color, " &
                        "COALESCE(r.is_calibrated, 'NO') as is_calibrated, r.calibrated_date, r.due_date, r.calibration_status " &
                        "FROM interchangeability i " &
                        "LEFT JOIN regular_calibration r ON i.ControlNo = r.control_no AND i.CycleName = r.CycleName " &
                        "WHERE 1=1"

            If selectedCycle <> "All" Then
                query &= $" AND i.CycleName = '{selectedCycle.Replace("'", "''")}'"
            End If

            If selectedYear <> "All" Then
                query &= $" AND YEAR(i.ActionDate) = {selectedYear}"
            End If

            If selectedColor <> "All" Then
                query &= $" AND i.Color = '{selectedColor.Replace("'", "''")}'"
            End If

            If selectedCalibStatus <> "All" Then
                query &= $" AND r.calibration_status = '{selectedCalibStatus.Replace("'", "''")}'"
            End If

            If selectedStatus <> "All" Then
                query &= $" AND i.Status = '{selectedStatus.Replace("'", "''")}'"
            End If

            If DpFrom.SelectedDate IsNot Nothing Then
                query &= $" AND i.ActionDate >= '{DpFrom.SelectedDate.Value.ToString("yyyy-MM-dd")}'"
            End If

            If DpTo.SelectedDate IsNot Nothing Then
                query &= $" AND i.ActionDate <= '{DpTo.SelectedDate.Value.ToString("yyyy-MM-dd")}'"
            End If

            If _selectedDept <> "All" Then
                query &= $" AND i.Department = '{_selectedDept.Replace("'", "''")}'"
            End If

            If Not String.IsNullOrEmpty(searchText) Then
                query &= $" AND (i.ControlNo LIKE '%{searchText.Replace("'", "''")}%' OR COALESCE(r.instrument_name, i.InstrumentName) LIKE '%{searchText.Replace("'", "''")}%')"
            End If

            query &= " ORDER BY instrument_name ASC"

            Dim dt = _mySql.ReadDatatable(query)
            AuditDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading audit data: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub TabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton As Button = DirectCast(sender, Button)

        ' Update Tab Buttons
        BtnAudit.Tag = If(clickedButton Is BtnAudit, "Selected", "")
        BtnBulk.Tag = If(clickedButton Is BtnBulk, "Selected", "")
        BtnExclusion.Tag = If(clickedButton Is BtnExclusion, "Selected", "")

        ' Switch Views
        AuditView.Visibility = If(clickedButton Is BtnAudit, Visibility.Visible, Visibility.Collapsed)
        BulkView.Visibility = If(clickedButton Is BtnBulk, Visibility.Visible, Visibility.Collapsed)
        ExclusionView.Visibility = If(clickedButton Is BtnExclusion, Visibility.Visible, Visibility.Collapsed)

        ' If switching to Bulk Search or Exclusion List, ensure default sub-tabs are displayed
        If clickedButton Is BtnBulk Then
            ' If no sub-tab is selected yet (initial state), select Bulk Scan
            If SubBulkScanView.Visibility = Visibility.Collapsed AndAlso
               SubControlView.Visibility = Visibility.Collapsed AndAlso
               SubDeptView.Visibility = Visibility.Collapsed AndAlso
               SubDatabaseView.Visibility = Visibility.Collapsed Then
                SubBulkScanView.Visibility = Visibility.Visible
                BtnSubBulkScan.Tag = "Selected"
            End If
        ElseIf clickedButton Is BtnExclusion Then
            ' If no sub-tab is selected yet (initial state), select Exclusion Addition
            If SubExclusionAddView.Visibility = Visibility.Collapsed AndAlso
               SubExclusionRevokeView.Visibility = Visibility.Collapsed Then
                SubExclusionAddView.Visibility = Visibility.Visible
                BtnSubExclusionAdd.Tag = "Selected"
            End If
        End If

        StopAllScanning()
    End Sub

    Private Sub SubTabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton As Button = DirectCast(sender, Button)

        ' Update Sub-tab Buttons
        BtnSubBulkScan.Tag = If(clickedButton Is BtnSubBulkScan, "Selected", "")
        BtnSubControl.Tag = If(clickedButton Is BtnSubControl, "Selected", "")
        BtnSubDept.Tag = If(clickedButton Is BtnSubDept, "Selected", "")
        BtnSubDatabase.Tag = If(clickedButton Is BtnSubDatabase, "Selected", "")
        BtnSubCalibAudit.Tag = If(clickedButton Is BtnSubCalibAudit, "Selected", "")
        If BtnSubExclusionAdd IsNot Nothing Then BtnSubExclusionAdd.Tag = If(clickedButton Is BtnSubExclusionAdd, "Selected", "")
        If BtnSubExclusionRevoke IsNot Nothing Then BtnSubExclusionRevoke.Tag = If(clickedButton Is BtnSubExclusionRevoke, "Selected", "")

        ' Switch Sub-views
        SubBulkScanView.Visibility = If(clickedButton Is BtnSubBulkScan, Visibility.Visible, Visibility.Collapsed)
        SubControlView.Visibility = If(clickedButton Is BtnSubControl, Visibility.Visible, Visibility.Collapsed)
        SubDeptView.Visibility = If(clickedButton Is BtnSubDept, Visibility.Visible, Visibility.Collapsed)
        SubDatabaseView.Visibility = If(clickedButton Is BtnSubDatabase, Visibility.Visible, Visibility.Collapsed)
        SubCalibAuditView.Visibility = If(clickedButton Is BtnSubCalibAudit, Visibility.Visible, Visibility.Collapsed)
        If SubExclusionAddView IsNot Nothing Then SubExclusionAddView.Visibility = If(clickedButton Is BtnSubExclusionAdd, Visibility.Visible, Visibility.Collapsed)
        If SubExclusionRevokeView IsNot Nothing Then
            SubExclusionRevokeView.Visibility = If(clickedButton Is BtnSubExclusionRevoke, Visibility.Visible, Visibility.Collapsed)
            If clickedButton Is BtnSubExclusionRevoke Then
                LoadRevokeExclusionList()
            End If
        End If

        StopAllScanning()
    End Sub

    Private Sub BtnScanning_Click(sender As Object, e As RoutedEventArgs)
        Dim scanButton As Button = TryCast(sender, Button)
        If scanButton Is Nothing Then Return

        Dim isGunScan As Boolean = True
        If scanButton Is BtnBulkScan Then
            isGunScan = RbBulkScanGun.IsChecked.GetValueOrDefault()
        ElseIf scanButton Is BtnControlScan Then
            ' Mandatory check for Control Wise tab
            If SubControlView.Visibility = Visibility.Visible AndAlso _controlWiseItems.Count = 0 Then
                MessageBox.Show("Please add at least one Control Number before scanning.", "List Empty", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
            isGunScan = RbControlGun.IsChecked.GetValueOrDefault()
        ElseIf scanButton Is BtnDeptScan Then
            isGunScan = RbDeptGun.IsChecked.GetValueOrDefault()
        ElseIf scanButton Is BtnDatabaseScan Then
            ' Mandatory check for Instrument/Gauge Wise tab
            If SubDatabaseView.Visibility = Visibility.Visible AndAlso _selectedBulkTypes.Count = 0 Then
                MessageBox.Show("Please select at least one instrument/gauge type before scanning.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
            isGunScan = RbDatabaseGun.IsChecked.GetValueOrDefault()
        ElseIf scanButton Is BtnExclusionScan Then
            isGunScan = RbExclusionGun.IsChecked.GetValueOrDefault()
        ElseIf scanButton Is BtnCalibAuditScan Then
            isGunScan = RbCalibAuditGun.IsChecked.GetValueOrDefault()
        End If

        If _isBulkScanning Then
            StopScanning(scanButton)
        Else
            ' Mandatory check for Department Wise tab
            If SubDeptView.Visibility = Visibility.Visible AndAlso _selectedBulkDepts.Count = 0 Then
                MessageBox.Show("Please select at least one department before scanning.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            StartScanning(scanButton, isGunScan)
        End If
    End Sub

    Private Sub StartScanning(btn As Button, isGunScan As Boolean)
        _isBulkScanning = True
        btn.Content = "Stop Scan"
        btn.Background = New Media.SolidColorBrush(Media.Color.FromRgb(220, 38, 38)) ' Red

        _bulkScanItems.Clear()
        _calibAuditItems.Clear()
        _calibratedCount = 0
        _notCalibratedCount = 0
        TxtCalibratedCount.Text = "0"
        TxtNotCalibratedCount.Text = "0"

        ' Note: We don't clear _controlWiseItems or _deptWiseItems because they are pre-added/loaded.
        ' However, we should reset their "Found" status when starting a fresh scan.
        If True Then ' Apply to both for safety
            For Each item In _controlWiseItems
                item.Found = "No"
            Next
            For Each item In _deptWiseItems
                item.Found = "No"
            Next
        End If

        If isGunScan Then
            Dim ip = "192.168.1.51"
            Dim port = "3000"
            If ProjectSettings.Current IsNot Nothing Then
                If Not String.IsNullOrEmpty(ProjectSettings.Current.RfidGunIp) Then ip = ProjectSettings.Current.RfidGunIp
                If Not String.IsNullOrEmpty(ProjectSettings.Current.RfidGunPort) Then port = ProjectSettings.Current.RfidGunPort
            End If

            _gunClient = New RFIDGunClient()
            AddHandler RFIDGunClient.RFIDScanned, AddressOf OnTagScanned
            Dim discardTask = _gunClient.StartStreamingAsync(ip, port)
        Else
            _usbScanCts = New System.Threading.CancellationTokenSource()
            Dim token = _usbScanCts.Token
            Task.Run(Sub()
                         While Not token.IsCancellationRequested
                             Dim tags = CommManager.ScanBulkTags(Nothing, 400)
                             If tags IsNot Nothing AndAlso tags.Count > 0 Then
                                 For Each t In tags
                                     OnTagScanned(t)
                                 Next
                             End If
                             System.Threading.Thread.Sleep(50)
                         End While
                     End Sub, token)
        End If
    End Sub

    Private Sub StopScanning(btn As Button)
        StopAllScanning()
    End Sub

    Private Sub StopAllScanning()
        If Not _isBulkScanning Then Return

        _isBulkScanning = False

        ' Reset UI for all possible scan buttons
        Dim scanBtns = {BtnBulkScan, BtnControlScan, BtnDeptScan, BtnDatabaseScan, BtnExclusionScan, BtnCalibAuditScan}
        For Each btn In scanBtns
            If btn IsNot Nothing Then
                btn.Content = "Scan"
                btn.ClearValue(Button.BackgroundProperty)
            End If
        Next

        ' Stop Logic
        If _gunClient IsNot Nothing Then
            RemoveHandler RFIDGunClient.RFIDScanned, AddressOf OnTagScanned
            _gunClient.StopStreaming()
            _gunClient = Nothing
        End If

        If _usbScanCts IsNot Nothing Then
            _usbScanCts.Cancel()
            _usbScanCts = Nothing
        End If
    End Sub

    Private Sub OnTagScanned(rfid As String)
        Dispatcher.Invoke(Sub()
                              Dim formattedRfid = FormatRfidWithSpacesIfNeeded(rfid)
                              Dim cleanTag = rfid.Replace(" ", "").ToUpper().Replace("'", "''")

                              ' 1. Handle Exclusion List Addition tab FIRST
                              If ExclusionView.Visibility = Visibility.Visible AndAlso SubExclusionAddView.Visibility = Visibility.Visible Then
                                  ' Check if already in exclusion database
                                  If _mySql.IsTagExcluded(rfid) Then Return

                                  ' Check if already in UI list
                                  If _exclusionAddItems.Any(Function(x) x.RfidTag.Replace(" ", "").Equals(rfid.Replace(" ", ""), StringComparison.OrdinalIgnoreCase)) Then Return

                                  ' Fetch info from interchangeability (any cycle)
                                  Dim dtInfo = _mySql.GetInterchangeInfoByTag(rfid)
                                  Dim ctrlNoEx = "N/A"
                                  Dim instrNameEx = "Unknown Tag"
                                  If dtInfo.Rows.Count > 0 Then
                                      ctrlNoEx = dtInfo.Rows(0)("ControlNo").ToString()
                                      instrNameEx = dtInfo.Rows(0)("InstrumentName").ToString()
                                  End If

                                  _exclusionAddItems.Add(New ExclusionItem With {
                                      .RfidTag = formattedRfid,
                                      .IsSelected = True,
                                      .ControlNo = ctrlNoEx,
                                      .InstrumentName = instrNameEx,
                                      .IsInInterchange = (dtInfo.Rows.Count > 0)
                                  })
                                  Return
                              End If

                              ' 2. Existing audit/search logic
                              Dim query = "SELECT i.ControlNo, i.InstrumentName, i.CycleName, i.Department, i.Color, i.Status, i.ActionDate, i.ActionTime, " &
                        "COALESCE(r.is_calibrated, 'NO') as is_calibrated, r.calibrated_date, r.calibration_status " &
                        "FROM interchangeability i " &
                        "LEFT JOIN regular_calibration r ON i.ControlNo = r.control_no AND i.CycleName = r.CycleName " &
                        $"WHERE UPPER(REPLACE(i.RFID_tag, ' ', '')) = '{cleanTag}' LIMIT 1"
                              Dim dt = _mySql.ReadDatatable(query)

                              If dt.Rows.Count > 0 Then
                                  Dim row = dt.Rows(0)
                                  Dim ctrlNo = row("ControlNo").ToString().Trim()
                                  Dim instrName = row("InstrumentName").ToString()
                                  Dim dept = row("Department").ToString()
                                  Dim color = row("Color").ToString()
                                  Dim cycle = row("CycleName").ToString()
                                  Dim status = row("Status").ToString()
                                  Dim aDate = If(row("ActionDate") Is DBNull.Value, "", Convert.ToDateTime(row("ActionDate")).ToString("dd/MM/yy"))
                                  Dim aTime = row("ActionTime").ToString()
                                  Dim isCalib = row("is_calibrated").ToString()
                                  Dim calibDate = If(row("calibrated_date") Is DBNull.Value, "", Convert.ToDateTime(row("calibrated_date")).ToString("dd/MM/yyyy"))
                                  Dim calibStatus = row("calibration_status").ToString()

                                  ' Handle Control Number Wise tab
                                  If SubControlView.Visibility = Visibility.Visible Then
                                      Dim existingItem = _controlWiseItems.FirstOrDefault(Function(x) x.ControlNo.Trim().Equals(ctrlNo, StringComparison.OrdinalIgnoreCase))
                                      If existingItem IsNot Nothing Then
                                          existingItem.Found = "Yes"
                                      Else
                                          _controlWiseItems.Add(New SearchItem With {
                            .RfidTag = formattedRfid,
                            .ControlNo = ctrlNo,
                            .InstrumentName = instrName,
                            .CycleName = cycle,
                            .Department = dept,
                            .Color = color,
                            .Status = status,
                            .ActionDate = aDate,
                            .ActionTime = aTime,
                            .IsCalibrated = isCalib,
                            .CalibratedDate = calibDate,
                            .CalibrationStatus = calibStatus,
                            .Found = "Yes"
                        })
                                      End If
                                      Return
                                  End If

                                  ' Handle Department Wise tab
                                  If SubDeptView.Visibility = Visibility.Visible Then
                                      ' Only process if the item's department is in the selected list
                                      If _selectedBulkDepts.Contains(dept) Then
                                          Dim existingItem = _deptWiseItems.FirstOrDefault(Function(x) x.ControlNo.Trim().Equals(ctrlNo, StringComparison.OrdinalIgnoreCase))
                                          If existingItem Is Nothing Then
                                              _deptWiseItems.Add(New SearchItem With {
                                .RfidTag = formattedRfid,
                                .ControlNo = ctrlNo,
                                .InstrumentName = instrName,
                                .CycleName = cycle,
                                .Department = dept,
                                .Color = color,
                                .Status = status,
                                .ActionDate = aDate,
                                .ActionTime = aTime,
                                .IsCalibrated = isCalib,
                                .CalibratedDate = calibDate,
                                .CalibrationStatus = calibStatus
                            })
                                          End If
                                      End If
                                      Return
                                  End If

                                  ' Handle Database Wise tab
                                  If SubDatabaseView.Visibility = Visibility.Visible Then
                                      ' Only process if the item's instrument name (type) is in the selected list
                                      If _selectedBulkTypes.Contains(instrName) Then
                                          Dim existingItem = _databaseWiseItems.FirstOrDefault(Function(x) x.ControlNo.Trim().Equals(ctrlNo, StringComparison.OrdinalIgnoreCase))
                                          If existingItem Is Nothing Then
                                              _databaseWiseItems.Add(New SearchItem With {
                                .RfidTag = formattedRfid,
                                .ControlNo = ctrlNo,
                                .InstrumentName = instrName,
                                .CycleName = cycle,
                                .Department = dept,
                                .Color = color,
                                .Status = status,
                                .ActionDate = aDate,
                                .ActionTime = aTime,
                                .IsCalibrated = isCalib,
                                .CalibratedDate = calibDate,
                                .CalibrationStatus = calibStatus
                            })
                                          End If
                                      End If
                                      Return
                                  End If

                                  ' Handle Calibration Audit tab
                                  If SubCalibAuditView.Visibility = Visibility.Visible Then
                                      Dim existingItem = _calibAuditItems.FirstOrDefault(Function(x) x.ControlNo.Trim().Equals(ctrlNo, StringComparison.OrdinalIgnoreCase))
                                      If existingItem Is Nothing Then
                                          _calibAuditItems.Add(New SearchItem With {
                                .RfidTag = formattedRfid,
                                .ControlNo = ctrlNo,
                                .InstrumentName = instrName,
                                .CycleName = cycle,
                                .Department = dept,
                                .Color = color,
                                .Status = status,
                                .ActionDate = aDate,
                                .ActionTime = aTime,
                                .IsCalibrated = isCalib,
                                .CalibratedDate = calibDate,
                                .CalibrationStatus = calibStatus
                            })
                                          ' Update Counts
                                          If isCalib.ToUpper() = "YES" Then
                                              _calibratedCount += 1
                                          Else
                                              _notCalibratedCount += 1
                                          End If
                                          TxtCalibratedCount.Text = _calibratedCount.ToString()
                                          TxtNotCalibratedCount.Text = _notCalibratedCount.ToString()
                                      End If
                                      Return
                                  End If

                                  ' Duplicate check for Bulk Scan
                                  For Each item In _bulkScanItems
                                      If item.ControlNo.Trim().Equals(ctrlNo, StringComparison.OrdinalIgnoreCase) Then Return
                                  Next

                                  Dim newItem As New SearchItem With {
                    .RfidTag = formattedRfid,
                    .ControlNo = ctrlNo,
                    .InstrumentName = instrName,
                    .CycleName = cycle,
                    .Department = dept,
                    .Color = color,
                    .Status = status,
                    .ActionDate = aDate,
                    .ActionTime = aTime,
                    .IsCalibrated = isCalib,
                    .CalibratedDate = calibDate,
                    .CalibrationStatus = calibStatus
                }
                                  _bulkScanItems.Add(newItem)
                              Else
                                  ' Tag not in interchangeability
                                  ' Note: Exclusion tab already handled at top
                              End If
                          End Sub)
    End Sub

    Private Function FormatRfidWithSpacesIfNeeded(rfid As String) As String
        Dim clean = rfid.Replace(" ", "").Trim()
        If clean.Length Mod 2 <> 0 OrElse clean.Length = 0 Then Return rfid
        Dim parts As New List(Of String)
        For i As Integer = 0 To clean.Length - 1 Step 2
            parts.Add(clean.Substring(i, 2))
        Next
        Return String.Join(" ", parts)
    End Function

    Private Sub LoadRevokeExclusionList()
        _exclusionRevokeItems.Clear()
        Dim dt = _mySql.GetExclusionList()
        For Each row As DataRow In dt.Rows
            _exclusionRevokeItems.Add(New ExclusionItem With {
                .RfidTag = row("rfid_tag").ToString(),
                .IsSelected = False
            })
        Next
    End Sub

    Private Sub BtnClearCalibAudit_Click(sender As Object, e As RoutedEventArgs)
        _calibAuditItems.Clear()
        _calibratedCount = 0
        _notCalibratedCount = 0
        TxtCalibratedCount.Text = "0"
        TxtNotCalibratedCount.Text = "0"
    End Sub

    Private Sub BtnRevokeExclusion_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedItems = _exclusionRevokeItems.Where(Function(x) x.IsSelected).ToList()
        If selectedItems.Count = 0 Then
            MessageBox.Show("Please select at least one tag to revoke.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim res = MessageBox.Show($"Are you sure you want to revoke {selectedItems.Count} tag(s)?", "Confirm Revocation", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If res = MessageBoxResult.Yes Then
            Dim successCount = 0
            For Each item In selectedItems
                If _mySql.DeleteExclusion(item.RfidTag) Then
                    successCount += 1
                End If
            Next

            If successCount > 0 Then
                MessageBox.Show($"{successCount} tag(s) revoked successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                LoadRevokeExclusionList()
            Else
                MessageBox.Show("Failed to revoke selected tags.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub


    Private Sub BtnAddExclusion_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedItems = _exclusionAddItems.Where(Function(x) x.IsSelected).ToList()
        If selectedItems.Count = 0 Then
            MessageBox.Show("Please select at least one tag to exclude.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim confirmWin As New ExclusionConfirmWindow(selectedItems)
        Dim result = confirmWin.ShowDialog()
        If result <> True Then Return

        ' Warning logic for interchangeability
        Dim warnings As New List(Of String)()
        For Each item In selectedItems
            Dim dtInterchange = _mySql.GetInterchangeInfoByTag(item.RfidTag)
            If dtInterchange.Rows.Count > 0 Then
                Dim ctrlNo = dtInterchange.Rows(0)("ControlNo").ToString()
                Dim instrName = dtInterchange.Rows(0)("InstrumentName").ToString()
                warnings.Add($"Tag: {item.RfidTag}, Control: {ctrlNo}, Name: {instrName}")
            End If
        Next

        If warnings.Count > 0 Then
            Dim warnMsg = "Warning: The following tags already exist in the interchangeability table. Do you still want to exclude them?" & vbCrLf & vbCrLf & String.Join(vbCrLf, warnings)
            Dim warnResult = MessageBox.Show(warnMsg, "Exclusion Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning)
            If warnResult <> MessageBoxResult.Yes Then Return
        End If

        ' Insertion logic
        Dim successCount = 0
        For Each item In selectedItems
            If _mySql.InsertExclusion(item.RfidTag) Then
                successCount += 1
            End If
        Next

        If successCount > 0 Then
            MessageBox.Show($"{successCount} tags added to exclusion list.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            _exclusionAddItems.Clear()
        End If
    End Sub

    Private Sub BtnAddControl_Click(sender As Object, e As RoutedEventArgs)
        Dim ctrlNo = TxtBulkControlNo.Text.Trim()
        If String.IsNullOrEmpty(ctrlNo) Then
            MessageBox.Show("Please enter a Control No.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' 1. Check duplicate in current list
        If _controlWiseItems.Any(Function(x) x.ControlNo.Equals(ctrlNo, StringComparison.OrdinalIgnoreCase)) Then
            MessageBox.Show("Control Number already in the list.", "Duplicate", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        ' 2. Search Master Tables
        Dim foundItem = SearchMasterTables(ctrlNo)
        If foundItem IsNot Nothing Then
            _controlWiseItems.Add(foundItem)
            TxtBulkControlNo.Clear()
        Else
            MessageBox.Show("Control Number not found in Master List.", "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning)
        End If
    End Sub

    Private Sub BtnClearControl_Click(sender As Object, e As RoutedEventArgs)
        If _controlWiseItems.Count > 0 Then
            _controlWiseItems.Clear()
        End If
    End Sub

    Private Sub BtnClearBulkScan_Click(sender As Object, e As RoutedEventArgs)
        If _bulkScanItems.Count > 0 Then
            _bulkScanItems.Clear()
        End If
    End Sub

    Private Sub BtnClearDeptWise_Click(sender As Object, e As RoutedEventArgs)
        If _deptWiseItems.Count > 0 Then
            _deptWiseItems.Clear()
        End If
    End Sub

    Private Sub BtnClearDatabaseWise_Click(sender As Object, e As RoutedEventArgs)
        If _databaseWiseItems.Count > 0 Then
            _databaseWiseItems.Clear()
        End If
    End Sub

    Private Sub BtnClearExclusionAdd_Click(sender As Object, e As RoutedEventArgs)
        If _exclusionAddItems.Count > 0 Then
            _exclusionAddItems.Clear()
        End If
    End Sub

    Private Sub BtnRemoveControl_Click(sender As Object, e As RoutedEventArgs)
        Dim toRemove = _controlWiseItems.Where(Function(x) x.IsSelected).ToList()
        For Each item In toRemove
            _controlWiseItems.Remove(item)
        Next
    End Sub

    Private Function SearchMasterTables(ctrlNo As String) As SearchItem
        Try
            Dim activeCycle = GetCurrentCycleName()
            Dim query = "SELECT i.ControlNo, i.InstrumentName, i.CycleName, i.Department, i.Color, i.Status, i.ActionDate, i.ActionTime, i.RFID_tag, " &
                        "COALESCE(r.is_calibrated, 'NO') as is_calibrated, r.calibrated_date, r.calibration_status " &
                        "FROM interchangeability i " &
                        "LEFT JOIN regular_calibration r ON i.ControlNo = r.control_no AND i.CycleName = r.CycleName " &
                        $"WHERE i.ControlNo = '{ctrlNo.Replace("'", "''")}' AND i.CycleName = '{activeCycle.Replace("'", "''")}' LIMIT 1"

            Dim dt = _mySql.ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                Dim item As New SearchItem()
                item.ControlNo = ctrlNo
                item.InstrumentName = row("InstrumentName").ToString()
                item.CycleName = row("CycleName").ToString()
                item.Department = row("Department").ToString()
                item.Color = row("Color").ToString()
                item.Status = row("Status").ToString()
                item.ActionDate = If(row("ActionDate") Is DBNull.Value, "", Convert.ToDateTime(row("ActionDate")).ToString("dd/MM/yy"))
                item.ActionTime = row("ActionTime").ToString()
                item.RfidTag = FormatRfidWithSpacesIfNeeded(row("RFID_tag").ToString())
                item.IsCalibrated = row("is_calibrated").ToString()
                item.CalibratedDate = If(row("calibrated_date") Is DBNull.Value, "", Convert.ToDateTime(row("calibrated_date")).ToString("dd/MM/yyyy"))
                item.CalibrationStatus = row("calibration_status").ToString()
                item.Found = "No"
                Return item
            End If
        Catch ex As Exception
            Console.WriteLine("Error searching master tables: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Private Sub BtnSelectDept_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DepartmentName FROM departmentmaster ORDER BY DepartmentName")
            Dim allDepts As New List(Of String)()
            For Each row As DataRow In dt.Rows
                allDepts.Add(row("DepartmentName").ToString())
            Next

            Dim selectWin As New DepartmentSelectionWindow(allDepts, _selectedBulkDepts)
            selectWin.Owner = Window.GetWindow(Me)
            If selectWin.ShowDialog() = True Then
                _selectedBulkDepts = selectWin.SelectedDepartments
                _deptWiseItems.Clear() ' Clear previous scans when selection changes
                If _selectedBulkDepts.Count > 0 Then
                    TxtSelectedDepts.Text = String.Join(", ", _selectedBulkDepts)
                Else
                    TxtSelectedDepts.Text = "None Selected"
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error selecting departments: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub LoadDeptWiseData()
        ' This is now just a placeholder or helper to clear data
        _deptWiseItems.Clear()
    End Sub

    Private Sub BtnToggleDeptFilter_Click(sender As Object, e As RoutedEventArgs)
        If DeptFilterContainer.Visibility = Visibility.Collapsed Then
            DeptFilterContainer.Visibility = Visibility.Visible
            TxtToggleDeptFilter.Text = "Hide Dept Filter"
        Else
            DeptFilterContainer.Visibility = Visibility.Collapsed
            TxtToggleDeptFilter.Text = "Department Filter"
        End If
    End Sub

    Private Sub BtnToggleMoreFilters_Click(sender As Object, e As RoutedEventArgs)
        If MoreFiltersContainer.Visibility = Visibility.Collapsed Then
            MoreFiltersContainer.Visibility = Visibility.Visible
        Else
            MoreFiltersContainer.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub FilterDatePicker_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isLoading Then Return
        LoadAuditData()
    End Sub

    Private Sub BtnClearDates_Click(sender As Object, e As RoutedEventArgs)
        _isLoading = True
        DpFrom.SelectedDate = Nothing
        DpTo.SelectedDate = Nothing
        _isLoading = False
        LoadAuditData()
    End Sub

    Private Sub DeptFilterButton_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        _selectedDept = btn.Content.ToString()

        ' Update UI tags for visual feedback
        For Each item In DeptFiltersItemsControl.ItemsSource
            Dim container = DeptFiltersItemsControl.ItemContainerGenerator.ContainerFromItem(item)
            If container IsNot Nothing Then
                Dim childBtn = FindVisualChild(Of Button)(container)
                If childBtn IsNot Nothing Then
                    childBtn.Tag = If(item.ToString() = _selectedDept, "Selected", "")
                End If
            End If
        Next

        LoadAuditData()
    End Sub

    Private Sub ControlNo_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim row = TryCast(btn.DataContext, DataRowView)
        If row Is Nothing Then Return

        Dim ctrl = row("control_no").ToString()
        If String.IsNullOrEmpty(ctrl) Then Return

        Try
            ' Find table name and item type by checking master list tables
            Dim tblName = "unknown"
            Dim itemType = ""
            Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName, Category FROM type_details")

            For Each r As DataRow In dtSettings.Rows
                Dim typeName = r("TypeName").ToString()
                Dim testTbl = MySQLClass.TypeNameToTableName(typeName, forInventory:=True)
                Dim chk = _mySql.ReadDatatable($"SELECT ControlNo FROM `{testTbl}` WHERE ControlNo='{ctrl.Replace("'", "''")}'")
                If chk.Rows.Count > 0 Then
                    tblName = testTbl
                    itemType = r("Category").ToString()
                    Exit For
                End If
            Next

            If tblName = "unknown" Then
                MessageBox.Show("Could not find master list entry for " & ctrl, "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            ' Open ItemDetailsWindow
            Dim detailsWin As New ItemDetailsWindow()
            detailsWin.WindowStartupLocation = WindowStartupLocation.CenterScreen
            detailsWin.LoadDetails(ctrl, itemType, tblName)
            detailsWin.ShowDialog()
        Catch ex As Exception
            MessageBox.Show("Error opening item details: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnSelectType_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim dt = _mySql.ReadDatatable("SELECT TypeName, Category FROM type_details ORDER BY TypeName")

            Dim win As New TypeSelectionWindow(dt, _selectedBulkTypes)
            win.Owner = Window.GetWindow(Me)
            If win.ShowDialog() = True Then
                _selectedBulkTypes = win.SelectedTypes
                _databaseWiseItems.Clear() ' Clear previous scans when selection changes
                If _selectedBulkTypes.Count > 0 Then
                    TxtSelectedTypes.Text = String.Join(", ", _selectedBulkTypes)
                Else
                    TxtSelectedTypes.Text = "None Selected"
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error selecting types: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub LoadTypeWiseData()
        ' Repurposed to just clear the list
        _databaseWiseItems.Clear()
    End Sub

    ' Helper to find child in template
    Private Function FindVisualChild(Of T As DependencyObject)(ByVal obj As DependencyObject) As T
        For i As Integer = 0 To Media.VisualTreeHelper.GetChildrenCount(obj) - 1
            Dim child = Media.VisualTreeHelper.GetChild(obj, i)
            If child IsNot Nothing AndAlso TypeOf child Is T Then
                Return DirectCast(child, T)
            Else
                Dim childOfChild = FindVisualChild(Of T)(child)
                If childOfChild IsNot Nothing Then Return childOfChild
            End If
        Next
        Return Nothing
    End Function
End Class
