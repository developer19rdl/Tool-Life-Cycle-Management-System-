Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Data
Imports Microsoft.Win32
Imports System.IO
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports ExcelDataReader

Class InterchangeablePage
    Private _mySql As New MySQLClass()
    Private _selectedDept As String = "All"
    Private _selectedInterDept As String = "All"
    Private _selectedHistoryDept As String = "All"
    Private _selectedDeptListCycle As String = ""
    Private _isLoading As Boolean = False
    Private _isBulkScanMode As Boolean = False
    Private _isUsbBulkScanMode As Boolean = False
    Private _usbScanCts As System.Threading.CancellationTokenSource
    Private _rfidGunClient As RFIDGunClient
    Private _issuedColl As New ObservableCollection(Of BulkScanSummaryWindow.IssuedItem)()
    Private _ignoredColl As New ObservableCollection(Of BulkScanSummaryWindow.IgnoredItem)()
    Private _processedTags As New HashSet(Of String)()
    Private _usbScanTask As Task
    Private _gunScanTask As Task

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Async Sub Page_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        MainProgressBar.Visibility = Visibility.Visible
        MainProgressBar.IsIndeterminate = True

        Try
            Dim activeCycle = GetCurrentCycleName()
            Await Task.Run(Sub()
                               _mySql.InitializeDepartmentListTable()
                               _mySql.EnsureAdminPassword()
                               ' _mySql.SyncDepartmentListToInterchange(activeCycle) ' Moved to explicit Import only
                               ' Cleanup any corrupted synchronization data (one-time check)
                               _mySql.ExecuteNonQuery("UPDATE department_list SET Department = 'Unknown' WHERE Department = 'System.Data.DataRowView'")
                               _mySql.ExecuteNonQuery("UPDATE interchangeability SET Department = 'Unknown' WHERE Department = 'System.Data.DataRowView'")
                           End Sub)

            LoadDepartments()
            LoadDeptListCycleDropdown(activeCycle)
            LoadData()
            InitializeInterchangeUI(activeCycle)
            _mySql.CleanupStaleRFIDs(activeCycle)
        Catch ex As Exception
            Console.WriteLine("Error in Page_Loaded: " & ex.Message)
        Finally
            AddHandler RFIDApiServer.RFIDScanned, AddressOf OnRFIDScanned
            AddHandler RFIDGunClient.RFIDScanned, AddressOf OnRFIDScanned
            MainProgressBar.Visibility = Visibility.Collapsed
            MainProgressBar.IsIndeterminate = False
        End Try
    End Sub

    Private Sub Page_Unloaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Unloaded
        RemoveHandler RFIDApiServer.RFIDScanned, AddressOf OnRFIDScanned
        RemoveHandler RFIDGunClient.RFIDScanned, AddressOf OnRFIDScanned
    End Sub

    Private Sub OnRFIDScanned(rfid As String)
        If _isBulkScanMode Then
            ProcessSingleTag(rfid)
        End If
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

    Private Sub LoadDeptListCycleDropdown(activeCycle As String)
        RemoveHandler ComboDeptListCycle.SelectionChanged, AddressOf ComboDeptListCycle_SelectionChanged
        ComboDeptListCycle.Items.Clear()

        Dim cycleList As New List(Of String)()
        If Not String.IsNullOrEmpty(activeCycle) Then cycleList.Add(activeCycle)

        ' Fetch historical cycles from DB
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT CycleName FROM interchangeability")
            For Each row As DataRow In dt.Rows
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
            ComboDeptListCycle.Items.Add(c)
        Next

        If ComboDeptListCycle.Items.Count > 0 Then
            ComboDeptListCycle.SelectedIndex = 0
            _selectedDeptListCycle = ComboDeptListCycle.Items(0).ToString()
        End If

        AddHandler ComboDeptListCycle.SelectionChanged, AddressOf ComboDeptListCycle_SelectionChanged
    End Sub

    Private Sub ComboDeptListCycle_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If ComboDeptListCycle.SelectedItem IsNot Nothing Then
            _selectedDeptListCycle = ComboDeptListCycle.SelectedItem.ToString()
        End If
        LoadData()
    End Sub

    Private Sub LoadData()
        If _isLoading Then Return
        _isLoading = True
        MainProgressBar.Visibility = Visibility.Visible

        Try
            Dim sortState = SaveSortState(DeptListDataGrid)
            Dim query As String = "SELECT * FROM department_list WHERE 1=1"

            ' Filter by selected cycle
            If Not String.IsNullOrEmpty(_selectedDeptListCycle) Then
                query &= $" AND CycleName = '{_selectedDeptListCycle.Replace("'", "''")}'"
            End If

            If _selectedDept <> "All" Then
                query &= $" AND Department = '{_selectedDept.Replace("'", "''")}'"
            End If
            query &= " ORDER BY id DESC"

            Dim dt = _mySql.ReadDatatable(query)
            DeptListDataGrid.ItemsSource = dt.DefaultView
            RestoreSortState(DeptListDataGrid, sortState)
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        Finally
            _isLoading = False
            MainProgressBar.Visibility = Visibility.Collapsed
        End Try
    End Sub

    Private Sub TabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton As Button = DirectCast(sender, Button)

        ' Reset all tab buttons
        BtnDeptList.Tag = ""
        BtnInterchange.Tag = ""
        BtnInterHistory.Tag = ""

        ' Set selected tab button
        clickedButton.Tag = "Selected"

        ' Hide all tab views
        DeptListTabView.Visibility = Visibility.Collapsed
        InterchangeTabView.Visibility = Visibility.Collapsed
        InterHistoryTabView.Visibility = Visibility.Collapsed

        ' Show corresponding tab view
        If clickedButton Is BtnDeptList Then
            DeptListTabView.Visibility = Visibility.Visible
            LoadDepartments()
            Dim activeCycle = GetCurrentCycleName()
            LoadDeptListCycleDropdown(activeCycle)
            LoadData()
        ElseIf clickedButton Is BtnInterchange Then
            InterchangeTabView.Visibility = Visibility.Visible
            Dim activeCycle = GetCurrentCycleName()
            ' _mySql.SyncDepartmentListToInterchange(activeCycle) ' Sync is now only explicit during Import
            LoadInterchangeDepartments()
            ResetSearchContext()
            InterchangeDataGrid.SelectedItem = Nothing
            LoadInterchangeData()
        ElseIf clickedButton Is BtnInterHistory Then
            InterHistoryTabView.Visibility = Visibility.Visible
            LoadInterHistoryDepartments()
            LoadHistoryData()
        End If

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

        LoadData()
    End Sub

    Private Async Sub BtnImport_Click(sender As Object, e As RoutedEventArgs)
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Excel Files (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv"
        If ofd.ShowDialog() = True Then
            Await ImportFromFileAsync(ofd.FileName)
        End If
    End Sub



    Private Async Function ImportFromFileAsync(filePath As String) As Task
        BtnImport.IsEnabled = False
        BtnImport.Content = "Importing..."
        Dim importedCount As Integer = 0
        Dim skippedCount As Integer = 0
        Dim batchTag As String = DateTime.Now.ToString("yyyyMMddHHmmss")

        Try
            MainProgressBar.Visibility = Visibility.Visible
            MainProgressBar.IsIndeterminate = False
            MainProgressBar.Value = 0

            Dim activeCycle = GetCurrentCycleName()

            Await Task.Run(Sub()
                               Try
                                   Using stream = File.Open(filePath, FileMode.Open, FileAccess.Read)
                                       Dim reader = If(filePath.ToLower().EndsWith(".csv"),
                                                     ExcelReaderFactory.CreateCsvReader(stream),
                                                     ExcelReaderFactory.CreateReader(stream))
                                       Using reader
                                           Dim result = reader.AsDataSet()
                                           Dim table = result.Tables(0)

                                           ' Find headers
                                           Dim idxDept = -1, idxName = -1, idxSize = -1, idxControl = -1, idxColor = -1, idxStatus = -1, idxRemarks = -1

                                           ' Search in first 15 rows for headers
                                           For r As Integer = 0 To Math.Min(15, table.Rows.Count - 1)
                                               For c As Integer = 0 To table.Columns.Count - 1
                                                   Dim colName As String = table.Rows(r)(c).ToString().Trim()
                                                   If IsMatch(colName, "Dept", "Department") Then idxDept = c
                                                   If IsMatch(colName, "Instrument Name", "Name", "InstrumentName") Then idxName = c
                                                   If IsMatch(colName, "Size & Range", "Size", "SizeandRange") Then idxSize = c
                                                   If IsMatch(colName, "Control No", "Control Number", "ControlNo") Then idxControl = c
                                                   If IsMatch(colName, "Color part", "Color", "Colour") Then idxColor = c
                                                   If IsMatch(colName, "Status") Then idxStatus = c
                                                   If IsMatch(colName, "Remark", "Remarks") Then idxRemarks = c
                                               Next
                                               ' Requirement: must have Control No or at least Name to proceed
                                               If idxControl <> -1 Or idxName <> -1 Then
                                                   ' Found headers at row r, data starts at r+1
                                                   Dim dataRows = table.Rows.Count - (r + 1)
                                                   If dataRows > 0 Then
                                                       Application.Current.Dispatcher.Invoke(Sub()
                                                                                                 MainProgressBar.Minimum = 0
                                                                                                 MainProgressBar.Maximum = dataRows
                                                                                             End Sub)
                                                   End If

                                                   Dim existingCtrls = _mySql.GetExistingControlNos()
                                                   Dim itemsToInsert As New List(Of Dictionary(Of String, Object))()

                                                   For i As Integer = r + 1 To table.Rows.Count - 1
                                                       Dim row = table.Rows(i)
                                                       Dim deptVal = GetVal(row, idxDept)
                                                       Dim nameVal = GetVal(row, idxName)
                                                       Dim sizeVal = GetVal(row, idxSize)
                                                       Dim ctrlVal = GetVal(row, idxControl)
                                                       Dim colorVal = GetVal(row, idxColor)
                                                       Dim statusVal = GetVal(row, idxStatus)
                                                       Dim remarksVal = GetVal(row, idxRemarks)

                                                       If Not String.IsNullOrWhiteSpace(nameVal) Or Not String.IsNullOrWhiteSpace(ctrlVal) Then
                                                           ' Duplicate check
                                                           If Not String.IsNullOrEmpty(ctrlVal) AndAlso existingCtrls.Contains(ctrlVal) Then
                                                               skippedCount += 1
                                                           Else
                                                               Dim item As New Dictionary(Of String, Object)()
                                                               item("dept") = deptVal
                                                               item("name") = nameVal
                                                               item("size") = sizeVal
                                                               item("ctrl") = ctrlVal
                                                               item("color") = colorVal
                                                               item("status") = If(String.IsNullOrEmpty(statusVal), "Pending", statusVal)
                                                               item("remarks") = remarksVal
                                                               itemsToInsert.Add(item)

                                                               If Not String.IsNullOrEmpty(ctrlVal) Then existingCtrls.Add(ctrlVal)
                                                           End If
                                                       End If

                                                       ' Update progress frequently
                                                       Dim currentProgress = i - r
                                                       If currentProgress Mod 10 = 0 Or i = table.Rows.Count - 1 Then
                                                           Application.Current.Dispatcher.Invoke(Sub() MainProgressBar.Value = currentProgress)
                                                       End If
                                                   Next

                                                   If itemsToInsert.Count > 0 Then
                                                       importedCount = _mySql.BulkInsertDeptListItems(itemsToInsert, batchTag, activeCycle)
                                                   End If
                                                   Exit For
                                               End If
                                           Next
                                       End Using
                                   End Using
                                   _mySql.SyncDepartmentListToInterchange(activeCycle)
                               Catch ex As Exception
                                   Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                               End Try
                           End Sub)

            MessageBox.Show($"Import Complete!" & vbCrLf & vbCrLf &
                           $"Added to list: {importedCount} records." & vbCrLf &
                           $"Omitted (Duplicates): {skippedCount} records." & vbCrLf & vbCrLf &
                           $"Data has been automatically synced to cycle: {activeCycle}", "Import & Sync Success", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadDepartments()
            LoadDeptListCycleDropdown(activeCycle)
            LoadData()
            If InterchangeTabView.Visibility = Visibility.Visible Then
                LoadInterchangeData()
            End If
        Catch ex As Exception
            MessageBox.Show("Error during import: " & ex.Message)
        Finally
            BtnImport.IsEnabled = True
            BtnImport.Content = "📤 Import"
            MainProgressBar.Visibility = Visibility.Collapsed
            MainProgressBar.IsIndeterminate = False
        End Try
    End Function

    Private Sub BtnUndoImport_Click(sender As Object, e As RoutedEventArgs)
        ' Undo logic: delete the records from the last import session
        Try
            ' Get the latest batch tag
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT ImportBatch FROM department_list WHERE ImportBatch IS NOT NULL AND ImportBatch != '' ORDER BY ImportBatch DESC LIMIT 1")
            If dt.Rows.Count = 0 Then
                MessageBox.Show("No import records found to undo.", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            Dim latestBatch = dt.Rows(0)(0).ToString()
            Dim countDt = _mySql.ReadDatatable($"SELECT COUNT(*) FROM department_list WHERE ImportBatch = '{latestBatch}'")
            Dim count = countDt.Rows(0)(0).ToString()

            Dim result = MessageBox.Show($"Are you sure you want to undo the last import? This will remove {count} records added at {latestBatch}.", "Confirm Undo", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.Yes Then
                If _mySql.ExecuteNonQuery($"DELETE FROM department_list WHERE ImportBatch = '{latestBatch}'") Then
                    MessageBox.Show("Undo successful.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                    LoadDepartments()
                    LoadData()
                Else
                    MessageBox.Show("Undo failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error performing undo: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnHistoryToggleFilter_Click(sender As Object, e As RoutedEventArgs)
        If HistoryDeptFilterContainer.Visibility = Visibility.Collapsed Then
            HistoryDeptFilterContainer.Visibility = Visibility.Visible
            TxtHistoryToggleFilter.Text = "Hide Filters"
        Else
            HistoryDeptFilterContainer.Visibility = Visibility.Collapsed
            TxtHistoryToggleFilter.Text = "Show Filters"
        End If
    End Sub


    Private Sub LoadInterHistoryDepartments()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DepartmentName FROM departmentmaster ORDER BY DepartmentName")
            Dim deptList As New List(Of String)()
            deptList.Add("All")
            For Each row As DataRow In dt.Rows
                deptList.Add(row("DepartmentName").ToString())
            Next
            HistoryDeptFiltersItemsControl.ItemsSource = deptList
        Catch ex As Exception
            Console.WriteLine("Error loading history departments: " & ex.Message)
        End Try
    End Sub

    Private Sub HistoryDeptFilterButton_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        _selectedHistoryDept = btn.Content.ToString()

        ' Update UI tags for visual feedback
        For Each item In HistoryDeptFiltersItemsControl.ItemsSource
            Dim container = HistoryDeptFiltersItemsControl.ItemContainerGenerator.ContainerFromItem(item)
            If container IsNot Nothing Then
                Dim childBtn = FindVisualChild(Of Button)(container)
                If childBtn IsNot Nothing Then
                    childBtn.Tag = If(item.ToString() = _selectedHistoryDept, "Selected", "")
                End If
            End If
        Next

        LoadHistoryData()
    End Sub

    Private Sub BtnInterToggleFilter_Click(sender As Object, e As RoutedEventArgs)
        If InterDeptFilterContainer.Visibility = Visibility.Collapsed Then
            InterDeptFilterContainer.Visibility = Visibility.Visible
            TxtInterToggleFilter.Text = "Hide Filters"
        Else
            InterDeptFilterContainer.Visibility = Visibility.Collapsed
            TxtInterToggleFilter.Text = "Show Filters"
        End If
    End Sub

    Private Sub LoadInterchangeDepartments()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DepartmentName FROM departmentmaster ORDER BY DepartmentName")
            Dim deptList As New List(Of String)()
            deptList.Add("All")
            For Each row As DataRow In dt.Rows
                deptList.Add(row("DepartmentName").ToString())
            Next
            InterDeptFiltersItemsControl.ItemsSource = deptList
        Catch ex As Exception
            Console.WriteLine("Error loading interchange departments: " & ex.Message)
        End Try
    End Sub

    Private Sub InterDeptFilterButton_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        _selectedInterDept = btn.Content.ToString()

        ' Update UI tags for visual feedback
        For Each item In InterDeptFiltersItemsControl.ItemsSource
            Dim container = InterDeptFiltersItemsControl.ItemContainerGenerator.ContainerFromItem(item)
            If container IsNot Nothing Then
                Dim childBtn = FindVisualChild(Of Button)(container)
                If childBtn IsNot Nothing Then
                    childBtn.Tag = If(item.ToString() = _selectedInterDept, "Selected", "")
                End If
            End If
        Next

        LoadInterchangeData()
    End Sub


    Private Sub TxtHistorySearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            BtnHistorySearch_Click(Nothing, Nothing)
            e.Handled = True
        End If
    End Sub

    Private Sub BtnHistorySearch_Click(sender As Object, e As RoutedEventArgs)
        LoadHistoryData()
    End Sub

    Private Sub LoadHistoryData()
        If _isLoading Then Return
        _isLoading = True
        MainProgressBar.Visibility = Visibility.Visible

        Try
            Dim sortState = SaveSortState(InterHistoryDataGrid)
            Dim searchText = TxtHistorySearch.Text.Trim()
            ' Query both live (interchangeability) and historical (interchange_history) data
            Dim query = "SELECT ControlNo, Department, Color, CycleName, Status FROM (" &
                        "  SELECT ControlNo, Department, Color, CycleName, Status FROM interchangeability " &
                        "  UNION ALL " &
                        "  SELECT ControlNo, Department, Color, CycleName, Status FROM interchange_history" &
                        ") as all_data WHERE 1=1"

            If _selectedHistoryDept <> "All" Then
                query &= $" AND Department = '{_selectedHistoryDept.Replace("'", "''")}'"
            End If
            If Not String.IsNullOrEmpty(searchText) Then
                query &= $" AND ControlNo LIKE '%{searchText.Replace("'", "''")}%'"
            End If

            Dim dtRaw = _mySql.ReadDatatable(query)

            If dtRaw.Rows.Count = 0 Then
                InterHistoryDataGrid.ItemsSource = Nothing
                Return
            End If

            ' Target Structure: SNo, Control No, Dept, Colour, [Cycle 1], [Cycle 2], ...
            Dim dtPivot As New DataTable()
            dtPivot.Columns.Add("SNo", GetType(Integer))
            dtPivot.Columns.Add("Control No", GetType(String))
            dtPivot.Columns.Add("Dept", GetType(String))
            dtPivot.Columns.Add("Colour", GetType(String))

            ' Identify all unique cycles
            Dim cycles = dtRaw.AsEnumerable().Where(Function(r) Not String.IsNullOrEmpty(r.Field(Of String)("CycleName"))).Select(Function(r) r.Field(Of String)("CycleName")).Distinct().ToList()
            ' Sort cycles chronologically (oldest to newest for columns)
            cycles.Sort(Function(a, b) ParseCycleForSort(a).CompareTo(ParseCycleForSort(b)))

            ' Identify all control no groups
            Dim ctrlGroups = dtRaw.AsEnumerable().GroupBy(Function(r) r.Field(Of String)("ControlNo")).ToList()

            ' Fetch event counts per ControlNo+CycleName across all transaction tables to mark history
            ' If a control no has 2 or more total events in a cycle, it is marked as "History"
            Dim reintroMap As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            If cycles.Count > 0 Then
                ' Filter out null group keys if any
                Dim ctrlList = ctrlGroups.Where(Function(g) g.Key IsNot Nothing).Select(Function(g) g.Key.ToString().Replace("'", "''")).ToList()
                If ctrlList.Count > 0 Then
                    Dim ctrlIn = String.Join("','", ctrlList)

                    ' Count distinct event types per ControlNo+CycleName from each transaction table
                    ' Each table counts as 1 event type (even if it has multiple records)
                    Dim eventCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                    Dim tables() As String = {"receive", "issue", "wop", "writeoff", "reintroduction", "temporary_issuance"}

                    For Each tbl In tables
                        Try
                            Dim dtEvents = _mySql.ReadDatatable($"SELECT ControlNo, CycleName FROM `{tbl}` WHERE ControlNo IN ('{ctrlIn}')")
                            For Each row As DataRow In dtEvents.Rows
                                Dim key = row("ControlNo").ToString().Trim() & "|" & row("CycleName").ToString().Trim()
                                If Not eventCounts.ContainsKey(key) Then
                                    eventCounts(key) = 0
                                End If
                                eventCounts(key) += 1
                            Next
                        Catch ex As Exception
                            Console.WriteLine($"Error reading {tbl}: {ex.Message}")
                        End Try
                    Next

                    ' Build the map of ControlNo+CycleName combos with 2+ events
                    For Each kvp In eventCounts
                        If kvp.Value >= 2 Then
                            Dim parts = kvp.Key.Split("|"c)
                            Dim cNo = parts(0)
                            Dim cName = parts(1)
                            If Not reintroMap.ContainsKey(cNo) Then reintroMap(cNo) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                            reintroMap(cNo).Add(cName)
                        End If
                    Next
                End If
            End If

            For Each c In cycles
                If Not dtPivot.Columns.Contains(c) Then dtPivot.Columns.Add(c, GetType(String))
            Next

            Dim sno = 1
            For Each group In ctrlGroups
                ' Sort records within the group by CycleName descending to get the LATEST metadata (Dept, Color)
                Dim sortedRecords = group.OrderByDescending(Function(r) ParseCycleForSort(r.Field(Of String)("CycleName"))).ToList()
                Dim latest = sortedRecords.First()

                Dim newRow = dtPivot.NewRow()
                newRow("SNo") = sno
                newRow("Control No") = group.Key
                newRow("Dept") = latest.Field(Of String)("Department")
                newRow("Colour") = latest.Field(Of String)("Color")

                ' Fill cycle columns
                For Each histRow In group
                    Dim cycle = histRow.Field(Of String)("CycleName")
                    Dim status = histRow.Field(Of String)("Status")

                    ' Mark as History if reintroduction exists (multiple actions in cycle)
                    Dim cleanCtrl = group.Key.ToString().Trim()
                    Dim cleanCycle = cycle.Trim()
                    If reintroMap.ContainsKey(cleanCtrl) AndAlso reintroMap(cleanCtrl).Contains(cleanCycle) Then
                        status = "History"
                    End If

                    newRow(cycle) = status
                Next
                dtPivot.Rows.Add(newRow)
                sno += 1
            Next

            InterHistoryDataGrid.ItemsSource = dtPivot.DefaultView
            RestoreSortState(InterHistoryDataGrid, sortState)
        Catch ex As Exception
            MessageBox.Show("Error loading history data: " & ex.Message)
        Finally
            _isLoading = False
            MainProgressBar.Visibility = Visibility.Collapsed
        End Try
    End Sub

    Private Sub InterHistoryDataGrid_AutoGeneratingColumn(sender As Object, e As DataGridAutoGeneratingColumnEventArgs)
        ' Formatting for the pivot table
        If e.PropertyName = "SNo" Then
            e.Column.Header = "S.no"
            e.Column.Width = 60
        ElseIf e.PropertyName = "Control No" Then
            e.Column.Width = 150
        ElseIf e.PropertyName = "Dept" Then
            e.Column.Width = 150
        ElseIf e.PropertyName = "Colour" Then
            e.Column.Header = "Colour"
            e.Column.Width = 100
        Else
            ' Cycle columns
            e.Column.Width = 150
        End If

        ' Apply consistent cell styling to all columns
        Dim style As New Style(GetType(TextBlock))
        style.Setters.Add(New Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center))
        style.Setters.Add(New Setter(TextBlock.PaddingProperty, New Thickness(16, 0, 16, 0)))
        style.Setters.Add(New Setter(TextBlock.FontSizeProperty, 14.0))
        style.Setters.Add(New Setter(TextBlock.ForegroundProperty, New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#1E293B"), Color))))

        ' Add triggers for history highlighting (blue underline for actions)
        If Not (e.PropertyName = "SNo" Or e.PropertyName = "Control No" Or e.PropertyName = "Dept" Or e.PropertyName = "Colour") Then
            Dim triggerValues() As String = {"Issued", "Received", "WOP", "Write off", "Re-introduction", "Temp issuance",
                                             "Pending(Reintroduced)", "Issued(Reintroduced)", "Received(Reintroduced)",
                                             "Issue(Reintroduced)", "Receive(Reintroduced)", "WOP(Reintroduced)",
                                             "Write off(Reintroduced)", "Temp issuance(Reintroduced)", "History"}
            For Each triggerVal In triggerValues
                Dim dtTrigger As New DataTrigger()
                dtTrigger.Binding = New Data.Binding("[" & e.PropertyName & "]")
                dtTrigger.Value = triggerVal
                dtTrigger.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Blue))
                dtTrigger.Setters.Add(New Setter(TextBlock.TextDecorationsProperty, TextDecorations.Underline))
                dtTrigger.Setters.Add(New Setter(TextBlock.CursorProperty, Cursors.Hand))
                style.Triggers.Add(dtTrigger)
            Next
        End If
        If TypeOf e.Column Is DataGridTextColumn Then
            DirectCast(e.Column, DataGridTextColumn).ElementStyle = style
        End If
    End Sub

    Private Sub BtnHistoryClear_Click(sender As Object, e As RoutedEventArgs)
        Dim result = MessageBox.Show("Are you sure you want to clear ALL interchange history records? This action cannot be undone.", "Confirm Clear History", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        If result = MessageBoxResult.Yes Then
            Try
                If _mySql.ExecuteNonQuery("TRUNCATE TABLE interchange_history") Then
                    MessageBox.Show("Interchange history has been cleared.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                    LoadHistoryData()
                Else
                    ' Fallback if truncate fails (e.g. constraints)
                    _mySql.ExecuteNonQuery("DELETE FROM interchange_history")
                    LoadHistoryData()
                End If
            Catch ex As Exception
                MessageBox.Show("Error clearing history: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub InterHistoryDataGrid_MouseUp(sender As Object, e As MouseButtonEventArgs)
        If InterHistoryDataGrid.CurrentCell.IsValid Then
            Dim col = InterHistoryDataGrid.CurrentCell.Column
            ' Ensure it's a dynamic cycle column (not S.no, Control No, etc)
            If col IsNot Nothing AndAlso col.Header IsNot Nothing AndAlso
               col.Header.ToString() <> "S.no" AndAlso col.Header.ToString() <> "Control No" AndAlso
               col.Header.ToString() <> "Dept" AndAlso col.Header.ToString() <> "Colour" Then

                Dim rowItem = InterHistoryDataGrid.CurrentCell.Item
                If TypeOf rowItem Is DataRowView Then
                    Dim row = DirectCast(rowItem, DataRowView)
                    Dim cycleName = col.Header.ToString()
                    Dim val = ""
                    Try
                        val = row(cycleName).ToString()
                    Catch
                    End Try

                    If val = "Issued" OrElse val = "Received" OrElse val = "WOP" OrElse val = "Write off" OrElse val = "Re-introduction" OrElse val = "Temp issuance" OrElse val.Contains("(Reintroduced)") OrElse val = "History" Then
                        Dim ctrlNo = row("Control No").ToString()
                        Dim eventType = val ' Pass full value like "WOP(Reintroduced)" or "History"

                        If val = "History" Then
                            ' Open the new Cycle History timeline window
                            Dim cycleWin As New CycleHistoryWindow(ctrlNo, cycleName)
                            cycleWin.ShowDialog()
                        Else
                            ' Open standard single event details
                            Dim details = _mySql.GetEventDetails(ctrlNo, cycleName, eventType)
                            Dim win As New EventDetailsWindow("Event Details: " & eventType, details)
                            win.ShowDialog()
                        End If

                        ' Reset selection to avoid accidental double clicks
                        InterHistoryDataGrid.SelectedItems.Clear()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub StatusDetails_Click(sender As Object, e As RoutedEventArgs)
        Dim link = DirectCast(sender, Hyperlink)
        Dim row = TryCast(link.DataContext, DataRowView)
        If row IsNot Nothing Then
            Dim ctrlNo = ""
            If row.DataView.Table.Columns.Contains("Control No") Then
                ctrlNo = row("Control No").ToString()
            ElseIf row.DataView.Table.Columns.Contains("ControlNo") Then
                ctrlNo = row("ControlNo").ToString()
            End If

            Dim status = row("Status").ToString()
            ' Try to get cycle from ComboCycle
            Dim selectedCycle = If(ComboCycle.SelectedItem IsNot Nothing, ComboCycle.SelectedItem.ToString(), "")

            If status = "History" Then
                ' Open the Cycle History timeline window
                Dim cycleWin As New CycleHistoryWindow(ctrlNo, selectedCycle)
                cycleWin.ShowDialog()
            Else
                ' Open standard single event details (shows remarks)
                Dim details = _mySql.GetEventDetails(ctrlNo, selectedCycle, status)
                Dim win As New EventDetailsWindow("Event Details: " & status, details)
                win.ShowDialog()
            End If
        End If
    End Sub

    Private Sub StatusLink_Click(sender As Object, e As RoutedEventArgs)
        ' Redirect to the shared details handler
        StatusDetails_Click(sender, e)
    End Sub

    Private Sub BtnResetInter_Click(sender As Object, e As RoutedEventArgs)
        ' The topmost cycle in the combobox is always the first item as they are sorted latest-first.
        Dim currentCycle = If(ComboCycle.Items.Count > 0, ComboCycle.Items(0).ToString(), GetCurrentCycleName())

        Dim dialog As New ResetInterchangeabilityDialog(currentCycle)
        If dialog.ShowDialog() = True Then
            If dialog.Result = ResetInterchangeabilityDialog.ResetType.CurrentCycle Then
                ResetCurrentCycle(currentCycle)
            ElseIf dialog.Result = ResetInterchangeabilityDialog.ResetType.All Then
                ResetAllData()
            End If
        End If
    End Sub

    Private Sub ResetCurrentCycle(cycleName As String)
        Dim result = MessageBox.Show($"Are you sure you want to RESET data for cycle: {cycleName}?" & vbCrLf & vbCrLf &
                                   "This will clear all interchange tracking, department list, and calibration records for THIS cycle ONLY." & vbCrLf &
                                   "Associated documents (WriteOff, Reintroduction, TempIssuance, Calibration Reports) will also be deleted." & vbCrLf &
                                   "The system will roll back to the previous cycle." & vbCrLf &
                                   "This action cannot be undone.", "Confirm Cycle Reset", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        If result = MessageBoxResult.Yes Then
            Try
                Dim pCycle As New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName)
                Dim root = ProjectSettings.GetFileStorageRoot()

                ' 1. Collect DocumentPaths for this cycle BEFORE deleting DB rows
                '    so we can remove only the files that belong to this specific cycle.
                Dim docPaths As New List(Of String)
                For Each tbl In {"writeoff", "reintroduction", "temporary_issuance"}
                    Dim dtDocs = _mySql.ReadDatatable($"SELECT DocumentPath FROM `{tbl}` WHERE CycleName = '{cycleName.Replace("'", "''")}' AND DocumentPath IS NOT NULL AND DocumentPath <> ''")
                    For Each row As DataRow In dtDocs.Rows
                        Dim rel = row(0).ToString().Trim()
                        If Not String.IsNullOrEmpty(rel) Then
                            docPaths.Add(IO.Path.Combine(root, rel))
                        End If
                    Next
                Next

                ' 2. Department List for this cycle
                _mySql.ExecuteNonQuery("DELETE FROM department_list WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))

                ' 3. Interchangeability data for this cycle
                _mySql.ExecuteNonQuery("DELETE FROM interchangeability WHERE CycleName = @cycle", pCycle)
                _mySql.ExecuteNonQuery("DELETE FROM writeoff WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))
                _mySql.ExecuteNonQuery("DELETE FROM reintroduction WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))
                _mySql.ExecuteNonQuery("DELETE FROM temporary_issuance WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))
                _mySql.ExecuteNonQuery("DELETE FROM wop WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))
                _mySql.ExecuteNonQuery("DELETE FROM issue WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))
                _mySql.ExecuteNonQuery("DELETE FROM receive WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))
                _mySql.ExecuteNonQuery("DELETE FROM interchange_history WHERE CycleName = @cycle", New MySql.Data.MySqlClient.MySqlParameter("@cycle", cycleName))

                ' 4. Calibration data for this cycle
                _mySql.WipeCalibrationDataByCycle(cycleName)

                ' 5. Delete only the specific document files that belonged to this cycle
                For Each filePath In docPaths
                    Try
                        If IO.File.Exists(filePath) Then IO.File.Delete(filePath)
                    Catch
                        ' Skip files that cannot be deleted
                    End Try
                Next

                ' 6. Roll back to previous cycle
                Dim previousCycle = GetPreviousCycleName(cycleName)

                ' Check if previous cycle actually has data (exists in interchangeability)
                Dim dtPrevCheck = _mySql.ReadDatatable($"SELECT COUNT(*) FROM interchangeability WHERE CycleName = '{previousCycle.Replace("'", "''")}'")
                Dim prevExists = (dtPrevCheck.Rows.Count > 0 AndAlso Convert.ToInt32(dtPrevCheck.Rows(0)(0)) > 0)

                If prevExists Then
                    _mySql.SetConfigValue("ActiveCycle", previousCycle)
                    MessageBox.Show($"Reset for cycle '{cycleName}' successful." & vbCrLf & $"Active cycle is now: {previousCycle}", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                Else
                    _mySql.SetConfigValue("ActiveCycle", "")
                    MessageBox.Show($"Reset for cycle '{cycleName}' successful." & vbCrLf & "No previous cycle data found. System reset to initial state.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                End If

                RefreshUIAfterReset()
            Catch ex As Exception
                MessageBox.Show("Error during cycle reset: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub ResetAllData()
        Dim result = MessageBox.Show("Are you sure you want to RESET ALL application data? This will clear:" & vbCrLf &
                                   "- Department List" & vbCrLf &
                                   "- All Interchangeability Tracking & Cycle History" & vbCrLf &
                                   "- All Calibration Status Tracking" & vbCrLf &
                                   "- All Individual Calibration Results (RECORD type)" & vbCrLf &
                                   "- All associated documents (WriteOff, Reintroduction, TempIssuance, Calibration Reports, Exports)" & vbCrLf & vbCrLf &
                                   "This action cannot be undone.", "Confirm Master Reset", MessageBoxButton.YesNo, MessageBoxImage.Warning)

        If result = MessageBoxResult.Yes Then
            Try
                ' 1. Interchangeability & Master List Reset
                _mySql.ExecuteNonQuery("TRUNCATE TABLE department_list")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE interchangeability")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE writeoff")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE reintroduction")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE temporary_issuance")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE wop")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE issue")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE receive")
                _mySql.ExecuteNonQuery("TRUNCATE TABLE interchange_history")

                ' 2. Calibration Reset
                _mySql.WipeCalibrationData()

                ' 3. Delete all associated documents from file system
                Dim root = ProjectSettings.GetFileStorageRoot()
                DeleteFolderContents(IO.Path.Combine(root, "Interchangeability", "WriteOff"))
                DeleteFolderContents(IO.Path.Combine(root, "Interchangeability", "Reintroduction"))
                DeleteFolderContents(IO.Path.Combine(root, "Interchangeability", "TempIssuance"))
                DeleteFolderContents(IO.Path.Combine(root, "Calibration", "Reports"))
                DeleteFolderContents(IO.Path.Combine(root, "Records", "Exports"))

                ' 4. Configuration & UI Reset
                _mySql.SetConfigValue("ActiveCycle", "")

                MessageBox.Show("Reset successful. All records, transactions, calibration history, and associated documents cleared.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)

                RefreshUIAfterReset()
            Catch ex As Exception
                MessageBox.Show("Error during full reset: " & ex.Message)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Deletes all files and subdirectories inside <paramref name="folderPath"/> but
    ''' keeps the folder itself so the directory structure remains intact.
    ''' Silently skips if the folder does not exist.
    ''' </summary>
    Private Sub DeleteFolderContents(folderPath As String)
        Try
            If Not IO.Directory.Exists(folderPath) Then Return
            For Each f In IO.Directory.GetFiles(folderPath)
                IO.File.Delete(f)
            Next
            For Each d In IO.Directory.GetDirectories(folderPath)
                IO.Directory.Delete(d, True)
            Next
        Catch ex As Exception
            Console.WriteLine($"DeleteFolderContents Error ({folderPath}): {ex.Message}")
        End Try
    End Sub

    Private Sub RefreshUIAfterReset()
        ' Re-navigate to a fresh InterchangeablePage to fully reinitialize with the new active cycle
        ' This ensures all tabs (Department List, Interchange, History) reflect the rolled-back cycle
        ' CalibrationPage will also pick up the new ActiveCycle when navigated to (it reads config on construction)
        Dim mainWindow = TryCast(Window.GetWindow(Me), MainWindow)
        If mainWindow IsNot Nothing Then
            mainWindow.MainFrame.Navigate(New InterchangeablePage())
        Else
            ' Fallback: refresh in-place
            Dim currentCycleName = GetCurrentCycleName()
            LoadDepartments()
            LoadDeptListCycleDropdown(currentCycleName)
            LoadData()
            LoadHistoryData()
            InitializeInterchangeUI(currentCycleName)
            LoadCycleDropdown(currentCycleName)
        End If
    End Sub

    Private Sub BtnToggleFilter_Click(sender As Object, e As RoutedEventArgs)
        If DeptFilterContainer.Visibility = Visibility.Collapsed Then
            DeptFilterContainer.Visibility = Visibility.Visible
            TxtToggleFilter.Text = "Hide Filters"
        Else
            DeptFilterContainer.Visibility = Visibility.Collapsed
            TxtToggleFilter.Text = "Show Filters"
        End If
    End Sub

    Private _currentContextControlNo As String = ""
    Private _currentContextStatus As String = ""
    Private _currentContextRemarks As String = ""
    Private _currentGridHasData As Boolean = False
    Private _currentContextDept As String = ""
    Private _currentContextName As String = ""
    Private _currentContextColor As String = ""
    Private _currentContextSize As String = ""

    Private Function GetCurrentCycleName() As String
        Dim overrideCycle = _mySql.GetConfigValue("ActiveCycle")
        If Not String.IsNullOrEmpty(overrideCycle) Then
            Return overrideCycle
        End If
        Return CycleManager.GetNaturalCycleNameForDate(DateTime.Today, _mySql)
    End Function


    Private Sub InitializeInterchangeUI(activeCycle As String)
        TxtActiveCycle.Text = $"Active today: {activeCycle}"

        ' Show range from Admin-configured dates (preferred) or fallback to parsing cycle name
        Try
            Dim gsVal = _mySql.GetConfigValue("GreenCycleStart")
            Dim geVal = _mySql.GetConfigValue("GreenCycleEnd")
            Dim ysVal = _mySql.GetConfigValue("YellowCycleStart")
            Dim yeVal = _mySql.GetConfigValue("YellowCycleEnd")

            If Not String.IsNullOrEmpty(gsVal) AndAlso Not String.IsNullOrEmpty(geVal) AndAlso
               Not String.IsNullOrEmpty(ysVal) AndAlso Not String.IsNullOrEmpty(yeVal) Then

                Dim greenStart = DateTime.Parse(gsVal)
                Dim greenEnd = DateTime.Parse(geVal)
                Dim yellowStart = DateTime.Parse(ysVal)
                Dim yellowEnd = DateTime.Parse(yeVal)

                Dim activeYear As Integer = DateTime.Today.Year
                Try
                    Dim apos = activeCycle.IndexOf("'"c)
                    If apos >= 0 AndAlso activeCycle.Length > apos + 2 Then
                        activeYear = 2000 + Integer.Parse(activeCycle.Substring(apos + 1, 2))
                    End If
                Catch
                End Try

                If activeCycle.ToLower().Contains("green") Then
                    TxtCycleRange.Text = $"{activeCycle} ({greenStart.ToString("dd-MMM")}-{activeYear} to {greenEnd.ToString("dd-MMM")}-{activeYear})"
                Else
                    TxtCycleRange.Text = $"{activeCycle} ({yellowStart.ToString("dd-MMM")}-{activeYear} to {yellowEnd.ToString("dd-MMM")}-{activeYear})"
                End If
            Else
                ' Fallback: derive range from cycle name (legacy behaviour)
                Dim parts = activeCycle.Split(" "c)
                Dim tag = parts(0)
                Dim month = tag.Split("'"c)(0)
                Dim year = 2000 + Integer.Parse(tag.Split("'"c)(1))
                If month = "Jan" Then
                    TxtCycleRange.Text = $"{activeCycle} (01-Jan-{year} to 30-Jun-{year})"
                Else
                    TxtCycleRange.Text = $"{activeCycle} (01-Jul-{year} to 31-Dec-{year})"
                End If
            End If
        Catch ex As Exception
            TxtCycleRange.Text = activeCycle
        End Try

        LoadCycleDropdown(activeCycle)
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

    Private Sub LoadCycleDropdown(activeCycle As String)
        RemoveHandler ComboCycle.SelectionChanged, AddressOf ComboCycle_SelectionChanged
        ComboCycle.Items.Clear()

        Dim cycleList As New List(Of String)()
        If Not String.IsNullOrEmpty(activeCycle) Then cycleList.Add(activeCycle)

        ' Fetch historical cycles from DB
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DISTINCT CycleName FROM interchangeability")
            For Each row As DataRow In dt.Rows
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
            ComboCycle.Items.Add(c)
        Next

        If ComboCycle.Items.Count > 0 Then
            ComboCycle.SelectedIndex = 0
        End If

        AddHandler ComboCycle.SelectionChanged, AddressOf ComboCycle_SelectionChanged
        LoadInterchangeData()
    End Sub

    Private Sub ComboCycle_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        LoadInterchangeData()
        EvaluateActionButtons()
    End Sub

    Private Sub LoadInterchangeData(Optional searchCtrlNo As String = "")
        If ComboCycle.SelectedItem Is Nothing Then Return
        Dim selectedCycle = ComboCycle.SelectedItem.ToString()

        Try
            Dim sortState = SaveSortState(InterchangeDataGrid)
            Dim query As String
            Dim whereClause As String = $"WHERE CycleName = '{selectedCycle.Replace("'", "''")}'"

            ' Add Department Filter
            If _selectedInterDept <> "All" Then
                whereClause &= $" AND Department = '{_selectedInterDept.Replace("'", "''")}'"
            End If

            ' Add Search Filter
            Dim searchText = TxtInterSearch.Text.Trim()
            If Not String.IsNullOrEmpty(searchText) Then
                whereClause &= $" AND ControlNo LIKE '%{searchText.Replace("'", "''")}%'"
            End If

            If Not String.IsNullOrEmpty(searchCtrlNo) Then
                ' Pin searched item to top
                query = $"SELECT * FROM interchangeability {whereClause} " &
                        $"ORDER BY CASE WHEN ControlNo = '{searchCtrlNo.Replace("'", "''")}' THEN 0 ELSE 1 END, ID DESC"
            Else
                ' Default sorting
                query = $"SELECT * FROM interchangeability {whereClause} ORDER BY ID DESC"
            End If

            Dim dt = _mySql.ReadDatatable(query)

            ' Add computed columns (RFID_tag is already in SELECT *)
            If Not dt.Columns.Contains("HasRFID") Then
                dt.Columns.Add("HasRFID", GetType(Boolean))
            End If
            If Not dt.Columns.Contains("IsReceivedWithRFID") Then
                dt.Columns.Add("IsReceivedWithRFID", GetType(Boolean))
            End If

            ' Fetch reintroduction list for the selected cycle to mark status
            Dim dtReintro = _mySql.ReadDatatable($"SELECT ControlNo FROM reintroduction WHERE CycleName = '{selectedCycle.Replace("'", "''")}'")
            Dim reintroCtrls As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each r As DataRow In dtReintro.Rows
                reintroCtrls.Add(r("ControlNo").ToString().Trim())
            Next

            For Each row As DataRow In dt.Rows
                ' Compute HasRFID from RFID_tag column
                Dim rfidVal As String = ""
                If dt.Columns.Contains("RFID_tag") Then
                    rfidVal = If(IsDBNull(row("RFID_tag")), "", row("RFID_tag").ToString().Trim())
                End If
                Dim hasRfid = Not String.IsNullOrEmpty(rfidVal)
                row("HasRFID") = hasRfid

                ' IsReceivedWithRFID: tag linked AND base status starts with "Received"
                ' Computed BEFORE the "(Reintroduced)" suffix is appended so it covers
                ' both "Received" and "Received(Reintroduced)" in one bool.
                Dim baseStatus = row("Status").ToString().Trim()
                row("IsReceivedWithRFID") = hasRfid AndAlso baseStatus.StartsWith("Received")

                Dim ctrl = row("ControlNo").ToString().Trim()
                If reintroCtrls.Contains(ctrl) Then
                    row("Status") = row("Status").ToString() & "(Reintroduced)"
                End If
            Next

            InterchangeDataGrid.ItemsSource = dt.DefaultView
            RestoreSortState(InterchangeDataGrid, sortState)

            ' Select the pinned item if it exists
            If Not String.IsNullOrEmpty(searchCtrlNo) AndAlso dt.Rows.Count > 0 Then
                InterchangeDataGrid.SelectedIndex = 0
                ' Scroll into view and focus the row after layout is updated
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ContextIdle,
                    CType(Sub()
                              Try
                                  If InterchangeDataGrid.SelectedItem IsNot Nothing Then
                                      InterchangeDataGrid.ScrollIntoView(InterchangeDataGrid.SelectedItem)
                                      Dim row As DataGridRow = CType(InterchangeDataGrid.ItemContainerGenerator.ContainerFromIndex(0), DataGridRow)
                                      If row IsNot Nothing Then
                                          row.MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))
                                      End If
                                  End If
                              Catch
                              End Try
                          End Sub, Action))
            End If
        Catch ex As Exception
            Console.WriteLine("Error loading interchange data: " & ex.Message)
        End Try

        EvaluateActionButtons()
    End Sub

    Private Sub InterchangeDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim row = TryCast(InterchangeDataGrid.SelectedItem, DataRowView)
        If row IsNot Nothing Then
            _currentContextControlNo = row("ControlNo").ToString()
            _currentContextDept = row("Department").ToString()
            _currentContextName = row("InstrumentName").ToString()
            _currentContextSize = row("SizeandRange").ToString()
            _currentContextColor = row("Color").ToString()
            _currentContextStatus = row("Status").ToString()
            _currentContextRemarks = row("Remarks").ToString()

            TxtSelectedControl.Text = $"Selected: {_currentContextControlNo} ({_currentContextName}) - Color: {_currentContextColor} - Status: {_currentContextStatus}"
            TxtSelectedControl.Visibility = Visibility.Visible
            EvaluateActionButtons()
        Else
            ResetSearchContext()
        End If
    End Sub

    Private Sub InterchangeTabView_MouseDown(sender As Object, e As MouseButtonEventArgs)
        ' Clear selection when clicking on the background (unhandled white space)
        InterchangeDataGrid.SelectedItem = Nothing
        Keyboard.ClearFocus()
    End Sub

    Private Sub DeptListTabView_MouseDown(sender As Object, e As MouseButtonEventArgs)
        ' Clear selection when clicking on the background (unhandled white space)
        DeptListDataGrid.SelectedItem = Nothing
        Keyboard.ClearFocus()
    End Sub

    Private Sub InterHistoryTabView_MouseDown(sender As Object, e As MouseButtonEventArgs)
        ' Clear selection when clicking on the background (unhandled white space)
        InterHistoryDataGrid.SelectedItem = Nothing
        Keyboard.ClearFocus()
    End Sub

    Private Sub InterchangeDataGrid_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        ' Double-click on a "Received" row:
        '   - No RFID (white)   -> open standard Receive / link-RFID window
        '   - Has RFID (purple) -> open Change-RFID window to overwrite existing tag
        If String.IsNullOrEmpty(_currentContextControlNo) Then Return
        If ComboCycle.SelectedIndex <> 0 Then Return  ' Active cycle only

        Dim cleanStatus = _currentContextStatus.Replace("(Reintroduced)", "").Trim()
        If Not cleanStatus.StartsWith("Received") Then Return

        ' Check HasRFID from the selected DataRowView
        Dim rowView = TryCast(InterchangeDataGrid.SelectedItem, DataRowView)
        If rowView IsNot Nothing Then
            Dim hasRfid As Boolean = False
            If rowView.DataView.Table.Columns.Contains("HasRFID") Then
                hasRfid = CBool(rowView("HasRFID"))
            End If

            If hasRfid Then
                ' Purple row -- RFID already linked: allow operator to change/re-confirm tag
                OpenChangeRFIDWindowForRow()
            Else
                ' White Received row -- no tag yet: link a new RFID tag
                OpenReceiveWindowForRow()
            End If
        End If
    End Sub

    Private Sub BtnSizeDetails_Click(sender As Object, e As RoutedEventArgs)
        Dim link = DirectCast(sender, Hyperlink)
        Dim row = TryCast(link.DataContext, DataRowView)
        If row IsNot Nothing Then
            Dim ctrlNo = ""
            If row.DataView.Table.Columns.Contains("Control No") Then
                ctrlNo = row("Control No").ToString()
            Else
                ctrlNo = row("ControlNo").ToString()
            End If
            OpenSizeDetailsForControl(ctrlNo)
        End If
    End Sub

    Private Sub OpenSizeDetailsForControl(ctrlNo As String)
        If String.IsNullOrEmpty(ctrlNo) Then Return

        ' Find table name and type
        Dim tblName = "unknown"
        Dim itemType = "unknown"
        Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName, Category FROM type_details")
        For Each r As DataRow In dtSettings.Rows
            Dim testTbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString(), forInventory:=True)
            Dim chk = _mySql.ReadDatatable($"SELECT ControlNo FROM `{testTbl}` WHERE ControlNo='{ctrlNo.Replace("'", "''")}'")
            If chk.Rows.Count > 0 Then
                tblName = testTbl
                itemType = r("Category").ToString()
                Exit For
            End If
        Next

        If tblName = "unknown" Then
            MessageBox.Show("Could not find master list entry for " & ctrlNo, "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim detailsWin As New ItemDetailsWindow()
        detailsWin.WindowStartupLocation = WindowStartupLocation.CenterScreen
        detailsWin.LoadDetails(ctrlNo, itemType, tblName)
        detailsWin.ShowDialog()
    End Sub

    Private Sub ControlNoHistory_Click(sender As Object, e As RoutedEventArgs)
        Dim link = DirectCast(sender, Hyperlink)
        Dim row = TryCast(link.DataContext, DataRowView)
        If row IsNot Nothing Then
            Dim ctrlNo = ""
            If row.DataView.Table.Columns.Contains("Control No") Then
                ctrlNo = row("Control No").ToString()
            Else
                ctrlNo = row("ControlNo").ToString()
            End If
            OpenHistoryForControl(ctrlNo)
        End If
    End Sub

    Private Sub OpenHistoryForControl(ctrlNo As String)
        If String.IsNullOrEmpty(ctrlNo) Then Return

        ' Find table name
        Dim tblName = "unknown"
        Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName FROM type_details")
        For Each r As DataRow In dtSettings.Rows
            Dim testTbl = MySQLClass.TypeNameToTableName(r(0).ToString(), forInventory:=True)
            Dim chk = _mySql.ReadDatatable($"SELECT ControlNo FROM `{testTbl}` WHERE ControlNo='{ctrlNo.Replace("'", "''")}'")
            If chk.Rows.Count > 0 Then tblName = testTbl : Exit For
        Next

        Dim histWin As New HistoryWindow(ctrlNo, "Lifecycle History", tblName)
        histWin.ShowDialog()
    End Sub

    Private Function GetNextCycleName(currentCycle As String) As String
        Try
            Dim safeCycle = currentCycle.Replace("""", "'")
            Dim isYellow = safeCycle.Contains("Yellow Cycle")
            Dim parts = safeCycle.Split(" "c)
            Dim yearPart = Integer.Parse(parts(0).Split("'"c)(1))

            Dim gsVal = _mySql.GetConfigValue("GreenCycleStart")
            Dim ysVal = _mySql.GetConfigValue("YellowCycleStart")
            Dim greenStart As DateTime
            Dim yellowStart As DateTime
            If Not DateTime.TryParse(gsVal, greenStart) Then greenStart = New DateTime(2000, 7, 1)
            If Not DateTime.TryParse(ysVal, yellowStart) Then yellowStart = New DateTime(2000, 1, 1)

            ' Yellow -> next is Green (same year); Green -> next is Yellow (year + 1)
            If isYellow Then
                Return $"{greenStart.ToString("MMM")}'{yearPart.ToString("00")} Green Cycle"
            Else
                Return $"{yellowStart.ToString("MMM")}'{(yearPart + 1).ToString("00")} Yellow Cycle"
            End If
        Catch ex As Exception
            Return "Next Cycle"
        End Try
    End Function

    Private Function ParseCycleToDate(cycleName As String) As DateTime
        ' e.g. Jan'26 Yellow Cycle -> 2026-01-01, Apr'26 Green Cycle -> 2026-04-01
        Try
            Dim parts = cycleName.Split(" "c)
            Dim tag = parts(0) ' e.g. Jan'26 or Apr'26
            Dim monthAbbr = tag.Split("'"c)(0)
            Dim year = 2000 + Integer.Parse(tag.Split("'"c)(1))
            Dim parsedDate As DateTime
            If DateTime.TryParseExact($"01 {monthAbbr} {year}", "dd MMM yyyy",
                                      System.Globalization.CultureInfo.InvariantCulture,
                                      System.Globalization.DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If
            Return DateTime.MinValue
        Catch ex As Exception
            Return DateTime.MinValue
        End Try
    End Function

    Private Function IsCycleHistorical(selectedCycle As String, currentActiveCycle As String) As Boolean
        Dim d1 = ParseCycleToDate(selectedCycle)
        Dim d2 = ParseCycleToDate(currentActiveCycle)
        Return d1 < d2
    End Function

    Private Sub BtnNextCycle_Click(sender As Object, e As RoutedEventArgs)
        Dim activeCycle = GetCurrentCycleName()
        Dim nextCycle = GetNextCycleName(activeCycle)

        Dim result = MessageBox.Show($"Are you sure you want to start the next cycle: {nextCycle}? {vbCrLf}{vbCrLf}" &
                                   "Outstanding items (WOP) will carry over. " &
                                   "Pending items will be marked as WOP (Reason: Not Issued). " &
                                   "Temp Issuance items will be converted to WOP in the next cycle. " &
                                   "Issued/Received items will reset to Pending.", "Start Next Cycle", MessageBoxButton.YesNo, MessageBoxImage.Warning)

        If result = MessageBoxResult.Yes Then
            Dim currentUser = Application.Current.Properties("Username")?.ToString()
            If CycleManager.PerformRollover(activeCycle, nextCycle, _mySql, currentUser) Then
                MessageBox.Show($"Cycle transitioned to {nextCycle} successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                InitializeInterchangeUI(nextCycle)
            Else
                MessageBox.Show("Failed to execute rollover.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub

    Private Function GetPreviousCycleName(currentCycle As String) As String
        Try
            Dim safeCycle = currentCycle.Replace("""", "'")
            Dim isYellow = safeCycle.Contains("Yellow Cycle")
            Dim parts = safeCycle.Split(" "c)
            Dim yearPart = Integer.Parse(parts(0).Split("'"c)(1))

            Dim gsVal = _mySql.GetConfigValue("GreenCycleStart")
            Dim ysVal = _mySql.GetConfigValue("YellowCycleStart")
            Dim greenStart As DateTime
            Dim yellowStart As DateTime
            If Not DateTime.TryParse(gsVal, greenStart) Then greenStart = New DateTime(2000, 7, 1)
            If Not DateTime.TryParse(ysVal, yellowStart) Then yellowStart = New DateTime(2000, 1, 1)

            ' Yellow -> prev is Green (year - 1); Green -> prev is Yellow (same year)
            If isYellow Then
                Return $"{greenStart.ToString("MMM")}'{(yearPart - 1).ToString("00")} Green Cycle"
            Else
                Return $"{yellowStart.ToString("MMM")}'{yearPart.ToString("00")} Yellow Cycle"
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub ResetSearchContext()
        _currentContextControlNo = ""
        _currentContextDept = ""
        _currentContextName = ""
        _currentContextSize = ""
        _currentContextColor = ""
        _currentContextStatus = ""
        TxtSelectedControl.Text = "Selected: () - Color: - Status:"
        TxtSelectedControl.Visibility = Visibility.Collapsed
        EvaluateActionButtons()
    End Sub

    Private Sub TxtInterSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            BtnInterSearch_Click(Nothing, Nothing)
            e.Handled = True
        End If
    End Sub

    Private Sub BtnInterSearch_Click(sender As Object, e As RoutedEventArgs)
        Dim ctrlNo = TxtInterSearch.Text.Trim()
        If String.IsNullOrEmpty(ctrlNo) Then
            ResetSearchContext()
            LoadInterchangeData() ' Restore default view
            Return
        End If

        Try
            ' Step 1: Does it exist in department_list?
            Dim dtDept = _mySql.ReadDatatable($"SELECT * FROM department_list WHERE `Control No` = '{ctrlNo.Replace("'", "''")}' LIMIT 1")
            If dtDept.Rows.Count = 0 Then
                MessageBox.Show("Control No not found in Department List.", "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning)
                ResetSearchContext()
                Return
            End If

            Dim row = dtDept.Rows(0)
            _currentContextControlNo = GetVal(row, dtDept.Columns.IndexOf("Control No"))
            _currentContextDept = GetVal(row, dtDept.Columns.IndexOf("Department"))
            _currentContextName = GetVal(row, dtDept.Columns.IndexOf("InstrumentName"))
            _currentContextSize = GetVal(row, dtDept.Columns.IndexOf("SizeandRange"))
            _currentContextColor = GetVal(row, dtDept.Columns.IndexOf("Color"))

            ' Step 2: What is its status in the CURRENT selected cycle?
            Dim selectedCycle = GetCurrentCycleName()
            Dim dtInter = _mySql.ReadDatatable($"SELECT Status FROM interchangeability WHERE ControlNo = '{_currentContextControlNo.Replace("'", "''")}' AND CycleName = '{selectedCycle.Replace("'", "''")}' ORDER BY ID DESC LIMIT 1")

            If dtInter.Rows.Count > 0 Then
                _currentContextStatus = dtInter.Rows(0)("Status").ToString()
            Else
                _currentContextStatus = "Pending"
            End If

            TxtSelectedControl.Text = $"Selected: {_currentContextControlNo} ({_currentContextName}) - Color: {_currentContextColor} - Status: {_currentContextStatus}"

            ' REFRESH TABLE TO PIN SEARCHED ITEM
            LoadInterchangeData(ctrlNo)

            EvaluateActionButtons()
        Catch ex As Exception
            MessageBox.Show("Search Error: " & ex.Message)
            ResetSearchContext()
        End Try
    End Sub

    Private Sub EvaluateActionButtons()
        ' Disable all first
        BtnIssued.IsEnabled = False
        BtnReceived.IsEnabled = False
        BtnWOP.IsEnabled = False ' Disabled by default until Active Cycle is confirmed
        BtnWriteOff.IsEnabled = False
        BtnTempIssuance.IsEnabled = False
        BtnReintroduction.IsEnabled = False
        BtnBulkScan.IsEnabled = False
        BtnNextCycle.IsEnabled = False

        ' Global Rule 1: Non-Active Cycle (Not Topmost) - ALL Grey
        If ComboCycle.SelectedIndex <> 0 Then
            ' Clear tags to avoid BulkMode styling (dark blue)
            BtnWOP.Tag = ""
            BtnWriteOff.Tag = ""
            BtnTempIssuance.Tag = ""
            BtnReintroduction.Tag = ""
            BtnBulkScan.Tag = ""
            BtnChangeColour.Visibility = Visibility.Collapsed
            Return
        End If

        ' Change Colour button: only visible during the very first cycle
        BtnChangeColour.Visibility = If(IsFirstCycle(), Visibility.Visible, Visibility.Collapsed)

        ' Active cycle defaults
        BtnWOP.IsEnabled = True ' Enabled by default for bulk mode
        BtnWriteOff.IsEnabled = True
        BtnTempIssuance.IsEnabled = True
        BtnReintroduction.IsEnabled = True
        BtnBulkScan.IsEnabled = True


        ' If we are on the current cycle, enable Next Cycle button
        BtnNextCycle.IsEnabled = True

        ' Global Rule 2: No instrument/Gauge Selected - Action buttons remain grey
        If String.IsNullOrEmpty(_currentContextControlNo) Then
            ' Set Bulk Mode Color (Dark Blue)
            BtnWOP.Tag = "BulkMode"
            BtnWriteOff.Tag = "BulkMode"
            BtnTempIssuance.Tag = "BulkMode"
            BtnReintroduction.Tag = "BulkMode"
            BtnBulkScan.Tag = "BulkMode"
            Return
        End If

        ' Row selected -> Reset to Normal Mode Color
        BtnWOP.Tag = ""
        BtnWriteOff.Tag = ""
        BtnTempIssuance.Tag = ""
        BtnReintroduction.Tag = ""
        BtnBulkScan.Tag = ""

        ' Once a row is selected, disable these by default
        BtnReintroduction.IsEnabled = False
        BtnWOP.IsEnabled = False
        BtnTempIssuance.IsEnabled = False

        Dim selectedCycle = ComboCycle.SelectedItem.ToString()

        Dim status = _currentContextStatus
        If status IsNot Nothing Then
            status = status.Replace("(Reintroduced)", "").Trim()
        End If
        Dim curColor = _currentContextColor

        ' Global Rule 3: Written Off - ALL Grey, except for Reintroduction
        ' If status = "Write off" Then
        '     BtnReintroduction.IsEnabled = True
        '     Return
        ' End If

        ' Cycle logic for Issued/Received (Based on Selected Cycle color)
        ' Rule: Same Color = Issue. Different Color = Receive.
        Dim cycleIsYellow = selectedCycle.Contains("Yellow")
        Dim cycleIsGreen = selectedCycle.Contains("Green")

        Dim gaugeIsYellow = curColor.Equals("Yellow", StringComparison.OrdinalIgnoreCase)
        Dim gaugeIsGreen = curColor.Equals("Green", StringComparison.OrdinalIgnoreCase)

        Dim colorsMatch = (cycleIsYellow AndAlso gaugeIsYellow) OrElse (cycleIsGreen AndAlso gaugeIsGreen)

        If colorsMatch Then
            ' Same Color -> Issue Only
            If status = "Pending" OrElse status = "Received" Then BtnIssued.IsEnabled = True
            BtnReceived.IsEnabled = False
        Else
            ' Different Color -> Receive Only
            If status = "Pending" OrElse status = "Issued" OrElse status = "Temp issuance" Then BtnReceived.IsEnabled = True
            BtnIssued.IsEnabled = False
        End If

        ' 4. WOP Button
        ' WOP Enabled if status is not Write off and not already WOP
        If status <> "Write off" AndAlso status <> "WOP" Then
            BtnWOP.IsEnabled = True
        End If

        ' 5. Write off Button
        ' Always enabled (managed at start)

        ' 6. Temp issuance Button
        ' Enabled ONLY if status is "Received"
        If status = "Received" Then
            BtnTempIssuance.IsEnabled = True
        End If

        ' 7. Reintroduction Button
        If _currentContextStatus = "WOP" OrElse _currentContextStatus = "Write off" Then
            BtnReintroduction.IsEnabled = True

            ' RESTRICTION: Do not allow reintroduction if it's Calibration NG
            If _currentContextRemarks.Contains("Calibration NG") Then
                BtnReintroduction.IsEnabled = False
            End If
        End If

        ' Bulk scan button is always enabled
        BtnBulkScan.IsEnabled = True
    End Sub

    ''' <summary>
    ''' Returns True if the currently active cycle is the very first cycle ever
    ''' (no more than one distinct CycleName exists in the interchangeability table).
    ''' </summary>
    Private Function IsFirstCycle() As Boolean
        Try
            Dim dt = _mySql.ReadDatatable(
                "SELECT COUNT(DISTINCT CycleName) AS cnt FROM interchangeability")
            If dt.Rows.Count > 0 Then
                Return Convert.ToInt32(dt.Rows(0)("cnt")) <= 1
            End If
        Catch ex As Exception
            Console.WriteLine("IsFirstCycle error: " & ex.Message)
        End Try
        Return False
    End Function

    Private Sub BtnChangeColour_Click(sender As Object, e As RoutedEventArgs)
        ' Safety guard — should never be reachable if not first cycle
        If Not IsFirstCycle() OrElse ComboCycle.SelectedIndex <> 0 Then
            MessageBox.Show("Colour change is only allowed during the first cycle.",
                            "Not Allowed", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Collect all selected rows
        Dim selectedRows As New List(Of DataRowView)()
        For Each item In InterchangeDataGrid.SelectedItems
            Dim drv = TryCast(item, DataRowView)
            If drv IsNot Nothing Then selectedRows.Add(drv)
        Next

        If selectedRows.Count = 0 Then
            MessageBox.Show("Please select one or more rows first.",
                            "No Selection", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        ' Determine majority colour of selection
        Dim greenCount As Integer = 0
        For Each drvi In selectedRows
            If drvi.Row("Color").ToString().Equals("Green", StringComparison.OrdinalIgnoreCase) Then
                greenCount += 1
            End If
        Next
        Dim fromColor As String = If(greenCount >= selectedRows.Count - greenCount, "Green", "Yellow")
        Dim toColor As String = If(fromColor = "Green", "Yellow", "Green")

        ' Build a short preview of control numbers
        Dim ctrlParts As New List(Of String)()
        For Each drvi In selectedRows
            If ctrlParts.Count < 5 Then ctrlParts.Add(drvi.Row("ControlNo").ToString())
        Next
        Dim ctrlList = String.Join(", ", ctrlParts)
        If selectedRows.Count > 5 Then ctrlList &= $" ... (+{selectedRows.Count - 5} more)"

        Dim prompt = $"Selected {selectedRows.Count} item(s) are currently {fromColor}." & vbCrLf &
                     $"Do you want to change their colour to {toColor}?" & vbCrLf & vbCrLf &
                     $"Control Nos: {ctrlList}"

        If MessageBox.Show(prompt, "Change Colour",
                           MessageBoxButton.YesNo, MessageBoxImage.Question) <> MessageBoxResult.Yes Then
            Return
        End If

        Dim selectedCycle = ComboCycle.SelectedItem.ToString()
        Dim successCount As Integer = 0
        Dim errorCount As Integer = 0

        For Each drv In selectedRows
            Dim ctrl = drv("ControlNo").ToString().Replace("'", "''")
            Dim cyc = selectedCycle.Replace("'", "''")
            Try
                _mySql.ExecuteNonQuery(
                    $"UPDATE interchangeability SET Color='{toColor}' WHERE ControlNo='{ctrl}' AND CycleName='{cyc}'")
                _mySql.ExecuteNonQuery(
                    $"UPDATE department_list SET Color='{toColor}' WHERE `Control No`='{ctrl}' AND CycleName='{cyc}'")
                successCount += 1
            Catch ex As Exception
                Console.WriteLine("BtnChangeColour error for " & ctrl & ": " & ex.Message)
                errorCount += 1
            End Try
        Next

        MessageBox.Show(
            If(errorCount = 0,
               $"Colour changed to {toColor} for {successCount} item(s).",
               $"Changed {successCount} item(s). {errorCount} item(s) failed."),
            "Change Colour", MessageBoxButton.OK,
            If(errorCount = 0, MessageBoxImage.Information, MessageBoxImage.Warning))

        LoadInterchangeData()
        EvaluateActionButtons()
    End Sub
    Private Sub BtnAction_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim action = btn.Content.ToString()

        Dim selectedCycle = ComboCycle.SelectedItem.ToString()

        If action = "Issued" Then
            If Not _mySql.IsControlNoCalibratedInAnyTab(_currentContextControlNo, selectedCycle) Then
                MessageBox.Show("Control no. not calibrated", "Not Calibrated", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            If _mySql.IsControlNoCalibrationNGInAnyTab(_currentContextControlNo, selectedCycle) Then
                MessageBox.Show("Calibration is done, but status is NG", "Not Calibrated", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim issueWin As New IssueRFIDWindow(_currentContextControlNo)
            If issueWin.ShowDialog() = True Then
                ' Now handled automatically inside InsertInterchangeRecord
                If _mySql.InsertInterchangeRecord(selectedCycle, _currentContextControlNo, _currentContextDept, _currentContextName, _currentContextSize, _currentContextColor, "Issued", issueWin.Remarks) Then
                    _mySql.ClearRFIDTag(_currentContextControlNo)
                    _currentContextStatus = "Issued"
                End If
            End If
        ElseIf action = "Received" Then
            Dim rfidWin As New ReceiveRFIDWindow(_currentContextControlNo)
            If rfidWin.ShowDialog() = True Then
                ' Now handled automatically inside InsertInterchangeRecord
                If _mySql.InsertInterchangeRecord(selectedCycle, _currentContextControlNo, _currentContextDept, _currentContextName, _currentContextSize, _currentContextColor, "Received", rfidWin.Remarks, rfidWin.RfidTag) Then
                    _currentContextStatus = "Received"
                End If
            End If
        ElseIf action = "WOP" Then
            If String.IsNullOrEmpty(_currentContextControlNo) Then
                Dim bulkWopWin As New BulkWOPWindow()
                If bulkWopWin.ShowDialog() = True Then
                    LoadInterchangeData()
                End If
                Return
            End If

            Dim wopWin As New WOPWindow(_currentContextControlNo, selectedCycle, _currentContextDept, _currentContextName, _currentContextColor, _currentContextSize)
            If wopWin.ShowDialog() = True Then
                _mySql.ClearRFIDTag(_currentContextControlNo) ' Clear RFID when marked as WOP
                _currentContextStatus = "WOP"
            End If
        ElseIf action = "Write off" Then
            If String.IsNullOrEmpty(_currentContextControlNo) Then
                Dim bulkWin As New BulkWriteOffWindow()
                If bulkWin.ShowDialog() = True Then
                    LoadInterchangeData()
                End If
                Return
            End If

            ' Single Write Off Logic
            Dim tblName = "unknown"
            Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName FROM type_details")
            For Each r As DataRow In dtSettings.Rows
                Dim testTbl = MySQLClass.TypeNameToTableName(r(0).ToString(), forInventory:=True)
                Dim chk = _mySql.ReadDatatable($"SELECT ControlNo FROM `{testTbl}` WHERE ControlNo='{_currentContextControlNo.Replace("'", "''")}'")
                If chk.Rows.Count > 0 Then tblName = testTbl : Exit For
            Next

            Dim woWin As New WriteOffWindow(_currentContextControlNo, tblName, selectedCycle, _currentContextSize)
            If woWin.ShowDialog() = True Then
                _mySql.ClearRFIDTag(_currentContextControlNo) ' Clear RFID when written off
                _currentContextStatus = "Write off"
                LoadInterchangeData()
            End If
        ElseIf action = "Temp issuance" Then
            If String.IsNullOrEmpty(_currentContextControlNo) Then
                Dim bulkTempWin As New BulkTempIssuanceWindow()
                If bulkTempWin.ShowDialog() = True Then
                    LoadInterchangeData()
                End If
                Return
            End If

            Dim tempWin As New TempIssuanceWindow(_currentContextControlNo, selectedCycle, _currentContextDept, _currentContextName, _currentContextColor, _currentContextSize)
            If tempWin.ShowDialog() = True Then
                _mySql.ClearRFIDTag(_currentContextControlNo) ' Clear RFID when issued temporarily
                _currentContextStatus = "Temp issuance"
            End If
        ElseIf action = "Reintroduction" Then
            If String.IsNullOrEmpty(_currentContextControlNo) Then ' Check if no item is selected
                Dim bulkReWin As New BulkReintroduceWindow()
                If bulkReWin.ShowDialog() = True Then
                    LoadInterchangeData() ' Refresh data after bulk operation
                End If
            Else
                ' Normal reintroduction logic for selected item
                Dim tblName = "unknown"
                Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName FROM type_details")
                For Each r As DataRow In dtSettings.Rows
                    Dim testTbl = MySQLClass.TypeNameToTableName(r(0).ToString(), forInventory:=True)
                    Dim chk = _mySql.ReadDatatable($"SELECT ControlNo FROM `{testTbl}` WHERE ControlNo='{_currentContextControlNo.Replace("'", "''")}'")
                    If chk.Rows.Count > 0 Then tblName = testTbl : Exit For
                Next
                Dim reWin As New ReintroductionWindow(_currentContextControlNo, tblName, selectedCycle, _currentContextSize)
                If reWin.ShowDialog() = True Then
                    _currentContextStatus = "Pending" ' or received, depending on reintroduction logic
                    LoadInterchangeData() ' Refresh data after single item operation
                End If
            End If
        End If

        TxtSelectedControl.Text = $"Selected: {_currentContextControlNo} ({_currentContextName}) - Color: {_currentContextColor} - Status: {_currentContextStatus}"
        LoadInterchangeData()
        EvaluateActionButtons()
    End Sub

    Private Sub OpenReceiveWindowForRow()
        Dim selectedCycle = ComboCycle.SelectedItem.ToString()
        Dim rfidWin As New ReceiveRFIDWindow(_currentContextControlNo)
        If rfidWin.ShowDialog() = True Then
            If _mySql.InsertInterchangeRecord(selectedCycle, _currentContextControlNo,
                    _currentContextDept, _currentContextName, _currentContextSize,
                    _currentContextColor, "Received", rfidWin.Remarks, rfidWin.RfidTag) Then
                _currentContextStatus = "Received"
            End If
            LoadInterchangeData()
            EvaluateActionButtons()
        End If
    End Sub

    ' Helpers
    Private Function GetVal(row As DataRow, idx As Integer) As String
        If idx = -1 OrElse IsDBNull(row(idx)) Then Return ""
        Return row(idx).ToString().Trim()
    End Function

    Private Function IsMatch(header As String, ParamArray targets() As String) As Boolean
        ' Strip out (Y) or similar suffixes at the end, and trim
        Dim cleanHeader = System.Text.RegularExpressions.Regex.Replace(header, "\s*\([^)]*\)$", "").Replace(" ", "").Trim()
        For Each target In targets
            If cleanHeader.Equals(target.Replace(" ", ""), StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function SaveSortState(dg As DataGrid) As List(Of Tuple(Of String, ListSortDirection))
        Dim sortState As New List(Of Tuple(Of String, ListSortDirection))()
        For Each sd As SortDescription In dg.Items.SortDescriptions
            sortState.Add(New Tuple(Of String, ListSortDirection)(sd.PropertyName, sd.Direction))
        Next
        Return sortState
    End Function

    Private Sub RestoreSortState(dg As DataGrid, sortState As List(Of Tuple(Of String, ListSortDirection)))
        If sortState Is Nothing OrElse sortState.Count = 0 Then Return

        dg.Items.SortDescriptions.Clear()
        For Each state In sortState
            dg.Items.SortDescriptions.Add(New SortDescription(state.Item1, state.Item2))

            ' Restore the arrow in the header
            For Each col In dg.Columns
                ' Check SortMemberPath or Binding Path
                Dim isMatch = False
                If col.SortMemberPath = state.Item1 Then
                    isMatch = True
                ElseIf TypeOf col Is DataGridBoundColumn Then
                    Dim boundCol = DirectCast(col, DataGridBoundColumn)
                    If TypeOf boundCol.Binding Is Data.Binding Then
                        Dim binding = DirectCast(boundCol.Binding, Data.Binding)
                        If binding.Path IsNot Nothing AndAlso binding.Path.Path = state.Item1 Then
                            isMatch = True
                        End If
                    End If
                End If

                If isMatch Then
                    col.SortDirection = state.Item2
                    Exit For
                End If
            Next
        Next
    End Sub

    Private Function FindVisualChild(Of T As DependencyObject)(ByVal obj As DependencyObject) As T
        If obj Is Nothing Then Return Nothing
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

    Private Sub BtnExportCSV_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim view As DataView = TryCast(InterchangeDataGrid.ItemsSource, DataView)
            If view Is Nothing OrElse view.Count = 0 Then
                MessageBox.Show("No data available to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If
            Dim dtExport As DataTable = view.ToTable()

            Dim selectedCycle As String = ""
            If ComboCycle.SelectedItem IsNot Nothing Then
                selectedCycle = ComboCycle.SelectedItem.ToString().Replace(" ", "_")
            End If

            Dim saveFileDialog As New Microsoft.Win32.SaveFileDialog()
            saveFileDialog.Filter = "CSV file (*.csv)|*.csv"
            saveFileDialog.Title = "Export to CSV"
            saveFileDialog.FileName = "Interchange_Export_" & selectedCycle & ".csv"

            If saveFileDialog.ShowDialog() = True Then
                Dim csvContent As New System.Text.StringBuilder()

                ' Dynamically get all columns from DataTable, excluding UI helper columns
                Dim columnHeaders As New List(Of String)()
                columnHeaders.Add("S.no")
                For Each col As DataColumn In dtExport.Columns
                    Dim colName = col.ColumnName
                    If colName <> "HasRFID" AndAlso colName <> "IsReceivedWithRFID" AndAlso colName <> "RFID_tag" Then
                        columnHeaders.Add(colName)
                    End If
                Next
                ' Append RFID Tag as a friendly-labelled column at the end
                columnHeaders.Add("RFID Tag")
                csvContent.AppendLine(String.Join(",", columnHeaders))

                Dim sno As Integer = 1
                For Each row As DataRow In dtExport.Rows
                    Dim values As New List(Of String)()
                    values.Add(sno.ToString())

                    For Each header In columnHeaders
                        If header = "S.no" Then Continue For
                        If header = "RFID Tag" Then
                            Dim rfidVal As String = ""
                            If dtExport.Columns.Contains("RFID_tag") Then
                                rfidVal = If(IsDBNull(row("RFID_tag")), "", row("RFID_tag").ToString().Trim())
                            End If
                            values.Add(EscapeCsv(rfidVal))
                        Else
                            Dim val = row(header).ToString()
                            values.Add(EscapeCsv(val))
                        End If
                    Next
                    csvContent.AppendLine(String.Join(",", values))
                    sno += 1
                Next

                System.IO.File.WriteAllText(saveFileDialog.FileName, csvContent.ToString())
                MessageBox.Show("Data exported successfully!", "Export to CSV", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting data: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnHistoryExportCSV_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim view As DataView = TryCast(InterHistoryDataGrid.ItemsSource, DataView)
            If view Is Nothing OrElse view.Count = 0 Then
                MessageBox.Show("No data available to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If
            Dim dtExport As DataTable = view.ToTable()

            Dim saveFileDialog As New Microsoft.Win32.SaveFileDialog()
            saveFileDialog.Filter = "CSV file (*.csv)|*.csv"
            saveFileDialog.Title = "Export History to CSV"
            saveFileDialog.FileName = "Interchange_History_Export_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".csv"

            If saveFileDialog.ShowDialog() = True Then
                Dim csvContent As New System.Text.StringBuilder()

                ' Get columns exactly as they appear in the Grid (including dynamic cycles)
                Dim headers As New List(Of String)()
                For Each col As DataColumn In dtExport.Columns
                    headers.Add(EscapeCsv(col.ColumnName))
                Next
                csvContent.AppendLine(String.Join(",", headers))

                ' Identify which columns are cycle columns (not SNo, Control No, Dept, Colour)
                Dim fixedCols As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {"SNo", "Control No", "Dept", "Colour"}

                For Each row As DataRow In dtExport.Rows
                    Dim values As New List(Of String)()
                    For Each col As DataColumn In dtExport.Columns
                        Dim cellVal = row(col).ToString()

                        ' Expand "History" to show actual event types
                        If cellVal = "History" AndAlso Not fixedCols.Contains(col.ColumnName) Then
                            Dim ctrlNo = row("Control No").ToString().Replace("'", "''")
                            Dim cycleName = col.ColumnName.Replace("'", "''")
                            cellVal = GetEventListForCycle(ctrlNo, cycleName)
                        End If

                        values.Add(EscapeCsv(cellVal))
                    Next
                    csvContent.AppendLine(String.Join(",", values))
                Next

                System.IO.File.WriteAllText(saveFileDialog.FileName, csvContent.ToString())
                MessageBox.Show("History exported successfully!", "Export to CSV", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting history: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Returns a slash-separated list of distinct event types for a ControlNo in a cycle.
    ''' Example output: "Received/ Temp issuance/ WOP"
    ''' </summary>
    Private Function GetEventListForCycle(ctrlNo As String, cycleName As String) As String
        Try
            Dim query = "SELECT Action FROM (" &
                        $"  SELECT CAST(CONCAT(ReintroductionDate, ' ', Time) AS DATETIME) AS Date, 'Re-introduction' AS Action FROM reintroduction WHERE ControlNo = '{ctrlNo}' AND CycleName = '{cycleName}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(WriteOffDate, ' ', Time) AS DATETIME) AS Date, 'Write off' AS Action FROM writeoff WHERE ControlNo = '{ctrlNo}' AND CycleName = '{cycleName}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(WOPDate, ' ', Time) AS DATETIME) AS Date, 'WOP' AS Action FROM wop WHERE ControlNo = '{ctrlNo}' AND CycleName = '{cycleName}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(IssueDate, ' ', Time) AS DATETIME) AS Date, 'Temp issuance' AS Action FROM temporary_issuance WHERE ControlNo = '{ctrlNo}' AND CycleName = '{cycleName}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(IssueDate, ' ', Time) AS DATETIME) AS Date, 'Issued' AS Action FROM issue WHERE ControlNo = '{ctrlNo}' AND CycleName = '{cycleName}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(ReceiveDate, ' ', Time) AS DATETIME) AS Date, 'Received' AS Action FROM receive WHERE ControlNo = '{ctrlNo}' AND CycleName = '{cycleName}'" &
                        ") as t " &
                        "ORDER BY Date ASC"

            Dim dt = _mySql.ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Dim events As New List(Of String)()
                For Each row As DataRow In dt.Rows
                    Dim act = row("Action").ToString()
                    If Not events.Contains(act) Then events.Add(act)
                Next
                Return String.Join("/ ", events)
            End If
        Catch
        End Try

        Return "History"
    End Function

    Private Sub BtnBulkScan_Click(sender As Object, e As RoutedEventArgs)
        ' If already scanning, StopBulkScan will be called by summary window close or manual check
        If _isBulkScanMode OrElse _isUsbBulkScanMode Then
            StopBulkScan()
            Return
        End If

        ' Show Selection Popup
        Dim selectionWin As New BulkScanSelectionWindow()
        If selectionWin.ShowDialog() = True Then
            _issuedColl.Clear()
            _ignoredColl.Clear()
            _processedTags.Clear()

            Dim activeCycle = GetCurrentCycleName()
            If ComboCycle.SelectedItem IsNot Nothing Then
                activeCycle = ComboCycle.SelectedItem.ToString()
            End If

            ' Start the appropriate scan mode
            If selectionWin.SelectedMethod = BulkScanSelectionWindow.ScanMethod.Gun Then
                _isBulkScanMode = True
                BtnBulkScan.Tag = "Scanning"
                BtnBulkScan.Content = "Stop Gun Scan"

                _rfidGunClient = New RFIDGunClient()
                Dim discardTask = _rfidGunClient.StartStreamingAsync(ProjectSettings.Current.RfidGunIp, ProjectSettings.Current.RfidGunPort)
            ElseIf selectionWin.SelectedMethod = BulkScanSelectionWindow.ScanMethod.Usb Then
                _isUsbBulkScanMode = True
                BtnBulkScan.Tag = "Scanning"
                BtnBulkScan.Content = "Stop USB Scan"
                StartUsbPolling()
            End If

            ' Open Summary Window immediately (Modally)
            ' This blocks the UI but the background task (USB) or events (Gun) will update the collections.
            Dim scanType = If(selectionWin.SelectedMethod = BulkScanSelectionWindow.ScanMethod.Usb,
                             BulkScanSummaryWindow.ScannerType.USB,
                             BulkScanSummaryWindow.ScannerType.RFID_Gun)
            Dim summary As New BulkScanSummaryWindow(_issuedColl, _ignoredColl, _mySql, activeCycle, scanType, Me)

            ' Ensure scan stops when summary window is closed
            summary.ShowDialog()

            ' After summary is closed
            StopBulkScan()
            LoadInterchangeData(activeCycle)
        End If
    End Sub

    Friend Sub StopBulkScan()
        If _isUsbBulkScanMode Then
            _usbScanCts?.Cancel()
            _isUsbBulkScanMode = False
        End If

        If _rfidGunClient IsNot Nothing Then
            _rfidGunClient.StopStreaming()
            _rfidGunClient = Nothing
        End If

        _isBulkScanMode = False
        BtnBulkScan.Tag = "BulkMode"
        BtnBulkScan.Content = "Bulk scan"
    End Sub

    Friend Async Sub StartGunScan()
        If _gunScanTask IsNot Nothing AndAlso Not _gunScanTask.IsCompleted Then
            Try
                Await _gunScanTask
            Catch ex As Exception
            End Try
        End If

        _isBulkScanMode = True
        BtnBulkScan.Tag = "Scanning"
        BtnBulkScan.Content = "Stop Gun Scan"
        _rfidGunClient = New RFIDGunClient()
        _gunScanTask = _rfidGunClient.StartStreamingAsync(ProjectSettings.Current.RfidGunIp, ProjectSettings.Current.RfidGunPort)
        Try
            Await _gunScanTask
        Catch ex As Exception
        End Try
    End Sub

    Friend Sub ClearProcessedTags()
        SyncLock _processedTags
            _processedTags.Clear()
        End SyncLock
    End Sub

    Friend Async Sub StartUsbPolling()
        If _usbScanTask IsNot Nothing AndAlso Not _usbScanTask.IsCompleted Then
            Try
                Await _usbScanTask
            Catch ex As Exception
            End Try
        End If

        _isUsbBulkScanMode = True
        BtnBulkScan.Tag = "Scanning"
        BtnBulkScan.Content = "Stop USB Scan"
        _usbScanCts = New System.Threading.CancellationTokenSource()
        Dim token = _usbScanCts.Token

        Try
            _usbScanTask = Task.Run(Sub()
                                        While Not token.IsCancellationRequested
                                            ' Call the new Bulk Scan method (Reduced to 800ms for snappy UI updates)
                                            Dim tags = CommManager.ScanBulkTags(Nothing, 400)
                                            If tags IsNot Nothing AndAlso tags.Count > 0 Then
                                                For Each t In tags
                                                    ProcessSingleTag(t)
                                                Next
                                            End If
                                            ' Short pause before next batch
                                            System.Threading.Thread.Sleep(10) ' V2: Minimal pause (10ms)
                                        End While
                                    End Sub, token)
            Await _usbScanTask
        Catch ex As TaskCanceledException
            ' Normal stop
        Catch ex As Exception
            MessageBox.Show("USB Scan Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            StopBulkScan()
        End Try
    End Sub

    Private Sub ProcessSingleTag(tag As String)
        ' Skip if already processed in this session (Thread-safe check)
        SyncLock _processedTags
            If _processedTags.Contains(tag) Then Return
            _processedTags.Add(tag)
        End SyncLock

        Task.Run(Sub()
                     Try
                         ' 0. Check if tag is excluded
                         If _mySql.IsTagExcluded(tag) Then
                             Me.Dispatcher.Invoke(Sub() _ignoredColl.Add(New BulkScanSummaryWindow.IgnoredItem() With {.RFID = tag, .Reason = "(Excluded Tag)", .DisplayColor = "#374151"}))
                             Return
                         End If

                         ' 1. Lookup the tool by RFID (Global search)
                         Dim dtCtrl = _mySql.ReadDatatable($"SELECT ControlNo FROM interchangeability WHERE UPPER(REPLACE(RFID_tag, ' ', '')) = '{tag.Replace(" ", "").ToUpper().Replace("'", "''")}' LIMIT 1")

                         If dtCtrl.Rows.Count > 0 Then
                             Dim ctrlNo = dtCtrl.Rows(0)("ControlNo").ToString()

                             ' Get active cycle for check
                             Dim activeCycle = ""
                             Me.Dispatcher.Invoke(Sub()
                                                      If ComboCycle.SelectedItem IsNot Nothing Then
                                                          activeCycle = ComboCycle.SelectedItem.ToString()
                                                      Else
                                                          activeCycle = GetCurrentCycleName()
                                                      End If
                                                  End Sub)

                             ' 2. Check if tool exists in the active cycle and is in-stock
                             Dim dtMatch = _mySql.ReadDatatable($"SELECT InstrumentName, Status, Department FROM interchangeability WHERE ControlNo = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{activeCycle.Replace("'", "''")}'")

                             If dtMatch.Rows.Count > 0 Then
                                 Dim name = dtMatch.Rows(0)("InstrumentName").ToString()
                                 Dim status = dtMatch.Rows(0)("Status").ToString()
                                 Dim dept = dtMatch.Rows(0)("Department").ToString()

                                 If status.Equals("Received", StringComparison.OrdinalIgnoreCase) OrElse status.Equals("Pending", StringComparison.OrdinalIgnoreCase) Then
                                     If Not _mySql.IsControlNoCalibratedInAnyTab(ctrlNo, activeCycle) Then
                                         Me.Dispatcher.Invoke(Sub() _ignoredColl.Add(New BulkScanSummaryWindow.IgnoredItem() With {.RFID = tag, .Reason = "(Not Calibrated)", .ControlNo = "- " & ctrlNo & " | " & name, .DisplayColor = "Red"}))
                                     ElseIf _mySql.IsControlNoCalibrationNGInAnyTab(ctrlNo, activeCycle) Then
                                         Me.Dispatcher.Invoke(Sub() _ignoredColl.Add(New BulkScanSummaryWindow.IgnoredItem() With {.RFID = tag, .Reason = "(Calibration NG)", .ControlNo = "- " & ctrlNo & " | " & name, .DisplayColor = "Red"}))
                                     Else
                                         ' Candidate for Issue
                                         Me.Dispatcher.Invoke(Sub()
                                                                  _issuedColl.Add(New BulkScanSummaryWindow.IssuedItem() With {
                                                                       .ControlNo = ctrlNo,
                                                                       .Name = name,
                                                                       .Department = dept,
                                                                       .RFID = tag
                                                                   })
                                                              End Sub)
                                     End If
                                 Else
                                     ' Wrong status (Already Issued/WOP/etc)
                                     Me.Dispatcher.Invoke(Sub() _ignoredColl.Add(New BulkScanSummaryWindow.IgnoredItem() With {.RFID = tag, .Reason = "(Already " & status & ")", .ControlNo = "- " & ctrlNo & " | " & name, .DisplayColor = "#374151"}))
                                 End If
                             Else
                                 ' Not in this cycle
                                 Me.Dispatcher.Invoke(Sub() _ignoredColl.Add(New BulkScanSummaryWindow.IgnoredItem() With {.RFID = tag, .Reason = "(Not in current cycle)", .ControlNo = "- " & ctrlNo, .DisplayColor = "#374151"}))
                             End If
                         Else
                             ' Unknown Tag
                             Me.Dispatcher.Invoke(Sub() _ignoredColl.Add(New BulkScanSummaryWindow.IgnoredItem() With {.RFID = tag, .Reason = "(Unknown Tag)", .DisplayColor = "#374151"}))
                         End If
                     Catch ex As Exception
                         Console.WriteLine("ProcessSingleTag Error: " & ex.Message)
                     End Try
                 End Sub)
    End Sub

    Private Function EscapeCsv(field As String) As String
        If String.IsNullOrEmpty(field) Then Return ""
        ' If the field contains a comma, quote, or newline, wrap it in quotes and double the quotes inside
        If field.Contains(",") OrElse field.Contains("""") OrElse field.Contains(vbCrLf) OrElse field.Contains(vbLf) Then
            Return $"""{field.Replace("""", """""")}"""
        End If
        Return field
    End Function


    ' ─────────────────────────────────────────────────────────────────────────
    '  IMPORT / EXPORT — Full Interchangeability Backup
    ' ─────────────────────────────────────────────────────────────────────────

    Private Sub ComboImportExport_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim combo = DirectCast(sender, ComboBox)
        Dim item = TryCast(combo.SelectedItem, ComboBoxItem)
        If item Is Nothing Then Return

        ' Reset immediately so the ComboBox is ready for next use
        combo.SelectedItem = Nothing

        Dim choice = item.Content.ToString().Replace("📥 ", "").Replace("📤 ", "").Trim()

        If choice = "Export" Then
            ExportAllTablesToExcel()
        ElseIf choice = "Import" Then
            ImportAllTablesFromExcel()
        End If
    End Sub

    ' ── Tables included in the backup ───────────────────────────────────────
    Private Function GetFullBackupTableList() As List(Of String)
        Dim tables As New List(Of String)({
            "departmentmaster",
            "department_list",
            "interchangeability",
            "interchange_history",
            "issue",
            "receive",
            "wop",
            "writeoff",
            "reintroduction",
            "temporary_issuance",
            "regular_calibration",
            "new_addition_calibration",
            "reintroduction_calibration",
            "temp_issuance_calibration",
            "result_list",
            "ng_list"
        })

        Try
            ' Discover dynamic calibration tables (tables with CycleName and RowType)
            Dim query = "SELECT TABLE_NAME FROM information_schema.COLUMNS WHERE COLUMN_NAME = 'CycleName' AND TABLE_SCHEMA = DATABASE()"
            Dim dtCycleTables = _mySql.ReadDatatable(query)

            For Each row As DataRow In dtCycleTables.Rows
                Dim tblName = row("TABLE_NAME").ToString()
                ' Ensure it also has RowType
                Dim dtHasRowType = _mySql.ReadDatatable($"SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE COLUMN_NAME = 'RowType' AND TABLE_NAME = '{tblName}' AND TABLE_SCHEMA = DATABASE()")
                If dtHasRowType.Rows.Count > 0 Then
                    If Not tables.Contains(tblName, StringComparer.OrdinalIgnoreCase) Then
                        tables.Add(tblName)
                    End If
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("GetFullBackupTableList Error: " & ex.Message)
        End Try

        Return tables
    End Function

    ' ── Logical key columns used for duplicate-detection on import ───────────
    ' Each entry maps a table name to the column names that together form a
    ' "natural" unique row identifier (id is excluded because it is auto-assigned).
    Private Function GetLogicalKeys(tableName As String, sheetCols As List(Of String)) As String()
        Select Case tableName.ToLower()
            Case "departmentmaster"
                Return New String() {"DepartmentName"}
            Case "department_list"
                Return New String() {"Control No", "CycleName"}
            Case "interchangeability", "interchange_history", "regular_calibration", "new_addition_calibration", "reintroduction_calibration", "temp_issuance_calibration", "result_list"
                Return New String() {"ControlNo", "CycleName"}
            Case "issue", "temporary_issuance"
                Return New String() {"ControlNo", "IssueDate", "CycleName"}
            Case "receive"
                Return New String() {"ControlNo", "ReceiveDate", "CycleName"}
            Case "wop"
                Return New String() {"ControlNo", "WOPDate", "CycleName"}
            Case "writeoff"
                Return New String() {"ControlNo", "WriteOffDate", "CycleName"}
            Case "reintroduction"
                Return New String() {"ControlNo", "ReintroductionDate", "CycleName"}
            Case "ng_list"
                Return New String() {"control_no", "calibrated_date", "CycleName"}
            Case Else
                ' Fallback for dynamic instrument tables
                Dim dynKeys As New List(Of String)()
                For Each potKey In {"ControlNo", "CycleName", "RowType", "Date", "Time"}
                    If sheetCols.Contains(potKey, StringComparer.OrdinalIgnoreCase) Then
                        dynKeys.Add(potKey)
                    End If
                Next
                If dynKeys.Count >= 2 Then
                    Return dynKeys.ToArray()
                End If
                Return New String() {}
        End Select
    End Function

    ' ════════════════════════════════════════════════════════════════════════
    '  EXPORT
    ' ════════════════════════════════════════════════════════════════════════
    Private Async Sub ExportAllTablesToExcel()
        Dim sfd As New Microsoft.Win32.SaveFileDialog()
        sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        sfd.Title = "Export Interchangeability Backup"
        sfd.FileName = "Interchangeability_Backup_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".xlsx"

        If sfd.ShowDialog() <> True Then Return

        Dim filePath = sfd.FileName
        MainProgressBar.Visibility = Visibility.Visible
        MainProgressBar.IsIndeterminate = True

        Dim exportedTables As New List(Of String)()
        Dim failedTables As New List(Of String)()

        Try
            Await Task.Run(Sub()
                               Try
                                   Using wb As New ClosedXML.Excel.XLWorkbook()
                                       Dim allTables = GetFullBackupTableList()
                                       For Each tbl In allTables
                                           Try
                                               Dim dt = _mySql.ReadDatatable($"SELECT * FROM `{tbl}`")
                                               Dim ws = wb.Worksheets.Add(tbl)

                                               ' ── Header row ──────────────────────────────────────
                                               For c As Integer = 0 To dt.Columns.Count - 1
                                                   Dim cell = ws.Cell(1, c + 1)
                                                   cell.Value = dt.Columns(c).ColumnName
                                                   cell.Style.Font.Bold = True
                                                   cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#F8FAFC")
                                                   cell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center
                                               Next

                                               ' ── Data rows ───────────────────────────────────────
                                               For r As Integer = 0 To dt.Rows.Count - 1
                                                   For c As Integer = 0 To dt.Columns.Count - 1
                                                       Dim val = dt.Rows(r)(c)
                                                       Dim cell = ws.Cell(r + 2, c + 1)
                                                       If IsDBNull(val) Then
                                                           cell.Value = ""
                                                       ElseIf TypeOf val Is DateTime Then
                                                           ' Store as ISO string to survive round-trip
                                                           cell.Value = DirectCast(val, DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                                                       ElseIf TypeOf val Is TimeSpan Then
                                                           cell.Value = DirectCast(val, TimeSpan).ToString("hh\:mm\:ss")
                                                       Else
                                                           cell.Value = val.ToString()
                                                       End If
                                                   Next
                                               Next

                                               ' ── Auto-fit columns (max width 60) ─────────────────
                                               ws.Columns().AdjustToContents(8.0, 60.0)

                                               exportedTables.Add($"{tbl} ({dt.Rows.Count} rows)")
                                           Catch exTbl As Exception
                                               failedTables.Add($"{tbl} — {exTbl.Message}")
                                           End Try
                                       Next

                                       wb.SaveAs(filePath)
                                   End Using
                               Catch exSave As Exception
                                   Application.Current.Dispatcher.Invoke(Sub()
                                                                             MessageBox.Show("Export failed: " & exSave.Message,
                                                                                             "Export Error", MessageBoxButton.OK, MessageBoxImage.Error)
                                                                         End Sub)
                               End Try
                           End Sub)

            ' ── Summary message ──────────────────────────────────────────────
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine($"✅ Export Complete!")
            sb.AppendLine($"File: {filePath}")
            sb.AppendLine()
            sb.AppendLine("Tables exported:")
            For Each t In exportedTables
                sb.AppendLine($"  • {t}")
            Next
            If failedTables.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("⚠️ Tables with errors:")
                For Each t In failedTables
                    sb.AppendLine($"  ✗ {t}")
                Next
            End If

            MessageBox.Show(sb.ToString(), "Export Success", MessageBoxButton.OK, MessageBoxImage.Information)

        Catch ex As Exception
            MessageBox.Show("Unexpected export error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            MainProgressBar.Visibility = Visibility.Collapsed
            MainProgressBar.IsIndeterminate = False
        End Try
    End Sub

    ' ════════════════════════════════════════════════════════════════════════
    '  IMPORT
    ' ════════════════════════════════════════════════════════════════════════
    Private Async Sub ImportAllTablesFromExcel()
        Dim ofd As New Microsoft.Win32.OpenFileDialog()
        ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        ofd.Title = "Import Interchangeability Backup"

        If ofd.ShowDialog() <> True Then Return

        ' ── Confirm before writing ───────────────────────────────────────────
        Dim confirm = MessageBox.Show(
            "This will INSERT rows from the selected backup file into all interchangeability tables." & vbCrLf & vbCrLf &
            "• Existing rows (matching their natural key) will be SKIPPED — they will NOT be overwritten." & vbCrLf &
            "• Auto-increment IDs from the file are ignored; the database assigns new IDs." & vbCrLf & vbCrLf &
            "Continue?",
            "Confirm Import", MessageBoxButton.YesNo, MessageBoxImage.Question)

        If confirm <> MessageBoxResult.Yes Then Return

        Dim filePath = ofd.FileName
        MainProgressBar.Visibility = Visibility.Visible
        MainProgressBar.IsIndeterminate = True

        Dim totalInserted As Integer = 0
        Dim totalSkipped As Integer = 0
        Dim tableResults As New List(Of String)()
        Dim errorMessages As New List(Of String)()

        Try
            Await Task.Run(Sub()
                               Try
                                   ' ── Read workbook via ExcelDataReader ───────────────
                                   Dim dataSet As DataSet
                                   Using stream = IO.File.Open(filePath, IO.FileMode.Open, IO.FileAccess.Read)
                                       Using reader = ExcelReaderFactory.CreateReader(stream)
                                           dataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                               .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                   .UseHeaderRow = True
                                               }
                                           })
                                       End Using
                                   End Using

                                   ' ── Process each expected table ──────────────────────
                                   Dim allTables = GetFullBackupTableList()
                                   For Each tblName In allTables
                                       Try
                                           ' Find the matching sheet (case-insensitive)
                                           Dim sheet As DataTable = Nothing
                                           For Each ds As DataTable In dataSet.Tables
                                               If String.Equals(ds.TableName, tblName, StringComparison.OrdinalIgnoreCase) Then
                                                   sheet = ds
                                                   Exit For
                                               End If
                                           Next

                                           If sheet Is Nothing Then
                                               tableResults.Add($"⚠️ {tblName} — sheet not found in file, skipped")
                                               Continue For
                                           End If

                                           If sheet.Rows.Count = 0 Then
                                               tableResults.Add($"  {tblName} — 0 rows in sheet, skipped")
                                               Continue For
                                           End If

                                           ' Get columns that exist in the sheet (excluding 'id')
                                           Dim sheetCols As New List(Of String)()
                                           For Each col As DataColumn In sheet.Columns
                                               If Not col.ColumnName.Equals("id", StringComparison.OrdinalIgnoreCase) Then
                                                   sheetCols.Add(col.ColumnName)
                                               End If
                                           Next

                                           Dim logicalKeys = GetLogicalKeys(tblName, sheetCols)
                                           Dim inserted As Integer = 0
                                           Dim skipped As Integer = 0

                                           ' ── Build existing-key set for fast lookup ────────
                                           Dim existingKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                                           If logicalKeys.Length > 0 Then
                                               Dim selCols = String.Join(", ", logicalKeys.Select(Function(k) $"`{k}`"))
                                               Dim dtExist = _mySql.ReadDatatable($"SELECT {selCols} FROM `{tblName}`")
                                               For Each exRow As DataRow In dtExist.Rows
                                                   Dim keyParts = logicalKeys.Select(Function(k)
                                                                                         Dim colIdx = dtExist.Columns.IndexOf(k)
                                                                                         Return If(colIdx >= 0 AndAlso Not IsDBNull(exRow(colIdx)), exRow(colIdx).ToString().Trim(), "")
                                                                                     End Function)
                                                   existingKeys.Add(String.Join("|", keyParts))
                                               Next
                                           End If

                                           ' ── Insert new rows ──────────────────────────────
                                           For Each row As DataRow In sheet.Rows
                                               Try
                                                   ' Build the natural key string for this row
                                                   Dim rowKey = ""
                                                   If logicalKeys.Length > 0 Then
                                                       Dim keyParts As New List(Of String)()
                                                       Dim allKeyColsPresent = True
                                                       For Each k In logicalKeys
                                                           If sheet.Columns.Contains(k) Then
                                                               keyParts.Add(If(IsDBNull(row(k)), "", row(k).ToString().Trim()))
                                                           Else
                                                               allKeyColsPresent = False
                                                               keyParts.Add("")
                                                           End If
                                                       Next
                                                       If allKeyColsPresent Then
                                                           rowKey = String.Join("|", keyParts)
                                                       End If
                                                   End If

                                                   ' Skip if already exists
                                                   If Not String.IsNullOrEmpty(rowKey) AndAlso existingKeys.Contains(rowKey) Then
                                                       skipped += 1
                                                       Continue For
                                                   End If

                                                   ' Build INSERT statement from columns present in the sheet
                                                   Dim colNames As New List(Of String)()
                                                   Dim colValues As New List(Of String)()
                                                   Dim parameters As New List(Of MySql.Data.MySqlClient.MySqlParameter)()

                                                   For Each colName In sheetCols
                                                       If Not sheet.Columns.Contains(colName) Then Continue For
                                                       Dim rawVal = row(colName)
                                                       Dim strVal = If(IsDBNull(rawVal), "", rawVal.ToString().Trim())

                                                       colNames.Add($"`{colName}`")
                                                       colValues.Add($"@p_{colNames.Count}")
                                                       parameters.Add(New MySql.Data.MySqlClient.MySqlParameter($"@p_{colNames.Count}", If(String.IsNullOrEmpty(strVal), DBNull.Value, CObj(strVal))))
                                                   Next

                                                   If colNames.Count = 0 Then Continue For

                                                   Dim sql = $"INSERT IGNORE INTO `{tblName}` ({String.Join(", ", colNames)}) VALUES ({String.Join(", ", colValues)})"
                                                   Dim success = _mySql.ExecuteNonQuery(sql, parameters.ToArray())

                                                   If success Then
                                                       inserted += 1
                                                       If Not String.IsNullOrEmpty(rowKey) Then existingKeys.Add(rowKey)
                                                   Else
                                                       skipped += 1
                                                   End If
                                               Catch exRow As Exception
                                                   skipped += 1
                                                   Console.WriteLine($"Row insert error [{tblName}]: {exRow.Message}")
                                               End Try
                                           Next

                                           totalInserted += inserted
                                           totalSkipped += skipped
                                           tableResults.Add($"  ✅ {tblName}: {inserted} inserted, {skipped} skipped")

                                       Catch exTbl As Exception
                                           errorMessages.Add($"  ✗ {tblName}: {exTbl.Message}")
                                       End Try
                                   Next

                               Catch exFile As Exception
                                   Application.Current.Dispatcher.Invoke(Sub()
                                                                             MessageBox.Show("Failed to read file: " & exFile.Message,
                                                                                             "Import Error", MessageBoxButton.OK, MessageBoxImage.Error)
                                                                         End Sub)
                               End Try
                           End Sub)

            ' ── Summary ──────────────────────────────────────────────────────
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("📥 Import Complete!")
            sb.AppendLine()
            sb.AppendLine($"Total rows inserted : {totalInserted}")
            sb.AppendLine($"Total rows skipped  : {totalSkipped} (duplicates / errors)")
            sb.AppendLine()
            sb.AppendLine("Per-table results:")
            For Each tblResult In tableResults
                sb.AppendLine(tblResult)
            Next
            If errorMessages.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("Errors:")
                For Each errMsg In errorMessages
                    sb.AppendLine(errMsg)
                Next
            End If

            MessageBox.Show(sb.ToString(), "Import Summary", MessageBoxButton.OK, MessageBoxImage.Information)

            ' Refresh UI to reflect any newly imported data
            Dim activeCycle = GetCurrentCycleName()
            LoadDepartments()
            LoadDeptListCycleDropdown(activeCycle)
            LoadData()
            If InterchangeTabView.Visibility = Visibility.Visible Then LoadInterchangeData()
            If InterHistoryTabView.Visibility = Visibility.Visible Then LoadHistoryData()

        Catch ex As Exception
            MessageBox.Show("Unexpected import error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            MainProgressBar.Visibility = Visibility.Collapsed
            MainProgressBar.IsIndeterminate = False
        End Try
    End Sub

    Private Sub OpenChangeRFIDWindowForRow()
        Dim currentRfid = ""
        Dim rowView = TryCast(InterchangeDataGrid.SelectedItem, DataRowView)
        If rowView IsNot Nothing AndAlso rowView.DataView.Table.Columns.Contains("RFID_tag") Then
            currentRfid = rowView("RFID_tag").ToString()
        End If

        Dim result = MessageBox.Show(
            $"The item {_currentContextControlNo} already has an RFID tag assigned ({currentRfid})." & vbCrLf &
            "Are you sure you want to overwrite the existing RFID tag?",
            "Confirm RFID Overwrite",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning)

        If result <> MessageBoxResult.Yes Then
            Return
        End If

        Dim selectedCycle = ComboCycle.SelectedItem.ToString()
        Dim rfidWin As New ReceiveRFIDWindow(_currentContextControlNo)
        rfidWin.Title = "Change RFID Tag"
        If rfidWin.ShowDialog() = True Then
            If _mySql.UpdateLinkedRFID(_currentContextControlNo, selectedCycle, rfidWin.RfidTag) Then
                ' Only the RFID tag was updated. No new logs were created.
            End If
            LoadInterchangeData()
            EvaluateActionButtons()
        End If
    End Sub

    Private Sub DeptListDataGrid_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles DeptListDataGrid.PreviewKeyDown
        If e.Key = Key.Enter Then
            Dim grid = DirectCast(sender, DataGrid)
            If grid.CurrentCell.IsValid AndAlso grid.CurrentCell.Column IsNot Nothing Then
                Dim headerStr = grid.CurrentCell.Column.Header?.ToString()
                If headerStr = "Size & Range" Then
                    Dim row = TryCast(grid.CurrentCell.Item, DataRowView)
                    If row IsNot Nothing Then
                        Dim ctrlNo = If(row.DataView.Table.Columns.Contains("Control No"), row("Control No").ToString(), row("ControlNo").ToString())
                        If Not String.IsNullOrEmpty(ctrlNo) Then
                            e.Handled = True
                            OpenSizeDetailsForControl(ctrlNo)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub InterchangeDataGrid_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles InterchangeDataGrid.PreviewKeyDown
        If e.Key = Key.Enter Then
            Dim grid = DirectCast(sender, DataGrid)
            If grid.CurrentCell.IsValid AndAlso grid.CurrentCell.Column IsNot Nothing Then
                Dim headerStr = grid.CurrentCell.Column.Header?.ToString()
                Dim row = TryCast(grid.CurrentCell.Item, DataRowView)
                If row IsNot Nothing Then
                    Dim ctrlNo = If(row.DataView.Table.Columns.Contains("Control No"), row("Control No").ToString(), row("ControlNo").ToString())

                    If headerStr = "Size & Range" Then
                        If Not String.IsNullOrEmpty(ctrlNo) Then
                            e.Handled = True
                            OpenSizeDetailsForControl(ctrlNo)
                        End If
                    ElseIf headerStr = "Control No" Then
                        If Not String.IsNullOrEmpty(ctrlNo) Then
                            e.Handled = True
                            OpenHistoryForControl(ctrlNo)
                        End If
                    Else
                        ' Trigger DoubleClick logic for other cells
                        e.Handled = True
                        InterchangeDataGrid_MouseDoubleClick(sender, Nothing)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub InterHistoryDataGrid_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles InterHistoryDataGrid.PreviewKeyDown
        If e.Key = Key.Enter Then
            Dim grid = DirectCast(sender, DataGrid)
            If grid.CurrentCell.IsValid AndAlso grid.CurrentCell.Column IsNot Nothing Then
                Dim headerStr = grid.CurrentCell.Column.Header?.ToString()
                Dim row = TryCast(grid.CurrentCell.Item, DataRowView)
                If row IsNot Nothing Then
                    If headerStr = "Control No" Then
                        Dim ctrlNo = If(row.DataView.Table.Columns.Contains("Control No"), row("Control No").ToString(), row("ControlNo").ToString())
                        If Not String.IsNullOrEmpty(ctrlNo) Then
                            e.Handled = True
                            OpenHistoryForControl(ctrlNo)
                        End If
                    Else
                        ' Try triggering the MouseUp logic which checks if it's a valid cycle cell
                        e.Handled = True
                        InterHistoryDataGrid_MouseUp(sender, Nothing)
                    End If
                End If
            End If
        End If
    End Sub

End Class
