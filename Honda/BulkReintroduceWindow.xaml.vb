Imports System.Collections.ObjectModel
Imports System.Data
Imports Microsoft.Win32
Imports System.IO

Public Class BulkReintroduceWindow
    Private _mySql As New MySQLClass()
    Private _addedItems As New ObservableCollection(Of ReintroItem)()
    Private _currentSearchItem As ReintroItem = Nothing
    Private _uploadedFilePath As String = ""
    Private _activeCycle As String = ""

    Public Sub New()
        InitializeComponent()
        GridAddedItems.ItemsSource = _addedItems
        _activeCycle = _mySql.GetActiveCycleName()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        LoadDepartments()
    End Sub

    Private Sub LoadDepartments()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT DepartmentName FROM departmentmaster ORDER BY DepartmentName")
            ComboDept.ItemsSource = dt.DefaultView
        Catch ex As Exception
            Console.WriteLine("Error loading departments: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnSearch_Click(sender As Object, e As RoutedEventArgs)
        Dim ctrlNo = TxtSearchControlNo.Text.Trim()
        If String.IsNullOrEmpty(ctrlNo) Then Return

        ' Reset state
        _currentSearchItem = Nothing
        PanelDeptSelect.Visibility = Visibility.Collapsed
        BtnAdd.IsEnabled = False
        TxtItemMeta.Text = "Searching..."
        TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#64748B"), Color))

        Try
            ' 1. Check duplicate in current list
            If _addedItems.Any(Function(x) x.ControlNo.Equals(ctrlNo, StringComparison.OrdinalIgnoreCase)) Then
                TxtItemMeta.Text = "Item already in the list."
                TxtItemMeta.Foreground = Brushes.Red
                Return
            End If

            ' 2. Search Master Tables FIRST (we ALWAYS need this for base metadata: TableName, InstrumentType, etc.)
            Dim foundMaster = SearchMasterTables(ctrlNo)
            If foundMaster Is Nothing Then
                TxtItemMeta.Text = "Control Number not found in Master List."
                TxtItemMeta.Foreground = Brushes.Red
                Return
            End If
            
            _currentSearchItem = foundMaster

            ' 3. Check Current Cycle
            Dim queryIC = $"SELECT Status, Department, Remarks FROM interchangeability WHERE ControlNo = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{_activeCycle.Replace("'", "''")}'"
            Dim dtIC = _mySql.ReadDatatable(queryIC)
            
            Dim foundInCurrentCycle As Boolean = False
            If dtIC.Rows.Count > 0 Then
                foundInCurrentCycle = True
                Dim status = dtIC.Rows(0)("Status").ToString()
                Dim remarks = dtIC.Rows(0)("Remarks").ToString()

                If Not status.Equals("WOP", StringComparison.OrdinalIgnoreCase) AndAlso Not status.Equals("Write off", StringComparison.OrdinalIgnoreCase) Then
                    TxtItemMeta.Text = $"Item is in current cycle with status '{status}'. Reintroduction requires 'WOP' or 'Write off'."
                    TxtItemMeta.Foreground = Brushes.Red
                    _currentSearchItem = Nothing
                    Return
                End If

                ' Special check for Calibration NG
                If remarks.Contains("Calibration NG") Then
                    _currentSearchItem.IsNewToInterchange = False
                    Dim dept = dtIC.Rows(0)("Department").ToString()
                    If Not String.IsNullOrEmpty(dept) Then _currentSearchItem.Department = dept
                    
                    TxtItemMeta.Text = $"Found Calibration NG: {_currentSearchItem.InstrumentName} ({_currentSearchItem.Department}). Reintroduction NOT possible."
                    TxtItemMeta.Foreground = Brushes.Red
                    BtnAdd.IsEnabled = False
                    Return
                End If

                ' It is WOP or Write off in current cycle!
                _currentSearchItem.IsNewToInterchange = False
                Dim deptVal = dtIC.Rows(0)("Department").ToString()
                If Not String.IsNullOrEmpty(deptVal) Then
                    _currentSearchItem.Department = deptVal ' Keep its latest department
                End If
            End If

            ' 4. Check Write-off history (to get MissingDate and WriteOffNo)
            Dim queryWO = $"SELECT w.*, COALESCE(w.WriteOffNo, '') as WO_No, COALESCE(w.WriteOffDate, CURDATE()) as WO_Date " &
                         $"FROM writeoff w WHERE w.ControlNo = '{ctrlNo.Replace("'", "''")}' ORDER BY w.WriteOffDate DESC LIMIT 1"
            Dim dtWO = _mySql.ReadDatatable(queryWO)

            If dtWO.Rows.Count > 0 Then
                Dim rowWO = dtWO.Rows(0)
                _currentSearchItem.WriteOffNo = rowWO("WO_No").ToString()
                _currentSearchItem.MissingDate = If(IsDBNull(rowWO("WO_Date")), DateTime.Today, Convert.ToDateTime(rowWO("WO_Date")))
                _currentSearchItem.IsNewToInterchange = False
                
                If foundInCurrentCycle Then
                    TxtItemMeta.Text = $"Found in Current Cycle & Write-off: {_currentSearchItem.InstrumentName} ({_currentSearchItem.Department}) - {_currentSearchItem.Color}"
                    TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#3B82F6"), Color)) ' Blue
                Else
                    TxtItemMeta.Text = $"Found in Write-off: {_currentSearchItem.InstrumentName} ({_currentSearchItem.Department}) - {_currentSearchItem.Color}"
                    TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#3B82F6"), Color)) ' Blue
                End If
                PanelDeptSelect.Visibility = Visibility.Collapsed
            Else
                ' No Write-off history
                _currentSearchItem.WriteOffNo = ""
                _currentSearchItem.MissingDate = DateTime.Today

                If foundInCurrentCycle Then
                    ' It's in current cycle as WOP (since it has no writeoff history and passed the status check)
                    _currentSearchItem.IsNewToInterchange = False
                    TxtItemMeta.Text = $"Found in Current Cycle (WOP): {_currentSearchItem.InstrumentName} ({_currentSearchItem.Department}) - {_currentSearchItem.Color}"
                    TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#3B82F6"), Color)) ' Blue
                    PanelDeptSelect.Visibility = Visibility.Collapsed
                Else
                    ' Not in current cycle, not in write off -> completely new
                    _currentSearchItem.IsNewToInterchange = True
                    TxtItemMeta.Text = $"Found in Master List (New to Interchange): {_currentSearchItem.InstrumentName} - {_currentSearchItem.Color}. Please select target Department."
                    TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#10B981"), Color)) ' Green
                    PanelDeptSelect.Visibility = Visibility.Visible
                    If Not String.IsNullOrEmpty(_currentSearchItem.Department) Then
                         ComboDept.SelectedValue = _currentSearchItem.Department
                    End If
                End If
            End If

            BtnAdd.IsEnabled = True

        Catch ex As Exception
            TxtItemMeta.Text = "Error during search: " & ex.Message
            TxtItemMeta.Foreground = Brushes.Red
        End Try
    End Sub

    Private Function SearchMasterTables(ctrlNo As String) As ReintroItem
        Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName, Category FROM type_details")
        For Each r As DataRow In dtSettings.Rows
            Dim tbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString(), forInventory:=True)
            Dim query = $"SELECT * FROM `{tbl}` WHERE ControlNo = '{ctrlNo.Replace("'", "''")}' LIMIT 1"
            Dim dt = _mySql.ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                Dim item As New ReintroItem()
                item.ControlNo = ctrlNo
                item.TableName = tbl
                item.InstrumentType = r("Category").ToString()
                item.Color = row("Color").ToString()
                item.Department = row("Location").ToString() ' Original dept
                
                If item.InstrumentType.Equals("Instrument", StringComparison.OrdinalIgnoreCase) Then
                    item.InstrumentName = row("InstrumentName").ToString()
                    item.SizeandRange = row("Size").ToString()
                Else
                    item.InstrumentName = row("GaugeName").ToString()
                    item.SizeandRange = row("Size").ToString()
                End If
                Return item
            End If
        Next
        Return Nothing
    End Function

    Private Function FindMasterTable(ctrlNo As String) As String
        Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName FROM type_details")
        For Each r As DataRow In dtSettings.Rows
            Dim tbl = MySQLClass.TypeNameToTableName(r(0).ToString(), forInventory:=True)
            Dim chk = _mySql.ReadDatatable($"SELECT ControlNo FROM `{tbl}` WHERE ControlNo='{ctrlNo.Replace("'", "''")}'")
            If chk.Rows.Count > 0 Then Return tbl
        Next
        Return "unknown"
    End Function

    Private Sub BtnAdd_Click(sender As Object, e As RoutedEventArgs)
        If _currentSearchItem Is Nothing Then Return

        ' If new to interchange, require department selection
        If _currentSearchItem.IsNewToInterchange Then
            If ComboDept.SelectedValue Is Nothing Then
                MessageBox.Show("Please select a department for this item.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
            _currentSearchItem.Department = ComboDept.SelectedValue.ToString()
        End If

        _addedItems.Add(_currentSearchItem)
        
        ' Reset UI
        TxtSearchControlNo.Clear()
        TxtItemMeta.Text = "Enter a control number and press search..."
        TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#64748B"), Color))
        PanelDeptSelect.Visibility = Visibility.Collapsed
        BtnAdd.IsEnabled = False
        _currentSearchItem = Nothing
        UpdateSelectionCount()
    End Sub

    Private Sub BtnRemoveItem_Click(sender As Object, e As RoutedEventArgs)
        Dim item = DirectCast(DirectCast(sender, Button).DataContext, ReintroItem)
        _addedItems.Remove(item)
        UpdateSelectionCount()
    End Sub

    Private Sub UpdateSelectionCount()
        lblSelectionCount.Text = $"Items in Queue: {_addedItems.Count}"
        BtnSubmit.IsEnabled = (_addedItems.Count > 0 AndAlso Not String.IsNullOrEmpty(_uploadedFilePath))
        lblSelectionCount.Foreground = If(_addedItems.Count > 0, New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#3B82F6"), Color)), Brushes.Gray)
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As RoutedEventArgs)
        Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
        openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
        If openFileDialog.ShowDialog() = True Then
            _uploadedFilePath = openFileDialog.FileName
            TxtFileName.Text = Path.GetFileName(_uploadedFilePath)
            UpdateSelectionCount()
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnSubmit_Click(sender As Object, e As RoutedEventArgs)
        If _addedItems.Count = 0 Then Return
        
        If String.IsNullOrEmpty(_uploadedFilePath) Then
            MessageBox.Show("Please upload a reintroduction PDF document.", "Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim result = MessageBox.Show($"Are you sure you want to process reintroduction for {_addedItems.Count} items?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result <> MessageBoxResult.Yes Then Return

        Try
            Dim successCount = 0
            Dim batchTag = "Reintro " & DateTime.Now.ToString("yyyyMMdd")
            Dim remark = TxtRemark.Text.Trim()
            Dim currentUser = Application.Current.Properties("Username")?.ToString()

            For Each item In _addedItems
                ' Copy file for each control no (as done in single reintro logic)
                Dim copiedPath = MySQLClass.CopyFileToDocuments(_uploadedFilePath, "Reintroduction", item.ControlNo)
                
                Dim emptyVal = ""
                Dim reintroSuccess = _mySql.InsertReintroduction(
                    DateTime.Today,
                    item.InstrumentType,
                    item.InstrumentName,
                    item.ControlNo,
                    "1", ' Quantity
                    "", ' Line (could be empty or from master)
                    item.Department,
                    item.Color,
                    remark,
                    item.WriteOffNo,
                    $"RE-{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    DateTime.Today,
                    item.MissingDate,
                    emptyVal, emptyVal, emptyVal, emptyVal, emptyVal,
                    currentUser,
                    emptyVal,
                    copiedPath,
                    _activeCycle
                )

                If item.IsNewToInterchange Then
                    ' 1. Auto Write-off for Master List items
                    Dim woNo = "WO-AUTO-" & DateTime.Now.ToString("yyyyMMddHHmmss")
                    _mySql.InsertWriteOff(
                        DateTime.Today, DateTime.Today,
                        item.InstrumentType, item.InstrumentName, item.ControlNo,
                        "1", "", item.Department, item.Color,
                        remark, "Write Off", "NG", copiedPath, woNo, "Write off",
                        currentUser, "1", _activeCycle
                    )
                    
                    ' Log to history for Write off
                    _mySql.LogInterchangeHistory(_activeCycle, item.ControlNo, item.Department, item.InstrumentName, item.Color, "Write off", remark)
                End If

                If reintroSuccess Then
                    ' 1. Update master table status to Active
                    If item.TableName <> "unknown" Then
                        _mySql.UpdateRecordStatus(item.TableName, item.ControlNo, 0, "Active")
                    End If

                    ' 2. Sync to department_list
                    _mySql.InsertDeptListItem(item.Department, item.InstrumentName, item.SizeandRange, item.ControlNo, item.Color, "Pending", remark, batchTag, _activeCycle)

                    ' 3. Sync to interchangeability
                    _mySql.InsertInterchangeRecord(_activeCycle, item.ControlNo, item.Department, item.InstrumentName, item.SizeandRange, item.Color, "Pending", remark)

                    ' NOTE: reintroduction_calibration entry is created later when the item is Received
                    '       (InsertInterchangeRecord detects the reintroduction row and routes accordingly)

                    ' 4. Register in History
                    _mySql.LogInterchangeHistory(_activeCycle, item.ControlNo, item.Department, item.InstrumentName, item.Color, "Pending(Reintroduced)", remark)

                    ' 5. Mark as having history to ensure visibility in grid
                    _mySql.ExecuteNonQuery($"UPDATE interchangeability SET HasHistory = 1 WHERE ControlNo = '{item.ControlNo.Replace("'", "''")}' AND CycleName = '{_activeCycle.Replace("'", "''")}'")

                    successCount += 1
                End If
            Next

            MessageBox.Show($"Successfully reintroduced {successCount} of {_addedItems.Count} items.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Error processing bulk reintroduction: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        GridAddedItems.SelectedItem = Nothing
    End Sub
End Class

Public Class ReintroItem
    Public Property ControlNo As String
    Public Property InstrumentName As String
    Public Property Department As String
    Public Property InstrumentType As String
    Public Property SizeandRange As String
    Public Property Color As String
    Public Property TableName As String ' Master Table Name
    Public Property WriteOffNo As String
    Public Property MissingDate As DateTime
    Public Property IsNewToInterchange As Boolean
End Class
