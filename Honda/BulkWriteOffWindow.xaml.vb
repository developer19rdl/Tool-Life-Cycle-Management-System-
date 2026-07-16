Imports System.Data
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class BulkWriteOffWindow
    Private _mySql As New MySQLClass()
    Private _activeCycle As String = ""
    Private _allRecords As New List(Of ItemData)()
    Private _recordsMap As New Dictionary(Of String, ItemData)(StringComparer.OrdinalIgnoreCase)
    Private _addedItems As New ObservableCollection(Of ItemData)()
    Private _currentlySearchedItem As ItemData = Nothing
    Private _uploadedFilePath As String = ""

    Public Class ItemData
        Implements INotifyPropertyChanged
        Public Property ControlNo As String
        Public Property Department As String
        Public Property InstrumentName As String
        Public Property Color As String
        Public Property Type As String ' Instrument or Gauge
        Public Property Size As String
        Public Property FullRow As DataRow

        Private _isChecked As Boolean
        Public Property IsChecked As Boolean
            Get
                Return _isChecked
            End Get
            Set(ByVal value As Boolean)
                _isChecked = value
                OnPropertyChanged("IsChecked")
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Protected Sub OnPropertyChanged(name As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
        End Sub
    End Class

    Public Sub New()
        InitializeComponent()
        GridAddedItems.ItemsSource = _addedItems
    End Sub


    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        _activeCycle = GetActiveCycle()
        LoadData()
    End Sub

    Private Function GetActiveCycle() As String
        ' Delegate to MySQLClass which reads Admin-configured date ranges + ActiveCycle override
        Return _mySql.GetActiveCycleName()
    End Function


    Private Async Sub LoadData()
        Try
            BtnSubmit.IsEnabled = False
            Dim activeCycle = _activeCycle ' Capture for background thread

            Dim data = Await Task.Run(Function()
                                          Dim query = $"SELECT *, SizeandRange FROM interchangeability WHERE CycleName = '{activeCycle.Replace("'", "''")}' AND Status <> 'Write off' ORDER BY ControlNo"
                                          Dim dt = _mySql.ReadDatatable(query)

                                          ' Fetch type mapping once
                                          Dim dtTypes = _mySql.ReadDatatable("SELECT TypeName, Category FROM type_details")
                                          Dim typeMap = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                                          For Each r As DataRow In dtTypes.Rows
                                              typeMap(r("TypeName").ToString()) = r("Category").ToString()
                                          Next

                                          Dim records As New List(Of ItemData)()
                                          For Each row As DataRow In dt.Rows
                                              Dim name = row("InstrumentName").ToString()
                                              Dim itemType = "Unknown"
                                              If typeMap.ContainsKey(name) Then itemType = typeMap(name)

                                              Dim item As New ItemData With {
                                                  .ControlNo = row("ControlNo").ToString(),
                                                  .Department = row("Department").ToString(),
                                                  .InstrumentName = name,
                                                  .Color = row("Color").ToString(),
                                                  .Type = itemType,
                                                  .Size = row("SizeandRange").ToString(),
                                                  .FullRow = row,
                                                  .IsChecked = False
                                              }
                                              records.Add(item)
                                          Next
                                          Return records
                                      End Function)

            _allRecords.Clear()
            _recordsMap.Clear()
            
            For Each item In data
                _allRecords.Add(item)
                _recordsMap(item.ControlNo) = item
            Next

            UpdateSelectionCount()
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnSearch_Click(sender As Object, e As RoutedEventArgs)
        Dim searchKey = TxtSearchControlNo.Text.Trim()
        If String.IsNullOrEmpty(searchKey) Then Return

        ' Reset state
        _currentlySearchedItem = Nothing
        BtnAdd.IsEnabled = False
        TxtItemMeta.Text = "Searching..."
        TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#64748B"), Color))

        ' 1. Check if already in the added list
        If _addedItems.Any(Function(x) x.ControlNo.Equals(searchKey, StringComparison.OrdinalIgnoreCase)) Then
            TxtItemMeta.Text = "Item already in the list."
            TxtItemMeta.Foreground = Brushes.Red
            Return
        End If

        ' 2. Check Interchangeability (Current Cycle)
        If _recordsMap.ContainsKey(searchKey) Then
            _currentlySearchedItem = _recordsMap(searchKey)
            Dim status = _currentlySearchedItem.FullRow("Status").ToString()
            TxtItemMeta.Text = $"Found in Interchange: {_currentlySearchedItem.InstrumentName} - Color: {_currentlySearchedItem.Color} - Status: {status}"
            TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#3B82F6"), Color)) ' Blue
            BtnAdd.IsEnabled = True
            Return
        End If

        ' 3. Check Master Tables if not in interchange
        Try
            Dim masterItem = SearchMasterTables(searchKey)
            If masterItem IsNot Nothing Then
                If masterItem.FullRow IsNot Nothing AndAlso masterRowHasStatus(masterItem.FullRow, "Write off") Then
                    TxtItemMeta.Text = "Item is already marked as 'Write off' in Master List."
                    TxtItemMeta.Foreground = Brushes.Red
                Else
                    _currentlySearchedItem = masterItem
                    TxtItemMeta.Text = $"Found in Master List: {_currentlySearchedItem.InstrumentName} ({_currentlySearchedItem.Department}) - {_currentlySearchedItem.Type}"
                    TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#10B981"), Color)) ' Green for master find
                    BtnAdd.IsEnabled = True
                End If
            Else
                TxtItemMeta.Text = "Control number not found in Interchange or Master List."
                TxtItemMeta.Foreground = Brushes.Red
            End If
        Catch ex As Exception
            TxtItemMeta.Text = "Error during master search: " & ex.Message
            TxtItemMeta.Foreground = Brushes.Red
        End Try
    End Sub

    Private Function masterRowHasStatus(row As DataRow, targetStatus As String) As Boolean
        If row.Table.Columns.Contains("Status") Then
            Return row("Status").ToString().Equals(targetStatus, StringComparison.OrdinalIgnoreCase)
        End If
        Return False
    End Function

    Private Function SearchMasterTables(ctrlNo As String) As ItemData
        Dim dtSettings = _mySql.ReadDatatable("SELECT TypeName, Category FROM type_details")
        For Each r As DataRow In dtSettings.Rows
            Dim tbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString())
            Dim query = $"SELECT * FROM `{tbl}` WHERE ControlNo = '{ctrlNo.Replace("'", "''")}' LIMIT 1"
            Dim dt = _mySql.ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                Dim item As New ItemData()
                item.ControlNo = ctrlNo
                item.InstrumentName = If(r("Category").ToString().Equals("Instrument", StringComparison.OrdinalIgnoreCase), row("InstrumentName").ToString(), row("GaugeName").ToString())
                item.Department = row("Location").ToString()
                item.Color = row("Color").ToString()
                item.Type = r("Category").ToString()
                item.Size = row("Size").ToString()
                item.FullRow = row
                item.IsChecked = False
                Return item
            End If
        Next
        Return Nothing
    End Function

    Private Sub BtnAdd_Click(sender As Object, e As RoutedEventArgs)
        If _currentlySearchedItem Is Nothing Then Return

        If _addedItems.Contains(_currentlySearchedItem) Then
            MessageBox.Show("Item already in the list.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        _addedItems.Add(_currentlySearchedItem)
        TxtSearchControlNo.Clear()
        TxtItemMeta.Text = "Item added. Search for another..."
        TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#64748B"), Color))
        _currentlySearchedItem = Nothing
        BtnAdd.IsEnabled = False
        UpdateSelectionCount()
    End Sub

    Private Sub BtnRemoveItem_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim item = DirectCast(btn.DataContext, ItemData)
        If item IsNot Nothing Then
            _addedItems.Remove(item)
            UpdateSelectionCount()
        End If
    End Sub






    Private Sub UpdateSelectionCount()
        Dim count = _addedItems.Count
        lblSelectionCount.Text = $"Items Selected: {count}"
        BtnSubmit.IsEnabled = (count > 0 AndAlso Not String.IsNullOrEmpty(_uploadedFilePath))
    End Sub


    Private Sub BtnBrowse_Click(sender As Object, e As RoutedEventArgs)
        Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
        openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
        If openFileDialog.ShowDialog() = True Then
            Dim source = openFileDialog.FileName
            ' Use "Bulk" as identifier for the centralized helper
            _uploadedFilePath = MySQLClass.CopyFileToDocuments(source, "WriteOff", "Bulk")
            
            If Not String.IsNullOrEmpty(_uploadedFilePath) Then
                TxtFileName.Text = IO.Path.GetFileName(source)
                UpdateSelectionCount()
            Else
                MessageBox.Show("Failed to copy file to Documents folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Async Sub BtnSubmit_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedItems = _addedItems.ToList()
        If selectedItems.Count = 0 Then Return

        Dim confirm = MessageBox.Show($"Are you sure you want to Write Off {selectedItems.Count} items?", "Confirm Bulk Write Off", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        If confirm <> MessageBoxResult.Yes Then Return

        Try
            Me.IsEnabled = False
            Mouse.OverrideCursor = Cursors.Wait
            
            Dim woNo = "WO-" & DateTime.Now.ToString("yyyyMMddHHmmss")
            Dim raisedBy = Application.Current.Properties("Username")?.ToString()
            Dim successCount = 0

            ' Run heavy processing in background
            Dim count = selectedItems.Count
            Dim uploadedPath = _uploadedFilePath
            Dim activeCycle = _activeCycle
            Dim remark = TxtRemark.Text

            Dim result = Await Task.Run(Function()
                                            Dim sCount = 0
                                            
                                            ' Optimization: Pre-map ControlNos to TableNames
                                            Dim controlTableMap As New Dictionary(Of String, String)()
                                            Dim dtTypes = _mySql.ReadDatatable("SELECT TypeName FROM type_details")
                                            For Each tr As DataRow In dtTypes.Rows
                                                Dim tName = MySQLClass.TypeNameToTableName(tr("TypeName").ToString())
                                                Dim dtIds = _mySql.ReadDatatable($"SELECT ControlNo FROM `{tName}`")
                                                For Each dr As DataRow In dtIds.Rows
                                                    controlTableMap(dr("ControlNo").ToString()) = tName
                                                Next
                                            Next

                                            For Each item In selectedItems
                                                ' 1. Insert into Write Off Table
                                                Dim success = _mySql.InsertWriteOff(
                                                    DateTime.Now,
                                                    DateTime.Now,
                                                    item.Type,
                                                    item.InstrumentName,
                                                    item.ControlNo,
                                                    "1",
                                                    "",
                                                    item.Department,
                                                    item.Color,
                                                    remark,
                                                    "Write Off",
                                                    "NG",
                                                    uploadedPath,
                                                    woNo,
                                                    "Write off",
                                                    raisedBy,
                                                    "1",
                                                    activeCycle
                                                )

                                                If success Then
                                                    ' 2. Update Interchangeability Status
                                                    _mySql.InsertInterchangeRecord(activeCycle, item.ControlNo, item.Department, item.InstrumentName, item.Size, item.Color, "Write off", remark)
                                                    _mySql.ClearRFIDTag(item.ControlNo) ' Clear RFID when written off
                                                    
                                                    ' Populate NG List for Bulk Write Off
                                                    Dim ngReason = If(String.IsNullOrWhiteSpace(remark), "Writeoff NG", remark)
                                                    _mySql.InsertNGRecord(item.ControlNo, item.InstrumentName, item.Department, activeCycle, ngReason, DateTime.Today, DateTime.Today, "Write off")

                                                    ' 3. Update Master List Table Status using pre-mapped table
                                                    If controlTableMap.ContainsKey(item.ControlNo) Then
                                                        Dim tblName = controlTableMap(item.ControlNo)
                                                        _mySql.UpdateRecordStatus(tblName, item.ControlNo, 1, "Write off")
                                                    End If

                                                    ' 4. Remove from department_list for this cycle
                                                    _mySql.ExecuteNonQuery($"DELETE FROM department_list WHERE `Control No` = '{item.ControlNo.Replace("'", "''")}' AND CycleName = '{activeCycle.Replace("'", "''")}'")
                                                    
                                                    sCount += 1
                                                End If
                                            Next
                                            Return sCount
                                        End Function)


            MessageBox.Show($"Successfully processed Write Off for {result} items.", "Bulk Complete", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Bulk Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            Mouse.OverrideCursor = Nothing
            Me.IsEnabled = True
        End Try
    End Sub

    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        GridAddedItems.SelectedItem = Nothing
    End Sub
End Class
