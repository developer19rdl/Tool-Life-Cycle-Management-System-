Imports System.Data
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class BulkWOPWindow
    Private _mySql As New MySQLClass()
    Private _activeCycle As String = ""
    Private _allRecords As New List(Of ItemData)()
    Private _recordsMap As New Dictionary(Of String, ItemData)(StringComparer.OrdinalIgnoreCase)
    Private _addedItems As New ObservableCollection(Of ItemData)()
    Private _currentlySearchedItem As ItemData = Nothing

    Public Class ItemData
        Implements INotifyPropertyChanged
        Public Property ControlNo As String
        Public Property Department As String
        Public Property InstrumentName As String
        Public Property Color As String
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
                                          Dim query = $"SELECT *, SizeandRange FROM interchangeability WHERE CycleName = '{activeCycle.Replace("'", "''")}' AND Status NOT IN ('Write off', 'WOP') ORDER BY ControlNo"
                                          Dim dt = _mySql.ReadDatatable(query)

                                          Dim records As New List(Of ItemData)()
                                          For Each row As DataRow In dt.Rows
                                              Dim item As New ItemData With {
                                                  .ControlNo = row("ControlNo").ToString(),
                                                  .Department = row("Department").ToString(),
                                                  .InstrumentName = row("InstrumentName").ToString(),
                                                  .Color = row("Color").ToString(),
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

        If _recordsMap.ContainsKey(searchKey) Then
            _currentlySearchedItem = _recordsMap(searchKey)
            Dim status = _currentlySearchedItem.FullRow("Status").ToString()
            TxtItemMeta.Text = $"Found: {_currentlySearchedItem.InstrumentName} - Color: {_currentlySearchedItem.Color} - Status: {status}"
            TxtItemMeta.Foreground = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#3B82F6"), Color)) ' Blue for found
            BtnAdd.IsEnabled = True
        Else
            _currentlySearchedItem = Nothing
            TxtItemMeta.Text = "Control number not found, written off, or already marked as WOP."
            TxtItemMeta.Foreground = Brushes.Red
            BtnAdd.IsEnabled = False
        End If
    End Sub

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
        BtnSubmit.IsEnabled = count > 0
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Async Sub BtnSubmit_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedItems = _addedItems.ToList()
        If selectedItems.Count = 0 Then Return

        Dim confirm = MessageBox.Show($"Are you sure you want to process WOP for {selectedItems.Count} items?", "Confirm Bulk WOP", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        If confirm <> MessageBoxResult.Yes Then Return

        Try
            Me.IsEnabled = False
            Mouse.OverrideCursor = Cursors.Wait
            
            Dim activeCycle = _activeCycle
            Dim remark = TxtRemark.Text

            Dim result = Await Task.Run(Function()
                                            Dim sCount = 0
                                            
                                            For Each item In selectedItems
                                                Dim success = _mySql.InsertWOPRecord(item.ControlNo, activeCycle, item.Department, item.InstrumentName, item.Color, remark)
                                                If success Then
                                                    _mySql.InsertInterchangeRecord(activeCycle, item.ControlNo, item.Department, item.InstrumentName, item.Size, item.Color, "WOP", remark)
                                                    _mySql.ClearRFIDTag(item.ControlNo) ' Clear RFID when marked as WOP
                                                    sCount += 1
                                                End If
                                            Next
                                            Return sCount
                                        End Function)


            MessageBox.Show($"Successfully processed WOP for {result} items.", "Bulk Complete", MessageBoxButton.OK, MessageBoxImage.Information)
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
