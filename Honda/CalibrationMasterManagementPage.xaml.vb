Imports System.Windows
Imports System.Windows.Controls
Imports System.Data
Imports Microsoft.Win32
Imports System.IO
Imports ExcelDataReader

Public Class CalibrationMasterManagementPage
    Inherits Page

    Private _mySql As New MySQLClass()
    Private _masterName As String
    Private _tableName As String
    Private _currentPage As Integer = 1
    Private _pageSize As Integer = 20
    Private _totalCount As Integer = 0
    Private _isLoading As Boolean = False

    Public Sub New(masterName As String)
        InitializeComponent()
        _masterName = masterName
        _tableName = MySQLClass.TypeNameToTableName(masterName)
        
        ' UI Updates
        TxtHeaderTitle.Text = $"{_masterName} Masterlist"
    End Sub

    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        LoadData()
    End Sub

    Private Sub LoadData()
        If _isLoading Then Return
        _isLoading = True
        MainProgressBar.Visibility = Visibility.Visible

        Try
            Dim keyword As String = TxtSearch.Text
            _totalCount = _mySql.GetCalibrationMasterTotalCount(_tableName, keyword)

            Dim offset = (_currentPage - 1) * _pageSize
            Dim dt = _mySql.GetCalibrationMasterData(_tableName, offset, _pageSize, keyword)

            MasterDataGrid.ItemsSource = dt.DefaultView
            UpdatePaginationUI()
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        Finally
            _isLoading = False
            MainProgressBar.Visibility = Visibility.Collapsed
        End Try
    End Sub

    Private Sub UpdatePaginationUI()
        Dim totalPages = Math.Ceiling(_totalCount / _pageSize)
        If totalPages = 0 Then totalPages = 1
        LblPageInfo.Text = $"Page {_currentPage} of {totalPages} (Total: {_totalCount})"

        BtnBackward.IsEnabled = (_currentPage > 1)
        BtnBegin.IsEnabled = (_currentPage > 1)
        BtnForward.IsEnabled = (_currentPage < totalPages)
        BtnEnd.IsEnabled = (_currentPage < totalPages)
    End Sub

    Private Sub BtnBack_Click(sender As Object, e As RoutedEventArgs)
        NavigationService.GoBack()
    End Sub

    Private Sub BtnSearch_Click(sender As Object, e As RoutedEventArgs) Handles BtnSearch.Click
        _currentPage = 1
        LoadData()
    End Sub

    Private Sub Description_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim rowView As DataRowView = DirectCast(btn.DataContext, DataRowView)
        
        If rowView IsNot Nothing Then
            Dim id = Convert.ToInt32(rowView("ID"))
            Dim win As New CalibrationMasterDetailsWindow()
            win.Owner = Window.GetWindow(Me)
            win.LoadDetails(id, _tableName)
            win.ShowDialog()
        End If
    End Sub

    ' --- PAGINATION HANDLERS ---
    Private Sub BtnBegin_Click(sender As Object, e As RoutedEventArgs) Handles BtnBegin.Click
        _currentPage = 1
        LoadData()
    End Sub

    Private Sub BtnBackward_Click(sender As Object, e As RoutedEventArgs) Handles BtnBackward.Click
        If _currentPage > 1 Then
            _currentPage -= 1
            LoadData()
        End If
    End Sub

    Private Sub BtnForward_Click(sender As Object, e As RoutedEventArgs) Handles BtnForward.Click
        Dim totalPages = Math.Ceiling(_totalCount / _pageSize)
        If _currentPage < totalPages Then
            _currentPage += 1
            LoadData()
        End If
    End Sub

    Private Sub BtnEnd_Click(sender As Object, e As RoutedEventArgs) Handles BtnEnd.Click
        Dim totalPages = Math.Ceiling(_totalCount / _pageSize)
        If totalPages > 0 Then
            _currentPage = CInt(totalPages)
            LoadData()
        End If
    End Sub

    ' --- CRUD ACTIONS ---
    Private Sub BtnAdd_Click(sender As Object, e As RoutedEventArgs) Handles BtnAdd.Click
        Dim win As New AddEditCalibrationMasterWindow(_tableName, _masterName)
        win.Owner = Window.GetWindow(Me)
        If win.ShowDialog() = True Then
            LoadData()
        End If
    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As RoutedEventArgs) Handles BtnEdit.Click
        If MasterDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a record to edit.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = MasterDataGrid.SelectedItem
        Dim win As New AddEditCalibrationMasterWindow(_tableName, _masterName)
        win.Owner = Window.GetWindow(Me)
        win.PopulateForm(rowView.Row)
        
        If win.ShowDialog() = True Then
            LoadData()
        End If
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As RoutedEventArgs) Handles BtnDelete.Click
        If MasterDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a record to delete.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim selectedRow As DataRowView = DirectCast(MasterDataGrid.SelectedItem, DataRowView)
        Dim id = Convert.ToInt32(selectedRow("ID"))
        Dim desc = selectedRow("Description").ToString()

        Dim result = MessageBox.Show($"Are you sure you want to delete '{desc}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)

        If result = MessageBoxResult.Yes Then
            If _mySql.DeleteRecord(_tableName, id) Then
                MessageBox.Show("Record successfully deleted.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                LoadData()
            Else
                MessageBox.Show("Delete failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub

    ' --- IMPORT / EXPORT ---
    Private Sub BtnDownload_Click(sender As Object, e As RoutedEventArgs) Handles BtnDownload.Click
        Try
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
            saveFileDialog.FileName = $"{_masterName.Replace(" ", "_")}_Export"
            
            If saveFileDialog.ShowDialog() = True Then
                Dim view As DataView = DirectCast(MasterDataGrid.ItemsSource, DataView)
                Dim dt = view.ToTable()
                
                Using sw As New StreamWriter(saveFileDialog.FileName)
                    Dim cols = {"ID", "Date", "Status", "Description", "LeastCount", "MasterUncertainty", "CalDate", "DueDate"}
                    sw.WriteLine(String.Join(",", cols))

                    For Each row As DataRow In dt.Rows
                        Dim lineValues As New List(Of String)
                        For Each col In cols
                            Dim val = If(row.Table.Columns.Contains(col), row(col).ToString(), "")
                            lineValues.Add($"""{val.Replace("""", """""")}""")
                        Next
                        sw.WriteLine(String.Join(",", lineValues))
                    Next
                End Using
                MessageBox.Show("Data exported successfully.", "Export Success", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Export Error: " & ex.Message)
        End Try
    End Sub

    Private Async Sub BtnImport_Click(sender As Object, e As RoutedEventArgs) Handles BtnImport.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Import Files (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv"
        If ofd.ShowDialog() = True Then
            Await ImportFromFileAsync(ofd.FileName)
        End If
    End Sub

    Private Async Function ImportFromFileAsync(filePath As String) As Task
        BtnImport.IsEnabled = False
        BtnImport.Content = "Importing..."
        Dim importedCount As Integer = 0

        Try
            MainProgressBar.Visibility = Visibility.Visible
            Await Task.Run(Sub()
                               Try
                                   Using stream = File.Open(filePath, FileMode.Open, FileAccess.Read)
                                       Dim reader = If(filePath.ToLower().EndsWith(".csv"), 
                                                    ExcelReaderFactory.CreateCsvReader(stream), 
                                                    ExcelReaderFactory.CreateReader(stream))
                                       Using reader
                                           Dim result = reader.AsDataSet()
                                           Dim table = result.Tables(0)
                                           
                                           ' Simple header check and row-by-row insert
                                           For i As Integer = 1 To table.Rows.Count - 1
                                               Dim row = table.Rows(i)
                                               ' Assuming columns: Description[0], LeastCount[1], Uncertainty[2], CalDate[3], DueDate[4]
                                               Dim desc = row(0).ToString()
                                               If String.IsNullOrEmpty(desc) Then Continue For
                                               
                                               Dim lc = row(1).ToString()
                                               Dim unc As Decimal = 0
                                               Decimal.TryParse(row(2).ToString(), unc)
                                               
                                               Dim calD As Date? = Nothing
                                               Dim dueD As Date? = Nothing
                                               
                                               Try
                                                   If Not IsDBNull(row(3)) Then calD = Convert.ToDateTime(row(3))
                                                   If Not IsDBNull(row(4)) Then dueD = Convert.ToDateTime(row(4))
                                               Catch : End Try

                                               If _mySql.InsertCalibrationMasterRecord(_tableName, desc, lc, unc, calD, dueD, "") Then
                                                   importedCount += 1
                                               End If
                                           Next
                                       End Using
                                   End Using
                               Catch ex As Exception
                                   Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                               End Try
                           End Sub)

            MessageBox.Show($"Import Complete! Added {importedCount} records.")
            LoadData()
        Finally
            BtnImport.IsEnabled = True
            BtnImport.Content = "📤 Import"
            MainProgressBar.Visibility = Visibility.Collapsed
        End Try
    End Function

    Private Sub RootGrid_MouseDown(sender As Object, e As Input.MouseButtonEventArgs)
        MasterDataGrid.SelectedItem = Nothing
    End Sub

End Class

