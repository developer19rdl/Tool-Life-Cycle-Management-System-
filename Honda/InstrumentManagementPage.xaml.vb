Imports System.Windows
Imports System.Windows.Controls
Imports System.Data
Imports Microsoft.Win32
Imports System.IO
Imports ExcelDataReader

Public Class InstrumentManagementPage
    Inherits Page

    Private _mySql As New MySQLClass()
    Private _currentPage As Integer = 1
    Private _pageSize As Integer = 20
    Private _totalCount As Integer = 0
    Private _isLoading As Boolean = False
    Private _typeNameFilter As String = ""
    Private _tableName As String = ""
    Private _selectedGroup As String = "All"

    Public Sub New(Optional typeName As String = "")
        InitializeComponent()
        _typeNameFilter = typeName
        ' forInventory:=True ensures the correct inventory table name is resolved,
        ' not a calibration table name, even if names collide with calibration hardcodes.
        _tableName = MySQLClass.TypeNameToTableName(typeName, forInventory:=True)
    End Sub

    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        _mySql.CheckAndAddUploadedDocColumn()
        LoadFilters()
        LoadGroups()
        LoadData()
    End Sub

    Private Sub LoadFilters()
        Try
            ' Populate filter dropdowns from this type's own table
            Dim dtLine = _mySql.ReadDatatable($"SELECT DISTINCT Line FROM `{_tableName}` WHERE Line IS NOT NULL AND Line != ''")
            ComboLine.Items.Clear()
            ComboLine.Items.Add("All Lines")
            For Each row As DataRow In dtLine.Rows
                ComboLine.Items.Add(row("Line").ToString())
            Next
            ComboLine.SelectedIndex = 0

            Dim dtSec = _mySql.ReadDatatable($"SELECT DISTINCT Section FROM `{_tableName}` WHERE Section IS NOT NULL AND Section != ''")
            ComboSection.Items.Clear()
            ComboSection.Items.Add("All Sections")
            For Each row As DataRow In dtSec.Rows
                ComboSection.Items.Add(row("Section").ToString())
            Next
            ComboSection.SelectedIndex = 0

            Dim dtLoc = _mySql.ReadDatatable($"SELECT DISTINCT Location FROM `{_tableName}` WHERE Location IS NOT NULL AND Location != ''")
            ComboLocation.Items.Clear()
            ComboLocation.Items.Add("All Locations")
            For Each row As DataRow In dtLoc.Rows
                ComboLocation.Items.Add(row("Location").ToString())
            Next
            ComboLocation.SelectedIndex = 0

            Dim dtSize = _mySql.ReadDatatable($"SELECT DISTINCT Size FROM `{_tableName}` WHERE Size IS NOT NULL AND Size != ''")
            ComboSize.Items.Clear()
            ComboSize.Items.Add("All Sizes")
            For Each row As DataRow In dtSize.Rows
                ComboSize.Items.Add(row("Size").ToString())
            Next
            ComboSize.SelectedIndex = 0

            ComboStatus.Items.Clear()
            ComboStatus.Items.Add("All Statuses")
            ComboStatus.Items.Add("Active")
            ComboStatus.Items.Add("In-Active")
            ComboStatus.SelectedIndex = 0
        Catch ex As Exception
            Console.WriteLine("Error loading filters: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadData()
        If _isLoading Then Return
        _isLoading = True
        MainProgressBar.Visibility = Visibility.Visible

        Try
            Dim filters As New Dictionary(Of String, String)()
            If ComboLine.SelectedIndex > 0 Then filters.Add("Line", ComboLine.SelectedItem.ToString())
            If ComboSection.SelectedIndex > 0 Then filters.Add("Section", ComboSection.SelectedItem.ToString())
            If ComboLocation.SelectedIndex > 0 Then filters.Add("Location", ComboLocation.SelectedItem.ToString())
            If ComboSize.SelectedIndex > 0 Then filters.Add("Size", ComboSize.SelectedItem.ToString())
            If ComboStatus.SelectedIndex > 0 Then
                Dim stat = If(ComboStatus.SelectedItem.ToString() = "Active", "Active", "In-Active")
                filters.Add("Status", stat)
            End If

            If _selectedGroup <> "All" Then
                filters.Add("Group", _selectedGroup)
            End If

            ' Update Title UI
            If Not String.IsNullOrEmpty(_typeNameFilter) Then
                TxtHeaderTitle.Text = $"{MySQLClass.ToTitleCase(_typeNameFilter)} Masterlist"
            End If

            Dim keyword As String = TxtSearch.Text
            _totalCount = _mySql.GetInstrumentInventoryCount(_tableName, filters, keyword)

            Dim offset = (_currentPage - 1) * _pageSize
            Dim dt = _mySql.GetInstrumentInventoryData(_tableName, offset, _pageSize, filters, keyword)

            InstrumentsDataGrid.ItemsSource = dt.DefaultView
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
        If NavigationService.CanGoBack Then
            NavigationService.GoBack()
        Else
            NavigationService.Navigate(New DatabasePage())
        End If
    End Sub

    ' --- EVENT HANDLERS ---

    Private Sub BtnSearch_Click(sender As Object, e As RoutedEventArgs) Handles BtnSearch.Click
        _currentPage = 1
        LoadData()
    End Sub

    Private Sub Filter_Changed(sender As Object, e As SelectionChangedEventArgs) Handles ComboLine.SelectionChanged, ComboSection.SelectionChanged, ComboLocation.SelectionChanged, ComboSize.SelectionChanged, ComboStatus.SelectionChanged
        If Not _isLoading Then
            _currentPage = 1
            LoadData()
        End If
    End Sub

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

    Private Sub BtnAdd_Click(sender As Object, e As RoutedEventArgs) Handles BtnAdd.Click
        Dim win As New AddEditInstrumentWindow(_tableName)
        win.Owner = Window.GetWindow(Me)

        If Not String.IsNullOrEmpty(_typeNameFilter) Then
            win.InstrumentType = _typeNameFilter
            win.TxtName.Text = _typeNameFilter
        End If

        If win.ShowDialog() = True Then
            LoadData()
        End If
    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As RoutedEventArgs) Handles BtnEdit.Click
        If InstrumentsDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select an instrument to edit.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = InstrumentsDataGrid.SelectedItem
        Dim win As New AddEditInstrumentWindow(_tableName)
        win.Owner = Window.GetWindow(Me)

        ' Set InstrumentType before populating so sizes can load
        If rowView.Row.Table.Columns.Contains("InstrumentName") AndAlso Not IsDBNull(rowView("InstrumentName")) Then
            win.InstrumentType = rowView("InstrumentName").ToString()
        End If

        win.PopulateForm(rowView.Row)
        If win.ShowDialog() = True Then
            LoadData()
        End If
    End Sub

    Private Sub InstrumentsDataGrid_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        If InstrumentsDataGrid.SelectedItem Is Nothing Then
            e.Handled = True ' Don't show menu if nothing selected
            Return
        End If

        Dim selectedRow As DataRowView = DirectCast(InstrumentsDataGrid.SelectedItem, DataRowView)
        Dim status = selectedRow("Status").ToString()

        If status = "In-Active" Then
            MenuWriteOff.Header = "Reintroduction"
        Else
            MenuWriteOff.Header = "Writeoff"
        End If
    End Sub

    ' --- CONTEXT MENU HANDLERS ---
    Private Sub MenuWriteOff_Click(sender As Object, e As RoutedEventArgs)
        If InstrumentsDataGrid.SelectedItem Is Nothing Then Return
        Dim selectedRow As DataRowView = DirectCast(InstrumentsDataGrid.SelectedItem, DataRowView)
        Dim ctrlNo = selectedRow("ControlNo").ToString()
        Dim status = selectedRow("Status").ToString()
        
        If status = "In-Active" Then
            Dim reIntroWin As New ReintroductionWindow(ctrlNo, _tableName)
            reIntroWin.Owner = Window.GetWindow(Me)
            reIntroWin.ShowDialog()
        Else
            Dim writeOffWin As New WriteOffWindow(ctrlNo, _tableName)
            writeOffWin.Owner = Window.GetWindow(Me)
            writeOffWin.ShowDialog()
        End If
        
        LoadData()
    End Sub

    Private Sub MenuMenuView_Click(sender As Object, e As RoutedEventArgs)
        If InstrumentsDataGrid.SelectedItem Is Nothing Then Return
        Dim selectedRow As DataRowView = DirectCast(InstrumentsDataGrid.SelectedItem, DataRowView)
        Dim ctrlNo = selectedRow("ControlNo").ToString()
        
        Dim historyWin As New ItemHistoryWindow(ctrlNo, "Instrument", _tableName)
        historyWin.ShowDialog()
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As RoutedEventArgs) Handles BtnDelete.Click
        If InstrumentsDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select an instrument to delete.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim selectedRow As DataRowView = DirectCast(InstrumentsDataGrid.SelectedItem, DataRowView)
        Dim ctrlNo = selectedRow("ControlNo").ToString()
        Dim id = Convert.ToInt32(selectedRow("ID"))

        Dim result = MessageBox.Show($"Are you sure you want to delete '{ctrlNo}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)

        If result = MessageBoxResult.Yes Then
            If _mySql.DeleteRecord(_tableName, id) Then
                MessageBox.Show($"{ctrlNo} successfully deleted.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                LoadData()
            Else
                MessageBox.Show("Delete failed. Please try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
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
        Dim skippedCount As Integer = 0

        Try
            MainProgressBar.Minimum = 0
            MainProgressBar.Value = 0
            MainProgressBar.Maximum = 100
            MainProgressBar.Visibility = Visibility.Visible

            Await Task.Run(Sub()
                               Try
                                   Using stream = File.Open(filePath, FileMode.Open, FileAccess.Read)
                                        Dim reader = If(filePath.ToLower().EndsWith(".csv"), 
                                                     ExcelReaderFactory.CreateCsvReader(stream), 
                                                     ExcelReaderFactory.CreateReader(stream))
                                        Using reader
                                           Dim conf = New ExcelDataSetConfiguration() With {
                                               .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                   .UseHeaderRow = False
                                               }
                                           }
                                           Dim result = reader.AsDataSet(conf)

                                           Dim totalRows As Integer = 0
                                           For Each table As DataTable In result.Tables
                                               totalRows += table.Rows.Count
                                           Next

                                           Application.Current.Dispatcher.Invoke(Sub() MainProgressBar.Maximum = totalRows)

                                           For Each table As DataTable In result.Tables
                                               If table.Rows.Count < 2 Then
                                                   Application.Current.Dispatcher.Invoke(Sub() MainProgressBar.Value += table.Rows.Count)
                                                   Continue For
                                               End If

                                               Dim headerRowIndex As Integer = -1
                                               For r As Integer = 0 To Math.Min(15, table.Rows.Count - 1)
                                                   For c As Integer = 0 To table.Columns.Count - 1
                                                       If table.Rows(r)(c).ToString().Trim().Equals("Control No", StringComparison.OrdinalIgnoreCase) Then
                                                           headerRowIndex = r
                                                           Exit For
                                                       End If
                                                   Next
                                                   If headerRowIndex <> -1 Then Exit For
                                               Next

                                               If headerRowIndex = -1 Then Continue For

                                               Dim headerRow = table.Rows(headerRowIndex)
                                               Dim idxControl = -1, idxName = -1, idxColor = -1, idxSize = -1, idxLine = -1
                                               Dim idxModel = -1, idxMaker = -1, idxSection = -1, idxLoc = -1, idxRemark = -1, idxReq = -1
                                               Dim idxAsset = -1, idxDesc = -1, idxRfid = -1, idxDD = -1, idxMM = -1, idxYY = -1

                                               For c As Integer = 0 To table.Columns.Count - 1
                                                   Dim colName As String = headerRow(c).ToString().Trim()
                                                   If IsMatch(colName, "Control No") Then idxControl = c
                                                   If IsMatch(colName, "Color", "Colour") Then idxColor = c
                                                   If IsMatch(colName, "Size") Then idxSize = c
                                                   If IsMatch(colName, "Maker Name", "MakerName") Then idxMaker = c
                                                   If IsMatch(colName, "Section") Then idxSection = c
                                                   If IsMatch(colName, "Location") Then idxLoc = c
                                                   If IsMatch(colName, "Remark", "Remarks") Then idxRemark = c
                                                   If IsMatch(colName, "Request No", "Request No.", "RequestNo") Then idxReq = c
                                                   If IsMatch(colName, "Instrument Name", "Name", "InstrumentName") Then idxName = c
                                                   If IsMatch(colName, "Instrument Description", "Description", "InstrumentDescription") Then idxDesc = c
                                                   If IsMatch(colName, "Model", "Model No", "InstrumentModelNo", "Model No.") Then idxModel = c
                                                   If IsMatch(colName, "Line") Then idxLine = c
                                                   If IsMatch(colName, "Asset", "Asset No", "Asset No.", "AssetNo") Then idxAsset = c
                                                   If IsMatch(colName, "RFID", "RFID Tag", "RFID Tag No.") Then idxRfid = c
                                                   If IsMatch(colName, "DD", "D", "Date") Then idxDD = c
                                                   If IsMatch(colName, "MM", "M", "Month") Then idxMM = c
                                                   If IsMatch(colName, "YY", "Y", "Year") Then idxYY = c
                                               Next

                                               For i As Integer = headerRowIndex + 1 To table.Rows.Count - 1
                                                   Dim row = table.Rows(i)
                                                   Dim controlNoVal As String = GetVal(row, idxControl)

                                                   If idxControl = -1 OrElse String.IsNullOrWhiteSpace(controlNoVal) Then
                                                       Application.Current.Dispatcher.Invoke(Sub() MainProgressBar.Value += 1)
                                                       Continue For
                                                   End If

                                                   If _mySql.IsControlNoDuplicate(_tableName, controlNoVal) Then
                                                       skippedCount += 1
                                                       Application.Current.Dispatcher.Invoke(Sub() MainProgressBar.Value += 1)
                                                       Continue For
                                                   End If

                                                   Dim customDate As Date? = Nothing
                                                   If idxDD <> -1 AndAlso idxMM <> -1 AndAlso idxYY <> -1 Then
                                                       Dim ddStr = GetVal(row, idxDD)
                                                       Dim mmStr = GetVal(row, idxMM)
                                                       Dim yyStr = GetVal(row, idxYY)
                                                       
                                                       If Not String.IsNullOrWhiteSpace(ddStr) AndAlso Not String.IsNullOrWhiteSpace(mmStr) AndAlso Not String.IsNullOrWhiteSpace(yyStr) Then
                                                           Dim d As Integer, m As Integer, y As Integer
                                                           If Integer.TryParse(ddStr, d) AndAlso Integer.TryParse(mmStr, m) AndAlso Integer.TryParse(yyStr, y) Then
                                                               If y < 100 Then
                                                                   y += 2000
                                                               End If
                                                               Try
                                                                   customDate = New Date(y, m, d)
                                                               Catch ex As Exception
                                                                   ' Ignore invalid date
                                                               End Try
                                                           End If
                                                       End If
                                                   End If

                                                   Dim nameStr = If(Not String.IsNullOrEmpty(_typeNameFilter), _typeNameFilter, GetVal(row, idxName))

                                                   Dim success As Boolean = _mySql.Insertinstrument(
                                                       _tableName,
                                                       nameStr,
                                                       GetVal(row, idxDesc),
                                                       controlNoVal,
                                                       GetVal(row, idxColor),
                                                       GetVal(row, idxModel),
                                                       GetVal(row, idxLine),
                                                       GetVal(row, idxAsset),
                                                       GetVal(row, idxSize),
                                                       GetVal(row, idxMaker),
                                                       GetVal(row, idxSection),
                                                       GetVal(row, idxLoc),
                                                       GetVal(row, idxReq),
                                                       GetVal(row, idxRemark),
                                                       "",
                                                       GetVal(row, idxRfid),
                                                       "",
                                                       0,
                                                       customDate
                                                   )

                                                   If success Then importedCount += 1
                                                   Application.Current.Dispatcher.Invoke(Sub() MainProgressBar.Value += 1)
                                               Next
                                           Next
                                       End Using
                                   End Using
                               Catch ex As IOException
                                   Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("The import file is currently open in another program. Please close it and try again.", "File in Use", MessageBoxButton.OK, MessageBoxImage.Warning))
                               Catch ex As Exception
                                   Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error))
                               End Try
                           End Sub)

            MessageBox.Show($"Import Complete!" & Environment.NewLine & Environment.NewLine &
                   $"Successfully added: {importedCount} records." & Environment.NewLine &
                   $"Skipped (Duplicates or Empty): {skippedCount} records.", "Import Summary", MessageBoxButton.OK, MessageBoxImage.Information)

            _mySql.SyncCategoryControlFromInventory()
            LoadFilters()
            LoadGroups()
            LoadData()

        Catch ex As Exception
            MessageBox.Show("Import Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            BtnImport.IsEnabled = True
            BtnImport.Content = "📤 Import"
            MainProgressBar.Visibility = Visibility.Collapsed
            MainProgressBar.Value = 0
        End Try
    End Function

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

    Private Sub BtnDownload_Click(sender As Object, e As RoutedEventArgs) Handles BtnDownload.Click
        Try
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
            
            Dim defaultName As String = "Instruments_Export"
            If Not String.IsNullOrEmpty(_typeNameFilter) Then
                defaultName = _typeNameFilter.Replace(" ", "_") & "_Export"
            End If
            saveFileDialog.FileName = defaultName
            
            If saveFileDialog.ShowDialog() = True Then
                Dim view As DataView = DirectCast(InstrumentsDataGrid.ItemsSource, DataView)
                Dim dt = view.ToTable()
                
                Using sw As New StreamWriter(saveFileDialog.FileName)
                    Dim cols = {"ID", "Date", "ControlNo", "InstrumentName", "InstrumentDescription", "Color", "CategoryControl", "InstrumentModelNo", "Line", "AssetNo", "Size", "MakerName", "Section", "Location", "Remark", "RequestNo"}
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
                
                MessageBox.Show("Data exported successfully to CSV.", "Export Success", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Export Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub LoadGroups()
        Try
            Dim prefixes = _mySql.GetControlNoPrefixes(_tableName)
            GroupsItemsControl.ItemsSource = prefixes
        Catch ex As Exception
            Console.WriteLine("Error loading groups: " & ex.Message)
        End Try
    End Sub

    Private Sub GroupButton_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim group = btn.Content.ToString()
        _selectedGroup = group

        For Each item In GroupsItemsControl.ItemsSource
            Dim container = GroupsItemsControl.ItemContainerGenerator.ContainerFromItem(item)
            If container IsNot Nothing Then
                Dim childBtn = FindVisualChild(Of Button)(container)
                If childBtn IsNot Nothing Then
                    childBtn.Tag = If(item.ToString() = _selectedGroup, "Selected", "")
                End If
            End If
        Next

        _currentPage = 1
        LoadData()
    End Sub

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

    Private Sub ControlNo_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Dim textBlock = DirectCast(sender, TextBlock)
        Dim ctrlNo = textBlock.Text
        
        Dim detailsWin As New ItemDetailsWindow()
        detailsWin.Owner = Window.GetWindow(Me)
        detailsWin.LoadDetails(ctrlNo, "Instrument", _tableName)
        detailsWin.ShowDialog()
    End Sub

    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        InstrumentsDataGrid.SelectedItem = Nothing
        Keyboard.ClearFocus()
    End Sub

End Class
