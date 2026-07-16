Imports System.Windows.Controls
Imports System.Data
Imports System.Windows
Imports Microsoft.Win32
Imports System.IO
Imports ExcelDataReader

Public Class SettingsPage
    Inherits Page

    Private _mySql As New MySQLClass()

    Private Sub Page_Loaded(sender As Object, e As Windows.RoutedEventArgs) Handles MyBase.Loaded
        ' Ensure TypeImage column exists
        _mySql.CheckAndAddTypeImageColumn()
        ' Load Types by default
        LoadTypes()
    End Sub

    Private Sub LoadTypes()
        Try
            Dim dt = _mySql.GetTypeDetails()
            ' Apply Title Case to TypeName for display
            For Each row As DataRow In dt.Rows
                row("TypeName") = MySQLClass.ToTitleCase(row("TypeName").ToString())
            Next
            TypesDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading types: " & ex.Message)
        End Try
    End Sub

    Private Sub TabButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles BtnTypes.Click, BtnDept.Click, BtnSizes.Click, BtnGaugeSizes.Click, BtnGroupSets.Click, BtnGroups.Click, BtnAdmin.Click, BtnCalibrationMapping.Click, BtnCalibMaster.Click
        ' Reset all tags
        BtnTypes.Tag = ""
        BtnDept.Tag = ""
        BtnSizes.Tag = ""
        BtnGaugeSizes.Tag = ""
        BtnCalibrationMapping.Tag = ""
        BtnGroupSets.Tag = ""
        BtnGroups.Tag = ""
        BtnAdmin.Tag = ""
        BtnCalibMaster.Tag = ""

        ' Set selected tag
        Dim clickedButton = DirectCast(sender, Button)
        clickedButton.Tag = "Selected"

        ' Handle content visibility
        Dim knownTabs As String() = {"BtnTypes", "BtnDept", "BtnSizes", "BtnGaugeSizes", "BtnCalibrationMapping", "BtnCalibMaster", "BtnAdmin"}
        TypesTabView.Visibility = If(clickedButton.Name = "BtnTypes", Windows.Visibility.Visible, Windows.Visibility.Collapsed)
        DeptTabView.Visibility = If(clickedButton.Name = "BtnDept", Windows.Visibility.Visible, Windows.Visibility.Collapsed)
        SizesTabView.Visibility = If(clickedButton.Name = "BtnSizes", Windows.Visibility.Visible, Windows.Visibility.Collapsed)
        GaugeSizesTabView.Visibility = If(clickedButton.Name = "BtnGaugeSizes", Windows.Visibility.Visible, Windows.Visibility.Collapsed)
        CalibrationMappingTabView.Visibility = If(clickedButton.Name = "BtnCalibrationMapping", Windows.Visibility.Visible, Windows.Visibility.Collapsed)
        CalibMasterTabView.Visibility = If(clickedButton.Name = "BtnCalibMaster", Windows.Visibility.Visible, Windows.Visibility.Collapsed)
        AdminTabView.Visibility = If(clickedButton.Name = "BtnAdmin", Windows.Visibility.Visible, Windows.Visibility.Collapsed)

        ' Show placeholder only if none of the known tabs are selected
        OtherTabPlaceholder.Visibility = If(knownTabs.Contains(clickedButton.Name), Windows.Visibility.Collapsed, Windows.Visibility.Visible)

        If clickedButton.Name = "BtnTypes" Then
            LoadTypes()
        ElseIf clickedButton.Name = "BtnDept" Then
            LoadDepartments()
        ElseIf clickedButton.Name = "BtnSizes" Then
            LoadInstrumentSizes()
        ElseIf clickedButton.Name = "BtnGaugeSizes" Then
            LoadGaugeSizes()
        ElseIf clickedButton.Name = "BtnCalibrationMapping" Then
            LoadCalibrationMapping()
        ElseIf clickedButton.Name = "BtnCalibMaster" Then
            LoadCalibrationMasterTypes()
        ElseIf clickedButton.Name = "BtnAdmin" Then
            LoadAdminSettings()
        End If
    End Sub

    ' --- Calibration Mapping Tab Logic ---

    Private Sub LoadCalibrationMapping()
        Try
            Dim dt = _mySql.ReadDatatable("SELECT id, category, type_name, prefix, calibration_category FROM calibrationmapping")
            CalibrationMappingDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading calibration mapping: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnAddMapping_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New AddEditCalibrationMappingWindow()
        win.Owner = Window.GetWindow(Me)
        If win.ShowDialog() = True Then
            LoadCalibrationMapping()
        End If
    End Sub

    Private Sub BtnEditMapping_Click(sender As Object, e As RoutedEventArgs)
        If CalibrationMappingDataGrid.SelectedItem IsNot Nothing Then
            Dim rowView As DataRowView = DirectCast(CalibrationMappingDataGrid.SelectedItem, DataRowView)
            Dim win As New AddEditCalibrationMappingWindow()
            win.Owner = Window.GetWindow(Me)
            win.PopulateForm(rowView.Row)
            If win.ShowDialog() = True Then
                LoadCalibrationMapping()
            End If
        Else
            MessageBox.Show("Please select a mapping to edit.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
        End If
    End Sub

    Private Sub BtnDeleteMapping_Click(sender As Object, e As RoutedEventArgs)
        If CalibrationMappingDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a record to delete.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = DirectCast(CalibrationMappingDataGrid.SelectedItem, DataRowView)
        Dim id = Convert.ToInt32(rowView("id"))
        Dim type = rowView("type_name").ToString()

        Dim result = MessageBox.Show($"Are you sure you want to delete calibration mapping for '{type}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result = MessageBoxResult.Yes Then
            If _mySql.DeleteRecord("calibrationmapping", id) Then
                LoadCalibrationMapping()
            Else
                MessageBox.Show("Delete failed.")
            End If
        End If
    End Sub

    ' --- Department Tab Logic ---

    Private Sub LoadDepartments()
        Try
            Dim dt = _mySql.GetDepartments()
            DeptDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading departments: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnAddDept_Click(sender As Object, e As RoutedEventArgs)
        Dim deptName = TxtDeptName.Text.Trim()
        If String.IsNullOrEmpty(deptName) Then
            MessageBox.Show("Please enter a department name.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        If _mySql.InsertDepartment(deptName) Then
            TxtDeptName.Clear()
            LoadDepartments()
        Else
            MessageBox.Show("Failed to add department. It might already exist.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub BtnDeleteDept_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim rowView = DirectCast(btn.DataContext, DataRowView)
        Dim id = Convert.ToInt32(rowView("ID"))
        Dim deptName = rowView("DepartmentName").ToString()

        Dim result = MessageBox.Show($"Are you sure you want to delete department '{deptName}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result = MessageBoxResult.Yes Then
            If _mySql.DeleteRecord("departmentmaster", id) Then
                LoadDepartments()
            Else
                MessageBox.Show("Delete failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub

    ' --- Calibration Master Types Tab Logic ---

    Private Sub LoadCalibrationMasterTypes()
        Try
            Dim dt = _mySql.GetCalibrationMasterTypes()
            CalibMasterDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading calibration master types: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnAddCalibMaster_Click(sender As Object, e As RoutedEventArgs)
        Dim rawText = TxtCalibMasterName.Text.Trim()
        Dim masterName = MySQLClass.ToTitleCase(rawText)

        If String.IsNullOrEmpty(masterName) Then
            MessageBox.Show("Please enter a calibration master type name.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Try
            If _mySql.InsertCalibrationMasterType(masterName) Then
                Dim tableName = MySQLClass.TypeNameToTableName(masterName)
                _mySql.CreateCalibrationMasterTable(tableName)
                TxtCalibMasterName.Clear()
                LoadCalibrationMasterTypes()
            Else
                MessageBox.Show("Failed to connect to database.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Error adding type: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnDeleteCalibMaster_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim rowView = DirectCast(btn.DataContext, DataRowView)
        Dim id = Convert.ToInt32(rowView("ID"))
        Dim masterName = rowView("CalibrationMasterName").ToString()

        Dim result = MessageBox.Show($"Are you sure you want to delete calibration master type '{masterName}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result = MessageBoxResult.Yes Then
            If _mySql.DeleteRecord("calibrationmaster", id) Then
                ' 1. Drop existing table
                Dim tableName = MySQLClass.TypeNameToTableName(masterName)
                _mySql.DropTypeTable(tableName)

                ' 2. Delete from calibrationmaster_details
                Try
                    Dim deleteDetailsSql = $"DELETE FROM calibrationmaster_details WHERE Description = '{masterName.Replace("'", "''")}'"
                    _mySql.ExecuteNonQuery(deleteDetailsSql)
                Catch exSync As Exception
                    Console.WriteLine("Sync detail deletion error: " & exSync.Message)
                End Try

                ' 3. Delete physical files folder
                Try
                    Dim appPath = AppDomain.CurrentDomain.BaseDirectory
                    Dim targetFolder = System.IO.Path.Combine(appPath, "Documents", "Calibration Master", masterName)
                    If System.IO.Directory.Exists(targetFolder) Then
                        System.IO.Directory.Delete(targetFolder, True)
                    End If
                Catch ex As Exception
                    Console.WriteLine("Folder deletion error: " & ex.Message)
                End Try

                LoadCalibrationMasterTypes()
            Else
                MessageBox.Show("Delete failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub

    Private Sub BtnAddType_Click(sender As Object, e As Windows.RoutedEventArgs)
        Dim win As New AddEditTypeWindow()
        win.Owner = Window.GetWindow(Me)
        If win.ShowDialog() = True Then
            LoadTypes()
        End If
    End Sub

    Private Sub BtnEditType_Click(sender As Object, e As Windows.RoutedEventArgs)
        If TypesDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a type to edit.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = TypesDataGrid.SelectedItem
        Dim win As New AddEditTypeWindow()
        win.Owner = Window.GetWindow(Me)
        win.PopulateForm(rowView.Row)
        If win.ShowDialog() = True Then
            LoadTypes()
        End If
    End Sub

    Private Sub BtnDeleteType_Click(sender As Object, e As Windows.RoutedEventArgs)
        If TypesDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a type to delete.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = TypesDataGrid.SelectedItem
        Dim id = Convert.ToInt32(rowView("ID"))
        Dim typeName = rowView("TypeName").ToString()

        Dim result = MessageBox.Show($"Are you sure you want to delete '{typeName}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result = MessageBoxResult.Yes Then
            ' forInventory:=True ensures gauge types like "Plain Plug Gauge" drop the inventory
            ' table (plain_plug_gauge) instead of the calibration table (plain_plug_gauge_calibration).
            Dim tableName = MySQLClass.TypeNameToTableName(typeName, forInventory:=True)

            If _mySql.DeleteRecord("type_details", id) Then
                ' Also drop the corresponding database table
                _mySql.DropTypeTable(tableName)

                MessageBox.Show($"{typeName} successfully deleted.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                LoadTypes()
            Else
                MessageBox.Show("Delete failed. Please try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub
    ' --- Instrument Sizes Tab Logic ---

    Private Sub LoadInstrumentSizes()
        Try
            Dim dt = _mySql.GetInstrumentSizes()
            ' Format to Title Case
            For Each row As DataRow In dt.Rows
                row("InstrumentType") = MySQLClass.ToTitleCase(row("InstrumentType").ToString())
            Next
            SizesDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading instrument sizes: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnAddSize_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New AddEditInstrumentSizeWindow()
        win.Owner = Window.GetWindow(Me)
        If win.ShowDialog() = True Then
            LoadInstrumentSizes()
        End If
    End Sub

    Private Sub BtnEditSize_Click(sender As Object, e As RoutedEventArgs)
        If SizesDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a record to edit.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = SizesDataGrid.SelectedItem
        Dim win As New AddEditInstrumentSizeWindow()
        win.Owner = Window.GetWindow(Me)
        win.PopulateForm(rowView.Row)
        If win.ShowDialog() = True Then
            LoadInstrumentSizes()
        End If
    End Sub

    Private Sub BtnDeleteSize_Click(sender As Object, e As RoutedEventArgs)
        If SizesDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a record to delete.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = SizesDataGrid.SelectedItem
        Dim id = Convert.ToInt32(rowView("ID"))
        Dim type = rowView("InstrumentType").ToString()
        Dim size = rowView("Size").ToString()

        Dim result = MessageBox.Show($"Are you sure you want to delete '{type}' (Size: {size})?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result = MessageBoxResult.Yes Then
            If _mySql.DeleteRecord("categorycontrol", id) Then
                MessageBox.Show("Record deleted successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                LoadInstrumentSizes()
            Else
                MessageBox.Show("Delete failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub

    ' --- Gauge Sizes Tab Logic ---

    Private Sub LoadGaugeSizes()
        Try
            Dim dt = _mySql.GetGaugeSizes()
            ' Format to Title Case
            For Each row As DataRow In dt.Rows
                row("InstrumentType") = MySQLClass.ToTitleCase(row("InstrumentType").ToString())
            Next
            GaugeSizesDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading gauge sizes: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnAddGaugeSize_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New AddEditGaugeSizeWindow()
        win.Owner = Window.GetWindow(Me)
        If win.ShowDialog() = True Then
            LoadGaugeSizes()
        End If
    End Sub

    Private Sub BtnEditGaugeSize_Click(sender As Object, e As RoutedEventArgs)
        If GaugeSizesDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a record to edit.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = GaugeSizesDataGrid.SelectedItem
        Dim win As New AddEditGaugeSizeWindow()
        win.Owner = Window.GetWindow(Me)
        win.PopulateForm(rowView.Row)
        If win.ShowDialog() = True Then
            LoadGaugeSizes()
        End If
    End Sub

    Private Sub BtnDeleteGaugeSize_Click(sender As Object, e As RoutedEventArgs)
        If GaugeSizesDataGrid.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a record to delete.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim rowView As DataRowView = GaugeSizesDataGrid.SelectedItem
        Dim id = Convert.ToInt32(rowView("ID"))
        Dim type = rowView("InstrumentType").ToString()
        Dim size = rowView("Size").ToString()

        Dim result = MessageBox.Show($"Are you sure you want to delete '{type}' (Size: {size})?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result = MessageBoxResult.Yes Then
            If _mySql.DeleteRecord("gauge_categorycontrol", id) Then
                MessageBox.Show("Record deleted successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                LoadGaugeSizes()
            Else
                MessageBox.Show("Delete failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub
    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        TypesDataGrid.SelectedItem = Nothing
        DeptDataGrid.SelectedItem = Nothing
        SizesDataGrid.SelectedItem = Nothing
        GaugeSizesDataGrid.SelectedItem = Nothing
        CalibrationMappingDataGrid.SelectedItem = Nothing
        CalibMasterDataGrid.SelectedItem = Nothing
    End Sub

    ' ========== SHARED IMPORT HELPERS ==========

    Private Function GetVal(row As DataRow, idx As Integer) As String
        If idx = -1 OrElse IsDBNull(row(idx)) Then Return ""
        Return row(idx).ToString().Trim()
    End Function

    Private Function IsMatch(header As String, ParamArray targets() As String) As Boolean
        Dim clean = System.Text.RegularExpressions.Regex.Replace(header, "\s*\([^)]*\)$", "").Replace(" ", "").Trim()
        For Each t In targets
            If clean.Equals(t.Replace(" ", ""), StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    Private Sub ShowProgress(show As Boolean)
        SettingsProgressBar.Visibility = If(show, Windows.Visibility.Visible, Windows.Visibility.Collapsed)
    End Sub

    Private Function ExportDataGridToCSV(grid As DataGrid, cols() As String, defaultName As String) As Boolean
        Try
            Dim sfd As New SaveFileDialog()
            sfd.Filter = "CSV Files (*.csv)|*.csv"
            sfd.FileName = defaultName
            If sfd.ShowDialog() <> True Then Return False
            Dim view As DataView = DirectCast(grid.ItemsSource, DataView)
            Dim dt = view.ToTable()
            Using sw As New StreamWriter(sfd.FileName)
                sw.WriteLine(String.Join(",", cols))
                For Each row As DataRow In dt.Rows
                    Dim vals As New List(Of String)
                    For Each col In cols
                        Dim v = If(dt.Columns.Contains(col), row(col).ToString(), "")
                        vals.Add($"""{v.Replace("""", """""")}""")
                    Next
                    sw.WriteLine(String.Join(",", vals))
                Next
            End Using
            MessageBox.Show("Exported successfully.", "Export", MessageBoxButton.OK, MessageBoxImage.Information)
            Return True
        Catch ex As Exception
            MessageBox.Show("Export Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Return False
        End Try
    End Function

    Private Function OpenImportFile() As String
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Import Files (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv"
        If ofd.ShowDialog() = True Then Return ofd.FileName
        Return Nothing
    End Function

    Private Function ReadImportTable(filePath As String) As DataTable
        Using stream = File.Open(filePath, FileMode.Open, FileAccess.Read)
            Dim reader = If(filePath.ToLower().EndsWith(".csv"),
                         ExcelReaderFactory.CreateCsvReader(stream),
                         ExcelReaderFactory.CreateReader(stream))
            Using reader
                Dim conf = New ExcelDataSetConfiguration() With {
                    .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {.UseHeaderRow = False}
                }
                Dim ds = reader.AsDataSet(conf)
                If ds.Tables.Count > 0 Then Return ds.Tables(0)
            End Using
        End Using
        Return Nothing
    End Function

    ' Find header row index by looking for a known column name
    Private Function FindHeaderRow(dt As DataTable, ParamArray knownCols() As String) As Integer
        For r As Integer = 0 To Math.Min(15, dt.Rows.Count - 1)
            For c As Integer = 0 To dt.Columns.Count - 1
                Dim cell = dt.Rows(r)(c).ToString().Trim()
                For Each kc In knownCols
                    If cell.Equals(kc, StringComparison.OrdinalIgnoreCase) Then Return r
                Next
            Next
        Next
        Return -1
    End Function

    ' Build a column-index map from a header row
    Private Function MapColumns(headerRow As DataRow, ParamArray mappings() As String) As Dictionary(Of String, Integer)
        ' mappings: "keyName:Header1,Header2,..."
        Dim result As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim pairs As New Dictionary(Of String, String())
        For Each m In mappings
            Dim parts = m.Split(":"c)
            pairs(parts(0)) = parts(1).Split(","c)
        Next
        For c As Integer = 0 To headerRow.Table.Columns.Count - 1
            Dim cell = headerRow(c).ToString().Trim()
            For Each kvp In pairs
                For Each aliasName In kvp.Value
                    If cell.Equals(aliasName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        If Not result.ContainsKey(kvp.Key) Then result(kvp.Key) = c
                    End If
                Next
            Next
        Next
        Return result
    End Function

    ' ========== TYPES TAB ==========

    Private Sub BtnDownloadTypes_Click(sender As Object, e As RoutedEventArgs)
        ExportDataGridToCSV(TypesDataGrid, {"Category", "TypeName", "PrefixMode", "BasePrefix", "SerialDigits"}, "Types_Export")
    End Sub

    Private Async Sub BtnImportTypes_Click(sender As Object, e As RoutedEventArgs)
        Dim path = OpenImportFile()
        If String.IsNullOrEmpty(path) Then Return
        BtnImportTypes.IsEnabled = False
        BtnImportTypes.Content = "Importing..."
        ShowProgress(True)
        Dim imported As Integer = 0, skipped As Integer = 0
        Try
            Await Task.Run(Sub()
                Try
                    Dim dt = ReadImportTable(path)
                    If dt Is Nothing Then Return
                    Dim hi = FindHeaderRow(dt, "Category", "TypeName", "Type Name")
                    If hi = -1 Then Return
                    Dim idx = MapColumns(dt.Rows(hi),
                        "Category:Category",
                        "TypeName:TypeName,Type Name",
                        "PrefixMode:PrefixMode,Prefix Mode",
                        "BasePrefix:BasePrefix,Base Prefix",
                        "SerialDigits:SerialDigits,Digits,Serial Digits")
                    For i = hi + 1 To dt.Rows.Count - 1
                        Dim row = dt.Rows(i)
                        Dim cat = If(idx.ContainsKey("Category"), GetVal(row, idx("Category")), "")
                        Dim name = If(idx.ContainsKey("TypeName"), GetVal(row, idx("TypeName")), "")
                        If String.IsNullOrWhiteSpace(name) Then skipped += 1 : Continue For
                        If _mySql.IsTypeDuplicate(name, cat) Then skipped += 1 : Continue For
                        Dim mode = If(idx.ContainsKey("PrefixMode"), GetVal(row, idx("PrefixMode")), "Manual")
                        Dim prefix = If(idx.ContainsKey("BasePrefix"), GetVal(row, idx("BasePrefix")), "")
                        Dim digits = If(idx.ContainsKey("SerialDigits"), GetVal(row, idx("SerialDigits")), "4")
                        Dim ok = _mySql.InsertTypeDetail(cat, name, mode, prefix, digits)
                        If ok Then
                            Dim tbl = MySQLClass.TypeNameToTableName(name, forInventory:=True)
                            _mySql.CreateTypeTable(tbl, cat)
                            imported += 1
                        Else
                            skipped += 1
                        End If
                    Next
                Catch ex As Exception
                    Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                End Try
            End Sub)
            MessageBox.Show($"Import Complete!{Environment.NewLine}Added: {imported} | Skipped: {skipped}", "Types Import", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadTypes()
        Catch ex As Exception
            MessageBox.Show("Import Error: " & ex.Message)
        Finally
            BtnImportTypes.IsEnabled = True
            BtnImportTypes.Content = "📤 Import"
            ShowProgress(False)
        End Try
    End Sub

    ' ========== DEPT TAB ==========

    Private Sub BtnDownloadDept_Click(sender As Object, e As RoutedEventArgs)
        ExportDataGridToCSV(DeptDataGrid, {"DepartmentName"}, "Departments_Export")
    End Sub

    Private Async Sub BtnImportDept_Click(sender As Object, e As RoutedEventArgs)
        Dim path = OpenImportFile()
        If String.IsNullOrEmpty(path) Then Return
        BtnImportDept.IsEnabled = False
        BtnImportDept.Content = "Importing..."
        ShowProgress(True)
        Dim imported As Integer = 0, skipped As Integer = 0
        Try
            Await Task.Run(Sub()
                Try
                    Dim dt = ReadImportTable(path)
                    If dt Is Nothing Then Return
                    Dim hi = FindHeaderRow(dt, "DepartmentName", "Department Name", "Department")
                    If hi = -1 Then Return
                    Dim idx = MapColumns(dt.Rows(hi), "DepartmentName:DepartmentName,Department Name,Department")
                    For i = hi + 1 To dt.Rows.Count - 1
                        Dim row = dt.Rows(i)
                        Dim name = If(idx.ContainsKey("DepartmentName"), GetVal(row, idx("DepartmentName")), "")
                        If String.IsNullOrWhiteSpace(name) Then skipped += 1 : Continue For
                        If _mySql.IsDepartmentDuplicate(name) Then skipped += 1 : Continue For
                        If _mySql.InsertDepartment(name) Then imported += 1 Else skipped += 1
                    Next
                Catch ex As Exception
                    Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                End Try
            End Sub)
            MessageBox.Show($"Import Complete!{Environment.NewLine}Added: {imported} | Skipped: {skipped}", "Dept Import", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadDepartments()
        Catch ex As Exception
            MessageBox.Show("Import Error: " & ex.Message)
        Finally
            BtnImportDept.IsEnabled = True
            BtnImportDept.Content = "📤 Import"
            ShowProgress(False)
        End Try
    End Sub

    ' ========== INSTRUMENT SIZES TAB ==========

    Private Sub BtnDownloadSizes_Click(sender As Object, e As RoutedEventArgs)
        ExportDataGridToCSV(SizesDataGrid, {"InstrumentType", "Size", "GroupCode", "Active", "Sort"}, "InstrumentSizes_Export")
    End Sub

    Private Async Sub BtnImportSizes_Click(sender As Object, e As RoutedEventArgs)
        Dim path = OpenImportFile()
        If String.IsNullOrEmpty(path) Then Return
        BtnImportSizes.IsEnabled = False
        BtnImportSizes.Content = "Importing..."
        ShowProgress(True)
        Dim imported As Integer = 0, skipped As Integer = 0
        Try
            Await Task.Run(Sub()
                Try
                    Dim dt = ReadImportTable(path)
                    If dt Is Nothing Then Return
                    Dim hi = FindHeaderRow(dt, "InstrumentType", "Instrument Type")
                    If hi = -1 Then Return
                    Dim idx = MapColumns(dt.Rows(hi),
                        "InstrumentType:InstrumentType,Instrument Type",
                        "Size:Size",
                        "GroupCode:GroupCode,Group Code",
                        "Active:Active",
                        "Sort:Sort")
                    For i = hi + 1 To dt.Rows.Count - 1
                        Dim row = dt.Rows(i)
                        Dim itype = If(idx.ContainsKey("InstrumentType"), GetVal(row, idx("InstrumentType")), "")
                        Dim sz = If(idx.ContainsKey("Size"), GetVal(row, idx("Size")), "")
                        If String.IsNullOrWhiteSpace(itype) OrElse String.IsNullOrWhiteSpace(sz) Then skipped += 1 : Continue For
                        If _mySql.IsInstrumentSizeDuplicate(itype, sz) Then skipped += 1 : Continue For
                        Dim gc = If(idx.ContainsKey("GroupCode"), GetVal(row, idx("GroupCode")), "")
                        Dim activeStr = If(idx.ContainsKey("Active"), GetVal(row, idx("Active")), "1")
                        Dim activeVal As Integer = If(activeStr = "0" OrElse activeStr.ToLower() = "false", 0, 1)
                        Dim sortStr = If(idx.ContainsKey("Sort"), GetVal(row, idx("Sort")), "0")
                        If _mySql.InsertInstrumentSize(itype, sz, gc, activeVal, sortStr) Then imported += 1 Else skipped += 1
                    Next
                Catch ex As Exception
                    Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                End Try
            End Sub)
            MessageBox.Show($"Import Complete!{Environment.NewLine}Added: {imported} | Skipped: {skipped}", "Instrument Sizes Import", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadInstrumentSizes()
        Catch ex As Exception
            MessageBox.Show("Import Error: " & ex.Message)
        Finally
            BtnImportSizes.IsEnabled = True
            BtnImportSizes.Content = "📤 Import"
            ShowProgress(False)
        End Try
    End Sub

    ' ========== GAUGE SIZES TAB ==========

    Private Sub BtnDownloadGaugeSizes_Click(sender As Object, e As RoutedEventArgs)
        ExportDataGridToCSV(GaugeSizesDataGrid, {"InstrumentType", "Size", "GroupCode", "Active", "Sort"}, "GaugeSizes_Export")
    End Sub

    Private Async Sub BtnImportGaugeSizes_Click(sender As Object, e As RoutedEventArgs)
        Dim path = OpenImportFile()
        If String.IsNullOrEmpty(path) Then Return
        BtnImportGaugeSizes.IsEnabled = False
        BtnImportGaugeSizes.Content = "Importing..."
        ShowProgress(True)
        Dim imported As Integer = 0, skipped As Integer = 0
        Try
            Await Task.Run(Sub()
                Try
                    Dim dt = ReadImportTable(path)
                    If dt Is Nothing Then Return
                    Dim hi = FindHeaderRow(dt, "InstrumentType", "Instrument Type", "Gauge Type")
                    If hi = -1 Then Return
                    Dim idx = MapColumns(dt.Rows(hi),
                        "InstrumentType:InstrumentType,Instrument Type,Gauge Type",
                        "Size:Size",
                        "GroupCode:GroupCode,Group Code",
                        "Active:Active",
                        "Sort:Sort")
                    For i = hi + 1 To dt.Rows.Count - 1
                        Dim row = dt.Rows(i)
                        Dim itype = If(idx.ContainsKey("InstrumentType"), GetVal(row, idx("InstrumentType")), "")
                        Dim sz = If(idx.ContainsKey("Size"), GetVal(row, idx("Size")), "")
                        If String.IsNullOrWhiteSpace(itype) OrElse String.IsNullOrWhiteSpace(sz) Then skipped += 1 : Continue For
                        If _mySql.IsGaugeSizeDuplicate(itype, sz) Then skipped += 1 : Continue For
                        Dim gc = If(idx.ContainsKey("GroupCode"), GetVal(row, idx("GroupCode")), "")
                        Dim activeStr = If(idx.ContainsKey("Active"), GetVal(row, idx("Active")), "1")
                        Dim activeVal As Integer = If(activeStr = "0" OrElse activeStr.ToLower() = "false", 0, 1)
                        Dim sortStr = If(idx.ContainsKey("Sort"), GetVal(row, idx("Sort")), "0")
                        If _mySql.InsertGaugeSize(itype, sz, gc, activeVal, sortStr) Then imported += 1 Else skipped += 1
                    Next
                Catch ex As Exception
                    Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                End Try
            End Sub)
            MessageBox.Show($"Import Complete!{Environment.NewLine}Added: {imported} | Skipped: {skipped}", "Gauge Sizes Import", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadGaugeSizes()
        Catch ex As Exception
            MessageBox.Show("Import Error: " & ex.Message)
        Finally
            BtnImportGaugeSizes.IsEnabled = True
            BtnImportGaugeSizes.Content = "📤 Import"
            ShowProgress(False)
        End Try
    End Sub

    ' ========== CALIBRATION MAPPING TAB ==========

    Private Sub BtnDownloadMapping_Click(sender As Object, e As RoutedEventArgs)
        ExportDataGridToCSV(CalibrationMappingDataGrid, {"category", "type_name", "prefix", "calibration_category"}, "CalibrationMapping_Export")
    End Sub

    Private Async Sub BtnImportMapping_Click(sender As Object, e As RoutedEventArgs)
        Dim path = OpenImportFile()
        If String.IsNullOrEmpty(path) Then Return
        BtnImportMapping.IsEnabled = False
        BtnImportMapping.Content = "Importing..."
        ShowProgress(True)
        Dim imported As Integer = 0, skipped As Integer = 0
        Try
            Await Task.Run(Sub()
                Try
                    Dim dt = ReadImportTable(path)
                    If dt Is Nothing Then Return
                    Dim hi = FindHeaderRow(dt, "category", "Category", "type_name", "Type Name")
                    If hi = -1 Then Return
                    Dim idx = MapColumns(dt.Rows(hi),
                        "category:category,Category",
                        "type_name:type_name,TypeName,Type Name",
                        "prefix:prefix,Prefix",
                        "calibration_category:calibration_category,Calibration Category,CalibrationCategory")
                    For i = hi + 1 To dt.Rows.Count - 1
                        Dim row = dt.Rows(i)
                        Dim cat = If(idx.ContainsKey("category"), GetVal(row, idx("category")), "")
                        Dim tname = If(idx.ContainsKey("type_name"), GetVal(row, idx("type_name")), "")
                        If String.IsNullOrWhiteSpace(tname) Then skipped += 1 : Continue For
                        If _mySql.IsCalibMappingDuplicate(tname, cat) Then skipped += 1 : Continue For
                        Dim pre = If(idx.ContainsKey("prefix"), GetVal(row, idx("prefix")), "")
                        Dim cc = If(idx.ContainsKey("calibration_category"), GetVal(row, idx("calibration_category")), "")
                        If _mySql.InsertCalibrationMappingRecord(cat, tname, pre, cc) Then imported += 1 Else skipped += 1
                    Next
                Catch ex As Exception
                    Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                End Try
            End Sub)
            MessageBox.Show($"Import Complete!{Environment.NewLine}Added: {imported} | Skipped: {skipped}", "Calibration Mapping Import", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadCalibrationMapping()
        Catch ex As Exception
            MessageBox.Show("Import Error: " & ex.Message)
        Finally
            BtnImportMapping.IsEnabled = True
            BtnImportMapping.Content = "📤 Import"
            ShowProgress(False)
        End Try
    End Sub

    ' ========== CALIBRATION MASTER TYPES TAB ==========

    Private Sub BtnDownloadCalibMaster_Click(sender As Object, e As RoutedEventArgs)
        ExportDataGridToCSV(CalibMasterDataGrid, {"CalibrationMasterName"}, "CalibrationMasterTypes_Export")
    End Sub

    Private Async Sub BtnImportCalibMaster_Click(sender As Object, e As RoutedEventArgs)
        Dim path = OpenImportFile()
        If String.IsNullOrEmpty(path) Then Return
        BtnImportCalibMaster.IsEnabled = False
        BtnImportCalibMaster.Content = "Importing..."
        ShowProgress(True)
        Dim imported As Integer = 0, skipped As Integer = 0
        Try
            Await Task.Run(Sub()
                Try
                    Dim dt = ReadImportTable(path)
                    If dt Is Nothing Then Return
                    Dim hi = FindHeaderRow(dt, "CalibrationMasterName", "Calibration Master Type", "Calibration Master Name")
                    If hi = -1 Then Return
                    Dim idx = MapColumns(dt.Rows(hi), "CalibrationMasterName:CalibrationMasterName,Calibration Master Type,Calibration Master Name")
                    For i = hi + 1 To dt.Rows.Count - 1
                        Dim row = dt.Rows(i)
                        Dim name = If(idx.ContainsKey("CalibrationMasterName"), GetVal(row, idx("CalibrationMasterName")), "")
                        If String.IsNullOrWhiteSpace(name) Then skipped += 1 : Continue For
                        Dim masterName = MySQLClass.ToTitleCase(name)
                        If _mySql.IsCalibMasterDuplicate(masterName) Then skipped += 1 : Continue For
                        If _mySql.InsertCalibrationMasterType(masterName) Then
                            Dim tbl = MySQLClass.TypeNameToTableName(masterName)
                            _mySql.CreateCalibrationMasterTable(tbl)
                            imported += 1
                        Else
                            skipped += 1
                        End If
                    Next
                Catch ex As Exception
                    Application.Current.Dispatcher.Invoke(Sub() MessageBox.Show("Import Error: " & ex.Message))
                End Try
            End Sub)
            MessageBox.Show($"Import Complete!{Environment.NewLine}Added: {imported} | Skipped: {skipped}", "Calibration Master Types Import", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadCalibrationMasterTypes()
        Catch ex As Exception
            MessageBox.Show("Import Error: " & ex.Message)
        Finally
            BtnImportCalibMaster.IsEnabled = True
            BtnImportCalibMaster.Content = "📤 Import"
            ShowProgress(False)
        End Try
    End Sub

    ' ========== ADMIN TAB ==========

    ''' <summary>
    ''' Loads saved Cycle Timing Settings from system_config into the Admin tab controls.
    ''' Also loads the File Storage path from settings.json.
    ''' </summary>
    Private Sub LoadAdminSettings()
        ' --- File Storage Path ---
        Try
            Dim savedPath = ProjectSettings.Current.FileStorageBasePath
            If Not String.IsNullOrWhiteSpace(savedPath) AndAlso Directory.Exists(savedPath) Then
                TxtFileStoragePath.Text = System.IO.Path.Combine(savedPath, "Tool Life Cycle Management System Files")
                TxtFileStoragePath.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#15803D"), System.Windows.Media.Color))
            Else
                TxtFileStoragePath.Text = "(Not configured — using application default directory)"
                TxtFileStoragePath.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#3B82F6"), System.Windows.Media.Color))
            End If
        Catch ex As Exception
            ' Non-critical – just leave default text
        End Try

        ' --- Last Backup Timestamp ---
        Try
            Dim ts = ProjectSettings.Current.LastBackupTimestamp
            TxtLastBackup.Text = If(String.IsNullOrWhiteSpace(ts), "No backup created yet", ts)
        Catch
        End Try

        Try
            ' --- Start Colour ---
            Dim startColor = _mySql.GetConfigValue("CycleStartColor")
            CboStartColor.SelectedIndex = 0 ' Default: Green
            For i As Integer = 0 To CboStartColor.Items.Count - 1
                Dim item = TryCast(CboStartColor.Items(i), ComboBoxItem)
                If item IsNot Nothing AndAlso item.Tag?.ToString() = startColor Then
                    CboStartColor.SelectedIndex = i
                    Exit For
                End If
            Next

            ' --- Read month/day TEMPLATES from DB (year is ignored here) ------
            Dim gsVal = _mySql.GetConfigValue("GreenCycleStart")
            Dim geVal = _mySql.GetConfigValue("GreenCycleEnd")
            Dim ysVal = _mySql.GetConfigValue("YellowCycleStart")
            Dim yeVal = _mySql.GetConfigValue("YellowCycleEnd")

            ' ── FRESH INSTALL: no cycle dates configured yet ─────────────────
            ' When all four DB values are empty, pre-fill using Honda Standard:
            '   Green  = July 1  → December 31
            '   Yellow = January 1 → June 30
            ' Dates are derived from today's month/year. Nothing is auto-saved —
            ' the user must still click "Save Cycle Timing Settings".
            If String.IsNullOrEmpty(gsVal) AndAlso String.IsNullOrEmpty(geVal) AndAlso
               String.IsNullOrEmpty(ysVal) AndAlso String.IsNullOrEmpty(yeVal) Then

                Dim today As DateTime = DateTime.Today
                Dim cy As Integer = today.Year

                If today.Month >= 7 Then
                    ' Green is CURRENT (Jul–Dec this year), Yellow is UPCOMING (Jan–Jun next year)
                    DpGreenStart.SelectedDate = CType(New DateTime(cy, 7, 1), DateTime?)
                    DpGreenEnd.SelectedDate = CType(New DateTime(cy, 12, 31), DateTime?)
                    DpYellowStart.SelectedDate = CType(New DateTime(cy + 1, 1, 1), DateTime?)
                    DpYellowEnd.SelectedDate = CType(New DateTime(cy + 1, 6, 30), DateTime?)
                    LblGreenCycleStatus.Text = "[Current Cycle]"
                    LblGreenCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#15803D"), System.Windows.Media.Color))
                    LblYellowCycleStatus.Text = "[Upcoming Cycle]"
                    LblYellowCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#B45309"), System.Windows.Media.Color))
                Else
                    ' Yellow is CURRENT (Jan–Jun this year), Green is UPCOMING (Jul–Dec this year)
                    DpYellowStart.SelectedDate = CType(New DateTime(cy, 1, 1), DateTime?)
                    DpYellowEnd.SelectedDate = CType(New DateTime(cy, 6, 30), DateTime?)
                    DpGreenStart.SelectedDate = CType(New DateTime(cy, 7, 1), DateTime?)
                    DpGreenEnd.SelectedDate = CType(New DateTime(cy, 12, 31), DateTime?)
                    LblYellowCycleStatus.Text = "[Current Cycle]"
                    LblYellowCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#B45309"), System.Windows.Media.Color))
                    LblGreenCycleStatus.Text = "[Upcoming Cycle]"
                    LblGreenCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#15803D"), System.Windows.Media.Color))
                End If

            Else
                ' ── EXISTING INSTALL: DB values present — use stored dates ────
                Dim greenStartTpl As DateTime? = If(String.IsNullOrEmpty(gsVal), Nothing, CType(DateTime.Parse(gsVal), DateTime?))
                Dim greenEndTpl As DateTime? = If(String.IsNullOrEmpty(geVal), Nothing, CType(DateTime.Parse(geVal), DateTime?))
                Dim yellowStartTpl As DateTime? = If(String.IsNullOrEmpty(ysVal), Nothing, CType(DateTime.Parse(ysVal), DateTime?))
                Dim yellowEndTpl As DateTime? = If(String.IsNullOrEmpty(yeVal), Nothing, CType(DateTime.Parse(yeVal), DateTime?))

                ' --- Get active cycle name and extract its year -------------------
                '     Format: "Jan'28 Yellow Cycle" or "Jul'27 Green Cycle"
                Dim activeCycle = _mySql.GetActiveCycleName()
                Dim activeYear As Integer = DateTime.Today.Year ' safe fallback
                Try
                    Dim apos = activeCycle.IndexOf("'"c)
                    If apos >= 0 AndAlso activeCycle.Length > apos + 2 Then
                        activeYear = 2000 + Integer.Parse(activeCycle.Substring(apos + 1, 2))
                    End If
                Catch
                    ' keep fallback year
                End Try

                ' --- Assign DatePickers using active year + DB month/day ----------
                If activeCycle.Contains("Green") Then
                    ' ── Green = CURRENT (activeYear) ──────────────────────────────
                    '    e.g. Jul'27 Green Cycle → Green Jul–Dec 2027
                    If greenStartTpl.HasValue Then
                        DpGreenStart.SelectedDate = CType(New DateTime(activeYear, greenStartTpl.Value.Month, greenStartTpl.Value.Day), DateTime?)
                        DpGreenEnd.SelectedDate = CType(New DateTime(activeYear, greenEndTpl.Value.Month, greenEndTpl.Value.Day), DateTime?)
                    End If

                    ' ── Yellow = UPCOMING (activeYear + 1) ────────────────────────
                    '    e.g. Upcoming Yellow Jan–Jun 2028
                    If yellowStartTpl.HasValue Then
                        Dim upcomingYear = activeYear + 1
                        DpYellowStart.SelectedDate = CType(New DateTime(upcomingYear, yellowStartTpl.Value.Month, yellowStartTpl.Value.Day), DateTime?)
                        DpYellowEnd.SelectedDate = CType(New DateTime(upcomingYear, yellowEndTpl.Value.Month, yellowEndTpl.Value.Day), DateTime?)
                    End If

                    LblGreenCycleStatus.Text = "[Current Cycle]"
                    LblGreenCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#15803D"), System.Windows.Media.Color))
                    LblYellowCycleStatus.Text = "[Upcoming Cycle]"
                    LblYellowCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#B45309"), System.Windows.Media.Color))
                Else
                    ' ── Yellow = CURRENT (activeYear) ─────────────────────────────
                    '    e.g. Jan'28 Yellow Cycle → Yellow Jan–Jun 2028
                    If yellowStartTpl.HasValue Then
                        DpYellowStart.SelectedDate = CType(New DateTime(activeYear, yellowStartTpl.Value.Month, yellowStartTpl.Value.Day), DateTime?)
                        DpYellowEnd.SelectedDate = CType(New DateTime(activeYear, yellowEndTpl.Value.Month, yellowEndTpl.Value.Day), DateTime?)
                    End If

                    ' ── Green = UPCOMING (same activeYear, second half) ────────────
                    '    e.g. Upcoming Green Jul–Dec 2028
                    If greenStartTpl.HasValue Then
                        DpGreenStart.SelectedDate = CType(New DateTime(activeYear, greenStartTpl.Value.Month, greenStartTpl.Value.Day), DateTime?)
                        DpGreenEnd.SelectedDate = CType(New DateTime(activeYear, greenEndTpl.Value.Month, greenEndTpl.Value.Day), DateTime?)
                    End If

                    LblYellowCycleStatus.Text = "[Current Cycle]"
                    LblYellowCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#B45309"), System.Windows.Media.Color))
                    LblGreenCycleStatus.Text = "[Upcoming Cycle]"
                    LblGreenCycleStatus.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#15803D"), System.Windows.Media.Color))
                End If

            End If
        Catch ex As Exception
            MessageBox.Show("Error loading admin settings: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Validates and saves cycle timing settings to system_config.
    ''' </summary>
    Private Sub BtnSaveCycleTiming_Click(sender As Object, e As RoutedEventArgs)
        ' --- Validation ---
        If Not DpGreenStart.SelectedDate.HasValue OrElse Not DpGreenEnd.SelectedDate.HasValue Then
            MessageBox.Show("Please select both Start and End dates for the Green cycle.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        If Not DpYellowStart.SelectedDate.HasValue OrElse Not DpYellowEnd.SelectedDate.HasValue Then
            MessageBox.Show("Please select both Start and End dates for the Yellow cycle.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        If DpGreenEnd.SelectedDate.Value < DpGreenStart.SelectedDate.Value Then
            MessageBox.Show("Green Cycle End date must be on or after the Start date.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        If DpYellowEnd.SelectedDate.Value < DpYellowStart.SelectedDate.Value Then
            MessageBox.Show("Yellow Cycle End date must be on or after the Start date.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        ' Check overlap: Green and Yellow ranges must not overlap
        Dim greenStart = DpGreenStart.SelectedDate.Value
        Dim greenEnd = DpGreenEnd.SelectedDate.Value
        Dim yellowStart = DpYellowStart.SelectedDate.Value
        Dim yellowEnd = DpYellowEnd.SelectedDate.Value
        If greenStart <= yellowEnd AndAlso yellowStart <= greenEnd Then
            MessageBox.Show("Green and Yellow cycle date ranges must not overlap. Please adjust the dates.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' --- Determine Start Colour ---
        Dim selectedColor = "Green"
        Dim selectedItem = TryCast(CboStartColor.SelectedItem, ComboBoxItem)
        If selectedItem IsNot Nothing Then
            selectedColor = selectedItem.Tag?.ToString()
        End If

        ' --- Save to DB ---
        Try
            _mySql.SetConfigValue("CycleStartColor", selectedColor)
            _mySql.SetConfigValue("GreenCycleStart", greenStart.ToString("yyyy-MM-dd"))
            _mySql.SetConfigValue("GreenCycleEnd", greenEnd.ToString("yyyy-MM-dd"))
            _mySql.SetConfigValue("YellowCycleStart", yellowStart.ToString("yyyy-MM-dd"))
            _mySql.SetConfigValue("YellowCycleEnd", yellowEnd.ToString("yyyy-MM-dd"))

            ' Build cycle name preview using actual month from the configured start dates
            Dim greenCycleName = $"{greenStart.ToString("MMM")}'{greenStart.ToString("yy")} Green Cycle"
            Dim yellowCycleName = $"{yellowStart.ToString("MMM")}'{yellowStart.ToString("yy")} Yellow Cycle"

            Dim msg As String = $"Cycle Timing Settings saved successfully!" & vbCrLf & vbCrLf &
                                $"🟢 Green Cycle : {greenCycleName}" & vbCrLf &
                                $"   ({greenStart:dd-MMM-yyyy}  to  {greenEnd:dd-MMM-yyyy})" & vbCrLf &
                                $"🟡 Yellow Cycle: {yellowCycleName}" & vbCrLf &
                                $"   ({yellowStart:dd-MMM-yyyy}  to  {yellowEnd:dd-MMM-yyyy})" & vbCrLf & vbCrLf &
                                $"▶ Start Colour after Reset All: {selectedColor}"
            MessageBox.Show(msg, "Settings Saved", MessageBoxButton.OK, MessageBoxImage.Information)
        Catch ex As Exception
            MessageBox.Show("Error saving settings: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
    ' ========== FILE LOCATION SECTION ==========

    ''' <summary>
    ''' Opens a folder browser so the user can pick the storage root directory.
    ''' Uses an OpenFileDialog trick for WPF-native folder selection.
    ''' </summary>
    Private Sub BtnBrowseFileStorage_Click(sender As Object, e As RoutedEventArgs)
        ' WPF folder-selection trick: use OpenFileDialog with ValidateNames = False
        Dim dlg As New Microsoft.Win32.OpenFileDialog()
        dlg.Title = "Select Root Directory for Tool Life Cycle Management System Files"
        dlg.CheckFileExists = False
        dlg.CheckPathExists = True
        dlg.ValidateNames = False
        dlg.FileName = "Select Folder"
        dlg.Filter = "Folder|*.none"

        ' Pre-select existing path
        Dim current = ProjectSettings.Current.FileStorageBasePath
        If Not String.IsNullOrWhiteSpace(current) AndAlso Directory.Exists(current) Then
            dlg.InitialDirectory = current
        End If

        If dlg.ShowDialog() = True Then
            ' The chosen "file" path gives us the directory
            Dim chosen = System.IO.Path.GetDirectoryName(dlg.FileName)
            If String.IsNullOrWhiteSpace(chosen) Then
                chosen = System.IO.Path.GetPathRoot(dlg.FileName)
            End If
            If Directory.Exists(chosen) Then
                ' Update UI
                TxtFileStoragePath.Text = System.IO.Path.Combine(chosen, "Tool Life Cycle Management System Files")
                TxtFileStoragePath.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#15803D"), System.Windows.Media.Color))

                ' Save path and immediately create the full folder structure
                ProjectSettings.Current.FileStorageBasePath = chosen
                ProjectSettings.Save()
                ProjectSettings.EnsureFileStorageFolders()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Clears the configured file storage path, reverting to the Documents default.
    ''' </summary>
    Private Sub BtnClearFileStorage_Click(sender As Object, e As RoutedEventArgs)
        Dim confirm = MessageBox.Show(
            "Are you sure you want to clear the configured file storage path?" & vbCrLf & vbCrLf &
            "The application will revert to the default Documents folder." & vbCrLf &
            "Note: Files already stored in the current location will NOT be moved.",
            "Clear File Storage Path",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning)

        If confirm = MessageBoxResult.Yes Then
            ProjectSettings.Current.FileStorageBasePath = ""
            ProjectSettings.Save()
            TxtFileStoragePath.Text = "(Not configured — using Documents default directory)"
            TxtFileStoragePath.Foreground = New System.Windows.Media.SolidColorBrush(DirectCast(System.Windows.Media.ColorConverter.ConvertFromString("#3B82F6"), System.Windows.Media.Color))
        End If
    End Sub

    ''' <summary>
    ''' Persists the chosen file storage path and creates the full folder hierarchy.
    ''' </summary>
    Private Sub BtnSaveFileLocation_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim basePath = ProjectSettings.Current.FileStorageBasePath

            If Not String.IsNullOrWhiteSpace(basePath) AndAlso Not Directory.Exists(basePath) Then
                MessageBox.Show("The selected folder does not exist. Please browse and select a valid directory.", "Invalid Path", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            ' Save to settings.json
            ProjectSettings.Save()

            ' Create folder structure
            ProjectSettings.EnsureFileStorageFolders()

            Dim rootDisplay = If(String.IsNullOrWhiteSpace(basePath),
                                 AppDomain.CurrentDomain.BaseDirectory,
                                 System.IO.Path.Combine(basePath, "Tool Life Cycle Management System Files"))

            MessageBox.Show($"File location saved successfully!{Environment.NewLine}{Environment.NewLine}" &
                            $"📁 Root Folder:{Environment.NewLine}   {rootDisplay}{Environment.NewLine}{Environment.NewLine}" &
                            $"All required sub-folders have been created.",
                            "File Location Saved", MessageBoxButton.OK, MessageBoxImage.Information)
        Catch ex As Exception
            MessageBox.Show("Error saving file location: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' ========== BACKUP & RESTORE SECTION ==========

    ''' <summary>Appends a line to the backup log TextBox (thread-safe via Dispatcher).</summary>
    Private Sub BackupManager_ProgressChanged(message As String)
        Dispatcher.Invoke(Sub()
            TxtBackupLog.Text &= If(TxtBackupLog.Text = "Ready. Press Create Backup or Restore Backup to begin.",
                                    message,
                                    Environment.NewLine & message)
            BackupLogScroll.ScrollToBottom()
            TxtLastBackup.Text = If(String.IsNullOrWhiteSpace(ProjectSettings.Current.LastBackupTimestamp),
                                    "No backup created yet",
                                    ProjectSettings.Current.LastBackupTimestamp)
        End Sub)
    End Sub

    ''' <summary>Creates a full .hbak backup archive chosen by the user.</summary>
    Private Async Sub BtnCreateBackup_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New Microsoft.Win32.SaveFileDialog()
        dlg.Title = "Save Backup File"
        dlg.Filter = "TLCMS Backup (*.hbak)|*.hbak"
        dlg.FileName = $"TLCMS_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.hbak"

        If dlg.ShowDialog() <> True Then Return

        ' Disable buttons while running
        BtnCreateBackup.IsEnabled = False
        BtnRestoreBackup.IsEnabled = False
        TxtBackupLog.Text = ""

        AddHandler BackupManager.ProgressChanged, AddressOf BackupManager_ProgressChanged
        Try
            Await BackupManager.CreateBackupAsync(dlg.FileName)
        Finally
            RemoveHandler BackupManager.ProgressChanged, AddressOf BackupManager_ProgressChanged
            BtnCreateBackup.IsEnabled = True
            BtnRestoreBackup.IsEnabled = True
        End Try
    End Sub

    ''' <summary>Restores a .hbak backup archive chosen by the user.</summary>
    Private Async Sub BtnRestoreBackup_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New Microsoft.Win32.OpenFileDialog()
        dlg.Title = "Select Backup File to Restore"
        dlg.Filter = "TLCMS Backup (*.hbak)|*.hbak"
        dlg.CheckFileExists = True

        If dlg.ShowDialog() <> True Then Return

        Dim confirm = MessageBox.Show(
            $"⚠️  WARNING — This will REPLACE all existing database records and files with the contents of:{Environment.NewLine}{Environment.NewLine}" &
            $"   {dlg.FileName}{Environment.NewLine}{Environment.NewLine}" &
            "This cannot be undone. Are you sure you want to continue?",
            "Confirm Restore",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning)

        If confirm <> MessageBoxResult.Yes Then Return

        BtnCreateBackup.IsEnabled = False
        BtnRestoreBackup.IsEnabled = False
        TxtBackupLog.Text = ""

        AddHandler BackupManager.ProgressChanged, AddressOf BackupManager_ProgressChanged
        Try
            Await BackupManager.RestoreBackupAsync(dlg.FileName)
        Finally
            RemoveHandler BackupManager.ProgressChanged, AddressOf BackupManager_ProgressChanged
            BtnCreateBackup.IsEnabled = True
            BtnRestoreBackup.IsEnabled = True
        End Try
    End Sub

End Class

