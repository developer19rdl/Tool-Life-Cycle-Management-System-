Imports System.Windows
Imports System.Windows.Controls
Imports System.IO.Ports
Imports System.Data

Public Class AddEditInstrumentWindow
    Private _mySql As New MySQLClass()
    Private _currentCategory As String = ""
    Private _isLoading As Boolean = True
    Private _tableName As String = "addinstrument" ' fallback default
    Private _uploadedDocPath As String = ""
    
    Public Property IsEditMode As Boolean = False
    Public Property RecordID As Integer = 0
    Public Property InstrumentType As String = "Instrument" ' Default

    Public Sub New(Optional tableName As String = "addinstrument")
        InitializeComponent()
        _tableName = tableName
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        If IsEditMode Then
            BtnSave.Content = "Update"
            Me.Title = "Edit Instrument Details"
        Else
            _isLoading = True
            LoadComboBoxData()
            BtnSave.Content = "Add"
            Me.Title = "Add New Instrument"
            ' Set default date
            DatePickerDate.SelectedDate = DateTime.Now

            ' Pre-fill description if name is already set
            If Not String.IsNullOrEmpty(TxtName.Text) AndAlso String.IsNullOrEmpty(TxtDescription.Text) Then
                TxtDescription.Text = TxtName.Text
            End If
            _isLoading = False
        End If
    End Sub

    Private Sub LoadComboBoxData()
        Try
            ' Populate Department
            Dim dtDept = _mySql.GetDepartments()
            If dtDept IsNot Nothing AndAlso dtDept.Rows.Count > 0 Then
                ComboSection.ItemsSource = dtDept.DefaultView
                ComboSection.DisplayMemberPath = "DepartmentName"
                ComboSection.SelectedValuePath = "DepartmentName"
            Else
                Console.WriteLine("Warning: No departments found in database.")
            End If

            ' Populate Size
            Dim sizeQuery As String = $"SELECT DISTINCT Size, GroupCode FROM categorycontrol WHERE InstrumentType = '{InstrumentType.Replace("'", "''")}' AND Active = 1 ORDER BY Sort, Size"
            Dim dtSizes = _mySql.ReadDatatable(sizeQuery)
            If dtSizes IsNot Nothing Then
                dtSizes.Columns.Add("DisplayString", GetType(String))
                For Each row As DataRow In dtSizes.Rows
                    row("DisplayString") = $"{row("GroupCode")} | {row("Size")}"
                Next
                ComboSize.ItemsSource = dtSizes.DefaultView
                ComboSize.DisplayMemberPath = "DisplayString"
                ComboSize.SelectedValuePath = "Size"
            End If


        Catch ex As Exception
            Console.WriteLine("Load Combo Error: " & ex.Message)
            MessageBox.Show("Error loading dropdown data: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub ResolveCategoryAndControlNo()
        If _isLoading Then Return
        If String.IsNullOrWhiteSpace(TxtName.Text) OrElse ComboSize.SelectedValue Is Nothing Then
            TxtControlNo.Clear()
            Return
        End If

        Try
            Dim selectedSize = ComboSize.SelectedValue.ToString()

            ' 1. Lookup Phase: Get GroupCode from categorycontrol by size + instrument type
            Dim groupCode = _mySql.GetGroupCodeBySize(selectedSize, InstrumentType)
            If String.IsNullOrEmpty(groupCode) Then
                TxtControlNo.Clear()
                Return
            End If

            ' 2. ID Construction: GroupCode IS the Base ID (e.g. J1 → J1-001)
            '    BasePrefix from type_details is NOT concatenated here — it was designed for gauges.
            '    For instruments, the user's spec says GroupCode = Base ID directly.
            Dim fullPrefix = groupCode
            _currentCategory = fullPrefix

            ' 3. Search & Increment: find highest existing sequence for this prefix in the per-type table
            Dim lastNum = _mySql.GetLastControlNumber(fullPrefix, _tableName)
            TxtControlNo.Text = $"{fullPrefix}-{(lastNum + 1).ToString("D3")}"
        Catch ex As Exception
            Console.WriteLine("Resolve Error: " & ex.Message)
        End Try
    End Sub

    Private Sub TxtName_LostFocus(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(TxtDescription.Text) Then
            TxtDescription.Text = TxtName.Text
        End If
        ResolveCategoryAndControlNo()
    End Sub

    Private Sub ComboSize_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If Not _isLoading Then
            ResolveCategoryAndControlNo()
        End If
    End Sub

    Private Async Sub BtnScanRFID_Click(sender As Object, e As RoutedEventArgs)
        Try
            TxtRFID.Text = "Scanning..."
            
            Dim scannedTag = Await Task.Run(Function() CommManager.ScanSingleTag(Nothing, 5000))

            If String.IsNullOrWhiteSpace(scannedTag) Then
                TxtRFID.Clear()
                MessageBox.Show("No RFID tag detected. Please place the tag near the reader and try again.", "Scan Failed", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            TxtRFID.Text = scannedTag

        Catch ex As Exception
            MessageBox.Show("RFID Scan Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub


    Private Sub BtnUpload_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
            openFileDialog.Filter = "All Supported Files|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.jpg;*.jpeg;*.png;*.bmp|PDF Files (*.pdf)|*.pdf|Word Documents (*.doc;*.docx)|*.doc;*.docx|Excel Spreadsheets (*.xls;*.xlsx)|*.xls;*.xlsx|Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*"

            If openFileDialog.ShowDialog() = True Then
                Dim sourceFile = openFileDialog.FileName
                ' Store relative path using centralized helper
                _uploadedDocPath = MySQLClass.CopyFileToDocuments(sourceFile, "Instrument", TxtControlNo.Text)
                
                If Not String.IsNullOrEmpty(_uploadedDocPath) Then
                    TxtFileName.Text = System.IO.Path.GetFileName(sourceFile)
                Else
                    MessageBox.Show("Failed to copy file to Documents folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("File Upload Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        ' Validation
        If String.IsNullOrWhiteSpace(TxtName.Text) Then
            MessageBox.Show("Instrument Name is required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim success As Boolean = False
        Dim colorText As String = ""
        If ComboColor.SelectedItem IsNot Nothing Then
            colorText = DirectCast(ComboColor.SelectedItem, ComboBoxItem).Content.ToString()
        End If

        Dim sectionVal = If(ComboSection.SelectedValue IsNot Nothing, ComboSection.SelectedValue.ToString(), "")
        Dim locationVal = sectionVal ' Use the same value for both as there is only one dropdown in UI
        Dim sizeVal = If(ComboSize.SelectedValue IsNot Nothing, ComboSize.SelectedValue.ToString(), "")

        Dim customDate As Date? = DatePickerDate.SelectedDate

        If IsEditMode Then
            success = _mySql.Updateinstrument(_tableName, RecordID, TxtName.Text, TxtDescription.Text, TxtControlNo.Text, colorText, TxtModel.Text, TxtLine.Text, TxtAsset.Text, sizeVal, TxtMaker.Text, sectionVal, locationVal, TxtRequestNo.Text, TxtRemark.Text, _currentCategory, TxtRFID.Text, _uploadedDocPath, 0, customDate)
        Else
            ' Check for duplicate RFID if provided
            If Not String.IsNullOrWhiteSpace(TxtRFID.Text) Then
                Dim dtRfid = _mySql.ReadDatatable($"SELECT ID FROM `{_tableName}` WHERE RFID_tag = '{TxtRFID.Text}'")
                If dtRfid.Rows.Count > 0 Then
                    MessageBox.Show("This RFID tag is already assigned to another instrument.", "Duplicate RFID", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return
                End If
            End If

            success = _mySql.Insertinstrument(_tableName, TxtName.Text, TxtDescription.Text, TxtControlNo.Text, colorText, TxtModel.Text, TxtLine.Text, TxtAsset.Text, sizeVal, TxtMaker.Text, sectionVal, locationVal, TxtRequestNo.Text, TxtRemark.Text, _currentCategory, TxtRFID.Text, _uploadedDocPath, 0, customDate)

            ' Automated Sync to Interchangeability for NEW records
            If success Then
                Dim activeCycle = _mySql.GetActiveCycleName()
                If Not String.IsNullOrEmpty(activeCycle) Then
                    ' 1. Add to department_list
                    Dim batchTag = "Add " & DateTime.Now.ToString("yyyyMMdd")
                    _mySql.InsertDeptListItem(locationVal, TxtName.Text, sizeVal, TxtControlNo.Text, colorText, "Pending", TxtRemark.Text, batchTag)

                    ' 2. Add to interchangeability
                    _mySql.InsertInterchangeRecord(activeCycle, TxtControlNo.Text, locationVal, TxtName.Text, sizeVal, colorText, "Pending", TxtRemark.Text)

                    ' 3. Add to new_addition_calibration
                    _mySql.AddNewAdditionCalibration(TxtName.Text, TxtControlNo.Text, "Pending", activeCycle, TxtRequestNo.Text)
                End If
            End If
        End If

        If success Then
            _mySql.SyncCategoryControlFromInventory()
            MessageBox.Show("Record processed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save record. Please check database connection.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub SetComboValue(cb As ComboBox, target As String)
        If String.IsNullOrEmpty(target) Then Return
        Dim cleanTarget = target.TrimEnd(";"c, " "c).Trim()

        ' Try direct assignment first
        cb.SelectedValue = cleanTarget
        If cb.SelectedValue IsNot Nothing AndAlso cb.SelectedValue.ToString().Equals(cleanTarget, StringComparison.OrdinalIgnoreCase) Then Return

        ' Scanning fallback
        For i As Integer = 0 To cb.Items.Count - 1
            Dim itm = cb.Items(i)
            Dim itmVal As String = ""

            If TypeOf itm Is DataRowView Then
                Dim drv = DirectCast(itm, DataRowView)
                If Not String.IsNullOrEmpty(cb.SelectedValuePath) Then
                    Try
                        itmVal = drv(cb.SelectedValuePath).ToString()
                    Catch
                        itmVal = itm.ToString()
                    End Try
                End If
            ElseIf TypeOf itm Is ComboBoxItem Then
                itmVal = DirectCast(itm, ComboBoxItem).Content.ToString()
            Else
                itmVal = itm.ToString()
            End If

            If itmVal.TrimEnd(";"c, " "c).Trim().Equals(cleanTarget, StringComparison.OrdinalIgnoreCase) Then
                cb.SelectedIndex = i
                Return
            End If
        Next
    End Sub

    Public Sub PopulateForm(row As DataRow)
        If row Is Nothing Then Return

        _isLoading = True
        IsEditMode = True

        ' Ensure dropdowns are loaded BEFORE setting values
        LoadComboBoxData()

        RecordID = If(row.Table.Columns.Contains("ID"), Convert.ToInt32(row("ID")), 0)

        If row.Table.Columns.Contains("InstrumentName") Then TxtName.Text = row("InstrumentName").ToString()
        If row.Table.Columns.Contains("InstrumentDescription") Then TxtDescription.Text = row("InstrumentDescription").ToString()
        If row.Table.Columns.Contains("ControlNo") Then TxtControlNo.Text = row("ControlNo").ToString()

        ' Handle ComboBox selection for Color
        If row.Table.Columns.Contains("Color") Then
            Dim colorVal = row("Color").ToString()
            For Each item As ComboBoxItem In ComboColor.Items
                If item.Content.ToString() = colorVal Then
                    ComboColor.SelectedItem = item
                    Exit For
                End If
            Next
        End If

        If row.Table.Columns.Contains("InstrumentModelNo") Then TxtModel.Text = row("InstrumentModelNo").ToString()
        If row.Table.Columns.Contains("Line") Then TxtLine.Text = row("Line").ToString()
        If row.Table.Columns.Contains("AssetNo") Then TxtAsset.Text = row("AssetNo").ToString()

        ' ComboBoxes
        If row.Table.Columns.Contains("Size") Then
            SetComboValue(ComboSize, row("Size").ToString())
        End If

        ' Robust check for Section/Location
        Dim dbSection = If(row.Table.Columns.Contains("Section"), row("Section").ToString(), "")
        Dim dbLocation = If(row.Table.Columns.Contains("Location"), row("Location").ToString(), "")
        Dim targetLoc = If(Not String.IsNullOrEmpty(dbSection), dbSection, dbLocation)
        SetComboValue(ComboSection, targetLoc)

        If row.Table.Columns.Contains("MakerName") Then TxtMaker.Text = row("MakerName").ToString().Trim()
        If row.Table.Columns.Contains("RequestNo") Then TxtRequestNo.Text = row("RequestNo").ToString().Trim()
        If row.Table.Columns.Contains("Remark") Then TxtRemark.Text = row("Remark").ToString().Trim()
        If row.Table.Columns.Contains("CategoryControl") Then _currentCategory = row("CategoryControl").ToString().Trim()
        If row.Table.Columns.Contains("RFID_tag") Then TxtRFID.Text = row("RFID_tag").ToString().Trim()

        If row.Table.Columns.Contains("uploaded_doc") Then
            _uploadedDocPath = row("uploaded_doc").ToString()
            If Not String.IsNullOrEmpty(_uploadedDocPath) Then
                TxtFileName.Text = System.IO.Path.GetFileName(_uploadedDocPath)
            End If
        End If

        ' Populate Date parts
        If row.Table.Columns.Contains("Date") AndAlso row("Date") IsNot DBNull.Value Then
            Try
                DatePickerDate.SelectedDate = Convert.ToDateTime(row("Date"))
            Catch ex As Exception
                DatePickerDate.SelectedDate = Nothing
            End Try
        End If

        _isLoading = False
    End Sub
End Class
