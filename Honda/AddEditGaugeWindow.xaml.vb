Imports System.Windows
Imports System.Windows.Controls
Imports System.IO.Ports
Imports System.Data

Public Class AddEditGaugeWindow
    Private _mySql As New MySQLClass()
    Private _currentCategory As String = ""
    Private _tableName As String = "addgauge" ' fallback default
    Private _uploadedDocPath As String = ""
    Private _isLoading As Boolean = False
    
    Public Property IsEditMode As Boolean = False
    Public Property RecordID As Integer = 0
    Public Property InstrumentType As String = "Gauge" ' Fixed to Gauge
    Public Property DefaultGaugeName As String = ""

    Public Sub New(Optional tableName As String = "addgauge")
        InitializeComponent()
        _tableName = tableName
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        If IsEditMode Then
            BtnSave.Content = "Update"
            Me.Title = "Edit Gauge Details"
        Else
            _isLoading = True
            LoadComboBoxData()
            BtnSave.Content = "Add"
            Me.Title = "Add New Gauge"
            ' Set default date
            DatePickerDate.SelectedDate = DateTime.Now
            
            If Not String.IsNullOrEmpty(DefaultGaugeName) Then
                TxtName.Text = DefaultGaugeName
            End If
            
            ' Pre-fill description if name is already set
            If Not String.IsNullOrEmpty(TxtName.Text) AndAlso String.IsNullOrEmpty(TxtDescription.Text) Then
                TxtDescription.Text = TxtName.Text
            End If

            CheckGaugeSizeConfiguration()
            _isLoading = False
        End If
    End Sub

    Private Sub LoadComboBoxData()
        Try
            Dim dtDept = _mySql.GetDepartments()
            If dtDept IsNot Nothing AndAlso dtDept.Rows.Count > 0 Then
                ComboLocation.ItemsSource = dtDept.DefaultView
                ComboLocation.DisplayMemberPath = "DepartmentName"
                ComboLocation.SelectedValuePath = "DepartmentName"
            End If
        Catch ex As Exception
            Console.WriteLine("Load Combo Error: " & ex.Message)
            MessageBox.Show("Error loading dropdown data: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub CheckGaugeSizeConfiguration()
        If String.IsNullOrWhiteSpace(TxtName.Text) Then Return
        
        Try
            Dim gaugeType = TxtName.Text.Trim()
            Dim sizeQuery As String = $"SELECT DISTINCT Size, GroupCode FROM gauge_categorycontrol WHERE InstrumentType = '{gaugeType.Replace("'", "''")}' AND Active = 1 ORDER BY Sort, Size"
            Dim dtSizes = _mySql.ReadDatatable(sizeQuery)

            If dtSizes IsNot Nothing AndAlso dtSizes.Rows.Count > 0 Then
                dtSizes.Columns.Add("DisplayString", GetType(String))
                For Each row As DataRow In dtSizes.Rows
                    row("DisplayString") = $"{row("GroupCode")} | {row("Size")}"
                Next
                ComboSize.ItemsSource = dtSizes.DefaultView
                ComboSize.DisplayMemberPath = "DisplayString"
                ComboSize.SelectedValuePath = "Size"
                ComboSize.Visibility = Visibility.Visible
                TxtSize.Visibility = Visibility.Collapsed
            Else
                ComboSize.ItemsSource = Nothing
                ComboSize.Visibility = Visibility.Collapsed
                TxtSize.Visibility = Visibility.Visible
            End If
        Catch ex As Exception
            Console.WriteLine("CheckGaugeSize Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ResolveCategoryAndControlNo()
        If _isLoading Then Return
        If String.IsNullOrWhiteSpace(TxtName.Text) Then
            _currentCategory = ""
            TxtControlNo.Clear()
            Return
        End If

        Try
            Dim gaugeName = TxtName.Text.Trim()

            If ComboSize.Visibility = Visibility.Visible Then
                If ComboSize.SelectedValue Is Nothing Then
                    TxtControlNo.Clear()
                    _currentCategory = ""
                    Return
                End If
                Dim selectedSize = ComboSize.SelectedValue.ToString()
                Dim groupCode = _mySql.GetGaugeGroupCodeBySize(selectedSize, gaugeName)
                
                If String.IsNullOrEmpty(groupCode) Then
                    TxtControlNo.Clear()
                    Return
                End If
                
                Dim cat = groupCode
                Dim lastNum = _mySql.GetLastControlNumber(cat, _tableName)
                _currentCategory = cat
                TxtControlNo.Text = $"{cat}-{(lastNum + 1).ToString("D3")}"
            Else
                Dim sizeStr = TxtSize.Text.Trim()
                If String.IsNullOrWhiteSpace(sizeStr) Then
                    TxtControlNo.Clear()
                    _currentCategory = ""
                    Return
                End If

                ' STEP 1: Get the BasePrefix for this gauge from type_details (e.g. PS for PLAIN PLUG GAUGE)
                Dim basePrefix As String = _mySql.GetGaugeBasePrefix(gaugeName)
                
                If String.IsNullOrEmpty(basePrefix) Then
                    MessageBox.Show($"No type definition found for '{gaugeName}'. Please add it in Settings > Types first.", "Gauge Type Not Found", MessageBoxButton.OK, MessageBoxImage.Warning)
                    TxtControlNo.Clear()
                    _currentCategory = ""
                    Return
                End If

                ' STEP 2: Extract the full integer (whole number) part of size (e.g. "57.4" -> "57", "125.9" -> "125")
                Dim wholeNum As String = ""
                Dim dotPos As Integer = sizeStr.IndexOf("."c)
                Dim numericPart As String = If(dotPos >= 0, sizeStr.Substring(0, dotPos), sizeStr).Trim()
                ' Only digits allowed
                For Each ch As Char In numericPart
                    If Char.IsDigit(ch) Then wholeNum &= ch
                Next

                If String.IsNullOrEmpty(wholeNum) Then
                    MessageBox.Show("Invalid size value. Please enter a numeric size.", "Invalid Size", MessageBoxButton.OK, MessageBoxImage.Warning)
                    TxtControlNo.Clear()
                    _currentCategory = ""
                    Return
                End If

                ' STEP 3: Combine to build candidate category (e.g. PS + 57 = PS57)
                Dim cat As String = basePrefix & wholeNum

                ' STEP 4: Check per-type table if records already exist for this category prefix
                Dim lastNum = _mySql.GetLastControlNumber(cat, _tableName)

                If lastNum = 0 Then
                    ' No records found - this is a new prefix/size combination
                    Dim result = MessageBox.Show($"No records exist for prefix '{cat}'. Would you like to create it and start at {cat}-001?", "New Prefix Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question)
                    If result = MessageBoxResult.No Then
                        TxtControlNo.Clear()
                        _currentCategory = ""
                        Return
                    End If
                End If

                ' STEP 5: Assign control number
                _currentCategory = cat
                TxtControlNo.Text = $"{cat}-{(lastNum + 1).ToString("D3")}"
            End If
        Catch ex As Exception
            Console.WriteLine("Resolve Error: " & ex.Message)
        End Try
    End Sub

    Private Sub TxtName_LostFocus(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(TxtDescription.Text) Then
            TxtDescription.Text = TxtName.Text
        End If
        CheckGaugeSizeConfiguration()
        ResolveCategoryAndControlNo()
    End Sub

    Private Sub TxtSize_LostFocus(sender As Object, e As RoutedEventArgs)
        If TxtSize.Visibility = Visibility.Visible Then
            ResolveCategoryAndControlNo()
        End If
    End Sub

    Private Sub ComboSize_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If ComboSize.Visibility = Visibility.Visible Then
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
                _uploadedDocPath = MySQLClass.CopyFileToDocuments(sourceFile, "Gauge", TxtControlNo.Text)
                
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
            MessageBox.Show("Gauge Name is required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim success As Boolean = False
        Dim colorText As String = "White"
        If ComboColor.SelectedItem IsNot Nothing Then
            colorText = DirectCast(ComboColor.SelectedItem, ComboBoxItem).Content.ToString()
        End If

        Dim dateAdditionStr = ""
        If DatePickerDate.SelectedDate.HasValue Then
            dateAdditionStr = DatePickerDate.SelectedDate.Value.ToString("dd.MM.yy")
        End If

        Dim locationVal = If(ComboLocation.SelectedValue IsNot Nothing, ComboLocation.SelectedValue.ToString(), "")
        
        Dim sizeVal As String = ""
        If ComboSize.Visibility = Visibility.Visible Then
            sizeVal = If(ComboSize.SelectedValue IsNot Nothing, ComboSize.SelectedValue.ToString(), "")
            If String.IsNullOrWhiteSpace(sizeVal) Then
                MessageBox.Show("Size is required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
        Else
            sizeVal = TxtSize.Text.Trim()
            If String.IsNullOrWhiteSpace(sizeVal) Then
                MessageBox.Show("Size is required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
        End If

        Dim customDate As Date? = DatePickerDate.SelectedDate

        If IsEditMode Then
            success = _mySql.Updategauge(_tableName, RecordID, TxtName.Text, TxtDescription.Text, TxtControlNo.Text, colorText, TxtDrgNo.Text, TxtMaker.Text, TxtModel.Text, TxtLine.Text, sizeVal, TxtTol.Text, TxtTol2.Text, TxtTol3.Text, TxtSection.Text, locationVal, dateAdditionStr, TxtRequestNo.Text, TxtRemark.Text, _currentCategory, TxtRFID.Text, _uploadedDocPath, 0, customDate)
        Else
            ' Check for duplicate RFID if provided
            If Not String.IsNullOrWhiteSpace(TxtRFID.Text) Then
                Dim dtRfid = _mySql.ReadDatatable($"SELECT ID FROM `{_tableName}` WHERE RFID_tag = '{TxtRFID.Text}'")
                If dtRfid.Rows.Count > 0 Then
                    MessageBox.Show("This RFID tag is already assigned to another gauge.", "Duplicate RFID", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return
                End If
            End If

            success = _mySql.Insertgauge(_tableName, TxtName.Text, TxtDescription.Text, TxtControlNo.Text, colorText, TxtDrgNo.Text, TxtMaker.Text, TxtModel.Text, TxtLine.Text, sizeVal, TxtTol.Text, TxtTol2.Text, TxtTol3.Text, TxtSection.Text, locationVal, dateAdditionStr, TxtRequestNo.Text, TxtRemark.Text, _currentCategory, TxtRFID.Text, _uploadedDocPath, 0, customDate)
            
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

        ' Scanning fallback for cases where SelectedValue assignment fails or timing is off
        For i As Integer = 0 To cb.Items.Count - 1
            Dim itm = cb.Items(i)
            Dim itmVal As String = ""
            
            If TypeOf itm Is DataRowView Then
                Dim drv = DirectCast(itm, DataRowView)
                If Not String.IsNullOrEmpty(cb.SelectedValuePath) Then
                    Try
                        itmVal = drv(cb.SelectedValuePath).ToString()
                    Catch
                        ' Fallback to string if path fails
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
        
        If row.Table.Columns.Contains("GaugeName") Then TxtName.Text = row("GaugeName").ToString()
        If row.Table.Columns.Contains("GaugeDescription") Then TxtDescription.Text = row("GaugeDescription").ToString()
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

        If row.Table.Columns.Contains("Model") Then TxtModel.Text = row("Model").ToString()
        If row.Table.Columns.Contains("Line") Then TxtLine.Text = row("Line").ToString()

        ' Handle Size (Combo or Text)
        CheckGaugeSizeConfiguration()
        If row.Table.Columns.Contains("Size") Then
            If ComboSize.Visibility = Visibility.Visible Then
                SetComboValue(ComboSize, row("Size").ToString())
            Else
                TxtSize.Text = row("Size").ToString().Trim()
            End If
        End If
        
        If row.Table.Columns.Contains("MakerName") Then TxtMaker.Text = row("MakerName").ToString().Trim()
        If row.Table.Columns.Contains("DrgNo") Then TxtDrgNo.Text = row("DrgNo").ToString().Trim()
        If row.Table.Columns.Contains("Tol") Then TxtTol.Text = row("Tol").ToString().Trim()
        If row.Table.Columns.Contains("Tol2") Then TxtTol2.Text = row("Tol2").ToString().Trim()
        If row.Table.Columns.Contains("Tol3") Then TxtTol3.Text = row("Tol3").ToString().Trim()
        
        If row.Table.Columns.Contains("Section") Then TxtSection.Text = row("Section").ToString().Trim()
        
        ' Handle Location ComboBox
        Dim dbLoc = If(row.Table.Columns.Contains("Location"), row("Location").ToString(), "")
        Dim dbSec = If(row.Table.Columns.Contains("Section"), row("Section").ToString(), "")
        Dim targetLoc = If(Not String.IsNullOrEmpty(dbLoc), dbLoc, dbSec)
        SetComboValue(ComboLocation, targetLoc)
        
        If row.Table.Columns.Contains("RequestNo") Then TxtRequestNo.Text = row("RequestNo").ToString()
        If row.Table.Columns.Contains("Remark") Then TxtRemark.Text = row("Remark").ToString()
        If row.Table.Columns.Contains("CategoryControl") Then _currentCategory = row("CategoryControl").ToString()
        If row.Table.Columns.Contains("RFID_tag") Then TxtRFID.Text = row("RFID_tag").ToString()
        
        If row.Table.Columns.Contains("uploaded_doc") Then
            _uploadedDocPath = row("uploaded_doc").ToString()
            if Not String.IsNullOrEmpty(_uploadedDocPath) Then
                TxtFileName.Text = System.IO.Path.GetFileName(_uploadedDocPath)
            End If
        End If

        ' Handle DateofAddition mapping to DatePicker
        If row.Table.Columns.Contains("DateofAddition") Then
            Dim dateAddition = row("DateofAddition").ToString()
            If Not String.IsNullOrWhiteSpace(dateAddition) Then
                Try
                    ' Attempt parsing "dd.MM.yy" format explicitly
                    Dim parsedDate As DateTime
                    If DateTime.TryParseExact(dateAddition, "dd.MM.yy", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, parsedDate) Then
                        DatePickerDate.SelectedDate = parsedDate
                    Else
                        DatePickerDate.SelectedDate = Convert.ToDateTime(dateAddition)
                    End If
                Catch ex As Exception
                    DatePickerDate.SelectedDate = Nothing
                End Try
            End If
        ElseIf row.Table.Columns.Contains("Date") AndAlso row("Date") IsNot DBNull.Value Then
             ' Fallback to Date column if DateofAddition is missing
             Try
                DatePickerDate.SelectedDate = Convert.ToDateTime(row("Date"))
             Catch ex As Exception
                DatePickerDate.SelectedDate = Nothing
             End Try
        End If

        _isLoading = False
    End Sub
End Class
