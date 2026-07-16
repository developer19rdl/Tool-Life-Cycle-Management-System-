Public Class ReintroductionWindow
    Private _controlNo As String
    Private _tableName As String = ""
    Private _itemType As String = ""
    Private _uploadedFilePath As String = ""
    
    Private _name As String = ""
    Private _line As String = ""
    Private _dept As String = ""
    Private _color As String = ""
    Private _quantity As String = ""
    Private _writeOffNo As String = ""
    Private _missingDate As DateTime = DateTime.Today
    Private _cycleName As String = ""
    Private _sizeRange As String = ""

    Public Sub New(controlNo As String, tableName As String, Optional cycleName As String = "", Optional sizeRange As String = "")
        InitializeComponent()
        _controlNo = controlNo
        _tableName = tableName
        _cycleName = cycleName
        _sizeRange = sizeRange
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim mysql As New MySQLClass()

        ' Load from the known per-type table directly
        Dim dt As System.Data.DataTable = mysql.ReadDatatable($"SELECT * FROM `{_tableName}` WHERE ControlNo = '{_controlNo}'")

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim row = dt.Rows(0)
            If row.Table.Columns.Contains("InstrumentName") Then
                _itemType = "Instrument"
                _name = row("InstrumentName").ToString()
            Else
                _itemType = "Gauge"
                _name = row("GaugeName").ToString()
            End If
            _line = row("Line").ToString()
            _dept = row("Section").ToString()
            _color = row("Color").ToString()
        End If

        ' Fetch WriteOff data to prefill
        Dim queryWO As String = "SELECT WriteOffNo, WriteOffDate FROM writeoff WHERE ControlNo = '" & _controlNo & "' ORDER BY WriteOffDate DESC LIMIT 1"
        Dim dtWO As System.Data.DataTable = mysql.ReadDatatable(queryWO)

        If dtWO IsNot Nothing AndAlso dtWO.Rows.Count > 0 Then
            _writeOffNo = dtWO.Rows(0)("WriteOffNo").ToString()
            If Not IsDBNull(dtWO.Rows(0)("WriteOffDate")) Then
                _missingDate = Convert.ToDateTime(dtWO.Rows(0)("WriteOffDate"))
            End If
        End If

        ' Setup Department selector for Master List mode
        If String.IsNullOrEmpty(_cycleName) Then
            lblLocation.Visibility = Visibility.Visible
            ComboLocation.Visibility = Visibility.Visible
            Dim dtDepts = mysql.GetDepartments()
            ComboLocation.ItemsSource = dtDepts.DefaultView
            ' Pre-select current department if available
            If Not String.IsNullOrEmpty(_dept) Then
                ComboLocation.SelectedValue = _dept
            End If
        End If

        TxtHeader.Text = $"Reintroduction → {_itemType} → {_controlNo}"
        BtnReintroduce.IsEnabled = False
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnReintroduce_Click(sender As Object, e As RoutedEventArgs) Handles BtnReintroduce.Click
        If String.IsNullOrWhiteSpace(_uploadedFilePath) Then
            MessageBox.Show("Please browse and select a PDF file first.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim mysql As New MySQLClass()
        
        ' If in Master List mode, use the selected department
        Dim finalDept = _dept
        If String.IsNullOrEmpty(_cycleName) Then
            If ComboLocation.SelectedValue Is Nothing Then
                MessageBox.Show("Please select a department for reintroduction.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
            finalDept = ComboLocation.SelectedValue.ToString()
        End If

            Dim targetCycle = _cycleName
            If String.IsNullOrEmpty(targetCycle) Then
                targetCycle = mysql.GetActiveCycleName()
            End If

        Dim emptyVal As String = ""
        
        Dim success As Boolean = mysql.InsertReintroduction(
            DateTime.Today,
            _itemType,
            _name,
            _controlNo,
            _quantity,
            _line,
            finalDept,
            _color,
            TxtReasonReIntro.Text,
            _writeOffNo,
            $"RE-{DateTime.Now.ToString("yyyyMMddHHmmss")}",
            DateTime.Today,
            _missingDate,
            emptyVal,
            emptyVal,
            emptyVal,
            emptyVal,
            emptyVal,
            Application.Current.Properties("Username")?.ToString(),
            emptyVal,
            _uploadedFilePath,
            targetCycle
        )

        If success Then
            mysql.UpdateRecordStatus(_tableName, _controlNo, 0, "Active")
            
            If Not String.IsNullOrEmpty(targetCycle) Then
                ' Ensure it's in the department_list first (automated sync rule)
                Dim batchTag = "Reintro " & DateTime.Now.ToString("yyyyMMdd")
                mysql.InsertDeptListItem(finalDept, _name, _sizeRange, _controlNo, _color, "Pending", TxtReasonReIntro.Text, batchTag, targetCycle)

                ' Update interchangeability status USING UPSERT
                mysql.InsertInterchangeRecord(targetCycle, _controlNo, finalDept, _name, _sizeRange, _color, "Pending", TxtReasonReIntro.Text)

                ' NOTE: reintroduction_calibration entry is created later when the item is Received
                '       (InsertInterchangeRecord detects the reintroduction row and routes accordingly)
            End If
            
            MessageBox.Show($"Reintroduction details for '{_controlNo}' saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save Reintroduction record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As RoutedEventArgs) Handles BtnBrowse.Click
        Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
        openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
        If openFileDialog.ShowDialog() = True Then
            Dim sourceFile = openFileDialog.FileName
            ' Store relative path using centralized helper
            _uploadedFilePath = MySQLClass.CopyFileToDocuments(sourceFile, "Reintroduction", _controlNo)
            
            If Not String.IsNullOrEmpty(_uploadedFilePath) Then
                Dim fileInfo As New System.IO.FileInfo(sourceFile)
                TxtFileName.Text = fileInfo.Name
                BtnReintroduce.IsEnabled = True
            Else
                MessageBox.Show("Failed to copy file to Documents folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub
End Class
