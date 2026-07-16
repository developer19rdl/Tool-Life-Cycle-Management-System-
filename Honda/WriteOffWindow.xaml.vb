Imports System.Windows

Public Class WriteOffWindow
    Private _controlNo As String
    Private _tableName As String = ""
    Private _itemType As String = ""
    Private _uploadedFilePath As String = ""
    
    ' Fields to store original data for the insert function
    Private _name As String = ""
    Private _line As String = ""
    Private _dept As String = ""
    Private _color As String = ""
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

        ' Load original data directly from the known per-type table
        Dim dt As System.Data.DataTable = mysql.ReadDatatable($"SELECT * FROM `{_tableName}` WHERE ControlNo = '{_controlNo}'")

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim row = dt.Rows(0)
            ' Detect type by column name
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

        ' Set Header
        TxtHeader.Text = $"Write Off → {_itemType} → {_controlNo}"
        
        ' Disable Write Off button initially
        BtnWriteOff.IsEnabled = False
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As RoutedEventArgs) Handles BtnWriteOff.Click
        If String.IsNullOrWhiteSpace(_uploadedFilePath) Then
            MessageBox.Show("Please browse and select a PDF file first.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim mysql As New MySQLClass()
        Dim emptyVal As String = ""
        
        Dim success As Boolean = mysql.InsertWriteOff(
            DateTime.Today,
            DateTime.Today,
            _itemType,
            _name,
            _controlNo,
            emptyVal,
            _line,
            _dept,
            _color,
            TxtReasonWriteOff.Text,
            emptyVal,
            emptyVal,
            _uploadedFilePath,
            $"WO-{DateTime.Now.ToString("yyyyMMddHHmmss")}",
            "WrittenOff",
            Application.Current.Properties("Username")?.ToString(),
            emptyVal,
            _cycleName
        )

        If success Then
            mysql.ExecuteNonQuery($"DELETE FROM department_list WHERE `Control No` = '{_controlNo.Replace("'", "''")}' AND CycleName = '{_cycleName.Replace("'", "''")}'")
            mysql.UpdateRecordStatus(_tableName, _controlNo, 1, "In-Active")
            
            If Not String.IsNullOrEmpty(_cycleName) Then
                ' Update interchangeability status USING UPSERT to avoid duplicates
                mysql.InsertInterchangeRecord(_cycleName, _controlNo, _dept, _name, _sizeRange, _color, "Write off", TxtReasonWriteOff.Text)
                
                ' Populate NG List for Write Off
                Dim ngReason = If(String.IsNullOrWhiteSpace(TxtReasonWriteOff.Text), "Writeoff NG", TxtReasonWriteOff.Text)
                mysql.InsertNGRecord(_controlNo, _name, _dept, _cycleName, ngReason, DateTime.Today, DateTime.Today, "Write off")
            End If
            
            mysql.ClearRFIDTag(_controlNo) ' Clear RFID when written off

            MessageBox.Show($"WriteOff details for '{_controlNo}' saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save WriteOff record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub BtnUpload_Click(sender As Object, e As RoutedEventArgs) Handles BtnBrowse.Click
        Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
        openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
        If openFileDialog.ShowDialog() = True Then
            Dim sourceFile = openFileDialog.FileName
            ' Store relative path using centralized helper
            _uploadedFilePath = MySQLClass.CopyFileToDocuments(sourceFile, "WriteOff", _controlNo)
            
            If Not String.IsNullOrEmpty(_uploadedFilePath) Then
                Dim fileInfo As New System.IO.FileInfo(sourceFile)
                TxtFileName.Text = fileInfo.Name
                BtnWriteOff.IsEnabled = True
                BtnWriteOff.Style = DirectCast(FindResource("BlueButtonStyle"), Style)
            Else
                MessageBox.Show("Failed to copy file to Documents folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub
End Class
