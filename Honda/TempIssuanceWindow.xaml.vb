Imports System.Windows

Public Class TempIssuanceWindow
    Private _controlNo As String
    Private _cycleName As String
    Private _dept As String
    Private _instName As String
    Private _color As String
    Private _sizeRange As String
    Private _uploadedFilePath As String = ""

    Public Sub New(controlNo As String, cycleName As String, dept As String, instName As String, color As String, sizeRange As String)
        InitializeComponent()
        _controlNo = controlNo
        _cycleName = cycleName
        _dept = dept
        _instName = instName
        _color = color
        _sizeRange = sizeRange
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        TxtHeader.Text = $"Temp Issuance → {_controlNo}"
        BtnSubmit.IsEnabled = False
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnSubmit_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(_uploadedFilePath) Then
            MessageBox.Show("Please browse and select a PDF file first.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim mysql As New MySQLClass()
        Dim success As Boolean = mysql.InsertTempIssuance(_controlNo, _cycleName, _dept, _instName, _color, TxtReasonTemp.Text, _uploadedFilePath)

        If success Then
            ' Update interchangeability status USING UPSERT to avoid duplicates
            mysql.InsertInterchangeRecord(_cycleName, _controlNo, _dept, _instName, _sizeRange, _color, "Temp issuance", TxtReasonTemp.Text)
            
            ' Add to temp issuance calibration
            mysql.AddTempIssuanceCalibration(_instName, _controlNo, "Temp issuance", _cycleName)
            
            mysql.ClearRFIDTag(_controlNo) ' Clear RFID when issued temporarily
            
            MessageBox.Show($"Temporary Issuance details for '{_controlNo}' saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save Temporary Issuance record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As RoutedEventArgs)
        Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
        openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
        If openFileDialog.ShowDialog() = True Then
            Dim sourceFile = openFileDialog.FileName
            ' Store relative path using centralized helper
            _uploadedFilePath = MySQLClass.CopyFileToDocuments(sourceFile, "TempIssuance", _controlNo)
            
            If Not String.IsNullOrEmpty(_uploadedFilePath) Then
                Dim fileInfo As New System.IO.FileInfo(sourceFile)
                TxtFileName.Text = fileInfo.Name
                BtnSubmit.IsEnabled = True
            Else
                MessageBox.Show("Failed to copy file to Documents folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End If
    End Sub
End Class
