Imports System.Windows
Imports System.Data
Imports System.IO

Public Class AddEditCalibrationMasterWindow
    Private _mySql As New MySQLClass()
    Private _tableName As String
    Private _masterName As String
    Private _uploadedDocPath As String = ""
    
    Public Property IsEditMode As Boolean = False
    Public Property RecordID As Integer = 0

    Public Sub New(tableName As String, masterName As String)
        InitializeComponent()
        _tableName = tableName
        _masterName = masterName
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        If IsEditMode Then
            BtnSave.Content = "Update"
            Me.Title = "Edit Calibration Master"
        Else
            BtnSave.Content = "Add"
            Me.Title = "Add Calibration Master"
            TxtDescription.Text = _masterName
            Dim today As DateTime = DateTime.Today
            DatePickerCal.SelectedDate = today
            DatePickerDue.SelectedDate = today.AddYears(1).AddDays(-1)
        End If
    End Sub

    Private Sub BtnUpload_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
            openFileDialog.Filter = "All Supported Files|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*"

            If openFileDialog.ShowDialog() = True Then
                Dim sourceFile = openFileDialog.FileName
                _uploadedDocPath = MySQLClass.CopyCalibrationMasterFile(sourceFile, _masterName)
                
                If Not String.IsNullOrEmpty(_uploadedDocPath) Then
                    TxtFileName.Text = Path.GetFileName(sourceFile)
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
        If String.IsNullOrWhiteSpace(TxtDescription.Text) Then
            MessageBox.Show("Description is required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Formatting
        Dim desc = MySQLClass.ToTitleCase(TxtDescription.Text)
        Dim lc = TxtLeastCount.Text.Trim()
        
        Dim uncertainty As Decimal = 0
        If Not Decimal.TryParse(TxtUncertainty.Text, uncertainty) Then
            ' Optional: handle invalid decimal
        End If

        Dim success As Boolean = False
        If IsEditMode Then
            success = _mySql.UpdateCalibrationMasterRecord(_tableName, RecordID, desc, lc, uncertainty, DatePickerCal.SelectedDate, DatePickerDue.SelectedDate, _uploadedDocPath)
        Else
            success = _mySql.InsertCalibrationMasterRecord(_tableName, desc, lc, uncertainty, DatePickerCal.SelectedDate, DatePickerDue.SelectedDate, _uploadedDocPath)
        End If

        If success Then
            MessageBox.Show("Record saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Public Sub PopulateForm(row As DataRow)
        If row Is Nothing Then Return
        
        IsEditMode = True
        RecordID = Convert.ToInt32(row("ID"))
        TxtDescription.Text = row("Description").ToString()
        TxtLeastCount.Text = row("LeastCount").ToString()
        TxtUncertainty.Text = row("MasterUncertainty").ToString()
        
        If row("CalDate") IsNot DBNull.Value Then DatePickerCal.SelectedDate = Convert.ToDateTime(row("CalDate"))
        If row("DueDate") IsNot DBNull.Value Then DatePickerDue.SelectedDate = Convert.ToDateTime(row("DueDate"))
        
        _uploadedDocPath = row("uploaded_doc").ToString()
        If Not String.IsNullOrEmpty(_uploadedDocPath) Then
            TxtFileName.Text = Path.GetFileName(_uploadedDocPath)
        End If
    End Sub
    Private Sub DatePickerCal_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs)
        If DatePickerCal.SelectedDate.HasValue Then
            Dim calDate As DateTime = DatePickerCal.SelectedDate.Value
            DatePickerDue.SelectedDate = calDate.AddYears(1).AddDays(-1)
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub
End Class
