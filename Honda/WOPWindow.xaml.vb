Imports System.Windows

Public Class WOPWindow
    Private _controlNo As String
    Private _cycleName As String
    Private _dept As String
    Private _instName As String
    Private _color As String
    Private _sizeRange As String

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
        TxtHeader.Text = $"WOP → {_controlNo}"
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnSubmit_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(TxtReasonWOP.Text) Then
            MessageBox.Show("Please enter a reported reason.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim mysql As New MySQLClass()
        Dim success As Boolean = mysql.InsertWOPRecord(_controlNo, _cycleName, _dept, _instName, _color, TxtReasonWOP.Text)

        If success Then
            ' Update interchangeability status USING UPSERT to avoid duplicates
            mysql.InsertInterchangeRecord(_cycleName, _controlNo, _dept, _instName, _sizeRange, _color, "WOP", TxtReasonWOP.Text)
            mysql.ClearRFIDTag(_controlNo) ' Clear RFID when marked as WOP
            
            MessageBox.Show($"WOP recorded for '{_controlNo}'.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save WOP record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub
End Class
