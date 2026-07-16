Imports System.Windows

Public Class PasswordPromptWindow
    Public Property Verified As Boolean = False

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        TxtPassword.Focus()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnVerify_Click(sender As Object, e As RoutedEventArgs)
        Dim mysql As New MySQLClass()
        Try
            ' Get admin password from settings
            Dim dt = mysql.ReadDatatable("SELECT val FROM settings WHERE property='AppPass'")
            If dt.Rows.Count > 0 Then
                Dim adminPass = dt.Rows(0)("val").ToString()
                If TxtPassword.Password = adminPass Then
                    Verified = True
                    Me.DialogResult = True
                    Me.Close()
                Else
                    TxtError.Visibility = Visibility.Visible
                    TxtPassword.SelectAll()
                    TxtPassword.Focus()
                End If
            Else
                MessageBox.Show("Administrator password not configured in system settings.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Error verifying password: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class
