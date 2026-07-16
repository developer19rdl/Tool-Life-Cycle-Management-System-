Imports System.Windows

Public Class LoginWindow
    Private _mysql As New MySQLClass()

    Private Sub BtnLogin_Click(sender As Object, e As RoutedEventArgs)
        Dim inputUser As String = TxtUsername.Text.Trim()
        Dim inputPass As String
        
        If TxtPasswordVisible.Visibility = Visibility.Visible Then
            inputPass = TxtPasswordVisible.Text.Trim()
        Else
            inputPass = TxtPassword.Password.Trim()
        End If

        If String.IsNullOrWhiteSpace(inputUser) OrElse String.IsNullOrWhiteSpace(inputPass) Then
            ShowError("Please enter username and password.")
            Return
        End If

        ' 1. Check Admin Credentials
        If inputUser = "Admin" AndAlso inputPass = "RDL123" Then
            NavigateToMain(inputUser)
            Return
        End If

        ' 2. Check Database
        Try
            If _mysql.CheckLogin(inputUser, inputPass) Then
                NavigateToMain(inputUser)
            Else
                ShowError("Invalid username or password.")
            End If
        Catch ex As Exception
            ShowError("Database connection failed.")
        End Try
    End Sub

    Private Sub BtnShowPassword_Click(sender As Object, e As RoutedEventArgs)
        If TxtPassword.Visibility = Visibility.Visible Then
            ' Show Password
            TxtPasswordVisible.Text = TxtPassword.Password
            TxtPassword.Visibility = Visibility.Collapsed
            TxtPasswordVisible.Visibility = Visibility.Visible
            IconShowPassword.Text = "🙈"
        Else
            ' Hide Password
            TxtPassword.Password = TxtPasswordVisible.Text
            TxtPasswordVisible.Visibility = Visibility.Collapsed
            TxtPassword.Visibility = Visibility.Visible
            IconShowPassword.Text = "👁️"
        End If
    End Sub

    Private Sub NavigateToMain(username As String)
        ' Store username globally (could use a Shared property in MainWindow)
        Application.Current.Properties("Username") = username
        
        Dim mainWin As New MainWindow()
        mainWin.Show()
        Me.Close()
    End Sub

    Private Sub ShowError(msg As String)
        LblError.Text = msg
        LblError.Visibility = Visibility.Visible
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Application.Current.Shutdown()
    End Sub
End Class
