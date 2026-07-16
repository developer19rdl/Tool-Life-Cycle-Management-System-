Imports System.Windows
Imports System.Windows.Controls

Public Class AddUserPage
    Inherits Page

    Private _mysql As New MySQLClass()
    Public Property IsEditMode As Boolean = False
    Public Property RecordID As Integer = 0

    Public Sub New()
        InitializeComponent()
    End Sub

    ' Constructor for Edit Mode
    Public Sub New(userId As Integer, username As String, email As String, phone As String)
        InitializeComponent()
        IsEditMode = True
        RecordID = userId
        TxtUsername.Text = username
        TxtEmail.Text = email
        TxtPhone.Text = phone
        ' Note: Password fields are left blank for security/manual re-entry if needed, or you can pre-fill
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        ' Gather password values based on visibility
        Dim passwordValue As String = If(TxtPasswordVisible.Visibility = Visibility.Visible, TxtPasswordVisible.Text, TxtPassword.Password)
        Dim confirmValue As String = If(TxtConfirmPasswordVisible.Visibility = Visibility.Visible, TxtConfirmPasswordVisible.Text, TxtConfirmPassword.Password)

        ' Validation
        If String.IsNullOrEmpty(TxtUsername.Text) OrElse String.IsNullOrEmpty(passwordValue) Then
            LblError.Text = "Username and Password are required."
            LblError.Visibility = Visibility.Visible
            Return
        End If

        If passwordValue <> confirmValue Then
            LblError.Text = "Passwords do not match."
            LblError.Visibility = Visibility.Visible
            Return
        End If

        Dim success As Boolean = False
        If IsEditMode Then
            success = _mysql.UpdateUser(RecordID, TxtUsername.Text, passwordValue, TxtEmail.Text, TxtPhone.Text)
        Else
            success = _mysql.InsertNewUser(TxtUsername.Text, passwordValue, TxtEmail.Text, TxtPhone.Text)
        End If

        If success Then
            MessageBox.Show("User saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            NavigationService.GoBack()
        Else
            MessageBox.Show("Failed to save user. Please check database connection.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
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

    Private Sub BtnShowConfirmPassword_Click(sender As Object, e As RoutedEventArgs)
        If TxtConfirmPassword.Visibility = Visibility.Visible Then
            ' Show Password
            TxtConfirmPasswordVisible.Text = TxtConfirmPassword.Password
            TxtConfirmPassword.Visibility = Visibility.Collapsed
            TxtConfirmPasswordVisible.Visibility = Visibility.Visible
            IconShowConfirmPassword.Text = "🙈"
        Else
            ' Hide Password
            TxtConfirmPassword.Password = TxtConfirmPasswordVisible.Text
            TxtConfirmPasswordVisible.Visibility = Visibility.Collapsed
            TxtConfirmPassword.Visibility = Visibility.Visible
            IconShowConfirmPassword.Text = "👁️"
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        NavigationService.GoBack()
    End Sub

End Class
