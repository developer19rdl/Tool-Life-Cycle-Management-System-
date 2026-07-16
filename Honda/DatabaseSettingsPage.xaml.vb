Imports System.Windows
Imports System.Windows.Media

Public Class DatabaseSettingsPage
    Public Sub New()
        InitializeComponent()
        LoadSettings()
    End Sub

    Private Sub LoadSettings()
        Dim s = ProjectSettings.Current
        txtServer.Text = s.Datasource
        txtPort.Text = s.Port
        txtDatabase.Text = s.Database
        txtUsername.Text = s.Username
        pbPassword.Password = s.Password
    End Sub

    Private Sub btnTestConnection_Click(sender As Object, e As RoutedEventArgs)
        Dim db As New MySQLClass()
        StatusLabel.Text = "Connection Status: Testing..."
        StatusLabel.Foreground = New SolidColorBrush(Color.FromRgb(148, 163, 184)) ' Gray

        If db.TestConnection(txtServer.Text, txtPort.Text, txtDatabase.Text, txtUsername.Text, pbPassword.Password) Then
            StatusLabel.Text = "Connection Status: Success!"
            StatusLabel.Foreground = New SolidColorBrush(Color.FromRgb(34, 197, 94)) ' Green
        Else
            StatusLabel.Text = "Connection Status: Failed!"
            StatusLabel.Foreground = New SolidColorBrush(Color.FromRgb(239, 68, 68)) ' Red
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs)
        Dim s = ProjectSettings.Current
        s.Datasource = txtServer.Text
        s.Port = txtPort.Text
        s.Database = txtDatabase.Text
        s.Username = txtUsername.Text
        s.Password = pbPassword.Password

        ProjectSettings.Save()
        MessageBox.Show("Settings saved successfully!", "Configuration", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub
End Class
