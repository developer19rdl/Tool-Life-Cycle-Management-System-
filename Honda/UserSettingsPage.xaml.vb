Imports System.Windows
Imports System.Windows.Controls
Imports System.Data ' Added for DataTable and DataRowView

Public Class UserSettingsPage
    Inherits Page

    Private _mysql As New MySQLClass()

    Public Sub New()
        InitializeComponent()
        LoadUsers()
    End Sub

    Private Sub LoadUsers()
        ' Using the same query as the source project
        Dim query As String = "SELECT ID, Username AS UserName, Password, Email, Phone FROM users"
        Dim dt As DataTable = _mysql.ReadDatatable(query)
        UsersDataGrid.ItemsSource = dt.DefaultView
    End Sub

    Private Sub BtnAddUser_Click(sender As Object, e As RoutedEventArgs)
        ' Navigate to AddUserPage
        NavigationService.Navigate(New AddUserPage())
    End Sub

    Private Sub BtnEditUser_Click(sender As Object, e As RoutedEventArgs)
        ' Navigate to AddUserPage with selected user
        Dim selectedRow As DataRowView = TryCast(UsersDataGrid.SelectedItem, DataRowView)
        If selectedRow IsNot Nothing Then
            Dim userId As Integer = Convert.ToInt32(selectedRow("ID"))
            Dim username As String = selectedRow("UserName").ToString()
            Dim email As String = selectedRow("Email").ToString()
            Dim phone As String = selectedRow("Phone").ToString()
            
            NavigationService.Navigate(New AddUserPage(userId, username, email, phone))
        Else
            MessageBox.Show("Please select a user from the list to edit.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
        End If
    End Sub

    Private Sub BtnDeleteUser_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRowView = TryCast(UsersDataGrid.SelectedItem, DataRowView)
        If selectedRow IsNot Nothing Then
            Dim userId As Integer = Convert.ToInt32(selectedRow("ID"))
            Dim username As String = selectedRow("UserName").ToString()

            Dim result = MessageBox.Show($"Are you sure you want to permanently delete user: {username}?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning)
            If result = MessageBoxResult.Yes Then
                If _mysql.DeleteUser(userId) Then
                    MessageBox.Show("User deleted successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                    LoadUsers()
                Else
                    MessageBox.Show("Failed to delete user. Please check database connection.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            End If
        Else
            MessageBox.Show("Please select a user to delete.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning)
        End If
    End Sub

    Private Sub BtnRefresh_Click(sender As Object, e As RoutedEventArgs)
        LoadUsers()
    End Sub

    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        UsersDataGrid.SelectedItem = Nothing
    End Sub

End Class
