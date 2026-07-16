Imports System.Windows

Public Class ExclusionConfirmWindow
    Inherits Window

    Public Sub New(items As IEnumerable(Of InventoryManagementPage.ExclusionItem))
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ConfirmDataGrid.ItemsSource = items
    End Sub

    Private Sub BtnYes_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub BtnNo_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub
End Class
