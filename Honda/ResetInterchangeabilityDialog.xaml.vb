Imports System.Windows

Public Class ResetInterchangeabilityDialog
    Public Property Result As ResetType = ResetType.Cancel
    Public Enum ResetType
        Cancel
        CurrentCycle
        All
    End Enum

    Public Sub New(cycleName As String)
        InitializeComponent()
        RadioCurrentCycle.Content = cycleName
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.Result = ResetType.Cancel
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnSubmit_Click(sender As Object, e As RoutedEventArgs)
        If RadioCurrentCycle.IsChecked = True Then
            Me.Result = ResetType.CurrentCycle
        ElseIf RadioAll.IsChecked = True Then
            Me.Result = ResetType.All
        End If
        Me.DialogResult = True
        Me.Close()
    End Sub
End Class
