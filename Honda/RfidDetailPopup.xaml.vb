Public Class RfidDetailPopup

    Public Sub New(results As List(Of RFIDLookupResult))
        ' This call is required by the designer.
        InitializeComponent()

        ' Set values
        DgResults.ItemsSource = results
        TxtCountHeader.Text = $"{results.Count} Tool(s) Found"
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        DgResults.SelectedItem = Nothing
    End Sub

End Class
