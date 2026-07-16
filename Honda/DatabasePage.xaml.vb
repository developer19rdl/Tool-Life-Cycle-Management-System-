Imports System.Windows
Imports System.Windows.Controls

Public Class DatabasePage
    Inherits Page

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InstrumentCard_Click(sender As Object, e As RoutedEventArgs)
        NavigationService.Navigate(New CategorySelectionPage("Instrument"))
    End Sub

    Private Sub GaugeCard_Click(sender As Object, e As RoutedEventArgs)
        NavigationService.Navigate(New CategorySelectionPage("Gauge"))
    End Sub

End Class
