Public Class BulkScanSelectionWindow
    Inherits Window

    Public Enum ScanMethod
        None
        Usb
        Gun
    End Enum

    Public Property SelectedMethod As ScanMethod = ScanMethod.None

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub BtnUsbScan_Click(sender As Object, e As RoutedEventArgs)
        SelectedMethod = ScanMethod.Usb
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub BtnGunScan_Click(sender As Object, e As RoutedEventArgs)
        SelectedMethod = ScanMethod.Gun
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        SelectedMethod = ScanMethod.None
        Me.DialogResult = False
        Me.Close()
    End Sub
End Class
