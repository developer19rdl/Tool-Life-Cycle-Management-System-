Imports System.Windows
Imports System.Windows.Controls
Imports System.IO.Ports

Public Class ReceiveRFIDWindow
    Inherits Window
    Private _controlNo As String
    Private _mySql As New MySQLClass()

    Public Property RfidTag As String = ""
    Public Property Remarks As String = ""
    Public Property IsConfirmed As Boolean = False

    Public Sub New(controlNo As String)
        InitializeComponent()
        _controlNo = controlNo
        TxtControlNo.Text = $"Control No: {_controlNo}"

        ' Hook up events in code behind to avoid XAML parser issues
        AddHandler Me.Loaded, Sub() TxtRemarks.Focus()
        AddHandler Me.PreviewKeyDown, AddressOf Window_PreviewKeyDown

        ' Set this HERE to ensure all components are initialized before event fires
        ChkRFID.IsChecked = True
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnReceive_Click(sender As Object, e As RoutedEventArgs)
        If ChkRFID.IsChecked = True AndAlso String.IsNullOrWhiteSpace(TxtRFID.Text) Then
            MessageBox.Show("Please scan the RFID tag or uncheck the RFID option.", "RFID Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Me.RfidTag = If(ChkRFID.IsChecked = True, TxtRFID.Text.Trim(), "")
        Me.Remarks = TxtRemarks.Text.Trim()
        Me.IsConfirmed = True
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub ChkRFID_Checked(sender As Object, e As RoutedEventArgs)
        ' Trigger scan automatically when checked
        ScanRFID()
        TxtRemarks.Focus()
    End Sub

    Private Sub ChkRFID_Unchecked(sender As Object, e As RoutedEventArgs)
        TxtRFID.Clear()
        BtnReceive.IsEnabled = True ' Always enable for manual comments
        TxtRemarks.Focus()
    End Sub

    Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
            If BtnReceive.IsEnabled Then
                BtnReceive_Click(Nothing, Nothing)
            End If
            e.Handled = True
        End If
    End Sub

    Private Async Sub ScanRFID()
        Try
            If TxtRFID Is Nothing Then Return ' Safety check
            TxtRFID.Text = "Scanning..."

            Dim scannedTag = Await Task.Run(Function() CommManager.ScanSingleTag(Nothing, 5000))

            ' If user unchecked it during the 5 second wait, just abort
            If ChkRFID.IsChecked = False Then
                TxtRFID.Clear()
                Return
            End If

            If String.IsNullOrWhiteSpace(scannedTag) Then
                TxtRFID.Clear()
                MessageBox.Show("No RFID tag detected. Please place the tag near the reader and try again.", "Scan Failed", MessageBoxButton.OK, MessageBoxImage.Warning)
                ChkRFID.IsChecked = False
                Return
            End If

            TxtRFID.Text = scannedTag

            ' Check if tag is excluded
            If _mySql.IsTagExcluded(scannedTag) Then
                MessageBox.Show("This tag is marked as excluded. Please use a different tag.", "Excluded Tag Scanned", MessageBoxButton.OK, MessageBoxImage.Warning)
                TxtRFID.Clear()
                Return
            End If

            ' Check if RFID already exists for another tool
            Dim existingCtrl = _mySql.CheckRFIDExists(scannedTag, _controlNo)
            If Not String.IsNullOrEmpty(existingCtrl) Then
                TxtError.Text = $"Duplicate RFID: Assigned to {existingCtrl}"
                TxtError.Foreground = System.Windows.Media.Brushes.Red
                TxtError.Visibility = Visibility.Visible
                BtnReceive.IsEnabled = False ' Disable on duplicate
                Return
            End If
            TxtError.Visibility = Visibility.Collapsed
            BtnReceive.IsEnabled = True ' Enable on valid scan

        Catch ex As Exception
            MessageBox.Show("RFID Scan Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class

