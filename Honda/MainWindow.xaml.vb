Imports System.Windows


Class MainWindow

    ''' <summary>Embedded REST API for RFID lookups (runs on port 8080).</summary>
    Private _apiServer As New RFIDApiServer()

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Display Logged-in Username
        Dim currentUser As String = "System"
        If Application.Current.Properties.Contains("Username") Then
            currentUser = Application.Current.Properties("Username").ToString()
            LblCurrentUsername.Text = currentUser
        End If

        ' Run Automated Cycle Checks
        Dim mysql As New MySQLClass()
        CycleManager.CheckAndRunAutomatedRollover(mysql, currentUser)
        CycleManager.CheckForUpcomingRolloverWarning(mysql)

        ' Load the Welcome Page inside the Frame when the app starts
        MainFrame.Navigate(New HomePage())

        ' Start the embedded RFID REST API server
        _apiServer.Start()
        BackendManager.StartRfidBackend()
    End Sub

    ''' <summary>Stop the API server cleanly when the window is closing.</summary>
    Private Sub MainWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        _apiServer.Stop()
    End Sub

    Private Sub BtnLogout_Click(sender As Object, e As RoutedEventArgs)
        Dim result = MessageBox.Show("Are you sure you want to log out?", "Logout", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If result = MessageBoxResult.Yes Then
            ' Clear session
            Application.Current.Properties.Remove("Username")
            
            ' Re-open LoginWindow
            Dim loginWin As New LoginWindow()
            loginWin.Show()
            Me.Close()
        End If
    End Sub

    ' Navigate to the Database page when the menu item is clicked
    Private Sub DatabaseMenu_Click(sender As Object, e As RoutedEventArgs)
        MainFrame.Navigate(New DatabasePage())
    End Sub

    ' Navigate to the DB Settings page when the menu item is clicked
    Private Sub DatabaseSettings_Click(sender As Object, e As RoutedEventArgs)
        MainFrame.Navigate(New DatabaseSettingsPage())
    End Sub
    ' Add this right below your DatabaseSettings_Click sub
    Private Sub UserSettings_Click(sender As Object, e As RoutedEventArgs)
        ' This will navigate to the User Settings page once we create it
        MainFrame.Navigate(New UserSettingsPage())
    End Sub

    Private Sub RFIDScanner_Click(sender As Object, e As RoutedEventArgs)
        Dim scannerWin As New RFIDScannerSelectionWindow()
        scannerWin.Owner = Me
        scannerWin.ShowDialog()
    End Sub

    Private Sub SettingsMenu_Click(sender As Object, e As RoutedEventArgs)
        MainFrame.Navigate(New SettingsPage())
    End Sub

    Private Sub CalibrationMenu_Click(sender As Object, e As RoutedEventArgs)
        MainFrame.Navigate(New CalibrationPage())
    End Sub

    Private Sub InterchangeableMenu_Click(sender As Object, e As RoutedEventArgs)
        MainFrame.Navigate(New InterchangeablePage())
    End Sub

    Private Sub RecordsMenu_Click(sender As Object, e As RoutedEventArgs)
        MainFrame.Navigate(New RecordsPage())
    End Sub

    Private Sub InventoryManagementMenu_Click(sender As Object, e As RoutedEventArgs)
        MainFrame.Navigate(New InventoryManagementPage())
    End Sub

    ' You can leave this empty or delete it if you removed "Click=MenuItem_Click" from the XAML
    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)

    End Sub

End Class