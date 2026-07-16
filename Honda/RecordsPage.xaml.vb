Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports System.Data

Class RecordsPage
    Private db As New MySQLClass()
    Private WithEvents timer As New DispatcherTimer()

    Public Sub New()
        InitializeComponent()
        
        ' Initialize timer for real-time updates (every 2 seconds)
        timer.Interval = TimeSpan.FromSeconds(2)
        
        ' Load initial data
        LoadRecordsData()
        LoadRegisteredData()
        LoadUnRegisteredData()
        
        ' Start timer
        timer.Start()
    End Sub

    Private Sub TabButton_Click(sender As Object, e As RoutedEventArgs)
        Dim clickedButton = TryCast(sender, Button)
        If clickedButton Is Nothing Then Return

        ' Reset all tab buttons
        BtnRecords.Tag = ""
        BtnRegistered.Tag = ""
        BtnUnRegistered.Tag = ""

        ' Set selected tab button
        clickedButton.Tag = "Selected"

        ' Toggle visibility of tab views
        RecordsTabView.Visibility = Visibility.Collapsed
        RegisteredTabView.Visibility = Visibility.Collapsed
        UnRegisteredTabView.Visibility = Visibility.Collapsed

        Select Case clickedButton.Name
            Case "BtnRecords"
                RecordsTabView.Visibility = Visibility.Visible
                LoadRecordsData() ' Refresh on click
            Case "BtnRegistered"
                RegisteredTabView.Visibility = Visibility.Visible
                LoadRegisteredData() ' Refresh on click
            Case "BtnUnRegistered"
                UnRegisteredTabView.Visibility = Visibility.Visible
                LoadUnRegisteredData() ' Refresh on click
        End Select
    End Sub

    Private Sub LoadRecordsData()
        Dim dt = db.GetScanLogs()
        RecordsDataGrid.ItemsSource = dt.DefaultView
    End Sub

    Private Sub LoadRegisteredData()
        Dim dt = db.GetRegisteredItems()
        RegisteredDataGrid.ItemsSource = dt.DefaultView
    End Sub

    Private Sub LoadUnRegisteredData()
        Dim dt = db.GetUnRegisteredItems()
        UnRegisteredDataGrid.ItemsSource = dt.DefaultView
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles timer.Tick
        ' Only refresh Records data if the Records tab is visible
        If RecordsTabView.Visibility = Visibility.Visible Then
            LoadRecordsData()
        End If
    End Sub
End Class
