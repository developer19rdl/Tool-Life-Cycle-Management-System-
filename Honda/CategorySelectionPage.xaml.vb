Imports System.Windows
Imports System.Windows.Controls
Imports System.Data

Public Class CategorySelectionPage
    Inherits Page

    Private _mySql As New MySQLClass()
    Private _category As String ' "Instrument" or "Gauge"

    Public Sub New(category As String)
        InitializeComponent()
        _category = category
        
        ' UI Updates
        TxtCategoryBreadcrumb.Text = IF(_category = "Gauge", "Gauges", "Instruments")
        TxtTitle.Text = IF(_category = "Gauge", "Select Gauge Type", "Select Instrument Type")
    End Sub

    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        LoadCategoryTypes()
    End Sub

    Private Sub LoadCategoryTypes()
        Try
            Dim dt = _mySql.GetTypeDetails()
            ' Filter the DataTable for the selected category
            Dim filteredData = dt.AsEnumerable().Where(Function(r) r.Field(Of String)("Category") = _category).ToList()
            
            If filteredData.Count > 0 Then
                Dim dtDisplay = filteredData.CopyToDataTable()
                For Each row As DataRow In dtDisplay.Rows
                    row("TypeName") = MySQLClass.ToTitleCase(row("TypeName").ToString())
                Next
                CardsContainer.ItemsSource = dtDisplay.DefaultView
            Else
                ' Show message if no types found
                MessageBox.Show($"No {_category} types configured. Please add them in Settings > Types.", "No Types Found", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading types: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnBack_Click(sender As Object, e As RoutedEventArgs)
        NavigationService.GoBack()
    End Sub

    Private Sub Card_Click(sender As Object, e As RoutedEventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim rowView = DirectCast(btn.DataContext, DataRowView)
        Dim typeName = rowView("TypeName").ToString()

        If _category = "Gauge" Then
            NavigationService.Navigate(New GaugeManagementPage(typeName))
        Else
            NavigationService.Navigate(New InstrumentManagementPage(typeName))
        End If
    End Sub
End Class
