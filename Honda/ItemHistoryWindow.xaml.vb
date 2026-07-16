Imports System.Windows
Imports System.Data

Public Class ItemHistoryWindow
    Private _controlNo As String
    Private _tableName As String
    Private _itemType As String

    Public Sub New(controlNo As String, itemType As String, tableName As String)
        InitializeComponent()
        _controlNo = controlNo
        _itemType = itemType
        _tableName = tableName
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        TxtHeader.Text = $"{_itemType} History → Lifecycle"
        TxtControlNo.Text = _controlNo
        LoadItemDetails()
        LoadHistoryData()
    End Sub

    Private Sub LoadItemDetails()
        Dim mysql As New MySQLClass()
        Try
            Dim dt = mysql.ReadDatatable($"SELECT * FROM `{_tableName}` WHERE ControlNo = '{_controlNo.Replace("'", "''")}'")
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                
                ' Set Item Name
                If dt.Columns.Contains("InstrumentName") Then
                    TxtItemName.Text = row("InstrumentName").ToString()
                ElseIf dt.Columns.Contains("GaugeName") Then
                    TxtItemName.Text = row("GaugeName").ToString()
                End If

                ' Set Status
                If dt.Columns.Contains("Status") Then
                    TxtStatus.Text = row("Status").ToString()
                ElseIf dt.Columns.Contains("InstrumentStatus") Then
                    TxtStatus.Text = row("InstrumentStatus").ToString()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error loading item details: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadHistoryData()
        Dim mysql As New MySQLClass()
        Try
            Dim escapedCtrl = _controlNo.Replace("'", "''")
            
            ' Query for Write-off and Reintroduction (including DocumentPath)
            Dim query = "SELECT Date, EventType, Remarks, DocumentPath FROM (" &
                        $"  SELECT CAST(CONCAT(WriteOffDate, ' ', Time) AS DATETIME) AS Date, 'Write-off' AS EventType, Reason AS Remarks, DocumentPath FROM writeoff WHERE ControlNo = '{escapedCtrl}' " &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(ReintroductionDate, ' ', Time) AS DATETIME) AS Date, 'Reintroduction' AS EventType, Reason AS Remarks, DocumentPath FROM reintroduction WHERE ControlNo = '{escapedCtrl}' " &
                        ") as t ORDER BY Date DESC"

            Dim dt = mysql.ReadDatatable(query)
            HistoryDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading history: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub OpenPdf_Click(sender As Object, e As RoutedEventArgs)
        Dim link As Documents.Hyperlink = TryCast(sender, Documents.Hyperlink)
        If link IsNot Nothing AndAlso link.Tag IsNot Nothing Then
            Dim filePath As String = link.Tag.ToString()
            If Not String.IsNullOrEmpty(filePath) AndAlso System.IO.File.Exists(filePath) Then
                Try
                    Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                Catch ex As Exception
                    MessageBox.Show("Could not open file: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
            Else
                MessageBox.Show("File path is empty or file does not exist.", "File Not Found", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        End If
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub
End Class
