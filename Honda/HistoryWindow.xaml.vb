Imports System.Windows
Imports System.Data

Public Class HistoryWindow
    Private _controlNo As String
    Private _itemType As String = ""
    Private _tableName As String = ""

    Public Sub New(controlNo As String, itemType As String, tableName As String)
        InitializeComponent()
        _controlNo = controlNo.Trim()
        _itemType = itemType
        _tableName = tableName
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        TxtHeader.Text = $"Lifecycle History → {_itemType}"
        TxtControlNo.Text = _controlNo

        Dim mysql As New MySQLClass()

        ' Load Header Details from the per-type table
        Dim queryDetails As String = $"SELECT * FROM `{_tableName}` WHERE ControlNo = '{_controlNo.Replace("'", "''")}'"
        Dim dtDetails As DataTable = mysql.ReadDatatable(queryDetails)

        If dtDetails IsNot Nothing AndAlso dtDetails.Rows.Count > 0 Then
            Dim row = dtDetails.Rows(0)

            ' Robust column checking to prevent crash
            If dtDetails.Columns.Contains("GaugeName") Then
                TxtType.Text = row("GaugeName").ToString()
            ElseIf dtDetails.Columns.Contains("InstrumentName") Then
                TxtType.Text = row("InstrumentName").ToString()
            Else
                TxtType.Text = "Unknown Type"
            End If

            If dtDetails.Columns.Contains("Size") Then
                TxtSize.Text = row("Size").ToString()
            ElseIf dtDetails.Columns.Contains("SizeandRange") Then
                TxtSize.Text = row("SizeandRange").ToString()
            End If

            If dtDetails.Columns.Contains("InstrumentStatus") Then
                TxtStatus.Text = row("InstrumentStatus").ToString()
            ElseIf dtDetails.Columns.Contains("Status") Then
                TxtStatus.Text = row("Status").ToString()
            End If
        End If

        ' ===== Load History Data =====
        Dim escapedCtrl = _controlNo.Replace("'", "''")
        Dim unionParts As New List(Of String)()

        ' 1. Add live interchangeability table (include original fields)
        unionParts.Add($"SELECT CAST(CONCAT(ActionDate, ' ', ActionTime) AS DATETIME) AS LogDate, CycleName, TRIM(Status) as Status, IFNULL(Remarks, '') as Remarks FROM interchangeability WHERE ControlNo = '{escapedCtrl}'")

        ' 2. Add the main unified history table
        unionParts.Add($"SELECT LogDate, CycleName, TRIM(Status) as Status, IFNULL(Remarks, '') as Remarks FROM interchange_history WHERE ControlNo = '{escapedCtrl}'")

        ' 3. Add transaction tables to ensure intermediate events are captured
        unionParts.Add($"SELECT CAST(CONCAT(WriteOffDate, ' ', Time) AS DATETIME) AS LogDate, CycleName, 'Write off' as Status, IFNULL(Reason, '') as Remarks FROM writeoff WHERE ControlNo = '{escapedCtrl}'")
        unionParts.Add($"SELECT CAST(CONCAT(WOPDate, ' ', Time) AS DATETIME) AS LogDate, CycleName, 'WOP' as Status, IFNULL(ReportedReason, '') as Remarks FROM wop WHERE ControlNo = '{escapedCtrl}'")
        unionParts.Add($"SELECT CAST(CONCAT(IssueDate, ' ', Time) AS DATETIME) AS LogDate, CycleName, 'Temp issuance' as Status, IFNULL(Reason, '') as Remarks FROM temporary_issuance WHERE ControlNo = '{escapedCtrl}'")
        unionParts.Add($"SELECT CAST(CONCAT(ReceiveDate, ' ', Time) AS DATETIME) AS LogDate, CycleName, 'Received' as Status, IFNULL(Remarks, '') as Remarks FROM receive WHERE ControlNo = '{escapedCtrl}'")
        unionParts.Add($"SELECT CAST(CONCAT(IssueDate, ' ', Time) AS DATETIME) AS LogDate, CycleName, 'Issued' as Status, IFNULL(Remarks, '') as Remarks FROM issue WHERE ControlNo = '{escapedCtrl}'")

        ' Step 3: Build the enriched query
        Dim sbUnion = String.Join(" UNION ALL ", unionParts)

        Dim fullQuery As String =
            $"SELECT MAX(u.LogDate) AS 'Date', u.CycleName AS Cycle, " &
            $"CASE " &
            $"  WHEN u.Status = 'WOP' THEN " &
            $"    IF((SELECT COUNT(*) FROM reintroduction r WHERE r.ControlNo = '{escapedCtrl}' AND r.CycleName = u.CycleName LIMIT 1) > 0, 'WOP Request(Reintroduced)', 'WOP Request') " &
            $"  WHEN u.Status = 'Write off' THEN " &
            $"    IF((SELECT COUNT(*) FROM reintroduction r WHERE r.ControlNo = '{escapedCtrl}' AND (r.CycleName = u.CycleName OR r.CycleName IS NULL OR r.CycleName = '') LIMIT 1) > 0, 'Write Off(Reintroduced)', 'Write Off') " &
            $"  WHEN u.Status = 'Temp issuance' THEN " &
            $"    IF((SELECT COUNT(*) FROM reintroduction r WHERE r.ControlNo = '{escapedCtrl}' AND r.CycleName = u.CycleName LIMIT 1) > 0, 'Temp Issuance(Reintroduced)', 'Temp Issuance') " &
            $"  ELSE " &
            $"    IF((SELECT COUNT(*) FROM reintroduction r WHERE r.ControlNo = '{escapedCtrl}' AND r.CycleName = u.CycleName LIMIT 1) > 0, CONCAT(u.Status, '(Reintroduced)'), u.Status) " &
            $"END AS Action, " &
            $"MAX(u.Remarks) AS Remarks, " &
            $"COALESCE(" &
            $"  (SELECT w.DocumentPath FROM writeoff w WHERE w.ControlNo = '{escapedCtrl}' AND (w.CycleName = u.CycleName OR w.CycleName IS NULL OR w.CycleName = '') LIMIT 1), " &
            $"  (SELECT t.DocumentPath FROM temporary_issuance t WHERE t.ControlNo = '{escapedCtrl}' AND t.CycleName = u.CycleName LIMIT 1), " &
            $"  '' " &
            $") AS DocumentPath, " &
            $"IF(u.CycleName IS NULL OR u.CycleName = '' OR u.CycleName = 'N/A', 0, " &
            $"   (CAST(SUBSTRING_INDEX(SUBSTRING_INDEX(u.CycleName, '''', -1), ' ', 1) AS UNSIGNED) * 10 + IF(LEFT(u.CycleName, 3) = 'Jul', 2, 1))) AS SortOrder " &
            $"FROM ({sbUnion}) AS u " &
            $"GROUP BY u.CycleName, u.Status " &
            $"UNION ALL " &
            $"SELECT CAST(CONCAT(ReintroductionDate, ' ', Time) AS DATETIME) AS 'Date', IFNULL(CycleName, 'N/A') AS Cycle, 'Reintroduction' AS Action, Reason AS Remarks, DocumentPath, " &
            $"IF(CycleName IS NULL OR CycleName = '' OR CycleName = 'N/A', 0, " &
            $"   (CAST(SUBSTRING_INDEX(SUBSTRING_INDEX(CycleName, '''', -1), ' ', 1) AS UNSIGNED) * 10 + IF(LEFT(CycleName, 3) = 'Jul', 2, 1))) AS SortOrder " &
            $"FROM reintroduction WHERE ControlNo = '{escapedCtrl}' " &
            $"ORDER BY SortOrder DESC, `Date` DESC, CASE WHEN Action = 'Reintroduction' THEN 0 ELSE 1 END DESC"

        Dim dtHistory As DataTable = mysql.ReadDatatable(fullQuery)

        If dtHistory IsNot Nothing AndAlso dtHistory.Rows.Count > 0 Then
            HistoryDataGrid.ItemsSource = dtHistory.DefaultView
        Else
            ' Log to console if no records found or query failed
            Console.WriteLine($"History for {escapedCtrl}: No records found in union.")
            ' Still set source to empty table instead of leaving item source null to avoid issues
            HistoryDataGrid.ItemsSource = dtHistory?.DefaultView
        End If
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub OpenPdf_Click(sender As Object, e As RoutedEventArgs)
        Dim link As Documents.Hyperlink = TryCast(sender, Documents.Hyperlink)
        If link IsNot Nothing AndAlso link.Tag IsNot Nothing Then
            Dim filePath As String = link.Tag.ToString()
            If Not String.IsNullOrEmpty(filePath) AndAlso System.IO.File.Exists(filePath) Then
                Try
                    Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                Catch ex As Exception
                    MessageBox.Show("Could not open PDF: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
            Else
                MessageBox.Show("File path is empty or file does not exist. It may not have been uploaded.", "File Not Found", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        End If
    End Sub
    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        HistoryDataGrid.SelectedItem = Nothing
    End Sub

    Private Sub Window_PreviewKeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs) Handles Me.PreviewKeyDown
        If e.Key = System.Windows.Input.Key.Escape Then
            Me.Close()
        End If
    End Sub
End Class
