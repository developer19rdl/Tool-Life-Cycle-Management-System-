Imports System.Data

Public Class CycleHistoryWindow
    Private _mySql As New MySQLClass()
    Private _controlNo As String
    Private _cycleName As String

    Public Sub New(controlNo As String, cycleName As String)
        InitializeComponent()
        _controlNo = controlNo
        _cycleName = cycleName
        TxtHeader.Text = $"Cycle History: {_cycleName} ({_controlNo})"
        LoadCycleActions()
    End Sub

    Private Sub LoadCycleActions()
        Try
            Dim escapedCtrl = _controlNo.Replace("'", "''")
            Dim escapedCycle = _cycleName.Replace("'", "''")

            ' Fetch only from specific transaction tables, deduplicated per action type
            Dim query = "SELECT MAX(Date) AS Date, Action, MAX(User) AS User FROM (" &
                        $"  SELECT CAST(CONCAT(ReintroductionDate, ' ', Time) AS DATETIME) AS Date, 'Re-introduction' AS Action, 'System' AS User FROM reintroduction WHERE ControlNo = '{escapedCtrl}' AND CycleName = '{escapedCycle}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(WriteOffDate, ' ', Time) AS DATETIME) AS Date, 'Write off' AS Action, TRIM(RaisedBy) AS User FROM writeoff WHERE ControlNo = '{escapedCtrl}' AND CycleName = '{escapedCycle}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(WOPDate, ' ', Time) AS DATETIME) AS Date, 'WOP' AS Action, TRIM(ReportedBy) AS User FROM wop WHERE ControlNo = '{escapedCtrl}' AND CycleName = '{escapedCycle}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(IssueDate, ' ', Time) AS DATETIME) AS Date, 'Temp issuance' AS Action, TRIM(IssuedBy) AS User FROM temporary_issuance WHERE ControlNo = '{escapedCtrl}' AND CycleName = '{escapedCycle}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(IssueDate, ' ', Time) AS DATETIME) AS Date, 'Issued' AS Action, TRIM(IssuedBy) AS User FROM issue WHERE ControlNo = '{escapedCtrl}' AND CycleName = '{escapedCycle}'" &
                        "  UNION ALL " &
                        $"  SELECT CAST(CONCAT(ReceiveDate, ' ', Time) AS DATETIME) AS Date, 'Received' AS Action, TRIM(ReceivedBy) AS User FROM receive WHERE ControlNo = '{escapedCtrl}' AND CycleName = '{escapedCycle}'" &
                        ") as t " &
                        "GROUP BY Action " &
                        "ORDER BY Date DESC"

            Dim dt = _mySql.ReadDatatable(query)
            ActionsDataGrid.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading cycle history: " & ex.Message)
        End Try
    End Sub

    Private Sub ActionsDataGrid_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim selectedItem = ActionsDataGrid.SelectedItem
        If selectedItem IsNot Nothing AndAlso TypeOf selectedItem Is DataRowView Then
            Dim row = DirectCast(selectedItem, DataRowView)
            Dim action = row("Action").ToString()
            
            ' Open Event Details only for specific actions that have remarks/docs
            Dim clickableActions As String() = {"WOP", "Write off", "Temp issuance", "Reintroduction", "Re-introduction", "Issued", "Received"}
            Dim isClickable = clickableActions.Any(Function(a) action.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0)

            If isClickable Then
                Dim details = _mySql.GetEventDetails(_controlNo, _cycleName, action)
                Dim win As New EventDetailsWindow("Event Details: " & action, details)
                win.ShowDialog()
            End If
        End If
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub
    Private Sub RootGrid_MouseDown(sender As Object, e As MouseButtonEventArgs)
        ActionsDataGrid.SelectedItem = Nothing
    End Sub
End Class
