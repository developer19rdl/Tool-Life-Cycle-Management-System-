Imports System.Windows
Imports System.Windows.Controls
Imports System.Data

Public Class ResultWindow
    Private _mySql As New MySQLClass()

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        LoadCycles()
    End Sub

    Private Sub LoadCycles()
        Try
            Dim dt = _mySql.GetCycleNames()
            CmbCycles.Items.Clear()
            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    CmbCycles.Items.Add(row("cycle_name").ToString())
                Next
            End If

            Dim activeCycle = _mySql.GetActiveCycleName()
            If Not String.IsNullOrEmpty(activeCycle) AndAlso CmbCycles.Items.Contains(activeCycle) Then
                CmbCycles.SelectedItem = activeCycle
            ElseIf CmbCycles.Items.Count > 0 Then
                CmbCycles.SelectedIndex = 0
            End If
        Catch ex As Exception
            MessageBox.Show("Failed to load cycles: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub CmbCycles_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CmbCycles.SelectionChanged
        If CmbCycles.SelectedItem IsNot Nothing Then
            Dim selectedCycle = CmbCycles.SelectedItem.ToString()
            LoadResults(selectedCycle)
        End If
    End Sub

    Private Sub LoadResults(cycleName As String)
        Try
            ' Protect against rudimentary SQL injection since we use ReadDatatable directly
            Dim safeCycleName = cycleName.Replace("'", "''")
            Dim query = "SELECT control_no, category, type_name, CycleName FROM result_list WHERE CycleName = '" & safeCycleName & "'"
            Dim dt = _mySql.ReadDatatable(query)
            If dt IsNot Nothing Then
                DataGridResults.ItemsSource = dt.DefaultView
            End If
        Catch ex As Exception
            MessageBox.Show("Failed to load results: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub ControlNo_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim textBlock = DirectCast(sender, TextBlock)
            Dim rowData = DirectCast(textBlock.DataContext, DataRowView)
            If rowData IsNot Nothing Then
                Dim controlNo = rowData("control_no").ToString()
                Dim cycleName = rowData("CycleName").ToString()
                
                Dim detailWin As New CalibrationHistoryDetailWindow(controlNo, cycleName)
                detailWin.Owner = Me
                detailWin.ShowDialog()
            End If
        Catch ex As Exception
            MessageBox.Show("Failed to open detail view: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class
