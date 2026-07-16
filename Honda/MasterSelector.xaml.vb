Imports System.Collections.ObjectModel
Imports System.Data

Public Class MasterSelectorItem
    Public Property IsSelected As Boolean
    Public Property Description As String
    Public Property MasterUncertainty As Decimal
    Public Property LeastCount As String
    Public Property CalDate As Date?
    Public Property DueDate As Date?
End Class

Public Class MasterSelector
    Private MySqlCls As New MySQLClass()
    Private _existingSelectionDescriptions As New List(Of String)
    Private _referenceDate As Date? = Nothing
    Public Property SelectedMasters As New List(Of MasterSelectorItem)

    Public Sub New(Optional existingSelectionDescriptions As List(Of String) = Nothing, Optional referenceDate As Date? = Nothing)
        InitializeComponent()
        If existingSelectionDescriptions IsNot Nothing Then
            _existingSelectionDescriptions = existingSelectionDescriptions
        End If
        _referenceDate = referenceDate
        LoadMasters()
    End Sub

    Private Sub LoadMasters()
        Try
            Dim query As String = "SELECT Description, MasterUncertainty, LeastCount, CalDate, DueDate FROM calibrationmaster_details "
            If _referenceDate.HasValue Then
                Dim formattedDate = _referenceDate.Value.ToString("yyyy-MM-dd HH:mm:ss")
                query &= $"WHERE CalDate <= '{formattedDate}' AND DueDate >= '{formattedDate}' "
            End If
            query &= "ORDER BY CalDate DESC"

            Dim dt = MySqlCls.ReadDatatable(query)
            Dim items As New ObservableCollection(Of MasterSelectorItem)

            For Each row As DataRow In dt.Rows
                Dim item As New MasterSelectorItem()
                item.Description = row("Description").ToString()
                item.MasterUncertainty = If(row("MasterUncertainty") Is DBNull.Value, 0, Convert.ToDecimal(row("MasterUncertainty")))
                item.LeastCount = row("LeastCount").ToString()
                item.CalDate = If(row("CalDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("CalDate")))
                item.DueDate = If(row("DueDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("DueDate")))

                ' Only apply IsSelected if it's in the existing selection list
                If _existingSelectionDescriptions.Contains(item.Description) Then
                    item.IsSelected = True
                End If

                items.Add(item)
            Next

            MasterGrid.ItemsSource = items
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As RoutedEventArgs)
        SelectedMasters = DirectCast(MasterGrid.ItemsSource, ObservableCollection(Of MasterSelectorItem)).Where(Function(i) i.IsSelected).ToList()
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub
End Class
