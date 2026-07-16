Imports System.Windows
Imports System.Data

Public Class AddEditInstrumentSizeWindow
    Private _mySql As New MySQLClass()
    Private _isEditMode As Boolean = False
    Private _currentId As Integer = 0

    Public Sub New()
        InitializeComponent()
        LoadTypes()
    End Sub

    Private Sub LoadTypes()
        Try
            Dim dt = _mySql.GetInstrumentTypesOnly()
            ' Format to Title Case
            For Each row As DataRow In dt.Rows
                row("TypeName") = MySQLClass.ToTitleCase(row("TypeName").ToString())
            Next
            ComboType.ItemsSource = dt.DefaultView
        Catch ex As Exception
            MessageBox.Show("Error loading types: " & ex.Message)
        End Try
    End Sub

    Public Sub PopulateForm(row As DataRow)
        _isEditMode = True
        _currentId = Convert.ToInt32(row("ID"))
        TxtTitle.Text = "Edit Instrument Size"
        
        ComboType.SelectedValue = row("InstrumentType").ToString()
        TxtSize.Text = row("Size").ToString()
        TxtGroupCode.Text = row("GroupCode").ToString()
        TxtSort.Text = row("Sort").ToString()
        CheckActive.IsChecked = If(IsDBNull(row("Active")), True, Convert.ToBoolean(row("Active")))
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        If ComboType.SelectedValue Is Nothing OrElse String.IsNullOrWhiteSpace(TxtSize.Text) OrElse String.IsNullOrWhiteSpace(TxtGroupCode.Text) Then
            MessageBox.Show("Please fill in all required fields (Type, Size, GroupCode).", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim type = ComboType.SelectedValue.ToString()
        Dim size = TxtSize.Text.Trim()
        Dim groupCode = TxtGroupCode.Text.Trim()
        Dim sort = TxtSort.Text.Trim()
        Dim active = If(CheckActive.IsChecked = True, 1, 0)

        Dim success As Boolean
        If _isEditMode Then
            success = _mySql.UpdateInstrumentSize(_currentId, type, size, groupCode, active, sort)
        Else
            success = _mySql.InsertInstrumentSize(type, size, groupCode, active, sort)
        End If

        If success Then
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save data. Please check database connection.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub
End Class
