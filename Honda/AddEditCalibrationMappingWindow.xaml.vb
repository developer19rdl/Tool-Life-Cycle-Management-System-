Imports System.Windows
Imports System.Windows.Controls
Imports System.Data

Public Class AddEditCalibrationMappingWindow
    Private _mySql As New MySQLClass()
    Private _isEdit As Boolean = False
    Private _mappingId As Integer = -1
    Private _matchedFormName As String = ""
    Private _isLoading As Boolean = False

    Public Sub New()
        InitializeComponent()
        LoadCategories()
    End Sub

    Private Sub LoadCategories()
        Dim categories = _mySql.GetUniqueCategories()
        ComboCategory.ItemsSource = categories
    End Sub

    Private Sub ComboCategory_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isLoading Then Return
        If ComboCategory.SelectedItem IsNot Nothing Then
            Dim selectedCategory = ComboCategory.SelectedItem.ToString()
            Dim types = _mySql.GetTypesByCategory(selectedCategory)
            ComboTypeName.ItemsSource = types
            ComboGroupCode.ItemsSource = Nothing
            ComboRange.ItemsSource = Nothing
            _matchedFormName = ""
        End If
    End Sub

    Private Sub ComboTypeName_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isLoading Then Return
        If ComboTypeName.SelectedItem IsNot Nothing AndAlso ComboCategory.SelectedItem IsNot Nothing Then
            Dim selectedType = ComboTypeName.SelectedItem.ToString()
            Dim selectedCategory = ComboCategory.SelectedItem.ToString()
            
            ' Fetch Group Codes
            Dim groupCodes = _mySql.GetGroupCodes(selectedCategory, selectedType)
            ComboGroupCode.ItemsSource = groupCodes
            If groupCodes.Count > 0 Then ComboGroupCode.SelectedIndex = 0

            ' Fetch Calibration Categories
            Dim categories = _mySql.GetCategoriesForType(selectedType)
            ComboRange.ItemsSource = categories
            _matchedFormName = ""
        End If
    End Sub

    Private Sub ComboRange_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isLoading Then Return
        If ComboRange.SelectedItem IsNot Nothing AndAlso ComboTypeName.SelectedItem IsNot Nothing Then
            Dim selectedType = ComboTypeName.SelectedItem.ToString()
            Dim selectedRange = ComboRange.SelectedItem.ToString()
            
            ' Fetch Form Name (Background)
            _matchedFormName = _mySql.GetFormByCalibrationCategory(selectedType, selectedRange)
        End If
    End Sub

    Public Sub PopulateForm(row As DataRow)
        If row Is Nothing Then Return
        _isLoading = True
        Try
            _isEdit = True
            _mappingId = Convert.ToInt32(row("id"))
            Me.Title = "Edit Calibration Mapping"
            BtnSave.Content = "Update"

            ' Set Category
            Dim cat = row("category").ToString()
            ComboCategory.SelectedItem = cat

            ' Set Types
            Dim types = _mySql.GetTypesByCategory(cat)
            ComboTypeName.ItemsSource = types
            Dim typeName = row("type_name").ToString()
            ComboTypeName.SelectedItem = typeName

            ' Set Group Codes
            Dim groupCodes = _mySql.GetGroupCodes(cat, typeName)
            ComboGroupCode.ItemsSource = groupCodes
            ComboGroupCode.Text = row("prefix").ToString() ' Use Text for ComboBox (could be editable or not)

            ' Set Calibration Categories
            Dim categories = _mySql.GetCategoriesForType(typeName)
            ComboRange.ItemsSource = categories
            Dim calibCategory = row("calibration_category").ToString()
            ComboRange.SelectedItem = calibCategory

            ' Form Name (Background)
            _matchedFormName = _mySql.GetFormByCalibrationCategory(typeName, calibCategory)

        Finally
            _isLoading = False
        End Try
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        If ComboCategory.SelectedItem Is Nothing OrElse ComboTypeName.SelectedItem Is Nothing OrElse 
           String.IsNullOrEmpty(ComboGroupCode.Text) OrElse ComboRange.SelectedItem Is Nothing Then
            MessageBox.Show("Please fill all required fields.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim category = ComboCategory.SelectedItem.ToString()
        Dim typeName = ComboTypeName.SelectedItem.ToString()
        Dim prefix = ComboGroupCode.Text
        Dim sizeRange = ComboRange.SelectedItem.ToString()
        Dim formName = _matchedFormName

        Dim success As Boolean = False
        If _isEdit Then
            success = _mySql.UpdateCalibrationMapping(_mappingId, category, typeName, prefix, sizeRange, formName)
        Else
            success = _mySql.InsertCalibrationMapping(category, typeName, prefix, sizeRange, formName)
        End If

        If success Then
            DialogResult = True
            Close()
        Else
            MessageBox.Show("Failed to save calibration mapping.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        DialogResult = False
        Close()
    End Sub
End Class
