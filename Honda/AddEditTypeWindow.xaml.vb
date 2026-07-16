Imports System.Windows
Imports System.Data
Imports System.IO
Imports Microsoft.Win32
Imports System.Windows.Media.Imaging

Public Class AddEditTypeWindow
    Private _mySql As New MySQLClass()
    
    Public Property IsEditMode As Boolean = False
    Public Property RecordID As Integer = 0
    Private _oldTypeName As String = ""
    Private _typeImageBase64 As String = ""

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        If IsEditMode Then
            Me.Title = "Edit Type Details"
            BtnSave.Content = "Update"
        Else
            Me.Title = "Add Type Details"
            BtnSave.Content = "Save"
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub BtnSelectImage_Click(sender As Object, e As RoutedEventArgs)
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image files (*.png;*.jpeg;*.jpg;*.bmp)|*.png;*.jpeg;*.jpg;*.bmp|All files (*.*)|*.*"
        
        If openFileDialog.ShowDialog() = True Then
            Try
                Dim fileBytes = File.ReadAllBytes(openFileDialog.FileName)
                _typeImageBase64 = Convert.ToBase64String(fileBytes)
                SetImageFromBase64(_typeImageBase64)
            Catch ex As Exception
                MessageBox.Show("Error loading image: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub BtnClearImage_Click(sender As Object, e As RoutedEventArgs)
        _typeImageBase64 = ""
        ImgPreview.Source = Nothing
        TxtNoImage.Visibility = Visibility.Visible
    End Sub

    Private Sub SetImageFromBase64(base64String As String)
        If String.IsNullOrEmpty(base64String) Then
            ImgPreview.Source = Nothing
            TxtNoImage.Visibility = Visibility.Visible
            Return
        End If

        Try
            Dim binaryData() As Byte = Convert.FromBase64String(base64String)
            Dim bi As New BitmapImage()
            bi.BeginInit()
            bi.StreamSource = New MemoryStream(binaryData)
            bi.EndInit()
            ImgPreview.Source = bi
            TxtNoImage.Visibility = Visibility.Collapsed
        Catch ex As Exception
            ImgPreview.Source = Nothing
            TxtNoImage.Visibility = Visibility.Visible
        End Try
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        ' Simple validation
        If ComboCategory.SelectedItem Is Nothing OrElse String.IsNullOrWhiteSpace(TxtTypeName.Text) Then
            MessageBox.Show("Please fill in Category and Type Name.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim category = DirectCast(ComboCategory.SelectedItem, ComboBoxItem).Content.ToString()
        Dim typeName = TxtTypeName.Text
        Dim prefixMode = If(ComboPrefixMode.SelectedItem IsNot Nothing, DirectCast(ComboPrefixMode.SelectedItem, ComboBoxItem).Content.ToString(), "Manual")
        Dim basePrefix = TxtBasePrefix.Text
        Dim serialDigits = TxtSerialDigits.Text

        Dim success As Boolean = False
        If IsEditMode Then
            ' If name changed, rename the table
            If Not typeName.Equals(_oldTypeName, StringComparison.OrdinalIgnoreCase) Then
                Dim oldTbl = MySQLClass.TypeNameToTableName(_oldTypeName, forInventory:=True)
                Dim newTbl = MySQLClass.TypeNameToTableName(typeName, forInventory:=True)
                _mySql.RenameTypeTable(oldTbl, newTbl)
            End If
            success = _mySql.UpdateTypeDetail(RecordID, category, typeName, prefixMode, basePrefix, serialDigits, _typeImageBase64)
        Else
            success = _mySql.InsertTypeDetail(category, typeName, prefixMode, basePrefix, serialDigits, _typeImageBase64)
            If success Then
                ' Create the per-type inventory table automatically.
                ' forInventory:=True ensures gauge types like "Plain Plug Gauge" get their own
                ' inventory table (plain_plug_gauge) and not the calibration table.
                Dim tableName As String = MySQLClass.TypeNameToTableName(typeName, forInventory:=True)
                Dim tableCreated = _mySql.CreateTypeTable(tableName, category)
                If Not tableCreated Then
                    MessageBox.Show($"Type saved but failed to create database table '{tableName}'. Please check DB connection.", "Table Creation Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                End If
            End If
        End If

        If success Then
            MessageBox.Show("Record saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("Failed to save record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Public Sub PopulateForm(row As DataRow)
        If row Is Nothing Then Return

        IsEditMode = True
        RecordID = Convert.ToInt32(row("ID"))
        
        ' Set Category
        Dim catVal = row("Category").ToString()
        For Each item As ComboBoxItem In ComboCategory.Items
            If item.Content.ToString() = catVal Then
                ComboCategory.SelectedItem = item
                Exit For
            End If
        Next

        _oldTypeName = row("TypeName").ToString()
        TxtTypeName.Text = _oldTypeName

        ' Set Prefix Mode
        Dim modeVal = row("PrefixMode").ToString()
        For Each item As ComboBoxItem In ComboPrefixMode.Items
            If item.Content.ToString() = modeVal Then
                ComboPrefixMode.SelectedItem = item
                Exit For
            End If
        Next

        TxtBasePrefix.Text = row("BasePrefix").ToString()
        TxtSerialDigits.Text = row("SerialDigits").ToString()

        If row.Table.Columns.Contains("TypeImage") Then
            _typeImageBase64 = row("TypeImage").ToString()
            SetImageFromBase64(_typeImageBase64)
        End If
    End Sub
End Class
