Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Diagnostics

Public Class CalibrationMasterDetailsWindow
    Private _mySql As New MySQLClass()
    Private _uploadedDocPath As String = ""

    Public Sub LoadDetails(id As Integer, tableName As String)
        Try
            Dim query As String = $"SELECT * FROM `{tableName}` WHERE ID = {id}"
            Dim dtMain As DataTable = _mySql.ReadDatatable(query)

            If dtMain Is Nothing OrElse dtMain.Rows.Count = 0 Then
                MessageBox.Show("No details found for this master.", "Not Found", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                Me.Close()
                Return
            End If

            Dim row As DataRow = dtMain.Rows(0)
            Dim masterName As String = If(IsDBNull(row("Description")), "- -", row("Description").ToString())

            TxtHeaderInfo.Text = masterName

            FieldsContainer.Children.Clear()
            _uploadedDocPath = ""
            BtnViewDocument.Visibility = Visibility.Collapsed

            ' Filter display columns
            Dim displayCols As New Dictionary(Of String, String) From {
                {"Date", "Entry Date"},
                {"Description", "Description"},
                {"LeastCount", "Least Count"},
                {"MasterUncertainty", "Master Uncertainty"},
                {"CalDate", "Calibration Date"},
                {"DueDate", "Due Date"}
            }

            For Each col In displayCols
                If dtMain.Columns.Contains(col.Key) Then
                    Dim rawValue As Object = row(col.Key)
                    Dim strValue As String = If(IsDBNull(rawValue) OrElse String.IsNullOrWhiteSpace(rawValue.ToString()), "- -", rawValue.ToString())

                    If TypeOf rawValue Is DateTime Then
                        strValue = Convert.ToDateTime(rawValue).ToString("dd-MM-yyyy")
                    End If

                    AddField(col.Value, strValue)
                End If
            Next

            ' Document logic
            If dtMain.Columns.Contains("uploaded_doc") Then
                _uploadedDocPath = If(IsDBNull(row("uploaded_doc")), "", row("uploaded_doc").ToString())
                If Not String.IsNullOrEmpty(_uploadedDocPath) Then
                    Dim appPath = AppDomain.CurrentDomain.BaseDirectory
                    Dim fullPath = System.IO.Path.Combine(appPath, _uploadedDocPath)

                    If System.IO.File.Exists(fullPath) Then
                        BtnViewDocument.Visibility = Visibility.Visible
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading details: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub AddField(header As String, value As String)
        Dim fieldPanel As New StackPanel()
        fieldPanel.HorizontalAlignment = HorizontalAlignment.Left
        fieldPanel.Margin = New Thickness(0, 0, 0, 16)

        Dim lblHeader As New TextBlock()
        lblHeader.Text = header
        lblHeader.FontSize = 12
        lblHeader.FontWeight = FontWeights.SemiBold
        lblHeader.Foreground = New SolidColorBrush(Color.FromRgb(100, 116, 139))
        lblHeader.Margin = New Thickness(0, 0, 0, 2)

        Dim lblValue As New TextBlock()
        lblValue.Text = value
        lblValue.FontSize = 14
        lblValue.Foreground = New SolidColorBrush(Color.FromRgb(30, 41, 59))
        lblValue.TextWrapping = TextWrapping.Wrap

        fieldPanel.Children.Add(lblHeader)
        fieldPanel.Children.Add(lblValue)
        FieldsContainer.Children.Add(fieldPanel)
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub BtnViewDocument_Click(sender As Object, e As RoutedEventArgs)
        Try
            If String.IsNullOrEmpty(_uploadedDocPath) Then Return
            Dim appPath = AppDomain.CurrentDomain.BaseDirectory
            Dim fullPath = System.IO.Path.Combine(appPath, _uploadedDocPath)

            If System.IO.File.Exists(fullPath) Then
                Process.Start(New ProcessStartInfo(fullPath) With {.UseShellExecute = True})
            Else
                MessageBox.Show("Document file not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Error opening document: " & ex.Message)
        End Try
    End Sub
End Class
