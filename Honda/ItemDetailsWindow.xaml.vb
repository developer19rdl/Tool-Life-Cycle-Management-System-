Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media

Public Class ItemDetailsWindow
    Private _mySql As New MySQLClass()
    Private _uploadedDocPath As String = ""

    Public Sub LoadDetails(controlNo As String, itemType As String, tableName As String)
        Try
            Dim nameColumn As String = If(itemType.Equals("Gauge", StringComparison.OrdinalIgnoreCase), "GaugeName", "InstrumentName")

            Dim query As String = $"SELECT * FROM `{tableName}` WHERE ControlNo = '{controlNo.Replace("'", "''")}'"
            Dim dtMain As DataTable = _mySql.ReadDatatable(query)

            If dtMain Is Nothing OrElse dtMain.Rows.Count = 0 Then
                MessageBox.Show("No details found for this item.", "Not Found", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                Me.Close()
                Return
            End If

            Dim row As DataRow = dtMain.Rows(0)
            Dim itemName As String = If(IsDBNull(row(nameColumn)), "- -", row(nameColumn).ToString())

            ' 2. Set Header Labels
            TxtTypeTitle.Text = itemType & " Details"
            TxtHeaderInfo.Text = controlNo & "  |  " & itemName

            ' 3. Clear container for fresh data
            FieldsContainer.Children.Clear()
            _uploadedDocPath = ""
            BtnViewDocument.Visibility = Visibility.Collapsed

            ' 4. Loop through main table columns and display them
            For Each col As DataColumn In dtMain.Columns
                Dim colName As String = col.ColumnName

                ' Skip system/hidden columns
                If colName = "ID" OrElse colName = "Time" OrElse colName = "Flag" OrElse colName = "DateofAddition" Then Continue For

                If colName = "uploaded_doc" Then
                    _uploadedDocPath = If(IsDBNull(row(colName)), "", row(colName).ToString())
                    If Not String.IsNullOrEmpty(_uploadedDocPath) Then
                        ' Check if file physically exists
                        Dim appPath = AppDomain.CurrentDomain.BaseDirectory
                        Dim fullPath = System.IO.Path.Combine(appPath, _uploadedDocPath)

                        If System.IO.File.Exists(fullPath) Then
                            BtnViewDocument.Visibility = Visibility.Visible
                            AddField("Uploaded Document", System.IO.Path.GetFileName(_uploadedDocPath))
                        End If
                    End If
                    Continue For
                End If

                Dim rawValue As Object = row(colName)
                Dim strValue As String = If(IsDBNull(rawValue) OrElse String.IsNullOrWhiteSpace(rawValue.ToString()), "- -", rawValue.ToString())

                ' Format dates nicely if it's a date column
                If TypeOf rawValue Is DateTime Then
                    strValue = Convert.ToDateTime(rawValue).ToString("dd-MM-yyyy")
                End If

                AddField(colName, strValue)
            Next

            ' 5. Fetch Write-Off Data
            Dim woQuery As String = $"SELECT WriteOffDate FROM writeoff WHERE ControlNo = '{controlNo.Replace("'", "''")}' ORDER BY ID DESC LIMIT 1"
            Dim dtWO As DataTable = _mySql.ReadDatatable(woQuery)
            Dim woVal As String = "- -"
            If dtWO IsNot Nothing AndAlso dtWO.Rows.Count > 0 AndAlso Not IsDBNull(dtWO.Rows(0)(0)) Then
                woVal = Convert.ToDateTime(dtWO.Rows(0)(0)).ToString("dd-MM-yyyy")
            End If
            AddField("Write-off Date", woVal)

            ' 6. Fetch Received Data
            Dim recQuery As String = $"SELECT Date FROM recieve WHERE ControlNo = '{controlNo.Replace("'", "''")}' ORDER BY ID DESC LIMIT 1"
            Dim dtRec As DataTable = _mySql.ReadDatatable(recQuery)
            Dim recVal As String = "- -"
            If dtRec IsNot Nothing AndAlso dtRec.Rows.Count > 0 AndAlso Not IsDBNull(dtRec.Rows(0)(0)) Then
                recVal = Convert.ToDateTime(dtRec.Rows(0)(0)).ToString("dd-MM-yyyy")
            End If
            AddField("Received Date", recVal)

        Catch ex As Exception
            MessageBox.Show("Error loading details: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub AddField(header As String, value As String)
        Dim fieldPanel As New StackPanel()
        fieldPanel.Width = 350
        fieldPanel.Margin = New Thickness(0, 0, 20, 16)

        Dim lblHeader As New TextBlock()
        lblHeader.Text = header
        lblHeader.FontSize = 12
        lblHeader.FontWeight = FontWeights.SemiBold
        lblHeader.Foreground = New SolidColorBrush(Color.FromRgb(100, 116, 139)) ' #64748B
        lblHeader.Margin = New Thickness(0, 0, 0, 2)

        Dim lblValue As New TextBlock()
        lblValue.Text = value
        lblValue.FontSize = 14
        lblValue.Foreground = New SolidColorBrush(Color.FromRgb(30, 41, 59)) ' #1E293B
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
                MessageBox.Show("Document file not found at " & fullPath, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Error opening document: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class
