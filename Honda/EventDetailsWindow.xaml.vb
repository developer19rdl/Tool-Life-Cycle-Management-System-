Imports System.Windows
Imports System.Diagnostics
Imports System.IO

Public Class EventDetailsWindow
    Private _docPath As String = ""
    Private _reintroDocPath As String = ""
    Private _actionName As String = ""

    Public Sub New(eventTitle As String, details As Dictionary(Of String, String))
        InitializeComponent()
        
        TxtEventTitle.Text = eventTitle
        
        ' Extract action name from title (e.g., "WOP" from "Event Details: WOP(Reintroduced)")
        Dim actionName = eventTitle.Replace("Event Details: ", "").Replace("(Reintroduced)", "").Trim()
        _actionName = actionName
        
        Dim hasActionRemarks = details.ContainsKey("ActionRemarks") AndAlso Not String.IsNullOrWhiteSpace(details("ActionRemarks"))
        Dim hasReintroRemarks = details.ContainsKey("ReintroRemarks") AndAlso Not String.IsNullOrWhiteSpace(details("ReintroRemarks"))

        ' 1. Special Case: Reintroduction (Simplify UI to match Write off)
        If actionName.ToLower().Contains("reintroduction") OrElse actionName.ToLower().Contains("re-introduction") Then
            TxtMainRemarksLabel.Text = "Remarks / Reason:"
            TxtMainRemarksLabel.Visibility = Visibility.Visible
            
            TxtActionLabel.Visibility = Visibility.Collapsed
            SeparatorReintro.Visibility = Visibility.Collapsed
            TxtReintroLabel.Visibility = Visibility.Collapsed
            TxtReintroRemarks.Visibility = Visibility.Collapsed
            
            ' Only take reintroduction remarks for this event (from reintroduction table)
            Dim finalRemarks = If(hasReintroRemarks, details("ReintroRemarks"), details("ActionRemarks"))
            
            TxtRemarks.Text = If(String.IsNullOrWhiteSpace(finalRemarks), "No remarks available.", finalRemarks)
            TxtRemarks.Visibility = Visibility.Visible
        Else
            ' 2. Normal Case: Other actions (WOP, Write off, etc.)
            Dim isSkipAction = actionName.Equals("Pending", StringComparison.OrdinalIgnoreCase)
            If isSkipAction OrElse Not hasActionRemarks Then
                TxtRemarks.Visibility = Visibility.Collapsed
                TxtMainRemarksLabel.Visibility = Visibility.Collapsed
                TxtActionLabel.Visibility = Visibility.Collapsed
            Else
                TxtRemarks.Text = details("ActionRemarks")
                TxtRemarks.Visibility = Visibility.Visible
                TxtMainRemarksLabel.Text = "Remarks / Reason:"
                TxtMainRemarksLabel.Visibility = Visibility.Visible
            End If

            ' Manage Document Path for Action
            If details.ContainsKey("ActionDoc") AndAlso Not String.IsNullOrWhiteSpace(details("ActionDoc")) Then
                _docPath = details("ActionDoc")
                BtnViewDoc.Visibility = Visibility.Visible
            End If

            ' Reintroduction Details (Sub-section for combined events like WOP(Reintroduced))
            If hasReintroRemarks Then
                If isSkipAction OrElse Not hasActionRemarks Then
                    TxtMainRemarksLabel.Visibility = Visibility.Visible
                    TxtMainRemarksLabel.Text = "Re-introduction Remarks:"
                Else
                    TxtMainRemarksLabel.Visibility = Visibility.Collapsed
                    TxtActionLabel.Visibility = Visibility.Visible
                    TxtActionLabel.Text = actionName & " Remarks:"
                End If

                SeparatorReintro.Visibility = If(isSkipAction OrElse Not hasActionRemarks, Visibility.Collapsed, Visibility.Visible)
                TxtReintroLabel.Visibility = If(isSkipAction OrElse Not hasActionRemarks, Visibility.Collapsed, Visibility.Visible)
                TxtReintroRemarks.Visibility = Visibility.Visible
                TxtReintroRemarks.Text = details("ReintroRemarks")
            End If
        End If

        ' 3. Manage Document Buttons
        If details.ContainsKey("ActionDoc") AndAlso Not String.IsNullOrWhiteSpace(details("ActionDoc")) Then
            _docPath = details("ActionDoc")
            BtnViewDoc.Visibility = Visibility.Visible
        End If
        If details.ContainsKey("ReintroDoc") AndAlso Not String.IsNullOrWhiteSpace(details("ReintroDoc")) Then
            _reintroDocPath = details("ReintroDoc")
            BtnViewDoc.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub BtnViewDoc_Click(sender As Object, e As RoutedEventArgs)
        If Not String.IsNullOrEmpty(_docPath) AndAlso Not String.IsNullOrEmpty(_reintroDocPath) Then
            ' Multiple documents - show selection menu
            Dim contextMenu As New ContextMenu()
            contextMenu.Style = CType(Me.FindResource("ModernContextMenu"), Style)
            
            Dim actionItem As New MenuItem()
            actionItem.Header = "View " & _actionName & " Document"
            actionItem.Style = CType(Me.FindResource("ModernMenuItem"), Style)
            AddHandler actionItem.Click, Sub() OpenDocument(_docPath)
            contextMenu.Items.Add(actionItem)
            
            Dim reintroItem As New MenuItem()
            reintroItem.Header = "View Re-introduction Document"
            reintroItem.Style = CType(Me.FindResource("ModernMenuItem"), Style)
            AddHandler reintroItem.Click, Sub() OpenDocument(_reintroDocPath)
            contextMenu.Items.Add(reintroItem)
            
            contextMenu.PlacementTarget = BtnViewDoc
            contextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom
            contextMenu.IsOpen = True
        ElseIf Not String.IsNullOrEmpty(_docPath) Then
            OpenDocument(_docPath)
        ElseIf Not String.IsNullOrEmpty(_reintroDocPath) Then
            OpenDocument(_reintroDocPath)
        End If
    End Sub

    Private Sub OpenDocument(path As String)
        Try
            If Not String.IsNullOrEmpty(path) AndAlso File.Exists(path) Then
                Process.Start(New ProcessStartInfo(path) With {.UseShellExecute = True})
            Else
                MessageBox.Show("The document file was not found at the specified path:" & vbCrLf & path, "File Not Found", MessageBoxButton.OK, MessageBoxImage.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to open the document. Error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class
