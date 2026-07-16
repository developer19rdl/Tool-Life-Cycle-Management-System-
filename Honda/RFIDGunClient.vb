Imports System.Net.Http
Imports System.IO
Imports System.Threading
Imports System.Text.Json
Imports System.Threading.Tasks

Public Class RFIDGunClient
    Public Shared Event RFIDScanned(rfid As String)

    Private _httpClient As HttpClient
    Private _cts As CancellationTokenSource

    Public Sub New()
        _httpClient = New HttpClient()
        _httpClient.Timeout = Timeout.InfiniteTimeSpan
    End Sub

    Public Async Function StartStreamingAsync(ip As String, port As String) As Task
        StopStreaming()
        _cts = New CancellationTokenSource()
        Dim token = _cts.Token

        Dim url = $"http://{ip}:{port}/api/session/stream"

        Try
            Await Task.Run(Async Function()
                               Dim request = New HttpRequestMessage(HttpMethod.Get, url)
                               request.Headers.Add("Accept", "text/event-stream")

                               Using response = Await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(False)
                                   response.EnsureSuccessStatusCode()

                                   Using stream = Await response.Content.ReadAsStreamAsync(token).ConfigureAwait(False)
                                       Using reader = New StreamReader(stream)
                                            While Not token.IsCancellationRequested AndAlso Not reader.EndOfStream
                                                Dim line = Await reader.ReadLineAsync().ConfigureAwait(False)
                                                If line Is Nothing Then Exit While

                                                If line.StartsWith("data:") Then
                                                    Dim json = line.Substring(5).Trim()
                                                    ProcessJson(json)
                                                End If
                                            End While
                                       End Using
                                   End Using
                               End Using
                           End Function, token)
        Catch ex As System.Net.Http.HttpRequestException
            Dim msg = $"[RFIDGunClient] Connection error: {ex.Message}. Ensure the RFID Gun is at {ip}:{port} and the service is running."
            Console.WriteLine(msg)
            ' Optional: Notify UI if needed
        Catch ex As OperationCanceledException
            ' Normal shutdown
        Catch ex As Exception
            Console.WriteLine($"[RFIDGunClient] Unexpected Stream Error: {ex.Message}")
        End Try
    End Function

    Public Sub StopStreaming()
        If _cts IsNot Nothing Then
            Try
                _cts.Cancel()
            Catch
            End Try
            _cts.Dispose()
            _cts = Nothing
        End If
    End Sub

    Private Sub ProcessJson(json As String)
        Try
            If String.IsNullOrWhiteSpace(json) Then Return

            ' Match ANY "rfid_tag": "VALUE" across the JSON dynamically because live streams 
            ' and session initializations use distinct JSON shapes.
            Dim regex As New System.Text.RegularExpressions.Regex("""rfid_tag""\s*:\s*""([^""]+)""")
            Dim matches = regex.Matches(json)
            
            For Each m As System.Text.RegularExpressions.Match In matches
                Dim rfid As String = m.Groups(1).Value
                If Not String.IsNullOrWhiteSpace(rfid) Then
                    rfid = FormatRfidWithSpaces(rfid)
                    RaiseEvent RFIDScanned(rfid)
                End If
            Next
        Catch ex As Exception
            Console.WriteLine($"[RFIDGunClient] Parse error: {ex.Message}")
        End Try
    End Sub

    Private Shared Function FormatRfidWithSpaces(rfid As String) As String
        If String.IsNullOrWhiteSpace(rfid) Then Return ""
        
        ' Strip out any existing spaces and uppercase it to normalize
        rfid = rfid.Replace(" ", "").ToUpper()
        
        ' Need evenly paired string to easily space it, if odd, just return it
        If rfid.Length Mod 2 <> 0 Then Return rfid
        
        Dim formatted As New System.Text.StringBuilder()
        For i As Integer = 0 To rfid.Length - 1 Step 2
            If i > 0 Then formatted.Append(" ")
            formatted.Append(rfid.Substring(i, 2))
        Next
        
        Return formatted.ToString()
    End Function
End Class
