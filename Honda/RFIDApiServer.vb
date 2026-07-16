Imports System.Net
Imports System.Text
Imports System.Threading
Imports MySql.Data.MySqlClient

''' <summary>
''' Lightweight embedded HTTP API server.
''' Listens on http://localhost:8080/api/lookup?rfid=XXXX
''' and returns a JSON response with control_number, type, name, colour.
''' </summary>
Public Class RFIDApiServer

    Private _listener As HttpListener
    Private _listenerThread As Thread
    Private _running As Boolean = False

    ''' <summary>The port the API listens on. Change if 8080 is already in use.</summary>
    Public Const Port As Integer = 8080

    ''' <summary>
    ''' Starts the HTTP listener on a background thread.
    ''' Call this once from MainWindow after the app loads.
    ''' NOTE: To allow external devices to reach this API, open port 8080 in Windows Firewall:
    '''   netsh advfirewall firewall add rule name="Honda RFID API" dir=in action=allow protocol=TCP localport=8080
    ''' </summary>
    Public Sub Start()
        Try
            _listener = New HttpListener()
            ' Listen on /api/tools/ path
            _listener.Prefixes.Add($"http://+:{Port}/api/tools/")
            _listener.Start()
            _running = True

            _listenerThread = New Thread(AddressOf ListenLoop)
            _listenerThread.IsBackground = True
            _listenerThread.Name = "RFIDApiListener"
            _listenerThread.Start()

            Console.WriteLine($"[RFIDApiServer] Started on port {Port}.")
        Catch ex As Exception
            Console.WriteLine($"[RFIDApiServer] Failed to start: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Stops the HTTP listener cleanly.
    ''' Call this from MainWindow.Closing event.
    ''' </summary>
    Public Sub [Stop]()
        Try
            _running = False
            _listener?.Stop()
            _listener?.Close()
            Console.WriteLine("[RFIDApiServer] Stopped.")
        Catch ex As Exception
            Console.WriteLine($"[RFIDApiServer] Error stopping: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Main loop — waits for incoming HTTP requests and processes them.
    ''' </summary>
    Private Sub ListenLoop()
        While _running
            Try
                Dim context As HttpListenerContext = _listener.GetContext()
                ' Handle each request on a separate thread to avoid blocking
                ThreadPool.QueueUserWorkItem(Sub(c) HandleRequest(CType(c, HttpListenerContext)), context)
            Catch ex As HttpListenerException
                ' Listener was stopped intentionally
                If _running Then
                    Console.WriteLine($"[RFIDApiServer] Listener error: {ex.Message}")
                End If
            Catch ex As Exception
                Console.WriteLine($"[RFIDApiServer] Unexpected error: {ex.Message}")
            End Try
        End While
    End Sub

    ''' <summary>
    ''' Event raised when an RFID is scanned via the direct /scan/ endpoint.
    ''' </summary>
    Public Shared Event RFIDScanned(rfid As String)

    ''' <summary>
    ''' Handles a single HTTP request. Parses the rfid query param,
    ''' queries MySQL, and writes the JSON response.
    ''' </summary>
    Private Sub HandleRequest(context As HttpListenerContext)
        Dim request As HttpListenerRequest = context.Request
        Dim response As HttpListenerResponse = context.Response

        ' Allow cross-origin requests (CORS) so browser clients can call this
        response.Headers.Add("Access-Control-Allow-Origin", "*")
        response.ContentType = "application/json; charset=utf-8"

        Try
            ' Expect path: /api/tools/epc/{rfid} OR /api/tools/show/{rfid} OR /api/tools/scan/{rfid}
            Dim path As String = request.Url.AbsolutePath.ToLower().TrimEnd("/"c)
            Dim isShowRequest As Boolean = path.StartsWith("/api/tools/show/")
            Dim isScanRequest As Boolean = path.StartsWith("/api/tools/scan/")
            
            Dim prefix As String = ""
            If isShowRequest Then
                prefix = "/api/tools/show/"
            ElseIf isScanRequest Then
                prefix = "/api/tools/scan/"
            Else
                prefix = "/api/tools/epc/"
            End If

            If request.HttpMethod.ToUpper() <> "GET" OrElse Not path.StartsWith(prefix) Then
                SendJson(response, 404, "{""error"":""Not found.""}") 
                Return
            End If

            ' Extract RFID list from the path segment
            Dim rfidParam As String = request.Url.AbsolutePath.Substring(prefix.Length).Trim()
            
            If String.IsNullOrWhiteSpace(rfidParam) Then
                SendJson(response, 400, "{""error"":""Missing required rfid parameter(s) in path""}")
                Return
            End If

            ' If it's a direct SCAN request (Bulk Scan Gun), just raise the event and return success
            If isScanRequest Then
                RaiseEvent RFIDScanned(rfidParam)
                SendJson(response, 200, "{""success"":true,""message"":""Scan received""}")
                Return
            End If

            ' Split by comma to support multiple RFIDs for Lookup/Show
            Dim rfidTags As String() = rfidParam.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)
            Dim results As New List(Of RFIDLookupResult)()

            For Each t In rfidTags
                Dim r As RFIDLookupResult = LookupByRFID(t.Trim())
                If r IsNot Nothing Then
                    results.Add(r)
                End If
            Next

            If results.Count = 0 Then
                SendJson(response, 404, $"{{""success"":false,""error"":""No records found for the provided RFID tag(s).""}}")
            Else
                ' If it's a "show" request, pop up the window on the UI thread
                If isShowRequest Then
                    Application.Current.Dispatcher.BeginInvoke(Sub()
                        Dim popup As New RfidDetailPopup(results)
                        popup.Topmost = True
                        popup.Show()
                    End Sub)
                End If

                ' Build JSON response for multiple results
                Dim dataParts As New List(Of String)()
                For Each r In results
                    dataParts.Add($"{{""epc"":""{EscapeJson(r.Rfid)}"",""tool_name"":""{EscapeJson(r.Name)}"",""tool_code"":""{EscapeJson(r.ControlNumber)}"",""colour"":""{EscapeJson(r.Colour)}"",""type"":""{EscapeJson(r.Type)}""}}")
                Next

                Dim json As String = "{" &
                    $"""success"":true," &
                    $"""count"":{results.Count}," &
                    $"""data"":[{String.Join(",", dataParts)}]" &
                    "}"
                SendJson(response, 200, json)
            End If

        Catch ex As Exception
            Console.WriteLine($"[RFIDApiServer] HandleRequest error: {ex.Message}")
            SendJson(response, 500, $"{{""error"":""Internal server error: {EscapeJson(ex.Message)}""}}")
        End Try
    End Sub

    ''' <summary>
    ''' Normalizes an RFID tag by stripping all spaces, converting to uppercase.
    ''' Handles both formats:
    '''   Android sends:  E28069950000501400000000  (no spaces)
    '''   DB stores:      E2 80 69 95 00 00 50 14...  (with spaces)
    ''' By stripping spaces from BOTH sides, they always match.
    ''' </summary>
    Private Function NormalizeRFID(rfid As String) As String
        If rfid Is Nothing Then Return ""
        Return rfid.Replace(" ", "").ToUpper().Trim()
    End Function

    ''' <summary>
    ''' Queries all per-type tables dynamically from type_details for the given RFID tag.
    ''' Accepts both spaced ("E2 80 69...") and no-space ("E28069...") formats.
    ''' Returns Nothing if no match found.
    ''' </summary>
    Private Function LookupByRFID(rfid As String) As RFIDLookupResult
        Try
            Dim normalizedRfid As String = NormalizeRFID(rfid)
            Dim settings = ProjectSettings.Current
            Dim conStr As String = $"datasource={settings.Datasource};port={settings.Port};database={settings.Database};username={settings.Username};password={settings.Password};"

            Using con As New MySqlConnection(conStr)
                con.Open()

                ' Step 1: Load all type tables from type_details
                Dim typeRows As New List(Of (tableName As String, category As String))()
                Using cmd As New MySqlCommand("SELECT TypeName, Category FROM type_details", con)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim typeName = reader("TypeName").ToString()
                            Dim cat = reader("Category").ToString()
                            Dim tbl = System.Text.RegularExpressions.Regex.Replace(typeName.ToLower().Trim(), "[^a-z0-9]+", "_").Trim("_"c)
                            typeRows.Add((tbl, cat))
                        End While
                    End Using
                End Using

                If typeRows.Count = 0 Then Return Nothing

                ' Step 2: Build UNION ALL query across all per-type tables
                Dim unionParts As New List(Of String)()
                For Each t In typeRows
                    If t.category = "Instrument" Then
                        unionParts.Add($"SELECT ControlNo, 'Instrument' AS Type, InstrumentName AS Name, Color, RFID_tag FROM `{t.tableName}` WHERE UPPER(REPLACE(RFID_tag, ' ', '')) = @rfid")
                    Else
                        unionParts.Add($"SELECT ControlNo, 'Gauge' AS Type, GaugeName AS Name, Color, RFID_tag FROM `{t.tableName}` WHERE UPPER(REPLACE(RFID_tag, ' ', '')) = @rfid")
                    End If
                Next

                Dim query As String = String.Join(" UNION ALL ", unionParts) & " LIMIT 1"

                Using cmd As New MySqlCommand(query, con)
                    cmd.Parameters.AddWithValue("@rfid", normalizedRfid)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return New RFIDLookupResult With {
                                .ControlNumber = reader("ControlNo").ToString(),
                                .Type = reader("Type").ToString(),
                                .Name = reader("Name").ToString(),
                                .Colour = reader("Color").ToString(),
                                .Rfid = NormalizeRFID(reader("RFID_tag").ToString())
                            }
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine($"[RFIDApiServer] LookupByRFID error: {ex.Message}")
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Writes the JSON string to the HTTP response with the given status code.
    ''' </summary>
    Private Sub SendJson(response As HttpListenerResponse, statusCode As Integer, json As String)
        Try
            response.StatusCode = statusCode
            Dim buffer As Byte() = Encoding.UTF8.GetBytes(json)
            response.ContentLength64 = buffer.Length
            response.OutputStream.Write(buffer, 0, buffer.Length)
        Catch ex As Exception
            Console.WriteLine($"[RFIDApiServer] SendJson error: {ex.Message}")
        Finally
            Try
                response.OutputStream.Close()
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' Escapes special characters in a string for safe embedding in JSON.
    ''' </summary>
    Private Function EscapeJson(value As String) As String
        If value Is Nothing Then Return ""
        Return value.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n").Replace(vbTab, "\t")
    End Function

End Class

''' <summary>
''' Holds the result of an RFID lookup from the database.
''' </summary>
Public Class RFIDLookupResult
    Public Property ControlNumber As String
    Public Property Type As String
    Public Property Name As String
    Public Property Colour As String
    Public Property Rfid As String
End Class
