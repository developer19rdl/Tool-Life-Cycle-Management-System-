Imports System.IO.Ports
Imports System.Management
Imports ThingMagic
Imports System.Threading
Imports System.Windows

Public Class CommManager
    ''' <summary>
    ''' Entry point for scanning a single RFID tag.
    ''' Switches between ThinkMagic and RDL implementations based on settings.
    ''' </summary>
    Public Shared Function ScanSingleTag(Optional portName As String = Nothing, Optional timeoutMs As Integer = 500) As String
        Dim scannerType = ProjectSettings.Current.RfidScannerType
        Dim startTime = DateTime.Now
        
        ' If timeout is > 1s, we assume it's a UI-driven scan and we loop/retry
        If timeoutMs > 1000 Then
            While (DateTime.Now - startTime).TotalMilliseconds < timeoutMs
                Dim result = ""
                If scannerType = "RDL" Then
                    result = ScanSingleTagRDL(portName, 200, True) ' Ultra-fast Window (200ms)
                Else
                    result = ScanSingleTagThinkMagic(portName, 200, True) ' Ultra-fast Window (200ms)
                End If
                
                If Not String.IsNullOrEmpty(result) AndAlso IsValidTagPrefix(result) Then Return result
                Thread.Sleep(10) ' Minimal pause (10ms)
            End While
            Return Nothing
        Else
            ' Background polling or short scan
            Dim finalResult = ""
            If scannerType = "RDL" Then
                finalResult = ScanSingleTagRDL(portName, timeoutMs, False)
            Else
                finalResult = ScanSingleTagThinkMagic(portName, timeoutMs, False)
            End If

            If IsValidTagPrefix(finalResult) Then
                Return finalResult
            Else
                Return Nothing
            End If
        End If
    End Function

    ''' <summary>
    ''' Entry point for scanning multiple RFID tags (Bulk Scan).
    ''' </summary>
    Public Shared Function ScanBulkTags(Optional portName As String = Nothing, Optional timeoutMs As Integer = 5000) As List(Of String)
        Dim scannerType = ProjectSettings.Current.RfidScannerType
        If scannerType = "RDL" Then
            Return ScanBulkTagsRDL(portName, timeoutMs)
        Else
            Return ScanBulkTagsThinkMagic(portName, timeoutMs)
        End If
    End Function

    ''' <summary>
    ''' Scans for a single RFID tag using the ThingMagic M6e Micro reader.
    ''' Returns the EPC as a space-separated hex string (e.g., "E2 80 69 95 00 00 50 14 00 00 00 00").
    ''' Returns Nothing if no tag is found or an error occurs.
    ''' </summary>
    Private Shared Function ScanSingleTagThinkMagic(Optional portName As String = Nothing, Optional timeoutMs As Integer = 500, Optional suppressErrors As Boolean = False) As String
        Dim reader As Reader = Nothing
        Try
            ' Priority: Provided -> Settings -> Auto-detect
            If String.IsNullOrEmpty(portName) Then
                portName = ProjectSettings.Current.RfidPort
                If String.IsNullOrEmpty(portName) Then
                    Dim ports = GetUSBSerialPorts()
                    If ports.Count = 0 Then
                        If Not suppressErrors Then
                            System.Windows.MessageBox.Show("No USB COM ports found. Please check device manager.", "Port Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
                        End If
                        Return Nothing
                    End If
                    portName = ports(0).PortName
                End If
            End If

            Dim uriString = $"tmr:///{portName}"
            
            Try
                reader = Reader.Create(uriString)
                ' Set baud rate before connecting since 115200 is standard for M6e Micro
                reader.ParamSet("/reader/baudRate", CInt(115200))
                reader.Connect()
            Catch exConn As Exception
                If Not suppressErrors Then
                    Dim errorMsg = $"Failed to connect to reader on {portName}. Is another app (like Universal Reader Assistant) using it? Error: {exConn.Message}"
                    Console.WriteLine(errorMsg)
                    System.Windows.MessageBox.Show(errorMsg, "RFID Connection Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
                End If
                Return Nothing
            End Try

            Try
                ' Configure for single tag read
                reader.ParamSet("/reader/region/id", Reader.Region.IN)
                reader.ParamSet("/reader/read/asyncOnTime", CInt(250))
                
                ' Apply Scan Power (Distance)
                Dim pwr = ProjectSettings.Current.RfidPower
                If pwr >= 5 AndAlso pwr <= 30 Then
                    reader.ParamSet("/reader/radio/readPower", CInt(pwr * 100)) ' Unit: centidBm
                End If

                ' Read tags with specified timeout
                Dim tagReads = reader.Read(timeoutMs)

                If tagReads IsNot Nothing AndAlso tagReads.Length > 0 Then
                    ' Get the first (strongest signal) tag
                    Dim bestTag = tagReads(0)
                    
                    ' TagReadData.EpcString gives the hex EPC directly
                    Dim epcHex = bestTag.EpcString
                    
                    ' Format as space-separated pairs (e.g., "E2 80 69 95 00 00 50 14")
                    If Not String.IsNullOrEmpty(epcHex) Then
                        Return FormatEpc(epcHex)
                    End If
                End If

                Return Nothing
            Catch exRead As Exception
                 If Not suppressErrors Then
                    System.Windows.MessageBox.Show($"Reader connected but failed to scan: {exRead.Message}", "RFID Read Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
                 End If
                 Return Nothing
            End Try

        Catch ex As Exception
            Dim errorMsg = $"Unexpected RFID error on {portName}: {ex.Message}"
            Console.WriteLine(errorMsg)
            System.Windows.MessageBox.Show(errorMsg, "RFID Scanner Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
            Return Nothing
        Finally
            Try
                If reader IsNot Nothing Then
                    reader.Destroy()
                End If
            Catch ex2 As Exception
                Console.WriteLine("Cleanup error: " & ex2.Message)
            End Try
        End Try
    End Function

    ''' <summary>
    ''' Shared buffer for incoming serial data (V.1.1.4 style).
    ''' </summary>
    Public Shared Message2 As String = ""
    Private Shared ReadOnly MessageLock As New Object()

    ''' <summary>
    ''' Event handler to capture incoming RDL hex data into Message2 buffer.
    ''' </summary>
    Private Shared Sub RdlDataReceived(sender As Object, e As SerialDataReceivedEventArgs)
        Try
            Dim comPort = CType(sender, SerialPort)
            Dim bytes As Integer = comPort.BytesToRead
            If bytes <= 0 Then Return

            Dim buffer(bytes - 1) As Byte
            comPort.Read(buffer, 0, bytes)

            ' Convert to hex string with spaces (e.g., "BB 01 22...")
            Dim builder As New System.Text.StringBuilder(buffer.Length * 3)
            For Each b As Byte In buffer
                builder.Append(Convert.ToString(b, 16).PadLeft(2, "0"c).ToUpper() & " ")
            Next
            
            SyncLock MessageLock
                Message2 &= builder.ToString()
            End SyncLock
        Catch ex As Exception
            Console.WriteLine("RDL Receive Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Scans for a single RFID tag using the RDL RFID scanner (SerialPort based).
    ''' Reverted to EXACT V.1.1.4 implementation style using Message2 buffer.
    ''' </summary>
    Private Shared Function ScanSingleTagRDL(Optional portName As String = Nothing, Optional timeoutMs As Integer = 500, Optional suppressErrors As Boolean = False) As String
        Dim comPort As New SerialPort()
        Try
            ' Priority: Provided -> Settings -> Auto-detect
            If String.IsNullOrEmpty(portName) Then
                portName = ProjectSettings.Current.RfidPort
                If String.IsNullOrEmpty(portName) Then
                    Dim ports = GetUSBSerialPorts()
                    If ports.Count = 0 Then
                        If Not suppressErrors Then
                            System.Windows.MessageBox.Show("No USB COM ports found. Please check device manager.", "Port Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
                        End If
                        Return Nothing
                    End If
                    portName = ports(0).PortName
                End If
            End If

            comPort.PortName = portName
            comPort.BaudRate = 115200
            comPort.DataBits = 8
            comPort.Parity = Parity.None
            comPort.StopBits = StopBits.One
            comPort.ReadTimeout = timeoutMs
            
            ' Add handler BEFORE opening (to catch early data)
            AddHandler comPort.DataReceived, AddressOf RdlDataReceived
            
            SyncLock MessageLock
                Message2 = "" ' Clear buffer before start
            End SyncLock
            
            comPort.Open()
            
            ' Apply Scan Power (Distance)
            SetPowerRDL(comPort, ProjectSettings.Current.RfidPower)
            
            ' Clear buffer AFTER power setting response is received to ensure a clean scan start
            SyncLock MessageLock
                Message2 = ""
            End SyncLock
            
            ' RDL Scan Command (Exact V.1.1.4 method call)
            WriteData(comPort, "BB 00 22 00 00 22 7E")

            ' Wait for FULL response (Optimized: 20ms refresh)
            Dim waitCount As Integer = 0
            Dim maxWait = timeoutMs / 20
            If maxWait < 10 Then maxWait = 10 
            
            While waitCount < maxWait
                Thread.Sleep(20) ' Fast hardware polling (20ms)
                
                Dim length As Integer = 0
                SyncLock MessageLock
                    length = Message2.Length
                End SyncLock
                ' Wait until enough data is received
                If length > 50 Then Exit While
                waitCount += 1
            End While
            
            ' Parsing (V3.1: Robust packet scanning - Find BB ... 7E)
            ' This avoids issues if there are multiple packets or noise in the buffer
            Dim raw As String = ""
            SyncLock MessageLock
                raw = Message2 ' Keep original for substring math
            End SyncLock

            If Not String.IsNullOrEmpty(raw) Then
                Dim searchIdx = 0
                While True
                    Dim start = raw.IndexOf("BB ", searchIdx)
                    If start = -1 Then Exit While
                    
                    Dim [end] = raw.IndexOf("7E ", start)
                    If [end] = -1 Then Exit While
                    
                    ' We have a potential packet from start to end
                    Dim packet = raw.Substring(start, [end] - start + 2).Trim()
                    Dim msg = packet.Split(" "c, StringSplitOptions.RemoveEmptyEntries)
                    
                    If msg.Length >= 20 Then
                        ' Extract tag (Index 8 to 19)
                        Dim tag = ""
                        For i As Integer = 8 To 19
                            If i < msg.Length Then tag &= msg(i) & " "
                        Next
                        Dim finalTag = tag.Trim().ToUpper()
                        If IsValidTagPrefix(finalTag) Then Return finalTag
                    End If
                    searchIdx = start + 3 ' Move past current BB
                End While
            End If

            Return Nothing
        Catch ex As Exception
            If Not suppressErrors Then
                Dim errorMsg = $"RDL Scanner error on {portName}: {ex.Message}"
                Console.WriteLine(errorMsg)
                System.Windows.MessageBox.Show(errorMsg, "RFID Scanner Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
            End If
            Return Nothing
        Finally
            Try
                RemoveHandler comPort.DataReceived, AddressOf RdlDataReceived
                If comPort.IsOpen Then comPort.Close()
            Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' Scans for multiple RFID tags using the RDL RFID scanner.
    ''' Uses the V.1.1.4 style buffer to capture multiple packets.
    ''' </summary>
    Private Shared Function ScanBulkTagsRDL(Optional portName As String = Nothing, Optional timeoutMs As Integer = 5000) As List(Of String)
        Dim tags As New HashSet(Of String)()
        Dim comPort As New SerialPort()
        Try
            ' Priority: Provided -> Settings -> Auto-detect
            If String.IsNullOrEmpty(portName) Then
                portName = ProjectSettings.Current.RfidPort
                If String.IsNullOrEmpty(portName) Then
                    Dim ports = GetUSBSerialPorts()
                    If ports.Count = 0 Then Return New List(Of String)()
                    portName = ports(0).PortName
                End If
            End If

            comPort.PortName = portName
            comPort.BaudRate = 115200
            comPort.DataBits = 8
            comPort.Parity = Parity.None
            comPort.StopBits = StopBits.One
            comPort.ReadTimeout = 1000
            
            AddHandler comPort.DataReceived, AddressOf RdlDataReceived
            
            SyncLock MessageLock
                Message2 = ""
            End SyncLock
            
            comPort.Open()

            ' Apply Scan Power (Distance)
            SetPowerRDL(comPort, ProjectSettings.Current.RfidPower)

            ' Send inventory command multiple times to ensure all tags are energized
            Dim startTime = DateTime.Now
            While (DateTime.Now - startTime).TotalMilliseconds < timeoutMs
                WriteData(comPort, "BB 00 22 00 00 22 7E")
                Thread.Sleep(50) ' Ultra-fast poll (50ms)

                ' Parse current buffer for any tags found so far
                Dim raw = ""
                SyncLock MessageLock
                    raw = Message2 ' Keep original for substring math
                End SyncLock

                If Not String.IsNullOrEmpty(raw) Then
                    Dim searchIdx = 0
                    Dim foundEndOfLastPacket = -1
                    
                    ' V3.1: Robust packet scanning (Find BB ... 7E)
                    ' This avoids accidental splitting when a tag EPC contains "BB"
                    While True
                       Dim start = raw.IndexOf("BB ", searchIdx)
                       If start = -1 Then Exit While
                       
                       Dim [end] = raw.IndexOf("7E ", start)
                       If [end] = -1 Then 
                           ' Packet might be incomplete, wait for next buffer poll
                           Exit While
                       End If
                       
                       ' We have a potential packet from start to end
                       Dim packet = raw.Substring(start, [end] - start + 2).Trim()
                       ' Use RemoveEmptyEntries to ensure indices are always consistent (e.g., EPC at index 8)
                       Dim msg = packet.Split(" "c, StringSplitOptions.RemoveEmptyEntries)
                       
                       If msg.Length >= 20 AndAlso msg.Length <= 30 Then
                           ' Extract tag (Index 8 to 19)
                           Dim tag = ""
                           For i As Integer = 8 To 19
                               If i < msg.Length Then tag &= msg(i) & " "
                           Next
                           If Not String.IsNullOrEmpty(tag) AndAlso IsValidTagPrefix(tag) Then tags.Add(tag.Trim().ToUpper())
                           
                           ' Successfully processed a full packet
                           foundEndOfLastPacket = [end] + 2
                           searchIdx = foundEndOfLastPacket
                       Else
                           ' Not a valid tool tag packet (maybe a different command response)
                           ' Move search index forward past this "false" BB
                           searchIdx = start + 3
                       End If
                    End While
                    
                    ' Streaming Buffer Management: Clear handled packets to keep performance high
                    If foundEndOfLastPacket > 0 Then
                        SyncLock MessageLock
                            If foundEndOfLastPacket < Message2.Length Then
                                Message2 = Message2.Substring(foundEndOfLastPacket)
                            Else
                                Message2 = ""
                            End If
                        End SyncLock
                    End If
                End If
            End While

            Return tags.ToList()
        Catch ex As Exception
            Console.WriteLine("Bulk RDL Error: " & ex.Message)
            Return tags.ToList()
        Finally
            Try
                RemoveHandler comPort.DataReceived, AddressOf RdlDataReceived
                If comPort.IsOpen Then comPort.Close()
            Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' Scans for multiple RFID tags using the ThinkMagic reader.
    ''' </summary>
    Private Shared Function ScanBulkTagsThinkMagic(Optional portName As String = Nothing, Optional timeoutMs As Integer = 5000) As List(Of String)
        Dim tags As New HashSet(Of String)()
        Dim reader As Reader = Nothing
        Try
            ' Priority: Provided -> Settings -> Auto-detect
            If String.IsNullOrEmpty(portName) Then
                portName = ProjectSettings.Current.RfidPort
                If String.IsNullOrEmpty(portName) Then
                    Dim ports = GetUSBSerialPorts()
                    If ports.Count = 0 Then Return New List(Of String)()
                    portName = ports(0).PortName
                End If
            End If

            Dim uriString = $"tmr:///{portName}"
            
            Try
                reader = Reader.Create(uriString)
                reader.ParamSet("/reader/baudRate", CInt(115200))
                reader.Connect()
            Catch exConn As Exception
                Console.WriteLine($"Bulk ThinkMagic Connect Error: {exConn.Message}")
                Return New List(Of String)()
            End Try

            Try
                ' Configure for multi-tag read
                reader.ParamSet("/reader/region/id", Reader.Region.IN)
                reader.ParamSet("/reader/read/asyncOnTime", CInt(250))

                ' Apply Scan Power (Distance)
                Dim pwr = ProjectSettings.Current.RfidPower
                If pwr >= 5 AndAlso pwr <= 30 Then
                    reader.ParamSet("/reader/radio/readPower", CInt(pwr * 100))
                End If

                ' Read tags with specified timeout
                Dim tagReads = reader.Read(timeoutMs)

                If tagReads IsNot Nothing Then
                    For Each tr In tagReads
                        Dim epc = FormatEpc(tr.EpcString)
                        ' Only add if it matches the configured prefix
                        If IsValidTagPrefix(epc) Then
                            tags.Add(epc.ToUpper())
                        End If
                    Next
                End If

                Return tags.ToList()
            Catch exRead As Exception
                 Console.WriteLine($"Bulk ThinkMagic Read Error: {exRead.Message}")
                 Return tags.ToList()
            End Try

        Catch ex As Exception
            Console.WriteLine($"Unexpected Bulk ThinkMagic Error: {ex.Message}")
            Return tags.ToList()
        Finally
            Try
                If reader IsNot Nothing Then
                    reader.Destroy()
                End If
            Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' Sets the RF read power for RDL scanners (Indie/MagicRF protocol).
    ''' </summary>
    Private Shared Sub SetPowerRDL(comPort As SerialPort, powerDbm As Integer)
        Try
            ' Ensure power is within typical module limits (5-30 dBm)
            If powerDbm < 5 Then powerDbm = 5
            If powerDbm > 30 Then powerDbm = 30

            ' Command: 0xB6 (Set Read Power)
            ' Using 2-byte power value (dBm * 100)
            Dim powerVal As Integer = powerDbm * 100
            Dim powerH As Byte = CByte((powerVal >> 8) And &HFF)
            Dim powerL As Byte = CByte(powerVal And &HFF)

            ' Calculate Checksum: (Address + Length + Command + PowerH + PowerL) mod 256
            Dim checksum As Byte = CByte((&H0 + &H2 + &HB6 + powerH + powerL) And &HFF)

            ' Packet: BB [Addr] [Cmd] [Len] [P1] [P2] [CS] 7E
            Dim packet As Byte() = {&HBB, &H0, &HB6, &H0, &H2, powerH, powerL, checksum, &H7E}
            
            If comPort.IsOpen Then
                comPort.Write(packet, 0, packet.Length)
                Thread.Sleep(50) ' Brief pause for hardware to process
            End If
        Catch ex As Exception
            Console.WriteLine("SetPowerRDL Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Sends an RDL command as a hex string (Exact V.1.1.4 style).
    ''' </summary>
    Private Shared Sub WriteData(comPort As SerialPort, msg As String)
        Try
            Dim newMsg As Byte() = HexToByte(msg)
            If comPort.IsOpen Then
                comPort.Write(newMsg, 0, newMsg.Length)
            End If
        Catch ex As Exception
            Console.WriteLine("WriteData Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Converts a hex string into a byte array (Exact V.1.1.4 style).
    ''' </summary>
    Private Shared Function HexToByte(ByVal msg As String) As Byte()
        Dim cleanMsg = msg.Replace(" ", "")
        Dim comBuffer As Byte() = New Byte(cleanMsg.Length / 2 - 1) {}
        For i As Integer = 0 To cleanMsg.Length - 1 Step 2
            comBuffer(i / 2) = CByte(Convert.ToByte(cleanMsg.Substring(i, 2), 16))
        Next
        Return comBuffer
    End Function

    ''' <summary>
    ''' Formats a raw hex EPC string into space-separated pairs.
    ''' </summary>
    Private Shared Function FormatEpc(epcHex As String) As String
        Dim formatted As New List(Of String)()
        Dim clean = epcHex.Replace(" ", "").ToUpper()
        For i As Integer = 0 To clean.Length - 1 Step 2
            If i + 1 < clean.Length Then
                formatted.Add(clean.Substring(i, 2))
            End If
        Next
        Return String.Join(" ", formatted)
    End Function

    ''' <summary>
    ''' Finds available serial ports. Prioritizes USB or higher COM ports over internal COM1/COM2.
    ''' </summary>
    ''' <summary>
    ''' Lists available serial ports, prioritized by USB device descriptions.
    ''' Uses WMI to distinguish actual USB-to-Serial adapters from Bluetooth or virtual ports.
    ''' </summary>
    Public Shared Function GetUSBSerialPorts() As List(Of SerialPortInfo)
        Dim ports As New List(Of SerialPortInfo)()

        Try
            ' 1. Use WMI to get detailed device info
            Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption LIKE '%(COM%)'")
                For Each mObj As ManagementObject In searcher.Get()
                    Dim caption = mObj("Caption")?.ToString()
                    If Not String.IsNullOrEmpty(caption) Then
                        ' Extract COM port name from caption like "Silicon Labs CP210x USB to UART Bridge (COM3)"
                        Dim match = System.Text.RegularExpressions.Regex.Match(caption, "\((COM\d+)\)")
                        If match.Success Then
                            Dim portName = match.Groups(1).Value
                            Dim desc = caption ' Keep original for display
                            
                            ' Flag as USB if it contains typical USB serial keywords
                            Dim lowerDesc = desc.ToLower()
                            Dim isUsb = lowerDesc.Contains("usb") OrElse lowerDesc.Contains("silicon labs") OrElse 
                                        lowerDesc.Contains("prolific") OrElse lowerDesc.Contains("ch340") OrElse
                                        lowerDesc.Contains("ftdi") OrElse lowerDesc.Contains("micro")

                            ports.Add(New SerialPortInfo With {
                                .PortName = portName,
                                .Description = desc,
                                .IsUsb = isUsb
                            })
                        End If
                    End If
                Next
            End Using

            ' 2. Fallback to basic port names if WMI returned nothing
            Dim allNames = SerialPort.GetPortNames()
            For Each name In allNames
                If Not ports.Any(Function(p) p.PortName = name) Then
                    ports.Add(New SerialPortInfo With {
                        .PortName = name,
                        .Description = name,
                        .IsUsb = False
                    })
                End If
            Next

            ' 3. Sort: USB/Micro first, then by COM number descending
            Dim sorted = ports.OrderByDescending(Function(p) p.IsUsb) _
                              .ThenByDescending(Function(p) 
                                                  Dim num As Integer
                                                  Integer.TryParse(p.PortName.Replace("COM", ""), num)
                                                  Return num
                                                End Function) _
                              .ToList()

            Return sorted
        Catch ex As Exception
            Console.WriteLine("[CommManager] Error listing ports: " & ex.Message)
            Return SerialPort.GetPortNames().OrderByDescending(Function(s) s) _
                             .Select(Function(s) New SerialPortInfo With {.PortName = s, .Description = s}) _
                             .ToList()
        End Try
    End Function

    ''' <summary>
    ''' Helper to determine if a tag matches the configured EPC prefix filter.
    ''' </summary>
    Private Shared Function IsValidTagPrefix(epc As String) As Boolean
        If String.IsNullOrEmpty(epc) Then Return False

        If ProjectSettings.Current.RfidFilterMode = "All" Then Return True

        Dim prefix = ProjectSettings.Current.RfidPrefix
        If String.IsNullOrEmpty(prefix) Then Return True ' If no prefix defined, allow all
        
        ' Check if it starts with the configured prefix (ignoring case/spaces)
        Dim cleanEpc = epc.Replace(" ", "").ToUpper().Trim()
        Dim cleanPrefix = prefix.Replace(" ", "").ToUpper().Trim()
        
        Return cleanEpc.StartsWith(cleanPrefix)
    End Function
End Class

Public Class SerialPortInfo
    Public Property PortName As String
    Public Property Description As String
    Public Property IsUsb As Boolean

    Public Overrides Function ToString() As String
        Return If(String.IsNullOrEmpty(Description), PortName, Description)
    End Function
End Class
