Imports System.Windows
Imports System.Linq
Imports System.Net.NetworkInformation
Imports System.Net.Sockets

Public Class RFIDScannerSelectionWindow
    Public Sub New()
        InitializeComponent()
        LoadPorts()
        LoadCurrentSettings()
    End Sub

    Private Sub LoadPorts()
        ComboPorts.Items.Clear()
        
        Dim ports = CommManager.GetUSBSerialPorts()
        For Each p In ports
            ComboPorts.Items.Add(p)
        Next
    End Sub

    Private Sub LoadCurrentSettings()
        Dim currentType = ProjectSettings.Current.RfidScannerType
        If currentType = "RDL" Then
            RbRDL.IsChecked = True
        Else
            RbThinkMagic.IsChecked = True
        End If

        Dim filterMode = ProjectSettings.Current.RfidFilterMode
        If filterMode = "All" Then
            RbFilterAll.IsChecked = True
        Else
            RbFilterPrefix.IsChecked = True
        End If

        TxtPrefix.Text = ProjectSettings.Current.RfidPrefix
        SldPower.Value = ProjectSettings.Current.RfidPower
        LblPowerValue.Text = $"{ProjectSettings.Current.RfidPower} dBm"
        
        ' Detect system LAN IP and show it at the bottom
        Dim systemIp As String = GetLocalIPAddress()
        LblSystemIp.Text = systemIp

        ' Always pre-fill IP field with the current system LAN IP
        TxtGunIp.Text = If(systemIp <> "Not Found" AndAlso Not systemIp.Contains("Error"), systemIp, "")

        TxtGunPort.Text = ProjectSettings.Current.RfidGunPort
        TxtBackendPath.Text = ProjectSettings.Current.RfidBackendPath
        
        Dim savedPort = ProjectSettings.Current.RfidPort
        If String.IsNullOrEmpty(savedPort) Then
            ' Smart Pre-selection: Prioritize USB/Micro devices
            Dim firstUsb = ComboPorts.Items.OfType(Of SerialPortInfo)().FirstOrDefault(Function(p) p.IsUsb)
            If firstUsb IsNot Nothing Then
                ComboPorts.SelectedItem = firstUsb
            ElseIf ComboPorts.Items.Count > 0 Then
                ComboPorts.SelectedIndex = 0
            End If
        Else
            Dim found = ComboPorts.Items.OfType(Of SerialPortInfo)().FirstOrDefault(Function(p) p.PortName = savedPort)
            If found IsNot Nothing Then
                ComboPorts.SelectedItem = found
            Else
                ComboPorts.Items.Add(savedPort)
                ComboPorts.SelectedItem = savedPort
            End If
        End If
        
        
        ' Dynamic update for slider
        AddHandler SldPower.ValueChanged, Sub(s, e) LblPowerValue.Text = $"{CInt(SldPower.Value)} dBm"
    End Sub

    Private Sub RbFilterMode_Changed(sender As Object, e As RoutedEventArgs)
        If TxtPrefix IsNot Nothing AndAlso RbFilterPrefix IsNot Nothing Then
            TxtPrefix.IsEnabled = (RbFilterPrefix.IsChecked = True)
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub BtnRefreshPorts_Click(sender As Object, e As RoutedEventArgs)
        LoadPorts()
        ' Auto-select the best port after refresh
        Dim firstUsb = ComboPorts.Items.OfType(Of SerialPortInfo)().FirstOrDefault(Function(p) p.IsUsb)
        If firstUsb IsNot Nothing Then
            ComboPorts.SelectedItem = firstUsb
        ElseIf ComboPorts.Items.Count > 0 Then
            ComboPorts.SelectedIndex = 0
        End If
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        If RbRDL.IsChecked = True Then
            ProjectSettings.Current.RfidScannerType = "RDL"
        Else
            ProjectSettings.Current.RfidScannerType = "ThinkMagic"
        End If

        If RbFilterAll.IsChecked = True Then
            ProjectSettings.Current.RfidFilterMode = "All"
        Else
            ProjectSettings.Current.RfidFilterMode = "Prefix"
        End If

        ProjectSettings.Current.RfidPrefix = TxtPrefix.Text.Trim()
        ProjectSettings.Current.RfidPower = CInt(SldPower.Value)
        ProjectSettings.Current.RfidGunIp = TxtGunIp.Text.Trim()
        ProjectSettings.Current.RfidGunPort = TxtGunPort.Text.Trim()
        ProjectSettings.Current.RfidBackendPath = TxtBackendPath.Text.Trim()

        If ComboPorts.SelectedItem Is Nothing Then
            ProjectSettings.Current.RfidPort = ""
        Else
            Dim selected = TryCast(ComboPorts.SelectedItem, SerialPortInfo)
            If selected IsNot Nothing Then
                ProjectSettings.Current.RfidPort = selected.PortName
            Else
                ProjectSettings.Current.RfidPort = ComboPorts.SelectedItem.ToString()
            End If
        End If

        ProjectSettings.Save()
        MessageBox.Show("RFID Scanner configuration saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        Me.Close()
    End Sub

    Private Async Sub BtnTestConnection_Click(sender As Object, e As RoutedEventArgs)
        Dim ip As String = TxtGunIp.Text.Trim()
        Dim portStr As String = TxtGunPort.Text.Trim()
        
        If String.IsNullOrWhiteSpace(ip) OrElse String.IsNullOrWhiteSpace(portStr) Then
            MessageBox.Show("Please enter both IP Address and Port.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        BtnTestConnection.IsEnabled = False
        BtnTestConnection.Content = "Testing..."

        Try
            Dim url = $"http://{ip}:{portStr}/api/session/stream"
            Using client As New System.Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(5)
                Using request = New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url)
                    request.Headers.Add("Accept", "text/event-stream")
                    Using response = Await client.SendAsync(request, System.Net.Http.HttpCompletionOption.ResponseHeadersRead)
                        If response.IsSuccessStatusCode Then
                            MessageBox.Show("Connection successful! Stream is reachable.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                        Else
                            MessageBox.Show($"Connection reachable but returned error: {response.StatusCode}", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As System.Net.Http.HttpRequestException
            Dim troubleshootingMsg = "Could not connect to the RFID Gun." & vbCrLf & vbCrLf &
                                   "Please check the following:" & vbCrLf &
                                   "1. Is the RFID Gun turned ON?" & vbCrLf &
                                   "2. Is it connected to the same Network/WiFi?" & vbCrLf &
                                   "3. Are the IP Address and Port correct?"
            
            MessageBox.Show(troubleshootingMsg, "Connection Failed", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As System.Threading.Tasks.TaskCanceledException
            MessageBox.Show("Connection timed out. The device is not responding. Please ensure the IP is correct and the device is awake.", "Timeout", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            BtnTestConnection.IsEnabled = True
            BtnTestConnection.Content = "Test Connection"
        End Try
    End Sub

    Private Function GetLocalIPAddress() As String
        Try
            For Each netInterface In NetworkInterface.GetAllNetworkInterfaces()
                If netInterface.OperationalStatus = OperationalStatus.Up AndAlso
                   netInterface.NetworkInterfaceType <> NetworkInterfaceType.Loopback Then
                    For Each addrInfo In netInterface.GetIPProperties().UnicastAddresses
                        If addrInfo.Address.AddressFamily = AddressFamily.InterNetwork Then
                            Dim ip As String = addrInfo.Address.ToString()
                            ' Return first private LAN IP: 192.168.x.x, 10.x.x.x, 172.x.x.x
                            If ip.StartsWith("192.168.") OrElse
                               ip.StartsWith("10.") OrElse
                               ip.StartsWith("172.") Then
                                Return ip
                            End If
                        End If
                    Next
                End If
            Next
            Return "Not Found"
        Catch ex As Exception
            Return "Error fetching IP"
        End Try
    End Function

    Private Sub BtnBrowseBackend_Click(sender As Object, e As RoutedEventArgs)
        Dim dialog As New Microsoft.Win32.OpenFileDialog()
        dialog.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
        dialog.Title = "Select RFID Backend Executable"
        
        If dialog.ShowDialog() = True Then
            TxtBackendPath.Text = dialog.FileName
        End If
    End Sub

    Private Sub BtnStartBackend_Click(sender As Object, e As RoutedEventArgs)
        ' Save current path first so the latest input is used
        ProjectSettings.Current.RfidBackendPath = TxtBackendPath.Text.Trim()
        ProjectSettings.Save()

        ' Trigger the backend manager
        BackendManager.StartRfidBackend()
    End Sub
End Class
