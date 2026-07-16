Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Windows

Public Class BackendManager
    Public Shared Sub StartRfidBackend()
        System.Threading.Tasks.Task.Run(Sub()
            Try
                Dim backendPath As String = ProjectSettings.Current.RfidBackendPath

                ' 1. Check if the path is set and file exists
                If String.IsNullOrWhiteSpace(backendPath) OrElse Not File.Exists(backendPath) Then
                    Return
                End If

                ' 2. Read the database name from settings
                Dim dbName As String = ProjectSettings.Current.Database

                ' 3. Setup the Process
                Dim processInfo As New ProcessStartInfo()
                processInfo.FileName = backendPath
                processInfo.UseShellExecute = False
                processInfo.RedirectStandardInput = True
                processInfo.CreateNoWindow = False
                processInfo.WindowStyle = ProcessWindowStyle.Minimized ' Force it to minimize

                Dim rfidProcess As Process = Process.Start(processInfo)

                If rfidProcess IsNot Nothing Then
                    Dim streamWriter As StreamWriter = rfidProcess.StandardInput
                    streamWriter.AutoFlush = True ' Ensure commands are sent immediately

                    ' Wait for the Node app to fully boot and show Prompt 1 (Server IP)
                    System.Threading.Thread.Sleep(2500)

                    ' --- Auto-detect the system's active LAN IP address ---
                    Dim localIp As String = "127.0.0.1" ' Fallback
                    For Each netInterface In NetworkInterface.GetAllNetworkInterfaces()
                        If netInterface.OperationalStatus = OperationalStatus.Up AndAlso
                           netInterface.NetworkInterfaceType <> NetworkInterfaceType.Loopback Then
                            For Each addrInfo In netInterface.GetIPProperties().UnicastAddresses
                                If addrInfo.Address.AddressFamily = AddressFamily.InterNetwork Then
                                    Dim ip As String = addrInfo.Address.ToString()
                                    ' Pick private LAN ranges: 192.168.x.x, 10.x.x.x, 172.16-31.x.x
                                    If ip.StartsWith("192.168.") OrElse
                                       ip.StartsWith("10.") OrElse
                                       ip.StartsWith("172.") Then
                                        localIp = ip
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                        If localIp <> "127.0.0.1" Then Exit For
                    Next

                    ' --- Prompt 1: SERVER_HOST → Auto-detected system IP ---
                    streamWriter.WriteLine(localIp)
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 2: PORT [3000] → <Enter> (accept default) ---
                    streamWriter.WriteLine("")
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 3: DB_HOST [localhost] → <Enter> (accept default) ---
                    streamWriter.WriteLine("")
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 4: DB_PORT [3306] → <Enter> (accept default) ---
                    streamWriter.WriteLine("")
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 5: DB_NAME [hondadb] → send actual DB name from settings ---
                    streamWriter.WriteLine(dbName)
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 6: DB_USER [root] → <Enter> (accept default) ---
                    streamWriter.WriteLine("")
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 7: DB_PASSWORD [] → <Enter> (accept default) ---
                    streamWriter.WriteLine("")
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 8: DB_ENCRYPT [false] → <Enter> (accept default) ---
                    streamWriter.WriteLine("")
                    System.Threading.Thread.Sleep(500)
                    
                    ' --- Prompt 9: DB_TRUST_SERVER_CERTIFICATE [true] → <Enter> (accept default) ---
                    streamWriter.WriteLine("")
                    System.Threading.Thread.Sleep(500)

                    ' --- Prompt 10: Save to .env? (y/n) [y] → y ---
                    streamWriter.WriteLine("y")

                    ' Leave streamWriter open so the process doesn't close
                End If

            Catch ex As Exception
                Application.Current.Dispatcher.Invoke(Sub()
                    MessageBox.Show("Failed to start Node.js backend: " & ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                End Sub)
            End Try
        End Sub)
    End Sub
End Class
