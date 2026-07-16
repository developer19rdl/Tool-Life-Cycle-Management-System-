Imports System.IO
Imports System.Text.Json

Public Class DbSettings
    Public Property Datasource As String = "localhost"
    Public Property Port As String = "3306"
    Public Property Database As String = ""
    Public Property Username As String = "root"
    Public Property Password As String = ""
    Public Property RfidScannerType As String = "ThinkMagic" ' Default
    Public Property RfidFilterMode As String = "All" ' Default filter mode (All or Prefix)
    Public Property RfidPrefix As String = "E" ' Default EPC prefix filter
    Public Property RfidPower As Integer = 20 ' Default power in dBm
    Public Property RfidGunIp As String = "192.168.1.51" ' Default IP for RFID Gun
    Public Property RfidGunPort As String = "3000" ' Default Port for RFID Gun
    Public Property RfidPort As String = "" ' Manual override for USB COM port
    Public Property RfidBackendPath As String = "" ' Path to the Node.js backend
    Public Property RfidServerIp As String = "192.168.1.3" ' IP the Node.js RFID server binds to
    Public Property RfidServerPort As String = "3000" ' Port the Node.js RFID server listens on
    ''' <summary>
    ''' The user-selected base directory for all uploaded/generated files.
    ''' When empty the application directory is used as the fallback.
    ''' </summary>
    Public Property FileStorageBasePath As String = ""
    ''' <summary>Timestamp of the last successful backup (display only).</summary>
    Public Property LastBackupTimestamp As String = ""
End Class

Public Class ProjectSettings
    Private Shared _instance As DbSettings
    Private Shared ReadOnly SettingsPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json")

    Public Shared ReadOnly Property Current As DbSettings
        Get
            If _instance Is Nothing Then
                Load()
            End If
            Return _instance
        End Get
    End Property

    Public Shared Sub Load()
        Try
            If File.Exists(SettingsPath) Then
                Dim json As String = File.ReadAllText(SettingsPath)
                _instance = JsonSerializer.Deserialize(Of DbSettings)(json)
            Else
                _instance = New DbSettings()
            End If
        Catch ex As Exception
            _instance = New DbSettings()
        End Try
    End Sub

    Public Shared Sub Save()
        Try
            Dim json As String = JsonSerializer.Serialize(_instance, New JsonSerializerOptions With {.WriteIndented = True})
            File.WriteAllText(SettingsPath, json)
        Catch ex As Exception
            ' Handle error
        End Try
    End Sub

    ''' <summary>
    ''' Returns the root folder for file storage.
    ''' If the user has configured a custom path the function returns:
    '''   <configured base>\Tool Life Cycle Management System Files
    ''' Otherwise it falls back to the user's Documents folder:
    '''   Documents\Tool Life Cycle Management System Files
    ''' </summary>
    Public Shared Function GetFileStorageRoot() As String
        Dim basePath = Current.FileStorageBasePath
        If Not String.IsNullOrWhiteSpace(basePath) AndAlso Directory.Exists(basePath) Then
            Return Path.Combine(basePath, "Tool Life Cycle Management System Files")
        End If
        ' Fallback: use the user's Documents folder
        Dim documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Return Path.Combine(documents, "Tool Life Cycle Management System Files")
    End Function

    ''' <summary>
    ''' Creates the full folder structure under the storage root.
    ''' Call this once after the user saves a new path.
    ''' </summary>
    Public Shared Sub EnsureFileStorageFolders()
        Dim root = GetFileStorageRoot()

        ' Create all required section sub-folders
        Dim paths As New List(Of String)
        paths.Add(Path.Combine(root, "Database", "Instrument"))
        paths.Add(Path.Combine(root, "Database", "Gauge"))
        paths.Add(Path.Combine(root, "Interchangeability", "Reintroduction"))
        paths.Add(Path.Combine(root, "Interchangeability", "WriteOff"))
        paths.Add(Path.Combine(root, "Interchangeability", "TempIssuance"))
        paths.Add(Path.Combine(root, "Calibration", "Reports"))
        paths.Add(Path.Combine(root, "Calibration", "Masters"))
        paths.Add(Path.Combine(root, "Inventory", "Documents"))
        paths.Add(Path.Combine(root, "Records", "Exports"))

        For Each fullPath In paths
            If Not Directory.Exists(fullPath) Then
                Directory.CreateDirectory(fullPath)
            End If
        Next
    End Sub
End Class
