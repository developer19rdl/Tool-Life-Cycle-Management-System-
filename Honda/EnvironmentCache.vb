Imports System.IO
Imports System.Text.Json

Public Class EnvironmentCache
    Private Shared ReadOnly CacheFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "environment_cache.json")
    Private Shared ReadOnly LockObject As New Object()

    Public Class EnvEntry
        Public Property ClassName As String
        Public Property CalibrationDate As String
        Public Property Temperature As String
        Public Property Humidity As String
        Public Property DepthError As String
        Public Property TimeIn As String
    End Class

    Public Shared Sub SaveEnvironment(className As String, calibDate As String, temp As String, humidity As String, timeIn As String, Optional depthError As String = "")
        SyncLock LockObject
            Try
                Dim cache = LoadAll()
                Dim key = $"{className}_{calibDate}"
                If Not cache.ContainsKey(key) Then 
                    cache(key) = New EnvEntry() With {.ClassName = className, .CalibrationDate = calibDate}
                End If
                cache(key).Temperature = temp
                cache(key).Humidity = humidity
                cache(key).DepthError = depthError
                cache(key).TimeIn = timeIn
                SaveAll(cache)
            Catch ex As Exception
            End Try
        End SyncLock
    End Sub

    Public Shared Function LoadEnvironment(className As String, calibDate As String) As EnvEntry
        SyncLock LockObject
            Try
                Dim cache = LoadAll()
                Dim key = $"{className}_{calibDate}"
                If cache.ContainsKey(key) Then
                    Return cache(key)
                End If
            Catch ex As Exception
            End Try
            Return Nothing
        End SyncLock
    End Function

    Private Shared Function LoadAll() As Dictionary(Of String, EnvEntry)
        Try
            If File.Exists(CacheFilePath) Then
                Dim json = File.ReadAllText(CacheFilePath)
                If Not String.IsNullOrWhiteSpace(json) Then
                    Return JsonSerializer.Deserialize(Of Dictionary(Of String, EnvEntry))(json)
                End If
            End If
        Catch ex As Exception
        End Try
        Return New Dictionary(Of String, EnvEntry)()
    End Function

    Private Shared Sub SaveAll(cache As Dictionary(Of String, EnvEntry))
        Try
            Dim json = JsonSerializer.Serialize(cache, New JsonSerializerOptions With {.WriteIndented = True})
            File.WriteAllText(CacheFilePath, json)
        Catch ex As Exception
        End Try
    End Sub
End Class
