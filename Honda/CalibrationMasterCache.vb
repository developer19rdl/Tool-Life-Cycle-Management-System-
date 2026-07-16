Imports System.IO
Imports System.Text.Json

Public Class CalibrationMasterCache
    Private Shared ReadOnly CacheFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "calibration_masters_cache.json")
    Private Shared ReadOnly LockObject As New Object()

    Public Class CacheEntry
        Public Property Masters As List(Of String)
        Public Property ObservationMasters As List(Of String)
        Public Property GeometricMasters As List(Of String)
    End Class

    Public Shared Sub SaveMasters(controlNo As String, cycleName As String, masters As List(Of MasterSelectorItem))
        SyncLock LockObject
            Try
                Dim cache = LoadAll()
                Dim key = $"{controlNo}_{cycleName}"
                If Not cache.ContainsKey(key) Then cache(key) = New CacheEntry()
                cache(key).Masters = masters.Select(Function(m) m.Description).ToList()
                SaveAll(cache)
            Catch ex As Exception
            End Try
        End SyncLock
    End Sub

    Public Shared Sub SaveMastersDual(controlNo As String, cycleName As String, obsMasters As List(Of MasterSelectorItem), geoMasters As List(Of MasterSelectorItem))
        SyncLock LockObject
            Try
                Dim cache = LoadAll()
                Dim key = $"{controlNo}_{cycleName}"
                If Not cache.ContainsKey(key) Then cache(key) = New CacheEntry()
                cache(key).ObservationMasters = obsMasters.Select(Function(m) m.Description).ToList()
                cache(key).GeometricMasters = geoMasters.Select(Function(m) m.Description).ToList()
                SaveAll(cache)
            Catch ex As Exception
            End Try
        End SyncLock
    End Sub

    Public Shared Function LoadMasters(controlNo As String, cycleName As String) As List(Of String)
        SyncLock LockObject
            Try
                Dim cache = LoadAll()
                Dim key = $"{controlNo}_{cycleName}"
                If cache.ContainsKey(key) AndAlso cache(key).Masters IsNot Nothing Then
                    Return cache(key).Masters
                End If
            Catch ex As Exception
            End Try
            Return New List(Of String)()
        End SyncLock
    End Function

    Public Shared Function LoadMastersDual(controlNo As String, cycleName As String) As (Observation As List(Of String), Geometric As List(Of String))
        SyncLock LockObject
            Try
                Dim cache = LoadAll()
                Dim key = $"{controlNo}_{cycleName}"
                If cache.ContainsKey(key) Then
                    Return (If(cache(key).ObservationMasters, New List(Of String)()),
                            If(cache(key).GeometricMasters, New List(Of String)()))
                End If
            Catch ex As Exception
            End Try
            Return (New List(Of String)(), New List(Of String)())
        End SyncLock
    End Function

    Private Shared Function LoadAll() As Dictionary(Of String, CacheEntry)
        Try
            If File.Exists(CacheFilePath) Then
                Dim json = File.ReadAllText(CacheFilePath)
                If Not String.IsNullOrWhiteSpace(json) Then
                    Return JsonSerializer.Deserialize(Of Dictionary(Of String, CacheEntry))(json)
                End If
            End If
        Catch ex As Exception
        End Try
        Return New Dictionary(Of String, CacheEntry)()
    End Function

    Private Shared Sub SaveAll(cache As Dictionary(Of String, CacheEntry))
        Try
            Dim json = JsonSerializer.Serialize(cache, New JsonSerializerOptions With {.WriteIndented = True})
            File.WriteAllText(CacheFilePath, json)
        Catch ex As Exception
        End Try
    End Sub
End Class
