Imports System.Data
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.Json
Imports MySql.Data.MySqlClient

''' <summary>Metadata written into every .hbak archive.</summary>
Public Class BackupManifest
    Public Property Version As String = "1.0"
    Public Property AppVersion As String = "1.3.9"
    Public Property Timestamp As String = ""
    Public Property DatabaseName As String = ""
    Public Property TableCount As Integer = 0
    Public Property FileCount As Integer = 0
End Class

''' <summary>
''' Zero-dependency backup/restore engine.
''' All DB work uses MySql.Data (already referenced).
''' All compression uses System.IO.Compression (built into .NET 8).
''' </summary>
Public Class BackupManager

    Public Shared Event ProgressChanged(message As String)
    Private Shared Sub Report(message As String)
        RaiseEvent ProgressChanged(message)
    End Sub

    ' ── BACKUP ───────────────────────────────────────────────────────────────

    Public Shared Async Function CreateBackupAsync(outputPath As String) As Task(Of Boolean)
        Dim tempDir = Path.Combine(Path.GetTempPath(), "TLCMS_Backup_" & DateTime.Now.ToString("yyyyMMddHHmmss"))
        Try
            Directory.CreateDirectory(tempDir)
            Report("📁 Preparing temporary workspace...")

            Report("🗄️  Exporting database tables...")
            Dim tableCount As Integer = 0
            Dim sql As String = ""
            Dim dumpOk As Boolean = False
            Await Task.Run(Sub()
                               Try
                                   sql = DumpDatabaseToSql(tableCount)
                                   dumpOk = Not String.IsNullOrEmpty(sql)
                               Catch ex As Exception
                                   Report("❌ DB export error: " & ex.Message)
                               End Try
                           End Sub)

            If Not dumpOk Then
                Report("❌ Database export failed. Backup aborted.")
                Return False
            End If
            File.WriteAllText(Path.Combine(tempDir, "database_dump.sql"), sql, Encoding.UTF8)
            Report($"✅ {tableCount} tables exported.")

            Dim fileCount As Integer = 0
            Dim storageRoot = ProjectSettings.GetFileStorageRoot()
            If Directory.Exists(storageRoot) Then
                Report("📄 Copying uploaded files...")
                Dim filesDestDir = Path.Combine(tempDir, "files")
                Directory.CreateDirectory(filesDestDir)
                Await Task.Run(Sub() fileCount = CopyDirRecursive(storageRoot, filesDestDir))
                Report($"✅ {fileCount} file(s) copied.")
            Else
                Report("ℹ️  No file storage folder found — skipping files.")
            End If

            Dim manifest As New BackupManifest With {
                .Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                .DatabaseName = ProjectSettings.Current.Database,
                .TableCount = tableCount,
                .FileCount = fileCount
            }
            File.WriteAllText(Path.Combine(tempDir, "backup_manifest.json"),
                JsonSerializer.Serialize(manifest, New JsonSerializerOptions With {.WriteIndented = True}))

            Report("🗜️  Compressing archive...")
            If File.Exists(outputPath) Then File.Delete(outputPath)
            Await Task.Run(Sub() ZipFile.CreateFromDirectory(tempDir, outputPath))

            ProjectSettings.Current.LastBackupTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            ProjectSettings.Save()

            Dim sizeMB = (New FileInfo(outputPath).Length / 1048576.0).ToString("F2")
            Report($"🎉 Backup complete!  Size: {sizeMB} MB")
            Report($"📍 File: {outputPath}")
            Return True

        Catch ex As Exception
            Report($"❌ Backup error: {ex.Message}")
            Return False
        Finally
            Try
                If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
            Catch
            End Try
        End Try
    End Function

    Private Shared Function DumpDatabaseToSql(ByRef tableCount As Integer) As String
        Dim sb As New StringBuilder()
        tableCount = 0
        Try
            Dim mysql As New MySQLClass()
            If mysql.MySQLDBConnect() <> 1 Then Return ""

            sb.AppendLine("-- ================================================")
            sb.AppendLine("-- TLCMS Database Backup")
            sb.AppendLine($"-- Generated : {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine($"-- Database  : {ProjectSettings.Current.Database}")
            sb.AppendLine("-- ================================================")
            sb.AppendLine()
            sb.AppendLine("SET FOREIGN_KEY_CHECKS=0;")
            sb.AppendLine("SET SQL_MODE='NO_AUTO_VALUE_ON_ZERO';")
            sb.AppendLine("SET NAMES utf8mb4;")
            sb.AppendLine()

            Dim tables = mysql.GetAllTableNames()
            tableCount = tables.Count

            For Each tblName In tables
                sb.AppendLine($"-- ── Table: `{tblName}` ──")
                sb.AppendLine($"DROP TABLE IF EXISTS `{tblName}`;")

                Dim createDt = mysql.ReadDatatable($"SHOW CREATE TABLE `{tblName}`")
                If createDt.Rows.Count > 0 Then
                    sb.AppendLine(createDt.Rows(0)(1).ToString() & ";")
                End If
                sb.AppendLine()

                Dim dataDt = mysql.ReadDatatable($"SELECT * FROM `{tblName}`")
                If dataDt.Rows.Count > 0 Then
                    Dim cols As New List(Of String)
                    For Each col As DataColumn In dataDt.Columns
                        cols.Add($"`{col.ColumnName}`")
                    Next
                    Dim colList = String.Join(", ", cols)

                    For Each row As DataRow In dataDt.Rows
                        Dim vals As New List(Of String)
                        For Each col As DataColumn In dataDt.Columns
                            vals.Add(FormatSqlValue(row(col)))
                        Next
                        sb.AppendLine($"INSERT INTO `{tblName}` ({colList}) VALUES ({String.Join(", ", vals)});")
                    Next
                End If
                sb.AppendLine()
            Next

            sb.AppendLine("SET FOREIGN_KEY_CHECKS=1;")
            Return sb.ToString()
        Catch ex As Exception
            Console.WriteLine("DumpDatabaseToSql Error: " & ex.Message)
            Return ""
        End Try
    End Function

    Private Shared Function FormatSqlValue(val As Object) As String
        If IsDBNull(val) OrElse val Is Nothing Then Return "NULL"
        If TypeOf val Is Boolean Then Return If(CBool(val), "1", "0")
        If TypeOf val Is Byte OrElse TypeOf val Is Short OrElse TypeOf val Is Integer OrElse
           TypeOf val Is Long OrElse TypeOf val Is Single OrElse TypeOf val Is Double OrElse
           TypeOf val Is Decimal Then Return val.ToString()
        If TypeOf val Is DateTime Then
            Dim dt = CDate(val)
            If dt = DateTime.MinValue Then Return "NULL"
            Return $"'{dt:yyyy-MM-dd HH:mm:ss}'"
        End If
        Dim s = val.ToString()
        s = s.Replace("\", "\\")
        s = s.Replace("'", "\'")
        s = s.Replace(vbCrLf, "\n")
        s = s.Replace(vbCr, "\n")
        s = s.Replace(vbLf, "\n")
        Return $"'{s}'"
    End Function

    Private Shared Function CopyDirRecursive(src As String, dest As String) As Integer
        Dim count As Integer = 0
        For Each f In Directory.GetFiles(src)
            File.Copy(f, Path.Combine(dest, Path.GetFileName(f)), True)
            count += 1
        Next
        For Each d In Directory.GetDirectories(src)
            Dim destSub = Path.Combine(dest, Path.GetFileName(d))
            Directory.CreateDirectory(destSub)
            count += CopyDirRecursive(d, destSub)
        Next
        Return count
    End Function

    ' ── RESTORE ──────────────────────────────────────────────────────────────

    Public Shared Async Function RestoreBackupAsync(backupPath As String) As Task(Of Boolean)
        Dim tempDir = Path.Combine(Path.GetTempPath(), "TLCMS_Restore_" & DateTime.Now.ToString("yyyyMMddHHmmss"))
        Try
            Report("📦 Reading backup archive...")
            If Not File.Exists(backupPath) Then
                Report("❌ Backup file not found.")
                Return False
            End If

            Await Task.Run(Sub() ZipFile.ExtractToDirectory(backupPath, tempDir, True))

            Dim manifestPath = Path.Combine(tempDir, "backup_manifest.json")
            If Not File.Exists(manifestPath) Then
                Report("❌ Invalid backup — manifest missing. File may be corrupted.")
                Return False
            End If

            Dim manifest = JsonSerializer.Deserialize(Of BackupManifest)(File.ReadAllText(manifestPath))
            Report($"📋 Backup date  : {manifest.Timestamp}")
            Report($"🗄️  Tables       : {manifest.TableCount}")
            Report($"📄 Files        : {manifest.FileCount}")
            Report("")

            Dim sqlPath = Path.Combine(tempDir, "database_dump.sql")
            If File.Exists(sqlPath) Then
                Report("🗄️  Restoring database...")
                Dim sql = File.ReadAllText(sqlPath, Encoding.UTF8)
                Dim dbOk As Boolean = False
                Await Task.Run(Sub()
                                   Try
                                       Dim mysql As New MySQLClass()
                                       dbOk = mysql.ExecuteSqlScript(sql)
                                   Catch ex As Exception
                                       Report("⚠️  DB error: " & ex.Message)
                                   End Try
                               End Sub)
                Report(If(dbOk, "✅ Database restored.", "⚠️  Database restore had errors."))
            End If

            Dim filesDir = Path.Combine(tempDir, "files")
            If Directory.Exists(filesDir) Then
                Report("📄 Restoring uploaded files...")
                Dim storageRoot = ProjectSettings.GetFileStorageRoot()
                Directory.CreateDirectory(storageRoot)
                Dim fileCount As Integer = 0
                Await Task.Run(Sub() fileCount = CopyDirRecursive(filesDir, storageRoot))
                Report($"✅ {fileCount} file(s) restored.")
            Else
                Report("ℹ️  No files in archive — skipping.")
            End If

            ProjectSettings.Current.LastBackupTimestamp =
                "Restored: " & manifest.Timestamp
            ProjectSettings.Save()

            Report("")
            Report("🎉 Restore complete!")

            ' ── Notify user about file storage location ───────────────────────
            If String.IsNullOrWhiteSpace(ProjectSettings.Current.FileStorageBasePath) Then
                Dim documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                Dim defaultPath = Path.Combine(documents, "Tool Life Cycle Management System Files")
                Report("")
                Report("📁 No custom file storage path is configured.")
                Report($"   Files have been restored to the default location:")
                Report($"   {defaultPath}")
                Report("   To change this, go to Settings → Admin → File Storage Location.")
            Else
                Report($"📁 Files restored to: {ProjectSettings.GetFileStorageRoot()}")
            End If

            Report("🔄 Please restart the application.")
            Return True

        Catch ex As Exception
            Report($"❌ Restore error: {ex.Message}")
            Return False
        Finally
            Try
                If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
            Catch
            End Try
        End Try
    End Function

End Class
