Imports System.Data
Imports MySql.Data.MySqlClient
Imports System.Text
Imports System.IO
Imports System.Linq

Public Class SpecialDepthDynParam
    Public Property ParamPrefix As String ' "Dia_1", "Distance_1", "Angle_1"
    Public Property Nominal As String
    Public Property PermErr As String
    Public Property MinLimit As String
    Public Property MaxLimit As String
    Public Property Obs1 As String
    Public Property Obs2 As String
    Public Property Obs3 As String
End Class

Public Class MySQLClass
    Private _con As New MySqlConnection()
    Private _dbLock As New Object() ' Synchronization object for thread safety
    Public Property LastError As String = ""
    Public Property CurrentTargetTrackingTable As String = "regular_calibration"

    Public Sub DBConnect()
        SyncLock _dbLock
            Try
                Dim settings = ProjectSettings.Current
                _con.ConnectionString = $"datasource={settings.Datasource};port={settings.Port};database={settings.Database};username={settings.Username};password={settings.Password};"

                If _con.State <> ConnectionState.Open Then
                    _con.Open()
                    InitializeInterchangeTables()
                End If
            Catch ex As Exception
                Console.WriteLine("Error connecting to database: " & ex.Message)
                Throw
            End Try
        End SyncLock
    End Sub

    Public Function MySQLDBConnect() As Integer
        Try
            If _con IsNot Nothing AndAlso _con.State = ConnectionState.Open Then
                Return 1
            End If

            DBConnect()

            If _con.State = ConnectionState.Open Then
                Return 1
            End If
        Catch ex As Exception
            Console.WriteLine("Error connecting to database: " & ex.Message)
        End Try
        Return 0
    End Function

    ' Basic test function to verify connectivity
    Public Function TestConnection(datasource As String, port As String, database As String, username As String, password As String) As Boolean
        Dim testCon As New MySqlConnection($"datasource={datasource};port={port};database={database};username={username};password={password};")
        Try
            testCon.Open()
            testCon.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function ReadDatatable(ByVal Query As String) As DataTable
        SyncLock _dbLock
            Try
                Dim dataTable As New DataTable()
                If MySQLDBConnect() = 1 Then
                    Using command As New MySqlCommand(Query, _con)
                        Dim adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End If
                Return dataTable
            Catch ex As Exception
                Console.WriteLine("MySQL Error reading data table: " & ex.Message)
                Return New DataTable()
            End Try
        End SyncLock
    End Function

    Public Function ReadDatatable(ByVal query As String, ByVal parameters As MySqlParameter()) As DataTable
        SyncLock _dbLock
            Try
                Dim dataTable As New DataTable()
                If MySQLDBConnect() = 1 Then
                    Using command As New MySqlCommand(query, _con)
                        If parameters IsNot Nothing Then
                            command.Parameters.AddRange(parameters)
                        End If
                        Dim adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End If
                Return dataTable
            Catch ex As Exception
                Console.WriteLine("MySQL Error reading data table: " & ex.Message)
                Return New DataTable()
            End Try
        End SyncLock
    End Function

    Public Function ExecuteScalar(ByVal query As String, ByVal parameters As MySqlParameter()) As Object
        SyncLock _dbLock
            Try
                If MySQLDBConnect() = 1 Then
                    Using command As New MySqlCommand(query, _con)
                        If parameters IsNot Nothing Then
                            command.Parameters.AddRange(parameters)
                        End If
                        Return command.ExecuteScalar()
                    End Using
                End If
            Catch ex As Exception
                Console.WriteLine("MySQL Error executing scalar: " & ex.Message)
            End Try
            Return Nothing
        End SyncLock
    End Function

    Public Function ExecuteQuery(ByVal query As String) As Boolean
        SyncLock _dbLock
            Try
                If MySQLDBConnect() = 1 Then
                    Using cmd As New MySqlCommand(query, _con)
                        cmd.ExecuteNonQuery()
                        Return True
                    End Using
                End If
            Catch ex As Exception
                Console.WriteLine("ExecuteQuery Error: " & ex.Message)
            End Try
            Return False
        End SyncLock
    End Function

    Public Function InsertNewUser(username As String, password As String, email As String, phoneno As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO users (Username, Password, Email, Phone) VALUES (@usernm, @passrd, @email, @phone)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@usernm", username)
                    cmd.Parameters.AddWithValue("@passrd", password)
                    cmd.Parameters.AddWithValue("@email", email)
                    cmd.Parameters.AddWithValue("@phone", phoneno)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertNewUser Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateUser(ByVal id As Integer, ByVal username As String, ByVal password As String, ByVal email As String, ByVal phone As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE users SET Username=@username, Password=@password, Email=@email, Phone=@phone WHERE ID=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@username", username)
                    cmd.Parameters.AddWithValue("@password", password)
                    cmd.Parameters.AddWithValue("@email", email)
                    cmd.Parameters.AddWithValue("@phone", phone)
                    cmd.Parameters.AddWithValue("@id", id)
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    Return (rowsAffected > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateUser Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function DeleteUser(ByVal id As Integer) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "DELETE FROM users WHERE ID=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@id", id)
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    Return (rowsAffected > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("DeleteUser Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function CheckLogin(username As String, password As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT COUNT(*) FROM users WHERE Username=@u AND Password=@p"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@u", username)
                    cmd.Parameters.AddWithValue("@p", password)
                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("CheckLogin Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function ColumnExists(tableName As String, columnName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = $"SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = '{tableName}' AND column_name = '{columnName}'"
                Dim dt = ReadDatatable(query)
                If dt.Rows.Count > 0 Then
                    Return Convert.ToInt32(dt.Rows(0)(0)) > 0
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error checking column {columnName} in {tableName}: {ex.Message}")
        End Try
        Return False
    End Function

    ' --- DYNAMIC TABLE HELPERS ---

    ''' <summary>
    ''' Converts a type name to a safe MySQL table name.
    ''' e.g. "Plain Plug Gauge" -> "plain_plug_gauge", "V Anvil Micrometer" -> "v_anvil_micrometer"
    ''' 
    ''' IMPORTANT: When forInventory:=True, the calibration-specific hardcodes are skipped so that
    ''' gauge type names like "Plain Plug Gauge" resolve to their inventory table ("plain_plug_gauge")
    ''' instead of the calibration table ("plain_plug_gauge_calibration").
    ''' Always pass forInventory:=True when creating or dropping per-type inventory tables in Settings.
    ''' </summary>
    Public Shared Function TypeNameToTableName(ByVal typeName As String, Optional forInventory As Boolean = False) As String
        If String.IsNullOrWhiteSpace(typeName) Then Return "unknown_type"

        ' --- Calibration-specific hardcodes ---
        ' These map gauge/instrument calibration FORM names to their dedicated calibration DB tables.
        ' Skip this block when resolving names for inventory table create/drop operations (forInventory=True).
        If Not forInventory Then
            ' Vernier Calliper (double-l spelling in system) → vernier_caliper (single-l prefix)
            ' Calibration tables are vernier_caliper_300, vernier_caliper_600 etc. (single l)
            If typeName.IndexOf("Calliper", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
               typeName.IndexOf("Vernier", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "vernier_caliper"
            End If

            If typeName.Equals("HeightMastergaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("HeightMasterGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Height Master Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "height_master_calibration"
            End If

            If typeName.Equals("KeyGrooveGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("KeyGrooveGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Key Groove Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "key_groove_guage_calibration"
            End If

            If typeName.Equals("SnapGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("SnapGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Snap Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "snap_gauge_calibration"
            End If

            If typeName.Equals("SpecialHeightGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("SpecialHeightGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Special Height Gauge", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Special Height", StringComparison.OrdinalIgnoreCase) Then
                Return "special_height_gauge_calibration"
            End If

            If typeName.Equals("PlainPlugGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("PlainPlugGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Plain Plug Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "plain_plug_gauge_calibration"
            End If

            If typeName.Equals("PlainRingGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("PlainRingGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Plain Ring Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "plain_ring_gauge_calibration"
            End If

            If typeName.Equals("SplineRingGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("SplineRingGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Spline Ring Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "spline_ring_gauge_calibration"
            End If

            If typeName.Equals("StraightPlugGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("StraightPlugGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Straight Plug Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "straight_plug_gauge_calibration"
            End If

            If typeName.Equals("ThreadRingGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("ThreadRingGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Thread Ring Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "thread_ring_gauge_calibration"
            End If

            If typeName.Equals("ThreadPlugGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("ThreadPlugGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Thread Plug Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "thread_plug_gauge_calibration"
            End If

            If typeName.Equals("TaperPlugGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("TaperPlugGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Taper Plug Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "taper_plug_gauge_calibration"
            End If

            If typeName.Equals("TaperRingGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("TaperRingGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Taper Ring Gauge", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Plain Taper Ring Gauge", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("PlainTaperRingGaugeCalibration", StringComparison.OrdinalIgnoreCase) Then
                Return "taper_ring_gauge_calibration"
            End If

            If typeName.Equals("FeelerGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("FeelerGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Feeler Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "feeler_gauge_calibration"
            End If

            If typeName.Equals("RadiusGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("RadiusGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Radius Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "radius_gauge_calibration"
            End If

            If typeName.Equals("ThreadPitchGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("ThreadPitchGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Thread Pitch Gauge", StringComparison.OrdinalIgnoreCase) Then
                Return "thread_pitch_gauge_calibration"
            End If

            If typeName.Equals("SpecialDepthGaugeCalibration", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("SpecialDepthGauge_cal", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Special Depth Gauge", StringComparison.OrdinalIgnoreCase) OrElse
               typeName.Equals("Special Depth", StringComparison.OrdinalIgnoreCase) Then
                Return "special_depth_gauge_calibration"
            End If
        End If
        ' --- End calibration hardcodes ---

        Try
            ' 1. Insert underscores between transitions: lower to Upper, alpha to Number, etc.
            ' Handle cases like VernierCaliper300 -> Vernier_Caliper_300
            Dim result As String = System.Text.RegularExpressions.Regex.Replace(typeName.Trim(), "([a-z])([A-Z])", "$1_$2")
            result = System.Text.RegularExpressions.Regex.Replace(result, "([a-zA-Z])([0-9])", "$1_$2")
            result = System.Text.RegularExpressions.Regex.Replace(result, "([0-9])([a-zA-Z])", "$1_$2")

            ' 2. Lowercase and replace any other non-alphanumeric with underscores
            result = System.Text.RegularExpressions.Regex.Replace(result.ToLower(), "[^a-z0-9]+", "_")

            ' 3. Cleanup underscores
            result = result.Trim("_"c)
            While result.Contains("__")
                result = result.Replace("__", "_")
            End While

            If String.IsNullOrEmpty(result) Then result = "unknown_type"
            Return result
        Catch ex As Exception
            ' Fallback to simple logic if regex fails for any reason
            Return typeName.ToLower().Replace(" ", "_").Trim()
        End Try
    End Function

    ''' <summary>
    ''' Searches across all master card tables (instruments and gauges) for a specific Control No.
    ''' This bypasses department_list and gets meta-data directly from the source card tables.
    ''' </summary>
    Public Function GetMasterCardData(ByVal controlNo As String) As DataTable
        Try
            ' 1. Get all defined types from settings
            Dim dtTypes = ReadDatatable("SELECT TypeName, Category FROM type_details")

            ' 2. Search each type's individual table
            For Each row As DataRow In dtTypes.Rows
                Dim typeName = row("TypeName").ToString()
                Dim category = row("Category").ToString()
                Dim tblName = TypeNameToTableName(typeName, forInventory:=True)

                ' Check if table exists before querying
                Dim checkTable = ReadDatatable($"SHOW TABLES LIKE '{tblName}'")
                If checkTable.Rows.Count > 0 Then
                    Dim selectSql As String
                    If category.Equals("Instrument", StringComparison.OrdinalIgnoreCase) Then
                        ' Instruments use InstrumentName
                        selectSql = $"SELECT *, 'Instrument' as MasterCardType, '{typeName.Replace("'", "''")}' as MasterCardTypeName FROM `{tblName}` WHERE ControlNo = @ctrl LIMIT 1"
                    Else
                        ' Gauges use GaugeName
                        selectSql = $"SELECT *, 'Gauge' as MasterCardType, '{typeName.Replace("'", "''")}' as MasterCardTypeName, GaugeName as InstrumentName FROM `{tblName}` WHERE ControlNo = @ctrl LIMIT 1"
                    End If

                    If MySQLDBConnect() = 1 Then
                        Using cmd As New MySqlCommand(selectSql, _con)
                            cmd.Parameters.AddWithValue("@ctrl", controlNo)
                            Using adapter As New MySqlDataAdapter(cmd)
                                Dim dtMatch As New DataTable()
                                adapter.Fill(dtMatch)
                                If dtMatch.Rows.Count > 0 Then
                                    Return dtMatch
                                End If
                            End Using
                        End Using
                    End If
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("GetMasterCardData Error: " & ex.Message)
        End Try
        ' Return empty table if not found
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Creates a per-type table using the instrument or gauge schema template.
    ''' Category should be "Instrument" or "Gauge".
    ''' </summary>
    Public Function CreateTypeTable(ByVal tableName As String, ByVal category As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim schema As String
                If category.Equals("Gauge", StringComparison.OrdinalIgnoreCase) Then
                    schema = $"CREATE TABLE IF NOT EXISTS `{tableName}` (" &
                             "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                             "Date DATE, Time TIME, GaugeName VARCHAR(255), GaugeDescription TEXT, " &
                             "ControlNo VARCHAR(100) UNIQUE, Color VARCHAR(100), DrgNo VARCHAR(100), " &
                             "MakerName VARCHAR(255), Model VARCHAR(255), Line VARCHAR(100), " &
                             "Size VARCHAR(100), Tol VARCHAR(100), Tol2 VARCHAR(100), Tol3 VARCHAR(100), " &
                             "Section VARCHAR(255), Location VARCHAR(255), DateofAddition VARCHAR(50), " &
                             "RequestNo VARCHAR(100), Remark TEXT, CategoryControl VARCHAR(100), " &
                             "RFID_tag VARCHAR(255), uploaded_doc TEXT, Flag INT DEFAULT 0, InstrumentStatus VARCHAR(50) DEFAULT 'Active')"
                Else
                    schema = $"CREATE TABLE IF NOT EXISTS `{tableName}` (" &
                              "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                              "Date DATE, Time TIME, InstrumentName VARCHAR(255), InstrumentDescription TEXT, " &
                              "ControlNo VARCHAR(100) UNIQUE, Color VARCHAR(100), InstrumentModelNo VARCHAR(255), " &
                              "Line VARCHAR(100), AssetNo VARCHAR(100), Size VARCHAR(100), MakerName VARCHAR(255), " &
                              "Section VARCHAR(255), Location VARCHAR(255), Remark TEXT, RequestNo VARCHAR(100), " &
                              "CategoryControl VARCHAR(100), RFID_tag VARCHAR(255), uploaded_doc TEXT, " &
                              "Flag INT DEFAULT 0, InstrumentStatus VARCHAR(50) DEFAULT 'Active')"
                End If
                Using cmd As New MySqlCommand(schema, _con)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("CreateTypeTable Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Creates a per-type table using the Calibration Master schema template.
    ''' </summary>
    Public Function CreateCalibrationMasterTable(ByVal tableName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim schema As String = $"CREATE TABLE IF NOT EXISTS `{tableName}` (" &
                             "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                             "Date DATE, Time TIME, " &
                             "Description TEXT, " &
                             "LeastCount VARCHAR(100), " &
                             "MasterUncertainty DECIMAL(18,10), " &
                             "CalDate DATE, " &
                             "DueDate DATE, " &
                             "uploaded_doc TEXT)"
                Using cmd As New MySqlCommand(schema, _con)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("CreateCalibrationMasterTable Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Drops a per-type table when a type card is deleted from Settings.
    ''' </summary>
    Public Function DropTypeTable(ByVal tableName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand($"DROP TABLE IF EXISTS `{tableName}`", _con)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("DropTypeTable Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Renames a per-type table when a type name is edited in Settings.
    ''' </summary>
    Public Function RenameTypeTable(ByVal oldTableName As String, ByVal newTableName As String) As Boolean
        Try
            If oldTableName.Equals(newTableName, StringComparison.OrdinalIgnoreCase) Then Return True
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand($"RENAME TABLE `{oldTableName}` TO `{newTableName}`", _con)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("RenameTypeTable Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Gets all type table names from type_details for a given category.
    ''' </summary>
    Public Function GetAllTypeTables(Optional category As String = "") As List(Of String)
        Dim tables As New List(Of String)()
        Try
            Dim query As String = If(String.IsNullOrEmpty(category),
                "SELECT TypeName, Category FROM type_details",
                $"SELECT TypeName, Category FROM type_details WHERE Category = '{category}'")
            Dim dt = ReadDatatable(query)
            For Each row As DataRow In dt.Rows
                tables.Add(TypeNameToTableName(row("TypeName").ToString()))
            Next
        Catch ex As Exception
            Console.WriteLine("GetAllTypeTables Error: " & ex.Message)
        End Try
        Return tables
    End Function

    ''' <summary>
    ''' Migrates existing per-type tables to include the uploaded_doc column if missing.
    ''' </summary>
    Public Sub CheckAndAddUploadedDocColumn()
        Try
            Dim tables = GetAllTypeTables()
            If MySQLDBConnect() = 1 Then
                For Each tbl In tables
                    Dim checkQuery As String = $"SELECT COUNT(*) FROM information_schema.columns WHERE table_name = '{tbl}' AND column_name = 'uploaded_doc' AND table_schema = DATABASE()"
                    Using cmdCheck As New MySqlCommand(checkQuery, _con)
                        Dim count = Convert.ToInt32(cmdCheck.ExecuteScalar())
                        If count = 0 Then
                            Dim alterQuery As String = $"ALTER TABLE `{tbl}` ADD COLUMN uploaded_doc TEXT"
                            Using cmdAlter As New MySqlCommand(alterQuery, _con)
                                cmdAlter.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("CheckAndAddUploadedDocColumn Error: " & ex.Message)
        End Try
    End Sub

    ' --- INSTRUMENT & GAUGE CRUD ---

    Public Function Insertinstrument(ByVal tableName As String, ByVal insName As String, ByVal insDesc As String, ByVal controlNo As String, ByVal color As String, ByVal modelNo As String, ByVal line As String, ByVal assetNo As String, ByVal size As String, ByVal makerName As String, ByVal section As String, ByVal location As String, ByVal requestNo As String, ByVal remark As String, ByVal catCtrl As String, ByVal rfidTag As String, ByVal uploadedDoc As String, ByVal flag As Integer, Optional customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = $"INSERT INTO `{tableName}` (Date, Time, InstrumentName, InstrumentDescription, ControlNo, Color, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, CategoryControl, RFID_tag, uploaded_doc, Flag, InstrumentStatus) " &
                                  "VALUES (@date, @time, @name, @desc, @control, @color, @model, @line, @asset, @size, @maker, @sec, @loc, @rem, @req, @cat, @rfid, @doc, @flag, 'Active')"
                Using cmd As New MySqlCommand(query, _con)
                    If customDate.HasValue Then
                        cmd.Parameters.AddWithValue("@date", customDate.Value)
                    Else
                        cmd.Parameters.AddWithValue("@date", DBNull.Value)
                    End If
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@name", insName)
                    cmd.Parameters.AddWithValue("@desc", insDesc)
                    cmd.Parameters.AddWithValue("@control", controlNo)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@model", modelNo)
                    cmd.Parameters.AddWithValue("@line", line)
                    cmd.Parameters.AddWithValue("@asset", assetNo)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@maker", makerName)
                    cmd.Parameters.AddWithValue("@sec", section)
                    cmd.Parameters.AddWithValue("@loc", location)
                    cmd.Parameters.AddWithValue("@rem", remark)
                    cmd.Parameters.AddWithValue("@req", requestNo)
                    cmd.Parameters.AddWithValue("@cat", catCtrl)
                    cmd.Parameters.AddWithValue("@rfid", rfidTag)
                    cmd.Parameters.AddWithValue("@doc", uploadedDoc)
                    cmd.Parameters.AddWithValue("@flag", flag)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Insertinstrument Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function Updateinstrument(ByVal tableName As String, ByVal id As Integer, ByVal insName As String, ByVal insDesc As String, ByVal controlNo As String, ByVal color As String, ByVal modelNo As String, ByVal line As String, ByVal assetNo As String, ByVal size As String, ByVal makerName As String, ByVal section As String, ByVal location As String, ByVal requestNo As String, ByVal remark As String, ByVal catCtrl As String, ByVal rfidTag As String, ByVal uploadedDoc As String, ByVal flag As Integer, Optional customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim dateClause = If(customDate.HasValue, ", Date=@date", "")
                Dim query As String = $"UPDATE `{tableName}` SET InstrumentName=@name, InstrumentDescription=@desc, ControlNo=@control, Color=@color, " &
                                  "InstrumentModelNo=@model, Line=@line, AssetNo=@asset, Size=@size, MakerName=@maker, Section=@sec, " &
                                  "Location=@loc, Remark=@rem, RequestNo=@req, CategoryControl=@cat, RFID_tag=@rfid, uploaded_doc=@doc, Flag=@flag" & dateClause & " WHERE ID=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@name", insName)
                    cmd.Parameters.AddWithValue("@desc", insDesc)
                    cmd.Parameters.AddWithValue("@control", controlNo)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@model", modelNo)
                    cmd.Parameters.AddWithValue("@line", line)
                    cmd.Parameters.AddWithValue("@asset", assetNo)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@maker", makerName)
                    cmd.Parameters.AddWithValue("@sec", section)
                    cmd.Parameters.AddWithValue("@loc", location)
                    cmd.Parameters.AddWithValue("@rem", remark)
                    cmd.Parameters.AddWithValue("@req", requestNo)
                    cmd.Parameters.AddWithValue("@cat", catCtrl)
                    cmd.Parameters.AddWithValue("@rfid", rfidTag)
                    cmd.Parameters.AddWithValue("@doc", uploadedDoc)
                    cmd.Parameters.AddWithValue("@flag", flag)
                    cmd.Parameters.AddWithValue("@id", id)
                    If customDate.HasValue Then cmd.Parameters.AddWithValue("@date", customDate.Value)
                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Updateinstrument Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function Insertgauge(ByVal tableName As String, ByVal type As String, ByVal gaugeDesc As String, ByVal controlNo As String, ByVal color As String, ByVal drgNo As String, ByVal makerName As String, ByVal model As String, ByVal line As String, ByVal size As String, ByVal tol As String, ByVal tol2 As String, ByVal tol3 As String, ByVal section As String, ByVal location As String, ByVal dateAdd As String, ByVal requestNo As String, ByVal remark As String, ByVal catCtrl As String, ByVal rfidTag As String, ByVal uploadedDoc As String, ByVal flag As Integer, Optional customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = $"INSERT INTO `{tableName}` (Date, Time, GaugeName, GaugeDescription, ControlNo, Color, DrgNo, MakerName, Model, Line, Size, Tol, Tol2, Tol3, Section, Location, DateofAddition, RequestNo, Remark, CategoryControl, RFID_tag, uploaded_doc, Flag, InstrumentStatus) " &
                                  "VALUES (@date, @time, @name, @desc, @control, @color, @drg, @maker, @model, @line, @size, @tol, @tol2, @tol3, @sec, @loc, @dateAdd, @req, @rem, @cat, @rfid, @doc, @flag, 'Active')"
                Using cmd As New MySqlCommand(query, _con)
                    If customDate.HasValue Then
                        cmd.Parameters.AddWithValue("@date", customDate.Value)
                    Else
                        cmd.Parameters.AddWithValue("@date", DBNull.Value)
                    End If
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@name", type)
                    cmd.Parameters.AddWithValue("@desc", gaugeDesc)
                    cmd.Parameters.AddWithValue("@control", controlNo)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@drg", drgNo)
                    cmd.Parameters.AddWithValue("@maker", makerName)
                    cmd.Parameters.AddWithValue("@model", model)
                    cmd.Parameters.AddWithValue("@line", line)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@tol", tol)
                    cmd.Parameters.AddWithValue("@tol2", tol2)
                    cmd.Parameters.AddWithValue("@tol3", tol3)
                    cmd.Parameters.AddWithValue("@sec", section)
                    cmd.Parameters.AddWithValue("@loc", location)
                    cmd.Parameters.AddWithValue("@dateAdd", dateAdd)
                    cmd.Parameters.AddWithValue("@req", requestNo)
                    cmd.Parameters.AddWithValue("@rem", remark)
                    cmd.Parameters.AddWithValue("@cat", catCtrl)
                    cmd.Parameters.AddWithValue("@rfid", rfidTag)
                    cmd.Parameters.AddWithValue("@doc", uploadedDoc)
                    cmd.Parameters.AddWithValue("@flag", flag)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Insertgauge Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function Updategauge(ByVal tableName As String, ByVal id As Integer, ByVal type As String, ByVal gaugeDesc As String, ByVal controlNo As String, ByVal color As String, ByVal drgNo As String, ByVal makerName As String, ByVal model As String, ByVal line As String, ByVal size As String, ByVal tol As String, ByVal tol2 As String, ByVal tol3 As String, ByVal section As String, ByVal location As String, ByVal dateAdd As String, ByVal requestNo As String, ByVal remark As String, ByVal catCtrl As String, ByVal rfidTag As String, ByVal uploadedDoc As String, ByVal flag As Integer, Optional customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim dateClause = If(customDate.HasValue, ", Date=@date", "")
                Dim query As String = $"UPDATE `{tableName}` SET GaugeName=@name, GaugeDescription=@desc, ControlNo=@control, Color=@color, " &
                                  "DrgNo=@drg, MakerName=@maker, Model=@model, Line=@line, Size=@size, Tol=@tol, Tol2=@tol2, Tol3=@tol3, " &
                                  "Section=@sec, Location=@loc, DateofAddition=@dateAdd, RequestNo=@req, Remark=@rem, CategoryControl=@cat, RFID_tag=@rfid, uploaded_doc=@doc, Flag=@flag" & dateClause & " WHERE ID=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@name", type)
                    cmd.Parameters.AddWithValue("@desc", gaugeDesc)
                    cmd.Parameters.AddWithValue("@control", controlNo)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@drg", drgNo)
                    cmd.Parameters.AddWithValue("@maker", makerName)
                    cmd.Parameters.AddWithValue("@model", model)
                    cmd.Parameters.AddWithValue("@line", line)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@tol", tol)
                    cmd.Parameters.AddWithValue("@tol2", tol2)
                    cmd.Parameters.AddWithValue("@tol3", tol3)
                    cmd.Parameters.AddWithValue("@sec", section)
                    cmd.Parameters.AddWithValue("@loc", location)
                    cmd.Parameters.AddWithValue("@dateAdd", dateAdd)
                    cmd.Parameters.AddWithValue("@req", requestNo)
                    cmd.Parameters.AddWithValue("@rem", remark)
                    cmd.Parameters.AddWithValue("@cat", catCtrl)
                    cmd.Parameters.AddWithValue("@rfid", rfidTag)
                    cmd.Parameters.AddWithValue("@doc", uploadedDoc)
                    cmd.Parameters.AddWithValue("@flag", flag)
                    cmd.Parameters.AddWithValue("@id", id)
                    If customDate.HasValue Then cmd.Parameters.AddWithValue("@date", customDate.Value)
                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Updategauge Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' --- SEARCH & PAGINATION & FILTERING ---

    ''' <summary>Generic per-type inventory query for Instrument-schema tables.</summary>
    Public Function GetInstrumentInventoryData(tableName As String, offset As Integer, limit As Integer, filters As Dictionary(Of String, String), Optional keyword As String = "") As DataTable
        Dim whereClause As String = " WHERE 1=1"
        Dim parameters As New List(Of MySqlParameter)()

        If filters IsNot Nothing Then
            For Each kvp In filters
                If kvp.Key = "Group" Then
                    whereClause &= $" AND ControlNo LIKE @grp"
                    parameters.Add(New MySqlParameter("@grp", kvp.Value & "-%"))
                Else
                    Dim dbCol As String = If(kvp.Key = "Status", "InstrumentStatus", kvp.Key)
                    whereClause &= $" AND `{dbCol}` = @f_{kvp.Key}"
                    parameters.Add(New MySqlParameter($"@f_{kvp.Key}", kvp.Value))
                End If
            Next
        End If

        If Not String.IsNullOrWhiteSpace(keyword) Then
            Dim words = keyword.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            Dim conds As New List(Of String)()
            Dim searchCols As String = "ControlNo, InstrumentName, InstrumentDescription, Color, CategoryControl, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag"
            For Each w In words
                Dim wc = w.Replace("'", "''")
                conds.Add($"(CONCAT_WS('|', {searchCols}) LIKE '%{wc}%')")
            Next
            whereClause &= " AND (" & String.Join(" AND ", conds) & ")"
        End If

        Dim query As String = $"SELECT ID, Date, Time, ControlNo, InstrumentName, InstrumentDescription, Color, CategoryControl, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag, uploaded_doc, InstrumentStatus AS Status FROM `{tableName}`" & whereClause & $" ORDER BY Date DESC, Time DESC LIMIT {limit} OFFSET {offset}"
        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddRange(parameters.ToArray())
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)
                        Return dt
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetInstrumentInventoryData Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>Generic per-type count for Instrument-schema tables.</summary>
    Public Function GetInstrumentInventoryCount(tableName As String, filters As Dictionary(Of String, String), Optional keyword As String = "") As Integer
        Dim whereClause As String = " WHERE 1=1"
        Dim parameters As New List(Of MySqlParameter)()

        If filters IsNot Nothing Then
            For Each kvp In filters
                If kvp.Key = "Group" Then
                    whereClause &= $" AND ControlNo LIKE @grp"
                    parameters.Add(New MySqlParameter("@grp", kvp.Value & "-%"))
                Else
                    Dim dbCol As String = If(kvp.Key = "Status", "InstrumentStatus", kvp.Key)
                    whereClause &= $" AND `{dbCol}` = @f_{kvp.Key}"
                    parameters.Add(New MySqlParameter($"@f_{kvp.Key}", kvp.Value))
                End If
            Next
        End If

        If Not String.IsNullOrWhiteSpace(keyword) Then
            Dim words = keyword.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            Dim conds As New List(Of String)()
            Dim searchCols As String = "ControlNo, InstrumentName, InstrumentDescription, Color, CategoryControl, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag"
            For Each w In words
                Dim wc = w.Replace("'", "''")
                conds.Add($"(CONCAT_WS('|', {searchCols}) LIKE '%{wc}%')")
            Next
            whereClause &= " AND (" & String.Join(" AND ", conds) & ")"
        End If

        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand($"SELECT COUNT(*) FROM `{tableName}`" & whereClause, _con)
                    cmd.Parameters.AddRange(parameters.ToArray())
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetInstrumentInventoryCount Error: " & ex.Message)
        End Try
        Return 0
    End Function

    ' [LEGACY ARCHITECTURE NOTE]
    ' The following functions were part of the old architecture using hardcoded `addinstrument` and `addgauge` tables.
    ' They are no longer used by the UI and have been replaced by dynamic queries on the `department_list` master table
    ' and specific per-type tables (e.g., `vernier_caliper`).

    ' Public Function GetCombinedInventoryData(offset As Integer, limit As Integer, filters As Dictionary(Of String, String), Optional keyword As String = "") As DataTable
    '     Dim whereClause As String = " WHERE 1=1"
    '     Dim parameters As New List(Of MySqlParameter)()
    '
    '     If filters IsNot Nothing Then
    '         For Each kvp In filters
    '             If kvp.Key = "Group" Then
    '                 whereClause &= $" AND ControlNo LIKE @{kvp.Key}"
    '                 parameters.Add(New MySqlParameter($"@{kvp.Key}", kvp.Value & "-%"))
    '             ElseIf kvp.Key = "Prefix" Then
    '                 whereClause &= $" AND ControlNo LIKE @{kvp.Key}Prefix"
    '                 parameters.Add(New MySqlParameter($"@{kvp.Key}Prefix", kvp.Value & "%"))
    '             ElseIf kvp.Key = "InstrumentName" Then
    '                 Dim cleanFilter = kvp.Value.ToLower().Replace(" ", "").Replace("-", "").Replace("'", "''")
    '                 whereClause &= $" AND (REPLACE(REPLACE(LOWER(InstrumentName), ' ', ''), '-', '') LIKE '%{cleanFilter}%' OR '{cleanFilter}' LIKE CONCAT('%', REPLACE(REPLACE(LOWER(InstrumentName), ' ', ''), '-', ''), '%'))"
    '             Else
    '                 Dim dbCol As String = kvp.Key
    '                 whereClause &= $" AND `{dbCol}` = @{kvp.Key}"
    '                 parameters.Add(New MySqlParameter($"@{kvp.Key}", kvp.Value))
    '             End If
    '         Next
    '     End If
    '
    '     If Not String.IsNullOrWhiteSpace(keyword) Then
    '         Dim words = keyword.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
    '         Dim searchConditions As New List(Of String)()
    '         Dim searchCols As String = "ControlNo, InstrumentName, InstrumentDescription, Color, CategoryControl, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag"
    '
    '         For Each w In words
    '             Dim wClean = w.Replace("'", "''")
    '             searchConditions.Add($"(CONCAT_WS('|', {searchCols}) LIKE '%{wClean}%')")
    '         Next
    '         whereClause &= " AND (" & String.Join(" AND ", searchConditions) & ")"
    '     End If
    '
    '     Dim query As String =
    '         "SELECT * FROM (" &
    '         " (SELECT ID, Date, Time, ControlNo, 'Instrument' as Type, InstrumentName, InstrumentDescription, Color, CategoryControl, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag, uploaded_doc, InstrumentStatus AS Status FROM addinstrument)" &
    '         " UNION ALL" &
    '         " (SELECT ID, Date, Time, ControlNo, 'Gauge' as Type, GaugeName as InstrumentName, GaugeDescription as InstrumentDescription, Color, CategoryControl, Model as InstrumentModelNo, Line, '' as AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag, uploaded_doc, InstrumentStatus AS Status FROM addgauge)" &
    '         ") as Combined" & whereClause &
    '         $" ORDER BY Date DESC, Time DESC LIMIT {limit} OFFSET {offset}"
    '
    '     Try
    '         If MySQLDBConnect() = 1 Then
    '             Using cmd As New MySqlCommand(query, _con)
    '                 cmd.Parameters.AddRange(parameters.ToArray())
    '                 Using adapter As New MySqlDataAdapter(cmd)
    '                     Dim dt As New DataTable()
    '                     adapter.Fill(dt)
    '                     Return dt
    '                 End Using
    '             End Using
    '         End If
    '     Catch ex As Exception
    '         Console.WriteLine("GetCombinedInventoryData Error: " & ex.Message)
    '     End Try
    '     Return New DataTable()
    ' End Function
    '
    ' Public Function GetFilteredCombinedTotalCount(filters As Dictionary(Of String, String), Optional keyword As String = "") As Integer
    '     Dim whereClause As String = " WHERE 1=1"
    '     Dim parameters As New List(Of MySqlParameter)()
    '
    '     If filters IsNot Nothing Then
    '         For Each kvp In filters
    '             If kvp.Key = "Group" Then
    '                 whereClause &= $" AND ControlNo LIKE @{kvp.Key}"
    '                 parameters.Add(New MySqlParameter($"@{kvp.Key}", kvp.Value & "-%"))
    '             ElseIf kvp.Key = "Prefix" Then
    '                 whereClause &= $" AND ControlNo LIKE @{kvp.Key}Prefix"
    '                 parameters.Add(New MySqlParameter($"@{kvp.Key}Prefix", kvp.Value & "%"))
    '             ElseIf kvp.Key = "InstrumentName" Then
    '                 Dim cleanFilter = kvp.Value.ToLower().Replace(" ", "").Replace("-", "").Replace("'", "''")
    '                 whereClause &= $" AND (REPLACE(REPLACE(LOWER(InstrumentName), ' ', ''), '-', '') LIKE '%{cleanFilter}%' OR '{cleanFilter}' LIKE CONCAT('%', REPLACE(REPLACE(LOWER(InstrumentName), ' ', ''), '-', ''), '%'))"
    '             Else
    '                 Dim dbCol As String = kvp.Key
    '                 whereClause &= $" AND `{dbCol}` = @{kvp.Key}"
    '                 parameters.Add(New MySqlParameter($"@{kvp.Key}", kvp.Value))
    '             End If
    '         Next
    '     End If
    '
    '     If Not String.IsNullOrWhiteSpace(keyword) Then
    '         Dim words = keyword.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
    '         Dim searchConditions As New List(Of String)()
    '         Dim searchCols As String = "ControlNo, InstrumentName, InstrumentDescription, Color, CategoryControl, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag"
    '
    '         For Each w In words
    '             Dim wClean = w.Replace("'", "''")
    '             searchConditions.Add($"(CONCAT_WS('|', {searchCols}) LIKE '%{wClean}%')")
    '         Next
    '         whereClause &= " AND (" & String.Join(" AND ", searchConditions) & ")"
    '     End If
    '
    '     Dim query As String = "SELECT COUNT(*) FROM (" &
    '         " (SELECT ControlNo, InstrumentName, InstrumentDescription, Color, CategoryControl, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag, InstrumentStatus AS Status FROM addinstrument)" &
    '         " UNION ALL" &
    '         " (SELECT ControlNo, GaugeName as InstrumentName, GaugeDescription as InstrumentDescription, Color, CategoryControl, Model as InstrumentModelNo, Line, '' as AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, RFID_tag, InstrumentStatus AS Status FROM addgauge)" &
    '         ") as Combined" & whereClause
    '
    '     Try
    '         If MySQLDBConnect() = 1 Then
    '             Using cmd As New MySqlCommand(query, _con)
    '                 cmd.Parameters.AddRange(parameters.ToArray())
    '                 Return Convert.ToInt32(cmd.ExecuteScalar())
    '             End Using
    '         End If
    '     Catch ex As Exception
    '         Console.WriteLine("GetFilteredCombinedTotalCount Error: " & ex.Message)
    '     End Try
    '     Return 0
    ' End Function

    Public Function GetGaugeInventoryData(tableName As String, offset As Integer, limit As Integer, filters As Dictionary(Of String, String), Optional keyword As String = "") As DataTable
        Dim whereClause As String = " WHERE 1=1"
        Dim parameters As New List(Of MySqlParameter)()

        If filters IsNot Nothing Then
            For Each kvp In filters
                If kvp.Key = "Group" Then
                    whereClause &= $" AND ControlNo LIKE @grp"
                    parameters.Add(New MySqlParameter("@grp", kvp.Value & "-%"))
                Else
                    Dim dbCol As String = If(kvp.Key = "Status", "InstrumentStatus", kvp.Key)
                    whereClause &= $" AND `{dbCol}` = @f_{kvp.Key}"
                    parameters.Add(New MySqlParameter($"@f_{kvp.Key}", kvp.Value))
                End If
            Next
        End If

        If Not String.IsNullOrWhiteSpace(keyword) Then
            Dim words = keyword.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            Dim searchConditions As New List(Of String)()
            Dim searchCols As String = "ControlNo, GaugeName, GaugeDescription, Color, CategoryControl, Model, Line, Size, Tol, Tol2, Tol3, DrgNo, MakerName, Section, Location, Remark, RequestNo, RFID_tag"

            For Each w In words
                Dim wClean = w.Replace("'", "''")
                searchConditions.Add($"(CONCAT_WS('|', {searchCols}) LIKE '%{wClean}%')")
            Next
            whereClause &= " AND (" & String.Join(" AND ", searchConditions) & ")"
        End If

        Dim query As String = $"SELECT ID, Date, Time, ControlNo, GaugeName, GaugeDescription, Color, CategoryControl, DrgNo, Model, Line, Size, Tol, Tol2, Tol3, MakerName, Section, Location, DateofAddition, Remark, RequestNo, RFID_tag, uploaded_doc, InstrumentStatus AS Status FROM `{tableName}`" & whereClause & $" ORDER BY Date DESC, Time DESC LIMIT {limit} OFFSET {offset}"

        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddRange(parameters.ToArray())
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)
                        Return dt
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetGaugeInventoryData Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function GetFilteredGaugeTotalCount(tableName As String, filters As Dictionary(Of String, String), Optional keyword As String = "") As Integer
        Dim whereClause As String = " WHERE 1=1"
        Dim parameters As New List(Of MySqlParameter)()

        If filters IsNot Nothing Then
            For Each kvp In filters
                If kvp.Key = "Group" Then
                    whereClause &= $" AND ControlNo LIKE @grp"
                    parameters.Add(New MySqlParameter("@grp", kvp.Value & "-%"))
                Else
                    Dim dbCol As String = If(kvp.Key = "Status", "InstrumentStatus", kvp.Key)
                    whereClause &= $" AND `{dbCol}` = @f_{kvp.Key}"
                    parameters.Add(New MySqlParameter($"@f_{kvp.Key}", kvp.Value))
                End If
            Next
        End If

        If Not String.IsNullOrWhiteSpace(keyword) Then
            Dim words = keyword.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            Dim searchConditions As New List(Of String)()
            Dim searchCols As String = "ControlNo, GaugeName, GaugeDescription, Color, CategoryControl, Model, Line, Size, Tol, Tol2, Tol3, DrgNo, MakerName, Section, Location, Remark, RequestNo, RFID_tag"

            For Each w In words
                Dim wClean = w.Replace("'", "''")
                searchConditions.Add($"(CONCAT_WS('|', {searchCols}) LIKE '%{wClean}%')")
            Next
            whereClause &= " AND (" & String.Join(" AND ", searchConditions) & ")"
        End If

        Dim query As String = $"SELECT COUNT(*) FROM `{tableName}`" & whereClause

        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddRange(parameters.ToArray())
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetFilteredGaugeTotalCount Error: " & ex.Message)
        End Try
        Return 0
    End Function

    ' --- TYPE DETAILS CRUD ---

    Public Function GetTypeDetails() As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM type_details ORDER BY Category, TypeName"
                Using cmd As New MySqlCommand(query, _con)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)
                        Return dt
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetTypeDetails Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function GetPrefixByTypeName(typeName As String) As String
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT BasePrefix FROM type_details WHERE TypeName = @typeName LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@typeName", typeName)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetPrefixByTypeName Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Sub CheckAndAddTypeImageColumn()
        Try
            If MySQLDBConnect() = 1 Then
                ' Check if column exists first for compatibility
                Dim checkQuery As String = "SELECT COUNT(*) FROM information_schema.columns WHERE table_name = 'type_details' AND column_name = 'TypeImage' AND table_schema = DATABASE()"
                Using cmdCheck As New MySqlCommand(checkQuery, _con)
                    Dim count = Convert.ToInt32(cmdCheck.ExecuteScalar())
                    If count = 0 Then
                        Dim alterQuery As String = "ALTER TABLE type_details ADD COLUMN TypeImage LONGTEXT"
                        Using cmdAlter As New MySqlCommand(alterQuery, _con)
                            cmdAlter.ExecuteNonQuery()
                        End Using
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("CheckAndAddTypeImageColumn Error: " & ex.Message)
        End Try
    End Sub

    Public Function InsertTypeDetail(category As String, typeName As String, prefixMode As String, basePrefix As String, serialDigits As String, Optional typeImage As String = "") As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO type_details (Category, TypeName, PrefixMode, BasePrefix, SerialDigits, TypeImage) VALUES (@cat, @name, @mode, @prefix, @digits, @img)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@cat", category)
                    cmd.Parameters.AddWithValue("@name", typeName)
                    cmd.Parameters.AddWithValue("@mode", prefixMode)
                    cmd.Parameters.AddWithValue("@prefix", basePrefix)
                    cmd.Parameters.AddWithValue("@digits", serialDigits)
                    cmd.Parameters.AddWithValue("@img", If(String.IsNullOrEmpty(typeImage), DBNull.Value, typeImage))
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertTypeDetail Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateTypeDetail(id As Integer, category As String, typeName As String, prefixMode As String, basePrefix As String, serialDigits As String, Optional typeImage As String = "") As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE type_details SET Category=@cat, TypeName=@name, PrefixMode=@mode, BasePrefix=@prefix, SerialDigits=@digits, TypeImage=@img WHERE ID=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.Parameters.AddWithValue("@cat", category)
                    cmd.Parameters.AddWithValue("@name", typeName)
                    cmd.Parameters.AddWithValue("@mode", prefixMode)
                    cmd.Parameters.AddWithValue("@prefix", basePrefix)
                    cmd.Parameters.AddWithValue("@digits", serialDigits)
                    cmd.Parameters.AddWithValue("@img", If(String.IsNullOrEmpty(typeImage), DBNull.Value, typeImage))
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateTypeDetail Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' --- CATEGORY & CONTROL NO HELPERS ---

    Public Function IsControlNoDuplicate(ByVal tableName As String, ByVal controlNo As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = $"SELECT COUNT(1) FROM {tableName} WHERE ControlNo = @controlNo"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    Return (Convert.ToInt32(cmd.ExecuteScalar()) > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("IsControlNoDuplicate Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Sub SyncCategoryControlFromInventory()
        Try
            If MySQLDBConnect() = 1 Then
                Dim existingCats As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim dtExisting As DataTable = ReadDatatable("SELECT DISTINCT CategoryControl, size FROM categorycontrol")
                For Each row As DataRow In dtExisting.Rows
                    existingCats.Add(row("CategoryControl").ToString() & "|" & row("size").ToString())
                Next

                ' Build UNION query across all type tables
                Dim typeRows = ReadDatatable("SELECT TypeName, Category FROM type_details")
                Dim unionParts As New List(Of String)()
                For Each typeRow As DataRow In typeRows.Rows
                    Dim tbl = TypeNameToTableName(typeRow("TypeName").ToString())
                    Dim cat = typeRow("Category").ToString()
                    If cat = "Instrument" Then
                        unionParts.Add($"SELECT ControlNo, InstrumentName AS Name, InstrumentDescription AS Description, Size, 'Instrument' AS Type FROM `{tbl}` WHERE ControlNo LIKE '%-%'")
                    Else
                        unionParts.Add($"SELECT ControlNo, GaugeName AS Name, GaugeDescription AS Description, Size, 'Gauge' AS Type FROM `{tbl}` WHERE ControlNo LIKE '%-%'")
                    End If
                Next

                If unionParts.Count = 0 Then Return
                Dim queryInv As String = String.Join(" UNION ALL ", unionParts)
                Dim dtInv As DataTable = ReadDatatable(queryInv)

                For Each row As DataRow In dtInv.Rows
                    Dim controlNo As String = row("ControlNo").ToString()
                    Dim size As String = row("Size").ToString()
                    Dim parts = controlNo.Split("-"c)
                    If parts.Length > 1 Then
                        Dim cat = parts(0).Trim()
                        If Not existingCats.Contains(cat & "|" & size) Then
                            Dim insertQuery As String = "INSERT INTO categorycontrol (Date, Time, Type, InstrumentName, InstrumentDescription, CategoryControl, size) " &
                                                      "VALUES (@date, @time, @type, @name, @desc, @cat, @size)"
                            Using cmd As New MySqlCommand(insertQuery, _con)
                                cmd.Parameters.AddWithValue("@date", DateTime.Today)
                                cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                                cmd.Parameters.AddWithValue("@type", row("Type"))
                                cmd.Parameters.AddWithValue("@name", row("Name"))
                                cmd.Parameters.AddWithValue("@desc", row("Description"))
                                cmd.Parameters.AddWithValue("@cat", cat)
                                cmd.Parameters.AddWithValue("@size", size)
                                cmd.ExecuteNonQuery()
                                existingCats.Add(cat & "|" & size)
                            End Using
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("SyncCategoryControlFromInventory Error: " & ex.Message)
        End Try
    End Sub

    ' --- MISC HELPERS ---

    Public Shared Function ToTitleCase(ByVal input As String) As String
        If String.IsNullOrWhiteSpace(input) Then Return input
        Dim words = input.Split(" "c)
        Dim result As New List(Of String)()
        Dim textInfo = System.Globalization.CultureInfo.CurrentCulture.TextInfo

        For Each word In words
            If String.IsNullOrWhiteSpace(word) Then Continue For
            ' Preserve acronyms ONLY if user typed them in UpperCase (2-5 chars)
            ' This allows CMM to stay CMM, but cmm becomes Cmm
            If word.Length >= 2 AndAlso word.Length <= 5 AndAlso word.All(AddressOf Char.IsUpper) Then
                result.Add(word)
            Else
                result.Add(textInfo.ToTitleCase(word.ToLower()))
            End If
        Next

        Return String.Join(" ", result)
    End Function

    Public Function DeleteRecord(ByVal tableName As String, ByVal id As Integer) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = $"DELETE FROM {tableName} WHERE ID = @id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@id", id)
                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("DeleteRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function ExecuteNonQuery(ByVal query As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    Return cmd.ExecuteNonQuery() >= 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("ExecuteNonQuery Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function ExecuteNonQuery(ByVal query As String, ByVal ParamArray params As MySqlParameter()) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    If params IsNot Nothing Then
                        cmd.Parameters.AddRange(params)
                    End If
                    Return cmd.ExecuteNonQuery() >= 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("ExecuteNonQuery (Params) Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetTableData(ByVal tableName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM {tableName}")
    End Function

    Public Function GetCalibrationMapping(ByVal controlNo As String) As String
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT Form FROM calibrationmapping WHERE ControlNo = @controlNo"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    Dim res = cmd.ExecuteScalar()
                    Return If(res IsNot Nothing, res.ToString(), "")
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetCalibrationMapping Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function InsertNGList(ByVal controlNo As String, ByVal type As String, ByVal name As String, ByVal size As String, ByVal deptSec As String, ByVal reason As String, ByVal color As String, ByVal dateComm As String, ByVal remarks As String, ByVal writeoffNo As Integer) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO ng_list (Date, Time, ControlNo, Type, Name, Size, Department_Section, Reason, Color, DateOfCommunication, Remarks, WriteOffNo) " &
                                  "VALUES (@date, @time, @ctrl, @type, @name, @size, @dept, @reason, @color, @dateComm, @rem, @wo)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@name", name)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@dept", deptSec)
                    cmd.Parameters.AddWithValue("@reason", reason)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@dateComm", dateComm)
                    cmd.Parameters.AddWithValue("@rem", remarks)
                    cmd.Parameters.AddWithValue("@wo", writeoffNo)
                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertNGList Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateWriteoffFlag(ByVal tableName As String, ByVal controlNo As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim q As String = $"UPDATE `{tableName}` SET Flag = 1 WHERE ControlNo = @ctrl"
                Using cmd As New MySqlCommand(q, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateWriteoffFlag Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' [LEGACY] Replaced by UpdateRecordStatus(tableName, controlNo, flag, status)
    ' which correctly targets the new type-specific tables (cards).
    ' Public Function UpdateInstrumentStatus(ByVal controlNo As String, ByVal statusValue As Integer) As Boolean
    '     Try
    '         If MySQLDBConnect() = 1 Then
    '             Dim q1 As String = "UPDATE addgauge SET InstrumentStatus = @stat WHERE ControlNo = @ctrl"
    '             Dim q2 As String = "UPDATE addinstrument SET InstrumentStatus = @stat WHERE ControlNo = @ctrl"
    '             Using cmd1 As New MySqlCommand(q1, _con), cmd2 As New MySqlCommand(q2, _con)
    '                 cmd1.Parameters.AddWithValue("@stat", statusValue)
    '                 cmd1.Parameters.AddWithValue("@ctrl", controlNo)
    '                 cmd2.Parameters.AddWithValue("@stat", statusValue)
    '                 cmd2.Parameters.AddWithValue("@ctrl", controlNo)
    '                 cmd1.ExecuteNonQuery()
    '                 cmd2.ExecuteNonQuery()
    '                 Return True
    '             End Using
    '         End If
    '     Catch ex As Exception
    '         Console.WriteLine("UpdateInstrumentStatus Error: " & ex.Message)
    '     End Try
    '     Return False
    ' End Function
    Public Function GetCategory(ByVal name As String, ByVal size As String, ByVal type As String) As String
        Try
            If MySQLDBConnect() = 1 Then
                Dim cleanName As String = name.ToLower().Replace(" ", "").Replace("'", "''")
                Dim cleanSize As String = size.ToLower().Replace(" ", "").Replace("'", "''")
                Dim query As String = "SELECT CategoryControl FROM categorycontrol " &
                                     $"WHERE REPLACE(LOWER(InstrumentName), ' ', '') = '{cleanName}' " &
                                     $"AND REPLACE(LOWER(size), ' ', '') = '{cleanSize}' " &
                                     $"AND Type = '{type}' LIMIT 1"
                Dim dt = ReadDatatable(query)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Return dt.Rows(0)("CategoryControl").ToString()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("GetCategory Error: " & ex.Message)
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Gets the BasePrefix configured in type_details for a given gauge name.
    ''' E.g., "PLAIN PLUG GAUGE" → "PS"
    ''' </summary>
    Public Function GetGaugeBasePrefix(ByVal name As String) As String
        Try
            If MySQLDBConnect() = 1 Then
                Dim cleanName As String = name.ToLower().Replace(" ", "").Replace("'", "''")
                Dim query As String = "SELECT BasePrefix FROM type_details " &
                                     $"WHERE REPLACE(LOWER(TypeName), ' ', '') = '{cleanName}' " &
                                     $"AND Category = 'Gauge' LIMIT 1"
                Dim dt = ReadDatatable(query)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Return dt.Rows(0)("BasePrefix").ToString().Trim()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("GetGaugeBasePrefix Error: " & ex.Message)
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Builds the category string from BasePrefix + first digit of size.
    ''' E.g., BasePrefix="PS", size="57.4" → category="PS5"
    ''' Then checks addgauge for existing records matching that category-xxx pattern.
    ''' Returns the category string if an existing match is found, or blank if none exists yet.
    ''' </summary>
    Public Function GetGaugeCategoryByFirstDigit(ByVal basePrefix As String, ByVal size As String) As String
        Try
            ' Extract first digit of size
            Dim firstDigit As String = ""
            For Each ch As Char In size.Trim()
                If Char.IsDigit(ch) Then
                    firstDigit = ch.ToString()
                    Exit For
                End If
            Next
            If String.IsNullOrEmpty(firstDigit) Then Return ""

            Dim candidate As String = basePrefix & firstDigit  ' e.g. PS5
            Return candidate
        Catch ex As Exception
            Console.WriteLine("GetGaugeCategoryByFirstDigit Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function GetNextAvailableCategory(ByVal name As String, ByVal type As String) As String
        Dim prefix As String = ""
        Dim maxNum As Integer = 0

        Try
            If MySQLDBConnect() = 1 Then
                Dim cleanName As String = name.ToLower().Replace(" ", "").Replace("'", "''")

                ' 1. Check if name exists to inheritance prefix from categorycontrol
                Dim qName As String = $"SELECT CategoryControl FROM categorycontrol WHERE REPLACE(LOWER(InstrumentName), ' ', '') = '{cleanName}' AND Type = '{type}'"
                Dim dtName = ReadDatatable(qName)

                If dtName IsNot Nothing AndAlso dtName.Rows.Count > 0 Then
                    For Each row As DataRow In dtName.Rows
                        Dim cat As String = row(0).ToString().Trim()
                        Dim match = System.Text.RegularExpressions.Regex.Match(cat, "^([a-zA-Z]+)(\d*)$")
                        If match.Success Then
                            Dim p As String = match.Groups(1).Value
                            Dim numStr As String = match.Groups(2).Value
                            Dim n As Integer = If(numStr = "", 0, Integer.Parse(numStr))
                            If prefix = "" Then prefix = p
                            If p.Equals(prefix, StringComparison.OrdinalIgnoreCase) AndAlso n > maxNum Then maxNum = n
                        End If
                    Next
                Else
                    ' 2. Try to get BasePrefix from type_details
                    Dim qTypeDetails As String = $"SELECT BasePrefix FROM type_details WHERE REPLACE(LOWER(TypeName), ' ', '') = '{cleanName}' AND Category = '{type}' LIMIT 1"
                    Dim dtType = ReadDatatable(qTypeDetails)

                    If dtType IsNot Nothing AndAlso dtType.Rows.Count > 0 Then
                        prefix = dtType.Rows(0)("BasePrefix").ToString().Trim()
                    Else
                        ' Fallback: Generate prefix from initials
                        Dim words() As String = name.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                        For Each w As String In words
                            If w.Length > 0 AndAlso Char.IsLetter(w(0)) Then prefix &= w(0).ToString().ToUpper()
                        Next
                        If prefix = "" Then prefix = "C"
                    End If

                    ' Check global max for this prefix in categorycontrol
                    Dim qGlobal As String = $"SELECT CategoryControl FROM categorycontrol WHERE CategoryControl REGEXP '^{prefix}[0-9]+$'"
                    Dim dtGlobal = ReadDatatable(qGlobal)
                    If dtGlobal IsNot Nothing AndAlso dtGlobal.Rows.Count > 0 Then
                        For Each row As DataRow In dtGlobal.Rows
                            Dim cat As String = row(0).ToString().Trim()
                            Dim match = System.Text.RegularExpressions.Regex.Match(cat, "^([a-zA-Z]+)(\d*)$")
                            If match.Success Then
                                Dim numStr As String = match.Groups(2).Value
                                Dim n As Integer = If(numStr = "", 0, Integer.Parse(numStr))
                                If n > maxNum Then maxNum = n
                            End If
                        Next
                    End If
                End If
                Return prefix & (maxNum + 1).ToString()
            End If
        Catch ex As Exception
            Console.WriteLine("GetNextAvailableCategory Error: " & ex.Message)
        End Try
        Return "C1"
    End Function

    Public Function GetLastControlNumber(ByVal category As String, ByVal tableName As String) As Integer
        Try
            If MySQLDBConnect() = 1 Then
                ' Improve query to sort numerically by the suffix
                Dim query As String = $"SELECT ControlNo FROM {tableName} " &
                                     $"WHERE ControlNo LIKE '{category.Replace("'", "''")}-%' " &
                                     $"ORDER BY CAST(SUBSTRING_INDEX(ControlNo, '-', -1) AS UNSIGNED) DESC LIMIT 1"
                Dim dt = ReadDatatable(query)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim lastControl As String = dt.Rows(0)("ControlNo").ToString()
                    Dim parts = lastControl.Split("-"c)
                    If parts.Length >= 2 Then
                        Dim numPart As String = parts(parts.Length - 1)
                        Dim num As Integer = 0
                        If Integer.TryParse(numPart, num) Then
                            Return num
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("GetLastControlNumber Error: " & ex.Message)
        End Try
        Return 0
    End Function
    Public Function InsertWriteOff(ByVal writeOffDate As DateTime, ByVal receivedDate As DateTime, ByVal type As String, ByVal instrumentName As String, ByVal controlNo As String, ByVal quantity As String, ByVal line As String, ByVal section As String, ByVal color As String, ByVal reason As String, ByVal action As String, ByVal actionNG As String, ByVal documentPath As String, ByVal writeOffNo As String, ByVal instrumentStatus As String, ByVal raisedBy As String, ByVal interupdated As String, Optional cycleName As String = "") As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO writeoff (WriteOffDate, Time, CycleName, Type, InstrumentName, ControlNo, Quantity, Line, Section, Color, Reason, Action, ActionNG, DocumentPath, WriteOffNo, RecievedDate, InstrumentStatus, RaisedBy, Interupdated) " &
                                  "VALUES (@woDate, @time, @cycle, @type, @insName, @ctrl, @qty, @line, @sec, @color, @reason, @action, @actionNG, @docPath, @woNo, @recDate, @status, @raisedBy, @interupdated)"

                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@woDate", writeOffDate)
                    cmd.Parameters.AddWithValue("@recDate", receivedDate)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@insName", instrumentName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@qty", quantity)
                    cmd.Parameters.AddWithValue("@line", line)
                    cmd.Parameters.AddWithValue("@sec", section)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@reason", reason)
                    cmd.Parameters.AddWithValue("@action", action)
                    cmd.Parameters.AddWithValue("@actionNG", actionNG)
                    cmd.Parameters.AddWithValue("@docPath", documentPath)
                    cmd.Parameters.AddWithValue("@woNo", writeOffNo)
                    cmd.Parameters.AddWithValue("@status", instrumentStatus)
                    cmd.Parameters.AddWithValue("@raisedBy", raisedBy)
                    cmd.Parameters.AddWithValue("@interupdated", interupdated)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertWriteOff Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateRecordStatus(ByVal tableName As String, ByVal controlNo As String, ByVal flag As Integer, ByVal status As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                ' IMPORTANT: Ensure table name is safe to avoid injection (it comes from internal logic, not user input)
                Dim query As String = $"UPDATE {tableName} SET Flag = @flag, InstrumentStatus = @status WHERE ControlNo = @controlNo"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@flag", flag)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateRecordStatus Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertReintroduction(ByVal reintroductionDate As DateTime, ByVal type As String, ByVal instrumentName As String, ByVal controlNo As String, ByVal quantity As String, ByVal line As String, ByVal section As String, ByVal color As String, ByVal reason As String, ByVal writeOffNo As String, ByVal reintroductionNo As String, ByVal foundDate As DateTime, ByVal missingDate As DateTime, ByVal userDeclaration As String, ByVal interStatus As String, ByVal interComent As String, ByVal interUpdated As String, ByVal calibrationResult As String, ByVal raisedBy As String, ByVal remarks As String, ByVal documentPath As String, Optional cycleName As String = "") As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO reintroduction " &
                                  "(ReintroductionDate, Time, CycleName, Type, InstrumentName, ControlNo, Quantity, Line, Section, Color, Reason, WriteOffNo, ReintroductionNo, FoundDate, MissingDate, UserDeclaration, InterStatus, InterComent, InterUpdated, CalibrationResult, RaisedBy, Remarks, DocumentPath) " &
                                  "VALUES " &
                                  "(@rDate, @time, @cycle, @type, @insName, @ctrl, @qty, @line, @sec, @color, @reason, @woNo, @reNo, @fDate, @mDate, @uDec, @iStat, @iCom, @iUp, @calRes, @raised, @rem, @doc)"

                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@rDate", reintroductionDate)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@insName", instrumentName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@qty", quantity)
                    cmd.Parameters.AddWithValue("@line", line)
                    cmd.Parameters.AddWithValue("@sec", section)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@reason", reason)
                    cmd.Parameters.AddWithValue("@woNo", writeOffNo)
                    cmd.Parameters.AddWithValue("@reNo", reintroductionNo)
                    cmd.Parameters.AddWithValue("@fDate", foundDate)
                    cmd.Parameters.AddWithValue("@mDate", missingDate)
                    cmd.Parameters.AddWithValue("@uDec", userDeclaration)
                    cmd.Parameters.AddWithValue("@iStat", interStatus)
                    cmd.Parameters.AddWithValue("@iCom", interComent)
                    cmd.Parameters.AddWithValue("@iUp", interUpdated)
                    cmd.Parameters.AddWithValue("@calRes", calibrationResult)
                    cmd.Parameters.AddWithValue("@raised", raisedBy)
                    cmd.Parameters.AddWithValue("@rem", remarks)
                    cmd.Parameters.AddWithValue("@doc", documentPath)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertReintroduction Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetDepartments() As DataTable
        Return ReadDatatable("SELECT * FROM departmentmaster ORDER BY DepartmentName ASC")
    End Function

    Public Function InsertDepartment(ByVal deptName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO departmentmaster (DepartmentName) VALUES (@name)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@name", deptName)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertDepartment Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' --- SETTINGS IMPORT DUPLICATE-CHECK HELPERS ---

    Public Function IsTypeDuplicate(typeName As String, category As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim clean = typeName.ToLower().Replace(" ", "").Replace("'", "''")
                Dim dt = ReadDatatable($"SELECT ID FROM type_details WHERE REPLACE(LOWER(TypeName),' ','') = '{clean}' AND Category = '{category.Replace("'", "''")}' LIMIT 1")
                Return dt.Rows.Count > 0
            End If
        Catch ex As Exception
            Console.WriteLine("IsTypeDuplicate Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function IsDepartmentDuplicate(deptName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim dt = ReadDatatable($"SELECT ID FROM departmentmaster WHERE LOWER(DepartmentName) = '{deptName.ToLower().Replace("'", "''")}' LIMIT 1")
                Return dt.Rows.Count > 0
            End If
        Catch ex As Exception
            Console.WriteLine("IsDepartmentDuplicate Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function IsInstrumentSizeDuplicate(instrumentType As String, size As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim dt = ReadDatatable($"SELECT ID FROM categorycontrol WHERE LOWER(InstrumentType) = '{instrumentType.ToLower().Replace("'", "''")}' AND LOWER(Size) = '{size.ToLower().Replace("'", "''")}' LIMIT 1")
                Return dt.Rows.Count > 0
            End If
        Catch ex As Exception
            Console.WriteLine("IsInstrumentSizeDuplicate Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function IsGaugeSizeDuplicate(instrumentType As String, size As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim dt = ReadDatatable($"SELECT ID FROM gauge_categorycontrol WHERE LOWER(InstrumentType) = '{instrumentType.ToLower().Replace("'", "''")}' AND LOWER(Size) = '{size.ToLower().Replace("'", "''")}' LIMIT 1")
                Return dt.Rows.Count > 0
            End If
        Catch ex As Exception
            Console.WriteLine("IsGaugeSizeDuplicate Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function IsCalibMappingDuplicate(typeName As String, category As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim dt = ReadDatatable($"SELECT id FROM calibrationmapping WHERE LOWER(type_name) = '{typeName.ToLower().Replace("'", "''")}' AND LOWER(category) = '{category.ToLower().Replace("'", "''")}' LIMIT 1")
                Return dt.Rows.Count > 0
            End If
        Catch ex As Exception
            Console.WriteLine("IsCalibMappingDuplicate Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function IsCalibMasterDuplicate(masterName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim dt = ReadDatatable($"SELECT ID FROM calibrationmaster WHERE LOWER(CalibrationMasterName) = '{masterName.ToLower().Replace("'", "''")}' LIMIT 1")
                Return dt.Rows.Count > 0
            End If
        Catch ex As Exception
            Console.WriteLine("IsCalibMasterDuplicate Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertCalibrationMappingRecord(category As String, typeName As String, prefix As String, calibCat As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO calibrationmapping (category, type_name, prefix, calibration_category) VALUES (@cat, @name, @pre, @cc)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@cat", category)
                    cmd.Parameters.AddWithValue("@name", typeName)
                    cmd.Parameters.AddWithValue("@pre", prefix)
                    cmd.Parameters.AddWithValue("@cc", calibCat)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertCalibrationMappingRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' --- CALIBRATION MASTER TYPES CRUD ---

    Public Function GetCalibrationMasterTypes() As DataTable
        Return ReadDatatable("SELECT * FROM calibrationmaster ORDER BY CalibrationMasterName ASC")
    End Function

    Public Function InsertCalibrationMasterType(ByVal masterName As String) As Boolean
        If MySQLDBConnect() = 1 Then
            Dim query As String = "INSERT INTO calibrationmaster (CalibrationMasterName) VALUES (@name)"
            Using cmd As New MySqlCommand(query, _con)
                cmd.Parameters.AddWithValue("@name", masterName)
                cmd.ExecuteNonQuery()
                Return True
            End Using
        End If
        Return False
    End Function

    Public Function GetControlNoPrefixes(ByVal tableName As String) As List(Of String)
        Dim prefixes As New List(Of String)
        prefixes.Add("All")
        Try
            If MySQLDBConnect() = 1 Then
                Dim query = $"SELECT DISTINCT ControlNo FROM `{tableName}` WHERE ControlNo LIKE '%-%'"
                Using cmd As New MySqlCommand(query, _con)
                    Using dr As MySqlDataReader = cmd.ExecuteReader()
                        Dim uniqueSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                        While dr.Read()
                            Dim ctrl = dr("ControlNo").ToString()
                            Dim parts = ctrl.Split("-"c)
                            If parts.Length > 0 Then
                                uniqueSet.Add(parts(0).Trim())
                            End If
                        End While
                        Dim sortedList = uniqueSet.ToList()
                        sortedList.Sort()
                        prefixes.AddRange(sortedList)
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetControlNoPrefixes Error: " & ex.Message)
        End Try
        Return prefixes
    End Function

    ' --- CATEGORY CONTROL (INSTRUMENT SIZES) CRUD ---

    Public Function GetInstrumentSizes() As DataTable
        Return ReadDatatable("SELECT ID, InstrumentType, Size, GroupCode, Active, Sort FROM categorycontrol ORDER BY InstrumentType, Sort")
    End Function

    Public Function InsertInstrumentSize(type As String, size As String, groupCode As String, active As Integer, sort As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO categorycontrol (InstrumentType, Size, GroupCode, Active, Sort) VALUES (@type, @size, @group, @active, @sort)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@group", groupCode)
                    cmd.Parameters.AddWithValue("@active", active)
                    cmd.Parameters.AddWithValue("@sort", sort)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertInstrumentSize Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateInstrumentSize(id As Integer, type As String, size As String, groupCode As String, active As Integer, sort As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE categorycontrol SET InstrumentType=@type, Size=@size, GroupCode=@group, Active=@active, Sort=@sort WHERE ID=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@group", groupCode)
                    cmd.Parameters.AddWithValue("@active", active)
                    cmd.Parameters.AddWithValue("@sort", sort)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateInstrumentSize Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetLocationList() As List(Of String)
        Dim list As New List(Of String)()
        Try
            ' Build union across all type tables
            Dim typeRows = ReadDatatable("SELECT TypeName, Category FROM type_details")
            Dim unionParts As New List(Of String)()
            For Each typeRow As DataRow In typeRows.Rows
                Dim tbl = TypeNameToTableName(typeRow("TypeName").ToString())
                unionParts.Add($"SELECT Location FROM `{tbl}`")
            Next
            If unionParts.Count = 0 Then Return list
            Dim query As String = $"SELECT DISTINCT Location FROM ({String.Join(" UNION ALL ", unionParts)}) as Combined WHERE Location IS NOT NULL AND Location <> '' ORDER BY Location"
            Dim dt = ReadDatatable(query)
            For Each row As DataRow In dt.Rows
                list.Add(row("Location").ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("GetLocationList Error: " & ex.Message)
        End Try
        Return list
    End Function

    Public Function GetGroupCodeBySize(size As String, type As String) As String
        Try
            Dim dt = ReadDatatable($"SELECT GroupCode FROM categorycontrol WHERE Size = '{size}' AND InstrumentType = '{type}' AND Active = 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)("GroupCode").ToString()
            End If
        Catch ex As Exception
            Console.WriteLine("GetGroupCodeBySize Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function GetGaugeGroupCodeBySize(size As String, type As String) As String
        Try
            Dim dt = ReadDatatable($"SELECT GroupCode FROM gauge_categorycontrol WHERE Size = '{size}' AND InstrumentType = '{type}' AND Active = 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)("GroupCode").ToString()
            End If
        Catch ex As Exception
            Console.WriteLine("GetGaugeGroupCodeBySize Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function GetInstrumentBasePrefix(typeName As String) As String
        Try
            Dim dt = ReadDatatable($"SELECT BasePrefix FROM type_details WHERE TypeName = '{typeName}' AND Category = 'Instrument'")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)("BasePrefix").ToString()
            End If
        Catch ex As Exception
            Console.WriteLine("GetInstrumentBasePrefix Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function GetInstrumentTypesOnly() As DataTable
        Return ReadDatatable("SELECT DISTINCT TypeName FROM type_details WHERE Category = 'Instrument' ORDER BY TypeName")
    End Function

    Public Function GetGaugeTypesOnly() As DataTable
        Return ReadDatatable("SELECT DISTINCT TypeName FROM type_details WHERE Category = 'Gauge' ORDER BY TypeName")
    End Function

    ' --- CATEGORY CONTROL (GAUGE SIZES) CRUD ---

    Public Function GetGaugeSizes() As DataTable
        Return ReadDatatable("SELECT ID, InstrumentType, Size, GroupCode, Active, Sort FROM gauge_categorycontrol ORDER BY InstrumentType, Sort")
    End Function

    Public Function InsertGaugeSize(type As String, size As String, groupCode As String, active As Integer, sort As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO gauge_categorycontrol (InstrumentType, Size, GroupCode, Active, Sort) VALUES (@type, @size, @group, @active, @sort)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@group", groupCode)
                    cmd.Parameters.AddWithValue("@active", active)
                    cmd.Parameters.AddWithValue("@sort", sort)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertGaugeSize Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateGaugeSize(id As Integer, type As String, size As String, groupCode As String, active As Integer, sort As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE gauge_categorycontrol SET InstrumentType=@type, Size=@size, GroupCode=@group, Active=@active, Sort=@sort WHERE ID=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@group", groupCode)
                    cmd.Parameters.AddWithValue("@active", active)
                    cmd.Parameters.AddWithValue("@sort", sort)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateGaugeSize Error: " & ex.Message)
        End Try
        Return False
    End Function
    Public Sub InitializeDepartmentListTable()
        Try
            If MySQLDBConnect() = 1 Then
                Dim createSql As String = "CREATE TABLE IF NOT EXISTS department_list (" &
                    "id INT AUTO_INCREMENT PRIMARY KEY, " &
                    "Department VARCHAR(255), " &
                    "InstrumentName VARCHAR(255), " &
                    "SizeandRange VARCHAR(255), " &
                    "`Control No` VARCHAR(100), " &
                    "Color VARCHAR(100), " &
                    "Status VARCHAR(255), " &
                    "Remarks TEXT, " &
                    "ImportBatch VARCHAR(100)" &
                    ")"
                Using cmd As New MySqlCommand(createSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                Try
                    Using cmdCol As New MySqlCommand("ALTER TABLE department_list ADD COLUMN Status VARCHAR(255)", _con)
                        cmdCol.ExecuteNonQuery()
                    End Using
                Catch ex2 As Exception
                End Try

                Try
                    Using cmdCol As New MySqlCommand("ALTER TABLE department_list ADD COLUMN Remarks TEXT", _con)
                        cmdCol.ExecuteNonQuery()
                    End Using
                Catch ex2 As Exception
                End Try

                Try
                    Using cmdCol As New MySqlCommand("ALTER TABLE department_list ADD COLUMN ImportBatch VARCHAR(100)", _con)
                        cmdCol.ExecuteNonQuery()
                    End Using
                Catch ex2 As Exception
                End Try

                ' Add CycleName column if not exists
                Try
                    Using cmdCol As New MySqlCommand("ALTER TABLE department_list ADD COLUMN CycleName VARCHAR(255)", _con)
                        cmdCol.ExecuteNonQuery()
                    End Using
                Catch ex2 As Exception
                End Try

                ' Optimization: Add index to Control No if not exists
                Try
                    Using cmdIdx As New MySqlCommand("CREATE INDEX idx_control_no ON department_list (`Control No`)", _con)
                        cmdIdx.ExecuteNonQuery()
                    End Using
                Catch ex3 As Exception
                    ' Ignore if index already exists
                End Try

                ' Backfill: Assign earliest cycle to existing rows without CycleName
                Try
                    Dim backfillSql = "UPDATE department_list SET CycleName = (SELECT MIN(CycleName) FROM interchangeability WHERE CycleName IS NOT NULL AND CycleName != '') WHERE CycleName IS NULL OR CycleName = ''"
                    Using cmdBackfill As New MySqlCommand(backfillSql, _con)
                        cmdBackfill.ExecuteNonQuery()
                    End Using
                Catch ex4 As Exception
                End Try
            End If
        Catch ex As Exception
            Console.WriteLine("InitDeptList Error: " & ex.Message)
        End Try
    End Sub

    Public Function InsertDeptListItem(dept As String, name As String, size As String, ctrl As String, color As String, status As String, remarks As String, batch As String, Optional cycleName As String = "") As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                ' Check for duplicate Control No + CycleName in department_list
                If Not String.IsNullOrEmpty(ctrl) Then
                    Dim checkQuery As String = "SELECT COUNT(*) FROM department_list WHERE `Control No` = @ctrl AND CycleName = @cycle"
                    Using checkCmd As New MySqlCommand(checkQuery, _con)
                        checkCmd.Parameters.AddWithValue("@ctrl", ctrl)
                        checkCmd.Parameters.AddWithValue("@cycle", cycleName)
                        Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If count > 0 Then Return False ' Duplicate found, skip
                    End Using
                End If

                Dim query As String = "INSERT INTO department_list (Department, InstrumentName, SizeandRange, `Control No`, Color, Status, Remarks, ImportBatch, CycleName) " &
                                    "VALUES (@dept, @name, @size, @ctrl, @color, @status, @rem, @batch, @cycle)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@dept", dept)
                    cmd.Parameters.AddWithValue("@name", name)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@ctrl", ctrl)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@rem", remarks)
                    cmd.Parameters.AddWithValue("@batch", batch)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertDeptListItem Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetExistingControlNos() As HashSet(Of String)
        Dim existingCtrls As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            If MySQLDBConnect() = 1 Then
                Dim dt = ReadDatatable("SELECT `Control No` FROM department_list WHERE `Control No` IS NOT NULL AND `Control No` != ''")
                For Each row As DataRow In dt.Rows
                    existingCtrls.Add(row(0).ToString().Trim())
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("GetExistingControlNos Error: " & ex.Message)
        End Try
        Return existingCtrls
    End Function

    Public Function BulkInsertDeptListItems(items As List(Of Dictionary(Of String, Object)), batchTag As String, Optional cycleName As String = "") As Integer
        If items Is Nothing OrElse items.Count = 0 Then Return 0
        Dim count As Integer = 0
        Try
            If MySQLDBConnect() = 1 Then
                Using transaction = _con.BeginTransaction()
                    Try
                        Dim query As String = "INSERT INTO department_list (Department, InstrumentName, SizeandRange, `Control No`, Color, Status, Remarks, ImportBatch, CycleName) " &
                                            "VALUES (@dept, @name, @size, @ctrl, @color, @status, @rem, @batch, @cycle)"
                        Using cmd As New MySqlCommand(query, _con, transaction)
                            cmd.Parameters.Add("@dept", MySqlDbType.VarChar)
                            cmd.Parameters.Add("@name", MySqlDbType.VarChar)
                            cmd.Parameters.Add("@size", MySqlDbType.VarChar)
                            cmd.Parameters.Add("@ctrl", MySqlDbType.VarChar)
                            cmd.Parameters.Add("@color", MySqlDbType.VarChar)
                            cmd.Parameters.Add("@status", MySqlDbType.VarChar)
                            cmd.Parameters.Add("@rem", MySqlDbType.Text)
                            cmd.Parameters.Add("@batch", MySqlDbType.VarChar)
                            cmd.Parameters.Add("@cycle", MySqlDbType.VarChar)

                            For Each item In items
                                cmd.Parameters("@dept").Value = item("dept")
                                cmd.Parameters("@name").Value = item("name")
                                cmd.Parameters("@size").Value = item("size")
                                cmd.Parameters("@ctrl").Value = item("ctrl")
                                cmd.Parameters("@color").Value = item("color")
                                cmd.Parameters("@status").Value = item("status")
                                cmd.Parameters("@rem").Value = item("remarks")
                                cmd.Parameters("@batch").Value = batchTag
                                cmd.Parameters("@cycle").Value = cycleName
                                cmd.ExecuteNonQuery()
                                count += 1
                            Next
                        End Using
                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Throw
                    End Try
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("BulkInsertDeptListItems Error: " & ex.Message)
            Throw
        End Try
        Return count
    End Function
    Public Function SyncDepartmentListFromCards(reportProgress As Action(Of Double)) As Integer
        Dim importedCount As Integer = 0
        Try
            If MySQLDBConnect() = 1 Then
                ' Get allowed departments from departmentmaster
                Dim allowedDepts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim dtAllowed = ReadDatatable("SELECT DepartmentName FROM departmentmaster")
                For Each row As DataRow In dtAllowed.Rows
                    allowedDepts.Add(row("DepartmentName").ToString().Trim())
                Next

                Dim typeRows = ReadDatatable("SELECT TypeName, Category FROM type_details")
                Dim totalTables = typeRows.Rows.Count
                Dim processedTables = 0
                Dim batchTag As String = "Card Sync " & DateTime.Now.ToString("yyyyMMdd")

                ' Optimize: Get existing Control Nos in a HashSet
                Dim existingCtrls As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Try
                    Dim dtExisting = ReadDatatable("SELECT `Control No` FROM department_list")
                    For Each row As DataRow In dtExisting.Rows
                        Dim c = row(0).ToString().Trim()
                        If Not String.IsNullOrEmpty(c) Then existingCtrls.Add(c)
                    Next
                Catch ex As Exception
                    Console.WriteLine("Error fetching existing controls: " & ex.Message)
                End Try

                For Each typeRow As DataRow In typeRows.Rows
                    Dim typeName = typeRow("TypeName").ToString()
                    Dim category = typeRow("Category").ToString()
                    Dim tbl = TypeNameToTableName(typeName)

                    ' Safety: Check if table actually exists
                    Dim checkTableDt = ReadDatatable($"SHOW TABLES LIKE '{tbl}'")
                    If checkTableDt.Rows.Count > 0 Then
                        Dim dtCards As DataTable = Nothing

                        Try
                            ' Attempt full query first
                            Dim selectQuery As String = ""
                            If category = "Instrument" Then
                                selectQuery = $"SELECT Location, InstrumentName, Size, ControlNo, Color FROM `{tbl}`"
                            Else ' Gauge
                                selectQuery = $"SELECT Location, GaugeName AS InstrumentName, Size, ControlNo, Color FROM `{tbl}`"
                            End If
                            dtCards = ReadDatatable(selectQuery)
                        Catch ex As Exception
                            ' Fallback: Some columns might be missing, try a safer approach
                            Try
                                Dim fallbackQuery = $"SELECT * FROM `{tbl}`"
                                dtCards = ReadDatatable(fallbackQuery)
                            Catch fallbackEx As Exception
                                Console.WriteLine($"Could not read table {tbl}: " & fallbackEx.Message)
                            End Try
                        End Try

                        If dtCards IsNot Nothing AndAlso dtCards.Rows.Count > 0 Then
                            For Each cardRow As DataRow In dtCards.Rows
                                ' Extra safety: check if column actually exists in the returned DataTable
                                Dim ctrl = ""
                                If dtCards.Columns.Contains("ControlNo") AndAlso Not IsDBNull(cardRow("ControlNo")) Then
                                    ctrl = cardRow("ControlNo").ToString().Trim()
                                End If

                                If String.IsNullOrEmpty(ctrl) OrElse existingCtrls.Contains(ctrl) Then Continue For

                                Dim dept = ""
                                If dtCards.Columns.Contains("Location") AndAlso Not IsDBNull(cardRow("Location")) Then
                                    dept = cardRow("Location").ToString().Trim()
                                End If

                                ' Only sync if department is allowed
                                If Not allowedDepts.Contains(dept) Then Continue For

                                Dim name = ""
                                If dtCards.Columns.Contains("InstrumentName") AndAlso Not IsDBNull(cardRow("InstrumentName")) Then
                                    name = cardRow("InstrumentName").ToString()
                                ElseIf dtCards.Columns.Contains("GaugeName") AndAlso Not IsDBNull(cardRow("GaugeName")) Then
                                    name = cardRow("GaugeName").ToString()
                                End If

                                Dim size = ""
                                If dtCards.Columns.Contains("Size") AndAlso Not IsDBNull(cardRow("Size")) Then size = cardRow("Size").ToString()

                                Dim color = ""
                                If dtCards.Columns.Contains("Color") AndAlso Not IsDBNull(cardRow("Color")) Then color = cardRow("Color").ToString()

                                Dim status = ""
                                If dtCards.Columns.Contains("Status") AndAlso Not IsDBNull(cardRow("Status")) Then status = cardRow("Status").ToString()

                                Dim rems = ""
                                If dtCards.Columns.Contains("Remark") AndAlso Not IsDBNull(cardRow("Remark")) Then rems = cardRow("Remark").ToString()
                                If dtCards.Columns.Contains("Remarks") AndAlso Not IsDBNull(cardRow("Remarks")) Then rems = cardRow("Remarks").ToString()

                                If InsertDeptListItem(dept, name, size, ctrl, color, status, rems, batchTag) Then
                                    importedCount += 1
                                    existingCtrls.Add(ctrl)
                                End If
                            Next
                        End If
                    End If

                    processedTables += 1
                    If reportProgress IsNot Nothing Then
                        Try
                            reportProgress((processedTables / totalTables) * 100)
                        Catch pEx As Exception
                            Console.WriteLine("Progress reporting error: " & pEx.Message)
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("SyncDepartmentListFromCards Error: " & ex.Message)
        End Try
        Return importedCount
    End Function

    Public Sub InitializeInterchangeTables()
        Try
            If MySQLDBConnect() = 1 Then
                ' interchangeability
                ' 1. Create table if not exists (including unique key)
                Dim createInterchangeSql As String = "CREATE TABLE IF NOT EXISTS interchangeability (" &
                    "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                    "CycleName VARCHAR(255), " &
                    "ControlNo VARCHAR(255), " &
                    "Department VARCHAR(255), " &
                    "InstrumentName VARCHAR(255), " &
                    "SizeandRange VARCHAR(255), " &
                    "Color VARCHAR(255), " &
                    "Status VARCHAR(255), " &
                    "ActionDate DATE, " &
                    "ActionTime TIME, " &
                    "Remarks TEXT, " &
                    "ActionBy VARCHAR(255), " &
                    "RFID_tag VARCHAR(255), " &
                    "UNIQUE KEY uk_cycle_ctrl (CycleName, ControlNo)" &
                    ")"
                Using cmd As New MySqlCommand(createInterchangeSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                ' 2. Migration: If table existed but unique key didn't, we might need to deduplicate and add it
                Try
                    ' Deduplicate: Keep the row with the latest ID for each (CycleName, ControlNo)
                    Dim dedupeSql As String = "DELETE i1 FROM interchangeability i1 " &
                                             "INNER JOIN interchangeability i2 " &
                                             "WHERE i1.ID < i2.ID " &
                                             "AND i1.CycleName = i2.CycleName " &
                                             "AND i1.ControlNo = i2.ControlNo"
                    Using cmdDedupe As New MySqlCommand(dedupeSql, _con)
                        cmdDedupe.ExecuteNonQuery()
                    End Using

                    ' Check if the unique key exists
                    Dim checkKeySql As String = "SHOW INDEX FROM interchangeability WHERE Key_name = 'uk_cycle_ctrl'"
                    Dim dtKey = ReadDatatable(checkKeySql)
                    If dtKey.Rows.Count = 0 Then
                        ' Add the key if missing
                        Try
                            Dim addKeySql As String = "ALTER TABLE interchangeability ADD UNIQUE KEY uk_cycle_ctrl (CycleName, ControlNo)"
                            Using cmdAddKey As New MySqlCommand(addKeySql, _con)
                                cmdAddKey.ExecuteNonQuery()
                            End Using
                        Catch keyEx As Exception
                            Console.WriteLine("Warning: Could not add unique key (might be duplicate values still): " & keyEx.Message)
                        End Try
                    End If
                Catch migEx As Exception
                    Console.WriteLine("Migration Error (interchangeability): " & migEx.Message)
                End Try

                ' 2.1 Ensure RFID_tag column exists in interchangeability (Migration)
                Try
                    If Not ColumnExists("interchangeability", "RFID_tag") Then
                        Using cmdCol As New MySqlCommand("ALTER TABLE interchangeability ADD COLUMN RFID_tag VARCHAR(255) AFTER ActionBy", _con)
                            cmdCol.ExecuteNonQuery()
                        End Using
                    End If
                Catch exRfid As Exception
                    Console.WriteLine("Migration Error (interchangeability RFID_tag): " & exRfid.Message)
                End Try

                ' system_config
                Dim createConfigSql As String = "CREATE TABLE IF NOT EXISTS system_config (" &
                    "ConfigKey VARCHAR(255) PRIMARY KEY, " &
                    "ConfigValue TEXT" &
                    ")"
                Using cmd As New MySqlCommand(createConfigSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                ' temporary_issuance
                Dim createTempIssuanceSql As String = "CREATE TABLE IF NOT EXISTS temporary_issuance (" &
                    "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                    "ControlNo VARCHAR(255), " &
                    "IssueDate DATE, " &
                    "Time TIME, " &
                    "CycleName VARCHAR(255), " &
                    "Department VARCHAR(255), " &
                    "InstrumentName VARCHAR(255), " &
                    "Color VARCHAR(255), " &
                    "Reason VARCHAR(255), " &
                    "DocumentPath VARCHAR(255), " &
                    "IssuedBy VARCHAR(255), " &
                    "IsReturned TINYINT(1) DEFAULT 0" &
                    ")"
                Using cmd As New MySqlCommand(createTempIssuanceSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                ' wop
                Dim createWopSql As String = "CREATE TABLE IF NOT EXISTS wop (" &
                    "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                    "ControlNo VARCHAR(255), " &
                    "WOPDate DATE, " &
                    "Time TIME, " &
                    "CycleName VARCHAR(255), " &
                    "Department VARCHAR(255), " &
                    "InstrumentName VARCHAR(255), " &
                    "Color VARCHAR(255), " &
                    "ReportedReason VARCHAR(255), " &
                    "ReportedBy VARCHAR(255), " &
                    "Resolution VARCHAR(100) DEFAULT ''" &
                    ")"
                Using cmd As New MySqlCommand(createWopSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                ' issue
                Dim createIssueSql As String = "CREATE TABLE IF NOT EXISTS issue (" &
                    "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                    "ControlNo VARCHAR(255), " &
                    "IssueDate DATE, " &
                    "Time TIME, " &
                    "CycleName VARCHAR(255), " &
                    "Department VARCHAR(255), " &
                    "InstrumentName VARCHAR(255), " &
                    "Color VARCHAR(255), " &
                    "Remarks TEXT, " &
                    "IssuedBy VARCHAR(255), " &
                    "RFID_tag VARCHAR(255)" &
                    ")"
                Using cmd As New MySqlCommand(createIssueSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                ' receive
                Dim createReceiveSql As String = "CREATE TABLE IF NOT EXISTS receive (" &
                    "ID INT AUTO_INCREMENT PRIMARY KEY, " &
                    "ControlNo VARCHAR(255), " &
                    "ReceiveDate DATE, " &
                    "Time TIME, " &
                    "CycleName VARCHAR(255), " &
                    "Department VARCHAR(255), " &
                    "InstrumentName VARCHAR(255), " &
                    "Color VARCHAR(255), " &
                    "Remarks TEXT, " &
                    "ReceivedBy VARCHAR(255), " &
                    "RFID_tag VARCHAR(255)" &
                    ")"
                Using cmd As New MySqlCommand(createReceiveSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                ' 6. writeoff column maintenance
                Try
                    If Not ColumnExists("writeoff", "CycleName") Then
                        Using cmdCol As New MySqlCommand("ALTER TABLE writeoff ADD COLUMN CycleName VARCHAR(255) AFTER Time", _con)
                            cmdCol.ExecuteNonQuery()
                        End Using
                    End If
                Catch exW As Exception
                End Try

                ' 7. reintroduction column maintenance
                Try
                    If Not ColumnExists("reintroduction", "CycleName") Then
                        Using cmdCol As New MySqlCommand("ALTER TABLE reintroduction ADD COLUMN CycleName VARCHAR(255) AFTER Time", _con)
                            cmdCol.ExecuteNonQuery()
                        End Using
                    End If
                Catch exR As Exception
                End Try

                ' rfid_exclusions
                Dim createExclusionsSql As String = "CREATE TABLE IF NOT EXISTS rfid_exclusions (" &
                    "id INT AUTO_INCREMENT PRIMARY KEY, " &
                    "rfid_tag VARCHAR(100), " &
                    "added_at DATETIME DEFAULT CURRENT_TIMESTAMP" &
                    ")"
                Using cmd As New MySqlCommand(createExclusionsSql, _con)
                    cmd.ExecuteNonQuery()
                End Using

                ' Update form name in calibration_template and calibrationmapping tables
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET form_name = 'Passameter60' WHERE form_name = 'Passameter'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template form_name): " & ex.Message)
                End Try

                Try
                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'Passameter60' WHERE form_name = 'Passameter'"
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping form_name): " & ex.Message)
                End Try

                ' Update mapping for HeightMastergaugeCalibration
                Try
                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'HeightMastergaugeCalibration' WHERE form_name = 'HeightMasterGauge_cal' OR form_name = 'heightMastergaugeCalibration'"
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping heightMaster): " & ex.Message)
                End Try

                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET form_name = 'HeightMastergaugeCalibration' WHERE form_name = 'HeightMasterGauge_cal' OR form_name = 'heightMastergaugeCalibration'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template heightMaster): " & ex.Message)
                End Try

                ' Update or insert KeyGrooveGaugeCalibration in calibration_template
                Try
                    Dim checkTemplate = ReadDatatable("SELECT 1 FROM calibration_template WHERE form_name = 'KeyGrooveGaugeCalibration' LIMIT 1")
                    If checkTemplate.Rows.Count = 0 Then
                        Dim updateOldTemplate As String = "UPDATE calibration_template SET form_name = 'KeyGrooveGaugeCalibration' WHERE form_name = 'KeyGrooveGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateOldTemplate, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertTemplateSql As String = "INSERT INTO calibration_template (type_name, calibration_category, form_name, created_at) VALUES ('Key Groove Gauge', 'All Ranges', 'KeyGrooveGaugeCalibration', NOW())"
                            Using cmd As New MySqlCommand(insertTemplateSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template KeyGroove): " & ex.Message)
                End Try

                ' Update or insert KeyGrooveGaugeCalibration in calibrationmapping
                Try
                    Dim checkMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE form_name = 'KeyGrooveGaugeCalibration' LIMIT 1")
                    If checkMapping.Rows.Count = 0 Then
                        Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'KeyGrooveGaugeCalibration' WHERE form_name = 'KeyGrooveGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateMappingSql, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Key Groove Gauge', 'Gauge', 'KeyGrooveGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping KeyGroove): " & ex.Message)
                End Try

                ' Update or insert PlainPlugGaugeCalibration in calibration_template
                Try
                    Dim checkTemplate = ReadDatatable("SELECT 1 FROM calibration_template WHERE form_name = 'PlainPlugGaugeCalibration' LIMIT 1")
                    If checkTemplate.Rows.Count = 0 Then
                        Dim updateOldTemplate As String = "UPDATE calibration_template SET form_name = 'PlainPlugGaugeCalibration' WHERE form_name = 'PlainPlugGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateOldTemplate, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertTemplateSql As String = "INSERT INTO calibration_template (type_name, calibration_category, form_name, created_at) VALUES ('Plain Plug Gauge', 'All Ranges', 'PlainPlugGaugeCalibration', NOW())"
                            Using cmd As New MySqlCommand(insertTemplateSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template PlainPlug): " & ex.Message)
                End Try

                ' Update or insert PlainPlugGaugeCalibration in calibrationmapping
                Try
                    Dim checkMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE form_name = 'PlainPlugGaugeCalibration' LIMIT 1")
                    If checkMapping.Rows.Count = 0 Then
                        Dim updateMappingSql = "UPDATE calibrationmapping SET form_name = 'PlainPlugGaugeCalibration' WHERE form_name = 'PlainPlugGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateMappingSql, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Plain Plug Gauge', 'Gauge', 'PlainPlugGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping PlainPlug): " & ex.Message)
                End Try

                ' Create feeler_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS feeler_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "MU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Measurement_Range VARCHAR(255), " &
                        "Min_Limit DECIMAL(10,4), " &
                        "Max_Limit DECIMAL(10,4), " &
                        "Obs_1 TEXT, " &
                        "Obs_2 TEXT, " &
                        "Obs_3 TEXT, " &
                        "Camber_Width_Obs_1 TEXT, " &
                        "Camber_Width_Obs_2 TEXT, " &
                        "Camber_Width_Obs_3 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT, " &
                        "CreatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating feeler_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for feeler_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM feeler_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO feeler_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_DEFAULT', 'All Sizes', '0.01', " &
                            "18.00, 22.00, 40.00, 60.00" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Feeler Gauge MASTER row: " & ex.Message)
                End Try

                ' Update or insert FeelerGaugeCalibration in calibration_template
                Try
                    Dim checkTemplate = ReadDatatable("SELECT 1 FROM calibration_template WHERE form_name = 'FeelerGaugeCalibration' LIMIT 1")
                    If checkTemplate.Rows.Count = 0 Then
                        Dim updateOldTemplate As String = "UPDATE calibration_template SET form_name = 'FeelerGaugeCalibration' WHERE form_name = 'FeelerGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateOldTemplate, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertTemplateSql As String = "INSERT INTO calibration_template (type_name, calibration_category, form_name, created_at) VALUES ('Feeler Gauge', 'All Ranges', 'FeelerGaugeCalibration', NOW())"
                            Using cmd As New MySqlCommand(insertTemplateSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template FeelerGauge): " & ex.Message)
                End Try

                ' Update or insert FeelerGaugeCalibration in calibrationmapping
                Try
                    Dim checkMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE form_name = 'FeelerGaugeCalibration' LIMIT 1")
                    If checkMapping.Rows.Count = 0 Then
                        Dim updateMappingSql = "UPDATE calibrationmapping SET form_name = 'FeelerGaugeCalibration' WHERE form_name = 'FeelerGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateMappingSql, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Feeler Gauge', 'Gauge', 'FeelerGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping FeelerGauge): " & ex.Message)
                End Try

                ' Create radius_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS radius_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Measurement_Range VARCHAR(255), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Min_Limit DECIMAL(10,4), " &
                        "Max_Limit DECIMAL(10,4), " &
                        "Inside_Radius_Obs_1 TEXT, " &
                        "Inside_Radius_Obs_2 TEXT, " &
                        "Inside_Radius_Obs_3 TEXT, " &
                        "Outside_Radius_Obs_1 TEXT, " &
                        "Outside_Radius_Obs_2 TEXT, " &
                        "Outside_Radius_Obs_3 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT, " &
                        "CreatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating radius_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for radius_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM radius_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO radius_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.001', 'All Sizes', '0.001', " &
                            "18.00, 22.00, 40.00, 60.00" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Radius Gauge MASTER row: " & ex.Message)
                End Try

                ' Update or insert RadiusGaugeCalibration in calibration_template
                Try
                    Dim checkTemplate = ReadDatatable("SELECT 1 FROM calibration_template WHERE form_name = 'RadiusGaugeCalibration' LIMIT 1")
                    If checkTemplate.Rows.Count = 0 Then
                        Dim updateOldTemplate As String = "UPDATE calibration_template SET form_name = 'RadiusGaugeCalibration' WHERE form_name = 'RadiusGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateOldTemplate, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertTemplateSql As String = "INSERT INTO calibration_template (type_name, calibration_category, form_name, created_at) VALUES ('Radius Gauge', 'All Ranges', 'RadiusGaugeCalibration', NOW())"
                            Using cmd As New MySqlCommand(insertTemplateSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template RadiusGauge): " & ex.Message)
                End Try

                ' Update or insert RadiusGaugeCalibration in calibrationmapping
                Try
                    Dim checkMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE form_name = 'RadiusGaugeCalibration' LIMIT 1")
                    If checkMapping.Rows.Count = 0 Then
                        Dim updateMappingSql = "UPDATE calibrationmapping SET form_name = 'RadiusGaugeCalibration' WHERE form_name = 'RadiusGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateMappingSql, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Radius Gauge', 'Gauge', 'RadiusGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping RadiusGauge): " & ex.Message)
                End Try

                ' Create thread_pitch_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS thread_pitch_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Measurement_Range VARCHAR(255), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Min_Limit DECIMAL(10,4), " &
                        "Max_Limit DECIMAL(10,4), " &
                        "Pitch_Obs_1 TEXT, " &
                        "Pitch_Obs_2 TEXT, " &
                        "Pitch_Obs_3 TEXT, " &
                        "Angle_Obs_1 TEXT, " &
                        "Angle_Obs_2 TEXT, " &
                        "Angle_Obs_3 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT, " &
                        "CreatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating thread_pitch_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for thread_pitch_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM thread_pitch_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO thread_pitch_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.001', 'All Sizes', '0.001', " &
                            "18.00, 22.00, 40.00, 60.00" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Thread Pitch Gauge MASTER row: " & ex.Message)
                End Try

                ' Update or insert ThreadPitchGaugeCalibration in calibration_template
                Try
                    Dim checkTemplate = ReadDatatable("SELECT 1 FROM calibration_template WHERE form_name = 'ThreadPitchGaugeCalibration' LIMIT 1")
                    If checkTemplate.Rows.Count = 0 Then
                        Dim updateOldTemplate As String = "UPDATE calibration_template SET form_name = 'ThreadPitchGaugeCalibration' WHERE form_name = 'ThreadPitchGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateOldTemplate, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertTemplateSql As String = "INSERT INTO calibration_template (type_name, calibration_category, form_name, created_at) VALUES ('Thread Pitch Gauge', 'All Ranges', 'ThreadPitchGaugeCalibration', NOW())"
                            Using cmd As New MySqlCommand(insertTemplateSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template ThreadPitchGauge): " & ex.Message)
                End Try

                ' Update or insert ThreadPitchGaugeCalibration in calibrationmapping
                Try
                    Dim checkMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE form_name = 'ThreadPitchGaugeCalibration' LIMIT 1")
                    If checkMapping.Rows.Count = 0 Then
                        Dim updateMappingSql = "UPDATE calibrationmapping SET form_name = 'ThreadPitchGaugeCalibration' WHERE form_name = 'ThreadPitchGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateMappingSql, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Thread Pitch Gauge', 'Gauge', 'ThreadPitchGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping ThreadPitchGauge): " & ex.Message)
                End Try

                ' Create plain_plug_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS plain_plug_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Go_Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Go_Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Go_Obs_1 TEXT, " &
                        "Go_Obs_2 TEXT, " &
                        "Go_Obs_3 TEXT, " &
                        "Go_Obs_4 TEXT, " &
                        "NoGo_Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "NoGo_Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "NoGo_Obs_1 TEXT, " &
                        "NoGo_Obs_2 TEXT, " &
                        "NoGo_Obs_3 TEXT, " &
                        "NoGo_Obs_4 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating plain_plug_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for plain_plug_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM plain_plug_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO plain_plug_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Go_Min_Limit, Go_Max_Limit, " &
                            "NoGo_Min_Limit, NoGo_Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.000001', 'All Sizes', '0.000001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL, " &
                            "NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Plain Plug MASTER row: " & ex.Message)
                End Try

                ' Update or insert PlainRingGaugeCalibration in calibration_template
                Try
                    Dim checkTemplate = ReadDatatable("SELECT 1 FROM calibration_template WHERE form_name = 'PlainRingGaugeCalibration' LIMIT 1")
                    If checkTemplate.Rows.Count = 0 Then
                        Dim updateOldTemplate As String = "UPDATE calibration_template SET form_name = 'PlainRingGaugeCalibration' WHERE form_name = 'PlainRingGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateOldTemplate, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertTemplateSql As String = "INSERT INTO calibration_template (type_name, calibration_category, form_name, created_at) VALUES ('Plain Ring Gauge', 'All Ranges', 'PlainRingGaugeCalibration', NOW())"
                            Using cmd As New MySqlCommand(insertTemplateSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template PlainRing): " & ex.Message)
                End Try

                ' Update or insert PlainRingGaugeCalibration in calibrationmapping
                Try
                    Dim checkMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE form_name = 'PlainRingGaugeCalibration' LIMIT 1")
                    If checkMapping.Rows.Count = 0 Then
                        Dim updateMappingSql = "UPDATE calibrationmapping SET form_name = 'PlainRingGaugeCalibration' WHERE form_name = 'PlainRingGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateMappingSql, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Plain Ring Gauge', 'Gauge', 'PlainRingGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping PlainRing): " & ex.Message)
                End Try

                ' Create plain_ring_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS plain_ring_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Obs_1 TEXT, " &
                        "Obs_2 TEXT, " &
                        "Obs_3 TEXT, " &
                        "Obs_4 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating plain_ring_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for plain_ring_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM plain_ring_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO plain_ring_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Min_Limit, Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.000001', 'All Sizes', '0.000001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Plain Ring MASTER row: " & ex.Message)
                End Try

                ' Update or insert SnapGaugeCalibration in calibration_template
                Try
                    Dim checkTemplate = ReadDatatable("SELECT 1 FROM calibration_template WHERE form_name = 'SnapGaugeCalibration' LIMIT 1")
                    If checkTemplate.Rows.Count = 0 Then
                        Dim updateOldTemplate As String = "UPDATE calibration_template SET form_name = 'SnapGaugeCalibration' WHERE form_name = 'SnapGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateOldTemplate, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertTemplateSql As String = "INSERT INTO calibration_template (type_name, calibration_category, form_name, created_at) VALUES ('Snap Gauge', 'All Ranges', 'SnapGaugeCalibration', NOW())"
                            Using cmd As New MySqlCommand(insertTemplateSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibration_template SnapGauge): " & ex.Message)
                End Try

                ' Update or insert SnapGaugeCalibration in calibrationmapping
                Try
                    Dim checkMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE form_name = 'SnapGaugeCalibration' LIMIT 1")
                    If checkMapping.Rows.Count = 0 Then
                        Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'SnapGaugeCalibration' WHERE form_name = 'SnapGauge_cal'"
                        Dim updatedCount As Integer = 0
                        Using cmd As New MySqlCommand(updateMappingSql, _con)
                            updatedCount = cmd.ExecuteNonQuery()
                        End Using

                        If updatedCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Snap Gauge', 'Gauge', 'SnapGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping SnapGauge): " & ex.Message)
                End Try

                ' Create snap_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS snap_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Go_Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Go_Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Go_Obs_1 TEXT, " &
                        "Go_Obs_2 TEXT, " &
                        "Go_Obs_3 TEXT, " &
                        "Go_Obs_4 TEXT, " &
                        "NoGo_Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "NoGo_Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "NoGo_Obs_1 TEXT, " &
                        "NoGo_Obs_2 TEXT, " &
                        "NoGo_Obs_3 TEXT, " &
                        "NoGo_Obs_4 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating snap_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for snap_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM snap_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO snap_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Go_Min_Limit, Go_Max_Limit, " &
                            "NoGo_Min_Limit, NoGo_Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.000001', 'All Sizes', '0.000001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL, " &
                            "NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Snap Gauge MASTER row: " & ex.Message)
                End Try

                ' Migrate legacy config and mapping names for Special Height Gauge
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET FormName = 'SpecialHeightGaugeCalibration' WHERE FormName = 'SpecialHeightGauge_cal'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'SpecialHeightGaugeCalibration' WHERE form_name = 'SpecialHeightGauge_cal'"
                    Dim updatedCount As Integer = 0
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        updatedCount = cmd.ExecuteNonQuery()
                    End Using

                    If updatedCount = 0 Then
                        Dim checkMappingSql As String = "SELECT COUNT(*) FROM calibrationmapping WHERE type_name = 'Special Height Gauge'"
                        Dim mappingCount As Integer = 0
                        Using cmd As New MySqlCommand(checkMappingSql, _con)
                            mappingCount = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using
                        If mappingCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Special Height Gauge', 'Gauge', 'SpecialHeightGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        Else
                            Dim updateExistingSql As String = "UPDATE calibrationmapping SET form_name = 'SpecialHeightGaugeCalibration' WHERE type_name = 'Special Height Gauge'"
                            Using cmd As New MySqlCommand(updateExistingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping SpecialHeight): " & ex.Message)
                End Try

                ' Create special_height_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS special_height_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Size_1_Nominal VARCHAR(100), " &
                        "Size_1_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Size_1_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Size_1_Obs_1 TEXT, " &
                        "Size_1_Obs_2 TEXT, " &
                        "Size_1_Obs_3 TEXT, " &
                        "Size_1_Obs_4 TEXT, " &
                        "Size_2_Nominal VARCHAR(100), " &
                        "Size_2_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Size_2_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Size_2_Obs_1 TEXT, " &
                        "Size_2_Obs_2 TEXT, " &
                        "Size_2_Obs_3 TEXT, " &
                        "Size_2_Obs_4 TEXT, " &
                        "Size_3_Nominal VARCHAR(100), " &
                        "Size_3_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Size_3_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Size_3_Obs_1 TEXT, " &
                        "Size_3_Obs_2 TEXT, " &
                        "Size_3_Obs_3 TEXT, " &
                        "Size_3_Obs_4 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating special_height_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for special_height_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM special_height_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO special_height_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Size_1_Min_Limit, Size_1_Max_Limit, " &
                            "Size_2_Min_Limit, Size_2_Max_Limit, " &
                            "Size_3_Min_Limit, Size_3_Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.0001', 'Special', '0.0001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL, " &
                            "NULL, NULL, " &
                            "NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Special Height MASTER row: " & ex.Message)
                End Try

                ' Migrate legacy config and mapping names for Spline Ring Gauge
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET FormName = 'SplineRingGaugeCalibration' WHERE FormName = 'SplineRingGauge_cal'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'SplineRingGaugeCalibration' WHERE form_name = 'SplineRingGauge_cal'"
                    Dim updatedCount As Integer = 0
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        updatedCount = cmd.ExecuteNonQuery()
                    End Using

                    If updatedCount = 0 Then
                        Dim checkMappingSql As String = "SELECT COUNT(*) FROM calibrationmapping WHERE type_name = 'Spline Ring Gauge'"
                        Dim mappingCount As Integer = 0
                        Using cmd As New MySqlCommand(checkMappingSql, _con)
                            mappingCount = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using
                        If mappingCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Spline Ring Gauge', 'Gauge', 'SplineRingGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        Else
                            Dim updateExistingSql As String = "UPDATE calibrationmapping SET form_name = 'SplineRingGaugeCalibration' WHERE type_name = 'Spline Ring Gauge'"
                            Using cmd As New MySqlCommand(updateExistingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping SplineRing): " & ex.Message)
                End Try

                ' Create spline_ring_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS spline_ring_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Minor_Dia_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Minor_Dia_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Minor_Dia_Obs_1 TEXT, " &
                        "Minor_Dia_Obs_2 TEXT, " &
                        "Minor_Dia_Obs_3 TEXT, " &
                        "Major_Dia_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Major_Dia_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Major_Dia_Obs_1 TEXT, " &
                        "Major_Dia_Obs_2 TEXT, " &
                        "Major_Dia_Obs_3 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating spline_ring_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for spline_ring_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM spline_ring_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO spline_ring_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Minor_Dia_Min_Limit, Minor_Dia_Max_Limit, " &
                            "Major_Dia_Min_Limit, Major_Dia_Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.001', 'All Sizes', '0.001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL, " &
                            "NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Spline Ring MASTER row: " & ex.Message)
                End Try

                ' Migrate legacy config and mapping names for Straight Plug Gauge
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET FormName = 'StraightPlugGaugeCalibration' WHERE FormName = 'StraightPlugGauge_cal'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'StraightPlugGaugeCalibration' WHERE form_name = 'StraightPlugGauge_cal'"
                    Dim updatedCount As Integer = 0
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        updatedCount = cmd.ExecuteNonQuery()
                    End Using

                    If updatedCount = 0 Then
                        Dim checkMappingSql As String = "SELECT COUNT(*) FROM calibrationmapping WHERE type_name = 'Straight Plug Gauge'"
                        Dim mappingCount As Integer = 0
                        Using cmd As New MySqlCommand(checkMappingSql, _con)
                            mappingCount = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using
                        If mappingCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Straight Plug Gauge', 'Gauge', 'StraightPlugGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        Else
                            Dim updateExistingSql As String = "UPDATE calibrationmapping SET form_name = 'StraightPlugGaugeCalibration' WHERE type_name = 'Straight Plug Gauge'"
                            Using cmd As New MySqlCommand(updateExistingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping StraightPlug): " & ex.Message)
                End Try

                ' Create straight_plug_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS straight_plug_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Obs_1 TEXT, " &
                        "Obs_2 TEXT, " &
                        "Obs_3 TEXT, " &
                        "Obs_4 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating straight_plug_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for straight_plug_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM straight_plug_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO straight_plug_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Min_Limit, Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.000001', 'All Sizes', '0.000001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Straight Plug MASTER row: " & ex.Message)
                End Try

                ' Migrate legacy config and mapping names for Thread Ring Gauge
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET FormName = 'ThreadRingGaugeCalibration' WHERE FormName = 'ThreadRingGauge_cal'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'ThreadRingGaugeCalibration' WHERE form_name = 'ThreadRingGauge_cal'"
                    Dim updatedCount As Integer = 0
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        updatedCount = cmd.ExecuteNonQuery()
                    End Using

                    If updatedCount = 0 Then
                        Dim checkMappingSql As String = "SELECT COUNT(*) FROM calibrationmapping WHERE type_name = 'Thread Ring Gauge'"
                        Dim mappingCount As Integer = 0
                        Using cmd As New MySqlCommand(checkMappingSql, _con)
                            mappingCount = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using
                        If mappingCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Thread Ring Gauge', 'Gauge', 'ThreadRingGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        Else
                            Dim updateExistingSql As String = "UPDATE calibrationmapping SET form_name = 'ThreadRingGaugeCalibration' WHERE type_name = 'Thread Ring Gauge'"
                            Using cmd As New MySqlCommand(updateExistingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping ThreadRing): " & ex.Message)
                End Try

                ' Create thread_ring_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS thread_ring_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Obs_1 TEXT, " &
                        "Obs_2 TEXT, " &
                        "Obs_3 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating thread_ring_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for thread_ring_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM thread_ring_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO thread_ring_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Min_Limit, Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.000001', 'All Sizes', '0.000001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Thread Ring MASTER row: " & ex.Message)
                End Try

                ' Create height_master_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS height_master_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "MU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Flatness_UL DECIMAL(10,4), Flatness_LL DECIMAL(10,4), " &
                        "Parallelism_UL DECIMAL(10,4), Parallelism_LL DECIMAL(10,4), " &
                        "Go_Size VARCHAR(50), " &
                        "Go_Min_Limit DECIMAL(10,4), " &
                        "Go_Max_Limit DECIMAL(10,4), " &
                        "Go_Obs_1 DECIMAL(10,4), " &
                        "Go_Obs_2 DECIMAL(10,4), " &
                        "Go_Obs_3 DECIMAL(10,4), " &
                        "NoGo_Size VARCHAR(50), " &
                        "NoGo_Min_Limit DECIMAL(10,4), " &
                        "NoGo_Max_Limit DECIMAL(10,4), " &
                        "NoGo_Obs_1 DECIMAL(10,4), " &
                        "NoGo_Obs_2 DECIMAL(10,4), " &
                        "NoGo_Obs_3 DECIMAL(10,4), " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Judgement VARCHAR(50), " &
                        "Remark TEXT, " &
                        "Status VARCHAR(50) DEFAULT 'ACTIVE'" &
                        ")"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating height_master_calibration: " & ex.Message)
                End Try

                ' Migrate: Add MU column to height_master_calibration if it doesn't exist
                Try
                    Dim alterMuSql As String = "ALTER TABLE height_master_calibration ADD COLUMN IF NOT EXISTS MU VARCHAR(100) NULL AFTER TMU"
                    Using cmd As New MySqlCommand(alterMuSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error migrating MU column to height_master_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for height_master_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM height_master_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO height_master_calibration (RowType, LC, Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, Flatness_LL, Flatness_UL, Parallelism_LL, Parallelism_UL, Go_Min_Limit, Go_Max_Limit, NoGo_Min_Limit, NoGo_Max_Limit) VALUES " &
                            "('MASTER', '0.001', 18.00, 22.00, 40.00, 60.00, 0.0000, 1.0000, 0.0000, 1.0000, -5.0000, 5.0000, -5.0000, 5.0000), " &
                            "('MASTER', '0.01', 18.00, 22.00, 40.00, 60.00, 0.0000, 1.0000, 0.0000, 1.0000, -5.0000, 5.0000, -5.0000, 5.0000), " &
                            "('MASTER', '', 18.00, 22.00, 40.00, 60.00, 0.0000, 1.0000, 0.0000, 1.0000, -5.0000, 5.0000, -5.0000, 5.0000)"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding MASTER row: " & ex.Message)
                End Try

                ' Migrate legacy config and mapping names for Thread Plug Gauge
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET FormName = 'ThreadPlugGaugeCalibration' WHERE FormName = 'ThreadPlugGauge_cal'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'ThreadPlugGaugeCalibration' WHERE form_name = 'ThreadPlugGauge_cal'"
                    Dim updatedCount As Integer = 0
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        updatedCount = cmd.ExecuteNonQuery()
                    End Using

                    If updatedCount = 0 Then
                        Dim checkMappingSql As String = "SELECT COUNT(*) FROM calibrationmapping WHERE type_name = 'Thread Plug Gauge'"
                        Dim mappingCount As Integer = 0
                        Using cmd As New MySqlCommand(checkMappingSql, _con)
                            mappingCount = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using
                        If mappingCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Thread Plug Gauge', 'Gauge', 'ThreadPlugGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        Else
                            Dim updateExistingSql As String = "UPDATE calibrationmapping SET form_name = 'ThreadPlugGaugeCalibration' WHERE type_name = 'Thread Plug Gauge'"
                            Using cmd As New MySqlCommand(updateExistingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping ThreadPlug): " & ex.Message)
                End Try

                ' Create thread_plug_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS thread_plug_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Go_Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Go_Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "Go_Obs_1 TEXT, " &
                        "Go_Obs_2 TEXT, " &
                        "Go_Obs_3 TEXT, " &
                        "Go_Obs_4 TEXT, " &
                        "NoGo_Min_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "NoGo_Max_Limit DECIMAL(14,6) DEFAULT NULL, " &
                        "NoGo_Obs_1 TEXT, " &
                        "NoGo_Obs_2 TEXT, " &
                        "NoGo_Obs_3 TEXT, " &
                        "NoGo_Obs_4 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating thread_plug_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for thread_plug_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM thread_plug_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO thread_plug_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Go_Min_Limit, Go_Max_Limit, NoGo_Min_Limit, NoGo_Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.000001', 'All Sizes', '0.000001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL, NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Thread Plug MASTER row: " & ex.Message)
                End Try

                ' Migrate legacy config and mapping names for Taper Plug Gauge
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET FormName = 'TaperPlugGaugeCalibration' WHERE FormName = 'TaperPlugGauge_cal'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'TaperPlugGaugeCalibration' WHERE form_name = 'TaperPlugGauge_cal'"
                    Dim updatedCount As Integer = 0
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        updatedCount = cmd.ExecuteNonQuery()
                    End Using

                    If updatedCount = 0 Then
                        Dim checkMappingSql As String = "SELECT COUNT(*) FROM calibrationmapping WHERE type_name = 'Taper Plug Gauge'"
                        Dim mappingCount As Integer = 0
                        Using cmd As New MySqlCommand(checkMappingSql, _con)
                            mappingCount = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using
                        If mappingCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Taper Plug Gauge', 'Gauge', 'TaperPlugGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        Else
                            Dim updateExistingSql As String = "UPDATE calibrationmapping SET form_name = 'TaperPlugGaugeCalibration' WHERE type_name = 'Taper Plug Gauge'"
                            Using cmd As New MySqlCommand(updateExistingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping TaperPlug): " & ex.Message)
                End Try

                ' Create taper_plug_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS taper_plug_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Taper_Angle_Degree VARCHAR(100), " &
                        "Major_Dia_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Major_Dia_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Major_Dia_Obs_1 TEXT, " &
                        "Major_Dia_Obs_2 TEXT, " &
                        "Major_Dia_Obs_3 TEXT, " &
                        "Angle_Sec_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Angle_Sec_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Angle_Sec_Obs_1 TEXT, " &
                        "Angle_Sec_Obs_2 TEXT, " &
                        "Angle_Sec_Obs_3 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating taper_plug_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for taper_plug_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM taper_plug_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO taper_plug_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Major_Dia_Min_Limit, Major_Dia_Max_Limit, Angle_Sec_Min_Limit, Angle_Sec_Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.0001', 'All Sizes', '0.0001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL, NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Taper Plug MASTER row: " & ex.Message)
                End Try

                ' Migrate legacy config and mapping names for Taper Ring Gauge
                Try
                    Dim updateTemplateSql As String = "UPDATE calibration_template SET FormName = 'TaperRingGaugeCalibration' WHERE FormName = 'TaperRingGauge_cal'"
                    Using cmd As New MySqlCommand(updateTemplateSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim updateMappingSql As String = "UPDATE calibrationmapping SET form_name = 'TaperRingGaugeCalibration' WHERE form_name = 'TaperRingGauge_cal'"
                    Dim updatedCount As Integer = 0
                    Using cmd As New MySqlCommand(updateMappingSql, _con)
                        updatedCount = cmd.ExecuteNonQuery()
                    End Using

                    If updatedCount = 0 Then
                        Dim checkMappingSql As String = "SELECT COUNT(*) FROM calibrationmapping WHERE type_name = 'Taper Ring Gauge'"
                        Dim mappingCount As Integer = 0
                        Using cmd As New MySqlCommand(checkMappingSql, _con)
                            mappingCount = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using
                        If mappingCount = 0 Then
                            Dim insertMappingSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Taper Ring Gauge', 'Gauge', 'TaperRingGaugeCalibration')"
                            Using cmd As New MySqlCommand(insertMappingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        Else
                            Dim updateExistingSql As String = "UPDATE calibrationmapping SET form_name = 'TaperRingGaugeCalibration' WHERE type_name = 'Taper Ring Gauge'"
                            Using cmd As New MySqlCommand(updateExistingSql, _con)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping TaperRing): " & ex.Message)
                End Try

                ' Ensure 'Plain Taper Ring Gauge' is also mapped to TaperRingGaugeCalibration
                ' (Handles the case where gauge type was saved with "Plain" prefix vs without)
                Try
                    Dim checkPlainTaperMapping = ReadDatatable("SELECT 1 FROM calibrationmapping WHERE type_name = 'Plain Taper Ring Gauge' LIMIT 1")
                    If checkPlainTaperMapping.Rows.Count = 0 Then
                        Dim insertPlainTaperSql As String = "INSERT INTO calibrationmapping (type_name, category, form_name) VALUES ('Plain Taper Ring Gauge', 'Gauge', 'TaperRingGaugeCalibration')"
                        Using cmd As New MySqlCommand(insertPlainTaperSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Migration Error (calibrationmapping PlainTaperRing): " & ex.Message)
                End Try

                ' Create taper_ring_gauge_calibration table
                Try
                    Dim createTableSql As String = "CREATE TABLE IF NOT EXISTS taper_ring_gauge_calibration (" &
                        "id INT AUTO_INCREMENT PRIMARY KEY, " &
                        "RowType VARCHAR(50) DEFAULT 'RECORD', " &
                        "ControlNo VARCHAR(100), " &
                        "CycleName VARCHAR(255), " &
                        "Date DATE, " &
                        "Time TIME, " &
                        "Type VARCHAR(255), " &
                        "Size VARCHAR(100), " &
                        "LC VARCHAR(100), " &
                        "Color VARCHAR(100), " &
                        "Location VARCHAR(255), " &
                        "Temperature VARCHAR(50), " &
                        "Humidity VARCHAR(50), " &
                        "TMU VARCHAR(100), " &
                        "Env_Temp_LL DECIMAL(10,2), " &
                        "Env_Temp_UL DECIMAL(10,2), " &
                        "Env_Hum_LL DECIMAL(10,2), " &
                        "Env_Hum_UL DECIMAL(10,2), " &
                        "Parameter VARCHAR(255), " &
                        "Location_Obs VARCHAR(255), " &
                        "Taper_Angle_Degree VARCHAR(100), " &
                        "Major_Dia_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Major_Dia_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Major_Dia_Obs_1 TEXT, " &
                        "Major_Dia_Obs_2 TEXT, " &
                        "Major_Dia_Obs_3 TEXT, " &
                        "Angle_Sec_Min_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Angle_Sec_Max_Limit DECIMAL(10,4) DEFAULT NULL, " &
                        "Angle_Sec_Obs_1 TEXT, " &
                        "Angle_Sec_Obs_2 TEXT, " &
                        "Angle_Sec_Obs_3 TEXT, " &
                        "TimeIn VARCHAR(50), " &
                        "TimeOut VARCHAR(50), " &
                        "TotalTime VARCHAR(50), " &
                        "Status VARCHAR(50), " &
                        "Remark TEXT, " &
                        "uploaded_doc TEXT" &
                        ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;"
                    Using cmd As New MySqlCommand(createTableSql, _con)
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Error creating taper_ring_gauge_calibration: " & ex.Message)
                End Try

                ' Seed default MASTER row for taper_ring_gauge_calibration if none exists
                Try
                    Dim checkMasterSql As String = "SELECT COUNT(*) FROM taper_ring_gauge_calibration WHERE RowType = 'MASTER'"
                    Dim masterCount As Integer = 0
                    Using cmd As New MySqlCommand(checkMasterSql, _con)
                        masterCount = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                    If masterCount = 0 Then
                        Dim seedSql As String = "INSERT INTO taper_ring_gauge_calibration (" &
                            "RowType, ControlNo, Size, LC, " &
                            "Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, " &
                            "Major_Dia_Min_Limit, Major_Dia_Max_Limit, Angle_Sec_Min_Limit, Angle_Sec_Max_Limit" &
                            ") VALUES (" &
                            "'MASTER', 'LIMITS_0.0001', 'All Sizes', '0.0001', " &
                            "18.00, 22.00, 40.00, 60.00, " &
                            "NULL, NULL, NULL, NULL" &
                            ");"
                        Using cmd As New MySqlCommand(seedSql, _con)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error seeding Taper Ring MASTER row: " & ex.Message)
                End Try

                InitializeNGListTable()
            End If
        Catch ex As Exception
            Console.WriteLine("InitInterchangeTables Error: " & ex.Message)
        End Try
    End Sub

    Public Sub InitializeNGListTable()
        Try
            If MySQLDBConnect() = 1 Then
                Dim createSql As String = "CREATE TABLE IF NOT EXISTS ng_list (" &
                    "id INT AUTO_INCREMENT PRIMARY KEY, " &
                    "instrument_name VARCHAR(255), " &
                    "control_no VARCHAR(100), " &
                    "status VARCHAR(255), " &
                    "calibrated_date DATE, " &
                    "due_date DATE, " &
                    "calibration_status VARCHAR(255), " &
                    "reason TEXT, " &
                    "CycleName VARCHAR(255)" &
                    ")"
                Using cmd As New MySqlCommand(createSql, _con)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InitializeNGListTable Error: " & ex.Message)
        End Try
    End Sub

    Public Sub SyncDepartmentListToInterchange(cycleName As String)
        Try
            If MySQLDBConnect() = 1 Then
                ' 1. Sync from department_list to interchangeability
                Dim sql As String = "INSERT IGNORE INTO interchangeability (CycleName, ControlNo, Department, InstrumentName, SizeandRange, Color, Status, ActionDate, ActionTime, Remarks, ActionBy) " &
                                    "SELECT @cycleName, d.`Control No`, d.Department, d.InstrumentName, d.SizeandRange, d.Color, " &
                                    "IF(d.Status IS NOT NULL AND d.Status != '', d.Status, 'Pending'), CURDATE(), CURTIME(), IFNULL(d.Remarks, ''), 'System' " &
                                    "FROM department_list d " &
                                    "LEFT JOIN interchangeability i ON i.CycleName = @cycleName AND i.ControlNo = d.`Control No` " &
                                    "WHERE i.ControlNo IS NULL " &
                                    "AND d.`Control No` IS NOT NULL AND d.`Control No` != '' " &
                                    "GROUP BY d.`Control No`, d.Department, d.InstrumentName, d.SizeandRange, d.Color, d.Status, d.Remarks"
                Using cmd As New MySqlCommand(sql, _con)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.ExecuteNonQuery()
                End Using

                ' 2. Populate transaction tables for special imported statuses (to ensure they show up in Event Details)
                ' 2. Populate transaction tables for special imported statuses (to ensure they show up in Event Details)
                ' We wrap each in an independent Try-Catch so one failure doesn't block the others.

                ' --- WOP ---
                Try
                    Dim wopSql = "INSERT IGNORE INTO `wop` (ControlNo, WOPDate, Time, CycleName, Department, InstrumentName, Color, ReportedReason, ReportedBy) " &
                               "SELECT d.`Control No`, CURDATE(), CURTIME(), @cycleName, d.`Department`, d.`InstrumentName`, d.`Color`, IFNULL(d.`Remarks`, 'Imported'), 'System' " &
                               "FROM `department_list` d " &
                               "JOIN `interchangeability` i ON i.`ControlNo` = d.`Control No` AND i.`CycleName` = @cycleName " &
                               "LEFT JOIN `wop` w ON w.`ControlNo` = d.`Control No` AND w.`CycleName` = @cycleName " &
                               "WHERE TRIM(UPPER(d.`Status`)) = 'WOP' AND w.`ControlNo` IS NULL AND i.`ActionBy` = 'System'"
                    Using cmdWop As New MySqlCommand(wopSql, _con)
                        cmdWop.Parameters.AddWithValue("@cycleName", cycleName)
                        cmdWop.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Sync WOP Error: " & ex.Message)
                End Try

                ' --- Write off ---
                Try
                    Dim woSql = "INSERT IGNORE INTO `writeoff` (WriteOffDate, Time, CycleName, Type, InstrumentName, ControlNo, Section, Color, Reason, InstrumentStatus, RaisedBy) " &
                               "SELECT CURDATE(), CURTIME(), @cycleName, 'Import', d.`InstrumentName`, d.`Control No`, d.`Department`, d.`Color`, IFNULL(d.`Remarks`, 'Imported'), 'In-Active', 'System' " &
                               "FROM `department_list` d " &
                               "JOIN `interchangeability` i ON i.`ControlNo` = d.`Control No` AND i.`CycleName` = @cycleName " &
                               "LEFT JOIN `writeoff` w ON w.`ControlNo` = d.`Control No` AND (w.`CycleName` = @cycleName OR w.`CycleName` IS NULL OR w.`CycleName` = '') " &
                               "WHERE TRIM(UPPER(d.`Status`)) = 'WRITE OFF' AND w.`ControlNo` IS NULL AND i.`ActionBy` = 'System'"
                    Using cmdWo As New MySqlCommand(woSql, _con)
                        cmdWo.Parameters.AddWithValue("@cycleName", cycleName)
                        cmdWo.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Sync WriteOff Error: " & ex.Message)
                End Try

                ' --- Reintroduction ---
                Try
                    Dim reSql = "INSERT IGNORE INTO `reintroduction` (ReintroductionDate, Time, CycleName, Department, InstrumentName, ControlNo, Color, Reason, ReintroducedBy) " &
                               "SELECT CURDATE(), CURTIME(), @cycleName, d.`Department`, d.`InstrumentName`, d.`Control No`, d.`Color`, IFNULL(d.`Remarks`, 'Imported'), 'System' " &
                               "FROM `department_list` d " &
                               "JOIN `interchangeability` i ON i.`ControlNo` = d.`Control No` AND i.`CycleName` = @cycleName " &
                               "LEFT JOIN `reintroduction` r ON r.`ControlNo` = d.`Control No` AND r.`CycleName` = @cycleName " &
                               "WHERE TRIM(UPPER(d.`Status`)) = 'RE-INTRODUCTION' AND r.`ControlNo` IS NULL AND i.`ActionBy` = 'System'"
                    Using cmdRe As New MySqlCommand(reSql, _con)
                        cmdRe.Parameters.AddWithValue("@cycleName", cycleName)
                        cmdRe.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Sync Reintroduction Error: " & ex.Message)
                End Try

                ' --- Issue ---
                Try
                    ' Using LIKE for better resilience against subtle variations/hidden characters
                    Dim issueImportSql = "INSERT IGNORE INTO `issue` (ControlNo, IssueDate, Time, CycleName, Department, InstrumentName, Color, Remarks, IssuedBy) " &
                                       "SELECT d.`Control No`, CURDATE(), CURTIME(), @cycleName, d.`Department`, d.`InstrumentName`, d.`Color`, IFNULL(d.`Remarks`, 'Imported'), 'System' " &
                                       "FROM `department_list` d " &
                                       "JOIN `interchangeability` i ON i.`ControlNo` = d.`Control No` AND i.`CycleName` = @cycleName " &
                                       "LEFT JOIN `issue` i2 ON i2.`ControlNo` = d.`Control No` AND i2.`CycleName` = @cycleName " &
                                       "WHERE (TRIM(UPPER(d.`Status`)) LIKE 'ISSUE%') AND i2.`ControlNo` IS NULL AND i.`ActionBy` = 'System'"
                    Using cmdIssue As New MySqlCommand(issueImportSql, _con)
                        cmdIssue.Parameters.AddWithValue("@cycleName", cycleName)
                        cmdIssue.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Sync Issue Error: " & ex.Message)
                End Try

                ' --- Receive ---
                Try
                    ' Using LIKE for better resilience against subtle variations and common misspellings (RECIEVE)
                    ' Note: department_list doesn't have RFID_tag, so we omit it from selection
                    Dim receiveImportSql = "INSERT IGNORE INTO `receive` (ControlNo, ReceiveDate, Time, CycleName, Department, InstrumentName, Color, Remarks, ReceivedBy) " &
                                         "SELECT d.`Control No`, CURDATE(), CURTIME(), @cycleName, d.`Department`, d.`InstrumentName`, d.`Color`, IFNULL(d.`Remarks`, 'Imported'), 'System' " &
                                         "FROM `department_list` d " &
                                         "JOIN `interchangeability` i ON i.`ControlNo` = d.`Control No` AND i.`CycleName` = @cycleName " &
                                         "LEFT JOIN `receive` r ON r.`ControlNo` = d.`Control No` AND r.`CycleName` = @cycleName " &
                                         "WHERE (TRIM(UPPER(d.`Status`)) LIKE 'RECEIV%' OR TRIM(UPPER(d.`Status`)) LIKE 'RECIEV%') " &
                                         "AND r.`ControlNo` IS NULL AND i.`ActionBy` = 'System'"
                    Using cmdReceive As New MySqlCommand(receiveImportSql, _con)
                        cmdReceive.Parameters.AddWithValue("@cycleName", cycleName)
                        cmdReceive.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Sync Receive Error: " & ex.Message)
                End Try

                ' --- Regular Calibration Sync ---
                Try
                    ' 1. Add missing "Received" items to regular_calibration
                    Dim calibInsertSql = "INSERT INTO `regular_calibration` (instrument_name, control_no, status, is_calibrated, CycleName) " &
                                       "SELECT i.`InstrumentName`, i.`ControlNo`, i.`Status`, 'NO', @cycleName " &
                                       "FROM `interchangeability` i " &
                                       "LEFT JOIN `regular_calibration` rc ON rc.`control_no` = i.`ControlNo` AND rc.`CycleName` = @cycleName " &
                                       "WHERE (TRIM(UPPER(i.`Status`)) LIKE 'RECEIV%' OR TRIM(UPPER(i.`Status`)) LIKE 'RECIEV%') " &
                                       "AND rc.`control_no` IS NULL AND i.`CycleName` = @cycleName AND i.`ActionBy` = 'System'"
                    Using cmdInsert As New MySqlCommand(calibInsertSql, _con)
                        cmdInsert.Parameters.AddWithValue("@cycleName", cycleName)
                        cmdInsert.ExecuteNonQuery()
                    End Using

                    ' 2. Remove items from regular_calibration that are no longer "Received" in interchangeability for this cycle
                    ' Only handles items previously synced/added by System (import/sync)
                    Dim calibDeleteSql = "DELETE rc FROM `regular_calibration` rc " &
                                       "JOIN `interchangeability` i ON i.`ControlNo` = rc.`control_no` AND i.`CycleName` = rc.`CycleName` " &
                                       "WHERE rc.`CycleName` = @cycleName " &
                                       "AND NOT (TRIM(UPPER(i.`Status`)) LIKE 'RECEIV%' OR TRIM(UPPER(i.`Status`)) LIKE 'RECIEV%') " &
                                       "AND i.`ActionBy` = 'System'"
                    Using cmdDelete As New MySqlCommand(calibDeleteSql, _con)
                        cmdDelete.Parameters.AddWithValue("@cycleName", cycleName)
                        cmdDelete.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    Console.WriteLine("Sync Calibration Error: " & ex.Message)
                End Try

                ' 3. Capture a fresh snapshot in history
                LogCycleHistoryBulk(cycleName)
            End If
        Catch ex As Exception
            Console.WriteLine("SyncDepartmentListToInterchange Error: " & ex.Message)
        End Try
    End Sub


    Public Function InsertInterchangeRecord(cycleName As String, controlNo As String, dept As String, instName As String, sizeRange As String, color As String, status As String, remark As String, Optional rfidTag As String = "") As Boolean
        If String.IsNullOrEmpty(controlNo) Then Return False
        Dim currentUser = Application.Current.Properties("Username")?.ToString()
        Try
            If MySQLDBConnect() = 1 Then
                ' 1. UPSERT into interchangeability
                Dim query As String = "INSERT INTO interchangeability (CycleName, ControlNo, Department, InstrumentName, SizeandRange, Color, Status, ActionDate, ActionTime, Remarks, ActionBy, RFID_tag) " &
                                     "VALUES (@cycle, @ctrl, @dept, @inst, @size, @color, @status, @date, @time, @rem, @user, @rfid) " &
                                     "ON DUPLICATE KEY UPDATE Status=@status, ActionDate=@date, ActionTime=@time, Remarks=@rem, ActionBy=@user, RFID_tag=@rfid"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@rem", remark)
                    cmd.Parameters.AddWithValue("@user", currentUser)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@dept", dept)
                    cmd.Parameters.AddWithValue("@inst", instName)
                    cmd.Parameters.AddWithValue("@size", sizeRange)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@rfid", rfidTag)

                    If cmd.ExecuteNonQuery() > 0 Then
                        ' 2. Log in Cycle History
                        LogInterchangeHistory(cycleName, controlNo, dept, instName, color, status, remark)

                        ' 3. Automatically populate transaction tables for common statuses to avoid missing markers
                        ' Standardize checks to handle case variations
                        Dim checkStatus = status.Trim().ToUpper()
                        If checkStatus.StartsWith("ISSUE") Then
                            InsertIssueRecord(controlNo, cycleName, dept, instName, color, remark, rfidTag)
                            ' If we are marking as Issued without providing a NEW RFID, clear the old one
                            If String.IsNullOrEmpty(rfidTag) Then ClearRFIDTag(controlNo)
                        ElseIf checkStatus.StartsWith("RECEIV") OrElse checkStatus.StartsWith("RECIEV") Then
                            InsertReceiveRecord(controlNo, cycleName, dept, instName, color, remark, rfidTag)
                        ElseIf checkStatus = "WOP" OrElse checkStatus = "WRITE OFF" OrElse checkStatus = "TEMP ISSUANCE" Then
                            ' Immediately clear RFID for terminal or long-term missing statuses
                            ClearRFIDTag(controlNo)
                        End If

                        ' 4. SYNC TO CALIBRATION TABLE
                        ' Route to reintroduction_calibration if item was reintroduced this cycle,
                        ' otherwise route to regular_calibration.
                        If checkStatus.StartsWith("RECEIV") OrElse checkStatus.StartsWith("RECIEV") Then
                            If IsReintroductionPending(controlNo, cycleName) Then
                                SyncToReintroductionCalibration(controlNo, instName, status, cycleName)
                            Else
                                SyncToRegularCalibration(controlNo, instName, status, cycleName)
                            End If
                        ElseIf checkStatus <> "WOP" Then
                            RemoveFromRegularCalibration(controlNo, cycleName)
                        End If

                        Return True
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertInterchangeRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Sub LogCycleHistoryBulk(cycleName As String)
        If String.IsNullOrEmpty(cycleName) Then Return
        Try
            If MySQLDBConnect() = 1 Then
                Dim tableName = GetHistoryTableName(DateTime.Now.Year)
                EnsureHistoryTableExists(tableName)

                ' 1. Delete existing records for this cycle in the history shard to ensure a fresh snapshot
                Dim deleteSql As String = $"DELETE FROM `{tableName}` WHERE CycleName = @cycleName"
                Using delCmd As New MySqlCommand(deleteSql, _con)
                    delCmd.Parameters.AddWithValue("@cycleName", cycleName)
                    delCmd.ExecuteNonQuery()
                End Using

                ' 2. Insert all current states from interchangeability for this cycle
                Dim bulkSql As String = $"INSERT INTO `{tableName}` (LogDate, CycleName, ControlNo, Department, InstrumentName, Color, Status, ActionBy, Remarks) " &
                                      "SELECT NOW(), CycleName, ControlNo, Department, InstrumentName, Color, Status, ActionBy, Remarks " &
                                      "FROM interchangeability " &
                                      "WHERE CycleName = @cycleName"
                Using insCmd As New MySqlCommand(bulkSql, _con)
                    insCmd.Parameters.AddWithValue("@cycleName", cycleName)
                    insCmd.ExecuteNonQuery()
                End Using

                Console.WriteLine($"History snapshot captured for cycle: {cycleName}")
            End If
        Catch ex As Exception
            Console.WriteLine("LogCycleHistoryBulk Error: " & ex.Message)
        End Try
    End Sub

    Public Function GetHistoryTableName(ByVal year As Integer) As String
        Return "interchange_history"
    End Function

    Public Sub EnsureHistoryTableExists(ByVal tableName As String)
        Try
            If MySQLDBConnect() = 1 Then
                Dim createSql As String = $"CREATE TABLE IF NOT EXISTS `{tableName}` (" &
                    "id INT AUTO_INCREMENT PRIMARY KEY, " &
                    "LogDate DATETIME, " &
                    "CycleName VARCHAR(255), " &
                    "ControlNo VARCHAR(255), " &
                    "Department VARCHAR(255), " &
                    "InstrumentName VARCHAR(255), " &
                    "Color VARCHAR(100), " &
                    "Status VARCHAR(255), " &
                    "ActionBy VARCHAR(255), " &
                    "Remarks TEXT, " &
                    "INDEX idx_ctrl_cycle (ControlNo, CycleName), " &
                    "INDEX idx_logdate (LogDate), " &
                    "INDEX idx_dept (Department), " &
                    "INDEX idx_status (Status)" &
                    ")"
                Using cmd As New MySqlCommand(createSql, _con)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("EnsureHistoryTableExists Error: " & ex.Message)
        End Try
    End Sub

    Public Sub LogInterchangeHistory(ByVal cycleName As String, ByVal controlNo As String, ByVal dept As String, ByVal instName As String, ByVal color As String, ByVal status As String, ByVal remark As String)
        Try
            Dim tableName = GetHistoryTableName(DateTime.Now.Year)
            EnsureHistoryTableExists(tableName)

            If MySQLDBConnect() = 1 Then
                Dim currentUser = Application.Current.Properties("Username")?.ToString()
                Dim query As String = $"INSERT INTO `{tableName}` (LogDate, CycleName, ControlNo, Department, InstrumentName, Color, Status, ActionBy, Remarks) " &
                                    "VALUES (NOW(), @cycle, @ctrl, @dept, @inst, @color, @status, @user, @rem)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@dept", dept)
                    cmd.Parameters.AddWithValue("@inst", instName)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@user", If(currentUser, "System"))
                    cmd.Parameters.AddWithValue("@rem", remark)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("LogInterchangeHistory Error: " & ex.Message)
        End Try
    End Sub

    Public Function GetItemCategory(controlNo As String) As String
        Try
            Dim dtSettings = ReadDatatable("SELECT TypeName, Category FROM type_details")
            For Each r As DataRow In dtSettings.Rows
                Dim testTbl = TypeNameToTableName(r("TypeName").ToString(), forInventory:=True)
                Dim chk = ReadDatatable($"SELECT ControlNo FROM `{testTbl}` WHERE ControlNo='{controlNo.Replace("'", "''")}'")
                If chk.Rows.Count > 0 Then
                    Return r("Category").ToString()
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("GetItemCategory Error: " & ex.Message)
        End Try
        Return "Unknown"
    End Function

    Public Sub EnsureAdminPassword()
        Try
            If MySQLDBConnect() = 1 Then
                Dim sql = "INSERT IGNORE INTO settings (property, val) VALUES ('AppPass', 'RDL123')"
                Using cmd As New MySqlCommand(sql, _con)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Function UpdateInterchangeStatus(controlNo As String, cycleName As String, newStatus As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim currentUser = Application.Current.Properties("Username")?.ToString()

                ' Get record details for logging before update (or just use parameters passed)
                Dim dt = ReadDatatable($"SELECT Department, InstrumentName, Color FROM interchangeability WHERE ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}'")
                Dim dept = "", inst = "", color = ""
                If dt.Rows.Count > 0 Then
                    dept = dt.Rows(0)("Department").ToString()
                    inst = dt.Rows(0)("InstrumentName").ToString()
                    color = dt.Rows(0)("Color").ToString()
                End If

                Dim query As String = "UPDATE interchangeability SET Status=@status, ActionBy=@user, ActionDate=@date, ActionTime=@time, Remarks=@rem WHERE ControlNo=@ctrl AND CycleName=@cycle"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@status", newStatus)
                    cmd.Parameters.AddWithValue("@user", currentUser)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@rem", remark)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    If cmd.ExecuteNonQuery() > 0 Then
                        LogInterchangeHistory(cycleName, controlNo, dept, inst, color, newStatus, remark)
                        Return True
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateInterchangeStatus Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertWOPRecord(controlNo As String, cycleName As String, dept As String, instName As String, color As String, reason As String) As Boolean
        Dim currentUser = Application.Current.Properties("Username")?.ToString()
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO wop (ControlNo, WOPDate, Time, CycleName, Department, InstrumentName, Color, ReportedReason, ReportedBy) " &
                                    "VALUES (@ctrl, @date, @time, @cycle, @dept, @name, @color, @rem, @user)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@dept", dept)
                    cmd.Parameters.AddWithValue("@name", instName)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@rem", reason)
                    cmd.Parameters.AddWithValue("@user", currentUser)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertWOPRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertTempIssuance(controlNo As String, cycleName As String, dept As String, instName As String, color As String, reason As String, docPath As String) As Boolean
        Dim currentUser = Application.Current.Properties("Username")?.ToString()
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO temporary_issuance (ControlNo, IssueDate, Time, CycleName, Department, InstrumentName, Color, Reason, DocumentPath, IssuedBy) " &
                                    "VALUES (@ctrl, @date, @time, @cycle, @dept, @name, @color, @rem, @doc, @user)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@dept", dept)
                    cmd.Parameters.AddWithValue("@name", instName)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@rem", reason)
                    cmd.Parameters.AddWithValue("@doc", docPath)
                    cmd.Parameters.AddWithValue("@user", currentUser)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertTempIssuance Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertIssueRecord(controlNo As String, cycleName As String, dept As String, instName As String, color As String, remarks As String, rfidTag As String) As Boolean
        Dim currentUser = Application.Current.Properties("Username")?.ToString()
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO issue (ControlNo, IssueDate, Time, CycleName, Department, InstrumentName, Color, Remarks, IssuedBy, RFID_tag) " &
                                    "VALUES (@ctrl, @date, @time, @cycle, @dept, @name, @color, @rem, @user, @rfid)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@dept", dept)
                    cmd.Parameters.AddWithValue("@name", instName)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@rem", remarks)
                    cmd.Parameters.AddWithValue("@user", currentUser)
                    cmd.Parameters.AddWithValue("@rfid", rfidTag)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertIssueRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertReceiveRecord(controlNo As String, cycleName As String, dept As String, instName As String, color As String, remarks As String, rfidTag As String) As Boolean
        Dim currentUser = Application.Current.Properties("Username")?.ToString()
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO receive (ControlNo, ReceiveDate, Time, CycleName, Department, InstrumentName, Color, Remarks, ReceivedBy, RFID_tag) " &
                                    "VALUES (@ctrl, @date, @time, @cycle, @dept, @name, @color, @rem, @user, @rfid)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@dept", dept)
                    cmd.Parameters.AddWithValue("@name", instName)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@rem", remarks)
                    cmd.Parameters.AddWithValue("@user", currentUser)
                    cmd.Parameters.AddWithValue("@rfid", rfidTag)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertReceiveRecord Error: " & ex.Message)
        End Try
        Return False
    End Function
    Public Function GetActiveCycleName() As String
        ' 1. Check for a manual override (set by Reset Current Cycle)
        Dim overrideCycle = GetConfigValue("ActiveCycle")
        If Not String.IsNullOrEmpty(overrideCycle) Then
            Return overrideCycle
        End If

        ' 2. Check user-configured date ranges from Admin settings
        Dim today = DateTime.Today
        Try
            Dim gsVal = GetConfigValue("GreenCycleStart")
            Dim geVal = GetConfigValue("GreenCycleEnd")
            Dim ysVal = GetConfigValue("YellowCycleStart")
            Dim yeVal = GetConfigValue("YellowCycleEnd")

            If Not String.IsNullOrEmpty(gsVal) AndAlso Not String.IsNullOrEmpty(geVal) AndAlso
               Not String.IsNullOrEmpty(ysVal) AndAlso Not String.IsNullOrEmpty(yeVal) Then

                Dim greenStart = DateTime.Parse(gsVal)
                Dim greenEnd = DateTime.Parse(geVal)
                Dim yellowStart = DateTime.Parse(ysVal)
                Dim yellowEnd = DateTime.Parse(yeVal)

                If today >= greenStart AndAlso today <= greenEnd Then
                    ' Use actual month abbreviation from the configured start date
                    Return $"{greenStart.ToString("MMM")}'{greenStart.ToString("yy")} Green Cycle"
                End If
                If today >= yellowStart AndAlso today <= yellowEnd Then
                    ' Use actual month abbreviation from the configured start date
                    Return $"{yellowStart.ToString("MMM")}'{yellowStart.ToString("yy")} Yellow Cycle"
                End If

                ' Today is outside both ranges — use configured start colour as fallback
                Dim startColor = GetConfigValue("CycleStartColor")
                If startColor = "Yellow" Then
                    Return $"{yellowStart.ToString("MMM")}'{yellowStart.ToString("yy")} Yellow Cycle"
                Else
                    Return $"{greenStart.ToString("MMM")}'{greenStart.ToString("yy")} Green Cycle"
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("GetActiveCycleName config read error: " & ex.Message)
        End Try

        ' 3. Last-resort fallback: old hardcoded date-based calculation
        Dim dt = DateTime.Today
        If dt.Month >= 1 AndAlso dt.Month <= 6 Then
            Return $"Jan'{dt.ToString("yy")} Yellow Cycle"
        Else
            Return $"Jul'{dt.ToString("yy")} Green Cycle"
        End If
    End Function



    Public Function GetCycleNames() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("cycle_name", GetType(String))
        Dim activeCycle = GetActiveCycleName()
        dt.Rows.Add(activeCycle)
        Try
            If MySQLDBConnect() = 1 Then
                Dim sql = "SELECT DISTINCT CycleName AS cycle_name FROM result_list WHERE CycleName IS NOT NULL AND CycleName != ''"
                Using cmd As New MySqlCommand(sql, _con)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim cycle = reader("cycle_name").ToString()
                            If cycle <> activeCycle Then
                                dt.Rows.Add(cycle)
                            End If
                        End While
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetCycleNames Error: " & ex.Message)
        End Try
        Return dt
    End Function

    Public Function GetPreviousCycleName(currentCycle As String) As String
        Try
            Dim isYellow = currentCycle.Contains("Yellow Cycle")
            Dim parts = currentCycle.Split(" "c)
            Dim yearPart = Integer.Parse(parts(0).Split("'"c)(1))

            Dim gsVal = GetConfigValue("GreenCycleStart")
            Dim ysVal = GetConfigValue("YellowCycleStart")
            Dim greenStart As DateTime
            Dim yellowStart As DateTime
            If Not DateTime.TryParse(gsVal, greenStart) Then greenStart = New DateTime(2000, 7, 1)
            If Not DateTime.TryParse(ysVal, yellowStart) Then yellowStart = New DateTime(2000, 1, 1)

            ' Yellow -> prev is Green (year - 1); Green -> prev is Yellow (same year)
            If isYellow Then
                Return $"{greenStart.ToString("MMM")}'{(yearPart - 1).ToString("00")} Green Cycle"
            Else
                Return $"{yellowStart.ToString("MMM")}'{yearPart.ToString("00")} Yellow Cycle"
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Sub CleanupStaleRFIDs(currentCycle As String)
        Try
            If MySQLDBConnect() = 1 Then
                Dim prevCycle = GetPreviousCycleName(currentCycle)
                If String.IsNullOrEmpty(prevCycle) Then Return

                ' Identify tools that were NOT "Received" in the previous cycle but have an RFID tag recorded ANYWHERE
                ' Rule: RFID is valid for Received Cycle + 1. 
                ' We use JOINs with both interchangeability and receive tables to find if the item HAD an RFID 
                ' anywhere, even if the immediate previous cycle record has it as NULL.
                ' FIX: Also exclude items that are ALREADY "Received" in the CURRENT cycle to prevent clearing tags after a tab switch.
                Dim sql = "SELECT DISTINCT i.ControlNo FROM interchangeability i " &
                          "LEFT JOIN interchangeability ih ON i.ControlNo = ih.ControlNo AND ih.RFID_tag IS NOT NULL AND ih.RFID_tag != '' " &
                          "LEFT JOIN receive r ON i.ControlNo = r.ControlNo AND r.RFID_tag IS NOT NULL AND r.RFID_tag != '' " &
                          "LEFT JOIN interchangeability ic ON i.ControlNo = ic.ControlNo AND ic.CycleName = @current AND (TRIM(UPPER(ic.Status)) LIKE 'RECEIV%' OR TRIM(UPPER(ic.Status)) LIKE 'RECIEV%') " &
                          "WHERE i.CycleName = @prev " &
                          "AND (TRIM(UPPER(i.Status)) NOT LIKE 'RECEIV%' AND TRIM(UPPER(i.Status)) NOT LIKE 'RECIEV%') " &
                          "AND (ih.ControlNo IS NOT NULL OR r.ControlNo IS NOT NULL) " &
                          "AND ic.ControlNo IS NULL"

                Using cmd As New MySqlCommand(sql, _con)
                    cmd.Parameters.AddWithValue("@prev", prevCycle)
                    cmd.Parameters.AddWithValue("@current", currentCycle)

                    Dim dt As New DataTable()
                    Dim adapter As New MySqlDataAdapter(cmd)
                    adapter.Fill(dt)

                    If dt.Rows.Count > 0 Then
                        Console.WriteLine($"Found {dt.Rows.Count} stale RFIDs from cycle {prevCycle}. Clearing now...")
                        For Each row As DataRow In dt.Rows
                            Dim ctrl = row("ControlNo").ToString()
                            ClearRFIDTag(ctrl)
                        Next
                    End If
                End Using

                ' Part 2: Identify tools already marked terminal (WOP, Write-off, Temp issuance) in the CURRENT cycle
                Dim sqlCurrent = "SELECT ControlNo FROM interchangeability " &
                                 "WHERE CycleName = @current " &
                                 "AND (TRIM(UPPER(Status)) IN ('WOP', 'WRITE OFF', 'TEMP ISSUANCE')) " &
                                 "AND RFID_tag IS NOT NULL AND RFID_tag != ''"

                Using cmd2 As New MySqlCommand(sqlCurrent, _con)
                    cmd2.Parameters.AddWithValue("@current", currentCycle)

                    Dim dt2 As New DataTable()
                    Dim adapter2 As New MySqlDataAdapter(cmd2)
                    adapter2.Fill(dt2)

                    If dt2.Rows.Count > 0 Then
                        Console.WriteLine($"Found {dt2.Rows.Count} items already terminal in current cycle {currentCycle}. Clearing RFIDs now...")
                        For Each row As DataRow In dt2.Rows
                            ClearRFIDTag(row("ControlNo").ToString())
                        Next
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("CleanupStaleRFIDs Error: " & ex.Message)
        End Try
    End Sub

    Public Function GetConfigValue(key As String) As String
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT ConfigValue FROM system_config WHERE ConfigKey = @key"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@key", key)
                    Dim val = cmd.ExecuteScalar()
                    Return If(val IsNot Nothing, val.ToString(), "")
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetConfigValue Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function SetConfigValue(key As String, value As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO system_config (ConfigKey, ConfigValue) VALUES (@key, @val) " &
                                    "ON DUPLICATE KEY UPDATE ConfigValue = VALUES(ConfigValue)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@key", key)
                    cmd.Parameters.AddWithValue("@val", value)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("SetConfigValue Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function ClearRegularCalibrationTable() As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sql = "TRUNCATE TABLE regular_calibration"
                Using cmd As New MySqlCommand(sql, _con)
                    Return cmd.ExecuteNonQuery() >= 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("ClearRegularCalibrationTable Error: " & ex.Message)
        End Try
        Return False
    End Function

    Private Sub SyncToRegularCalibration(controlNo As String, instName As String, status As String, cycleName As String)
        Try
            If MySQLDBConnect() = 1 Then
                ' Check if record exists for this specific instrument and cycle
                Dim checkSql = "SELECT id FROM regular_calibration WHERE control_no = @ctrl AND CycleName = @cycle"
                Dim existingId As Object = Nothing
                Using cmdCheck As New MySqlCommand(checkSql, _con)
                    cmdCheck.Parameters.AddWithValue("@ctrl", controlNo)
                    cmdCheck.Parameters.AddWithValue("@cycle", cycleName)
                    existingId = cmdCheck.ExecuteScalar()
                End Using

                If existingId IsNot Nothing Then
                    ' Update existing record for this cycle
                    Dim updateSql = "UPDATE regular_calibration SET instrument_name=@name, status=@status, is_calibrated='NO' WHERE control_no=@ctrl AND CycleName=@cycle"
                    Using cmdUpdate As New MySqlCommand(updateSql, _con)
                        cmdUpdate.Parameters.AddWithValue("@name", instName)
                        cmdUpdate.Parameters.AddWithValue("@status", status)
                        cmdUpdate.Parameters.AddWithValue("@ctrl", controlNo)
                        cmdUpdate.Parameters.AddWithValue("@cycle", cycleName)
                        cmdUpdate.ExecuteNonQuery()
                    End Using
                Else
                    ' Insert new record for this cycle
                    Dim insertSql = "INSERT INTO regular_calibration (instrument_name, control_no, status, is_calibrated, CycleName) " &
                                   "VALUES (@name, @ctrl, @status, 'NO', @cycle)"
                    Using cmdInsert As New MySqlCommand(insertSql, _con)
                        cmdInsert.Parameters.AddWithValue("@name", instName)
                        cmdInsert.Parameters.AddWithValue("@ctrl", controlNo)
                        cmdInsert.Parameters.AddWithValue("@status", status)
                        cmdInsert.Parameters.AddWithValue("@cycle", cycleName)
                        cmdInsert.ExecuteNonQuery()
                    End Using
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("SyncToRegularCalibration Error: " & ex.Message)
        End Try
    End Sub

    Private Sub RemoveFromRegularCalibration(controlNo As String, cycleName As String)
        Try
            If MySQLDBConnect() = 1 Then
                Dim sql = "DELETE FROM regular_calibration WHERE control_no = @ctrl AND CycleName = @cycle"
                Using cmd As New MySqlCommand(sql, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("RemoveFromRegularCalibration Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Returns True if the given controlNo has a reintroduction record for the given cycle.
    ''' Used to decide whether a Received item should be tracked under reintroduction_calibration.
    ''' </summary>
    Private Function IsReintroductionPending(controlNo As String, cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sql = "SELECT COUNT(*) FROM reintroduction WHERE ControlNo = @ctrl AND CycleName = @cycle"
                Using cmd As New MySqlCommand(sql, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    Dim cnt = Convert.ToInt32(cmd.ExecuteScalar())
                    Return cnt > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("IsReintroductionPending Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Adds or updates an entry in reintroduction_calibration for the given control number and cycle.
    ''' Called when a reintroduced item is received, so calibration is tracked in the Reintro tab.
    ''' </summary>
    Private Sub SyncToReintroductionCalibration(controlNo As String, instName As String, status As String, cycleName As String)
        Try
            If MySQLDBConnect() = 1 Then
                Dim checkSql = "SELECT id FROM reintroduction_calibration WHERE control_no = @ctrl AND CycleName = @cycle"
                Dim existingId As Object = Nothing
                Using cmdCheck As New MySqlCommand(checkSql, _con)
                    cmdCheck.Parameters.AddWithValue("@ctrl", controlNo)
                    cmdCheck.Parameters.AddWithValue("@cycle", cycleName)
                    existingId = cmdCheck.ExecuteScalar()
                End Using

                If existingId IsNot Nothing Then
                    Dim updateSql = "UPDATE reintroduction_calibration SET instrument_name=@name, status=@status, is_calibrated='NO' WHERE control_no=@ctrl AND CycleName=@cycle"
                    Using cmdUpdate As New MySqlCommand(updateSql, _con)
                        cmdUpdate.Parameters.AddWithValue("@name", instName)
                        cmdUpdate.Parameters.AddWithValue("@status", status)
                        cmdUpdate.Parameters.AddWithValue("@ctrl", controlNo)
                        cmdUpdate.Parameters.AddWithValue("@cycle", cycleName)
                        cmdUpdate.ExecuteNonQuery()
                    End Using
                Else
                    Dim insertSql = "INSERT INTO reintroduction_calibration (instrument_name, control_no, status, is_calibrated, CycleName) " &
                                   "VALUES (@name, @ctrl, @status, 'NO', @cycle)"
                    Using cmdInsert As New MySqlCommand(insertSql, _con)
                        cmdInsert.Parameters.AddWithValue("@name", instName)
                        cmdInsert.Parameters.AddWithValue("@ctrl", controlNo)
                        cmdInsert.Parameters.AddWithValue("@status", status)
                        cmdInsert.Parameters.AddWithValue("@cycle", cycleName)
                        cmdInsert.ExecuteNonQuery()
                    End Using
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("SyncToReintroductionCalibration Error: " & ex.Message)
        End Try
    End Sub

    Public Function GetEventDetails(controlNo As String, cycleName As String, eventType As String) As Dictionary(Of String, String)
        Dim results As New Dictionary(Of String, String)()
        results("ActionRemarks") = ""
        results("ActionDoc") = ""
        results("ReintroRemarks") = ""
        results("ReintroDoc") = ""

        Try
            If MySQLDBConnect() = 1 Then
                ' 1. Fetch Reintroduction if needed
                If eventType.Contains("(Reintroduced)") OrElse eventType.ToLower().Contains("reintroduction") Then
                    Dim dtRe = ReadDatatable($"SELECT Reason, DocumentPath AS DocPath FROM reintroduction WHERE ControlNo='{controlNo.Replace("'", "''")}' AND (CycleName='{cycleName.Replace("'", "''")}' OR CycleName IS NULL OR CycleName='') ORDER BY ID DESC LIMIT 1")
                    If dtRe IsNot Nothing AndAlso dtRe.Rows.Count > 0 Then
                        results("ReintroRemarks") = dtRe.Rows(0)("Reason").ToString()
                        results("ReintroDoc") = dtRe.Rows(0)("DocPath").ToString()
                    End If
                End If

                ' 2. Fetch specific action details
                Dim baseAction = eventType.Replace("(Reintroduced)", "").Trim()

                ' 1. Fetch from specific transaction tables first (Source of Truth for Documents)
                If Not baseAction.ToLower().Contains("reintroduction") Then
                    Dim dt As DataTable = Nothing
                    Select Case True
                        Case baseAction.Equals("WOP", StringComparison.OrdinalIgnoreCase) OrElse baseAction.Equals("WOP Request", StringComparison.OrdinalIgnoreCase)
                            dt = ReadDatatable($"SELECT ReportedReason AS Reason, '' AS DocPath FROM wop WHERE ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY ID DESC LIMIT 1")
                        Case baseAction.Equals("Write off", StringComparison.OrdinalIgnoreCase) OrElse baseAction.Equals("Write Off", StringComparison.OrdinalIgnoreCase)
                            dt = ReadDatatable($"SELECT Reason, DocumentPath AS DocPath FROM writeoff WHERE ControlNo='{controlNo.Replace("'", "''")}' AND (CycleName='{cycleName.Replace("'", "''")}' OR CycleName IS NULL OR CycleName='') ORDER BY ID DESC LIMIT 1")
                        Case baseAction.Equals("Temp issuance", StringComparison.OrdinalIgnoreCase) OrElse baseAction.Equals("Temp Issuance", StringComparison.OrdinalIgnoreCase)
                            dt = ReadDatatable($"SELECT Reason, DocumentPath AS DocPath FROM temporary_issuance WHERE ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY ID DESC LIMIT 1")
                        Case baseAction.StartsWith("Issue", StringComparison.OrdinalIgnoreCase)
                            dt = ReadDatatable($"SELECT Remarks AS Reason, '' AS DocPath FROM issue WHERE ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY ID DESC LIMIT 1")
                        Case baseAction.StartsWith("Receiv", StringComparison.OrdinalIgnoreCase) OrElse baseAction.StartsWith("Reciev", StringComparison.OrdinalIgnoreCase)
                            dt = ReadDatatable($"SELECT Remarks AS Reason, '' AS DocPath FROM receive WHERE ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY ID DESC LIMIT 1")
                    End Select

                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        results("ActionRemarks") = dt.Rows(0)("Reason").ToString()
                        results("ActionDoc") = dt.Rows(0)("DocPath").ToString()
                    End If
                End If

                ' 2. Fallback to interchangeability or interchange_history if ActionRemarks is still empty
                If String.IsNullOrEmpty(results("ActionRemarks")) Then
                    ' Try interchangeability (current state)
                    Dim dtInter = ReadDatatable($"SELECT Remarks FROM interchangeability WHERE ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}'")
                    If dtInter IsNot Nothing AndAlso dtInter.Rows.Count > 0 Then
                        results("ActionRemarks") = dtInter.Rows(0)("Remarks").ToString()
                    End If

                    If String.IsNullOrEmpty(results("ActionRemarks")) Then
                        ' Try fallback to interchange_history
                        Dim dtHist = ReadDatatable($"SELECT Remarks FROM interchange_history WHERE ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' AND (Status LIKE '%{baseAction}%' OR Status LIKE '%Import%') ORDER BY LogDate DESC LIMIT 1")
                        If dtHist IsNot Nothing AndAlso dtHist.Rows.Count > 0 Then
                            results("ActionRemarks") = dtHist.Rows(0)("Remarks").ToString()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("GetEventDetails Error: " & ex.Message)
        End Try
        Return results
    End Function

    Public Function GetRFIDTag(controlNo As String) As String
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT RFID_tag FROM interchangeability WHERE ControlNo = @ctrl AND RFID_tag IS NOT NULL AND RFID_tag <> '' ORDER BY ActionDate DESC, ActionTime DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetRFIDTag Error: " & ex.Message)
        End Try
        Return ""
    End Function
    Public Function ClearRFIDTag(controlNo As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                ' 1. Clear from interchangeability (All Cycles)
                Dim query As String = "UPDATE interchangeability SET RFID_tag = NULL WHERE ControlNo = @ctrl"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.ExecuteNonQuery()
                End Using

                ' 2. Clear from receive table logs
                Try
                    Dim clearRecSql = "UPDATE `receive` SET RFID_tag = NULL WHERE ControlNo = @ctrl"
                    Using cmdRec As New MySqlCommand(clearRecSql, _con)
                        cmdRec.Parameters.AddWithValue("@ctrl", controlNo)
                        cmdRec.ExecuteNonQuery()
                    End Using
                Catch exRec As Exception
                    Console.WriteLine("ClearRFIDTag Receive Error: " & exRec.Message)
                End Try

                ' 3. Clear from all type-specific tables
                Dim typeTables = GetAllTypeTables()
                For Each tbl In typeTables
                    Try
                        Dim clearSql = $"UPDATE `{tbl}` SET RFID_tag = NULL WHERE ControlNo = @ctrl"
                        Using cmd2 As New MySqlCommand(clearSql, _con)
                            cmd2.Parameters.AddWithValue("@ctrl", controlNo)
                            cmd2.ExecuteNonQuery()
                        End Using
                    Catch exTbl As Exception
                        ' Individual table fail shouldn't block others
                        Console.WriteLine($"ClearRFIDTag Table Error ({tbl}): " & exTbl.Message)
                    End Try
                Next
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine("ClearRFIDTag Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Copies a source file to a centralized Documents subfolder and returns the relative path.
    ''' </summary>
    ''' <param name="sourcePath">The absolute path of the selected file.</param>
    ''' <param name="subFolder">The subfolder name (e.g., "Instrument", "WriteOff").</param>
    ''' <param name="controlNo">The Control Number to include in the filename.</param>
    ''' <returns>The relative path (e.g., "Documents\WriteOff\filename.pdf") or empty string if failed.</returns>
    ''' <summary>
    ''' Maps a subfolder name to its section folder inside the storage root.
    ''' E.g. "Instrument" -> "Database\Instrument", "WriteOff" -> "Interchangeability\WriteOff"
    ''' </summary>
    Private Shared Function GetSectionedSubFolder(subFolder As String) As String
        Select Case subFolder.ToLower()
            Case "instrument", "gauge"
                Return Path.Combine("Database", subFolder)
            Case "reintroduction", "writeoff", "tempissuance"
                Return Path.Combine("Interchangeability", subFolder)
            Case "reports", "masters"
                Return Path.Combine("Calibration", subFolder)
            Case "documents"
                Return Path.Combine("Inventory", subFolder)
            Case "exports"
                Return Path.Combine("Records", subFolder)
            Case Else
                Return subFolder
        End Select
    End Function

    Public Shared Function CopyFileToDocuments(sourcePath As String, subFolder As String, controlNo As String) As String
        Try
            If String.IsNullOrEmpty(sourcePath) OrElse Not File.Exists(sourcePath) Then Return ""

            ' ── Block upload if no custom storage directory has been configured ──
            If String.IsNullOrWhiteSpace(ProjectSettings.Current.FileStorageBasePath) Then
                System.Windows.Application.Current.Dispatcher.Invoke(Sub()
                    System.Windows.MessageBox.Show(
                        "📁  File Storage Location Not Configured" & vbCrLf & vbCrLf &
                        "You must set a storage folder before uploading files." & vbCrLf & vbCrLf &
                        "To configure the storage location, go to:" & vbCrLf & vbCrLf &
                        "     ⚙️  Settings  ➜  Admin  ➜  File Location Settings" & vbCrLf & vbCrLf &
                        "Once a folder is set, uploads will work normally.",
                        "File Storage Location Not Set",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning)
                End Sub)
                Return ""   ' ← blocks the upload entirely
            End If
            ' ─────────────────────────────────────────────────────────────────────


            ' Use the user-configured storage root (falls back to app directory if not set)
            Dim storageRoot = ProjectSettings.GetFileStorageRoot()

            ' Build the sectioned subfolder path
            Dim sectionedSub = GetSectionedSubFolder(subFolder)
            Dim targetFolder = Path.Combine(storageRoot, sectionedSub)

            ' Create directory if it doesn't exist
            If Not Directory.Exists(targetFolder) Then
                Directory.CreateDirectory(targetFolder)
            End If

            Dim originalFileName = Path.GetFileName(sourcePath)
            ' Generate unique name: yyyyMMddHHmmss_ControlNo_OriginalName
            Dim uniqueName = $"{DateTime.Now:yyyyMMddHHmmss}_{controlNo.Replace("/", "-").Replace("\", "-")}_{originalFileName}"
            Dim destinationPath = Path.Combine(targetFolder, uniqueName)

            ' Copy the file
            File.Copy(sourcePath, destinationPath, True)

            ' Return relative path for database storage (relative to storage root)
            Return Path.Combine(sectionedSub, uniqueName)
        Catch ex As Exception
            Console.WriteLine($"CopyFileToDocuments Error ({subFolder}): " & ex.Message)
            Return ""
        End Try
    End Function
    Public Function CheckRFIDExists(rfid As String, Optional excludeControlNo As String = "") As String
        Try
            If MySQLDBConnect() = 1 Then
                Dim cleanRfid = rfid.Replace(" ", "").ToUpper()
                ' 1. Check interchangeability
                Dim q1 = $"SELECT ControlNo FROM interchangeability WHERE UPPER(REPLACE(RFID_tag, ' ', '')) = '{cleanRfid.Replace("'", "''")}'"
                If Not String.IsNullOrEmpty(excludeControlNo) Then
                    q1 &= $" AND ControlNo != '{excludeControlNo.Replace("'", "''")}'"
                End If

                Dim dt1 = ReadDatatable(q1)
                If dt1.Rows.Count > 0 Then
                    Return dt1.Rows(0)("ControlNo").ToString()
                End If

                ' 2. Check all type tables
                Dim tables = GetAllTypeTables()
                For Each tbl In tables
                    Dim q2 = $"SELECT ControlNo FROM `{tbl}` WHERE UPPER(REPLACE(RFID_tag, ' ', '')) = '{cleanRfid.Replace("'", "''")}'"
                    If Not String.IsNullOrEmpty(excludeControlNo) Then
                        q2 &= $" AND ControlNo != '{excludeControlNo.Replace("'", "''")}'"
                    End If
                    Dim dt2 = ReadDatatable(q2)
                    If dt2.Rows.Count > 0 Then
                        Return dt2.Rows(0)("ControlNo").ToString()
                    End If
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("CheckRFIDExists Error: " & ex.Message)
        End Try
        Return ""
    End Function
    ' --- CALIBRATION MAPPING HELPERS ---

    Public Function GetUniqueCategories() As List(Of String)
        Dim categories As New List(Of String)()
        Try
            Dim dt = ReadDatatable("SELECT DISTINCT Category FROM type_details")
            For Each row As DataRow In dt.Rows
                categories.Add(row("Category").ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("GetUniqueCategories Error: " & ex.Message)
        End Try
        Return categories
    End Function

    Public Function GetTypesByCategory(category As String) As List(Of String)
        Dim types As New List(Of String)()
        Try
            Dim dt = ReadDatatable($"SELECT TypeName FROM type_details WHERE Category = '{category.Replace("'", "''")}'")
            For Each row As DataRow In dt.Rows
                types.Add(row("TypeName").ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("GetTypesByCategory Error: " & ex.Message)
        End Try
        Return types
    End Function

    Public Function GetPrefixForType(category As String, typeName As String) As String
        Try
            Dim dt = ReadDatatable($"SELECT BasePrefix FROM type_details WHERE Category = '{category.Replace("'", "''")}' AND TypeName = '{typeName.Replace("'", "''")}'")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)("BasePrefix").ToString()
            End If
        Catch ex As Exception
            Console.WriteLine("GetPrefixForType Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function GetCategoriesForType(typeName As String) As List(Of String)
        Dim categories As New List(Of String)()
        Try
            ' First try exact match
            Dim dt = ReadDatatable($"SELECT calibration_category FROM calibration_template WHERE LOWER(type_name) = LOWER('{typeName.Replace("'", "''")}')")
            If dt.Rows.Count = 0 Then
                ' Fallback: try partial/fuzzy match on each word token
                Dim tokens = typeName.Split(" "c).Where(Function(t) t.Length > 3).ToArray()
                Dim whereClause = String.Join(" AND ", tokens.Select(Function(t) $"type_name LIKE '%{t.Replace("'", "''")}%'"))
                If Not String.IsNullOrEmpty(whereClause) Then
                    dt = ReadDatatable($"SELECT calibration_category FROM calibration_template WHERE {whereClause}")
                End If
            End If
            For Each row As DataRow In dt.Rows
                Dim cat = row("calibration_category").ToString()
                If Not String.IsNullOrEmpty(cat) Then categories.Add(cat)
            Next
        Catch ex As Exception
            Console.WriteLine("GetCategoriesForType Error: " & ex.Message)
        End Try
        Return categories
    End Function

    Public Function GetFormByCalibrationCategory(typeName As String, calibrationCategory As String) As String
        Try
            Dim dt = ReadDatatable($"SELECT form_name FROM calibration_template WHERE type_name = '{typeName.Replace("'", "''")}' AND calibration_category = '{calibrationCategory.Replace("'", "''")}'")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)("form_name").ToString()
            End If
        Catch ex As Exception
            Console.WriteLine("GetFormByCalibrationCategory Error: " & ex.Message)
        End Try
        Return ""
    End Function

    Public Function GetGroupCodes(category As String, typeName As String) As List(Of String)
        Dim codes As New List(Of String)()
        Try
            Dim tableName = If(category.Equals("Gauge", StringComparison.OrdinalIgnoreCase), "gauge_categorycontrol", "categorycontrol")
            Dim dt = ReadDatatable($"SELECT DISTINCT GroupCode FROM `{tableName}` WHERE InstrumentType = '{typeName.Replace("'", "''")}'")
            For Each row As DataRow In dt.Rows
                codes.Add(row("GroupCode").ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("GetGroupCodes Error: " & ex.Message)
        End Try
        Return codes
    End Function

    Public Function InsertCalibrationMapping(category As String, type_name As String, prefix As String, calibration_category As String, form_name As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO calibrationmapping (category, type_name, prefix, calibration_category, form_name) VALUES (@cat, @type, @pre, @range, @form)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@cat", category)
                    cmd.Parameters.AddWithValue("@type", type_name)
                    cmd.Parameters.AddWithValue("@pre", prefix)
                    cmd.Parameters.AddWithValue("@range", calibration_category)
                    cmd.Parameters.AddWithValue("@form", form_name)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertCalibrationMapping Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateCalibrationMapping(id As Integer, category As String, type_name As String, prefix As String, calibration_category As String, form_name As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE calibrationmapping SET category=@cat, type_name=@type, prefix=@pre, calibration_category=@range, form_name=@form WHERE id=@id"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@cat", category)
                    cmd.Parameters.AddWithValue("@type", type_name)
                    cmd.Parameters.AddWithValue("@pre", prefix)
                    cmd.Parameters.AddWithValue("@range", calibration_category)
                    cmd.Parameters.AddWithValue("@form", form_name)
                    cmd.Parameters.AddWithValue("@id", id)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateCalibrationMapping Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetInstrumentSizeByControlNo(controlNo As String) As String
        Try
            ' Check in interchangeability first (current cycle source)
            Dim query = $"SELECT SizeandRange FROM interchangeability WHERE ControlNo = @ctrl LIMIT 1"
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End If

            ' Fallback to department_list
            Dim fallbackQuery = $"SELECT SizeandRange FROM department_list WHERE `Control No` = @ctrl LIMIT 1"
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(fallbackQuery, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetInstrumentSizeByControlNo Error: " & ex.Message)
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Resolves a calibration form name based on control number prefix and instrument size.
    ''' </summary>
    Public Function GetCalibrationMappingForm(controlNo As String, size As String) As String
        Try
            If MySQLDBConnect() = 1 Then
                ' 1. Try to find mappings for this prefix
                Dim query = "SELECT prefix, calibration_category, form_name FROM calibrationmapping"
                Dim dt = ReadDatatable(query)

                Dim matches As New List(Of DataRow)()
                For Each row As DataRow In dt.Rows
                    Dim prefix = row("prefix").ToString()
                    If controlNo.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                        matches.Add(row)
                    End If
                Next

                If matches.Count = 1 Then
                    ' Single match for prefix - use it regardless of size categorisation details
                    Return matches(0)("form_name").ToString()
                ElseIf matches.Count > 1 Then
                    ' Multiple matches - try to find the best fit for the size
                    ' For now, prioritize longest prefix match
                    Dim bestMatchPrefix = ""
                    Dim matchedForm = ""
                    For Each row In matches
                        Dim prefix = row("prefix").ToString()
                        If prefix.Length > bestMatchPrefix.Length Then
                            bestMatchPrefix = prefix
                            matchedForm = row("form_name").ToString()
                        End If
                    Next
                    Return matchedForm
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("GetCalibrationMappingForm Error: " & ex.Message)
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from vernier_caliper_300.
    ''' </summary>
    Public Function GetVernierMasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM vernier_caliper_300 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetVernierMasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into vernier_caliper_300.
    ''' </summary>
    Public Function InsertVernierCalibration300(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal ext0 As Object(), ByVal ext20 As Object(), ByVal ext50 As Object(), ByVal ext100 As Object(), ByVal ext150 As Object(), ByVal ext200 As Object(), ByVal ext250 As Object(), ByVal ext300 As Object(),
                                          ByVal int0 As Object(), ByVal int20 As Object(), ByVal int50 As Object(), ByVal int100 As Object(), ByVal int150 As Object(), ByVal int200 As Object(), ByVal int250 As Object(), ByVal int300 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          ByVal calibMaster As String, ByVal masterUncertainty As Decimal, ByVal masterUncList As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                ' Future: Add columns for Calibration Master if needed (Currently disabled as per user request)
                ' Try
                '     Using alterCmd As New MySqlCommand("ALTER TABLE vernier_caliper_300 MODIFY COLUMN CalibrationMaster TEXT, ADD COLUMN IF NOT EXISTS CalibrationMasterUncertainty DECIMAL(18,10), ADD COLUMN IF NOT EXISTS MasterUncertaintiesList TEXT", _con)
                '         alterCmd.ExecuteNonQuery()
                '     End Using
                ' Catch ex As Exception
                ' End Try

                Dim query As String = "INSERT INTO vernier_caliper_300 " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Ext0_1, Ext0_2, Ext0_3, Ext20_1, Ext20_2, Ext20_3, Ext50_1, Ext50_2, Ext50_3, " &
                                   "Ext100_1, Ext100_2, Ext100_3, Ext150_1, Ext150_2, Ext150_3, Ext200_1, Ext200_2, Ext200_3, " &
                                   "Ext250_1, Ext250_2, Ext250_3, Ext300_1, Ext300_2, Ext300_3, " &
                                   "Int0_1, Int0_2, Int0_3, Int20_1, Int20_2, Int20_3, Int50_1, Int50_2, Int50_3, " &
                                   "Int100_1, Int100_2, Int100_3, Int150_1, Int150_2, Int150_3, Int200_1, Int200_2, Int200_3, " &
                                   "Int250_1, Int250_2, Int250_3, Int300_1, Int300_2, Int300_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@ext0_1, @ext0_2, @ext0_3, @ext20_1, @ext20_2, @ext20_3, @ext50_1, @ext50_2, @ext50_3, " &
                                   "@ext100_1, @ext100_2, @ext100_3, @ext150_1, @ext150_2, @ext150_3, @ext200_1, @ext200_2, @ext200_3, " &
                                   "@ext250_1, @ext250_2, @ext250_3, @ext300_1, @ext300_2, @ext300_3, " &
                                   "@int0_1, @int0_2, @int0_3, @int20_1, @int20_2, @int20_3, @int50_1, @int50_2, @int50_3, " &
                                   "@int100_1, @int100_2, @int100_3, @int150_1, @int150_2, @int150_3, @int200_1, @int200_2, @int200_3, " &
                                   "@int250_1, @int250_2, @int250_3, @int300_1, @int300_2, @int300_3, " &
                                   "@depthError, " &
                                   "@timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    ' Helper function to handle DBNull or specific values for parameters
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    ' --- Automatic & Header Data ---
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    ' --- EXTERNAL MEASUREMENTS (Arrays 0, 1, 2) ---
                    cmd.Parameters.AddWithValue("@ext0_1", SafeVal(ext0(0)))
                    cmd.Parameters.AddWithValue("@ext0_2", SafeVal(ext0(1)))
                    cmd.Parameters.AddWithValue("@ext0_3", SafeVal(ext0(2)))

                    cmd.Parameters.AddWithValue("@ext20_1", SafeVal(ext20(0)))
                    cmd.Parameters.AddWithValue("@ext20_2", SafeVal(ext20(1)))
                    cmd.Parameters.AddWithValue("@ext20_3", SafeVal(ext20(2)))

                    cmd.Parameters.AddWithValue("@ext50_1", SafeVal(ext50(0)))
                    cmd.Parameters.AddWithValue("@ext50_2", SafeVal(ext50(1)))
                    cmd.Parameters.AddWithValue("@ext50_3", SafeVal(ext50(2)))

                    cmd.Parameters.AddWithValue("@ext100_1", SafeVal(ext100(0)))
                    cmd.Parameters.AddWithValue("@ext100_2", SafeVal(ext100(1)))
                    cmd.Parameters.AddWithValue("@ext100_3", SafeVal(ext100(2)))

                    cmd.Parameters.AddWithValue("@ext150_1", SafeVal(ext150(0)))
                    cmd.Parameters.AddWithValue("@ext150_2", SafeVal(ext150(1)))
                    cmd.Parameters.AddWithValue("@ext150_3", SafeVal(ext150(2)))

                    cmd.Parameters.AddWithValue("@ext200_1", SafeVal(ext200(0)))
                    cmd.Parameters.AddWithValue("@ext200_2", SafeVal(ext200(1)))
                    cmd.Parameters.AddWithValue("@ext200_3", SafeVal(ext200(2)))

                    cmd.Parameters.AddWithValue("@ext250_1", SafeVal(ext250(0)))
                    cmd.Parameters.AddWithValue("@ext250_2", SafeVal(ext250(1)))
                    cmd.Parameters.AddWithValue("@ext250_3", SafeVal(ext250(2)))

                    cmd.Parameters.AddWithValue("@ext300_1", SafeVal(ext300(0)))
                    cmd.Parameters.AddWithValue("@ext300_2", SafeVal(ext300(1)))
                    cmd.Parameters.AddWithValue("@ext300_3", SafeVal(ext300(2)))

                    ' --- INTERNAL MEASUREMENTS (Arrays 0, 1, 2) ---
                    cmd.Parameters.AddWithValue("@int0_1", SafeVal(int0(0)))
                    cmd.Parameters.AddWithValue("@int0_2", SafeVal(int0(1)))
                    cmd.Parameters.AddWithValue("@int0_3", SafeVal(int0(2)))

                    cmd.Parameters.AddWithValue("@int20_1", SafeVal(int20(0)))
                    cmd.Parameters.AddWithValue("@int20_2", SafeVal(int20(1)))
                    cmd.Parameters.AddWithValue("@int20_3", SafeVal(int20(2)))

                    cmd.Parameters.AddWithValue("@int50_1", SafeVal(int50(0)))
                    cmd.Parameters.AddWithValue("@int50_2", SafeVal(int50(1)))
                    cmd.Parameters.AddWithValue("@int50_3", SafeVal(int50(2)))

                    cmd.Parameters.AddWithValue("@int100_1", SafeVal(int100(0)))
                    cmd.Parameters.AddWithValue("@int100_2", SafeVal(int100(1)))
                    cmd.Parameters.AddWithValue("@int100_3", SafeVal(int100(2)))

                    cmd.Parameters.AddWithValue("@int150_1", SafeVal(int150(0)))
                    cmd.Parameters.AddWithValue("@int150_2", SafeVal(int150(1)))
                    cmd.Parameters.AddWithValue("@int150_3", SafeVal(int150(2)))

                    cmd.Parameters.AddWithValue("@int200_1", SafeVal(int200(0)))
                    cmd.Parameters.AddWithValue("@int200_2", SafeVal(int200(1)))
                    cmd.Parameters.AddWithValue("@int200_3", SafeVal(int200(2)))

                    cmd.Parameters.AddWithValue("@int250_1", SafeVal(int250(0)))
                    cmd.Parameters.AddWithValue("@int250_2", SafeVal(int250(1)))
                    cmd.Parameters.AddWithValue("@int250_3", SafeVal(int250(2)))

                    cmd.Parameters.AddWithValue("@int300_1", SafeVal(int300(0)))
                    cmd.Parameters.AddWithValue("@int300_2", SafeVal(int300(1)))
                    cmd.Parameters.AddWithValue("@int300_3", SafeVal(int300(2)))

                    ' --- DEPTH ERROR (Single Value) ---
                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))

                    ' --- FOOTER DATA ---
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertVernierCalibration300 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetVernierCalibration300Data(controlNo As String, cycleName As String) As DataTable
        Try
            Dim query As String = "SELECT * FROM vernier_caliper_300 WHERE ControlNo = @control AND CycleName = @cycle AND RowType = 'RECORD' LIMIT 1"
            Dim dParam As MySqlParameter() = {
                New MySqlParameter("@control", controlNo),
                New MySqlParameter("@cycle", cycleName)
            }
            Return ReadDatatable(query, dParam)
        Catch ex As Exception
            Console.WriteLine("GetVernierCalibration300Data Error: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function UpdateVernierCalibration300(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal ext0 As Object(), ByVal ext20 As Object(), ByVal ext50 As Object(), ByVal ext100 As Object(), ByVal ext150 As Object(), ByVal ext200 As Object(), ByVal ext250 As Object(), ByVal ext300 As Object(),
                                           ByVal int0 As Object(), ByVal int20 As Object(), ByVal int50 As Object(), ByVal int100 As Object(), ByVal int150 As Object(), ByVal int200 As Object(), ByVal int250 As Object(), ByVal int300 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           ByVal calibMaster As String, ByVal masterUncertainty As Decimal, ByVal masterUncList As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE vernier_caliper_300 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Ext0_1=@ext0_1, Ext0_2=@ext0_2, Ext0_3=@ext0_3, Ext20_1=@ext20_1, Ext20_2=@ext20_2, Ext20_3=@ext20_3, Ext50_1=@ext50_1, Ext50_2=@ext50_2, Ext50_3=@ext50_3, " &
                                   "Ext100_1=@ext100_1, Ext100_2=@ext100_2, Ext100_3=@ext100_3, Ext150_1=@ext150_1, Ext150_2=@ext150_2, Ext150_3=@ext150_3, Ext200_1=@ext200_1, Ext200_2=@ext200_2, Ext200_3=@ext200_3, " &
                                   "Ext250_1=@ext250_1, Ext250_2=@ext250_2, Ext250_3=@ext250_3, Ext300_1=@ext300_1, Ext300_2=@ext300_2, Ext300_3=@ext300_3, " &
                                   "Int0_1=@int0_1, Int0_2=@int0_2, Int0_3=@int0_3, Int20_1=@int20_1, Int20_2=@int20_2, Int20_3=@int20_3, Int50_1=@int50_1, Int50_2=@int50_2, Int50_3=@int50_3, " &
                                   "Int100_1=@int100_1, Int100_2=@int100_2, Int100_3=@int100_3, Int150_1=@int150_1, Int150_2=@int150_2, Int150_3=@int150_3, Int200_1=@int200_1, Int200_2=@int200_2, Int200_3=@int200_3, " &
                                   "Int250_1=@int250_1, Int250_2=@int250_2, Int250_3=@int250_3, Int300_1=@int300_1, Int300_2=@int300_2, Int300_3=@int300_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@ext0_1", SafeVal(ext0(0)))
                    cmd.Parameters.AddWithValue("@ext0_2", SafeVal(ext0(1)))
                    cmd.Parameters.AddWithValue("@ext0_3", SafeVal(ext0(2)))
                    cmd.Parameters.AddWithValue("@ext20_1", SafeVal(ext20(0)))
                    cmd.Parameters.AddWithValue("@ext20_2", SafeVal(ext20(1)))
                    cmd.Parameters.AddWithValue("@ext20_3", SafeVal(ext20(2)))
                    cmd.Parameters.AddWithValue("@ext50_1", SafeVal(ext50(0)))
                    cmd.Parameters.AddWithValue("@ext50_2", SafeVal(ext50(1)))
                    cmd.Parameters.AddWithValue("@ext50_3", SafeVal(ext50(2)))
                    cmd.Parameters.AddWithValue("@ext100_1", SafeVal(ext100(0)))
                    cmd.Parameters.AddWithValue("@ext100_2", SafeVal(ext100(1)))
                    cmd.Parameters.AddWithValue("@ext100_3", SafeVal(ext100(2)))
                    cmd.Parameters.AddWithValue("@ext150_1", SafeVal(ext150(0)))
                    cmd.Parameters.AddWithValue("@ext150_2", SafeVal(ext150(1)))
                    cmd.Parameters.AddWithValue("@ext150_3", SafeVal(ext150(2)))
                    cmd.Parameters.AddWithValue("@ext200_1", SafeVal(ext200(0)))
                    cmd.Parameters.AddWithValue("@ext200_2", SafeVal(ext200(1)))
                    cmd.Parameters.AddWithValue("@ext200_3", SafeVal(ext200(2)))
                    cmd.Parameters.AddWithValue("@ext250_1", SafeVal(ext250(0)))
                    cmd.Parameters.AddWithValue("@ext250_2", SafeVal(ext250(1)))
                    cmd.Parameters.AddWithValue("@ext250_3", SafeVal(ext250(2)))
                    cmd.Parameters.AddWithValue("@ext300_1", SafeVal(ext300(0)))
                    cmd.Parameters.AddWithValue("@ext300_2", SafeVal(ext300(1)))
                    cmd.Parameters.AddWithValue("@ext300_3", SafeVal(ext300(2)))

                    cmd.Parameters.AddWithValue("@int0_1", SafeVal(int0(0)))
                    cmd.Parameters.AddWithValue("@int0_2", SafeVal(int0(1)))
                    cmd.Parameters.AddWithValue("@int0_3", SafeVal(int0(2)))
                    cmd.Parameters.AddWithValue("@int20_1", SafeVal(int20(0)))
                    cmd.Parameters.AddWithValue("@int20_2", SafeVal(int20(1)))
                    cmd.Parameters.AddWithValue("@int20_3", SafeVal(int20(2)))
                    cmd.Parameters.AddWithValue("@int50_1", SafeVal(int50(0)))
                    cmd.Parameters.AddWithValue("@int50_2", SafeVal(int50(1)))
                    cmd.Parameters.AddWithValue("@int50_3", SafeVal(int50(2)))
                    cmd.Parameters.AddWithValue("@int100_1", SafeVal(int100(0)))
                    cmd.Parameters.AddWithValue("@int100_2", SafeVal(int100(1)))
                    cmd.Parameters.AddWithValue("@int100_3", SafeVal(int100(2)))
                    cmd.Parameters.AddWithValue("@int150_1", SafeVal(int150(0)))
                    cmd.Parameters.AddWithValue("@int150_2", SafeVal(int150(1)))
                    cmd.Parameters.AddWithValue("@int150_3", SafeVal(int150(2)))
                    cmd.Parameters.AddWithValue("@int200_1", SafeVal(int200(0)))
                    cmd.Parameters.AddWithValue("@int200_2", SafeVal(int200(1)))
                    cmd.Parameters.AddWithValue("@int200_3", SafeVal(int200(2)))
                    cmd.Parameters.AddWithValue("@int250_1", SafeVal(int250(0)))
                    cmd.Parameters.AddWithValue("@int250_2", SafeVal(int250(1)))
                    cmd.Parameters.AddWithValue("@int250_3", SafeVal(int250(2)))
                    cmd.Parameters.AddWithValue("@int300_1", SafeVal(int300(0)))
                    cmd.Parameters.AddWithValue("@int300_2", SafeVal(int300(1)))
                    cmd.Parameters.AddWithValue("@int300_3", SafeVal(int300(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateVernierCalibration300 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from vernier_caliper_low_force.
    ''' </summary>
    Public Function GetVernierMasterLimitsLowForce(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM vernier_caliper_low_force WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetVernierMasterLimitsLowForce Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into vernier_caliper_low_force.
    ''' </summary>
    Public Function InsertVernierCalibrationLowForce(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs0 As Object(), ByVal obs20 As Object(), ByVal obs50 As Object(), ByVal obs100 As Object(), ByVal obs150 As Object(), ByVal obs190 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO vernier_caliper_low_force " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs0_1, Obs0_2, Obs0_3, Obs20_1, Obs20_2, Obs20_3, Obs50_1, Obs50_2, Obs50_3, " &
                                   "Obs100_1, Obs100_2, Obs100_3, Obs150_1, Obs150_2, Obs150_3, Obs190_1, Obs190_2, Obs190_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs0_1, @obs0_2, @obs0_3, @obs20_1, @obs20_2, @obs20_3, @obs50_1, @obs50_2, @obs50_3, " &
                                   "@obs100_1, @obs100_2, @obs100_3, @obs150_1, @obs150_2, @obs150_3, @obs190_1, @obs190_2, @obs190_3, " &
                                   "@depthError, " &
                                   "@timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))

                    cmd.Parameters.AddWithValue("@obs20_1", SafeVal(obs20(0)))
                    cmd.Parameters.AddWithValue("@obs20_2", SafeVal(obs20(1)))
                    cmd.Parameters.AddWithValue("@obs20_3", SafeVal(obs20(2)))

                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))

                    cmd.Parameters.AddWithValue("@obs100_1", SafeVal(obs100(0)))
                    cmd.Parameters.AddWithValue("@obs100_2", SafeVal(obs100(1)))
                    cmd.Parameters.AddWithValue("@obs100_3", SafeVal(obs100(2)))

                    cmd.Parameters.AddWithValue("@obs150_1", SafeVal(obs150(0)))
                    cmd.Parameters.AddWithValue("@obs150_2", SafeVal(obs150(1)))
                    cmd.Parameters.AddWithValue("@obs150_3", SafeVal(obs150(2)))

                    cmd.Parameters.AddWithValue("@obs190_1", SafeVal(obs190(0)))
                    cmd.Parameters.AddWithValue("@obs190_2", SafeVal(obs190(1)))
                    cmd.Parameters.AddWithValue("@obs190_3", SafeVal(obs190(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertVernierCalibrationLowForce Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetVernierCalibrationLowForceData(controlNo As String, cycleName As String) As DataTable
        Try
            Dim query As String = "SELECT * FROM vernier_caliper_low_force WHERE ControlNo = @control AND CycleName = @cycle AND RowType = 'RECORD' LIMIT 1"
            Dim dParam As MySqlParameter() = {
                New MySqlParameter("@control", controlNo),
                New MySqlParameter("@cycle", cycleName)
            }
            Return ReadDatatable(query, dParam)
        Catch ex As Exception
            Console.WriteLine("GetVernierCalibrationLowForceData Error: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function UpdateVernierCalibrationLowForce(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal obs0 As Object(), ByVal obs20 As Object(), ByVal obs50 As Object(), ByVal obs100 As Object(), ByVal obs150 As Object(), ByVal obs190 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE vernier_caliper_low_force SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs0_1=@obs0_1, Obs0_2=@obs0_2, Obs0_3=@obs0_3, Obs20_1=@obs20_1, Obs20_2=@obs20_2, Obs20_3=@obs20_3, Obs50_1=@obs50_1, Obs50_2=@obs50_2, Obs50_3=@obs50_3, " &
                                   "Obs100_1=@obs100_1, Obs100_2=@obs100_2, Obs100_3=@obs100_3, Obs150_1=@obs150_1, Obs150_2=@obs150_2, Obs150_3=@obs150_3, Obs190_1=@obs190_1, Obs190_2=@obs190_2, Obs190_3=@obs190_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))
                    cmd.Parameters.AddWithValue("@obs20_1", SafeVal(obs20(0)))
                    cmd.Parameters.AddWithValue("@obs20_2", SafeVal(obs20(1)))
                    cmd.Parameters.AddWithValue("@obs20_3", SafeVal(obs20(2)))
                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))
                    cmd.Parameters.AddWithValue("@obs100_1", SafeVal(obs100(0)))
                    cmd.Parameters.AddWithValue("@obs100_2", SafeVal(obs100(1)))
                    cmd.Parameters.AddWithValue("@obs100_3", SafeVal(obs100(2)))
                    cmd.Parameters.AddWithValue("@obs150_1", SafeVal(obs150(0)))
                    cmd.Parameters.AddWithValue("@obs150_2", SafeVal(obs150(1)))
                    cmd.Parameters.AddWithValue("@obs150_3", SafeVal(obs150(2)))
                    cmd.Parameters.AddWithValue("@obs190_1", SafeVal(obs190(0)))
                    cmd.Parameters.AddWithValue("@obs190_2", SafeVal(obs190(1)))
                    cmd.Parameters.AddWithValue("@obs190_3", SafeVal(obs190(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateVernierCalibrationLowForce Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from depth_vernier_calibration.
    ''' </summary>
    Public Function GetDepthVernierMasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM depth_vernier_calibration WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetDepthVernierMasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into depth_vernier_calibration.
    ''' </summary>
    Public Function InsertDepthVernierCalibration(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs0 As Object(), ByVal obs25 As Object(), ByVal obs50 As Object(), ByVal obs75 As Object(), ByVal obs100 As Object(), ByVal obs125 As Object(), ByVal obs150 As Object(), ByVal obs175 As Object(), ByVal obs200 As Object(), ByVal obs225 As Object(), ByVal obs250 As Object(), ByVal obs275 As Object(), ByVal obs300 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO depth_vernier_calibration " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs0_1, Obs0_2, Obs0_3, Obs25_1, Obs25_2, Obs25_3, Obs50_1, Obs50_2, Obs50_3, Obs75_1, Obs75_2, Obs75_3, " &
                                   "Obs100_1, Obs100_2, Obs100_3, Obs125_1, Obs125_2, Obs125_3, Obs150_1, Obs150_2, Obs150_3, Obs175_1, Obs175_2, Obs175_3, " &
                                   "Obs200_1, Obs200_2, Obs200_3, Obs225_1, Obs225_2, Obs225_3, Obs250_1, Obs250_2, Obs250_3, Obs275_1, Obs275_2, Obs275_3, Obs300_1, Obs300_2, Obs300_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs0_1, @obs0_2, @obs0_3, @obs25_1, @obs25_2, @obs25_3, @obs50_1, @obs50_2, @obs50_3, @obs75_1, @obs75_2, @obs75_3, " &
                                   "@obs100_1, @obs100_2, @obs100_3, @obs125_1, @obs125_2, @obs125_3, @obs150_1, @obs150_2, @obs150_3, @obs175_1, @obs175_2, @obs175_3, " &
                                   "@obs200_1, @obs200_2, @obs200_3, @obs225_1, @obs225_2, @obs225_3, @obs250_1, @obs250_2, @obs250_3, @obs275_1, @obs275_2, @obs275_3, @obs300_1, @obs300_2, @obs300_3, " &
                                   "@depthError, " &
                                   "@timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))

                    cmd.Parameters.AddWithValue("@obs25_1", SafeVal(obs25(0)))
                    cmd.Parameters.AddWithValue("@obs25_2", SafeVal(obs25(1)))
                    cmd.Parameters.AddWithValue("@obs25_3", SafeVal(obs25(2)))

                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))

                    cmd.Parameters.AddWithValue("@obs75_1", SafeVal(obs75(0)))
                    cmd.Parameters.AddWithValue("@obs75_2", SafeVal(obs75(1)))
                    cmd.Parameters.AddWithValue("@obs75_3", SafeVal(obs75(2)))

                    cmd.Parameters.AddWithValue("@obs100_1", SafeVal(obs100(0)))
                    cmd.Parameters.AddWithValue("@obs100_2", SafeVal(obs100(1)))
                    cmd.Parameters.AddWithValue("@obs100_3", SafeVal(obs100(2)))

                    cmd.Parameters.AddWithValue("@obs125_1", SafeVal(obs125(0)))
                    cmd.Parameters.AddWithValue("@obs125_2", SafeVal(obs125(1)))
                    cmd.Parameters.AddWithValue("@obs125_3", SafeVal(obs125(2)))

                    cmd.Parameters.AddWithValue("@obs150_1", SafeVal(obs150(0)))
                    cmd.Parameters.AddWithValue("@obs150_2", SafeVal(obs150(1)))
                    cmd.Parameters.AddWithValue("@obs150_3", SafeVal(obs150(2)))

                    cmd.Parameters.AddWithValue("@obs175_1", SafeVal(obs175(0)))
                    cmd.Parameters.AddWithValue("@obs175_2", SafeVal(obs175(1)))
                    cmd.Parameters.AddWithValue("@obs175_3", SafeVal(obs175(2)))

                    cmd.Parameters.AddWithValue("@obs200_1", SafeVal(obs200(0)))
                    cmd.Parameters.AddWithValue("@obs200_2", SafeVal(obs200(1)))
                    cmd.Parameters.AddWithValue("@obs200_3", SafeVal(obs200(2)))

                    cmd.Parameters.AddWithValue("@obs225_1", SafeVal(obs225(0)))
                    cmd.Parameters.AddWithValue("@obs225_2", SafeVal(obs225(1)))
                    cmd.Parameters.AddWithValue("@obs225_3", SafeVal(obs225(2)))

                    cmd.Parameters.AddWithValue("@obs250_1", SafeVal(obs250(0)))
                    cmd.Parameters.AddWithValue("@obs250_2", SafeVal(obs250(1)))
                    cmd.Parameters.AddWithValue("@obs250_3", SafeVal(obs250(2)))

                    cmd.Parameters.AddWithValue("@obs275_1", SafeVal(obs275(0)))
                    cmd.Parameters.AddWithValue("@obs275_2", SafeVal(obs275(1)))
                    cmd.Parameters.AddWithValue("@obs275_3", SafeVal(obs275(2)))

                    cmd.Parameters.AddWithValue("@obs300_1", SafeVal(obs300(0)))
                    cmd.Parameters.AddWithValue("@obs300_2", SafeVal(obs300(1)))
                    cmd.Parameters.AddWithValue("@obs300_3", SafeVal(obs300(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertDepthVernierCalibration Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetDepthVernierCalibrationData(controlNo As String, cycleName As String) As DataTable
        Try
            Dim query As String = "SELECT * FROM depth_vernier_calibration WHERE ControlNo = @control AND CycleName = @cycle AND RowType = 'RECORD' LIMIT 1"
            Dim dParam As MySqlParameter() = {
                New MySqlParameter("@control", controlNo),
                New MySqlParameter("@cycle", cycleName)
            }
            Return ReadDatatable(query, dParam)
        Catch ex As Exception
            Console.WriteLine("GetDepthVernierCalibrationData Error: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function UpdateDepthVernierCalibration(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal obs0 As Object(), ByVal obs25 As Object(), ByVal obs50 As Object(), ByVal obs75 As Object(), ByVal obs100 As Object(), ByVal obs125 As Object(), ByVal obs150 As Object(), ByVal obs175 As Object(), ByVal obs200 As Object(), ByVal obs225 As Object(), ByVal obs250 As Object(), ByVal obs275 As Object(), ByVal obs300 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE depth_vernier_calibration SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs0_1=@obs0_1, Obs0_2=@obs0_2, Obs0_3=@obs0_3, Obs25_1=@obs25_1, Obs25_2=@obs25_2, Obs25_3=@obs25_3, Obs50_1=@obs50_1, Obs50_2=@obs50_2, Obs50_3=@obs50_3, Obs75_1=@obs75_1, Obs75_2=@obs75_2, Obs75_3=@obs75_3, " &
                                   "Obs100_1=@obs100_1, Obs100_2=@obs100_2, Obs100_3=@obs100_3, Obs125_1=@obs125_1, Obs125_2=@obs125_2, Obs125_3=@obs125_3, Obs150_1=@obs150_1, Obs150_2=@obs150_2, Obs150_3=@obs150_3, Obs175_1=@obs175_1, Obs175_2=@obs175_2, Obs175_3=@obs175_3, " &
                                   "Obs200_1=@obs200_1, Obs200_2=@obs200_2, Obs200_3=@obs200_3, Obs225_1=@obs225_1, Obs225_2=@obs225_2, Obs225_3=@obs225_3, Obs250_1=@obs250_1, Obs250_2=@obs250_2, Obs250_3=@obs250_3, Obs275_1=@obs275_1, Obs275_2=@obs275_2, Obs275_3=@obs275_3, Obs300_1=@obs300_1, Obs300_2=@obs300_2, Obs300_3=@obs300_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))
                    cmd.Parameters.AddWithValue("@obs25_1", SafeVal(obs25(0)))
                    cmd.Parameters.AddWithValue("@obs25_2", SafeVal(obs25(1)))
                    cmd.Parameters.AddWithValue("@obs25_3", SafeVal(obs25(2)))
                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))
                    cmd.Parameters.AddWithValue("@obs75_1", SafeVal(obs75(0)))
                    cmd.Parameters.AddWithValue("@obs75_2", SafeVal(obs75(1)))
                    cmd.Parameters.AddWithValue("@obs75_3", SafeVal(obs75(2)))

                    cmd.Parameters.AddWithValue("@obs100_1", SafeVal(obs100(0)))
                    cmd.Parameters.AddWithValue("@obs100_2", SafeVal(obs100(1)))
                    cmd.Parameters.AddWithValue("@obs100_3", SafeVal(obs100(2)))
                    cmd.Parameters.AddWithValue("@obs125_1", SafeVal(obs125(0)))
                    cmd.Parameters.AddWithValue("@obs125_2", SafeVal(obs125(1)))
                    cmd.Parameters.AddWithValue("@obs125_3", SafeVal(obs125(2)))
                    cmd.Parameters.AddWithValue("@obs150_1", SafeVal(obs150(0)))
                    cmd.Parameters.AddWithValue("@obs150_2", SafeVal(obs150(1)))
                    cmd.Parameters.AddWithValue("@obs150_3", SafeVal(obs150(2)))
                    cmd.Parameters.AddWithValue("@obs175_1", SafeVal(obs175(0)))
                    cmd.Parameters.AddWithValue("@obs175_2", SafeVal(obs175(1)))
                    cmd.Parameters.AddWithValue("@obs175_3", SafeVal(obs175(2)))

                    cmd.Parameters.AddWithValue("@obs200_1", SafeVal(obs200(0)))
                    cmd.Parameters.AddWithValue("@obs200_2", SafeVal(obs200(1)))
                    cmd.Parameters.AddWithValue("@obs200_3", SafeVal(obs200(2)))
                    cmd.Parameters.AddWithValue("@obs225_1", SafeVal(obs225(0)))
                    cmd.Parameters.AddWithValue("@obs225_2", SafeVal(obs225(1)))
                    cmd.Parameters.AddWithValue("@obs225_3", SafeVal(obs225(2)))
                    cmd.Parameters.AddWithValue("@obs250_1", SafeVal(obs250(0)))
                    cmd.Parameters.AddWithValue("@obs250_2", SafeVal(obs250(1)))
                    cmd.Parameters.AddWithValue("@obs250_3", SafeVal(obs250(2)))
                    cmd.Parameters.AddWithValue("@obs275_1", SafeVal(obs275(0)))
                    cmd.Parameters.AddWithValue("@obs275_2", SafeVal(obs275(1)))
                    cmd.Parameters.AddWithValue("@obs275_3", SafeVal(obs275(2)))
                    cmd.Parameters.AddWithValue("@obs300_1", SafeVal(obs300(0)))
                    cmd.Parameters.AddWithValue("@obs300_2", SafeVal(obs300(1)))
                    cmd.Parameters.AddWithValue("@obs300_3", SafeVal(obs300(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateDepthVernierCalibration Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from vernier_caliper_600.
    ''' </summary>
    Public Function GetVernierMasterLimits600(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM vernier_caliper_600 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetVernierMasterLimits600 Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into vernier_caliper_600.
    ''' </summary>
    Public Function InsertVernierCalibration600(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal ext0 As Object(), ByVal ext20 As Object(), ByVal ext50 As Object(), ByVal ext100 As Object(), ByVal ext150 As Object(), ByVal ext200 As Object(), ByVal ext250 As Object(), ByVal ext300 As Object(),
                                          ByVal ext350 As Object(), ByVal ext400 As Object(), ByVal ext450 As Object(), ByVal ext500 As Object(), ByVal ext550 As Object(), ByVal ext600 As Object(),
                                          ByVal int0 As Object(), ByVal int20 As Object(), ByVal int50 As Object(), ByVal int100 As Object(), ByVal int150 As Object(), ByVal int200 As Object(), ByVal int250 As Object(), ByVal int300 As Object(),
                                          ByVal int350 As Object(), ByVal int400 As Object(), ByVal int450 As Object(), ByVal int500 As Object(), ByVal int550 As Object(), ByVal int600 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          ByVal calibMaster As String, ByVal masterUncertainty As Decimal, ByVal masterUncList As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO vernier_caliper_600 " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Ext0_1, Ext0_2, Ext0_3, Ext20_1, Ext20_2, Ext20_3, Ext50_1, Ext50_2, Ext50_3, " &
                                   "Ext100_1, Ext100_2, Ext100_3, Ext150_1, Ext150_2, Ext150_3, Ext200_1, Ext200_2, Ext200_3, " &
                                   "Ext250_1, Ext250_2, Ext250_3, Ext300_1, Ext300_2, Ext300_3, " &
                                   "Ext350_1, Ext350_2, Ext350_3, Ext400_1, Ext400_2, Ext400_3, Ext450_1, Ext450_2, Ext450_3, " &
                                   "Ext500_1, Ext500_2, Ext500_3, Ext550_1, Ext550_2, Ext550_3, Ext600_1, Ext600_2, Ext600_3, " &
                                   "Int0_1, Int0_2, Int0_3, Int20_1, Int20_2, Int20_3, Int50_1, Int50_2, Int50_3, " &
                                   "Int100_1, Int100_2, Int100_3, Int150_1, Int150_2, Int150_3, Int200_1, Int200_2, Int200_3, " &
                                   "Int250_1, Int250_2, Int250_3, Int300_1, Int300_2, Int300_3, " &
                                   "Int350_1, Int350_2, Int350_3, Int400_1, Int400_2, Int400_3, Int450_1, Int450_2, Int450_3, " &
                                   "Int500_1, Int500_2, Int500_3, Int550_1, Int550_2, Int550_3, Int600_1, Int600_2, Int600_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@ext0_1, @ext0_2, @ext0_3, @ext20_1, @ext20_2, @ext20_3, @ext50_1, @ext50_2, @ext50_3, " &
                                   "@ext100_1, @ext100_2, @ext100_3, @ext150_1, @ext150_2, @ext150_3, @ext200_1, @ext200_2, @ext200_3, " &
                                   "@ext250_1, @ext250_2, @ext250_3, @ext300_1, @ext300_2, @ext300_3, " &
                                   "@ext350_1, @ext350_2, @ext350_3, @ext400_1, @ext400_2, @ext400_3, @ext450_1, @ext450_2, @ext450_3, " &
                                   "@ext500_1, @ext500_2, @ext500_3, @ext550_1, @ext550_2, @ext550_3, @ext600_1, @ext600_2, @ext600_3, " &
                                   "@int0_1, @int0_2, @int0_3, @int20_1, @int20_2, @int20_3, @int50_1, @int50_2, @int50_3, " &
                                   "@int100_1, @int100_2, @int100_3, @int150_1, @int150_2, @int150_3, @int200_1, @int200_2, @int200_3, " &
                                   "@int250_1, @int250_2, @int250_3, @int300_1, @int300_2, @int300_3, " &
                                   "@int350_1, @int350_2, @int350_3, @int400_1, @int400_2, @int400_3, @int450_1, @int450_2, @int450_3, " &
                                   "@int500_1, @int500_2, @int500_3, @int550_1, @int550_2, @int550_3, @int600_1, @int600_2, @int600_3, " &
                                   "@depthError, " &
                                   "@timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@ext0_1", SafeVal(ext0(0)))
                    cmd.Parameters.AddWithValue("@ext0_2", SafeVal(ext0(1)))
                    cmd.Parameters.AddWithValue("@ext0_3", SafeVal(ext0(2)))
                    cmd.Parameters.AddWithValue("@ext20_1", SafeVal(ext20(0)))
                    cmd.Parameters.AddWithValue("@ext20_2", SafeVal(ext20(1)))
                    cmd.Parameters.AddWithValue("@ext20_3", SafeVal(ext20(2)))
                    cmd.Parameters.AddWithValue("@ext50_1", SafeVal(ext50(0)))
                    cmd.Parameters.AddWithValue("@ext50_2", SafeVal(ext50(1)))
                    cmd.Parameters.AddWithValue("@ext50_3", SafeVal(ext50(2)))
                    cmd.Parameters.AddWithValue("@ext100_1", SafeVal(ext100(0)))
                    cmd.Parameters.AddWithValue("@ext100_2", SafeVal(ext100(1)))
                    cmd.Parameters.AddWithValue("@ext100_3", SafeVal(ext100(2)))
                    cmd.Parameters.AddWithValue("@ext150_1", SafeVal(ext150(0)))
                    cmd.Parameters.AddWithValue("@ext150_2", SafeVal(ext150(1)))
                    cmd.Parameters.AddWithValue("@ext150_3", SafeVal(ext150(2)))
                    cmd.Parameters.AddWithValue("@ext200_1", SafeVal(ext200(0)))
                    cmd.Parameters.AddWithValue("@ext200_2", SafeVal(ext200(1)))
                    cmd.Parameters.AddWithValue("@ext200_3", SafeVal(ext200(2)))
                    cmd.Parameters.AddWithValue("@ext250_1", SafeVal(ext250(0)))
                    cmd.Parameters.AddWithValue("@ext250_2", SafeVal(ext250(1)))
                    cmd.Parameters.AddWithValue("@ext250_3", SafeVal(ext250(2)))
                    cmd.Parameters.AddWithValue("@ext300_1", SafeVal(ext300(0)))
                    cmd.Parameters.AddWithValue("@ext300_2", SafeVal(ext300(1)))
                    cmd.Parameters.AddWithValue("@ext300_3", SafeVal(ext300(2)))
                    cmd.Parameters.AddWithValue("@ext350_1", SafeVal(ext350(0)))
                    cmd.Parameters.AddWithValue("@ext350_2", SafeVal(ext350(1)))
                    cmd.Parameters.AddWithValue("@ext350_3", SafeVal(ext350(2)))
                    cmd.Parameters.AddWithValue("@ext400_1", SafeVal(ext400(0)))
                    cmd.Parameters.AddWithValue("@ext400_2", SafeVal(ext400(1)))
                    cmd.Parameters.AddWithValue("@ext400_3", SafeVal(ext400(2)))
                    cmd.Parameters.AddWithValue("@ext450_1", SafeVal(ext450(0)))
                    cmd.Parameters.AddWithValue("@ext450_2", SafeVal(ext450(1)))
                    cmd.Parameters.AddWithValue("@ext450_3", SafeVal(ext450(2)))
                    cmd.Parameters.AddWithValue("@ext500_1", SafeVal(ext500(0)))
                    cmd.Parameters.AddWithValue("@ext500_2", SafeVal(ext500(1)))
                    cmd.Parameters.AddWithValue("@ext500_3", SafeVal(ext500(2)))
                    cmd.Parameters.AddWithValue("@ext550_1", SafeVal(ext550(0)))
                    cmd.Parameters.AddWithValue("@ext550_2", SafeVal(ext550(1)))
                    cmd.Parameters.AddWithValue("@ext550_3", SafeVal(ext550(2)))
                    cmd.Parameters.AddWithValue("@ext600_1", SafeVal(ext600(0)))
                    cmd.Parameters.AddWithValue("@ext600_2", SafeVal(ext600(1)))
                    cmd.Parameters.AddWithValue("@ext600_3", SafeVal(ext600(2)))

                    cmd.Parameters.AddWithValue("@int0_1", SafeVal(int0(0)))
                    cmd.Parameters.AddWithValue("@int0_2", SafeVal(int0(1)))
                    cmd.Parameters.AddWithValue("@int0_3", SafeVal(int0(2)))
                    cmd.Parameters.AddWithValue("@int20_1", SafeVal(int20(0)))
                    cmd.Parameters.AddWithValue("@int20_2", SafeVal(int20(1)))
                    cmd.Parameters.AddWithValue("@int20_3", SafeVal(int20(2)))
                    cmd.Parameters.AddWithValue("@int50_1", SafeVal(int50(0)))
                    cmd.Parameters.AddWithValue("@int50_2", SafeVal(int50(1)))
                    cmd.Parameters.AddWithValue("@int50_3", SafeVal(int50(2)))
                    cmd.Parameters.AddWithValue("@int100_1", SafeVal(int100(0)))
                    cmd.Parameters.AddWithValue("@int100_2", SafeVal(int100(1)))
                    cmd.Parameters.AddWithValue("@int100_3", SafeVal(int100(2)))
                    cmd.Parameters.AddWithValue("@int150_1", SafeVal(int150(0)))
                    cmd.Parameters.AddWithValue("@int150_2", SafeVal(int150(1)))
                    cmd.Parameters.AddWithValue("@int150_3", SafeVal(int150(2)))
                    cmd.Parameters.AddWithValue("@int200_1", SafeVal(int200(0)))
                    cmd.Parameters.AddWithValue("@int200_2", SafeVal(int200(1)))
                    cmd.Parameters.AddWithValue("@int200_3", SafeVal(int200(2)))
                    cmd.Parameters.AddWithValue("@int250_1", SafeVal(int250(0)))
                    cmd.Parameters.AddWithValue("@int250_2", SafeVal(int250(1)))
                    cmd.Parameters.AddWithValue("@int250_3", SafeVal(int250(2)))
                    cmd.Parameters.AddWithValue("@int300_1", SafeVal(int300(0)))
                    cmd.Parameters.AddWithValue("@int300_2", SafeVal(int300(1)))
                    cmd.Parameters.AddWithValue("@int300_3", SafeVal(int300(2)))
                    cmd.Parameters.AddWithValue("@int350_1", SafeVal(int350(0)))
                    cmd.Parameters.AddWithValue("@int350_2", SafeVal(int350(1)))
                    cmd.Parameters.AddWithValue("@int350_3", SafeVal(int350(2)))
                    cmd.Parameters.AddWithValue("@int400_1", SafeVal(int400(0)))
                    cmd.Parameters.AddWithValue("@int400_2", SafeVal(int400(1)))
                    cmd.Parameters.AddWithValue("@int400_3", SafeVal(int400(2)))
                    cmd.Parameters.AddWithValue("@int450_1", SafeVal(int450(0)))
                    cmd.Parameters.AddWithValue("@int450_2", SafeVal(int450(1)))
                    cmd.Parameters.AddWithValue("@int450_3", SafeVal(int450(2)))
                    cmd.Parameters.AddWithValue("@int500_1", SafeVal(int500(0)))
                    cmd.Parameters.AddWithValue("@int500_2", SafeVal(int500(1)))
                    cmd.Parameters.AddWithValue("@int500_3", SafeVal(int500(2)))
                    cmd.Parameters.AddWithValue("@int550_1", SafeVal(int550(0)))
                    cmd.Parameters.AddWithValue("@int550_2", SafeVal(int550(1)))
                    cmd.Parameters.AddWithValue("@int550_3", SafeVal(int550(2)))
                    cmd.Parameters.AddWithValue("@int600_1", SafeVal(int600(0)))
                    cmd.Parameters.AddWithValue("@int600_2", SafeVal(int600(1)))
                    cmd.Parameters.AddWithValue("@int600_3", SafeVal(int600(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertVernierCalibration600 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetVernierCalibration600Data(controlNo As String, cycleName As String) As DataTable
        Try
            Dim query As String = "SELECT * FROM vernier_caliper_600 WHERE ControlNo = @control AND CycleName = @cycle AND RowType = 'RECORD' LIMIT 1"
            Dim dParam As MySqlParameter() = {
                New MySqlParameter("@control", controlNo),
                New MySqlParameter("@cycle", cycleName)
            }
            Return ReadDatatable(query, dParam)
        Catch ex As Exception
            Console.WriteLine("GetVernierCalibration600Data Error: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function UpdateVernierCalibration600(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal ext0 As Object(), ByVal ext20 As Object(), ByVal ext50 As Object(), ByVal ext100 As Object(), ByVal ext150 As Object(), ByVal ext200 As Object(), ByVal ext250 As Object(), ByVal ext300 As Object(),
                                           ByVal ext350 As Object(), ByVal ext400 As Object(), ByVal ext450 As Object(), ByVal ext500 As Object(), ByVal ext550 As Object(), ByVal ext600 As Object(),
                                           ByVal int0 As Object(), ByVal int20 As Object(), ByVal int50 As Object(), ByVal int100 As Object(), ByVal int150 As Object(), ByVal int200 As Object(), ByVal int250 As Object(), ByVal int300 As Object(),
                                           ByVal int350 As Object(), ByVal int400 As Object(), ByVal int450 As Object(), ByVal int500 As Object(), ByVal int550 As Object(), ByVal int600 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           ByVal calibMaster As String, ByVal masterUncertainty As Decimal, ByVal masterUncList As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE vernier_caliper_600 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Ext0_1=@ext0_1, Ext0_2=@ext0_2, Ext0_3=@ext0_3, Ext20_1=@ext20_1, Ext20_2=@ext20_2, Ext20_3=@ext20_3, Ext50_1=@ext50_1, Ext50_2=@ext50_2, Ext50_3=@ext50_3, " &
                                   "Ext100_1=@ext100_1, Ext100_2=@ext100_2, Ext100_3=@ext100_3, Ext150_1=@ext150_1, Ext150_2=@ext150_2, Ext150_3=@ext150_3, Ext200_1=@ext200_1, Ext200_2=@ext200_2, Ext200_3=@ext200_3, " &
                                   "Ext250_1=@ext250_1, Ext250_2=@ext250_2, Ext250_3=@ext250_3, Ext300_1=@ext300_1, Ext300_2=@ext300_2, Ext300_3=@ext300_3, " &
                                   "Ext350_1=@ext350_1, Ext350_2=@ext350_2, Ext350_3=@ext350_3, Ext400_1=@ext400_1, Ext400_2=@ext400_2, Ext400_3=@ext400_3, Ext450_1=@ext450_1, Ext450_2=@ext450_2, Ext450_3=@ext450_3, " &
                                   "Ext500_1=@ext500_1, Ext500_2=@ext500_2, Ext500_3=@ext500_3, Ext550_1=@ext550_1, Ext550_2=@ext550_2, Ext550_3=@ext550_3, Ext600_1=@ext600_1, Ext600_2=@ext600_2, Ext600_3=@ext600_3, " &
                                   "Int0_1=@int0_1, Int0_2=@int0_2, Int0_3=@int0_3, Int20_1=@int20_1, Int20_2=@int20_2, Int20_3=@int20_3, Int50_1=@int50_1, Int50_2=@int50_2, Int50_3=@int50_3, " &
                                   "Int100_1=@int100_1, Int100_2=@int100_2, Int100_3=@int100_3, Int150_1=@int150_1, Int150_2=@int150_2, Int150_3=@int150_3, Int200_1=@int200_1, Int200_2=@int200_2, Int200_3=@int200_3, " &
                                   "Int250_1=@int250_1, Int250_2=@int250_2, Int250_3=@int250_3, Int300_1=@int300_1, Int300_2=@int300_2, Int300_3=@int300_3, " &
                                   "Int350_1=@int350_1, Int350_2=@int350_2, Int350_3=@int350_3, Int400_1=@int400_1, Int400_2=@int400_2, Int400_3=@int400_3, Int450_1=@int450_1, Int450_2=@int450_2, Int450_3=@int450_3, " &
                                   "Int500_1=@int500_1, Int500_2=@int500_2, Int500_3=@int500_3, Int550_1=@int550_1, Int550_2=@int550_2, Int550_3=@int550_3, Int600_1=@int600_1, Int600_2=@int600_2, Int600_3=@int600_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@ext0_1", SafeVal(ext0(0)))
                    cmd.Parameters.AddWithValue("@ext0_2", SafeVal(ext0(1)))
                    cmd.Parameters.AddWithValue("@ext0_3", SafeVal(ext0(2)))
                    cmd.Parameters.AddWithValue("@ext20_1", SafeVal(ext20(0)))
                    cmd.Parameters.AddWithValue("@ext20_2", SafeVal(ext20(1)))
                    cmd.Parameters.AddWithValue("@ext20_3", SafeVal(ext20(2)))
                    cmd.Parameters.AddWithValue("@ext50_1", SafeVal(ext50(0)))
                    cmd.Parameters.AddWithValue("@ext50_2", SafeVal(ext50(1)))
                    cmd.Parameters.AddWithValue("@ext50_3", SafeVal(ext50(2)))
                    cmd.Parameters.AddWithValue("@ext100_1", SafeVal(ext100(0)))
                    cmd.Parameters.AddWithValue("@ext100_2", SafeVal(ext100(1)))
                    cmd.Parameters.AddWithValue("@ext100_3", SafeVal(ext100(2)))
                    cmd.Parameters.AddWithValue("@ext150_1", SafeVal(ext150(0)))
                    cmd.Parameters.AddWithValue("@ext150_2", SafeVal(ext150(1)))
                    cmd.Parameters.AddWithValue("@ext150_3", SafeVal(ext150(2)))
                    cmd.Parameters.AddWithValue("@ext200_1", SafeVal(ext200(0)))
                    cmd.Parameters.AddWithValue("@ext200_2", SafeVal(ext200(1)))
                    cmd.Parameters.AddWithValue("@ext200_3", SafeVal(ext200(2)))
                    cmd.Parameters.AddWithValue("@ext250_1", SafeVal(ext250(0)))
                    cmd.Parameters.AddWithValue("@ext250_2", SafeVal(ext250(1)))
                    cmd.Parameters.AddWithValue("@ext250_3", SafeVal(ext250(2)))
                    cmd.Parameters.AddWithValue("@ext300_1", SafeVal(ext300(0)))
                    cmd.Parameters.AddWithValue("@ext300_2", SafeVal(ext300(1)))
                    cmd.Parameters.AddWithValue("@ext300_3", SafeVal(ext300(2)))
                    cmd.Parameters.AddWithValue("@ext350_1", SafeVal(ext350(0)))
                    cmd.Parameters.AddWithValue("@ext350_2", SafeVal(ext350(1)))
                    cmd.Parameters.AddWithValue("@ext350_3", SafeVal(ext350(2)))
                    cmd.Parameters.AddWithValue("@ext400_1", SafeVal(ext400(0)))
                    cmd.Parameters.AddWithValue("@ext400_2", SafeVal(ext400(1)))
                    cmd.Parameters.AddWithValue("@ext400_3", SafeVal(ext400(2)))
                    cmd.Parameters.AddWithValue("@ext450_1", SafeVal(ext450(0)))
                    cmd.Parameters.AddWithValue("@ext450_2", SafeVal(ext450(1)))
                    cmd.Parameters.AddWithValue("@ext450_3", SafeVal(ext450(2)))
                    cmd.Parameters.AddWithValue("@ext500_1", SafeVal(ext500(0)))
                    cmd.Parameters.AddWithValue("@ext500_2", SafeVal(ext500(1)))
                    cmd.Parameters.AddWithValue("@ext500_3", SafeVal(ext500(2)))
                    cmd.Parameters.AddWithValue("@ext550_1", SafeVal(ext550(0)))
                    cmd.Parameters.AddWithValue("@ext550_2", SafeVal(ext550(1)))
                    cmd.Parameters.AddWithValue("@ext550_3", SafeVal(ext550(2)))
                    cmd.Parameters.AddWithValue("@ext600_1", SafeVal(ext600(0)))
                    cmd.Parameters.AddWithValue("@ext600_2", SafeVal(ext600(1)))
                    cmd.Parameters.AddWithValue("@ext600_3", SafeVal(ext600(2)))

                    cmd.Parameters.AddWithValue("@int0_1", SafeVal(int0(0)))
                    cmd.Parameters.AddWithValue("@int0_2", SafeVal(int0(1)))
                    cmd.Parameters.AddWithValue("@int0_3", SafeVal(int0(2)))
                    cmd.Parameters.AddWithValue("@int20_1", SafeVal(int20(0)))
                    cmd.Parameters.AddWithValue("@int20_2", SafeVal(int20(1)))
                    cmd.Parameters.AddWithValue("@int20_3", SafeVal(int20(2)))
                    cmd.Parameters.AddWithValue("@int50_1", SafeVal(int50(0)))
                    cmd.Parameters.AddWithValue("@int50_2", SafeVal(int50(1)))
                    cmd.Parameters.AddWithValue("@int50_3", SafeVal(int50(2)))
                    cmd.Parameters.AddWithValue("@int100_1", SafeVal(int100(0)))
                    cmd.Parameters.AddWithValue("@int100_2", SafeVal(int100(1)))
                    cmd.Parameters.AddWithValue("@int100_3", SafeVal(int100(2)))
                    cmd.Parameters.AddWithValue("@int150_1", SafeVal(int150(0)))
                    cmd.Parameters.AddWithValue("@int150_2", SafeVal(int150(1)))
                    cmd.Parameters.AddWithValue("@int150_3", SafeVal(int150(2)))
                    cmd.Parameters.AddWithValue("@int200_1", SafeVal(int200(0)))
                    cmd.Parameters.AddWithValue("@int200_2", SafeVal(int200(1)))
                    cmd.Parameters.AddWithValue("@int200_3", SafeVal(int200(2)))
                    cmd.Parameters.AddWithValue("@int250_1", SafeVal(int250(0)))
                    cmd.Parameters.AddWithValue("@int250_2", SafeVal(int250(1)))
                    cmd.Parameters.AddWithValue("@int250_3", SafeVal(int250(2)))
                    cmd.Parameters.AddWithValue("@int300_1", SafeVal(int300(0)))
                    cmd.Parameters.AddWithValue("@int300_2", SafeVal(int300(1)))
                    cmd.Parameters.AddWithValue("@int300_3", SafeVal(int300(2)))
                    cmd.Parameters.AddWithValue("@int350_1", SafeVal(int350(0)))
                    cmd.Parameters.AddWithValue("@int350_2", SafeVal(int350(1)))
                    cmd.Parameters.AddWithValue("@int350_3", SafeVal(int350(2)))
                    cmd.Parameters.AddWithValue("@int400_1", SafeVal(int400(0)))
                    cmd.Parameters.AddWithValue("@int400_2", SafeVal(int400(1)))
                    cmd.Parameters.AddWithValue("@int400_3", SafeVal(int400(2)))
                    cmd.Parameters.AddWithValue("@int450_1", SafeVal(int450(0)))
                    cmd.Parameters.AddWithValue("@int450_2", SafeVal(int450(1)))
                    cmd.Parameters.AddWithValue("@int450_3", SafeVal(int450(2)))
                    cmd.Parameters.AddWithValue("@int500_1", SafeVal(int500(0)))
                    cmd.Parameters.AddWithValue("@int500_2", SafeVal(int500(1)))
                    cmd.Parameters.AddWithValue("@int500_3", SafeVal(int500(2)))
                    cmd.Parameters.AddWithValue("@int550_1", SafeVal(int550(0)))
                    cmd.Parameters.AddWithValue("@int550_2", SafeVal(int550(1)))
                    cmd.Parameters.AddWithValue("@int550_3", SafeVal(int550(2)))
                    cmd.Parameters.AddWithValue("@int600_1", SafeVal(int600(0)))
                    cmd.Parameters.AddWithValue("@int600_2", SafeVal(int600(1)))
                    cmd.Parameters.AddWithValue("@int600_3", SafeVal(int600(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateVernierCalibration600 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from height_gauge_600.
    ''' </summary>
    Public Function GetHeightGauge600MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM height_gauge_600 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetHeightGauge600MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into height_gauge_600.
    ''' </summary>
    Public Function InsertHeightGauge600Calibration(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal obs0 As Object(), ByVal obs20 As Object(), ByVal obs50 As Object(), ByVal obs100 As Object(), ByVal obs150 As Object(), ByVal obs200 As Object(), ByVal obs250 As Object(), ByVal obs300 As Object(),
                                           ByVal obs350 As Object(), ByVal obs400 As Object(), ByVal obs450 As Object(), ByVal obs500 As Object(), ByVal obs550 As Object(), ByVal obs600 As Object(),
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO height_gauge_600 " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs0_0_1, Obs0_0_2, Obs0_0_3, Obs20_0_1, Obs20_0_2, Obs20_0_3, Obs50_0_1, Obs50_0_2, Obs50_0_3, " &
                                   "Obs100_0_1, Obs100_0_2, Obs100_0_3, Obs150_0_1, Obs150_0_2, Obs150_0_3, Obs200_0_1, Obs200_0_2, Obs200_0_3, " &
                                   "Obs250_0_1, Obs250_0_2, Obs250_0_3, Obs300_0_1, Obs300_0_2, Obs300_0_3, " &
                                   "Obs350_0_1, Obs350_0_2, Obs350_0_3, Obs400_0_1, Obs400_0_2, Obs400_0_3, Obs450_0_1, Obs450_0_2, Obs450_0_3, " &
                                   "Obs500_0_1, Obs500_0_2, Obs500_0_3, Obs550_0_1, Obs550_0_2, Obs550_0_3, Obs600_0_1, Obs600_0_2, Obs600_0_3, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs0_1, @obs0_2, @obs0_3, @obs20_1, @obs20_2, @obs20_3, @obs50_1, @obs50_2, @obs50_3, " &
                                   "@obs100_1, @obs100_2, @obs100_3, @obs150_1, @obs150_2, @obs150_3, @obs200_1, @obs200_2, @obs200_3, " &
                                   "@obs250_1, @obs250_2, @obs250_3, @obs300_1, @obs300_2, @obs300_3, " &
                                   "@obs350_1, @obs350_2, @obs350_3, @obs400_1, @obs400_2, @obs400_3, @obs450_1, @obs450_2, @obs450_3, " &
                                   "@obs500_1, @obs500_2, @obs500_3, @obs550_1, @obs550_2, @obs550_3, @obs600_1, @obs600_2, @obs600_3, " &
                                   "@timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))
                    cmd.Parameters.AddWithValue("@obs20_1", SafeVal(obs20(0)))
                    cmd.Parameters.AddWithValue("@obs20_2", SafeVal(obs20(1)))
                    cmd.Parameters.AddWithValue("@obs20_3", SafeVal(obs20(2)))
                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))
                    cmd.Parameters.AddWithValue("@obs100_1", SafeVal(obs100(0)))
                    cmd.Parameters.AddWithValue("@obs100_2", SafeVal(obs100(1)))
                    cmd.Parameters.AddWithValue("@obs100_3", SafeVal(obs100(2)))
                    cmd.Parameters.AddWithValue("@obs150_1", SafeVal(obs150(0)))
                    cmd.Parameters.AddWithValue("@obs150_2", SafeVal(obs150(1)))
                    cmd.Parameters.AddWithValue("@obs150_3", SafeVal(obs150(2)))
                    cmd.Parameters.AddWithValue("@obs200_1", SafeVal(obs200(0)))
                    cmd.Parameters.AddWithValue("@obs200_2", SafeVal(obs200(1)))
                    cmd.Parameters.AddWithValue("@obs200_3", SafeVal(obs200(2)))
                    cmd.Parameters.AddWithValue("@obs250_1", SafeVal(obs250(0)))
                    cmd.Parameters.AddWithValue("@obs250_2", SafeVal(obs250(1)))
                    cmd.Parameters.AddWithValue("@obs250_3", SafeVal(obs250(2)))
                    cmd.Parameters.AddWithValue("@obs300_1", SafeVal(obs300(0)))
                    cmd.Parameters.AddWithValue("@obs300_2", SafeVal(obs300(1)))
                    cmd.Parameters.AddWithValue("@obs300_3", SafeVal(obs300(2)))
                    cmd.Parameters.AddWithValue("@obs350_1", SafeVal(obs350(0)))
                    cmd.Parameters.AddWithValue("@obs350_2", SafeVal(obs350(1)))
                    cmd.Parameters.AddWithValue("@obs350_3", SafeVal(obs350(2)))
                    cmd.Parameters.AddWithValue("@obs400_1", SafeVal(obs400(0)))
                    cmd.Parameters.AddWithValue("@obs400_2", SafeVal(obs400(1)))
                    cmd.Parameters.AddWithValue("@obs400_3", SafeVal(obs400(2)))
                    cmd.Parameters.AddWithValue("@obs450_1", SafeVal(obs450(0)))
                    cmd.Parameters.AddWithValue("@obs450_2", SafeVal(obs450(1)))
                    cmd.Parameters.AddWithValue("@obs450_3", SafeVal(obs450(2)))
                    cmd.Parameters.AddWithValue("@obs500_1", SafeVal(obs500(0)))
                    cmd.Parameters.AddWithValue("@obs500_2", SafeVal(obs500(1)))
                    cmd.Parameters.AddWithValue("@obs500_3", SafeVal(obs500(2)))
                    cmd.Parameters.AddWithValue("@obs550_1", SafeVal(obs550(0)))
                    cmd.Parameters.AddWithValue("@obs550_2", SafeVal(obs550(1)))
                    cmd.Parameters.AddWithValue("@obs550_3", SafeVal(obs550(2)))
                    cmd.Parameters.AddWithValue("@obs600_1", SafeVal(obs600(0)))
                    cmd.Parameters.AddWithValue("@obs600_2", SafeVal(obs600(1)))
                    cmd.Parameters.AddWithValue("@obs600_3", SafeVal(obs600(2)))

                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertHeightGauge600Calibration Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetHeightGauge600CalibrationData(controlNo As String, cycleName As String) As DataTable
        Try
            Dim query As String = "SELECT * FROM height_gauge_600 WHERE ControlNo = @control AND CycleName = @cycle AND RowType = 'RECORD' LIMIT 1"
            Dim dParam As MySqlParameter() = {
                New MySqlParameter("@control", controlNo),
                New MySqlParameter("@cycle", cycleName)
            }
            Return ReadDatatable(query, dParam)
        Catch ex As Exception
            Console.WriteLine("GetHeightGauge600CalibrationData Error: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function UpdateHeightGauge600Calibration(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal obs0 As Object(), ByVal obs20 As Object(), ByVal obs50 As Object(), ByVal obs100 As Object(), ByVal obs150 As Object(), ByVal obs200 As Object(), ByVal obs250 As Object(), ByVal obs300 As Object(),
                                           ByVal obs350 As Object(), ByVal obs400 As Object(), ByVal obs450 As Object(), ByVal obs500 As Object(), ByVal obs550 As Object(), ByVal obs600 As Object(),
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE height_gauge_600 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs0_0_1=@obs0_1, Obs0_0_2=@obs0_2, Obs0_0_3=@obs0_3, Obs20_0_1=@obs20_1, Obs20_0_2=@obs20_2, Obs20_0_3=@obs20_3, Obs50_0_1=@obs50_1, Obs50_0_2=@obs50_2, Obs50_0_3=@obs50_3, " &
                                   "Obs100_0_1=@obs100_1, Obs100_0_2=@obs100_2, Obs100_0_3=@obs100_3, Obs150_0_1=@obs150_1, Obs150_0_2=@obs150_2, Obs150_0_3=@obs150_3, Obs200_0_1=@obs200_1, Obs200_0_2=@obs200_2, Obs200_0_3=@obs200_3, " &
                                   "Obs250_0_1=@obs250_1, Obs250_0_2=@obs250_2, Obs250_0_3=@obs250_3, Obs300_0_1=@obs300_1, Obs300_0_2=@obs300_2, Obs300_0_3=@obs300_3, " &
                                   "Obs350_0_1=@obs350_1, Obs350_0_2=@obs350_2, Obs350_0_3=@obs350_3, Obs400_0_1=@obs400_1, Obs400_0_2=@obs400_2, Obs400_0_3=@obs400_3, Obs450_0_1=@obs450_1, Obs450_0_2=@obs450_2, Obs450_0_3=@obs450_3, " &
                                   "Obs500_0_1=@obs500_1, Obs500_0_2=@obs500_2, Obs500_0_3=@obs500_3, Obs550_0_1=@obs550_1, Obs550_0_2=@obs550_2, Obs550_0_3=@obs550_3, Obs600_0_1=@obs600_1, Obs600_0_2=@obs600_2, Obs600_0_3=@obs600_3, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))
                    cmd.Parameters.AddWithValue("@obs20_1", SafeVal(obs20(0)))
                    cmd.Parameters.AddWithValue("@obs20_2", SafeVal(obs20(1)))
                    cmd.Parameters.AddWithValue("@obs20_3", SafeVal(obs20(2)))
                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))
                    cmd.Parameters.AddWithValue("@obs100_1", SafeVal(obs100(0)))
                    cmd.Parameters.AddWithValue("@obs100_2", SafeVal(obs100(1)))
                    cmd.Parameters.AddWithValue("@obs100_3", SafeVal(obs100(2)))
                    cmd.Parameters.AddWithValue("@obs150_1", SafeVal(obs150(0)))
                    cmd.Parameters.AddWithValue("@obs150_2", SafeVal(obs150(1)))
                    cmd.Parameters.AddWithValue("@obs150_3", SafeVal(obs150(2)))
                    cmd.Parameters.AddWithValue("@obs200_1", SafeVal(obs200(0)))
                    cmd.Parameters.AddWithValue("@obs200_2", SafeVal(obs200(1)))
                    cmd.Parameters.AddWithValue("@obs200_3", SafeVal(obs200(2)))
                    cmd.Parameters.AddWithValue("@obs250_1", SafeVal(obs250(0)))
                    cmd.Parameters.AddWithValue("@obs250_2", SafeVal(obs250(1)))
                    cmd.Parameters.AddWithValue("@obs250_3", SafeVal(obs250(2)))
                    cmd.Parameters.AddWithValue("@obs300_1", SafeVal(obs300(0)))
                    cmd.Parameters.AddWithValue("@obs300_2", SafeVal(obs300(1)))
                    cmd.Parameters.AddWithValue("@obs300_3", SafeVal(obs300(2)))
                    cmd.Parameters.AddWithValue("@obs350_1", SafeVal(obs350(0)))
                    cmd.Parameters.AddWithValue("@obs350_2", SafeVal(obs350(1)))
                    cmd.Parameters.AddWithValue("@obs350_3", SafeVal(obs350(2)))
                    cmd.Parameters.AddWithValue("@obs400_1", SafeVal(obs400(0)))
                    cmd.Parameters.AddWithValue("@obs400_2", SafeVal(obs400(1)))
                    cmd.Parameters.AddWithValue("@obs400_3", SafeVal(obs400(2)))
                    cmd.Parameters.AddWithValue("@obs450_1", SafeVal(obs450(0)))
                    cmd.Parameters.AddWithValue("@obs450_2", SafeVal(obs450(1)))
                    cmd.Parameters.AddWithValue("@obs450_3", SafeVal(obs450(2)))
                    cmd.Parameters.AddWithValue("@obs500_1", SafeVal(obs500(0)))
                    cmd.Parameters.AddWithValue("@obs500_2", SafeVal(obs500(1)))
                    cmd.Parameters.AddWithValue("@obs500_3", SafeVal(obs500(2)))
                    cmd.Parameters.AddWithValue("@obs550_1", SafeVal(obs550(0)))
                    cmd.Parameters.AddWithValue("@obs550_2", SafeVal(obs550(1)))
                    cmd.Parameters.AddWithValue("@obs550_3", SafeVal(obs550(2)))
                    cmd.Parameters.AddWithValue("@obs600_1", SafeVal(obs600(0)))
                    cmd.Parameters.AddWithValue("@obs600_2", SafeVal(obs600(1)))
                    cmd.Parameters.AddWithValue("@obs600_3", SafeVal(obs600(2)))

                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateHeightGauge600Calibration Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Updates the calibration status, setting the calibration flag to YES and calculating the due date.
    ''' </summary>
    Public Function UpdateRegularCalibrationStatus(controlNo As String, calibDate As DateTime, statusOkNg As String, cycleName As String, Optional tableName As String = "") As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim targetTable = If(String.IsNullOrEmpty(tableName), CurrentTargetTrackingTable, tableName)
                Dim dueDate = calibDate.AddYears(1).AddDays(-1)
                Dim sql = $"UPDATE `{targetTable}` SET is_calibrated='YES', calibrated_date=@cDate, due_date=@dDate, calibration_status=@status WHERE control_no=@ctrl AND CycleName=@cycle"
                Using cmd As New MySqlCommand(sql, _con)
                    cmd.Parameters.AddWithValue("@cDate", calibDate.ToString("yyyy-MM-dd"))
                    cmd.Parameters.AddWithValue("@dDate", dueDate.ToString("yyyy-MM-dd"))
                    cmd.Parameters.AddWithValue("@status", statusOkNg)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateRegularCalibrationStatus Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function AddNewAdditionCalibration(instrumentName As String, controlNo As String, status As String, cycleName As String, Optional requestNo As String = "") As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query = "INSERT INTO new_addition_calibration (instrument_name, control_no, status, is_calibrated, CycleName, RequestNo) VALUES (@name, @ctrl, @status, 'NO', @cycle, @reqNo)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@name", instrumentName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.Parameters.AddWithValue("@reqNo", If(String.IsNullOrEmpty(requestNo), DBNull.Value, CObj(requestNo)))
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("AddNewAdditionCalibration Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function AddReintroductionCalibration(instrumentName As String, controlNo As String, status As String, cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query = "INSERT INTO reintroduction_calibration (instrument_name, control_no, status, is_calibrated, CycleName) VALUES (@name, @ctrl, @status, 'NO', @cycle)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@name", instrumentName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("AddReintroductionCalibration Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function AddTempIssuanceCalibration(instrumentName As String, controlNo As String, status As String, cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query = "INSERT INTO temp_issuance_calibration (instrument_name, control_no, status, is_calibrated, CycleName) VALUES (@name, @ctrl, @status, 'NO', @cycle)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@name", instrumentName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("AddTempIssuanceCalibration Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Inserts a record into the result_list when a regular calibration receives an OK status.
    ''' </summary>
    Public Function InsertResultRecord(controlNo As String, category As String, typeName As String, cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sql = "INSERT INTO result_list (control_no, category, type_name, CycleName) VALUES (@c, @cat, @type, @cyc)"
                Using cmd As New MySqlCommand(sql, _con)
                    cmd.Parameters.AddWithValue("@c", controlNo)
                    cmd.Parameters.AddWithValue("@cat", category)
                    cmd.Parameters.AddWithValue("@type", typeName)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertResultRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Inserts a record into the ng_list when a calibration fails with NG status.
    ''' </summary>
    Public Function InsertNGRecord(controlNo As String, instrumentName As String, department As String, cycleName As String, reason As String, calDate As Date, dueDate As Date, calStatus As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sql = "INSERT INTO ng_list (instrument_name, control_no, Department, status, calibrated_date, due_date, calibration_status, reason, CycleName) " &
                          "VALUES (@name, @ctrl, @dept, 'NG', @calDate, @dueDate, @status, @reason, @cycle)"
                Using cmd As New MySqlCommand(sql, _con)
                    cmd.Parameters.AddWithValue("@name", instrumentName)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@dept", department)
                    cmd.Parameters.AddWithValue("@calDate", calDate)
                    cmd.Parameters.AddWithValue("@dueDate", dueDate)
                    cmd.Parameters.AddWithValue("@status", calStatus)
                    cmd.Parameters.AddWithValue("@reason", reason)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertNGRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches detailed calibration data for a specific control number and cycle.
    ''' Resolves the table name using the prefix-to-form mapping in calibrationmapping.
    ''' </summary>
    Public Function GetCalibrationDetail(controlNo As String, cycleName As String) As DataTable
        Try
            ' 1. Extract prefix (e.g. J1-001 -> J1)
            Dim prefix As String = controlNo
            If controlNo.Contains("-") Then
                prefix = controlNo.Split("-"c)(0)
            End If

            ' 2. Map prefix to form_name in calibrationmapping
            Dim mapQuery As String = $"SELECT form_name FROM calibrationmapping WHERE prefix = '{prefix.Replace("'", "''")}'"
            Dim dtMap = ReadDatatable(mapQuery)

            If dtMap.Rows.Count = 0 Then Return New DataTable()

            ' 3. Resolve table name (e.g. VernierCaliper300 -> vernier_caliper_300)
            Dim formName As String = dtMap.Rows(0)("form_name").ToString()
            Dim tableName As String = TypeNameToTableName(formName)

            ' 4. Fetch the record from the specific table
            Dim query As String = $"SELECT * FROM `{tableName}` WHERE ControlNo = @ctrl AND CycleName = @cycle AND RowType = 'RECORD' LIMIT 1"
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetCalibrationDetail Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Wipes all calibration-related transactional data while preserving master/inventory data.
    ''' Truncates status tables and deletes individual test records.
    ''' </summary>
    Public Function WipeCalibrationData() As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                ' 1. Truncate status tracking tables
                ExecuteNonQuery("TRUNCATE TABLE regular_calibration")
                ExecuteNonQuery("TRUNCATE TABLE new_addition_calibration")
                ExecuteNonQuery("TRUNCATE TABLE reintroduction_calibration")
                ExecuteNonQuery("TRUNCATE TABLE temp_issuance_calibration")
                ExecuteNonQuery("TRUNCATE TABLE result_list")
                ExecuteNonQuery("TRUNCATE TABLE ng_list")

                ' 2. Discover all tables with RowType column and delete 'RECORD' entries
                ' This handles tables like vernier_caliper_300 dynamically
                Dim query = "SELECT TABLE_NAME FROM information_schema.COLUMNS WHERE COLUMN_NAME = 'RowType' AND TABLE_SCHEMA = DATABASE()"
                Dim dtTables = ReadDatatable(query)

                For Each row As DataRow In dtTables.Rows
                    Dim tblName = row("TABLE_NAME").ToString()
                    ExecuteNonQuery($"DELETE FROM `{tblName}` WHERE RowType = 'RECORD'")
                Next
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine("WipeCalibrationData Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Wipes calibration-related data for a specific cycle.
    ''' </summary>
    Public Function WipeCalibrationDataByCycle(cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim pCycle As New MySqlParameter("@cycle", cycleName)

                ' 1. Delete from regular_calibration status table
                ExecuteNonQuery("DELETE FROM regular_calibration WHERE CycleName = @cycle", pCycle)
                
                ' 1.1 Delete from other status tables
                ExecuteNonQuery("DELETE FROM new_addition_calibration WHERE CycleName = @cycle", New MySqlParameter("@cycle", cycleName))
                ExecuteNonQuery("DELETE FROM reintroduction_calibration WHERE CycleName = @cycle", New MySqlParameter("@cycle", cycleName))
                ExecuteNonQuery("DELETE FROM temp_issuance_calibration WHERE CycleName = @cycle", New MySqlParameter("@cycle", cycleName))

                ' 2. Delete from result_list report mapping
                ExecuteNonQuery("DELETE FROM result_list WHERE CycleName = @cycle", New MySqlParameter("@cycle", cycleName))

                ' 2.1 Delete from ng_list
                ExecuteNonQuery("DELETE FROM ng_list WHERE CycleName = @cycle", pCycle)

                ' 3. Discover all tables with RowType column and delete 'RECORD' entries for THIS cycle
                Dim query = "SELECT TABLE_NAME FROM information_schema.COLUMNS WHERE COLUMN_NAME = 'CycleName' AND TABLE_SCHEMA = DATABASE()"
                Dim dtTables = ReadDatatable(query)

                For Each row As DataRow In dtTables.Rows
                    Dim tblName = row("TABLE_NAME").ToString()
                    ' Check if it also has RowType before deleting RECORDs
                    Dim dtHasRowType = ReadDatatable($"SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE COLUMN_NAME = 'RowType' AND TABLE_NAME = '{tblName}' AND TABLE_SCHEMA = DATABASE()")

                    If dtHasRowType.Rows.Count > 0 Then
                        ExecuteNonQuery($"DELETE FROM `{tblName}` WHERE CycleName = @cycle AND RowType = 'RECORD'", New MySqlParameter("@cycle", cycleName))
                    End If
                Next
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine("WipeCalibrationDataByCycle Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' --- CALIBRATION MASTER RECORD CRUD ---

    Public Function GetCalibrationMasterData(tableName As String, offset As Integer, limit As Integer, Optional keyword As String = "") As DataTable
        Dim whereClause As String = " WHERE 1=1"
        If Not String.IsNullOrWhiteSpace(keyword) Then
            Dim wClean = keyword.Replace("'", "''")
            whereClause &= $" AND (Description LIKE '%{wClean}%' OR LeastCount LIKE '%{wClean}%' OR Status LIKE '%{wClean}%')"
        End If

        Dim query As String = $"SELECT * FROM `{tableName}`" & whereClause & $" ORDER BY Date DESC, Time DESC LIMIT {limit} OFFSET {offset}"
        Return ReadDatatable(query)
    End Function

    Public Function GetCalibrationMasterTotalCount(tableName As String, Optional keyword As String = "") As Integer
        Dim whereClause As String = " WHERE 1=1"
        If Not String.IsNullOrWhiteSpace(keyword) Then
            Dim wClean = keyword.Replace("'", "''")
            whereClause &= $" AND (Description LIKE '%{wClean}%' OR LeastCount LIKE '%{wClean}%' OR Status LIKE '%{wClean}%')"
        End If

        Dim query As String = $"SELECT COUNT(*) FROM `{tableName}`" & whereClause
        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand(query, _con)
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetCalibrationMasterTotalCount Error: " & ex.Message)
        End Try
        Return 0
    End Function

    Public Function InsertCalibrationMasterRecord(tableName As String, description As String, lc As String, uncertainty As Decimal, calDate As Date?, dueDate As Date?, uploadedDoc As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = $"INSERT INTO `{tableName}` (Date, Time, Description, LeastCount, MasterUncertainty, CalDate, DueDate, uploaded_doc) " &
                                  "VALUES (@date, @time, @desc, @lc, @unc, @cal, @due, @doc)"
                Dim success As Boolean = False
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@date", DateTime.Today)
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@desc", description)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@unc", uncertainty)
                    cmd.Parameters.AddWithValue("@cal", If(calDate.HasValue, calDate.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@due", If(dueDate.HasValue, dueDate.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@doc", uploadedDoc)
                    success = cmd.ExecuteNonQuery() > 0
                End Using

                If success Then
                    ' Sync to calibrationmaster_details
                    Dim syncQuery As String = "INSERT INTO calibrationmaster_details (Date, Time, Description, LeastCount, MasterUncertainty, CalDate, DueDate, uploaded_doc) " &
                                          "VALUES (@date, @time, @desc, @lc, @unc, @cal, @due, @doc)"
                    Using syncCmd As New MySqlCommand(syncQuery, _con)
                        syncCmd.Parameters.AddWithValue("@date", DateTime.Today)
                        syncCmd.Parameters.AddWithValue("@time", DateTime.Now.TimeOfDay)
                        syncCmd.Parameters.AddWithValue("@desc", description)
                        syncCmd.Parameters.AddWithValue("@lc", lc)
                        syncCmd.Parameters.AddWithValue("@unc", uncertainty)
                        syncCmd.Parameters.AddWithValue("@cal", If(calDate.HasValue, calDate.Value, DBNull.Value))
                        syncCmd.Parameters.AddWithValue("@due", If(dueDate.HasValue, dueDate.Value, DBNull.Value))
                        syncCmd.Parameters.AddWithValue("@doc", uploadedDoc)
                        syncCmd.ExecuteNonQuery()
                    End Using
                End If
                Return success
            End If
        Catch ex As Exception
            Console.WriteLine("InsertCalibrationMasterRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateCalibrationMasterRecord(tableName As String, id As Integer, description As String, lc As String, uncertainty As Decimal, calDate As Date?, dueDate As Date?, uploadedDoc As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = $"UPDATE `{tableName}` SET Description=@desc, LeastCount=@lc, MasterUncertainty=@unc, CalDate=@cal, DueDate=@due, uploaded_doc=@doc WHERE ID=@id"
                ' 1. Get current record details to match in sync table (using ID, Description, CalDate as a composite key if ID doesn't match)
                ' Actually, we can just use the Description, CalDate, and Time if we want to be precise, or just ID if calibrationmaster_details had a reference.
                ' Since calibrationmaster_details doesn't have a direct back-reference ID to the dynamic table, 
                ' we'll match by Description and previous CalDate if available, but a better way is to fetch current values first.

                Dim oldCalDate As Object = DBNull.Value
                Dim oldDesc As String = ""
                Dim fetchQuery As String = $"SELECT Description, CalDate FROM `{tableName}` WHERE ID = @id"
                Using fetchCmd As New MySqlCommand(fetchQuery, _con)
                    fetchCmd.Parameters.AddWithValue("@id", id)
                    Dim dt = New DataTable()
                    Using adapter As New MySqlDataAdapter(fetchCmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        oldDesc = dt.Rows(0)("Description").ToString()
                        oldCalDate = dt.Rows(0)("CalDate")
                    End If
                End Using

                Dim success As Boolean = False
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.Parameters.AddWithValue("@desc", description)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@unc", uncertainty)
                    cmd.Parameters.AddWithValue("@cal", If(calDate.HasValue, calDate.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@due", If(dueDate.HasValue, dueDate.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@doc", uploadedDoc)
                    success = cmd.ExecuteNonQuery() > 0
                End Using

                If success AndAlso Not String.IsNullOrEmpty(oldDesc) Then
                    ' Sync Update to calibrationmaster_details
                    ' We match by Description and old CalDate to find the central record
                    Dim syncUpdateQuery As String = "UPDATE calibrationmaster_details SET Description=@desc, LeastCount=@lc, MasterUncertainty=@unc, CalDate=@cal, DueDate=@due, uploaded_doc=@doc " &
                                                 "WHERE Description=@oldDesc AND (CalDate=@oldCal OR (@oldCal IS NULL AND CalDate IS NULL))"
                    Using syncCmd As New MySqlCommand(syncUpdateQuery, _con)
                        syncCmd.Parameters.AddWithValue("@desc", description)
                        syncCmd.Parameters.AddWithValue("@lc", lc)
                        syncCmd.Parameters.AddWithValue("@unc", uncertainty)
                        syncCmd.Parameters.AddWithValue("@cal", If(calDate.HasValue, calDate.Value, DBNull.Value))
                        syncCmd.Parameters.AddWithValue("@due", If(dueDate.HasValue, dueDate.Value, DBNull.Value))
                        syncCmd.Parameters.AddWithValue("@doc", uploadedDoc)
                        syncCmd.Parameters.AddWithValue("@oldDesc", oldDesc)
                        syncCmd.Parameters.AddWithValue("@oldCal", If(oldCalDate Is DBNull.Value, DBNull.Value, oldCalDate))
                        syncCmd.ExecuteNonQuery()
                    End Using
                End If
                Return success
            End If
        Catch ex As Exception
            Console.WriteLine("UpdateCalibrationMasterRecord Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Shared Function CopyCalibrationMasterFile(sourcePath As String, masterName As String) As String
        Try
            If String.IsNullOrEmpty(sourcePath) OrElse Not File.Exists(sourcePath) Then Return ""

            Dim appPath = AppDomain.CurrentDomain.BaseDirectory
            ' Documents \ Calibration Master \ [Master Name]
            Dim targetFolder = Path.Combine(appPath, "Documents", "Calibration Master", masterName)

            If Not Directory.Exists(targetFolder) Then
                Directory.CreateDirectory(targetFolder)
            End If

            Dim originalFileName = Path.GetFileName(sourcePath)
            Dim uniqueName = $"{DateTime.Now:yyyyMMddHHmmss}_{originalFileName}"
            Dim destinationPath = Path.Combine(targetFolder, uniqueName)

            File.Copy(sourcePath, destinationPath, True)

            Return Path.Combine("Documents", "Calibration Master", masterName, uniqueName)
        Catch ex As Exception
            Console.WriteLine("CopyCalibrationMasterFile Error: " & ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from thread_pitch_micrometer_25.
    ''' </summary>
    Public Function GetThreadPitchMicrometer25MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_25 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer25MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into thread_pitch_micrometer_25.
    ''' </summary>
    Public Function InsertThreadPitchMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs2_5 As Object(), ByVal obs5_1 As Object(), ByVal obs7_7 As Object(), ByVal obs10_3 As Object(), ByVal obs12_9 As Object(), ByVal obs15_0 As Object(), ByVal obs17_6 As Object(), ByVal obs20_2 As Object(), ByVal obs22_8 As Object(), ByVal obs25_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_pitch_micrometer_25 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs2_5_1, Obs2_5_2, Obs2_5_3, Obs5_1_1, Obs5_1_2, Obs5_1_3, Obs7_7_1, Obs7_7_2, Obs7_7_3, " &
                                   "Obs10_3_1, Obs10_3_2, Obs10_3_3, Obs12_9_1, Obs12_9_2, Obs12_9_3, Obs15_0_1, Obs15_0_2, Obs15_0_3, " &
                                   "Obs17_6_1, Obs17_6_2, Obs17_6_3, Obs20_2_1, Obs20_2_2, Obs20_2_3, Obs22_8_1, Obs22_8_2, Obs22_8_3, " &
                                   "Obs25_0_1, Obs25_0_2, Obs25_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs2_5_1, @obs2_5_2, @obs2_5_3, @obs5_1_1, @obs5_1_2, @obs5_1_3, @obs7_7_1, @obs7_7_2, @obs7_7_3, " &
                                   "@obs10_3_1, @obs10_3_2, @obs10_3_3, @obs12_9_1, @obs12_9_2, @obs12_9_3, @obs15_0_1, @obs15_0_2, @obs15_0_3, " &
                                   "@obs17_6_1, @obs17_6_2, @obs17_6_3, @obs20_2_1, @obs20_2_2, @obs20_2_3, @obs22_8_1, @obs22_8_2, @obs22_8_3, " &
                                   "@obs25_0_1, @obs25_0_2, @obs25_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs2_5_1", SafeVal(obs2_5(0)))
                    cmd.Parameters.AddWithValue("@obs2_5_2", SafeVal(obs2_5(1)))
                    cmd.Parameters.AddWithValue("@obs2_5_3", SafeVal(obs2_5(2)))

                    cmd.Parameters.AddWithValue("@obs5_1_1", SafeVal(obs5_1(0)))
                    cmd.Parameters.AddWithValue("@obs5_1_2", SafeVal(obs5_1(1)))
                    cmd.Parameters.AddWithValue("@obs5_1_3", SafeVal(obs5_1(2)))

                    cmd.Parameters.AddWithValue("@obs7_7_1", SafeVal(obs7_7(0)))
                    cmd.Parameters.AddWithValue("@obs7_7_2", SafeVal(obs7_7(1)))
                    cmd.Parameters.AddWithValue("@obs7_7_3", SafeVal(obs7_7(2)))

                    cmd.Parameters.AddWithValue("@obs10_3_1", SafeVal(obs10_3(0)))
                    cmd.Parameters.AddWithValue("@obs10_3_2", SafeVal(obs10_3(1)))
                    cmd.Parameters.AddWithValue("@obs10_3_3", SafeVal(obs10_3(2)))

                    cmd.Parameters.AddWithValue("@obs12_9_1", SafeVal(obs12_9(0)))
                    cmd.Parameters.AddWithValue("@obs12_9_2", SafeVal(obs12_9(1)))
                    cmd.Parameters.AddWithValue("@obs12_9_3", SafeVal(obs12_9(2)))

                    cmd.Parameters.AddWithValue("@obs15_0_1", SafeVal(obs15_0(0)))
                    cmd.Parameters.AddWithValue("@obs15_0_2", SafeVal(obs15_0(1)))
                    cmd.Parameters.AddWithValue("@obs15_0_3", SafeVal(obs15_0(2)))

                    cmd.Parameters.AddWithValue("@obs17_6_1", SafeVal(obs17_6(0)))
                    cmd.Parameters.AddWithValue("@obs17_6_2", SafeVal(obs17_6(1)))
                    cmd.Parameters.AddWithValue("@obs17_6_3", SafeVal(obs17_6(2)))

                    cmd.Parameters.AddWithValue("@obs20_2_1", SafeVal(obs20_2(0)))
                    cmd.Parameters.AddWithValue("@obs20_2_2", SafeVal(obs20_2(1)))
                    cmd.Parameters.AddWithValue("@obs20_2_3", SafeVal(obs20_2(2)))

                    cmd.Parameters.AddWithValue("@obs22_8_1", SafeVal(obs22_8(0)))
                    cmd.Parameters.AddWithValue("@obs22_8_2", SafeVal(obs22_8(1)))
                    cmd.Parameters.AddWithValue("@obs22_8_3", SafeVal(obs22_8(2)))

                    cmd.Parameters.AddWithValue("@obs25_0_1", SafeVal(obs25_0(0)))
                    cmd.Parameters.AddWithValue("@obs25_0_2", SafeVal(obs25_0(1)))
                    cmd.Parameters.AddWithValue("@obs25_0_3", SafeVal(obs25_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertThreadPitchMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from thread_pitch_micrometer_25.
    ''' </summary>
    Public Function GetThreadPitchMicrometer25Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_25 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer25Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in thread_pitch_micrometer_25.
    ''' </summary>
    Public Function UpdateThreadPitchMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs2_5 As Object(), ByVal obs5_1 As Object(), ByVal obs7_7 As Object(), ByVal obs10_3 As Object(), ByVal obs12_9 As Object(), ByVal obs15_0 As Object(), ByVal obs17_6 As Object(), ByVal obs20_2 As Object(), ByVal obs22_8 As Object(), ByVal obs25_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE thread_pitch_micrometer_25 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs2_5_1=@obs2_5_1, Obs2_5_2=@obs2_5_2, Obs2_5_3=@obs2_5_3, Obs5_1_1=@obs5_1_1, Obs5_1_2=@obs5_1_2, Obs5_1_3=@obs5_1_3, Obs7_7_1=@obs7_7_1, Obs7_7_2=@obs7_7_2, Obs7_7_3=@obs7_7_3, " &
                                   "Obs10_3_1=@obs10_3_1, Obs10_3_2=@obs10_3_2, Obs10_3_3=@obs10_3_3, Obs12_9_1=@obs12_9_1, Obs12_9_2=@obs12_9_2, Obs12_9_3=@obs12_9_3, Obs15_0_1=@obs15_0_1, Obs15_0_2=@obs15_0_2, Obs15_0_3=@obs15_0_3, " &
                                   "Obs17_6_1=@obs17_6_1, Obs17_6_2=@obs17_6_2, Obs17_6_3=@obs17_6_3, Obs20_2_1=@obs20_2_1, Obs20_2_2=@obs20_2_2, Obs20_2_3=@obs20_2_3, Obs22_8_1=@obs22_8_1, Obs22_8_2=@obs22_8_2, Obs22_8_3=@obs22_8_3, " &
                                   "Obs25_0_1=@obs25_0_1, Obs25_0_2=@obs25_0_2, Obs25_0_3=@obs25_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs2_5_1", SafeVal(obs2_5(0)))
                    cmd.Parameters.AddWithValue("@obs2_5_2", SafeVal(obs2_5(1)))
                    cmd.Parameters.AddWithValue("@obs2_5_3", SafeVal(obs2_5(2)))

                    cmd.Parameters.AddWithValue("@obs5_1_1", SafeVal(obs5_1(0)))
                    cmd.Parameters.AddWithValue("@obs5_1_2", SafeVal(obs5_1(1)))
                    cmd.Parameters.AddWithValue("@obs5_1_3", SafeVal(obs5_1(2)))

                    cmd.Parameters.AddWithValue("@obs7_7_1", SafeVal(obs7_7(0)))
                    cmd.Parameters.AddWithValue("@obs7_7_2", SafeVal(obs7_7(1)))
                    cmd.Parameters.AddWithValue("@obs7_7_3", SafeVal(obs7_7(2)))

                    cmd.Parameters.AddWithValue("@obs10_3_1", SafeVal(obs10_3(0)))
                    cmd.Parameters.AddWithValue("@obs10_3_2", SafeVal(obs10_3(1)))
                    cmd.Parameters.AddWithValue("@obs10_3_3", SafeVal(obs10_3(2)))

                    cmd.Parameters.AddWithValue("@obs12_9_1", SafeVal(obs12_9(0)))
                    cmd.Parameters.AddWithValue("@obs12_9_2", SafeVal(obs12_9(1)))
                    cmd.Parameters.AddWithValue("@obs12_9_3", SafeVal(obs12_9(2)))

                    cmd.Parameters.AddWithValue("@obs15_0_1", SafeVal(obs15_0(0)))
                    cmd.Parameters.AddWithValue("@obs15_0_2", SafeVal(obs15_0(1)))
                    cmd.Parameters.AddWithValue("@obs15_0_3", SafeVal(obs15_0(2)))

                    cmd.Parameters.AddWithValue("@obs17_6_1", SafeVal(obs17_6(0)))
                    cmd.Parameters.AddWithValue("@obs17_6_2", SafeVal(obs17_6(1)))
                    cmd.Parameters.AddWithValue("@obs17_6_3", SafeVal(obs17_6(2)))

                    cmd.Parameters.AddWithValue("@obs20_2_1", SafeVal(obs20_2(0)))
                    cmd.Parameters.AddWithValue("@obs20_2_2", SafeVal(obs20_2(1)))
                    cmd.Parameters.AddWithValue("@obs20_2_3", SafeVal(obs20_2(2)))

                    cmd.Parameters.AddWithValue("@obs22_8_1", SafeVal(obs22_8(0)))
                    cmd.Parameters.AddWithValue("@obs22_8_2", SafeVal(obs22_8(1)))
                    cmd.Parameters.AddWithValue("@obs22_8_3", SafeVal(obs22_8(2)))

                    cmd.Parameters.AddWithValue("@obs25_0_1", SafeVal(obs25_0(0)))
                    cmd.Parameters.AddWithValue("@obs25_0_2", SafeVal(obs25_0(1)))
                    cmd.Parameters.AddWithValue("@obs25_0_3", SafeVal(obs25_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateThreadPitchMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from groove_micrometer_25.
    ''' </summary>
    Public Function GetGrooveMicrometer25MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM groove_micrometer_25 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetGrooveMicrometer25MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into groove_micrometer_25.
    ''' </summary>
    Public Function InsertGrooveMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal ext0 As Object(), ByVal ext5 As Object(), ByVal ext10 As Object(), ByVal ext15 As Object(), ByVal ext20 As Object(), ByVal ext25 As Object(),
                                          ByVal int5 As Object(), ByVal int10 As Object(), ByVal int15 As Object(), ByVal int20 As Object(), ByVal int25 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO groove_micrometer_25 " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Ext0_1, Ext0_2, Ext0_3, Ext5_1, Ext5_2, Ext5_3, Ext10_1, Ext10_2, Ext10_3, " &
                                   "Ext15_1, Ext15_2, Ext15_3, Ext20_1, Ext20_2, Ext20_3, Ext25_1, Ext25_2, Ext25_3, " &
                                   "Int5_1, Int5_2, Int5_3, Int10_1, Int10_2, Int10_3, Int15_1, Int15_2, Int15_3, Int20_1, Int20_2, Int20_3, Int25_1, Int25_2, Int25_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@ext0_1, @ext0_2, @ext0_3, @ext5_1, @ext5_2, @ext5_3, @ext10_1, @ext10_2, @ext10_3, " &
                                   "@ext15_1, @ext15_2, @ext15_3, @ext20_1, @ext20_2, @ext20_3, @ext25_1, @ext25_2, @ext25_3, " &
                                   "@int5_1, @int5_2, @int5_3, @int10_1, @int10_2, @int10_3, @int15_1, @int15_2, @int15_3, @int20_1, @int20_2, @int20_3, @int25_1, @int25_2, @int25_3, " &
                                   "@depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@ext0_1", SafeVal(ext0(0)))
                    cmd.Parameters.AddWithValue("@ext0_2", SafeVal(ext0(1)))
                    cmd.Parameters.AddWithValue("@ext0_3", SafeVal(ext0(2)))

                    cmd.Parameters.AddWithValue("@ext5_1", SafeVal(ext5(0)))
                    cmd.Parameters.AddWithValue("@ext5_2", SafeVal(ext5(1)))
                    cmd.Parameters.AddWithValue("@ext5_3", SafeVal(ext5(2)))

                    cmd.Parameters.AddWithValue("@ext10_1", SafeVal(ext10(0)))
                    cmd.Parameters.AddWithValue("@ext10_2", SafeVal(ext10(1)))
                    cmd.Parameters.AddWithValue("@ext10_3", SafeVal(ext10(2)))

                    cmd.Parameters.AddWithValue("@ext15_1", SafeVal(ext15(0)))
                    cmd.Parameters.AddWithValue("@ext15_2", SafeVal(ext15(1)))
                    cmd.Parameters.AddWithValue("@ext15_3", SafeVal(ext15(2)))

                    cmd.Parameters.AddWithValue("@ext20_1", SafeVal(ext20(0)))
                    cmd.Parameters.AddWithValue("@ext20_2", SafeVal(ext20(1)))
                    cmd.Parameters.AddWithValue("@ext20_3", SafeVal(ext20(2)))

                    cmd.Parameters.AddWithValue("@ext25_1", SafeVal(ext25(0)))
                    cmd.Parameters.AddWithValue("@ext25_2", SafeVal(ext25(1)))
                    cmd.Parameters.AddWithValue("@ext25_3", SafeVal(ext25(2)))

                    cmd.Parameters.AddWithValue("@int5_1", SafeVal(int5(0)))
                    cmd.Parameters.AddWithValue("@int5_2", SafeVal(int5(1)))
                    cmd.Parameters.AddWithValue("@int5_3", SafeVal(int5(2)))

                    cmd.Parameters.AddWithValue("@int10_1", SafeVal(int10(0)))
                    cmd.Parameters.AddWithValue("@int10_2", SafeVal(int10(1)))
                    cmd.Parameters.AddWithValue("@int10_3", SafeVal(int10(2)))

                    cmd.Parameters.AddWithValue("@int15_1", SafeVal(int15(0)))
                    cmd.Parameters.AddWithValue("@int15_2", SafeVal(int15(1)))
                    cmd.Parameters.AddWithValue("@int15_3", SafeVal(int15(2)))

                    cmd.Parameters.AddWithValue("@int20_1", SafeVal(int20(0)))
                    cmd.Parameters.AddWithValue("@int20_2", SafeVal(int20(1)))
                    cmd.Parameters.AddWithValue("@int20_3", SafeVal(int20(2)))

                    cmd.Parameters.AddWithValue("@int25_1", SafeVal(int25(0)))
                    cmd.Parameters.AddWithValue("@int25_2", SafeVal(int25(1)))
                    cmd.Parameters.AddWithValue("@int25_3", SafeVal(int25(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertGrooveMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetGrooveMicrometer25Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM groove_micrometer_25 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetGrooveMicrometer25Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function UpdateGrooveMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal ext0 As Object(), ByVal ext5 As Object(), ByVal ext10 As Object(), ByVal ext15 As Object(), ByVal ext20 As Object(), ByVal ext25 As Object(),
                                           ByVal int5 As Object(), ByVal int10 As Object(), ByVal int15 As Object(), ByVal int20 As Object(), ByVal int25 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE groove_micrometer_25 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Ext0_1=@ext0_1, Ext0_2=@ext0_2, Ext0_3=@ext0_3, Ext5_1=@ext5_1, Ext5_2=@ext5_2, Ext5_3=@ext5_3, Ext10_1=@ext10_1, Ext10_2=@ext10_2, Ext10_3=@ext10_3, " &
                                   "Ext15_1=@ext15_1, Ext15_2=@ext15_2, Ext15_3=@ext15_3, Ext20_1=@ext20_1, Ext20_2=@ext20_2, Ext20_3=@ext20_3, Ext25_1=@ext25_1, Ext25_2=@ext25_2, Ext25_3=@ext25_3, " &
                                   "Int5_1=@int5_1, Int5_2=@int5_2, Int5_3=@int5_3, Int10_1=@int10_1, Int10_2=@int10_2, Int10_3=@int10_3, Int15_1=@int15_1, Int15_2=@int15_2, Int15_3=@int15_3, Int20_1=@int20_1, Int20_2=@int20_2, Int20_3=@int20_3, Int25_1=@int25_1, Int25_2=@int25_2, Int25_3=@int25_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@ext0_1", SafeVal(ext0(0)))
                    cmd.Parameters.AddWithValue("@ext0_2", SafeVal(ext0(1)))
                    cmd.Parameters.AddWithValue("@ext0_3", SafeVal(ext0(2)))
                    cmd.Parameters.AddWithValue("@ext5_1", SafeVal(ext5(0)))
                    cmd.Parameters.AddWithValue("@ext5_2", SafeVal(ext5(1)))
                    cmd.Parameters.AddWithValue("@ext5_3", SafeVal(ext5(2)))
                    cmd.Parameters.AddWithValue("@ext10_1", SafeVal(ext10(0)))
                    cmd.Parameters.AddWithValue("@ext10_2", SafeVal(ext10(1)))
                    cmd.Parameters.AddWithValue("@ext10_3", SafeVal(ext10(2)))
                    cmd.Parameters.AddWithValue("@ext15_1", SafeVal(ext15(0)))
                    cmd.Parameters.AddWithValue("@ext15_2", SafeVal(ext15(1)))
                    cmd.Parameters.AddWithValue("@ext15_3", SafeVal(ext15(2)))
                    cmd.Parameters.AddWithValue("@ext20_1", SafeVal(ext20(0)))
                    cmd.Parameters.AddWithValue("@ext20_2", SafeVal(ext20(1)))
                    cmd.Parameters.AddWithValue("@ext20_3", SafeVal(ext20(2)))
                    cmd.Parameters.AddWithValue("@ext25_1", SafeVal(ext25(0)))
                    cmd.Parameters.AddWithValue("@ext25_2", SafeVal(ext25(1)))
                    cmd.Parameters.AddWithValue("@ext25_3", SafeVal(ext25(2)))

                    cmd.Parameters.AddWithValue("@int5_1", SafeVal(int5(0)))
                    cmd.Parameters.AddWithValue("@int5_2", SafeVal(int5(1)))
                    cmd.Parameters.AddWithValue("@int5_3", SafeVal(int5(2)))
                    cmd.Parameters.AddWithValue("@int10_1", SafeVal(int10(0)))
                    cmd.Parameters.AddWithValue("@int10_2", SafeVal(int10(1)))
                    cmd.Parameters.AddWithValue("@int10_3", SafeVal(int10(2)))
                    cmd.Parameters.AddWithValue("@int15_1", SafeVal(int15(0)))
                    cmd.Parameters.AddWithValue("@int15_2", SafeVal(int15(1)))
                    cmd.Parameters.AddWithValue("@int15_3", SafeVal(int15(2)))
                    cmd.Parameters.AddWithValue("@int20_1", SafeVal(int20(0)))
                    cmd.Parameters.AddWithValue("@int20_2", SafeVal(int20(1)))
                    cmd.Parameters.AddWithValue("@int20_3", SafeVal(int20(2)))
                    cmd.Parameters.AddWithValue("@int25_1", SafeVal(int25(0)))
                    cmd.Parameters.AddWithValue("@int25_2", SafeVal(int25(1)))
                    cmd.Parameters.AddWithValue("@int25_3", SafeVal(int25(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateGrooveMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from groove_micrometer_50.
    ''' </summary>
    Public Function GetGrooveMicrometer50MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM groove_micrometer_50 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetGrooveMicrometer50MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into groove_micrometer_50.
    ''' </summary>
    Public Function InsertGrooveMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal ext25 As Object(), ByVal ext30 As Object(), ByVal ext35 As Object(), ByVal ext40 As Object(), ByVal ext45 As Object(), ByVal ext50 As Object(),
                                          ByVal int25 As Object(), ByVal int30 As Object(), ByVal int35 As Object(), ByVal int40 As Object(), ByVal int45 As Object(), ByVal int50 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO groove_micrometer_50 " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Ext25_1, Ext25_2, Ext25_3, Ext30_1, Ext30_2, Ext30_3, Ext35_1, Ext35_2, Ext35_3, " &
                                   "Ext40_1, Ext40_2, Ext40_3, Ext45_1, Ext45_2, Ext45_3, Ext50_1, Ext50_2, Ext50_3, " &
                                   "Int25_1, Int25_2, Int25_3, Int30_1, Int30_2, Int30_3, Int35_1, Int35_2, Int35_3, Int40_1, Int40_2, Int40_3, Int45_1, Int45_2, Int45_3, Int50_1, Int50_2, Int50_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@ext25_1, @ext25_2, @ext25_3, @ext30_1, @ext30_2, @ext30_3, @ext35_1, @ext35_2, @ext35_3, " &
                                   "@ext40_1, @ext40_2, @ext40_3, @ext45_1, @ext45_2, @ext45_3, @ext50_1, @ext50_2, @ext50_3, " &
                                   "@int25_1, @int25_2, @int25_3, @int30_1, @int30_2, @int30_3, @int35_1, @int35_2, @int35_3, @int40_1, @int40_2, @int40_3, @int45_1, @int45_2, @int45_3, @int50_1, @int50_2, @int50_3, " &
                                   "@depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@ext25_1", SafeVal(ext25(0)))
                    cmd.Parameters.AddWithValue("@ext25_2", SafeVal(ext25(1)))
                    cmd.Parameters.AddWithValue("@ext25_3", SafeVal(ext25(2)))
                    cmd.Parameters.AddWithValue("@ext30_1", SafeVal(ext30(0)))
                    cmd.Parameters.AddWithValue("@ext30_2", SafeVal(ext30(1)))
                    cmd.Parameters.AddWithValue("@ext30_3", SafeVal(ext30(2)))
                    cmd.Parameters.AddWithValue("@ext35_1", SafeVal(ext35(0)))
                    cmd.Parameters.AddWithValue("@ext35_2", SafeVal(ext35(1)))
                    cmd.Parameters.AddWithValue("@ext35_3", SafeVal(ext35(2)))
                    cmd.Parameters.AddWithValue("@ext40_1", SafeVal(ext40(0)))
                    cmd.Parameters.AddWithValue("@ext40_2", SafeVal(ext40(1)))
                    cmd.Parameters.AddWithValue("@ext40_3", SafeVal(ext40(2)))
                    cmd.Parameters.AddWithValue("@ext45_1", SafeVal(ext45(0)))
                    cmd.Parameters.AddWithValue("@ext45_2", SafeVal(ext45(1)))
                    cmd.Parameters.AddWithValue("@ext45_3", SafeVal(ext45(2)))
                    cmd.Parameters.AddWithValue("@ext50_1", SafeVal(ext50(0)))
                    cmd.Parameters.AddWithValue("@ext50_2", SafeVal(ext50(1)))
                    cmd.Parameters.AddWithValue("@ext50_3", SafeVal(ext50(2)))

                    cmd.Parameters.AddWithValue("@int25_1", SafeVal(int25(0)))
                    cmd.Parameters.AddWithValue("@int25_2", SafeVal(int25(1)))
                    cmd.Parameters.AddWithValue("@int25_3", SafeVal(int25(2)))
                    cmd.Parameters.AddWithValue("@int30_1", SafeVal(int30(0)))
                    cmd.Parameters.AddWithValue("@int30_2", SafeVal(int30(1)))
                    cmd.Parameters.AddWithValue("@int30_3", SafeVal(int30(2)))
                    cmd.Parameters.AddWithValue("@int35_1", SafeVal(int35(0)))
                    cmd.Parameters.AddWithValue("@int35_2", SafeVal(int35(1)))
                    cmd.Parameters.AddWithValue("@int35_3", SafeVal(int35(2)))
                    cmd.Parameters.AddWithValue("@int40_1", SafeVal(int40(0)))
                    cmd.Parameters.AddWithValue("@int40_2", SafeVal(int40(1)))
                    cmd.Parameters.AddWithValue("@int40_3", SafeVal(int40(2)))
                    cmd.Parameters.AddWithValue("@int45_1", SafeVal(int45(0)))
                    cmd.Parameters.AddWithValue("@int45_2", SafeVal(int45(1)))
                    cmd.Parameters.AddWithValue("@int45_3", SafeVal(int45(2)))
                    cmd.Parameters.AddWithValue("@int50_1", SafeVal(int50(0)))
                    cmd.Parameters.AddWithValue("@int50_2", SafeVal(int50(1)))
                    cmd.Parameters.AddWithValue("@int50_3", SafeVal(int50(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertGrooveMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetGrooveMicrometer50Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM groove_micrometer_50 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetGrooveMicrometer50Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function UpdateGrooveMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal ext25 As Object(), ByVal ext30 As Object(), ByVal ext35 As Object(), ByVal ext40 As Object(), ByVal ext45 As Object(), ByVal ext50 As Object(),
                                           ByVal int25 As Object(), ByVal int30 As Object(), ByVal int35 As Object(), ByVal int40 As Object(), ByVal int45 As Object(), ByVal int50 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE groove_micrometer_50 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Ext25_1=@ext25_1, Ext25_2=@ext25_2, Ext25_3=@ext25_3, Ext30_1=@ext30_1, Ext30_2=@ext30_2, Ext30_3=@ext30_3, Ext35_1=@ext35_1, Ext35_2=@ext35_2, Ext35_3=@ext35_3, " &
                                   "Ext40_1=@ext40_1, Ext40_2=@ext40_2, Ext40_3=@ext40_3, Ext45_1=@ext45_1, Ext45_2=@ext45_2, Ext45_3=@ext45_3, Ext50_1=@ext50_1, Ext50_2=@ext50_2, Ext50_3=@ext50_3, " &
                                   "Int25_1=@int25_1, Int25_2=@int25_2, Int25_3=@int25_3, Int30_1=@int30_1, Int30_2=@int30_2, Int30_3=@int30_3, Int35_1=@int35_1, Int35_2=@int35_2, Int35_3=@int35_3, Int40_1=@int40_1, Int40_2=@int40_2, Int40_3=@int40_3, Int45_1=@int45_1, Int45_2=@int45_2, Int45_3=@int45_3, Int50_1=@int50_1, Int50_2=@int50_2, Int50_3=@int50_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@ext25_1", SafeVal(ext25(0)))
                    cmd.Parameters.AddWithValue("@ext25_2", SafeVal(ext25(1)))
                    cmd.Parameters.AddWithValue("@ext25_3", SafeVal(ext25(2)))
                    cmd.Parameters.AddWithValue("@ext30_1", SafeVal(ext30(0)))
                    cmd.Parameters.AddWithValue("@ext30_2", SafeVal(ext30(1)))
                    cmd.Parameters.AddWithValue("@ext30_3", SafeVal(ext30(2)))
                    cmd.Parameters.AddWithValue("@ext35_1", SafeVal(ext35(0)))
                    cmd.Parameters.AddWithValue("@ext35_2", SafeVal(ext35(1)))
                    cmd.Parameters.AddWithValue("@ext35_3", SafeVal(ext35(2)))
                    cmd.Parameters.AddWithValue("@ext40_1", SafeVal(ext40(0)))
                    cmd.Parameters.AddWithValue("@ext40_2", SafeVal(ext40(1)))
                    cmd.Parameters.AddWithValue("@ext40_3", SafeVal(ext40(2)))
                    cmd.Parameters.AddWithValue("@ext45_1", SafeVal(ext45(0)))
                    cmd.Parameters.AddWithValue("@ext45_2", SafeVal(ext45(1)))
                    cmd.Parameters.AddWithValue("@ext45_3", SafeVal(ext45(2)))
                    cmd.Parameters.AddWithValue("@ext50_1", SafeVal(ext50(0)))
                    cmd.Parameters.AddWithValue("@ext50_2", SafeVal(ext50(1)))
                    cmd.Parameters.AddWithValue("@ext50_3", SafeVal(ext50(2)))

                    cmd.Parameters.AddWithValue("@int25_1", SafeVal(int25(0)))
                    cmd.Parameters.AddWithValue("@int25_2", SafeVal(int25(1)))
                    cmd.Parameters.AddWithValue("@int25_3", SafeVal(int25(2)))
                    cmd.Parameters.AddWithValue("@int30_1", SafeVal(int30(0)))
                    cmd.Parameters.AddWithValue("@int30_2", SafeVal(int30(1)))
                    cmd.Parameters.AddWithValue("@int30_3", SafeVal(int30(2)))
                    cmd.Parameters.AddWithValue("@int35_1", SafeVal(int35(0)))
                    cmd.Parameters.AddWithValue("@int35_2", SafeVal(int35(1)))
                    cmd.Parameters.AddWithValue("@int35_3", SafeVal(int35(2)))
                    cmd.Parameters.AddWithValue("@int40_1", SafeVal(int40(0)))
                    cmd.Parameters.AddWithValue("@int40_2", SafeVal(int40(1)))
                    cmd.Parameters.AddWithValue("@int40_3", SafeVal(int40(2)))
                    cmd.Parameters.AddWithValue("@int45_1", SafeVal(int45(0)))
                    cmd.Parameters.AddWithValue("@int45_2", SafeVal(int45(1)))
                    cmd.Parameters.AddWithValue("@int45_3", SafeVal(int45(2)))
                    cmd.Parameters.AddWithValue("@int50_1", SafeVal(int50(0)))
                    cmd.Parameters.AddWithValue("@int50_2", SafeVal(int50(1)))
                    cmd.Parameters.AddWithValue("@int50_3", SafeVal(int50(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateGrooveMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from geartooth_micrometer_50.
    ''' </summary>
    Public Function GetGeartoothMicrometer50MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM geartooth_micrometer_50 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetGeartoothMicrometer50MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into geartooth_micrometer_50.
    ''' </summary>
    Public Function InsertGeartoothMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal obs0 As Object(), ByVal obs5 As Object(), ByVal obs10 As Object(), ByVal obs15 As Object(), ByVal obs20 As Object(), ByVal obs25 As Object(),
                                           ByVal obs30 As Object(), ByVal obs35 As Object(), ByVal obs40 As Object(), ByVal obs45 As Object(), ByVal obs50 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO geartooth_micrometer_50 " &
                                   "(RowType, `Date`, `Time`, `Type`, ControlNo, CycleName, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs0_0_1, Obs0_0_2, Obs0_0_3, Obs5_0_1, Obs5_0_2, Obs5_0_3, Obs10_0_1, Obs10_0_2, Obs10_0_3, " &
                                   "Obs15_0_1, Obs15_0_2, Obs15_0_3, Obs20_0_1, Obs20_0_2, Obs20_0_3, Obs25_0_1, Obs25_0_2, Obs25_0_3, " &
                                   "Obs30_0_1, Obs30_0_2, Obs30_0_3, Obs35_0_1, Obs35_0_2, Obs35_0_3, Obs40_0_1, Obs40_0_2, Obs40_0_3, " &
                                   "Obs45_0_1, Obs45_0_2, Obs45_0_3, Obs50_0_1, Obs50_0_2, Obs50_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @date, @time, @type, @controlNo, @cycleName, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs0_1, @obs0_2, @obs0_3, @obs5_1, @obs5_2, @obs5_3, @obs10_1, @obs10_2, @obs10_3, " &
                                   "@obs15_1, @obs15_2, @obs15_3, @obs20_1, @obs20_2, @obs20_3, @obs25_1, @obs25_2, @obs25_3, " &
                                   "@obs30_1, @obs30_2, @obs30_3, @obs35_1, @obs35_2, @obs35_3, @obs40_1, @obs40_2, @obs40_3, " &
                                   "@obs45_1, @obs45_2, @obs45_3, @obs50_1, @obs50_2, @obs50_3, " &
                                   "@depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))
                    cmd.Parameters.AddWithValue("@obs5_1", SafeVal(obs5(0)))
                    cmd.Parameters.AddWithValue("@obs5_2", SafeVal(obs5(1)))
                    cmd.Parameters.AddWithValue("@obs5_3", SafeVal(obs5(2)))
                    cmd.Parameters.AddWithValue("@obs10_1", SafeVal(obs10(0)))
                    cmd.Parameters.AddWithValue("@obs10_2", SafeVal(obs10(1)))
                    cmd.Parameters.AddWithValue("@obs10_3", SafeVal(obs10(2)))
                    cmd.Parameters.AddWithValue("@obs15_1", SafeVal(obs15(0)))
                    cmd.Parameters.AddWithValue("@obs15_2", SafeVal(obs15(1)))
                    cmd.Parameters.AddWithValue("@obs15_3", SafeVal(obs15(2)))
                    cmd.Parameters.AddWithValue("@obs20_1", SafeVal(obs20(0)))
                    cmd.Parameters.AddWithValue("@obs20_2", SafeVal(obs20(1)))
                    cmd.Parameters.AddWithValue("@obs20_3", SafeVal(obs20(2)))
                    cmd.Parameters.AddWithValue("@obs25_1", SafeVal(obs25(0)))
                    cmd.Parameters.AddWithValue("@obs25_2", SafeVal(obs25(1)))
                    cmd.Parameters.AddWithValue("@obs25_3", SafeVal(obs25(2)))
                    cmd.Parameters.AddWithValue("@obs30_1", SafeVal(obs30(0)))
                    cmd.Parameters.AddWithValue("@obs30_2", SafeVal(obs30(1)))
                    cmd.Parameters.AddWithValue("@obs30_3", SafeVal(obs30(2)))
                    cmd.Parameters.AddWithValue("@obs35_1", SafeVal(obs35(0)))
                    cmd.Parameters.AddWithValue("@obs35_2", SafeVal(obs35(1)))
                    cmd.Parameters.AddWithValue("@obs35_3", SafeVal(obs35(2)))
                    cmd.Parameters.AddWithValue("@obs40_1", SafeVal(obs40(0)))
                    cmd.Parameters.AddWithValue("@obs40_2", SafeVal(obs40(1)))
                    cmd.Parameters.AddWithValue("@obs40_3", SafeVal(obs40(2)))
                    cmd.Parameters.AddWithValue("@obs45_1", SafeVal(obs45(0)))
                    cmd.Parameters.AddWithValue("@obs45_2", SafeVal(obs45(1)))
                    cmd.Parameters.AddWithValue("@obs45_3", SafeVal(obs45(2)))
                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertGeartoothMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetGeartoothMicrometer50Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM geartooth_micrometer_50 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetGeartoothMicrometer50Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function UpdateGeartoothMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal obs0 As Object(), ByVal obs5 As Object(), ByVal obs10 As Object(), ByVal obs15 As Object(), ByVal obs20 As Object(), ByVal obs25 As Object(),
                                           ByVal obs30 As Object(), ByVal obs35 As Object(), ByVal obs40 As Object(), ByVal obs45 As Object(), ByVal obs50 As Object(),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE geartooth_micrometer_50 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs0_0_1=@obs0_1, Obs0_0_2=@obs0_2, Obs0_0_3=@obs0_3, Obs5_0_1=@obs5_1, Obs5_0_2=@obs5_2, Obs5_0_3=@obs5_3, " &
                                   "Obs10_0_1=@obs10_1, Obs10_0_2=@obs10_2, Obs10_0_3=@obs10_3, Obs15_0_1=@obs15_1, Obs15_0_2=@obs15_2, Obs15_0_3=@obs15_3, " &
                                   "Obs20_0_1=@obs20_1, Obs20_0_2=@obs20_2, Obs20_0_3=@obs20_3, Obs25_0_1=@obs25_1, Obs25_0_2=@obs25_2, Obs25_0_3=@obs25_3, " &
                                   "Obs30_0_1=@obs30_1, Obs30_0_2=@obs30_2, Obs30_0_3=@obs30_3, Obs35_0_1=@obs35_1, Obs35_0_2=@obs35_2, Obs35_0_3=@obs35_3, " &
                                   "Obs40_0_1=@obs40_1, Obs40_0_2=@obs40_2, Obs40_0_3=@obs40_3, Obs45_0_1=@obs45_1, Obs45_0_2=@obs45_2, Obs45_0_3=@obs45_3, " &
                                   "Obs50_0_1=@obs50_1, Obs50_0_2=@obs50_2, Obs50_0_3=@obs50_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo = @controlNo AND CycleName = @cycleName AND RowType = 'RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs0_1", SafeVal(obs0(0)))
                    cmd.Parameters.AddWithValue("@obs0_2", SafeVal(obs0(1)))
                    cmd.Parameters.AddWithValue("@obs0_3", SafeVal(obs0(2)))
                    cmd.Parameters.AddWithValue("@obs5_1", SafeVal(obs5(0)))
                    cmd.Parameters.AddWithValue("@obs5_2", SafeVal(obs5(1)))
                    cmd.Parameters.AddWithValue("@obs5_3", SafeVal(obs5(2)))
                    cmd.Parameters.AddWithValue("@obs10_1", SafeVal(obs10(0)))
                    cmd.Parameters.AddWithValue("@obs10_2", SafeVal(obs10(1)))
                    cmd.Parameters.AddWithValue("@obs10_3", SafeVal(obs10(2)))
                    cmd.Parameters.AddWithValue("@obs15_1", SafeVal(obs15(0)))
                    cmd.Parameters.AddWithValue("@obs15_2", SafeVal(obs15(1)))
                    cmd.Parameters.AddWithValue("@obs15_3", SafeVal(obs15(2)))
                    cmd.Parameters.AddWithValue("@obs20_1", SafeVal(obs20(0)))
                    cmd.Parameters.AddWithValue("@obs20_2", SafeVal(obs20(1)))
                    cmd.Parameters.AddWithValue("@obs20_3", SafeVal(obs20(2)))
                    cmd.Parameters.AddWithValue("@obs25_1", SafeVal(obs25(0)))
                    cmd.Parameters.AddWithValue("@obs25_2", SafeVal(obs25(1)))
                    cmd.Parameters.AddWithValue("@obs25_3", SafeVal(obs25(2)))
                    cmd.Parameters.AddWithValue("@obs30_1", SafeVal(obs30(0)))
                    cmd.Parameters.AddWithValue("@obs30_2", SafeVal(obs30(1)))
                    cmd.Parameters.AddWithValue("@obs30_3", SafeVal(obs30(2)))
                    cmd.Parameters.AddWithValue("@obs35_1", SafeVal(obs35(0)))
                    cmd.Parameters.AddWithValue("@obs35_2", SafeVal(obs35(1)))
                    cmd.Parameters.AddWithValue("@obs35_3", SafeVal(obs35(2)))
                    cmd.Parameters.AddWithValue("@obs40_1", SafeVal(obs40(0)))
                    cmd.Parameters.AddWithValue("@obs40_2", SafeVal(obs40(1)))
                    cmd.Parameters.AddWithValue("@obs40_3", SafeVal(obs40(2)))
                    cmd.Parameters.AddWithValue("@obs45_1", SafeVal(obs45(0)))
                    cmd.Parameters.AddWithValue("@obs45_2", SafeVal(obs45(1)))
                    cmd.Parameters.AddWithValue("@obs45_3", SafeVal(obs45(2)))
                    cmd.Parameters.AddWithValue("@obs50_1", SafeVal(obs50(0)))
                    cmd.Parameters.AddWithValue("@obs50_2", SafeVal(obs50(1)))
                    cmd.Parameters.AddWithValue("@obs50_3", SafeVal(obs50(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateGeartoothMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from point_micrometer_25.
    ''' </summary>
    Public Function GetPointMicrometer25MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM point_micrometer_25 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetPointMicrometer25MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into point_micrometer_25.
    ''' </summary>
    Public Function InsertPointMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs2_5 As Object(), ByVal obs5_1 As Object(), ByVal obs7_7 As Object(), ByVal obs10_3 As Object(), ByVal obs12_9 As Object(), ByVal obs15_0 As Object(), ByVal obs17_6 As Object(), ByVal obs20_2 As Object(), ByVal obs22_8 As Object(), ByVal obs25_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO point_micrometer_25 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs2_5_1, Obs2_5_2, Obs2_5_3, Obs5_1_1, Obs5_1_2, Obs5_1_3, Obs7_7_1, Obs7_7_2, Obs7_7_3, " &
                                   "Obs10_3_1, Obs10_3_2, Obs10_3_3, Obs12_9_1, Obs12_9_2, Obs12_9_3, Obs15_0_1, Obs15_0_2, Obs15_0_3, " &
                                   "Obs17_6_1, Obs17_6_2, Obs17_6_3, Obs20_2_1, Obs20_2_2, Obs20_2_3, Obs22_8_1, Obs22_8_2, Obs22_8_3, " &
                                   "Obs25_0_1, Obs25_0_2, Obs25_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs2_5_1, @obs2_5_2, @obs2_5_3, @obs5_1_1, @obs5_1_2, @obs5_1_3, @obs7_7_1, @obs7_7_2, @obs7_7_3, " &
                                   "@obs10_3_1, @obs10_3_2, @obs10_3_3, @obs12_9_1, @obs12_9_2, @obs12_9_3, @obs15_0_1, @obs15_0_2, @obs15_0_3, " &
                                   "@obs17_6_1, @obs17_6_2, @obs17_6_3, @obs20_2_1, @obs20_2_2, @obs20_2_3, @obs22_8_1, @obs22_8_2, @obs22_8_3, " &
                                   "@obs25_0_1, @obs25_0_2, @obs25_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs2_5_1", SafeVal(obs2_5(0)))
                    cmd.Parameters.AddWithValue("@obs2_5_2", SafeVal(obs2_5(1)))
                    cmd.Parameters.AddWithValue("@obs2_5_3", SafeVal(obs2_5(2)))

                    cmd.Parameters.AddWithValue("@obs5_1_1", SafeVal(obs5_1(0)))
                    cmd.Parameters.AddWithValue("@obs5_1_2", SafeVal(obs5_1(1)))
                    cmd.Parameters.AddWithValue("@obs5_1_3", SafeVal(obs5_1(2)))

                    cmd.Parameters.AddWithValue("@obs7_7_1", SafeVal(obs7_7(0)))
                    cmd.Parameters.AddWithValue("@obs7_7_2", SafeVal(obs7_7(1)))
                    cmd.Parameters.AddWithValue("@obs7_7_3", SafeVal(obs7_7(2)))

                    cmd.Parameters.AddWithValue("@obs10_3_1", SafeVal(obs10_3(0)))
                    cmd.Parameters.AddWithValue("@obs10_3_2", SafeVal(obs10_3(1)))
                    cmd.Parameters.AddWithValue("@obs10_3_3", SafeVal(obs10_3(2)))

                    cmd.Parameters.AddWithValue("@obs12_9_1", SafeVal(obs12_9(0)))
                    cmd.Parameters.AddWithValue("@obs12_9_2", SafeVal(obs12_9(1)))
                    cmd.Parameters.AddWithValue("@obs12_9_3", SafeVal(obs12_9(2)))

                    cmd.Parameters.AddWithValue("@obs15_0_1", SafeVal(obs15_0(0)))
                    cmd.Parameters.AddWithValue("@obs15_0_2", SafeVal(obs15_0(1)))
                    cmd.Parameters.AddWithValue("@obs15_0_3", SafeVal(obs15_0(2)))

                    cmd.Parameters.AddWithValue("@obs17_6_1", SafeVal(obs17_6(0)))
                    cmd.Parameters.AddWithValue("@obs17_6_2", SafeVal(obs17_6(1)))
                    cmd.Parameters.AddWithValue("@obs17_6_3", SafeVal(obs17_6(2)))

                    cmd.Parameters.AddWithValue("@obs20_2_1", SafeVal(obs20_2(0)))
                    cmd.Parameters.AddWithValue("@obs20_2_2", SafeVal(obs20_2(1)))
                    cmd.Parameters.AddWithValue("@obs20_2_3", SafeVal(obs20_2(2)))

                    cmd.Parameters.AddWithValue("@obs22_8_1", SafeVal(obs22_8(0)))
                    cmd.Parameters.AddWithValue("@obs22_8_2", SafeVal(obs22_8(1)))
                    cmd.Parameters.AddWithValue("@obs22_8_3", SafeVal(obs22_8(2)))

                    cmd.Parameters.AddWithValue("@obs25_0_1", SafeVal(obs25_0(0)))
                    cmd.Parameters.AddWithValue("@obs25_0_2", SafeVal(obs25_0(1)))
                    cmd.Parameters.AddWithValue("@obs25_0_3", SafeVal(obs25_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertPointMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from point_micrometer_25.
    ''' </summary>
    Public Function GetPointMicrometer25Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM point_micrometer_25 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetPointMicrometer25Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in point_micrometer_25.
    ''' </summary>
    Public Function UpdatePointMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs2_5 As Object(), ByVal obs5_1 As Object(), ByVal obs7_7 As Object(), ByVal obs10_3 As Object(), ByVal obs12_9 As Object(), ByVal obs15_0 As Object(), ByVal obs17_6 As Object(), ByVal obs20_2 As Object(), ByVal obs22_8 As Object(), ByVal obs25_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE point_micrometer_25 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs2_5_1=@obs2_5_1, Obs2_5_2=@obs2_5_2, Obs2_5_3=@obs2_5_3, Obs5_1_1=@obs5_1_1, Obs5_1_2=@obs5_1_2, Obs5_1_3=@obs5_1_3, Obs7_7_1=@obs7_7_1, Obs7_7_2=@obs7_7_2, Obs7_7_3=@obs7_7_3, " &
                                   "Obs10_3_1=@obs10_3_1, Obs10_3_2=@obs10_3_2, Obs10_3_3=@obs10_3_3, Obs12_9_1=@obs12_9_1, Obs12_9_2=@obs12_9_2, Obs12_9_3=@obs12_9_3, Obs15_0_1=@obs15_0_1, Obs15_0_2=@obs15_0_2, Obs15_0_3=@obs15_0_3, " &
                                   "Obs17_6_1=@obs17_6_1, Obs17_6_2=@obs17_6_2, Obs17_6_3=@obs17_6_3, Obs20_2_1=@obs20_2_1, Obs20_2_2=@obs20_2_2, Obs20_2_3=@obs20_2_3, Obs22_8_1=@obs22_8_1, Obs22_8_2=@obs22_8_2, Obs22_8_3=@obs22_8_3, " &
                                   "Obs25_0_1=@obs25_0_1, Obs25_0_2=@obs25_0_2, Obs25_0_3=@obs25_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs2_5_1", SafeVal(obs2_5(0)))
                    cmd.Parameters.AddWithValue("@obs2_5_2", SafeVal(obs2_5(1)))
                    cmd.Parameters.AddWithValue("@obs2_5_3", SafeVal(obs2_5(2)))

                    cmd.Parameters.AddWithValue("@obs5_1_1", SafeVal(obs5_1(0)))
                    cmd.Parameters.AddWithValue("@obs5_1_2", SafeVal(obs5_1(1)))
                    cmd.Parameters.AddWithValue("@obs5_1_3", SafeVal(obs5_1(2)))

                    cmd.Parameters.AddWithValue("@obs7_7_1", SafeVal(obs7_7(0)))
                    cmd.Parameters.AddWithValue("@obs7_7_2", SafeVal(obs7_7(1)))
                    cmd.Parameters.AddWithValue("@obs7_7_3", SafeVal(obs7_7(2)))

                    cmd.Parameters.AddWithValue("@obs10_3_1", SafeVal(obs10_3(0)))
                    cmd.Parameters.AddWithValue("@obs10_3_2", SafeVal(obs10_3(1)))
                    cmd.Parameters.AddWithValue("@obs10_3_3", SafeVal(obs10_3(2)))

                    cmd.Parameters.AddWithValue("@obs12_9_1", SafeVal(obs12_9(0)))
                    cmd.Parameters.AddWithValue("@obs12_9_2", SafeVal(obs12_9(1)))
                    cmd.Parameters.AddWithValue("@obs12_9_3", SafeVal(obs12_9(2)))

                    cmd.Parameters.AddWithValue("@obs15_0_1", SafeVal(obs15_0(0)))
                    cmd.Parameters.AddWithValue("@obs15_0_2", SafeVal(obs15_0(1)))
                    cmd.Parameters.AddWithValue("@obs15_0_3", SafeVal(obs15_0(2)))

                    cmd.Parameters.AddWithValue("@obs17_6_1", SafeVal(obs17_6(0)))
                    cmd.Parameters.AddWithValue("@obs17_6_2", SafeVal(obs17_6(1)))
                    cmd.Parameters.AddWithValue("@obs17_6_3", SafeVal(obs17_6(2)))

                    cmd.Parameters.AddWithValue("@obs20_2_1", SafeVal(obs20_2(0)))
                    cmd.Parameters.AddWithValue("@obs20_2_2", SafeVal(obs20_2(1)))
                    cmd.Parameters.AddWithValue("@obs20_2_3", SafeVal(obs20_2(2)))

                    cmd.Parameters.AddWithValue("@obs22_8_1", SafeVal(obs22_8(0)))
                    cmd.Parameters.AddWithValue("@obs22_8_2", SafeVal(obs22_8(1)))
                    cmd.Parameters.AddWithValue("@obs22_8_3", SafeVal(obs22_8(2)))

                    cmd.Parameters.AddWithValue("@obs25_0_1", SafeVal(obs25_0(0)))
                    cmd.Parameters.AddWithValue("@obs25_0_2", SafeVal(obs25_0(1)))
                    cmd.Parameters.AddWithValue("@obs25_0_3", SafeVal(obs25_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdatePointMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from point_micrometer_50.
    ''' </summary>

    Public Function GetPointMicrometer50MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM point_micrometer_50 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetPointMicrometer50MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into point_micrometer_50.
    ''' </summary>
    Public Function InsertPointMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs27_5 As Object(), ByVal obs30_1 As Object(), ByVal obs32_7 As Object(), ByVal obs35_3 As Object(), ByVal obs37_9 As Object(), ByVal obs40_0 As Object(), ByVal obs42_6 As Object(), ByVal obs45_2 As Object(), ByVal obs47_8 As Object(), ByVal obs50_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO point_micrometer_50 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs27_5_1, Obs27_5_2, Obs27_5_3, Obs30_1_1, Obs30_1_2, Obs30_1_3, Obs32_7_1, Obs32_7_2, Obs32_7_3, " &
                                   "Obs35_3_1, Obs35_3_2, Obs35_3_3, Obs37_9_1, Obs37_9_2, Obs37_9_3, Obs40_0_1, Obs40_0_2, Obs40_0_3, " &
                                   "Obs42_6_1, Obs42_6_2, Obs42_6_3, Obs45_2_1, Obs45_2_2, Obs45_2_3, Obs47_8_1, Obs47_8_2, Obs47_8_3, " &
                                   "Obs50_0_1, Obs50_0_2, Obs50_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs27_5_1, @obs27_5_2, @obs27_5_3, @obs30_1_1, @obs30_1_2, @obs30_1_3, @obs32_7_1, @obs32_7_2, @obs32_7_3, " &
                                   "@obs35_3_1, @obs35_3_2, @obs35_3_3, @obs37_9_1, @obs37_9_2, @obs37_9_3, @obs40_0_1, @obs40_0_2, @obs40_0_3, " &
                                   "@obs42_6_1, @obs42_6_2, @obs42_6_3, @obs45_2_1, @obs45_2_2, @obs45_2_3, @obs47_8_1, @obs47_8_2, @obs47_8_3, " &
                                   "@obs50_0_1, @obs50_0_2, @obs50_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_1_1", SafeVal(obs30_1(0)))
                    cmd.Parameters.AddWithValue("@obs30_1_2", SafeVal(obs30_1(1)))
                    cmd.Parameters.AddWithValue("@obs30_1_3", SafeVal(obs30_1(2)))

                    cmd.Parameters.AddWithValue("@obs32_7_1", SafeVal(obs32_7(0)))
                    cmd.Parameters.AddWithValue("@obs32_7_2", SafeVal(obs32_7(1)))
                    cmd.Parameters.AddWithValue("@obs32_7_3", SafeVal(obs32_7(2)))

                    cmd.Parameters.AddWithValue("@obs35_3_1", SafeVal(obs35_3(0)))
                    cmd.Parameters.AddWithValue("@obs35_3_2", SafeVal(obs35_3(1)))
                    cmd.Parameters.AddWithValue("@obs35_3_3", SafeVal(obs35_3(2)))

                    cmd.Parameters.AddWithValue("@obs37_9_1", SafeVal(obs37_9(0)))
                    cmd.Parameters.AddWithValue("@obs37_9_2", SafeVal(obs37_9(1)))
                    cmd.Parameters.AddWithValue("@obs37_9_3", SafeVal(obs37_9(2)))

                    cmd.Parameters.AddWithValue("@obs40_0_1", SafeVal(obs40_0(0)))
                    cmd.Parameters.AddWithValue("@obs40_0_2", SafeVal(obs40_0(1)))
                    cmd.Parameters.AddWithValue("@obs40_0_3", SafeVal(obs40_0(2)))

                    cmd.Parameters.AddWithValue("@obs42_6_1", SafeVal(obs42_6(0)))
                    cmd.Parameters.AddWithValue("@obs42_6_2", SafeVal(obs42_6(1)))
                    cmd.Parameters.AddWithValue("@obs42_6_3", SafeVal(obs42_6(2)))

                    cmd.Parameters.AddWithValue("@obs45_2_1", SafeVal(obs45_2(0)))
                    cmd.Parameters.AddWithValue("@obs45_2_2", SafeVal(obs45_2(1)))
                    cmd.Parameters.AddWithValue("@obs45_2_3", SafeVal(obs45_2(2)))

                    cmd.Parameters.AddWithValue("@obs47_8_1", SafeVal(obs47_8(0)))
                    cmd.Parameters.AddWithValue("@obs47_8_2", SafeVal(obs47_8(1)))
                    cmd.Parameters.AddWithValue("@obs47_8_3", SafeVal(obs47_8(2)))

                    cmd.Parameters.AddWithValue("@obs50_0_1", SafeVal(obs50_0(0)))
                    cmd.Parameters.AddWithValue("@obs50_0_2", SafeVal(obs50_0(1)))
                    cmd.Parameters.AddWithValue("@obs50_0_3", SafeVal(obs50_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertPointMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from point_micrometer_50.
    ''' </summary>
    Public Function GetPointMicrometer50Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM point_micrometer_50 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetPointMicrometer50Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in point_micrometer_50.
    ''' </summary>
    Public Function UpdatePointMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs27_5 As Object(), ByVal obs30_1 As Object(), ByVal obs32_7 As Object(), ByVal obs35_3 As Object(), ByVal obs37_9 As Object(), ByVal obs40_0 As Object(), ByVal obs42_6 As Object(), ByVal obs45_2 As Object(), ByVal obs47_8 As Object(), ByVal obs50_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE point_micrometer_50 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs27_5_1=@obs27_5_1, Obs27_5_2=@obs27_5_2, Obs27_5_3=@obs27_5_3, Obs30_1_1=@obs30_1_1, Obs30_1_2=@obs30_1_2, Obs30_1_3=@obs30_1_3, Obs32_7_1=@obs32_7_1, Obs32_7_2=@obs32_7_2, Obs32_7_3=@obs32_7_3, " &
                                   "Obs35_3_1=@obs35_3_1, Obs35_3_2=@obs35_3_2, Obs35_3_3=@obs35_3_3, Obs37_9_1=@obs37_9_1, Obs37_9_2=@obs37_9_2, Obs37_9_3=@obs37_9_3, Obs40_0_1=@obs40_0_1, Obs40_0_2=@obs40_0_2, Obs40_0_3=@obs40_0_3, " &
                                   "Obs42_6_1=@obs42_6_1, Obs42_6_2=@obs42_6_2, Obs42_6_3=@obs42_6_3, Obs45_2_1=@obs45_2_1, Obs45_2_2=@obs45_2_2, Obs45_2_3=@obs45_2_3, Obs47_8_1=@obs47_8_1, Obs47_8_2=@obs47_8_2, Obs47_8_3=@obs47_8_3, " &
                                   "Obs50_0_1=@obs50_0_1, Obs50_0_2=@obs50_0_2, Obs50_0_3=@obs50_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_1_1", SafeVal(obs30_1(0)))
                    cmd.Parameters.AddWithValue("@obs30_1_2", SafeVal(obs30_1(1)))
                    cmd.Parameters.AddWithValue("@obs30_1_3", SafeVal(obs30_1(2)))

                    cmd.Parameters.AddWithValue("@obs32_7_1", SafeVal(obs32_7(0)))
                    cmd.Parameters.AddWithValue("@obs32_7_2", SafeVal(obs32_7(1)))
                    cmd.Parameters.AddWithValue("@obs32_7_3", SafeVal(obs32_7(2)))

                    cmd.Parameters.AddWithValue("@obs35_3_1", SafeVal(obs35_3(0)))
                    cmd.Parameters.AddWithValue("@obs35_3_2", SafeVal(obs35_3(1)))
                    cmd.Parameters.AddWithValue("@obs35_3_3", SafeVal(obs35_3(2)))

                    cmd.Parameters.AddWithValue("@obs37_9_1", SafeVal(obs37_9(0)))
                    cmd.Parameters.AddWithValue("@obs37_9_2", SafeVal(obs37_9(1)))
                    cmd.Parameters.AddWithValue("@obs37_9_3", SafeVal(obs37_9(2)))

                    cmd.Parameters.AddWithValue("@obs40_0_1", SafeVal(obs40_0(0)))
                    cmd.Parameters.AddWithValue("@obs40_0_2", SafeVal(obs40_0(1)))
                    cmd.Parameters.AddWithValue("@obs40_0_3", SafeVal(obs40_0(2)))

                    cmd.Parameters.AddWithValue("@obs42_6_1", SafeVal(obs42_6(0)))
                    cmd.Parameters.AddWithValue("@obs42_6_2", SafeVal(obs42_6(1)))
                    cmd.Parameters.AddWithValue("@obs42_6_3", SafeVal(obs42_6(2)))

                    cmd.Parameters.AddWithValue("@obs45_2_1", SafeVal(obs45_2(0)))
                    cmd.Parameters.AddWithValue("@obs45_2_2", SafeVal(obs45_2(1)))
                    cmd.Parameters.AddWithValue("@obs45_2_3", SafeVal(obs45_2(2)))

                    cmd.Parameters.AddWithValue("@obs47_8_1", SafeVal(obs47_8(0)))
                    cmd.Parameters.AddWithValue("@obs47_8_2", SafeVal(obs47_8(1)))
                    cmd.Parameters.AddWithValue("@obs47_8_3", SafeVal(obs47_8(2)))

                    cmd.Parameters.AddWithValue("@obs50_0_1", SafeVal(obs50_0(0)))
                    cmd.Parameters.AddWithValue("@obs50_0_2", SafeVal(obs50_0(1)))
                    cmd.Parameters.AddWithValue("@obs50_0_3", SafeVal(obs50_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdatePointMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from thread_pitch_micrometer_50.
    ''' </summary>
    Public Function GetThreadPitchMicrometer50MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_50 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer50MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into thread_pitch_micrometer_50.
    ''' </summary>
    Public Function InsertThreadPitchMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs27_5 As Object(), ByVal obs30_1 As Object(), ByVal obs32_7 As Object(), ByVal obs35_3 As Object(), ByVal obs37_9 As Object(), ByVal obs40_0 As Object(), ByVal obs42_6 As Object(), ByVal obs45_2 As Object(), ByVal obs47_8 As Object(), ByVal obs50_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_pitch_micrometer_50 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs27_5_1, Obs27_5_2, Obs27_5_3, Obs30_1_1, Obs30_1_2, Obs30_1_3, Obs32_7_1, Obs32_7_2, Obs32_7_3, " &
                                   "Obs35_3_1, Obs35_3_2, Obs35_3_3, Obs37_9_1, Obs37_9_2, Obs37_9_3, Obs40_0_1, Obs40_0_2, Obs40_0_3, " &
                                   "Obs42_6_1, Obs42_6_2, Obs42_6_3, Obs45_2_1, Obs45_2_2, Obs45_2_3, Obs47_8_1, Obs47_8_2, Obs47_8_3, " &
                                   "Obs50_0_1, Obs50_0_2, Obs50_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs27_5_1, @obs27_5_2, @obs27_5_3, @obs30_1_1, @obs30_1_2, @obs30_1_3, @obs32_7_1, @obs32_7_2, @obs32_7_3, " &
                                   "@obs35_3_1, @obs35_3_2, @obs35_3_3, @obs37_9_1, @obs37_9_2, @obs37_9_3, @obs40_0_1, @obs40_0_2, @obs40_0_3, " &
                                   "@obs42_6_1, @obs42_6_2, @obs42_6_3, @obs45_2_1, @obs45_2_2, @obs45_2_3, @obs47_8_1, @obs47_8_2, @obs47_8_3, " &
                                   "@obs50_0_1, @obs50_0_2, @obs50_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_1_1", SafeVal(obs30_1(0)))
                    cmd.Parameters.AddWithValue("@obs30_1_2", SafeVal(obs30_1(1)))
                    cmd.Parameters.AddWithValue("@obs30_1_3", SafeVal(obs30_1(2)))

                    cmd.Parameters.AddWithValue("@obs32_7_1", SafeVal(obs32_7(0)))
                    cmd.Parameters.AddWithValue("@obs32_7_2", SafeVal(obs32_7(1)))
                    cmd.Parameters.AddWithValue("@obs32_7_3", SafeVal(obs32_7(2)))

                    cmd.Parameters.AddWithValue("@obs35_3_1", SafeVal(obs35_3(0)))
                    cmd.Parameters.AddWithValue("@obs35_3_2", SafeVal(obs35_3(1)))
                    cmd.Parameters.AddWithValue("@obs35_3_3", SafeVal(obs35_3(2)))

                    cmd.Parameters.AddWithValue("@obs37_9_1", SafeVal(obs37_9(0)))
                    cmd.Parameters.AddWithValue("@obs37_9_2", SafeVal(obs37_9(1)))
                    cmd.Parameters.AddWithValue("@obs37_9_3", SafeVal(obs37_9(2)))

                    cmd.Parameters.AddWithValue("@obs40_0_1", SafeVal(obs40_0(0)))
                    cmd.Parameters.AddWithValue("@obs40_0_2", SafeVal(obs40_0(1)))
                    cmd.Parameters.AddWithValue("@obs40_0_3", SafeVal(obs40_0(2)))

                    cmd.Parameters.AddWithValue("@obs42_6_1", SafeVal(obs42_6(0)))
                    cmd.Parameters.AddWithValue("@obs42_6_2", SafeVal(obs42_6(1)))
                    cmd.Parameters.AddWithValue("@obs42_6_3", SafeVal(obs42_6(2)))

                    cmd.Parameters.AddWithValue("@obs45_2_1", SafeVal(obs45_2(0)))
                    cmd.Parameters.AddWithValue("@obs45_2_2", SafeVal(obs45_2(1)))
                    cmd.Parameters.AddWithValue("@obs45_2_3", SafeVal(obs45_2(2)))

                    cmd.Parameters.AddWithValue("@obs47_8_1", SafeVal(obs47_8(0)))
                    cmd.Parameters.AddWithValue("@obs47_8_2", SafeVal(obs47_8(1)))
                    cmd.Parameters.AddWithValue("@obs47_8_3", SafeVal(obs47_8(2)))

                    cmd.Parameters.AddWithValue("@obs50_0_1", SafeVal(obs50_0(0)))
                    cmd.Parameters.AddWithValue("@obs50_0_2", SafeVal(obs50_0(1)))
                    cmd.Parameters.AddWithValue("@obs50_0_3", SafeVal(obs50_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertThreadPitchMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from thread_pitch_micrometer_50.
    ''' </summary>
    Public Function GetThreadPitchMicrometer50Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_50 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer50Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in thread_pitch_micrometer_50.
    ''' </summary>
    Public Function UpdateThreadPitchMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs27_5 As Object(), ByVal obs30_1 As Object(), ByVal obs32_7 As Object(), ByVal obs35_3 As Object(), ByVal obs37_9 As Object(), ByVal obs40_0 As Object(), ByVal obs42_6 As Object(), ByVal obs45_2 As Object(), ByVal obs47_8 As Object(), ByVal obs50_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE thread_pitch_micrometer_50 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs27_5_1=@obs27_5_1, Obs27_5_2=@obs27_5_2, Obs27_5_3=@obs27_5_3, Obs30_1_1=@obs30_1_1, Obs30_1_2=@obs30_1_2, Obs30_1_3=@obs30_1_3, Obs32_7_1=@obs32_7_1, Obs32_7_2=@obs32_7_2, Obs32_7_3=@obs32_7_3, " &
                                   "Obs35_3_1=@obs35_3_1, Obs35_3_2=@obs35_3_2, Obs35_3_3=@obs35_3_3, Obs37_9_1=@obs37_9_1, Obs37_9_2=@obs37_9_2, Obs37_9_3=@obs37_9_3, Obs40_0_1=@obs40_0_1, Obs40_0_2=@obs40_0_2, Obs40_0_3=@obs40_0_3, " &
                                   "Obs42_6_1=@obs42_6_1, Obs42_6_2=@obs42_6_2, Obs42_6_3=@obs42_6_3, Obs45_2_1=@obs45_2_1, Obs45_2_2=@obs45_2_2, Obs45_2_3=@obs45_2_3, Obs47_8_1=@obs47_8_1, Obs47_8_2=@obs47_8_2, Obs47_8_3=@obs47_8_3, " &
                                   "Obs50_0_1=@obs50_0_1, Obs50_0_2=@obs50_0_2, Obs50_0_3=@obs50_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_1_1", SafeVal(obs30_1(0)))
                    cmd.Parameters.AddWithValue("@obs30_1_2", SafeVal(obs30_1(1)))
                    cmd.Parameters.AddWithValue("@obs30_1_3", SafeVal(obs30_1(2)))

                    cmd.Parameters.AddWithValue("@obs32_7_1", SafeVal(obs32_7(0)))
                    cmd.Parameters.AddWithValue("@obs32_7_2", SafeVal(obs32_7(1)))
                    cmd.Parameters.AddWithValue("@obs32_7_3", SafeVal(obs32_7(2)))

                    cmd.Parameters.AddWithValue("@obs35_3_1", SafeVal(obs35_3(0)))
                    cmd.Parameters.AddWithValue("@obs35_3_2", SafeVal(obs35_3(1)))
                    cmd.Parameters.AddWithValue("@obs35_3_3", SafeVal(obs35_3(2)))

                    cmd.Parameters.AddWithValue("@obs37_9_1", SafeVal(obs37_9(0)))
                    cmd.Parameters.AddWithValue("@obs37_9_2", SafeVal(obs37_9(1)))
                    cmd.Parameters.AddWithValue("@obs37_9_3", SafeVal(obs37_9(2)))

                    cmd.Parameters.AddWithValue("@obs40_0_1", SafeVal(obs40_0(0)))
                    cmd.Parameters.AddWithValue("@obs40_0_2", SafeVal(obs40_0(1)))
                    cmd.Parameters.AddWithValue("@obs40_0_3", SafeVal(obs40_0(2)))

                    cmd.Parameters.AddWithValue("@obs42_6_1", SafeVal(obs42_6(0)))
                    cmd.Parameters.AddWithValue("@obs42_6_2", SafeVal(obs42_6(1)))
                    cmd.Parameters.AddWithValue("@obs42_6_3", SafeVal(obs42_6(2)))

                    cmd.Parameters.AddWithValue("@obs45_2_1", SafeVal(obs45_2(0)))
                    cmd.Parameters.AddWithValue("@obs45_2_2", SafeVal(obs45_2(1)))
                    cmd.Parameters.AddWithValue("@obs45_2_3", SafeVal(obs45_2(2)))

                    cmd.Parameters.AddWithValue("@obs47_8_1", SafeVal(obs47_8(0)))
                    cmd.Parameters.AddWithValue("@obs47_8_2", SafeVal(obs47_8(1)))
                    cmd.Parameters.AddWithValue("@obs47_8_3", SafeVal(obs47_8(2)))

                    cmd.Parameters.AddWithValue("@obs50_0_1", SafeVal(obs50_0(0)))
                    cmd.Parameters.AddWithValue("@obs50_0_2", SafeVal(obs50_0(1)))
                    cmd.Parameters.AddWithValue("@obs50_0_3", SafeVal(obs50_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateThreadPitchMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function


    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from inside_micrometer_25.
    ''' </summary>
    Public Function GetInsideMicrometer25MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM inside_micrometer_25 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetInsideMicrometer25MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into inside_micrometer_25.
    ''' </summary>
    Public Function InsertInsideMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs5_1 As Object(), ByVal obs7_7 As Object(), ByVal obs10_3 As Object(), ByVal obs12_9 As Object(), ByVal obs15_0 As Object(), ByVal obs17_6 As Object(), ByVal obs20_2 As Object(), ByVal obs22_8 As Object(), ByVal obs25_0 As Object(), ByVal obs27_5 As Object(), ByVal obs30_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO inside_micrometer_25 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs5_1_1, Obs5_1_2, Obs5_1_3, Obs7_7_1, Obs7_7_2, Obs7_7_3, Obs10_3_1, Obs10_3_2, Obs10_3_3, " &
                                   "Obs12_9_1, Obs12_9_2, Obs12_9_3, Obs15_0_1, Obs15_0_2, Obs15_0_3, Obs17_6_1, Obs17_6_2, Obs17_6_3, " &
                                   "Obs20_2_1, Obs20_2_2, Obs20_2_3, Obs22_8_1, Obs22_8_2, Obs22_8_3, Obs25_0_1, Obs25_0_2, Obs25_0_3, " &
                                   "Obs27_5_1, Obs27_5_2, Obs27_5_3, Obs30_0_1, Obs30_0_2, Obs30_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs5_1_1, @obs5_1_2, @obs5_1_3, @obs7_7_1, @obs7_7_2, @obs7_7_3, @obs10_3_1, @obs10_3_2, @obs10_3_3, " &
                                   "@obs12_9_1, @obs12_9_2, @obs12_9_3, @obs15_0_1, @obs15_0_2, @obs15_0_3, @obs17_6_1, @obs17_6_2, @obs17_6_3, " &
                                   "@obs20_2_1, @obs20_2_2, @obs20_2_3, @obs22_8_1, @obs22_8_2, @obs22_8_3, @obs25_0_1, @obs25_0_2, @obs25_0_3, " &
                                   "@obs27_5_1, @obs27_5_2, @obs27_5_3, @obs30_0_1, @obs30_0_2, @obs30_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs5_1_1", SafeVal(obs5_1(0)))
                    cmd.Parameters.AddWithValue("@obs5_1_2", SafeVal(obs5_1(1)))
                    cmd.Parameters.AddWithValue("@obs5_1_3", SafeVal(obs5_1(2)))

                    cmd.Parameters.AddWithValue("@obs7_7_1", SafeVal(obs7_7(0)))
                    cmd.Parameters.AddWithValue("@obs7_7_2", SafeVal(obs7_7(1)))
                    cmd.Parameters.AddWithValue("@obs7_7_3", SafeVal(obs7_7(2)))

                    cmd.Parameters.AddWithValue("@obs10_3_1", SafeVal(obs10_3(0)))
                    cmd.Parameters.AddWithValue("@obs10_3_2", SafeVal(obs10_3(1)))
                    cmd.Parameters.AddWithValue("@obs10_3_3", SafeVal(obs10_3(2)))

                    cmd.Parameters.AddWithValue("@obs12_9_1", SafeVal(obs12_9(0)))
                    cmd.Parameters.AddWithValue("@obs12_9_2", SafeVal(obs12_9(1)))
                    cmd.Parameters.AddWithValue("@obs12_9_3", SafeVal(obs12_9(2)))

                    cmd.Parameters.AddWithValue("@obs15_0_1", SafeVal(obs15_0(0)))
                    cmd.Parameters.AddWithValue("@obs15_0_2", SafeVal(obs15_0(1)))
                    cmd.Parameters.AddWithValue("@obs15_0_3", SafeVal(obs15_0(2)))

                    cmd.Parameters.AddWithValue("@obs17_6_1", SafeVal(obs17_6(0)))
                    cmd.Parameters.AddWithValue("@obs17_6_2", SafeVal(obs17_6(1)))
                    cmd.Parameters.AddWithValue("@obs17_6_3", SafeVal(obs17_6(2)))

                    cmd.Parameters.AddWithValue("@obs20_2_1", SafeVal(obs20_2(0)))
                    cmd.Parameters.AddWithValue("@obs20_2_2", SafeVal(obs20_2(1)))
                    cmd.Parameters.AddWithValue("@obs20_2_3", SafeVal(obs20_2(2)))

                    cmd.Parameters.AddWithValue("@obs22_8_1", SafeVal(obs22_8(0)))
                    cmd.Parameters.AddWithValue("@obs22_8_2", SafeVal(obs22_8(1)))
                    cmd.Parameters.AddWithValue("@obs22_8_3", SafeVal(obs22_8(2)))

                    cmd.Parameters.AddWithValue("@obs25_0_1", SafeVal(obs25_0(0)))
                    cmd.Parameters.AddWithValue("@obs25_0_2", SafeVal(obs25_0(1)))
                    cmd.Parameters.AddWithValue("@obs25_0_3", SafeVal(obs25_0(2)))

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_0_1", SafeVal(obs30_0(0)))
                    cmd.Parameters.AddWithValue("@obs30_0_2", SafeVal(obs30_0(1)))
                    cmd.Parameters.AddWithValue("@obs30_0_3", SafeVal(obs30_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertInsideMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from inside_micrometer_25.
    ''' </summary>
    Public Function GetInsideMicrometer25Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM inside_micrometer_25 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetInsideMicrometer25Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in inside_micrometer_25.
    ''' </summary>
    Public Function UpdateInsideMicrometer25(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs5_1 As Object(), ByVal obs7_7 As Object(), ByVal obs10_3 As Object(), ByVal obs12_9 As Object(), ByVal obs15_0 As Object(), ByVal obs17_6 As Object(), ByVal obs20_2 As Object(), ByVal obs22_8 As Object(), ByVal obs25_0 As Object(), ByVal obs27_5 As Object(), ByVal obs30_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE inside_micrometer_25 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs5_1_1=@obs5_1_1, Obs5_1_2=@obs5_1_2, Obs5_1_3=@obs5_1_3, Obs7_7_1=@obs7_7_1, Obs7_7_2=@obs7_7_2, Obs7_7_3=@obs7_7_3, Obs10_3_1=@obs10_3_1, Obs10_3_2=@obs10_3_2, Obs10_3_3=@obs10_3_3, " &
                                   "Obs12_9_1=@obs12_9_1, Obs12_9_2=@obs12_9_2, Obs12_9_3=@obs12_9_3, Obs15_0_1=@obs15_0_1, Obs15_0_2=@obs15_0_2, Obs15_0_3=@obs15_0_3, Obs17_6_1=@obs17_6_1, Obs17_6_2=@obs17_6_2, Obs17_6_3=@obs17_6_3, " &
                                   "Obs20_2_1=@obs20_2_1, Obs20_2_2=@obs20_2_2, Obs20_2_3=@obs20_2_3, Obs22_8_1=@obs22_8_1, Obs22_8_2=@obs22_8_2, Obs22_8_3=@obs22_8_3, Obs25_0_1=@obs25_0_1, Obs25_0_2=@obs25_0_2, Obs25_0_3=@obs25_0_3, " &
                                   "Obs27_5_1=@obs27_5_1, Obs27_5_2=@obs27_5_2, Obs27_5_3=@obs27_5_3, Obs30_0_1=@obs30_0_1, Obs30_0_2=@obs30_0_2, Obs30_0_3=@obs30_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs5_1_1", SafeVal(obs5_1(0)))
                    cmd.Parameters.AddWithValue("@obs5_1_2", SafeVal(obs5_1(1)))
                    cmd.Parameters.AddWithValue("@obs5_1_3", SafeVal(obs5_1(2)))

                    cmd.Parameters.AddWithValue("@obs7_7_1", SafeVal(obs7_7(0)))
                    cmd.Parameters.AddWithValue("@obs7_7_2", SafeVal(obs7_7(1)))
                    cmd.Parameters.AddWithValue("@obs7_7_3", SafeVal(obs7_7(2)))

                    cmd.Parameters.AddWithValue("@obs10_3_1", SafeVal(obs10_3(0)))
                    cmd.Parameters.AddWithValue("@obs10_3_2", SafeVal(obs10_3(1)))
                    cmd.Parameters.AddWithValue("@obs10_3_3", SafeVal(obs10_3(2)))

                    cmd.Parameters.AddWithValue("@obs12_9_1", SafeVal(obs12_9(0)))
                    cmd.Parameters.AddWithValue("@obs12_9_2", SafeVal(obs12_9(1)))
                    cmd.Parameters.AddWithValue("@obs12_9_3", SafeVal(obs12_9(2)))

                    cmd.Parameters.AddWithValue("@obs15_0_1", SafeVal(obs15_0(0)))
                    cmd.Parameters.AddWithValue("@obs15_0_2", SafeVal(obs15_0(1)))
                    cmd.Parameters.AddWithValue("@obs15_0_3", SafeVal(obs15_0(2)))

                    cmd.Parameters.AddWithValue("@obs17_6_1", SafeVal(obs17_6(0)))
                    cmd.Parameters.AddWithValue("@obs17_6_2", SafeVal(obs17_6(1)))
                    cmd.Parameters.AddWithValue("@obs17_6_3", SafeVal(obs17_6(2)))

                    cmd.Parameters.AddWithValue("@obs20_2_1", SafeVal(obs20_2(0)))
                    cmd.Parameters.AddWithValue("@obs20_2_2", SafeVal(obs20_2(1)))
                    cmd.Parameters.AddWithValue("@obs20_2_3", SafeVal(obs20_2(2)))

                    cmd.Parameters.AddWithValue("@obs22_8_1", SafeVal(obs22_8(0)))
                    cmd.Parameters.AddWithValue("@obs22_8_2", SafeVal(obs22_8(1)))
                    cmd.Parameters.AddWithValue("@obs22_8_3", SafeVal(obs22_8(2)))

                    cmd.Parameters.AddWithValue("@obs25_0_1", SafeVal(obs25_0(0)))
                    cmd.Parameters.AddWithValue("@obs25_0_2", SafeVal(obs25_0(1)))
                    cmd.Parameters.AddWithValue("@obs25_0_3", SafeVal(obs25_0(2)))

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_0_1", SafeVal(obs30_0(0)))
                    cmd.Parameters.AddWithValue("@obs30_0_2", SafeVal(obs30_0(1)))
                    cmd.Parameters.AddWithValue("@obs30_0_3", SafeVal(obs30_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateInsideMicrometer25 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from inside_micrometer_50.
    ''' </summary>
    Public Function GetInsideMicrometer50MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM inside_micrometer_50 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetInsideMicrometer50MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into inside_micrometer_50.
    ''' </summary>
    Public Function InsertInsideMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs27_5 As Object(), ByVal obs30_1 As Object(), ByVal obs32_7 As Object(), ByVal obs35_3 As Object(), ByVal obs37_9 As Object(), ByVal obs40_0 As Object(), ByVal obs42_6 As Object(), ByVal obs45_2 As Object(), ByVal obs47_8 As Object(), ByVal obs50_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO inside_micrometer_50 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs27_5_1, Obs27_5_2, Obs27_5_3, Obs30_1_1, Obs30_1_2, Obs30_1_3, Obs32_7_1, Obs32_7_2, Obs32_7_3, " &
                                   "Obs35_3_1, Obs35_3_2, Obs35_3_3, Obs37_9_1, Obs37_9_2, Obs37_9_3, Obs40_0_1, Obs40_0_2, Obs40_0_3, " &
                                   "Obs42_6_1, Obs42_6_2, Obs42_6_3, Obs45_2_1, Obs45_2_2, Obs45_2_3, Obs47_8_1, Obs47_8_2, Obs47_8_3, " &
                                   "Obs50_0_1, Obs50_0_2, Obs50_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs27_5_1, @obs27_5_2, @obs27_5_3, @obs30_1_1, @obs30_1_2, @obs30_1_3, @obs32_7_1, @obs32_7_2, @obs32_7_3, " &
                                   "@obs35_3_1, @obs35_3_2, @obs35_3_3, @obs37_9_1, @obs37_9_2, @obs37_9_3, @obs40_0_1, @obs40_0_2, @obs40_0_3, " &
                                   "@obs42_6_1, @obs42_6_2, @obs42_6_3, @obs45_2_1, @obs45_2_2, @obs45_2_3, @obs47_8_1, @obs47_8_2, @obs47_8_3, " &
                                   "@obs50_0_1, @obs50_0_2, @obs50_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_1_1", SafeVal(obs30_1(0)))
                    cmd.Parameters.AddWithValue("@obs30_1_2", SafeVal(obs30_1(1)))
                    cmd.Parameters.AddWithValue("@obs30_1_3", SafeVal(obs30_1(2)))

                    cmd.Parameters.AddWithValue("@obs32_7_1", SafeVal(obs32_7(0)))
                    cmd.Parameters.AddWithValue("@obs32_7_2", SafeVal(obs32_7(1)))
                    cmd.Parameters.AddWithValue("@obs32_7_3", SafeVal(obs32_7(2)))

                    cmd.Parameters.AddWithValue("@obs35_3_1", SafeVal(obs35_3(0)))
                    cmd.Parameters.AddWithValue("@obs35_3_2", SafeVal(obs35_3(1)))
                    cmd.Parameters.AddWithValue("@obs35_3_3", SafeVal(obs35_3(2)))

                    cmd.Parameters.AddWithValue("@obs37_9_1", SafeVal(obs37_9(0)))
                    cmd.Parameters.AddWithValue("@obs37_9_2", SafeVal(obs37_9(1)))
                    cmd.Parameters.AddWithValue("@obs37_9_3", SafeVal(obs37_9(2)))

                    cmd.Parameters.AddWithValue("@obs40_0_1", SafeVal(obs40_0(0)))
                    cmd.Parameters.AddWithValue("@obs40_0_2", SafeVal(obs40_0(1)))
                    cmd.Parameters.AddWithValue("@obs40_0_3", SafeVal(obs40_0(2)))

                    cmd.Parameters.AddWithValue("@obs42_6_1", SafeVal(obs42_6(0)))
                    cmd.Parameters.AddWithValue("@obs42_6_2", SafeVal(obs42_6(1)))
                    cmd.Parameters.AddWithValue("@obs42_6_3", SafeVal(obs42_6(2)))

                    cmd.Parameters.AddWithValue("@obs45_2_1", SafeVal(obs45_2(0)))
                    cmd.Parameters.AddWithValue("@obs45_2_2", SafeVal(obs45_2(1)))
                    cmd.Parameters.AddWithValue("@obs45_2_3", SafeVal(obs45_2(2)))

                    cmd.Parameters.AddWithValue("@obs47_8_1", SafeVal(obs47_8(0)))
                    cmd.Parameters.AddWithValue("@obs47_8_2", SafeVal(obs47_8(1)))
                    cmd.Parameters.AddWithValue("@obs47_8_3", SafeVal(obs47_8(2)))

                    cmd.Parameters.AddWithValue("@obs50_0_1", SafeVal(obs50_0(0)))
                    cmd.Parameters.AddWithValue("@obs50_0_2", SafeVal(obs50_0(1)))
                    cmd.Parameters.AddWithValue("@obs50_0_3", SafeVal(obs50_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertInsideMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from inside_micrometer_50.
    ''' </summary>
    Public Function GetInsideMicrometer50Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM inside_micrometer_50 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetInsideMicrometer50Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in inside_micrometer_50.
    ''' </summary>
    Public Function UpdateInsideMicrometer50(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs27_5 As Object(), ByVal obs30_1 As Object(), ByVal obs32_7 As Object(), ByVal obs35_3 As Object(), ByVal obs37_9 As Object(), ByVal obs40_0 As Object(), ByVal obs42_6 As Object(), ByVal obs45_2 As Object(), ByVal obs47_8 As Object(), ByVal obs50_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE inside_micrometer_50 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs27_5_1=@obs27_5_1, Obs27_5_2=@obs27_5_2, Obs27_5_3=@obs27_5_3, Obs30_1_1=@obs30_1_1, Obs30_1_2=@obs30_1_2, Obs30_1_3=@obs30_1_3, Obs32_7_1=@obs32_7_1, Obs32_7_2=@obs32_7_2, Obs32_7_3=@obs32_7_3, " &
                                   "Obs35_3_1=@obs35_3_1, Obs35_3_2=@obs35_3_2, Obs35_3_3=@obs35_3_3, Obs37_9_1=@obs37_9_1, Obs37_9_2=@obs37_9_2, Obs37_9_3=@obs37_9_3, Obs40_0_1=@obs40_0_1, Obs40_0_2=@obs40_0_2, Obs40_0_3=@obs40_0_3, " &
                                   "Obs42_6_1=@obs42_6_1, Obs42_6_2=@obs42_6_2, Obs42_6_3=@obs42_6_3, Obs45_2_1=@obs45_2_1, Obs45_2_2=@obs45_2_2, Obs45_2_3=@obs45_2_3, Obs47_8_1=@obs47_8_1, Obs47_8_2=@obs47_8_2, Obs47_8_3=@obs47_8_3, " &
                                   "Obs50_0_1=@obs50_0_1, Obs50_0_2=@obs50_0_2, Obs50_0_3=@obs50_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs27_5_1", SafeVal(obs27_5(0)))
                    cmd.Parameters.AddWithValue("@obs27_5_2", SafeVal(obs27_5(1)))
                    cmd.Parameters.AddWithValue("@obs27_5_3", SafeVal(obs27_5(2)))

                    cmd.Parameters.AddWithValue("@obs30_1_1", SafeVal(obs30_1(0)))
                    cmd.Parameters.AddWithValue("@obs30_1_2", SafeVal(obs30_1(1)))
                    cmd.Parameters.AddWithValue("@obs30_1_3", SafeVal(obs30_1(2)))

                    cmd.Parameters.AddWithValue("@obs32_7_1", SafeVal(obs32_7(0)))
                    cmd.Parameters.AddWithValue("@obs32_7_2", SafeVal(obs32_7(1)))
                    cmd.Parameters.AddWithValue("@obs32_7_3", SafeVal(obs32_7(2)))

                    cmd.Parameters.AddWithValue("@obs35_3_1", SafeVal(obs35_3(0)))
                    cmd.Parameters.AddWithValue("@obs35_3_2", SafeVal(obs35_3(1)))
                    cmd.Parameters.AddWithValue("@obs35_3_3", SafeVal(obs35_3(2)))

                    cmd.Parameters.AddWithValue("@obs37_9_1", SafeVal(obs37_9(0)))
                    cmd.Parameters.AddWithValue("@obs37_9_2", SafeVal(obs37_9(1)))
                    cmd.Parameters.AddWithValue("@obs37_9_3", SafeVal(obs37_9(2)))

                    cmd.Parameters.AddWithValue("@obs40_0_1", SafeVal(obs40_0(0)))
                    cmd.Parameters.AddWithValue("@obs40_0_2", SafeVal(obs40_0(1)))
                    cmd.Parameters.AddWithValue("@obs40_0_3", SafeVal(obs40_0(2)))

                    cmd.Parameters.AddWithValue("@obs42_6_1", SafeVal(obs42_6(0)))
                    cmd.Parameters.AddWithValue("@obs42_6_2", SafeVal(obs42_6(1)))
                    cmd.Parameters.AddWithValue("@obs42_6_3", SafeVal(obs42_6(2)))

                    cmd.Parameters.AddWithValue("@obs45_2_1", SafeVal(obs45_2(0)))
                    cmd.Parameters.AddWithValue("@obs45_2_2", SafeVal(obs45_2(1)))
                    cmd.Parameters.AddWithValue("@obs45_2_3", SafeVal(obs45_2(2)))

                    cmd.Parameters.AddWithValue("@obs47_8_1", SafeVal(obs47_8(0)))
                    cmd.Parameters.AddWithValue("@obs47_8_2", SafeVal(obs47_8(1)))
                    cmd.Parameters.AddWithValue("@obs47_8_3", SafeVal(obs47_8(2)))

                    cmd.Parameters.AddWithValue("@obs50_0_1", SafeVal(obs50_0(0)))
                    cmd.Parameters.AddWithValue("@obs50_0_2", SafeVal(obs50_0(1)))
                    cmd.Parameters.AddWithValue("@obs50_0_3", SafeVal(obs50_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateInsideMicrometer50 Error: " & ex.Message)
        End Try
        Return False
    End Function


    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from thread_pitch_micrometer_75.
    ''' </summary>
    Public Function GetThreadPitchMicrometer75MasterLimits(ByVal lc As String) As DataRow

        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_75 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer75MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into thread_pitch_micrometer_75.
    ''' </summary>
    Public Function InsertThreadPitchMicrometer75(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs52_5 As Object(), ByVal obs55_1 As Object(), ByVal obs57_7 As Object(), ByVal obs60_3 As Object(), ByVal obs62_9 As Object(), ByVal obs65_0 As Object(), ByVal obs67_6 As Object(), ByVal obs70_2 As Object(), ByVal obs72_8 As Object(), ByVal obs75_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_pitch_micrometer_75 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs52_5_1, Obs52_5_2, Obs52_5_3, Obs55_1_1, Obs55_1_2, Obs55_1_3, Obs57_7_1, Obs57_7_2, Obs57_7_3, " &
                                   "Obs60_3_1, Obs60_3_2, Obs60_3_3, Obs62_9_1, Obs62_9_2, Obs62_9_3, Obs65_0_1, Obs65_0_2, Obs65_0_3, " &
                                   "Obs67_6_1, Obs67_6_2, Obs67_6_3, Obs70_2_1, Obs70_2_2, Obs70_2_3, Obs72_8_1, Obs72_8_2, Obs72_8_3, " &
                                   "Obs75_0_1, Obs75_0_2, Obs75_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs52_5_1, @obs52_5_2, @obs52_5_3, @obs55_1_1, @obs55_1_2, @obs55_1_3, @obs57_7_1, @obs57_7_2, @obs57_7_3, " &
                                   "@obs60_3_1, @obs60_3_2, @obs60_3_3, @obs62_9_1, @obs62_9_2, @obs62_9_3, @obs65_0_1, @obs65_0_2, @obs65_0_3, " &
                                   "@obs67_6_1, @obs67_6_2, @obs67_6_3, @obs70_2_1, @obs70_2_2, @obs70_2_3, @obs72_8_1, @obs72_8_2, @obs72_8_3, " &
                                   "@obs75_0_1, @obs75_0_2, @obs75_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs52_5_1", SafeVal(obs52_5(0)))
                    cmd.Parameters.AddWithValue("@obs52_5_2", SafeVal(obs52_5(1)))
                    cmd.Parameters.AddWithValue("@obs52_5_3", SafeVal(obs52_5(2)))

                    cmd.Parameters.AddWithValue("@obs55_1_1", SafeVal(obs55_1(0)))
                    cmd.Parameters.AddWithValue("@obs55_1_2", SafeVal(obs55_1(1)))
                    cmd.Parameters.AddWithValue("@obs55_1_3", SafeVal(obs55_1(2)))

                    cmd.Parameters.AddWithValue("@obs57_7_1", SafeVal(obs57_7(0)))
                    cmd.Parameters.AddWithValue("@obs57_7_2", SafeVal(obs57_7(1)))
                    cmd.Parameters.AddWithValue("@obs57_7_3", SafeVal(obs57_7(2)))

                    cmd.Parameters.AddWithValue("@obs60_3_1", SafeVal(obs60_3(0)))
                    cmd.Parameters.AddWithValue("@obs60_3_2", SafeVal(obs60_3(1)))
                    cmd.Parameters.AddWithValue("@obs60_3_3", SafeVal(obs60_3(2)))

                    cmd.Parameters.AddWithValue("@obs62_9_1", SafeVal(obs62_9(0)))
                    cmd.Parameters.AddWithValue("@obs62_9_2", SafeVal(obs62_9(1)))
                    cmd.Parameters.AddWithValue("@obs62_9_3", SafeVal(obs62_9(2)))

                    cmd.Parameters.AddWithValue("@obs65_0_1", SafeVal(obs65_0(0)))
                    cmd.Parameters.AddWithValue("@obs65_0_2", SafeVal(obs65_0(1)))
                    cmd.Parameters.AddWithValue("@obs65_0_3", SafeVal(obs65_0(2)))

                    cmd.Parameters.AddWithValue("@obs67_6_1", SafeVal(obs67_6(0)))
                    cmd.Parameters.AddWithValue("@obs67_6_2", SafeVal(obs67_6(1)))
                    cmd.Parameters.AddWithValue("@obs67_6_3", SafeVal(obs67_6(2)))

                    cmd.Parameters.AddWithValue("@obs70_2_1", SafeVal(obs70_2(0)))
                    cmd.Parameters.AddWithValue("@obs70_2_2", SafeVal(obs70_2(1)))
                    cmd.Parameters.AddWithValue("@obs70_2_3", SafeVal(obs70_2(2)))

                    cmd.Parameters.AddWithValue("@obs72_8_1", SafeVal(obs72_8(0)))
                    cmd.Parameters.AddWithValue("@obs72_8_2", SafeVal(obs72_8(1)))
                    cmd.Parameters.AddWithValue("@obs72_8_3", SafeVal(obs72_8(2)))

                    cmd.Parameters.AddWithValue("@obs75_0_1", SafeVal(obs75_0(0)))
                    cmd.Parameters.AddWithValue("@obs75_0_2", SafeVal(obs75_0(1)))
                    cmd.Parameters.AddWithValue("@obs75_0_3", SafeVal(obs75_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertThreadPitchMicrometer75 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from thread_pitch_micrometer_75.
    ''' </summary>
    Public Function GetThreadPitchMicrometer75Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_75 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer75Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in thread_pitch_micrometer_75.
    ''' </summary>
    Public Function UpdateThreadPitchMicrometer75(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs52_5 As Object(), ByVal obs55_1 As Object(), ByVal obs57_7 As Object(), ByVal obs60_3 As Object(), ByVal obs62_9 As Object(), ByVal obs65_0 As Object(), ByVal obs67_6 As Object(), ByVal obs70_2 As Object(), ByVal obs72_8 As Object(), ByVal obs75_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE thread_pitch_micrometer_75 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs52_5_1=@obs52_5_1, Obs52_5_2=@obs52_5_2, Obs52_5_3=@obs52_5_3, Obs55_1_1=@obs55_1_1, Obs55_1_2=@obs55_1_2, Obs55_1_3=@obs55_1_3, Obs57_7_1=@obs57_7_1, Obs57_7_2=@obs57_7_2, Obs57_7_3=@obs57_7_3, " &
                                   "Obs60_3_1=@obs60_3_1, Obs60_3_2=@obs60_3_2, Obs60_3_3=@obs60_3_3, Obs62_9_1=@obs62_9_1, Obs62_9_2=@obs62_9_2, Obs62_9_3=@obs62_9_3, Obs65_0_1=@obs65_0_1, Obs65_0_2=@obs65_0_2, Obs65_0_3=@obs65_0_3, " &
                                   "Obs67_6_1=@obs67_6_1, Obs67_6_2=@obs67_6_2, Obs67_6_3=@obs67_6_3, Obs70_2_1=@obs70_2_1, Obs70_2_2=@obs70_2_2, Obs70_2_3=@obs70_2_3, Obs72_8_1=@obs72_8_1, Obs72_8_2=@obs72_8_2, Obs72_8_3=@obs72_8_3, " &
                                   "Obs75_0_1=@obs75_0_1, Obs75_0_2=@obs75_0_2, Obs75_0_3=@obs75_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs52_5_1", SafeVal(obs52_5(0)))
                    cmd.Parameters.AddWithValue("@obs52_5_2", SafeVal(obs52_5(1)))
                    cmd.Parameters.AddWithValue("@obs52_5_3", SafeVal(obs52_5(2)))

                    cmd.Parameters.AddWithValue("@obs55_1_1", SafeVal(obs55_1(0)))
                    cmd.Parameters.AddWithValue("@obs55_1_2", SafeVal(obs55_1(1)))
                    cmd.Parameters.AddWithValue("@obs55_1_3", SafeVal(obs55_1(2)))

                    cmd.Parameters.AddWithValue("@obs57_7_1", SafeVal(obs57_7(0)))
                    cmd.Parameters.AddWithValue("@obs57_7_2", SafeVal(obs57_7(1)))
                    cmd.Parameters.AddWithValue("@obs57_7_3", SafeVal(obs57_7(2)))

                    cmd.Parameters.AddWithValue("@obs60_3_1", SafeVal(obs60_3(0)))
                    cmd.Parameters.AddWithValue("@obs60_3_2", SafeVal(obs60_3(1)))
                    cmd.Parameters.AddWithValue("@obs60_3_3", SafeVal(obs60_3(2)))

                    cmd.Parameters.AddWithValue("@obs62_9_1", SafeVal(obs62_9(0)))
                    cmd.Parameters.AddWithValue("@obs62_9_2", SafeVal(obs62_9(1)))
                    cmd.Parameters.AddWithValue("@obs62_9_3", SafeVal(obs62_9(2)))

                    cmd.Parameters.AddWithValue("@obs65_0_1", SafeVal(obs65_0(0)))
                    cmd.Parameters.AddWithValue("@obs65_0_2", SafeVal(obs65_0(1)))
                    cmd.Parameters.AddWithValue("@obs65_0_3", SafeVal(obs65_0(2)))

                    cmd.Parameters.AddWithValue("@obs67_6_1", SafeVal(obs67_6(0)))
                    cmd.Parameters.AddWithValue("@obs67_6_2", SafeVal(obs67_6(1)))
                    cmd.Parameters.AddWithValue("@obs67_6_3", SafeVal(obs67_6(2)))

                    cmd.Parameters.AddWithValue("@obs70_2_1", SafeVal(obs70_2(0)))
                    cmd.Parameters.AddWithValue("@obs70_2_2", SafeVal(obs70_2(1)))
                    cmd.Parameters.AddWithValue("@obs70_2_3", SafeVal(obs70_2(2)))

                    cmd.Parameters.AddWithValue("@obs72_8_1", SafeVal(obs72_8(0)))
                    cmd.Parameters.AddWithValue("@obs72_8_2", SafeVal(obs72_8(1)))
                    cmd.Parameters.AddWithValue("@obs72_8_3", SafeVal(obs72_8(2)))

                    cmd.Parameters.AddWithValue("@obs75_0_1", SafeVal(obs75_0(0)))
                    cmd.Parameters.AddWithValue("@obs75_0_2", SafeVal(obs75_0(1)))
                    cmd.Parameters.AddWithValue("@obs75_0_3", SafeVal(obs75_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateThreadPitchMicrometer75 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from thread_pitch_micrometer_100.
    ''' </summary>
    Public Function GetThreadPitchMicrometer100MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_100 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer100MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into thread_pitch_micrometer_100.
    ''' </summary>
    Public Function InsertThreadPitchMicrometer100(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs77_5 As Object(), ByVal obs80_1 As Object(), ByVal obs82_7 As Object(), ByVal obs85_3 As Object(), ByVal obs87_9 As Object(), ByVal obs90_0 As Object(), ByVal obs92_6 As Object(), ByVal obs95_2 As Object(), ByVal obs97_8 As Object(), ByVal obs100_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_pitch_micrometer_100 " &
                                   "(RowType, ControlNo, CycleName, `Date`, `Time`, `Type`, Size, LC, Color, Location, Temperature, Humidity, TMU, " &
                                   "Obs77_5_1, Obs77_5_2, Obs77_5_3, Obs80_1_1, Obs80_1_2, Obs80_1_3, Obs82_7_1, Obs82_7_2, Obs82_7_3, " &
                                   "Obs85_3_1, Obs85_3_2, Obs85_3_3, Obs87_9_1, Obs87_9_2, Obs87_9_3, Obs90_0_1, Obs90_0_2, Obs90_0_3, " &
                                   "Obs92_6_1, Obs92_6_2, Obs92_6_3, Obs95_2_1, Obs95_2_2, Obs95_2_3, Obs97_8_1, Obs97_8_2, Obs97_8_3, " &
                                   "Obs100_0_1, Obs100_0_2, Obs100_0_3, " &
                                   "DepthError, " &
                                   "TimeIn, TimeOut, TotalTime, Status, Remark) " &
                                   "VALUES " &
                                   "('RECORD', @controlNo, @cycleName, @date, @time, @type, @size, @lc, @color, @location, @temp, @humidity, @tmu, " &
                                   "@obs77_5_1, @obs77_5_2, @obs77_5_3, @obs80_1_1, @obs80_1_2, @obs80_1_3, @obs82_7_1, @obs82_7_2, @obs82_7_3, " &
                                   "@obs85_3_1, @obs85_3_2, @obs85_3_3, @obs87_9_1, @obs87_9_2, @obs87_9_3, @obs90_0_1, @obs90_0_2, @obs90_0_3, " &
                                   "@obs92_6_1, @obs92_6_2, @obs92_6_3, @obs95_2_1, @obs95_2_2, @obs95_2_3, @obs97_8_1, @obs97_8_2, @obs97_8_3, " &
                                   "@obs100_0_1, @obs100_0_2, @obs100_0_3, @depthError, @timeIn, @timeOut, @totalTime, @status, @remark)"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs77_5_1", SafeVal(obs77_5(0)))
                    cmd.Parameters.AddWithValue("@obs77_5_2", SafeVal(obs77_5(1)))
                    cmd.Parameters.AddWithValue("@obs77_5_3", SafeVal(obs77_5(2)))

                    cmd.Parameters.AddWithValue("@obs80_1_1", SafeVal(obs80_1(0)))
                    cmd.Parameters.AddWithValue("@obs80_1_2", SafeVal(obs80_1(1)))
                    cmd.Parameters.AddWithValue("@obs80_1_3", SafeVal(obs80_1(2)))

                    cmd.Parameters.AddWithValue("@obs82_7_1", SafeVal(obs82_7(0)))
                    cmd.Parameters.AddWithValue("@obs82_7_2", SafeVal(obs82_7(1)))
                    cmd.Parameters.AddWithValue("@obs82_7_3", SafeVal(obs82_7(2)))

                    cmd.Parameters.AddWithValue("@obs85_3_1", SafeVal(obs85_3(0)))
                    cmd.Parameters.AddWithValue("@obs85_3_2", SafeVal(obs85_3(1)))
                    cmd.Parameters.AddWithValue("@obs85_3_3", SafeVal(obs85_3(2)))

                    cmd.Parameters.AddWithValue("@obs87_9_1", SafeVal(obs87_9(0)))
                    cmd.Parameters.AddWithValue("@obs87_9_2", SafeVal(obs87_9(1)))
                    cmd.Parameters.AddWithValue("@obs87_9_3", SafeVal(obs87_9(2)))

                    cmd.Parameters.AddWithValue("@obs90_0_1", SafeVal(obs90_0(0)))
                    cmd.Parameters.AddWithValue("@obs90_0_2", SafeVal(obs90_0(1)))
                    cmd.Parameters.AddWithValue("@obs90_0_3", SafeVal(obs90_0(2)))

                    cmd.Parameters.AddWithValue("@obs92_6_1", SafeVal(obs92_6(0)))
                    cmd.Parameters.AddWithValue("@obs92_6_2", SafeVal(obs92_6(1)))
                    cmd.Parameters.AddWithValue("@obs92_6_3", SafeVal(obs92_6(2)))

                    cmd.Parameters.AddWithValue("@obs95_2_1", SafeVal(obs95_2(0)))
                    cmd.Parameters.AddWithValue("@obs95_2_2", SafeVal(obs95_2(1)))
                    cmd.Parameters.AddWithValue("@obs95_2_3", SafeVal(obs95_2(2)))

                    cmd.Parameters.AddWithValue("@obs97_8_1", SafeVal(obs97_8(0)))
                    cmd.Parameters.AddWithValue("@obs97_8_2", SafeVal(obs97_8(1)))
                    cmd.Parameters.AddWithValue("@obs97_8_3", SafeVal(obs97_8(2)))

                    cmd.Parameters.AddWithValue("@obs100_0_1", SafeVal(obs100_0(0)))
                    cmd.Parameters.AddWithValue("@obs100_0_2", SafeVal(obs100_0(1)))
                    cmd.Parameters.AddWithValue("@obs100_0_3", SafeVal(obs100_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertThreadPitchMicrometer100 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Fetches the calibration data for a specific control number and cycle from thread_pitch_micrometer_100.
    ''' </summary>
    Public Function GetThreadPitchMicrometer100Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM thread_pitch_micrometer_100 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetThreadPitchMicrometer100Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Updates an existing calibration record in thread_pitch_micrometer_100.
    ''' </summary>
    Public Function UpdateThreadPitchMicrometer100(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal obs77_5 As Object(), ByVal obs80_1 As Object(), ByVal obs82_7 As Object(), ByVal obs85_3 As Object(), ByVal obs87_9 As Object(), ByVal obs90_0 As Object(), ByVal obs92_6 As Object(), ByVal obs95_2 As Object(), ByVal obs97_8 As Object(), ByVal obs100_0 As Object(),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE thread_pitch_micrometer_100 SET " &
                                   "`Date`=@date, `Time`=@time, `Type`=@type, Size=@size, LC=@lc, Color=@color, Location=@location, Temperature=@temp, Humidity=@humidity, TMU=@tmu, " &
                                   "Obs77_5_1=@obs77_5_1, Obs77_5_2=@obs77_5_2, Obs77_5_3=@obs77_5_3, Obs80_1_1=@obs80_1_1, Obs80_1_2=@obs80_1_2, Obs80_1_3=@obs80_1_3, Obs82_7_1=@obs82_7_1, Obs82_7_2=@obs82_7_2, Obs82_7_3=@obs82_7_3, " &
                                   "Obs85_3_1=@obs85_3_1, Obs85_3_2=@obs85_3_2, Obs85_3_3=@obs85_3_3, Obs87_9_1=@obs87_9_1, Obs87_9_2=@obs87_9_2, Obs87_9_3=@obs87_9_3, Obs90_0_1=@obs90_0_1, Obs90_0_2=@obs90_0_2, Obs90_0_3=@obs90_0_3, " &
                                   "Obs92_6_1=@obs92_6_1, Obs92_6_2=@obs92_6_2, Obs92_6_3=@obs92_6_3, Obs95_2_1=@obs95_2_1, Obs95_2_2=@obs95_2_2, Obs95_2_3=@obs95_2_3, Obs97_8_1=@obs97_8_1, Obs97_8_2=@obs97_8_2, Obs97_8_3=@obs97_8_3, " &
                                   "Obs100_0_1=@obs100_0_1, Obs100_0_2=@obs100_0_2, Obs100_0_3=@obs100_0_3, " &
                                   "DepthError=@depthError, " &
                                   "TimeIn=@timeIn, TimeOut=@timeOut, TotalTime=@totalTime, Status=@status, Remark=@remark " &
                                   "WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val), DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@obs77_5_1", SafeVal(obs77_5(0)))
                    cmd.Parameters.AddWithValue("@obs77_5_2", SafeVal(obs77_5(1)))
                    cmd.Parameters.AddWithValue("@obs77_5_3", SafeVal(obs77_5(2)))

                    cmd.Parameters.AddWithValue("@obs80_1_1", SafeVal(obs80_1(0)))
                    cmd.Parameters.AddWithValue("@obs80_1_2", SafeVal(obs80_1(1)))
                    cmd.Parameters.AddWithValue("@obs80_1_3", SafeVal(obs80_1(2)))

                    cmd.Parameters.AddWithValue("@obs82_7_1", SafeVal(obs82_7(0)))
                    cmd.Parameters.AddWithValue("@obs82_7_2", SafeVal(obs82_7(1)))
                    cmd.Parameters.AddWithValue("@obs82_7_3", SafeVal(obs82_7(2)))

                    cmd.Parameters.AddWithValue("@obs85_3_1", SafeVal(obs85_3(0)))
                    cmd.Parameters.AddWithValue("@obs85_3_2", SafeVal(obs85_3(1)))
                    cmd.Parameters.AddWithValue("@obs85_3_3", SafeVal(obs85_3(2)))

                    cmd.Parameters.AddWithValue("@obs87_9_1", SafeVal(obs87_9(0)))
                    cmd.Parameters.AddWithValue("@obs87_9_2", SafeVal(obs87_9(1)))
                    cmd.Parameters.AddWithValue("@obs87_9_3", SafeVal(obs87_9(2)))

                    cmd.Parameters.AddWithValue("@obs90_0_1", SafeVal(obs90_0(0)))
                    cmd.Parameters.AddWithValue("@obs90_0_2", SafeVal(obs90_0(1)))
                    cmd.Parameters.AddWithValue("@obs90_0_3", SafeVal(obs90_0(2)))

                    cmd.Parameters.AddWithValue("@obs92_6_1", SafeVal(obs92_6(0)))
                    cmd.Parameters.AddWithValue("@obs92_6_2", SafeVal(obs92_6(1)))
                    cmd.Parameters.AddWithValue("@obs92_6_3", SafeVal(obs92_6(2)))

                    cmd.Parameters.AddWithValue("@obs95_2_1", SafeVal(obs95_2(0)))
                    cmd.Parameters.AddWithValue("@obs95_2_2", SafeVal(obs95_2(1)))
                    cmd.Parameters.AddWithValue("@obs95_2_3", SafeVal(obs95_2(2)))

                    cmd.Parameters.AddWithValue("@obs97_8_1", SafeVal(obs97_8(0)))
                    cmd.Parameters.AddWithValue("@obs97_8_2", SafeVal(obs97_8(1)))
                    cmd.Parameters.AddWithValue("@obs97_8_3", SafeVal(obs97_8(2)))

                    cmd.Parameters.AddWithValue("@obs100_0_1", SafeVal(obs100_0(0)))
                    cmd.Parameters.AddWithValue("@obs100_0_2", SafeVal(obs100_0(1)))
                    cmd.Parameters.AddWithValue("@obs100_0_3", SafeVal(obs100_0(2)))

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateThreadPitchMicrometer100 Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' --- RECORDS TAB HELPERS ---

    Public Function GetScanLogs() As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM tbl_scanlog ORDER BY ScannedAt DESC"
                Using cmd As New MySqlCommand(query, _con)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)
                        Return dt
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetScanLogs Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function GetRegisteredItems() As DataTable
        Try
            ' Build union across all type tables for items with RFID tags
            Dim typeRows = ReadDatatable("SELECT TypeName, Category FROM type_details")
            Dim unionParts As New List(Of String)()
            For Each typeRow As DataRow In typeRows.Rows
                Dim typeName = typeRow("TypeName").ToString()
                Dim category = typeRow("Category").ToString()
                Dim tbl = TypeNameToTableName(typeName)

                If category.Equals("Instrument", StringComparison.OrdinalIgnoreCase) Then
                    unionParts.Add($"SELECT RFID_tag, ControlNo, 'Instrument' AS ItemType, InstrumentName AS ItemName FROM `{tbl}` WHERE RFID_tag IS NOT NULL AND RFID_tag != ''")
                Else
                    unionParts.Add($"SELECT RFID_tag, ControlNo, 'Gauge' AS ItemType, GaugeName AS ItemName FROM `{tbl}` WHERE RFID_tag IS NOT NULL AND RFID_tag != ''")
                End If
            Next
            If unionParts.Count = 0 Then Return New DataTable()
            Dim query As String = $"{String.Join(" UNION ALL ", unionParts)} ORDER BY ControlNo"
            Return ReadDatatable(query)
        Catch ex As Exception
            Console.WriteLine("GetRegisteredItems Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function GetUnRegisteredItems() As DataTable
        Try
            ' Build union across all type tables for items without RFID tags
            Dim typeRows = ReadDatatable("SELECT TypeName, Category FROM type_details")
            Dim unionParts As New List(Of String)()
            For Each typeRow As DataRow In typeRows.Rows
                Dim typeName = typeRow("TypeName").ToString()
                Dim category = typeRow("Category").ToString()
                Dim tbl = TypeNameToTableName(typeName)

                If category.Equals("Instrument", StringComparison.OrdinalIgnoreCase) Then
                    unionParts.Add($"SELECT ControlNo, 'Instrument' AS ItemType, InstrumentName AS ItemName FROM `{tbl}` WHERE RFID_tag IS NULL OR RFID_tag = ''")
                Else
                    unionParts.Add($"SELECT ControlNo, 'Gauge' AS ItemType, GaugeName AS ItemName FROM `{tbl}` WHERE RFID_tag IS NULL OR RFID_tag = ''")
                End If
            Next
            If unionParts.Count = 0 Then Return New DataTable()
            Dim query As String = $"{String.Join(" UNION ALL ", unionParts)} ORDER BY ControlNo"
            Return ReadDatatable(query)
        Catch ex As Exception
            Console.WriteLine("GetUnRegisteredItems Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    ''' <summary>
    ''' Fetches the MASTER tolerance limits for a specific Least Count (LC) from external_micrometer_100.
    ''' </summary>
    Public Function GetExternalMicrometer100MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM external_micrometer_100 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetExternalMicrometer100MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Inserts a new calibration record into external_micrometer_100.
    ''' </summary>
    Public Function InsertExternalMicrometer100(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal anvilFlatness As Object(), ByVal spindleFlatness As Object(), ByVal parallelism As Object(),
                                          ByVal obsArr As Dictionary(Of Decimal, Object()),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim columns As New List(Of String) From {
                    "RowType", "ControlNo", "CycleName", "`Date`", "`Time`", "`Type`", "Size", "LC", "Color", "Location", "Temperature", "Humidity", "TMU",
                    "Flatness_Anvil_1", "Flatness_Anvil_2", "Flatness_Anvil_3",
                    "Flatness_Spindle_1", "Flatness_Spindle_2", "Flatness_Spindle_3",
                    "Parallelism_1", "Parallelism_2", "Parallelism_3",
                    "DepthError", "TimeIn", "TimeOut", "TotalTime", "Status", "Remark"
                }

                Dim params As New List(Of String) From {
                    "'RECORD'", "@controlNo", "@cycleName", "@date", "@time", "@type", "@size", "@lc", "@color", "@location", "@temp", "@humidity", "@tmu",
                    "@anvil_1", "@anvil_2", "@anvil_3",
                    "@spindle_1", "@spindle_2", "@spindle_3",
                    "@parallel_1", "@parallel_2", "@parallel_3",
                    "@depthError", "@timeIn", "@timeOut", "@totalTime", "@status", "@remark"
                }

                For Each kvp In obsArr
                    Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                    columns.Add($"Obs{nomStr}_1")
                    columns.Add($"Obs{nomStr}_2")
                    columns.Add($"Obs{nomStr}_3")
                    params.Add($"@obs{nomStr}_1")
                    params.Add($"@obs{nomStr}_2")
                    params.Add($"@obs{nomStr}_3")
                Next

                Dim query As String = $"INSERT INTO external_micrometer_100 ({String.Join(", ", columns)}) VALUES ({String.Join(", ", params)})"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@anvil_1", SafeVal(anvilFlatness(0)))
                    cmd.Parameters.AddWithValue("@anvil_2", SafeVal(anvilFlatness(1)))
                    cmd.Parameters.AddWithValue("@anvil_3", SafeVal(anvilFlatness(2)))
                    cmd.Parameters.AddWithValue("@spindle_1", SafeVal(spindleFlatness(0)))
                    cmd.Parameters.AddWithValue("@spindle_2", SafeVal(spindleFlatness(1)))
                    cmd.Parameters.AddWithValue("@spindle_3", SafeVal(spindleFlatness(2)))
                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))

                    For Each kvp In obsArr
                        Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertExternalMicrometer100 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetExternalMicrometer100Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM external_micrometer_100 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetExternalMicrometer100Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function UpdateExternalMicrometer100(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal anvilFlatness As Object(), ByVal spindleFlatness As Object(), ByVal parallelism As Object(),
                                           ByVal obsArr As Dictionary(Of Decimal, Object()),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sets As New List(Of String) From {
                    "`Date`=@date", "`Time`=@time", "`Type`=@type", "Size=@size", "LC=@lc", "Color=@color", "Location=@location", "Temperature=@temp", "Humidity=@humidity", "TMU=@tmu",
                    "Flatness_Anvil_1=@anvil_1", "Flatness_Anvil_2=@anvil_2", "Flatness_Anvil_3=@anvil_3",
                    "Flatness_Spindle_1=@spindle_1", "Flatness_Spindle_2=@spindle_2", "Flatness_Spindle_3=@spindle_3",
                    "Parallelism_1=@parallel_1", "Parallelism_2=@parallel_2", "Parallelism_3=@parallel_3",
                    "DepthError=@depthError", "TimeIn=@timeIn", "TimeOut=@timeOut", "TotalTime=@totalTime", "Status=@status", "Remark=@remark"
                }

                For Each kvp In obsArr
                    Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                    sets.Add($"Obs{nomStr}_1=@obs{nomStr}_1")
                    sets.Add($"Obs{nomStr}_2=@obs{nomStr}_2")
                    sets.Add($"Obs{nomStr}_3=@obs{nomStr}_3")
                Next

                Dim query As String = $"UPDATE external_micrometer_100 SET {String.Join(", ", sets)} WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@anvil_1", SafeVal(anvilFlatness(0)))
                    cmd.Parameters.AddWithValue("@anvil_2", SafeVal(anvilFlatness(1)))
                    cmd.Parameters.AddWithValue("@anvil_3", SafeVal(anvilFlatness(2)))
                    cmd.Parameters.AddWithValue("@spindle_1", SafeVal(spindleFlatness(0)))
                    cmd.Parameters.AddWithValue("@spindle_2", SafeVal(spindleFlatness(1)))
                    cmd.Parameters.AddWithValue("@spindle_3", SafeVal(spindleFlatness(2)))
                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))

                    For Each kvp In obsArr
                        Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateExternalMicrometer100 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetSizesForGroup(prefix As String, category As String) As List(Of String)
        Dim sizes As New List(Of String)()
        Dim table As String = If(category?.ToLower() = "gauge", "gauge_categorycontrol", "categorycontrol")
        Try
            Dim dt = ReadDatatable($"SELECT DISTINCT Size FROM `{table}` WHERE GroupCode = '{prefix.Replace("'", "''")}' AND (Active = 1 OR Active IS NULL) ORDER BY ID ASC")
            For Each row As DataRow In dt.Rows
                sizes.Add(row("Size").ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("GetSizesForGroup Error: " & ex.Message)
        End Try
        Return sizes
    End Function

    Public Shared Function ExtractPrefix(controlNo As String) As String
        If String.IsNullOrWhiteSpace(controlNo) Then Return ""
        Dim parts = controlNo.Split("-"c)
        Return parts(0)
    End Function

    Public Function GetInterchangeInfoByTag(tag As String) As DataTable
        Try
            Dim cleanTag = tag.Replace(" ", "").ToUpper().Replace("'", "''")
            ' Get the most recent interchange info for this tag
            Dim query = "SELECT ControlNo, InstrumentName FROM interchangeability " &
                        $"WHERE UPPER(REPLACE(RFID_tag, ' ', '')) = '{cleanTag}' " &
                        "ORDER BY ActionDate DESC, ActionTime DESC LIMIT 1"
            Return ReadDatatable(query)
        Catch ex As Exception
            Console.WriteLine("GetInterchangeInfoByTag Error: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function InsertExclusion(tag As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                ' Check if already exists to avoid duplicates
                If IsTagExcluded(tag) Then Return True

                Dim query = "INSERT INTO rfid_exclusions (rfid_tag, added_at) VALUES (@tag, NOW())"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@tag", tag.Trim())
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("InsertExclusion Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function IsTagExcluded(tag As String) As Boolean
        Try
            Dim cleanTag = tag.Replace(" ", "").ToUpper().Replace("'", "''")
            Dim query = $"SELECT COUNT(*) FROM rfid_exclusions WHERE UPPER(REPLACE(rfid_tag, ' ', '')) = '{cleanTag}'"
            Dim dt = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return Convert.ToInt32(dt.Rows(0)(0)) > 0
            End If
        Catch ex As Exception
            Console.WriteLine("IsTagExcluded Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetExclusionList() As DataTable
        Try
            Dim query = "SELECT rfid_tag, added_at FROM rfid_exclusions ORDER BY added_at DESC"
            Return ReadDatatable(query)
        Catch ex As Exception
            Console.WriteLine("GetExclusionList Error: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function DeleteExclusion(tagId As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query = "DELETE FROM rfid_exclusions WHERE rfid_tag = @tag"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@tag", tagId)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("DeleteExclusion Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetDiscMicrometer75MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM disc_micrometer_75 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetDiscMicrometer75MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function InsertDiscMicrometer75(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal anvilFlatness As Object(), ByVal spindleFlatness As Object(), ByVal parallelism As Object(),
                                          ByVal obsArr As Dictionary(Of Decimal, Object()),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim columns As New List(Of String) From {
                    "RowType", "ControlNo", "CycleName", "`Date`", "`Time`", "`Type`", "Size", "LC", "Color", "Location", "Temperature", "Humidity", "TMU",
                    "Flatness_Anvil_1", "Flatness_Anvil_2", "Flatness_Anvil_3",
                    "Flatness_Spindle_1", "Flatness_Spindle_2", "Flatness_Spindle_3",
                    "Parallelism_1", "Parallelism_2", "Parallelism_3",
                    "DepthError", "TimeIn", "TimeOut", "TotalTime", "Status", "Remark"
                }

                Dim params As New List(Of String) From {
                    "'RECORD'", "@controlNo", "@cycleName", "@date", "@time", "@type", "@size", "@lc", "@color", "@location", "@temp", "@humidity", "@tmu",
                    "@anvil_1", "@anvil_2", "@anvil_3",
                    "@spindle_1", "@spindle_2", "@spindle_3",
                    "@parallel_1", "@parallel_2", "@parallel_3",
                    "@depthError", "@timeIn", "@timeOut", "@totalTime", "@status", "@remark"
                }

                For Each kvp In obsArr
                    Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                    columns.Add($"Obs{nomStr}_1")
                    columns.Add($"Obs{nomStr}_2")
                    columns.Add($"Obs{nomStr}_3")
                    params.Add($"@obs{nomStr}_1")
                    params.Add($"@obs{nomStr}_2")
                    params.Add($"@obs{nomStr}_3")
                Next

                Dim query As String = $"INSERT INTO disc_micrometer_75 ({String.Join(", ", columns)}) VALUES ({String.Join(", ", params)})"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@anvil_1", SafeVal(anvilFlatness(0)))
                    cmd.Parameters.AddWithValue("@anvil_2", SafeVal(anvilFlatness(1)))
                    cmd.Parameters.AddWithValue("@anvil_3", SafeVal(anvilFlatness(2)))
                    cmd.Parameters.AddWithValue("@spindle_1", SafeVal(spindleFlatness(0)))
                    cmd.Parameters.AddWithValue("@spindle_2", SafeVal(spindleFlatness(1)))
                    cmd.Parameters.AddWithValue("@spindle_3", SafeVal(spindleFlatness(2)))
                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))

                    For Each kvp In obsArr
                        Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertDiscMicrometer75 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetDiscMicrometer75Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM disc_micrometer_75 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetDiscMicrometer75Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function UpdateDiscMicrometer75(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal anvilFlatness As Object(), ByVal spindleFlatness As Object(), ByVal parallelism As Object(),
                                           ByVal obsArr As Dictionary(Of Decimal, Object()),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sets As New List(Of String) From {
                    "`Date`=@date", "`Time`=@time", "`Type`=@type", "Size=@size", "LC=@lc", "Color=@color", "Location=@location", "Temperature=@temp", "Humidity=@humidity", "TMU=@tmu",
                    "Flatness_Anvil_1=@anvil_1", "Flatness_Anvil_2=@anvil_2", "Flatness_Anvil_3=@anvil_3",
                    "Flatness_Spindle_1=@spindle_1", "Flatness_Spindle_2=@spindle_2", "Flatness_Spindle_3=@spindle_3",
                    "Parallelism_1=@parallel_1", "Parallelism_2=@parallel_2", "Parallelism_3=@parallel_3",
                    "DepthError=@depthError", "TimeIn=@timeIn", "TimeOut=@timeOut", "TotalTime=@totalTime", "Status=@status", "Remark=@remark"
                }

                For Each kvp In obsArr
                    Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                    sets.Add($"Obs{nomStr}_1=@obs{nomStr}_1")
                    sets.Add($"Obs{nomStr}_2=@obs{nomStr}_2")
                    sets.Add($"Obs{nomStr}_3=@obs{nomStr}_3")
                Next

                Dim query As String = $"UPDATE disc_micrometer_75 SET {String.Join(", ", sets)} WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@anvil_1", SafeVal(anvilFlatness(0)))
                    cmd.Parameters.AddWithValue("@anvil_2", SafeVal(anvilFlatness(1)))
                    cmd.Parameters.AddWithValue("@anvil_3", SafeVal(anvilFlatness(2)))
                    cmd.Parameters.AddWithValue("@spindle_1", SafeVal(spindleFlatness(0)))
                    cmd.Parameters.AddWithValue("@spindle_2", SafeVal(spindleFlatness(1)))
                    cmd.Parameters.AddWithValue("@spindle_3", SafeVal(spindleFlatness(2)))
                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))

                    For Each kvp In obsArr
                        Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateDiscMicrometer75 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetDepthMicrometer150MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM depth_micrometer_150 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetDepthMicrometer150MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function InsertDepthMicrometer150(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                          ByVal flatnessMeas As Object(), ByVal flatnessRef As Object(), ByVal parallelism As Object(), ByVal zeroError As Object(),
                                          ByVal obsArr As Dictionary(Of Decimal, Object()),
                                          ByVal depthError As Object,
                                          ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                          Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim columns As New List(Of String) From {
                    "RowType", "ControlNo", "CycleName", "`Date`", "`Time`", "`Type`", "Size", "LC", "Color", "Location", "Temperature", "Humidity", "TMU",
                    "Flatness_Meas_1", "Flatness_Meas_2", "Flatness_Meas_3",
                    "Flatness_Ref_1", "Flatness_Ref_2", "Flatness_Ref_3",
                    "Parallelism_1", "Parallelism_2", "Parallelism_3",
                    "Zero_Error_1", "Zero_Error_2", "Zero_Error_3",
                    "DepthError", "TimeIn", "TimeOut", "TotalTime", "Status", "Remark"
                }

                Dim params As New List(Of String) From {
                    "'RECORD'", "@controlNo", "@cycleName", "@date", "@time", "@type", "@size", "@lc", "@color", "@location", "@temp", "@humidity", "@tmu",
                    "@meas_1", "@meas_2", "@meas_3",
                    "@ref_1", "@ref_2", "@ref_3",
                    "@parallel_1", "@parallel_2", "@parallel_3",
                    "@zero_1", "@zero_2", "@zero_3",
                    "@depthError", "@timeIn", "@timeOut", "@totalTime", "@status", "@remark"
                }

                For Each kvp In obsArr
                    Dim nomDbStr = kvp.Key.ToString("0")
                    columns.Add($"Obs{nomDbStr}_1")
                    columns.Add($"Obs{nomDbStr}_2")
                    columns.Add($"Obs{nomDbStr}_3")
                    params.Add($"@obs{nomDbStr}_1")
                    params.Add($"@obs{nomDbStr}_2")
                    params.Add($"@obs{nomDbStr}_3")
                Next

                Dim query As String = $"INSERT INTO depth_micrometer_150 ({String.Join(", ", columns)}) VALUES ({String.Join(", ", params)})"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@meas_1", SafeVal(flatnessMeas(0)))
                    cmd.Parameters.AddWithValue("@meas_2", SafeVal(flatnessMeas(1)))
                    cmd.Parameters.AddWithValue("@meas_3", SafeVal(flatnessMeas(2)))
                    cmd.Parameters.AddWithValue("@ref_1", SafeVal(flatnessRef(0)))
                    cmd.Parameters.AddWithValue("@ref_2", SafeVal(flatnessRef(1)))
                    cmd.Parameters.AddWithValue("@ref_3", SafeVal(flatnessRef(2)))
                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))
                    cmd.Parameters.AddWithValue("@zero_1", SafeVal(zeroError(0)))
                    cmd.Parameters.AddWithValue("@zero_2", SafeVal(zeroError(1)))
                    cmd.Parameters.AddWithValue("@zero_3", SafeVal(zeroError(2)))

                    For Each kvp In obsArr
                        Dim nomDbStr = kvp.Key.ToString("0")
                        cmd.Parameters.AddWithValue($"@obs{nomDbStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomDbStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomDbStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertDepthMicrometer150 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetDepthMicrometer150Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM depth_micrometer_150 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetDepthMicrometer150Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function UpdateDepthMicrometer150(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal flatnessMeas As Object(), ByVal flatnessRef As Object(), ByVal parallelism As Object(), ByVal zeroError As Object(),
                                           ByVal obsArr As Dictionary(Of Decimal, Object()),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sets As New List(Of String) From {
                    "`Date`=@date", "`Time`=@time", "`Type`=@type", "Size=@size", "LC=@lc", "Color=@color", "Location=@location", "Temperature=@temp", "Humidity=@humidity", "TMU=@tmu",
                    "Flatness_Meas_1=@meas_1", "Flatness_Meas_2=@meas_2", "Flatness_Meas_3=@meas_3",
                    "Flatness_Ref_1=@ref_1", "Flatness_Ref_2=@ref_2", "Flatness_Ref_3=@ref_3",
                    "Parallelism_1=@parallel_1", "Parallelism_2=@parallel_2", "Parallelism_3=@parallel_3",
                    "Zero_Error_1=@zero_1", "Zero_Error_2=@zero_2", "Zero_Error_3=@zero_3",
                    "DepthError=@depthError", "TimeIn=@timeIn", "TimeOut=@timeOut", "TotalTime=@totalTime", "Status=@status", "Remark=@remark"
                }

                For Each kvp In obsArr
                    Dim nomDbStr = kvp.Key.ToString("0")
                    sets.Add($"Obs{nomDbStr}_1=@obs{nomDbStr}_1")
                    sets.Add($"Obs{nomDbStr}_2=@obs{nomDbStr}_2")
                    sets.Add($"Obs{nomDbStr}_3=@obs{nomDbStr}_3")
                Next

                Dim query As String = $"UPDATE depth_micrometer_150 SET {String.Join(", ", sets)} WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@meas_1", SafeVal(flatnessMeas(0)))
                    cmd.Parameters.AddWithValue("@meas_2", SafeVal(flatnessMeas(1)))
                    cmd.Parameters.AddWithValue("@meas_3", SafeVal(flatnessMeas(2)))
                    cmd.Parameters.AddWithValue("@ref_1", SafeVal(flatnessRef(0)))
                    cmd.Parameters.AddWithValue("@ref_2", SafeVal(flatnessRef(1)))
                    cmd.Parameters.AddWithValue("@ref_3", SafeVal(flatnessRef(2)))
                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))
                    cmd.Parameters.AddWithValue("@zero_1", SafeVal(zeroError(0)))
                    cmd.Parameters.AddWithValue("@zero_2", SafeVal(zeroError(1)))
                    cmd.Parameters.AddWithValue("@zero_3", SafeVal(zeroError(2)))

                    For Each kvp In obsArr
                        Dim nomDbStr = kvp.Key.ToString("0")
                        cmd.Parameters.AddWithValue($"@obs{nomDbStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomDbStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomDbStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateDepthMicrometer150 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetBladeMicrometer75MasterLimits(ByVal lc As String) As DataRow
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM blade_micrometer_75 WHERE RowType='MASTER' AND LC=@lc"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    Dim dt As New DataTable()
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetBladeMicrometer75MasterLimits Error: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function InsertBladeMicrometer75(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal parallelism As Object(),
                                           ByVal obsArr As Dictionary(Of Decimal, Object()),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim columns As New List(Of String) From {
                    "RowType", "ControlNo", "CycleName", "`Date`", "`Time`", "`Type`", "Size", "LC", "Color", "Location", "Temperature", "Humidity", "TMU",
                    "Parallelism_1", "Parallelism_2", "Parallelism_3",
                    "DepthError", "TimeIn", "TimeOut", "TotalTime", "Status", "Remark"
                }

                Dim params As New List(Of String) From {
                    "'RECORD'", "@controlNo", "@cycleName", "@date", "@time", "@type", "@size", "@lc", "@color", "@location", "@temp", "@humidity", "@tmu",
                    "@parallel_1", "@parallel_2", "@parallel_3",
                    "@depthError", "@timeIn", "@timeOut", "@totalTime", "@status", "@remark"
                }

                For Each kvp In obsArr
                    Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                    columns.Add($"Obs{nomStr}_1")
                    columns.Add($"Obs{nomStr}_2")
                    columns.Add($"Obs{nomStr}_3")
                    params.Add($"@obs{nomStr}_1")
                    params.Add($"@obs{nomStr}_2")
                    params.Add($"@obs{nomStr}_3")
                Next

                Dim query As String = $"INSERT INTO blade_micrometer_75 ({String.Join(", ", columns)}) VALUES ({String.Join(", ", params)})"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))

                    For Each kvp In obsArr
                        Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("InsertBladeMicrometer75 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetBladeMicrometer75Data(controlNo As String, cycleName As String) As DataTable
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "SELECT * FROM blade_micrometer_75 WHERE ControlNo = @ctrl AND CycleName = @cyc AND RowType = 'RECORD' ORDER BY Date DESC, Time DESC LIMIT 1"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dtResult As New DataTable()
                        adapter.Fill(dtResult)
                        Return dtResult
                    End Using
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("GetBladeMicrometer75Data Error: " & ex.Message)
        End Try
        Return New DataTable()
    End Function

    Public Function UpdateBladeMicrometer75(ByVal type As String, ByVal controlNo As String, ByVal cycleName As String, ByVal size As String, ByVal lc As String, ByVal color As String, ByVal location As String, ByVal temp As String, ByVal humidity As String, ByVal tmu As String,
                                           ByVal parallelism As Object(),
                                           ByVal obsArr As Dictionary(Of Decimal, Object()),
                                           ByVal depthError As Object,
                                           ByVal timeIn As String, ByVal timeOut As String, ByVal totalTime As String, ByVal status As String, ByVal remark As String,
                                           Optional ByVal customDate As Date? = Nothing) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim sets As New List(Of String) From {
                    "`Date`=@date", "`Time`=@time", "`Type`=@type", "Size=@size", "LC=@lc", "Color=@color", "Location=@location", "Temperature=@temp", "Humidity=@humidity", "TMU=@tmu",
                    "Parallelism_1=@parallel_1", "Parallelism_2=@parallel_2", "Parallelism_3=@parallel_3",
                    "DepthError=@depthError", "TimeIn=@timeIn", "TimeOut=@timeOut", "TotalTime=@totalTime", "Status=@status", "Remark=@remark"
                }

                For Each kvp In obsArr
                    Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                    sets.Add($"Obs{nomStr}_1=@obs{nomStr}_1")
                    sets.Add($"Obs{nomStr}_2=@obs{nomStr}_2")
                    sets.Add($"Obs{nomStr}_3=@obs{nomStr}_3")
                Next

                Dim query As String = $"UPDATE blade_micrometer_75 SET {String.Join(", ", sets)} WHERE ControlNo=@controlNo AND CycleName=@cycleName AND RowType='RECORD'"

                Using cmd As New MySqlCommand(query, _con)
                    Dim SafeVal As Func(Of Object, Object) = Function(val) If(val Is Nothing OrElse IsDBNull(val) OrElse val.ToString() = "-", DBNull.Value, val)

                    cmd.Parameters.AddWithValue("@controlNo", controlNo)
                    cmd.Parameters.AddWithValue("@cycleName", cycleName)
                    cmd.Parameters.AddWithValue("@date", If(customDate.HasValue, customDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@type", type)
                    cmd.Parameters.AddWithValue("@size", size)
                    cmd.Parameters.AddWithValue("@lc", lc)
                    cmd.Parameters.AddWithValue("@color", color)
                    cmd.Parameters.AddWithValue("@location", location)
                    cmd.Parameters.AddWithValue("@temp", temp)
                    cmd.Parameters.AddWithValue("@humidity", humidity)
                    cmd.Parameters.AddWithValue("@tmu", tmu)

                    cmd.Parameters.AddWithValue("@parallel_1", SafeVal(parallelism(0)))
                    cmd.Parameters.AddWithValue("@parallel_2", SafeVal(parallelism(1)))
                    cmd.Parameters.AddWithValue("@parallel_3", SafeVal(parallelism(2)))

                    For Each kvp In obsArr
                        Dim nomStr = kvp.Key.ToString("0.0").Replace(".", "_")
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_1", SafeVal(kvp.Value(0)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_2", SafeVal(kvp.Value(1)))
                        cmd.Parameters.AddWithValue($"@obs{nomStr}_3", SafeVal(kvp.Value(2)))
                    Next

                    cmd.Parameters.AddWithValue("@depthError", SafeVal(depthError))
                    cmd.Parameters.AddWithValue("@timeIn", timeIn)
                    cmd.Parameters.AddWithValue("@timeOut", timeOut)
                    cmd.Parameters.AddWithValue("@totalTime", totalTime)
                    cmd.Parameters.AddWithValue("@status", status)
                    cmd.Parameters.AddWithValue("@remark", remark)

                    Return (cmd.ExecuteNonQuery() > 0)
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("UpdateBladeMicrometer75 Error: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetMastersByDescriptions(descriptions As List(Of String)) As List(Of MasterSelectorItem)
        Dim result As New List(Of MasterSelectorItem)()
        If descriptions Is Nothing OrElse descriptions.Count = 0 Then Return result

        Try
            Dim descList = String.Join("','", descriptions.Select(Function(d) d.Replace("'", "''")))
            Dim query As String = $"SELECT Description, MasterUncertainty, LeastCount, CalDate, DueDate FROM calibrationmaster_details WHERE Description IN ('{descList}')"

            Dim dt = ReadDatatable(query)
            For Each row As DataRow In dt.Rows
                Dim item As New MasterSelectorItem()
                item.Description = row("Description").ToString()
                item.MasterUncertainty = If(row("MasterUncertainty") Is DBNull.Value, 0, Convert.ToDecimal(row("MasterUncertainty")))
                item.LeastCount = row("LeastCount").ToString()
                item.CalDate = If(row("CalDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("CalDate")))
                item.DueDate = If(row("DueDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("DueDate")))
                item.IsSelected = True
                result.Add(item)
            Next
        Catch ex As Exception
            Console.WriteLine("Error in GetMastersByDescriptions: " & ex.Message)
        End Try
        Return result
    End Function

    ' ========================================    ' Passameter60 Methods
    ' ==========================================
    Public Function GetPassameter60MasterLimits(size As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM passameter_60 WHERE RowType='MASTER' AND Size='{size.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetPassameter60MasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetPassameter60Data(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM passameter_60 WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertPassameter60(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, anvil As Object(), spindle As Object(), parallelism As Object(), plusObs As Dictionary(Of String, Object()), minusObs As Dictionary(Of String, Object()), depthError As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO passameter_60 (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, Flatness_Anvil_1, Flatness_Anvil_2, Flatness_Anvil_3, Flatness_Spindle_1, Flatness_Spindle_2, Flatness_Spindle_3, Parallelism_1, Parallelism_2, Parallelism_3, Plus_10_1, Plus_10_2, Plus_10_3, Plus_20_1, Plus_20_2, Plus_20_3, Plus_30_1, Plus_30_2, Plus_30_3, Plus_40_1, Plus_40_2, Plus_40_3, Plus_50_1, Plus_50_2, Plus_50_3, Plus_60_1, Plus_60_2, Plus_60_3, Minus_10_1, Minus_10_2, Minus_10_3, Minus_20_1, Minus_20_2, Minus_20_3, Minus_30_1, Minus_30_2, Minus_30_3, Minus_40_1, Minus_40_2, Minus_40_3, Minus_50_1, Minus_50_2, Minus_50_3, Minus_60_1, Minus_60_2, Minus_60_3, DepthError, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @Flatness_Anvil_1, @Flatness_Anvil_2, @Flatness_Anvil_3, @Flatness_Spindle_1, @Flatness_Spindle_2, @Flatness_Spindle_3, @Parallelism_1, @Parallelism_2, @Parallelism_3, @Plus_10_1, @Plus_10_2, @Plus_10_3, @Plus_20_1, @Plus_20_2, @Plus_20_3, @Plus_30_1, @Plus_30_2, @Plus_30_3, @Plus_40_1, @Plus_40_2, @Plus_40_3, @Plus_50_1, @Plus_50_2, @Plus_50_3, @Plus_60_1, @Plus_60_2, @Plus_60_3, @Minus_10_1, @Minus_10_2, @Minus_10_3, @Minus_20_1, @Minus_20_2, @Minus_20_3, @Minus_30_1, @Minus_30_2, @Minus_30_3, @Minus_40_1, @Minus_40_2, @Minus_40_3, @Minus_50_1, @Minus_50_2, @Minus_50_3, @Minus_60_1, @Minus_60_2, @Minus_60_3, @DepthError, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Flatness_Anvil_1", anvil(0))
                    cmd.Parameters.AddWithValue("@Flatness_Anvil_2", anvil(1))
                    cmd.Parameters.AddWithValue("@Flatness_Anvil_3", anvil(2))
                    cmd.Parameters.AddWithValue("@Flatness_Spindle_1", spindle(0))
                    cmd.Parameters.AddWithValue("@Flatness_Spindle_2", spindle(1))
                    cmd.Parameters.AddWithValue("@Flatness_Spindle_3", spindle(2))
                    cmd.Parameters.AddWithValue("@Parallelism_1", parallelism(0))
                    cmd.Parameters.AddWithValue("@Parallelism_2", parallelism(1))
                    cmd.Parameters.AddWithValue("@Parallelism_3", parallelism(2))

                    For Each kvp In plusObs
                        cmd.Parameters.AddWithValue($"@Plus_{kvp.Key}_1", kvp.Value(0))
                        cmd.Parameters.AddWithValue($"@Plus_{kvp.Key}_2", kvp.Value(1))
                        cmd.Parameters.AddWithValue($"@Plus_{kvp.Key}_3", kvp.Value(2))
                    Next

                    For Each kvp In minusObs
                        cmd.Parameters.AddWithValue($"@Minus_{kvp.Key}_1", kvp.Value(0))
                        cmd.Parameters.AddWithValue($"@Minus_{kvp.Key}_2", kvp.Value(1))
                        cmd.Parameters.AddWithValue($"@Minus_{kvp.Key}_3", kvp.Value(2))
                    Next

                    cmd.Parameters.AddWithValue("@DepthError", depthError)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertPassameter60: " & ex.Message)
            Return False
        End Try
        Return False
    End Function

    Public Function UpdatePassameter60(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, anvil As Object(), spindle As Object(), parallelism As Object(), plusObs As Dictionary(Of String, Object()), minusObs As Dictionary(Of String, Object()), depthError As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE passameter_60 SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, Flatness_Anvil_1=@Flatness_Anvil_1, Flatness_Anvil_2=@Flatness_Anvil_2, Flatness_Anvil_3=@Flatness_Anvil_3, Flatness_Spindle_1=@Flatness_Spindle_1, Flatness_Spindle_2=@Flatness_Spindle_2, Flatness_Spindle_3=@Flatness_Spindle_3, Parallelism_1=@Parallelism_1, Parallelism_2=@Parallelism_2, Parallelism_3=@Parallelism_3, Plus_10_1=@Plus_10_1, Plus_10_2=@Plus_10_2, Plus_10_3=@Plus_10_3, Plus_20_1=@Plus_20_1, Plus_20_2=@Plus_20_2, Plus_20_3=@Plus_20_3, Plus_30_1=@Plus_30_1, Plus_30_2=@Plus_30_2, Plus_30_3=@Plus_30_3, Plus_40_1=@Plus_40_1, Plus_40_2=@Plus_40_2, Plus_40_3=@Plus_40_3, Plus_50_1=@Plus_50_1, Plus_50_2=@Plus_50_2, Plus_50_3=@Plus_50_3, Plus_60_1=@Plus_60_1, Plus_60_2=@Plus_60_2, Plus_60_3=@Plus_60_3, Minus_10_1=@Minus_10_1, Minus_10_2=@Minus_10_2, Minus_10_3=@Minus_10_3, Minus_20_1=@Minus_20_1, Minus_20_2=@Minus_20_2, Minus_20_3=@Minus_20_3, Minus_30_1=@Minus_30_1, Minus_30_2=@Minus_30_2, Minus_30_3=@Minus_30_3, Minus_40_1=@Minus_40_1, Minus_40_2=@Minus_40_2, Minus_40_3=@Minus_40_3, Minus_50_1=@Minus_50_1, Minus_50_2=@Minus_50_2, Minus_50_3=@Minus_50_3, Minus_60_1=@Minus_60_1, Minus_60_2=@Minus_60_2, Minus_60_3=@Minus_60_3, DepthError=@DepthError, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.TimeOfDay)
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Flatness_Anvil_1", anvil(0))
                    cmd.Parameters.AddWithValue("@Flatness_Anvil_2", anvil(1))
                    cmd.Parameters.AddWithValue("@Flatness_Anvil_3", anvil(2))
                    cmd.Parameters.AddWithValue("@Flatness_Spindle_1", spindle(0))
                    cmd.Parameters.AddWithValue("@Flatness_Spindle_2", spindle(1))
                    cmd.Parameters.AddWithValue("@Flatness_Spindle_3", spindle(2))
                    cmd.Parameters.AddWithValue("@Parallelism_1", parallelism(0))
                    cmd.Parameters.AddWithValue("@Parallelism_2", parallelism(1))
                    cmd.Parameters.AddWithValue("@Parallelism_3", parallelism(2))

                    For Each kvp In plusObs
                        cmd.Parameters.AddWithValue($"@Plus_{kvp.Key}_1", kvp.Value(0))
                        cmd.Parameters.AddWithValue($"@Plus_{kvp.Key}_2", kvp.Value(1))
                        cmd.Parameters.AddWithValue($"@Plus_{kvp.Key}_3", kvp.Value(2))
                    Next

                    For Each kvp In minusObs
                        cmd.Parameters.AddWithValue($"@Minus_{kvp.Key}_1", kvp.Value(0))
                        cmd.Parameters.AddWithValue($"@Minus_{kvp.Key}_2", kvp.Value(1))
                        cmd.Parameters.AddWithValue($"@Minus_{kvp.Key}_3", kvp.Value(2))
                    Next

                    cmd.Parameters.AddWithValue("@DepthError", depthError)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdatePassameter60: " & ex.Message)
            Return False
        End Try
        Return False
    End Function

    Public Function GetDialGuage100MasterLimits(gaugeType As String, lc As String, sizeRange As String) As DataRow
        Try
            Dim query As String = "SELECT * FROM dial_gauge_100 WHERE RowType='MASTER' AND Type=@Type AND LC=@LC AND Size=@Size LIMIT 1"
            Dim dt As New DataTable()
            Using cmd As New MySqlCommand(query, _con)
                cmd.Parameters.AddWithValue("@Type", gaugeType)
                cmd.Parameters.AddWithValue("@LC", lc)
                cmd.Parameters.AddWithValue("@Size", sizeRange)
                Using sda As New MySqlDataAdapter(cmd)
                    sda.Fill(dt)
                End Using
            End Using
            If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Catch ex As Exception
            Console.WriteLine("Error in GetDialGuage100MasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetDialGuage100Data(controlNo As String, cycleName As String) As DataTable
        Try
            Dim query As String = "SELECT * FROM dial_gauge_100 WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName ORDER BY id DESC LIMIT 1"
            Dim dt As New DataTable()
            Using cmd As New MySqlCommand(query, _con)
                cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                cmd.Parameters.AddWithValue("@CycleName", cycleName)
                Using sda As New MySqlDataAdapter(cmd)
                    sda.Fill(dt)
                End Using
            End Using
            Return dt
        Catch ex As Exception
            Console.WriteLine("Error in GetDialGuage100Data: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function InsertDialGuage100(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, obs1_10Rev As Object(), obs1_2Rev As Object(), obs1Rev As Object(), obsWholeRange As Object(), obsHysteresis As Object(), obsRepeatability As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO dial_gauge_100 (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, Obs_1_10_Rev_1, Obs_1_10_Rev_2, Obs_1_10_Rev_3, Obs_1_2_Rev_1, Obs_1_2_Rev_2, Obs_1_2_Rev_3, Obs_1_Rev_1, Obs_1_Rev_2, Obs_1_Rev_3, Obs_Whole_Range_1, Obs_Whole_Range_2, Obs_Whole_Range_3, Obs_Hysteresis_1, Obs_Hysteresis_2, Obs_Hysteresis_3, Obs_Repeatability_1, Obs_Repeatability_2, Obs_Repeatability_3, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @Obs_1_10_Rev_1, @Obs_1_10_Rev_2, @Obs_1_10_Rev_3, @Obs_1_2_Rev_1, @Obs_1_2_Rev_2, @Obs_1_2_Rev_3, @Obs_1_Rev_1, @Obs_1_Rev_2, @Obs_1_Rev_3, @Obs_Whole_Range_1, @Obs_Whole_Range_2, @Obs_Whole_Range_3, @Obs_Hysteresis_1, @Obs_Hysteresis_2, @Obs_Hysteresis_3, @Obs_Repeatability_1, @Obs_Repeatability_2, @Obs_Repeatability_3, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Obs_1_10_Rev_1", obs1_10Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_1_10_Rev_2", obs1_10Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_1_10_Rev_3", obs1_10Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_1_2_Rev_1", obs1_2Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_1_2_Rev_2", obs1_2Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_1_2_Rev_3", obs1_2Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_1_Rev_1", obs1Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_1_Rev_2", obs1Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_1_Rev_3", obs1Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_Whole_Range_1", obsWholeRange(0))
                    cmd.Parameters.AddWithValue("@Obs_Whole_Range_2", obsWholeRange(1))
                    cmd.Parameters.AddWithValue("@Obs_Whole_Range_3", obsWholeRange(2))

                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_1", obsHysteresis(0))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_2", obsHysteresis(1))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_3", obsHysteresis(2))

                    cmd.Parameters.AddWithValue("@Obs_Repeatability_1", obsRepeatability(0))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_2", obsRepeatability(1))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_3", obsRepeatability(2))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertDialGuage100: " & ex.Message)
            Return False
        End Try
        Return False
    End Function

    Public Function UpdateDialGuage100(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, obs1_10Rev As Object(), obs1_2Rev As Object(), obs1Rev As Object(), obsWholeRange As Object(), obsHysteresis As Object(), obsRepeatability As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE dial_gauge_100 SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, Obs_1_10_Rev_1=@Obs_1_10_Rev_1, Obs_1_10_Rev_2=@Obs_1_10_Rev_2, Obs_1_10_Rev_3=@Obs_1_10_Rev_3, Obs_1_2_Rev_1=@Obs_1_2_Rev_1, Obs_1_2_Rev_2=@Obs_1_2_Rev_2, Obs_1_2_Rev_3=@Obs_1_2_Rev_3, Obs_1_Rev_1=@Obs_1_Rev_1, Obs_1_Rev_2=@Obs_1_Rev_2, Obs_1_Rev_3=@Obs_1_Rev_3, Obs_Whole_Range_1=@Obs_Whole_Range_1, Obs_Whole_Range_2=@Obs_Whole_Range_2, Obs_Whole_Range_3=@Obs_Whole_Range_3, Obs_Hysteresis_1=@Obs_Hysteresis_1, Obs_Hysteresis_2=@Obs_Hysteresis_2, Obs_Hysteresis_3=@Obs_Hysteresis_3, Obs_Repeatability_1=@Obs_Repeatability_1, Obs_Repeatability_2=@Obs_Repeatability_2, Obs_Repeatability_3=@Obs_Repeatability_3, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Obs_1_10_Rev_1", obs1_10Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_1_10_Rev_2", obs1_10Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_1_10_Rev_3", obs1_10Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_1_2_Rev_1", obs1_2Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_1_2_Rev_2", obs1_2Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_1_2_Rev_3", obs1_2Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_1_Rev_1", obs1Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_1_Rev_2", obs1Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_1_Rev_3", obs1Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_Whole_Range_1", obsWholeRange(0))
                    cmd.Parameters.AddWithValue("@Obs_Whole_Range_2", obsWholeRange(1))
                    cmd.Parameters.AddWithValue("@Obs_Whole_Range_3", obsWholeRange(2))

                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_1", obsHysteresis(0))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_2", obsHysteresis(1))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_3", obsHysteresis(2))

                    cmd.Parameters.AddWithValue("@Obs_Repeatability_1", obsRepeatability(0))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_2", obsRepeatability(1))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_3", obsRepeatability(2))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateDialGuage100: " & ex.Message)
            Return False
        End Try
        Return False
    End Function

    ' ==========================================
    ' Dial Test Indicator Methods
    ' ==========================================
    Public Function GetDialtestindicator100MasterLimits(lc As String, sizeRange As String, probeSizeL1 As String) As DataRow
        Try
            Dim query As String = "SELECT * FROM dial_test_indicator_100 WHERE RowType='MASTER' AND LC=@LC AND Size=@Size AND ProbeSize_L1=@ProbeSize_L1 LIMIT 1"
            Dim dt As New DataTable()
            Using cmd As New MySqlCommand(query, _con)
                cmd.Parameters.AddWithValue("@LC", lc)
                cmd.Parameters.AddWithValue("@Size", sizeRange)
                cmd.Parameters.AddWithValue("@ProbeSize_L1", probeSizeL1)
                Using sda As New MySqlDataAdapter(cmd)
                    sda.Fill(dt)
                End Using
            End Using
            If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Catch ex As Exception
            Console.WriteLine("Error in GetDialtestindicator100MasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetDialtestindicator100Data(controlNo As String, cycleName As String) As DataTable
        Try
            Dim query As String = "SELECT * FROM dial_test_indicator_100 WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName ORDER BY id DESC LIMIT 1"
            Dim dt As New DataTable()
            Using cmd As New MySqlCommand(query, _con)
                cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                cmd.Parameters.AddWithValue("@CycleName", cycleName)
                Using sda As New MySqlDataAdapter(cmd)
                    sda.Fill(dt)
                End Using
            End Using
            Return dt
        Catch ex As Exception
            Console.WriteLine("Error in GetDialtestindicator100Data: " & ex.Message)
            Return New DataTable()
        End Try
    End Function

    Public Function InsertDialtestindicator100(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, probeSizeL1 As String, color As String, location As String, temp As String, humidity As String, tmu As String, obsScaleDiv As Object(), obsOver1Rev As Object(), obsOverMeas As Object(), obsHysteresis As Object(), obsRepeatability As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO dial_test_indicator_100 (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, ProbeSize_L1, Color, Location, Temperature, Humidity, TMU, Obs_10_Scale_Div_1, Obs_10_Scale_Div_2, Obs_10_Scale_Div_3, Obs_Over_1_Rev_1, Obs_Over_1_Rev_2, Obs_Over_1_Rev_3, Obs_Over_Meas_Range_1, Obs_Over_Meas_Range_2, Obs_Over_Meas_Range_3, Obs_Hysteresis_1, Obs_Hysteresis_2, Obs_Hysteresis_3, Obs_Repeatability_1, Obs_Repeatability_2, Obs_Repeatability_3, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @ProbeSize_L1, @Color, @Location, @Temperature, @Humidity, @TMU, @Obs_10_Scale_Div_1, @Obs_10_Scale_Div_2, @Obs_10_Scale_Div_3, @Obs_Over_1_Rev_1, @Obs_Over_1_Rev_2, @Obs_Over_1_Rev_3, @Obs_Over_Meas_Range_1, @Obs_Over_Meas_Range_2, @Obs_Over_Meas_Range_3, @Obs_Hysteresis_1, @Obs_Hysteresis_2, @Obs_Hysteresis_3, @Obs_Repeatability_1, @Obs_Repeatability_2, @Obs_Repeatability_3, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@ProbeSize_L1", probeSizeL1)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Obs_10_Scale_Div_1", obsScaleDiv(0))
                    cmd.Parameters.AddWithValue("@Obs_10_Scale_Div_2", obsScaleDiv(1))
                    cmd.Parameters.AddWithValue("@Obs_10_Scale_Div_3", obsScaleDiv(2))

                    cmd.Parameters.AddWithValue("@Obs_Over_1_Rev_1", obsOver1Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_Over_1_Rev_2", obsOver1Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_Over_1_Rev_3", obsOver1Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_Over_Meas_Range_1", obsOverMeas(0))
                    cmd.Parameters.AddWithValue("@Obs_Over_Meas_Range_2", obsOverMeas(1))
                    cmd.Parameters.AddWithValue("@Obs_Over_Meas_Range_3", obsOverMeas(2))

                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_1", obsHysteresis(0))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_2", obsHysteresis(1))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_3", obsHysteresis(2))

                    cmd.Parameters.AddWithValue("@Obs_Repeatability_1", obsRepeatability(0))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_2", obsRepeatability(1))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_3", obsRepeatability(2))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertDialtestindicator100: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateDialtestindicator100(type As String, controlNo As String, cycleName As String, size As String, lc As String, probeSizeL1 As String, color As String, location As String, temp As String, humidity As String, tmu As String, obsScaleDiv As Object(), obsOver1Rev As Object(), obsOverMeas As Object(), obsHysteresis As Object(), obsRepeatability As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE dial_test_indicator_100 SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, ProbeSize_L1=@ProbeSize_L1, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, Obs_10_Scale_Div_1=@Obs_10_Scale_Div_1, Obs_10_Scale_Div_2=@Obs_10_Scale_Div_2, Obs_10_Scale_Div_3=@Obs_10_Scale_Div_3, Obs_Over_1_Rev_1=@Obs_Over_1_Rev_1, Obs_Over_1_Rev_2=@Obs_Over_1_Rev_2, Obs_Over_1_Rev_3=@Obs_Over_1_Rev_3, Obs_Over_Meas_Range_1=@Obs_Over_Meas_Range_1, Obs_Over_Meas_Range_2=@Obs_Over_Meas_Range_2, Obs_Over_Meas_Range_3=@Obs_Over_Meas_Range_3, Obs_Hysteresis_1=@Obs_Hysteresis_1, Obs_Hysteresis_2=@Obs_Hysteresis_2, Obs_Hysteresis_3=@Obs_Hysteresis_3, Obs_Repeatability_1=@Obs_Repeatability_1, Obs_Repeatability_2=@Obs_Repeatability_2, Obs_Repeatability_3=@Obs_Repeatability_3, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@ProbeSize_L1", probeSizeL1)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Obs_10_Scale_Div_1", obsScaleDiv(0))
                    cmd.Parameters.AddWithValue("@Obs_10_Scale_Div_2", obsScaleDiv(1))
                    cmd.Parameters.AddWithValue("@Obs_10_Scale_Div_3", obsScaleDiv(2))

                    cmd.Parameters.AddWithValue("@Obs_Over_1_Rev_1", obsOver1Rev(0))
                    cmd.Parameters.AddWithValue("@Obs_Over_1_Rev_2", obsOver1Rev(1))
                    cmd.Parameters.AddWithValue("@Obs_Over_1_Rev_3", obsOver1Rev(2))

                    cmd.Parameters.AddWithValue("@Obs_Over_Meas_Range_1", obsOverMeas(0))
                    cmd.Parameters.AddWithValue("@Obs_Over_Meas_Range_2", obsOverMeas(1))
                    cmd.Parameters.AddWithValue("@Obs_Over_Meas_Range_3", obsOverMeas(2))

                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_1", obsHysteresis(0))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_2", obsHysteresis(1))
                    cmd.Parameters.AddWithValue("@Obs_Hysteresis_3", obsHysteresis(2))

                    cmd.Parameters.AddWithValue("@Obs_Repeatability_1", obsRepeatability(0))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_2", obsRepeatability(1))
                    cmd.Parameters.AddWithValue("@Obs_Repeatability_3", obsRepeatability(2))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateDialtestindicator100: " & ex.Message)
        End Try
        Return False
    End Function

    ' =========================================================================
    ' Bore Gauge Methods
    ' =========================================================================
    Public Function GetBoreGauge3MasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM bore_gauge_3 WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetBoreGauge3MasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetBoreGauge3Data(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM bore_gauge_3 WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertBoreGauge3(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, wideRange As Object(), adjacent As Object(), repeatability As Object(), depthError As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO bore_gauge_3 (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, Wide_Range_1, Wide_Range_2, Wide_Range_3, Adjacent_Error_1, Adjacent_Error_2, Adjacent_Error_3, Repeatability_1, Repeatability_2, Repeatability_3, DepthError, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @Wide_Range_1, @Wide_Range_2, @Wide_Range_3, @Adjacent_Error_1, @Adjacent_Error_2, @Adjacent_Error_3, @Repeatability_1, @Repeatability_2, @Repeatability_3, @DepthError, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Wide_Range_1", wideRange(0))
                    cmd.Parameters.AddWithValue("@Wide_Range_2", wideRange(1))
                    cmd.Parameters.AddWithValue("@Wide_Range_3", wideRange(2))

                    cmd.Parameters.AddWithValue("@Adjacent_Error_1", adjacent(0))
                    cmd.Parameters.AddWithValue("@Adjacent_Error_2", adjacent(1))
                    cmd.Parameters.AddWithValue("@Adjacent_Error_3", adjacent(2))

                    cmd.Parameters.AddWithValue("@Repeatability_1", repeatability(0))
                    cmd.Parameters.AddWithValue("@Repeatability_2", repeatability(1))
                    cmd.Parameters.AddWithValue("@Repeatability_3", repeatability(2))

                    cmd.Parameters.AddWithValue("@DepthError", depthError)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertBoreGauge3: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateBoreGauge3(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, wideRange As Object(), adjacent As Object(), repeatability As Object(), depthError As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE bore_gauge_3 SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, Wide_Range_1=@Wide_Range_1, Wide_Range_2=@Wide_Range_2, Wide_Range_3=@Wide_Range_3, Adjacent_Error_1=@Adjacent_Error_1, Adjacent_Error_2=@Adjacent_Error_2, Adjacent_Error_3=@Adjacent_Error_3, Repeatability_1=@Repeatability_1, Repeatability_2=@Repeatability_2, Repeatability_3=@Repeatability_3, DepthError=@DepthError, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)

                    cmd.Parameters.AddWithValue("@Wide_Range_1", wideRange(0))
                    cmd.Parameters.AddWithValue("@Wide_Range_2", wideRange(1))
                    cmd.Parameters.AddWithValue("@Wide_Range_3", wideRange(2))

                    cmd.Parameters.AddWithValue("@Adjacent_Error_1", adjacent(0))
                    cmd.Parameters.AddWithValue("@Adjacent_Error_2", adjacent(1))
                    cmd.Parameters.AddWithValue("@Adjacent_Error_3", adjacent(2))

                    cmd.Parameters.AddWithValue("@Repeatability_1", repeatability(0))
                    cmd.Parameters.AddWithValue("@Repeatability_2", repeatability(1))
                    cmd.Parameters.AddWithValue("@Repeatability_3", repeatability(2))

                    cmd.Parameters.AddWithValue("@DepthError", depthError)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateBoreGauge3: " & ex.Message)
        End Try
        Return False
    End Function

    ' =========================================================================
    ' Height Master Gauge Methods
    ' =========================================================================
    Public Function GetHeightMasterGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM height_master_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM height_master_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetHeightMasterGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetHeightMasterGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM height_master_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertHeightMasterGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, flatnessLimits As Object(), parallelismLimits As Object(), goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO height_master_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Flatness_UL, Flatness_LL, Parallelism_UL, Parallelism_LL, Go_Size, Go_Min_Limit, Go_Max_Limit, Go_Obs_1, Go_Obs_2, Go_Obs_3, NoGo_Size, NoGo_Min_Limit, NoGo_Max_Limit, NoGo_Obs_1, NoGo_Obs_2, NoGo_Obs_3, TimeIn, TimeOut, TotalTime, Judgement, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Flatness_UL, @Flatness_LL, @Parallelism_UL, @Parallelism_LL, @Go_Size, @Go_Min_Limit, @Go_Max_Limit, @Go_Obs_1, @Go_Obs_2, @Go_Obs_3, @NoGo_Size, @NoGo_Min_Limit, @NoGo_Max_Limit, @NoGo_Obs_1, @NoGo_Obs_2, @NoGo_Obs_3, @TimeIn, @TimeOut, @TotalTime, @Judgement, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))

                    cmd.Parameters.AddWithValue("@Flatness_UL", flatnessLimits(0))
                    cmd.Parameters.AddWithValue("@Flatness_LL", flatnessLimits(1))
                    cmd.Parameters.AddWithValue("@Parallelism_UL", parallelismLimits(0))
                    cmd.Parameters.AddWithValue("@Parallelism_LL", parallelismLimits(1))

                    cmd.Parameters.AddWithValue("@Go_Size", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Size", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Judgement", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertHeightMasterGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateHeightMasterGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, flatnessLimits As Object(), parallelismLimits As Object(), goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE height_master_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Flatness_UL=@Flatness_UL, Flatness_LL=@Flatness_LL, Parallelism_UL=@Parallelism_UL, Parallelism_LL=@Parallelism_LL, Go_Size=@Go_Size, Go_Min_Limit=@Go_Min_Limit, Go_Max_Limit=@Go_Max_Limit, Go_Obs_1=@Go_Obs_1, Go_Obs_2=@Go_Obs_2, Go_Obs_3=@Go_Obs_3, NoGo_Size=@NoGo_Size, NoGo_Min_Limit=@NoGo_Min_Limit, NoGo_Max_Limit=@NoGo_Max_Limit, NoGo_Obs_1=@NoGo_Obs_1, NoGo_Obs_2=@NoGo_Obs_2, NoGo_Obs_3=@NoGo_Obs_3, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Judgement=@Judgement, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))

                    cmd.Parameters.AddWithValue("@Flatness_UL", flatnessLimits(0))
                    cmd.Parameters.AddWithValue("@Flatness_LL", flatnessLimits(1))
                    cmd.Parameters.AddWithValue("@Parallelism_UL", parallelismLimits(0))
                    cmd.Parameters.AddWithValue("@Parallelism_LL", parallelismLimits(1))

                    cmd.Parameters.AddWithValue("@Go_Size", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Size", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Judgement", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateHeightMasterGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    ' =========================================================================
    ' Key Groove Gauge Methods
    ' =========================================================================
    Public Function GetKeyGrooveGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM key_groove_guage_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM key_groove_guage_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetKeyGrooveGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetKeyGrooveGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM key_groove_guage_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertKeyGrooveGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO key_groove_guage_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Parameter, Location_Obs, Go_Min_Limit, Go_Max_Limit, Go_Obs_1, Go_Obs_2, Go_Obs_3, Go_Obs_4, NoGo_Min_Limit, NoGo_Max_Limit, NoGo_Obs_1, NoGo_Obs_2, NoGo_Obs_3, NoGo_Obs_4, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Parameter, @Location_Obs, @Go_Min_Limit, @Go_Max_Limit, @Go_Obs_1, @Go_Obs_2, @Go_Obs_3, @Go_Obs_4, @NoGo_Min_Limit, @NoGo_Max_Limit, @NoGo_Obs_1, @NoGo_Obs_2, @NoGo_Obs_3, @NoGo_Obs_4, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("Error in InsertKeyGrooveGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateKeyGrooveGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE key_groove_guage_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Parameter=@Parameter, Location_Obs=@Location_Obs, Go_Min_Limit=@Go_Min_Limit, Go_Max_Limit=@Go_Max_Limit, Go_Obs_1=@Go_Obs_1, Go_Obs_2=@Go_Obs_2, Go_Obs_3=@Go_Obs_3, Go_Obs_4=@Go_Obs_4, NoGo_Min_Limit=@NoGo_Min_Limit, NoGo_Max_Limit=@NoGo_Max_Limit, NoGo_Obs_1=@NoGo_Obs_1, NoGo_Obs_2=@NoGo_Obs_2, NoGo_Obs_3=@NoGo_Obs_3, NoGo_Obs_4=@NoGo_Obs_4, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            LastError = ex.Message
            Console.WriteLine("Error in UpdateKeyGrooveGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetPlainPlugGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM plain_plug_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM plain_plug_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetPlainPlugGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetPlainPlugGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM plain_plug_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertPlainPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO plain_plug_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Parameter, Location_Obs, Go_Min_Limit, Go_Max_Limit, Go_Obs_1, Go_Obs_2, Go_Obs_3, Go_Obs_4, NoGo_Min_Limit, NoGo_Max_Limit, NoGo_Obs_1, NoGo_Obs_2, NoGo_Obs_3, NoGo_Obs_4, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Parameter, @Location_Obs, @Go_Min_Limit, @Go_Max_Limit, @Go_Obs_1, @Go_Obs_2, @Go_Obs_3, @Go_Obs_4, @NoGo_Min_Limit, @NoGo_Max_Limit, @NoGo_Obs_1, @NoGo_Obs_2, @NoGo_Obs_3, @NoGo_Obs_4, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertPlainPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdatePlainPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE plain_plug_gauge_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Parameter=@Parameter, Location_Obs=@Location_Obs, Go_Min_Limit=@Go_Min_Limit, Go_Max_Limit=@Go_Max_Limit, Go_Obs_1=@Go_Obs_1, Go_Obs_2=@Go_Obs_2, Go_Obs_3=@Go_Obs_3, Go_Obs_4=@Go_Obs_4, NoGo_Min_Limit=@NoGo_Min_Limit, NoGo_Max_Limit=@NoGo_Max_Limit, NoGo_Obs_1=@NoGo_Obs_1, NoGo_Obs_2=@NoGo_Obs_2, NoGo_Obs_3=@NoGo_Obs_3, NoGo_Obs_4=@NoGo_Obs_4, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdatePlainPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetPlainRingGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM plain_ring_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM plain_ring_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetPlainRingGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetPlainRingGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM plain_ring_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertPlainRingGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, minLimit As Object, maxLimit As Object, obsDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO plain_ring_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Parameter, Location_Obs, Min_Limit, Max_Limit, Obs_1, Obs_2, Obs_3, Obs_4, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Parameter, @Location_Obs, @Min_Limit, @Max_Limit, @Obs_1, @Obs_2, @Obs_3, @Obs_4, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Min_Limit", minLimit)
                    cmd.Parameters.AddWithValue("@Max_Limit", maxLimit)
                    cmd.Parameters.AddWithValue("@Obs_1", obsDetails(0))
                    cmd.Parameters.AddWithValue("@Obs_2", obsDetails(1))
                    cmd.Parameters.AddWithValue("@Obs_3", obsDetails(2))
                    cmd.Parameters.AddWithValue("@Obs_4", obsDetails(3))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertPlainRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdatePlainRingGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, minLimit As Object, maxLimit As Object, obsDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE plain_ring_gauge_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Parameter=@Parameter, Location_Obs=@Location_Obs, Min_Limit=@Min_Limit, Max_Limit=@Max_Limit, Obs_1=@Obs_1, Obs_2=@Obs_2, Obs_3=@Obs_3, Obs_4=@Obs_4, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Min_Limit", minLimit)
                    cmd.Parameters.AddWithValue("@Max_Limit", maxLimit)
                    cmd.Parameters.AddWithValue("@Obs_1", obsDetails(0))
                    cmd.Parameters.AddWithValue("@Obs_2", obsDetails(1))
                    cmd.Parameters.AddWithValue("@Obs_3", obsDetails(2))
                    cmd.Parameters.AddWithValue("@Obs_4", obsDetails(3))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdatePlainRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetSnapGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM snap_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM snap_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetSnapGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetSnapGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM snap_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertSnapGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO snap_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Parameter, Location_Obs, Go_Min_Limit, Go_Max_Limit, Go_Obs_1, Go_Obs_2, Go_Obs_3, Go_Obs_4, NoGo_Min_Limit, NoGo_Max_Limit, NoGo_Obs_1, NoGo_Obs_2, NoGo_Obs_3, NoGo_Obs_4, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Parameter, @Location_Obs, @Go_Min_Limit, @Go_Max_Limit, @Go_Obs_1, @Go_Obs_2, @Go_Obs_3, @Go_Obs_4, @NoGo_Min_Limit, @NoGo_Max_Limit, @NoGo_Obs_1, @NoGo_Obs_2, @NoGo_Obs_3, @NoGo_Obs_4, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertSnapGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateSnapGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE snap_gauge_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Parameter=@Parameter, Location_Obs=@Location_Obs, Go_Min_Limit=@Go_Min_Limit, Go_Max_Limit=@Go_Max_Limit, Go_Obs_1=@Go_Obs_1, Go_Obs_2=@Go_Obs_2, Go_Obs_3=@Go_Obs_3, Go_Obs_4=@Go_Obs_4, NoGo_Min_Limit=@NoGo_Min_Limit, NoGo_Max_Limit=@NoGo_Max_Limit, NoGo_Obs_1=@NoGo_Obs_1, NoGo_Obs_2=@NoGo_Obs_2, NoGo_Obs_3=@NoGo_Obs_3, NoGo_Obs_4=@NoGo_Obs_4, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateSnapGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetSpecialHeightGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM special_height_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM special_height_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetSpecialHeightGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetSpecialHeightGaugeMasterData(controlNo As String, cycleName As String) As DataRow
        Dim dt = ReadDatatable($"SELECT * FROM special_height_gauge_calibration WHERE RowType='MASTER' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Return Nothing
    End Function

    Public Function GetSpecialHeightGaugeRecordData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM special_height_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id ASC")
    End Function

    Public Sub DeleteSpecialHeightGaugeCalibration(controlNo As String, cycleName As String)
        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand("DELETE FROM special_height_gauge_calibration WHERE ControlNo=@cno AND CycleName=@cyc", _con)
                    cmd.Parameters.AddWithValue("@cno", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error deleting SpecialHeightGauge: " & ex.Message)
        End Try
    End Sub

    Public Function InsertSpecialHeightGaugeMaster(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, tempMin As Object, tempMax As Object, humMin As Object, humMax As Object, status As String, remark As String, preparedBy As String, timeIn As String, timeOut As String, totalTime As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO special_height_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, Status, Remark, PreparedBy, TimeIn, TimeOut, TotalTime) VALUES ('MASTER', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Env_Temp_LL, @Env_Temp_UL, @Env_Hum_LL, @Env_Hum_UL, @Status, @Remark, @PreparedBy, @TimeIn, @TimeOut, @TotalTime)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Env_Temp_LL", tempMin)
                    cmd.Parameters.AddWithValue("@Env_Temp_UL", tempMax)
                    cmd.Parameters.AddWithValue("@Env_Hum_LL", humMin)
                    cmd.Parameters.AddWithValue("@Env_Hum_UL", humMax)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.Parameters.AddWithValue("@PreparedBy", preparedBy)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertSpecialHeightGaugeMaster: " & ex.Message)
        End Try
        Return False
    End Function

    Public Sub InsertSpecialHeightGaugeRecord(controlNo As String, cycleName As String, nominal As String, minLimit As Object, maxLimit As Object, o1 As Object, o2 As Object, o3 As Object, o4 As Object)
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO special_height_gauge_calibration (RowType, ControlNo, CycleName, Nominal, Min_Limit, Max_Limit, Obs_1, Obs_2, Obs_3, Obs_4) VALUES ('RECORD', @ControlNo, @CycleName, @Nominal, @MinLimit, @MaxLimit, @O1, @O2, @O3, @O4)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Nominal", nominal)
                    cmd.Parameters.AddWithValue("@MinLimit", If(minLimit IsNot Nothing, minLimit, DBNull.Value))
                    cmd.Parameters.AddWithValue("@MaxLimit", If(maxLimit IsNot Nothing, maxLimit, DBNull.Value))
                    cmd.Parameters.AddWithValue("@O1", If(o1 IsNot Nothing, o1, DBNull.Value))
                    cmd.Parameters.AddWithValue("@O2", If(o2 IsNot Nothing, o2, DBNull.Value))
                    cmd.Parameters.AddWithValue("@O3", If(o3 IsNot Nothing, o3, DBNull.Value))
                    cmd.Parameters.AddWithValue("@O4", If(o4 IsNot Nothing, o4, DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertSpecialHeightGaugeRecord: " & ex.Message)
        End Try
    End Sub

    Public Function GetSpecialDepthGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM special_depth_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM special_depth_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetSpecialDepthGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetSpecialDepthGaugeMasterData(controlNo As String, cycleName As String) As DataRow
        Dim dt = ReadDatatable($"SELECT * FROM special_depth_gauge_calibration WHERE RowType='MASTER' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Return Nothing
    End Function

    Public Function GetSpecialDepthGaugeRecordData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM special_depth_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id ASC")
    End Function

    Public Sub DeleteSpecialDepthGaugeCalibration(controlNo As String, cycleName As String)
        Try
            If MySQLDBConnect() = 1 Then
                Using cmd As New MySqlCommand("DELETE FROM special_depth_gauge_calibration WHERE ControlNo=@cno AND CycleName=@cyc", _con)
                    cmd.Parameters.AddWithValue("@cno", controlNo)
                    cmd.Parameters.AddWithValue("@cyc", cycleName)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error deleting SpecialDepthGauge: " & ex.Message)
        End Try
    End Sub

    Public Sub EnsureParameterConfigColumn()
        Try
            If MySQLDBConnect() = 1 Then
                Dim checkCmd As New MySqlCommand("SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'special_depth_gauge_calibration' AND COLUMN_NAME = 'ParameterConfig'", _con)
                Dim count = Convert.ToInt32(checkCmd.ExecuteScalar())
                If count = 0 Then
                    Dim addCmd As New MySqlCommand("ALTER TABLE special_depth_gauge_calibration ADD COLUMN ParameterConfig VARCHAR(100) NULL", _con)
                    addCmd.ExecuteNonQuery()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error in EnsureParameterConfigColumn: " & ex.Message)
        End Try
    End Sub

    Public Sub EnsureSpecialDepthColumns(diaCount As Integer, distCount As Integer, angCount As Integer)
        Try
            If MySQLDBConnect() = 1 Then
                Dim checkCmd As New MySqlCommand("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'special_depth_gauge_calibration'", _con)
                Dim existingCols As New System.Collections.Generic.HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Using reader = checkCmd.ExecuteReader()
                    While reader.Read()
                        existingCols.Add(reader("COLUMN_NAME").ToString())
                    End While
                End Using

                Dim toAdd As New System.Collections.Generic.List(Of String)()
                Dim AddColsForPrefix = Sub(prefix As String)
                                         If Not existingCols.Contains(prefix & "_Nominal") Then toAdd.Add($"ADD COLUMN {prefix}_Nominal VARCHAR(100) NULL")
                                         If Not existingCols.Contains(prefix & "_Permissible_Error") Then toAdd.Add($"ADD COLUMN {prefix}_Permissible_Error VARCHAR(100) NULL")
                                         If Not existingCols.Contains(prefix & "_Min_Limit") Then toAdd.Add($"ADD COLUMN {prefix}_Min_Limit DECIMAL(10,4) NULL")
                                         If Not existingCols.Contains(prefix & "_Max_Limit") Then toAdd.Add($"ADD COLUMN {prefix}_Max_Limit DECIMAL(10,4) NULL")
                                         If Not existingCols.Contains(prefix & "_Obs_1") Then toAdd.Add($"ADD COLUMN {prefix}_Obs_1 TEXT NULL")
                                         If Not existingCols.Contains(prefix & "_Obs_2") Then toAdd.Add($"ADD COLUMN {prefix}_Obs_2 TEXT NULL")
                                         If Not existingCols.Contains(prefix & "_Obs_3") Then toAdd.Add($"ADD COLUMN {prefix}_Obs_3 TEXT NULL")
                                       End Sub

                For i = 1 To diaCount
                    AddColsForPrefix("Dia_" & i)
                Next
                For i = 1 To distCount
                    AddColsForPrefix("Distance_" & i)
                Next
                For i = 1 To angCount
                    AddColsForPrefix("Angle_" & i)
                Next

                If toAdd.Count > 0 Then
                    Dim alterQuery = "ALTER TABLE special_depth_gauge_calibration " & String.Join(", ", toAdd)
                    Dim alterCmd As New MySqlCommand(alterQuery, _con)
                    alterCmd.ExecuteNonQuery()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error in EnsureSpecialDepthColumns: " & ex.Message)
        End Try
    End Sub

    Public Function InsertSpecialDepthGaugeMaster(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, tempMin As Object, tempMax As Object, humMin As Object, humMax As Object, status As String, remark As String, timeIn As String, timeOut As String, totalTime As String, paramConfig As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO special_depth_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, Status, Remark, TimeIn, TimeOut, TotalTime, ParameterConfig) VALUES ('MASTER', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Env_Temp_LL, @Env_Temp_UL, @Env_Hum_LL, @Env_Hum_UL, @Status, @Remark, @TimeIn, @TimeOut, @TotalTime, @ParameterConfig)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Env_Temp_LL", tempMin)
                    cmd.Parameters.AddWithValue("@Env_Temp_UL", tempMax)
                    cmd.Parameters.AddWithValue("@Env_Hum_LL", humMin)
                    cmd.Parameters.AddWithValue("@Env_Hum_UL", humMax)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@ParameterConfig", paramConfig)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertSpecialDepthGaugeMaster: " & ex.Message)
        End Try
        Return False
    End Function

    Public Sub InsertSpecialDepthGaugeDynamicRecord(controlNo As String, cycleName As String, params As System.Collections.Generic.List(Of SpecialDepthDynParam))
        Try
            If MySQLDBConnect() = 1 Then
                Dim cols As New System.Collections.Generic.List(Of String)()
                Dim vals As New System.Collections.Generic.List(Of String)()
                Dim cmd As New MySqlCommand("", _con)
                
                cols.Add("RowType") : vals.Add("'RECORD'")
                cols.Add("ControlNo") : vals.Add("@CNO")
                cols.Add("CycleName") : vals.Add("@CYC")
                cmd.Parameters.AddWithValue("@CNO", controlNo)
                cmd.Parameters.AddWithValue("@CYC", cycleName)

                For i As Integer = 0 To params.Count - 1
                    Dim p = params(i)
                    Dim pfx = p.ParamPrefix
                    Dim idx = i.ToString()
                    
                    cols.Add($"{pfx}_Nominal") : vals.Add($"@Nom{idx}")
                    cols.Add($"{pfx}_Permissible_Error") : vals.Add($"@Perm{idx}")
                    cols.Add($"{pfx}_Min_Limit") : vals.Add($"@Min{idx}")
                    cols.Add($"{pfx}_Max_Limit") : vals.Add($"@Max{idx}")
                    cols.Add($"{pfx}_Obs_1") : vals.Add($"@O1_{idx}")
                    cols.Add($"{pfx}_Obs_2") : vals.Add($"@O2_{idx}")
                    cols.Add($"{pfx}_Obs_3") : vals.Add($"@O3_{idx}")

                    cmd.Parameters.AddWithValue($"@Nom{idx}", If(String.IsNullOrEmpty(p.Nominal), "-", p.Nominal))
                    cmd.Parameters.AddWithValue($"@Perm{idx}", If(String.IsNullOrEmpty(p.PermErr), "-", p.PermErr))
                    cmd.Parameters.AddWithValue($"@Min{idx}", If(String.IsNullOrEmpty(p.MinLimit), DBNull.Value, p.MinLimit))
                    cmd.Parameters.AddWithValue($"@Max{idx}", If(String.IsNullOrEmpty(p.MaxLimit), DBNull.Value, p.MaxLimit))
                    cmd.Parameters.AddWithValue($"@O1_{idx}", If(String.IsNullOrEmpty(p.Obs1), "-", p.Obs1))
                    cmd.Parameters.AddWithValue($"@O2_{idx}", If(String.IsNullOrEmpty(p.Obs2), "-", p.Obs2))
                    cmd.Parameters.AddWithValue($"@O3_{idx}", If(String.IsNullOrEmpty(p.Obs3), "-", p.Obs3))
                Next

                Dim query = $"INSERT INTO special_depth_gauge_calibration ({String.Join(", ", cols)}) VALUES ({String.Join(", ", vals)})"
                cmd.CommandText = query
                cmd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertSpecialDepthGaugeDynamicRecord: " & ex.Message)
        End Try
    End Sub

    Public Function GetSplineRingGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM spline_ring_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM spline_ring_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetSplineRingGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetSplineRingGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM spline_ring_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertSplineRingGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, minorDetails As Object(), majorDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO spline_ring_gauge_calibration (" &
                    "RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, " &
                    "Minor_Dia_Min_Limit, Minor_Dia_Max_Limit, Minor_Dia_Obs_1, Minor_Dia_Obs_2, Minor_Dia_Obs_3, " &
                    "Major_Dia_Min_Limit, Major_Dia_Max_Limit, Major_Dia_Obs_1, Major_Dia_Obs_2, Major_Dia_Obs_3, " &
                    "TimeIn, TimeOut, TotalTime, Status, Remark) VALUES (" &
                    "'RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, " &
                    "@MinorMin, @MinorMax, @MinorO1, @MinorO2, @MinorO3, " &
                    "@MajorMin, @MajorMax, @MajorO1, @MajorO2, @MajorO3, " &
                    "@TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))

                    cmd.Parameters.AddWithValue("@MinorMin", minorDetails(1))
                    cmd.Parameters.AddWithValue("@MinorMax", minorDetails(2))
                    cmd.Parameters.AddWithValue("@MinorO1", minorDetails(3))
                    cmd.Parameters.AddWithValue("@MinorO2", minorDetails(4))
                    cmd.Parameters.AddWithValue("@MinorO3", minorDetails(5))

                    cmd.Parameters.AddWithValue("@MajorMin", majorDetails(1))
                    cmd.Parameters.AddWithValue("@MajorMax", majorDetails(2))
                    cmd.Parameters.AddWithValue("@MajorO1", majorDetails(3))
                    cmd.Parameters.AddWithValue("@MajorO2", majorDetails(4))
                    cmd.Parameters.AddWithValue("@MajorO3", majorDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertSplineRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateSplineRingGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, minorDetails As Object(), majorDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE spline_ring_gauge_calibration SET " &
                    "Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, " &
                    "Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, " &
                    "Minor_Dia_Min_Limit=@MinorMin, Minor_Dia_Max_Limit=@MinorMax, Minor_Dia_Obs_1=@MinorO1, Minor_Dia_Obs_2=@MinorO2, Minor_Dia_Obs_3=@MinorO3, " &
                    "Major_Dia_Min_Limit=@MajorMin, Major_Dia_Max_Limit=@MajorMax, Major_Dia_Obs_1=@MajorO1, Major_Dia_Obs_2=@MajorO2, Major_Dia_Obs_3=@MajorO3, " &
                    "TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark " &
                    "WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))

                    cmd.Parameters.AddWithValue("@MinorMin", minorDetails(1))
                    cmd.Parameters.AddWithValue("@MinorMax", minorDetails(2))
                    cmd.Parameters.AddWithValue("@MinorO1", minorDetails(3))
                    cmd.Parameters.AddWithValue("@MinorO2", minorDetails(4))
                    cmd.Parameters.AddWithValue("@MinorO3", minorDetails(5))

                    cmd.Parameters.AddWithValue("@MajorMin", majorDetails(1))
                    cmd.Parameters.AddWithValue("@MajorMax", majorDetails(2))
                    cmd.Parameters.AddWithValue("@MajorO1", majorDetails(3))
                    cmd.Parameters.AddWithValue("@MajorO2", majorDetails(4))
                    cmd.Parameters.AddWithValue("@MajorO3", majorDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateSplineRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetStraightPlugGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM straight_plug_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM straight_plug_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetStraightPlugGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetStraightPlugGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM straight_plug_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertStraightPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, minLimit As Object, maxLimit As Object, obs1 As String, obs2 As String, obs3 As String, obs4 As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO straight_plug_gauge_calibration (" &
                    "RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, " &
                    "Parameter, Location_Obs, Min_Limit, Max_Limit, Obs_1, Obs_2, Obs_3, Obs_4, " &
                    "TimeIn, TimeOut, TotalTime, Status, Remark) VALUES (" &
                    "'RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, " &
                    "@Parameter, @LocationObs, @MinLimit, @MaxLimit, @Obs1, @Obs2, @Obs3, @Obs4, " &
                    "@TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@LocationObs", locationObs)
                    cmd.Parameters.AddWithValue("@MinLimit", If(minLimit Is Nothing, DBNull.Value, minLimit))
                    cmd.Parameters.AddWithValue("@MaxLimit", If(maxLimit Is Nothing, DBNull.Value, maxLimit))
                    cmd.Parameters.AddWithValue("@Obs1", obs1)
                    cmd.Parameters.AddWithValue("@Obs2", obs2)
                    cmd.Parameters.AddWithValue("@Obs3", obs3)
                    cmd.Parameters.AddWithValue("@Obs4", obs4)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertStraightPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateStraightPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, minLimit As Object, maxLimit As Object, obs1 As String, obs2 As String, obs3 As String, obs4 As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE straight_plug_gauge_calibration SET " &
                    "Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, " &
                    "Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, " &
                    "Parameter=@Parameter, Location_Obs=@LocationObs, Min_Limit=@MinLimit, Max_Limit=@MaxLimit, " &
                    "Obs_1=@Obs1, Obs_2=@Obs2, Obs_3=@Obs3, Obs_4=@Obs4, " &
                    "TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark " &
                    "WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@LocationObs", locationObs)
                    cmd.Parameters.AddWithValue("@MinLimit", If(minLimit Is Nothing, DBNull.Value, minLimit))
                    cmd.Parameters.AddWithValue("@MaxLimit", If(maxLimit Is Nothing, DBNull.Value, maxLimit))
                    cmd.Parameters.AddWithValue("@Obs1", obs1)
                    cmd.Parameters.AddWithValue("@Obs2", obs2)
                    cmd.Parameters.AddWithValue("@Obs3", obs3)
                    cmd.Parameters.AddWithValue("@Obs4", obs4)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateStraightPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetThreadRingGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim dt = ReadDatatable($"SELECT * FROM thread_ring_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM thread_ring_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetThreadRingGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetThreadRingGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM thread_ring_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertThreadRingGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, minLimit As Object, maxLimit As Object, obs1 As String, obs2 As String, obs3 As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_ring_gauge_calibration (" &
                    "RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, " &
                    "Parameter, Location_Obs, Min_Limit, Max_Limit, Obs_1, Obs_2, Obs_3, " &
                    "TimeIn, TimeOut, TotalTime, Status, Remark) VALUES (" &
                    "'RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, " &
                    "@Parameter, @LocationObs, @MinLimit, @MaxLimit, @Obs1, @Obs2, @Obs3, " &
                    "@TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@LocationObs", locationObs)
                    cmd.Parameters.AddWithValue("@MinLimit", If(minLimit Is Nothing, DBNull.Value, minLimit))
                    cmd.Parameters.AddWithValue("@MaxLimit", If(maxLimit Is Nothing, DBNull.Value, maxLimit))
                    cmd.Parameters.AddWithValue("@Obs1", obs1)
                    cmd.Parameters.AddWithValue("@Obs2", obs2)
                    cmd.Parameters.AddWithValue("@Obs3", obs3)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertThreadRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateThreadRingGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, minLimit As Object, maxLimit As Object, obs1 As String, obs2 As String, obs3 As String, timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE thread_ring_gauge_calibration SET " &
                    "Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, " &
                    "Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, " &
                    "Parameter=@Parameter, Location_Obs=@LocationObs, Min_Limit=@MinLimit, Max_Limit=@MaxLimit, " &
                    "Obs_1=@Obs1, Obs_2=@Obs2, Obs_3=@Obs3, " &
                    "TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark " &
                    "WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@LocationObs", locationObs)
                    cmd.Parameters.AddWithValue("@MinLimit", If(minLimit Is Nothing, DBNull.Value, minLimit))
                    cmd.Parameters.AddWithValue("@MaxLimit", If(maxLimit Is Nothing, DBNull.Value, maxLimit))
                    cmd.Parameters.AddWithValue("@Obs1", obs1)
                    cmd.Parameters.AddWithValue("@Obs2", obs2)
                    cmd.Parameters.AddWithValue("@Obs3", obs3)
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateThreadRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetThreadPlugGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM thread_plug_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            ' Fallback if specific LC not found
            dt = ReadDatatable("SELECT * FROM thread_plug_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetThreadPlugGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetThreadPlugGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM thread_plug_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertThreadPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_plug_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Parameter, Location_Obs, Go_Min_Limit, Go_Max_Limit, Go_Obs_1, Go_Obs_2, Go_Obs_3, Go_Obs_4, NoGo_Min_Limit, NoGo_Max_Limit, NoGo_Obs_1, NoGo_Obs_2, NoGo_Obs_3, NoGo_Obs_4, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Parameter, @Location_Obs, @Go_Min_Limit, @Go_Max_Limit, @Go_Obs_1, @Go_Obs_2, @Go_Obs_3, @Go_Obs_4, @NoGo_Min_Limit, @NoGo_Max_Limit, @NoGo_Obs_1, @NoGo_Obs_2, @NoGo_Obs_3, @NoGo_Obs_4, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertThreadPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateThreadPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, goDetails As Object(), noGoDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE thread_plug_gauge_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Parameter=@Parameter, Location_Obs=@Location_Obs, Go_Min_Limit=@Go_Min_Limit, Go_Max_Limit=@Go_Max_Limit, Go_Obs_1=@Go_Obs_1, Go_Obs_2=@Go_Obs_2, Go_Obs_3=@Go_Obs_3, Go_Obs_4=@Go_Obs_4, NoGo_Min_Limit=@NoGo_Min_Limit, NoGo_Max_Limit=@NoGo_Max_Limit, NoGo_Obs_1=@NoGo_Obs_1, NoGo_Obs_2=@NoGo_Obs_2, NoGo_Obs_3=@NoGo_Obs_3, NoGo_Obs_4=@NoGo_Obs_4, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)

                    cmd.Parameters.AddWithValue("@Go_Min_Limit", goDetails(0))
                    cmd.Parameters.AddWithValue("@Go_Max_Limit", goDetails(1))
                    cmd.Parameters.AddWithValue("@Go_Obs_1", goDetails(2))
                    cmd.Parameters.AddWithValue("@Go_Obs_2", goDetails(3))
                    cmd.Parameters.AddWithValue("@Go_Obs_3", goDetails(4))
                    cmd.Parameters.AddWithValue("@Go_Obs_4", goDetails(5))

                    cmd.Parameters.AddWithValue("@NoGo_Min_Limit", noGoDetails(0))
                    cmd.Parameters.AddWithValue("@NoGo_Max_Limit", noGoDetails(1))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_1", noGoDetails(2))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_2", noGoDetails(3))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_3", noGoDetails(4))
                    cmd.Parameters.AddWithValue("@NoGo_Obs_4", noGoDetails(5))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateThreadPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetTaperPlugGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM taper_plug_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            dt = ReadDatatable("SELECT * FROM taper_plug_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetTaperPlugGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetTaperPlugGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM taper_plug_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertTaperPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, taperAngleDegree As String, majorDetails As Object(), angleDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO taper_plug_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Parameter, Location_Obs, Taper_Angle_Degree, Major_Dia_Min_Limit, Major_Dia_Max_Limit, Major_Dia_Obs_1, Major_Dia_Obs_2, Major_Dia_Obs_3, Angle_Sec_Min_Limit, Angle_Sec_Max_Limit, Angle_Sec_Obs_1, Angle_Sec_Obs_2, Angle_Sec_Obs_3, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Parameter, @Location_Obs, @Taper_Angle_Degree, @Major_Dia_Min_Limit, @Major_Dia_Max_Limit, @Major_Dia_Obs_1, @Major_Dia_Obs_2, @Major_Dia_Obs_3, @Angle_Sec_Min_Limit, @Angle_Sec_Max_Limit, @Angle_Sec_Obs_1, @Angle_Sec_Obs_2, @Angle_Sec_Obs_3, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)
                    cmd.Parameters.AddWithValue("@Taper_Angle_Degree", taperAngleDegree)

                    cmd.Parameters.AddWithValue("@Major_Dia_Min_Limit", majorDetails(0))
                    cmd.Parameters.AddWithValue("@Major_Dia_Max_Limit", majorDetails(1))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_1", majorDetails(2))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_2", majorDetails(3))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_3", majorDetails(4))

                    cmd.Parameters.AddWithValue("@Angle_Sec_Min_Limit", angleDetails(0))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Max_Limit", angleDetails(1))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_1", angleDetails(2))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_2", angleDetails(3))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_3", angleDetails(4))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertTaperPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateTaperPlugGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, taperAngleDegree As String, majorDetails As Object(), angleDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE taper_plug_gauge_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Parameter=@Parameter, Location_Obs=@Location_Obs, Taper_Angle_Degree=@Taper_Angle_Degree, Major_Dia_Min_Limit=@Major_Dia_Min_Limit, Major_Dia_Max_Limit=@Major_Dia_Max_Limit, Major_Dia_Obs_1=@Major_Dia_Obs_1, Major_Dia_Obs_2=@Major_Dia_Obs_2, Major_Dia_Obs_3=@Major_Dia_Obs_3, Angle_Sec_Min_Limit=@Angle_Sec_Min_Limit, Angle_Sec_Max_Limit=@Angle_Sec_Max_Limit, Angle_Sec_Obs_1=@Angle_Sec_Obs_1, Angle_Sec_Obs_2=@Angle_Sec_Obs_2, Angle_Sec_Obs_3=@Angle_Sec_Obs_3, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)
                    cmd.Parameters.AddWithValue("@Taper_Angle_Degree", taperAngleDegree)

                    cmd.Parameters.AddWithValue("@Major_Dia_Min_Limit", majorDetails(0))
                    cmd.Parameters.AddWithValue("@Major_Dia_Max_Limit", majorDetails(1))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_1", majorDetails(2))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_2", majorDetails(3))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_3", majorDetails(4))

                    cmd.Parameters.AddWithValue("@Angle_Sec_Min_Limit", angleDetails(0))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Max_Limit", angleDetails(1))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_1", angleDetails(2))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_2", angleDetails(3))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_3", angleDetails(4))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateTaperPlugGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetTaperRingGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM taper_ring_gauge_calibration WHERE RowType='MASTER' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            dt = ReadDatatable("SELECT * FROM taper_ring_gauge_calibration WHERE RowType='MASTER' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetTaperRingGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetTaperRingGaugeCalibrationData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM taper_ring_gauge_calibration WHERE RowType='RECORD' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id DESC LIMIT 1")
    End Function

    Public Function InsertTaperRingGaugeCalibration(type As String, controlNo As String, cycleName As String, calibDate As DateTime?, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, taperAngleDegree As String, majorDetails As Object(), angleDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO taper_ring_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Parameter, Location_Obs, Taper_Angle_Degree, Major_Dia_Min_Limit, Major_Dia_Max_Limit, Major_Dia_Obs_1, Major_Dia_Obs_2, Major_Dia_Obs_3, Angle_Sec_Min_Limit, Angle_Sec_Max_Limit, Angle_Sec_Obs_1, Angle_Sec_Obs_2, Angle_Sec_Obs_3, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Parameter, @Location_Obs, @Taper_Angle_Degree, @Major_Dia_Min_Limit, @Major_Dia_Max_Limit, @Major_Dia_Obs_1, @Major_Dia_Obs_2, @Major_Dia_Obs_3, @Angle_Sec_Min_Limit, @Angle_Sec_Max_Limit, @Angle_Sec_Obs_1, @Angle_Sec_Obs_2, @Angle_Sec_Obs_3, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)
                    cmd.Parameters.AddWithValue("@Taper_Angle_Degree", taperAngleDegree)

                    cmd.Parameters.AddWithValue("@Major_Dia_Min_Limit", majorDetails(0))
                    cmd.Parameters.AddWithValue("@Major_Dia_Max_Limit", majorDetails(1))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_1", majorDetails(2))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_2", majorDetails(3))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_3", majorDetails(4))

                    cmd.Parameters.AddWithValue("@Angle_Sec_Min_Limit", angleDetails(0))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Max_Limit", angleDetails(1))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_1", angleDetails(2))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_2", angleDetails(3))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_3", angleDetails(4))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertTaperRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateTaperRingGaugeCalibration(type As String, controlNo As String, cycleName As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, parameter As String, locationObs As String, taperAngleDegree As String, majorDetails As Object(), angleDetails As Object(), timeIn As String, timeOut As String, totalTime As String, status As String, remark As String, calibDate As DateTime?) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "UPDATE taper_ring_gauge_calibration SET Date=@Date, Time=@Time, Type=@Type, Size=@Size, LC=@LC, Color=@Color, Location=@Location, Temperature=@Temperature, Humidity=@Humidity, TMU=@TMU, MU=@MU, Parameter=@Parameter, Location_Obs=@Location_Obs, Taper_Angle_Degree=@Taper_Angle_Degree, Major_Dia_Min_Limit=@Major_Dia_Min_Limit, Major_Dia_Max_Limit=@Major_Dia_Max_Limit, Major_Dia_Obs_1=@Major_Dia_Obs_1, Major_Dia_Obs_2=@Major_Dia_Obs_2, Major_Dia_Obs_3=@Major_Dia_Obs_3, Angle_Sec_Min_Limit=@Angle_Sec_Min_Limit, Angle_Sec_Max_Limit=@Angle_Sec_Max_Limit, Angle_Sec_Obs_1=@Angle_Sec_Obs_1, Angle_Sec_Obs_2=@Angle_Sec_Obs_2, Angle_Sec_Obs_3=@Angle_Sec_Obs_3, TimeIn=@TimeIn, TimeOut=@TimeOut, TotalTime=@TotalTime, Status=@Status, Remark=@Remark WHERE RowType='RECORD' AND ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Parameter", parameter)
                    cmd.Parameters.AddWithValue("@Location_Obs", locationObs)
                    cmd.Parameters.AddWithValue("@Taper_Angle_Degree", taperAngleDegree)

                    cmd.Parameters.AddWithValue("@Major_Dia_Min_Limit", majorDetails(0))
                    cmd.Parameters.AddWithValue("@Major_Dia_Max_Limit", majorDetails(1))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_1", majorDetails(2))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_2", majorDetails(3))
                    cmd.Parameters.AddWithValue("@Major_Dia_Obs_3", majorDetails(4))

                    cmd.Parameters.AddWithValue("@Angle_Sec_Min_Limit", angleDetails(0))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Max_Limit", angleDetails(1))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_1", angleDetails(2))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_2", angleDetails(3))
                    cmd.Parameters.AddWithValue("@Angle_Sec_Obs_3", angleDetails(4))

                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateTaperRingGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetFeelerGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM feeler_gauge_calibration WHERE RowType='MASTER' AND ControlNo LIKE 'LIMITS_%' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            dt = ReadDatatable("SELECT * FROM feeler_gauge_calibration WHERE RowType='MASTER' AND ControlNo LIKE 'LIMITS_%' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetFeelerGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetFeelerGaugeMasterData(controlNo As String, cycleName As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM feeler_gauge_calibration WHERE (RowType='RECORD' OR RowType='MASTER') AND (Measurement_Range IS NULL OR Measurement_Range = '') AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetFeelerGaugeMasterData: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetFeelerGaugeRecordData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM feeler_gauge_calibration WHERE RowType='RECORD' AND Measurement_Range IS NOT NULL AND Measurement_Range <> '' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id ASC")
    End Function

    Public Function DeleteFeelerGaugeCalibration(controlNo As String, cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "DELETE FROM feeler_gauge_calibration WHERE ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in DeleteFeelerGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertFeelerGaugeMaster(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, tempMin As Decimal?, tempMax As Decimal?, humMin As Decimal?, humMax As Decimal?, judgement As String, remark As String, createdBy As String, timeIn As String, timeOut As String, totalTime As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO feeler_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, TimeIn, TimeOut, TotalTime, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Env_Temp_LL, @Env_Temp_UL, @Env_Hum_LL, @Env_Hum_UL, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Env_Temp_LL", If(tempMin.HasValue, tempMin.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Temp_UL", If(tempMax.HasValue, tempMax.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Hum_LL", If(humMin.HasValue, humMin.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Hum_UL", If(humMax.HasValue, humMax.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", judgement)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertFeelerGaugeMaster: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertFeelerGaugeRecord(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, measurementRange As String, minLimit As Decimal?, maxLimit As Decimal?, t1 As Decimal?, t2 As Decimal?, t3 As Decimal?, c1 As Decimal?, c2 As Decimal?, c3 As Decimal?, judgement As String, remark As String, createdBy As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO feeler_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Measurement_Range, Min_Limit, Max_Limit, Obs_1, Obs_2, Obs_3, Camber_Width_Obs_1, Camber_Width_Obs_2, Camber_Width_Obs_3, Status, Remark) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Measurement_Range, @Min_Limit, @Max_Limit, @Obs_1, @Obs_2, @Obs_3, @Camber_Width_Obs_1, @Camber_Width_Obs_2, @Camber_Width_Obs_3, @Status, @Remark)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Measurement_Range", measurementRange)
                    cmd.Parameters.AddWithValue("@Min_Limit", If(minLimit.HasValue, minLimit.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Max_Limit", If(maxLimit.HasValue, maxLimit.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Obs_1", If(t1.HasValue, t1.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Obs_2", If(t2.HasValue, t2.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Obs_3", If(t3.HasValue, t3.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Camber_Width_Obs_1", If(c1.HasValue, c1.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Camber_Width_Obs_2", If(c2.HasValue, c2.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Camber_Width_Obs_3", If(c3.HasValue, c3.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Status", judgement)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertFeelerGaugeRecord: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetRadiusGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM radius_gauge_calibration WHERE RowType='MASTER' AND ControlNo LIKE 'LIMITS_%' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            dt = ReadDatatable("SELECT * FROM radius_gauge_calibration WHERE RowType='MASTER' AND ControlNo LIKE 'LIMITS_%' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetRadiusGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetRadiusGaugeMasterData(controlNo As String, cycleName As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM radius_gauge_calibration WHERE (RowType='RECORD' OR RowType='MASTER') AND (Measurement_Range IS NULL OR Measurement_Range = '') AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetRadiusGaugeMasterData: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetRadiusGaugeRecordData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM radius_gauge_calibration WHERE RowType='RECORD' AND Measurement_Range IS NOT NULL AND Measurement_Range <> '' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id ASC")
    End Function

    Public Function DeleteRadiusGaugeCalibration(controlNo As String, cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "DELETE FROM radius_gauge_calibration WHERE ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in DeleteRadiusGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertRadiusGaugeMaster(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, tempMin As Decimal?, tempMax As Decimal?, humMin As Decimal?, humMax As Decimal?, status As String, remark As String, createdBy As String, timeIn As String, timeOut As String, totalTime As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO radius_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, TimeIn, TimeOut, TotalTime, Status, Remark, CreatedBy) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Env_Temp_LL, @Env_Temp_UL, @Env_Hum_LL, @Env_Hum_UL, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark, @CreatedBy)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Env_Temp_LL", If(tempMin.HasValue, tempMin.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Temp_UL", If(tempMax.HasValue, tempMax.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Hum_LL", If(humMin.HasValue, humMin.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Hum_UL", If(humMax.HasValue, humMax.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertRadiusGaugeMaster: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertRadiusGaugeRecord(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, measurementRange As String, minLimit As Decimal?, maxLimit As Decimal?, i1 As Decimal?, i2 As Decimal?, i3 As Decimal?, o1 As Decimal?, o2 As Decimal?, o3 As Decimal?, status As String, remark As String, createdBy As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO radius_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Measurement_Range, Min_Limit, Max_Limit, Inside_Radius_Obs_1, Inside_Radius_Obs_2, Inside_Radius_Obs_3, Outside_Radius_Obs_1, Outside_Radius_Obs_2, Outside_Radius_Obs_3, Status, Remark, CreatedBy) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Measurement_Range, @Min_Limit, @Max_Limit, @Inside_Radius_Obs_1, @Inside_Radius_Obs_2, @Inside_Radius_Obs_3, @Outside_Radius_Obs_1, @Outside_Radius_Obs_2, @Outside_Radius_Obs_3, @Status, @Remark, @CreatedBy)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Measurement_Range", measurementRange)
                    cmd.Parameters.AddWithValue("@Min_Limit", If(minLimit.HasValue, minLimit.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Max_Limit", If(maxLimit.HasValue, maxLimit.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Inside_Radius_Obs_1", If(i1.HasValue, i1.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Inside_Radius_Obs_2", If(i2.HasValue, i2.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Inside_Radius_Obs_3", If(i3.HasValue, i3.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Outside_Radius_Obs_1", If(o1.HasValue, o1.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Outside_Radius_Obs_2", If(o2.HasValue, o2.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Outside_Radius_Obs_3", If(o3.HasValue, o3.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertRadiusGaugeRecord: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function GetThreadPitchGaugeMasterLimits(lc As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM thread_pitch_gauge_calibration WHERE RowType='MASTER' AND ControlNo LIKE 'LIMITS_%' AND LC='{lc.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            dt = ReadDatatable("SELECT * FROM thread_pitch_gauge_calibration WHERE RowType='MASTER' AND ControlNo LIKE 'LIMITS_%' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetThreadPitchGaugeMasterLimits: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetThreadPitchGaugeMasterData(controlNo As String, cycleName As String) As DataRow
        Try
            Dim query As String = $"SELECT * FROM thread_pitch_gauge_calibration WHERE (RowType='RECORD' OR RowType='MASTER') AND (Measurement_Range IS NULL OR Measurement_Range = '') AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' LIMIT 1"
            Dim dt As DataTable = ReadDatatable(query)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in GetThreadPitchGaugeMasterData: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function GetThreadPitchGaugeRecordData(controlNo As String, cycleName As String) As DataTable
        Return ReadDatatable($"SELECT * FROM thread_pitch_gauge_calibration WHERE RowType='RECORD' AND Measurement_Range IS NOT NULL AND Measurement_Range <> '' AND ControlNo='{controlNo.Replace("'", "''")}' AND CycleName='{cycleName.Replace("'", "''")}' ORDER BY id ASC")
    End Function

    Public Function DeleteThreadPitchGaugeCalibration(controlNo As String, cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "DELETE FROM thread_pitch_gauge_calibration WHERE ControlNo=@ControlNo AND CycleName=@CycleName"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in DeleteThreadPitchGaugeCalibration: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertThreadPitchGaugeMaster(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, temp As String, humidity As String, tmu As String, mu As String, tempMin As Decimal?, tempMax As Decimal?, humMin As Decimal?, humMax As Decimal?, status As String, remark As String, createdBy As String, timeIn As String, timeOut As String, totalTime As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_pitch_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Temperature, Humidity, TMU, MU, Env_Temp_LL, Env_Temp_UL, Env_Hum_LL, Env_Hum_UL, TimeIn, TimeOut, TotalTime, Status, Remark, CreatedBy) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Temperature, @Humidity, @TMU, @MU, @Env_Temp_LL, @Env_Temp_UL, @Env_Hum_LL, @Env_Hum_UL, @TimeIn, @TimeOut, @TotalTime, @Status, @Remark, @CreatedBy)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Temperature", temp)
                    cmd.Parameters.AddWithValue("@Humidity", humidity)
                    cmd.Parameters.AddWithValue("@TMU", tmu)
                    cmd.Parameters.AddWithValue("@MU", If(String.IsNullOrWhiteSpace(mu), DBNull.Value, mu))
                    cmd.Parameters.AddWithValue("@Env_Temp_LL", If(tempMin.HasValue, tempMin.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Temp_UL", If(tempMax.HasValue, tempMax.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Hum_LL", If(humMin.HasValue, humMin.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Env_Hum_UL", If(humMax.HasValue, humMax.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@TimeIn", timeIn)
                    cmd.Parameters.AddWithValue("@TimeOut", timeOut)
                    cmd.Parameters.AddWithValue("@TotalTime", totalTime)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertThreadPitchGaugeMaster: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function InsertThreadPitchGaugeRecord(controlNo As String, cycleName As String, calibDate As DateTime?, type As String, size As String, lc As String, color As String, location As String, measurementRange As String, minLimit As Decimal?, maxLimit As Decimal?, p1 As Decimal?, p2 As Decimal?, p3 As Decimal?, a1 As Decimal?, a2 As Decimal?, a3 As Decimal?, status As String, remark As String, createdBy As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim query As String = "INSERT INTO thread_pitch_gauge_calibration (RowType, ControlNo, CycleName, Date, Time, Type, Size, LC, Color, Location, Measurement_Range, Min_Limit, Max_Limit, Pitch_Obs_1, Pitch_Obs_2, Pitch_Obs_3, Angle_Obs_1, Angle_Obs_2, Angle_Obs_3, Status, Remark, CreatedBy) VALUES ('RECORD', @ControlNo, @CycleName, @Date, @Time, @Type, @Size, @LC, @Color, @Location, @Measurement_Range, @Min_Limit, @Max_Limit, @Pitch_Obs_1, @Pitch_Obs_2, @Pitch_Obs_3, @Angle_Obs_1, @Angle_Obs_2, @Angle_Obs_3, @Status, @Remark, @CreatedBy)"
                Using cmd As New MySqlCommand(query, _con)
                    cmd.Parameters.AddWithValue("@ControlNo", controlNo)
                    cmd.Parameters.AddWithValue("@CycleName", cycleName)
                    cmd.Parameters.AddWithValue("@Date", If(calibDate.HasValue, calibDate.Value, DateTime.Today))
                    cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@Type", type)
                    cmd.Parameters.AddWithValue("@Size", size)
                    cmd.Parameters.AddWithValue("@LC", lc)
                    cmd.Parameters.AddWithValue("@Color", color)
                    cmd.Parameters.AddWithValue("@Location", location)
                    cmd.Parameters.AddWithValue("@Measurement_Range", measurementRange)
                    cmd.Parameters.AddWithValue("@Min_Limit", If(minLimit.HasValue, minLimit.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Max_Limit", If(maxLimit.HasValue, maxLimit.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Pitch_Obs_1", If(p1.HasValue, p1.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Pitch_Obs_2", If(p2.HasValue, p2.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Pitch_Obs_3", If(p3.HasValue, p3.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Angle_Obs_1", If(a1.HasValue, a1.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Angle_Obs_2", If(a2.HasValue, a2.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Angle_Obs_3", If(a3.HasValue, a3.Value.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@Remark", remark)
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy)
                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End If
        Catch ex As Exception
            Console.WriteLine("Error in InsertThreadPitchGaugeRecord: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateLinkedRFID(controlNo As String, cycleName As String, newRfidTag As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                ' 1. Update interchangeability table strictly for RFID
                Dim updateInterchange As String = "UPDATE interchangeability SET RFID_tag = @rfid WHERE ControlNo = @ctrl AND CycleName = @cycle"
                Using cmd As New MySqlCommand(updateInterchange, _con)
                    cmd.Parameters.AddWithValue("@rfid", newRfidTag)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.ExecuteNonQuery()
                End Using

                ' 2. Update the latest receive record for this item to reflect the new RFID
                Dim updateReceive As String = "UPDATE receive SET RFID_tag = @rfid WHERE ControlNo = @ctrl AND CycleName = @cycle ORDER BY ID DESC LIMIT 1"
                Using cmd As New MySqlCommand(updateReceive, _con)
                    cmd.Parameters.AddWithValue("@rfid", newRfidTag)
                    cmd.Parameters.AddWithValue("@ctrl", controlNo)
                    cmd.Parameters.AddWithValue("@cycle", cycleName)
                    cmd.ExecuteNonQuery()
                End Using
                
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine("Error in UpdateLinkedRFID: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function IsControlNoCalibratedInAnyTab(ByVal controlNo As String, ByVal cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim tables() As String = {"regular_calibration", "new_addition_calibration", "reintroduction_calibration", "temp_issuance_calibration"}
                For Each tbl In tables
                    Dim query As String = $"SELECT COUNT(*) FROM {tbl} WHERE control_no = @ctrl AND CycleName = @cycle AND (is_calibrated = 'YES' Or is_calibrated = 'Yes')"
                    Using cmd As New MySqlCommand(query, _con)
                        cmd.Parameters.AddWithValue("@ctrl", controlNo)
                        cmd.Parameters.AddWithValue("@cycle", cycleName)
                        Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        If count > 0 Then
                            Return True
                        End If
                    End Using
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("IsControlNoCalibratedInAnyTab Error: " & ex.Message)
        End Try
        Return False
    End Function
    Public Function IsControlNoCalibrationNGInAnyTab(ByVal controlNo As String, ByVal cycleName As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim tables() As String = {"regular_calibration", "new_addition_calibration", "reintroduction_calibration", "temp_issuance_calibration"}
                For Each tbl In tables
                    Dim query As String = $"SELECT COUNT(*) FROM {tbl} WHERE control_no = @ctrl AND CycleName = @cycle AND (is_calibrated = 'YES' Or is_calibrated = 'Yes') AND UPPER(calibration_status) = 'NG'"
                    Using cmd As New MySqlCommand(query, _con)
                        cmd.Parameters.AddWithValue("@ctrl", controlNo)
                        cmd.Parameters.AddWithValue("@cycle", cycleName)
                        Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        If count > 0 Then
                            Return True
                        End If
                    End Using
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("IsControlNoCalibrationNGInAnyTab Error: " & ex.Message)
        End Try
        Return False
    End Function

    ' ========== BACKUP HELPERS ==========

    ''' <summary>Returns every table name in the connected database.</summary>
    Public Function GetAllTableNames() As List(Of String)
        Dim tables As New List(Of String)()
        Try
            If MySQLDBConnect() = 1 Then
                Dim dt = ReadDatatable("SHOW TABLES")
                For Each row As DataRow In dt.Rows
                    tables.Add(row(0).ToString())
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("GetAllTableNames Error: " & ex.Message)
        End Try
        Return tables
    End Function

    ''' <summary>Exposes the underlying MySqlConnection for use by MySqlScript.</summary>
    Public Function GetConnection() As MySqlConnection
        MySQLDBConnect()
        Return _con
    End Function

    ''' <summary>
    ''' Executes a full multi-statement SQL dump script (e.g. from backup restore).
    ''' Uses MySqlScript which handles delimiter splitting correctly.
    ''' </summary>
    Public Function ExecuteSqlScript(sql As String) As Boolean
        Try
            If MySQLDBConnect() = 1 Then
                Dim script As New MySqlScript(_con, sql)
                script.Execute()
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine("ExecuteSqlScript Error: " & ex.Message)
        End Try
        Return False
    End Function

End Class
