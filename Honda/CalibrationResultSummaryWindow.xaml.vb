Imports System.Data
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Controls
Imports MySql.Data.MySqlClient
Imports ClosedXML.Excel

Partial Class CalibrationResultSummaryWindow
    Private _mySql As New MySQLClass()
    Private _instrumentName As String
    Private _cycleName As String

    ' ─── CONFIG: one sequential observation group (for gauge-type tables) ──
    Private Class ObsGroupConfig
        ''' <summary>Display text shown above the group columns.</summary>
        Public Property GroupHeader As String = ""
        ''' <summary>Full DB column prefix including trailing underscore, e.g. "Go_Obs_", "Obs_", "Wide_Range_".</summary>
        Public Property DbObsPrefix As String = ""
        ''' <summary>Number of trial observations (e.g. 3 or 4).</summary>
        Public Property TrialCount As Integer = 4
        ''' <summary>DB column name for Min Limit (empty if none).</summary>
        Public Property MinLimitCol As String = ""
        ''' <summary>DB column name for Max Limit (empty if none).</summary>
        Public Property MaxLimitCol As String = ""
        ''' <summary>DB column name for Nominal value (optional, e.g. Special Depth Gauge).</summary>
        Public Property NominalCol As String = ""
        ''' <summary>DB column name for Permissible Error (optional).</summary>
        Public Property PermErrCol As String = ""
    End Class

    ' ─── Holds config + data for each tab so export can access them ────────
    Private Class TabData
        Public Property Config As CalibrationTabConfig
        Public Property Table  As DataTable
    End Class

    ' ─── CONFIG: defines how each calibration template tab looks ──────
    Private Class CalibrationTabConfig
        Public Property TabLabel As String = ""
        Public Property Ranges As Decimal() = {}
        Public Property TrialCount As Integer = 3
        ''' <summary>True for Ext+Int range-based layout (e.g. Vernier Caliper 300/600).</summary>
        Public Property HasExtInt As Boolean = False
        Public Property DbObsPrefix As String = ""
        Public Property DbExtPrefix As String = "ext"
        Public Property DbIntPrefix As String = "int"
        Public Property GroupHeaderText As String = "Observation (mm) at Different Ranges*"
        Public Property ExtGroupHeader As String = "External error(t) (mm)  at Different Ranges*"
        Public Property IntGroupHeader As String = "Internal error(t) (mm)  at Different Ranges*"
        Public Property RangeToDbId As New Dictionary(Of Decimal, String)
        ' ── Sequential-obs properties (gauge-type tables) ──
        ''' <summary>True when the table uses sequential obs columns (Go_Obs_N, Obs_N, Wide_Range_N, etc.).</summary>
        Public Property IsSequentialObs As Boolean = False
        ''' <summary>Ordered list of observation groups detected from the DB schema.</summary>
        Public Property ObsGroups As New List(Of ObsGroupConfig)()
    End Class

    Public Sub New(instrumentName As String, cycleName As String)
        InitializeComponent()
        _instrumentName = instrumentName
        _cycleName = cycleName

        TxtTitle.Text = instrumentName
        LoadAllTabs()
    End Sub



    ''' <summary>Returns a human-readable group header from the raw DB prefix (e.g. "Go_Obs" → "Go Gauge Obs. (mm)").</summary>
    Private Function GetObsGroupHeaderText(grpPrefix As String) As String
        Select Case grpPrefix.ToUpper()
            Case "GO_OBS" : Return "Go Gauge Obs. (mm)"
            Case "NOGO_OBS" : Return "NoGo Gauge Obs. (mm)"
            Case "OBS" : Return "Observations (mm)"
            Case "MINOR_DIA_OBS" : Return "Minor Dia. Obs. (mm)"
            Case "MAJOR_DIA_OBS" : Return "Major Dia. Obs. (mm)"
            Case "ANGLE_SEC_OBS" : Return "Angle Obs. (arc-sec)"
            Case "SIZE_1_OBS" : Return "Size 1 Obs. (mm)"
            Case "SIZE_2_OBS" : Return "Size 2 Obs. (mm)"
            Case "SIZE_3_OBS" : Return "Size 3 Obs. (mm)"
            Case "CAMBER_WIDTH_OBS" : Return "Camber Width Obs."
            Case "INSIDE_RADIUS_OBS" : Return "Inside Radius Obs."
            Case "OUTSIDE_RADIUS_OBS" : Return "Outside Radius Obs."
            Case "PITCH_OBS" : Return "Pitch Obs. (mm)"
            Case "ANGLE_OBS" : Return "Angle Obs. (°)"
            Case "DIA_1_OBS" : Return "Dia 1 Obs. (mm)"
            Case "DIA_2_OBS" : Return "Dia 2 Obs. (mm)"
            Case "DISTANCE_OBS" : Return "Distance Obs. (mm)"
            Case "ANGLE_1_OBS" : Return "Angle 1 Obs. (°)"
            Case "ANGLE_2_OBS" : Return "Angle 2 Obs. (°)"
            Case "WIDE_RANGE" : Return "Wide Range"
            Case "ADJACENT_ERROR" : Return "Adjacent Error"
            Case "REPEATABILITY" : Return "Repeatability"
            Case "FLATNESS_ANVIL" : Return "Flatness - Anvil"
            Case "FLATNESS_SPINDLE" : Return "Flatness - Spindle"
            Case "PARALLELISM" : Return "Parallelism"
            Case Else
                ' Auto: replace underscore with space and apply title case
                Dim parts = grpPrefix.Split("_"c)
                Return String.Join(" ", parts.Select(Function(p) If(p.Length > 0, Char.ToUpper(p(0)) & If(p.Length > 1, p.Substring(1).ToLower(), ""), "")))
        End Select
    End Function

    ''' <summary>
    ''' Three-pass auto-detection of column layout from INFORMATION_SCHEMA.
    ''' Pass 1: Range-based Ext/Int/Obs (instruments like Vernier Caliper).
    ''' Pass 2: Sequential _Obs_N grouping (Go_Obs_1, Minor_Dia_Obs_1 etc.).
    ''' Pass 3: Generic _N grouping (Wide_Range_1, Plus_10_1 etc.) for bore/dial/passameter.
    ''' </summary>
    Private Function AutoDetectConfig(tableName As String, formName As String) As CalibrationTabConfig
        Dim config As New CalibrationTabConfig()

        Try
            Dim dtCols = _mySql.ReadDatatable(
                "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " &
                "WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = @tbl " &
                "ORDER BY ORDINAL_POSITION",
                {New MySqlParameter("@tbl", tableName)})

            If dtCols Is Nothing OrElse dtCols.Rows.Count = 0 Then
                config.TabLabel = formName
                Return config
            End If

            Dim allCols = dtCols.Rows.Cast(Of DataRow).Select(Function(r) r("COLUMN_NAME").ToString()).ToList()

            ' ── PASS 1: Range-based Ext/Int/Obs (existing logic) ────────────
            Dim rx1 As New Regex("^(Ext|Int|Obs|v)([\d_.]+)_(\d+)$", RegexOptions.IgnoreCase)
            Dim hasExt As Boolean = False, hasInt As Boolean = False
            Dim obsPrefix As String = "", detectedExtPrefix As String = "Ext", detectedIntPrefix As String = "Int"
            Dim ranges As New SortedSet(Of Decimal)
            Dim maxTrial As Integer = 1

            For Each colName In allCols
                Dim m = rx1.Match(colName)
                If m.Success Then
                    Dim pfx = m.Groups(1).Value
                    Dim rangeVal = Decimal.Parse(m.Groups(2).Value.Replace("_", "."), System.Globalization.CultureInfo.InvariantCulture)
                    Dim trial = Integer.Parse(m.Groups(3).Value)
                    ranges.Add(rangeVal)
                    maxTrial = Math.Max(maxTrial, trial)
                    If pfx.Equals("Ext", StringComparison.OrdinalIgnoreCase) Then
                        hasExt = True : detectedExtPrefix = pfx
                    ElseIf pfx.Equals("Int", StringComparison.OrdinalIgnoreCase) Then
                        hasInt = True : detectedIntPrefix = pfx
                    Else
                        obsPrefix = pfx
                    End If
                End If
            Next

            If ranges.Count > 0 Then
                ' Range-based instrument detected
                config.Ranges = ranges.ToArray()
                config.TrialCount = maxTrial
                Dim rangeMatchedCols As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each colName In allCols
                    Dim m = rx1.Match(colName)
                    If m.Success Then
                        Dim rangeStr = m.Groups(2).Value
                        Dim rangeVal = Decimal.Parse(rangeStr.Replace("_", "."), System.Globalization.CultureInfo.InvariantCulture)
                        If Not config.RangeToDbId.ContainsKey(rangeVal) Then config.RangeToDbId(rangeVal) = rangeStr
                        rangeMatchedCols.Add(colName)
                    End If
                Next
                config.TabLabel = "0-" & ranges.Max()
                If hasExt AndAlso hasInt Then
                    config.HasExtInt = True
                    config.DbExtPrefix = detectedExtPrefix : config.DbIntPrefix = detectedIntPrefix
                    config.ExtGroupHeader = "External error(t) (mm)  at Different Ranges*"
                    config.IntGroupHeader = "Internal error(t) (mm)  at Different Ranges*"
                Else
                    config.HasExtInt = False
                    config.DbObsPrefix = If(obsPrefix <> "", obsPrefix, "Obs")
                    config.GroupHeaderText = "Observation (mm) at Different Ranges*"
                End If

                ' Also scan non-range columns for geometric factor groups (Flatness_Anvil, Parallelism, etc.)
                Dim skipCols As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "id", "RowType", "ControlNo", "CycleName", "Date", "Time", "Type",
                    "Size", "LC", "Color", "Location", "Temperature", "Humidity", "TMU", "MU",
                    "TimeIn", "TimeOut", "TotalTime", "Status", "Remark", "DepthError",
                    "CreatedAt", "uploaded_doc"
                }
                Dim rxGeo As New Regex("^([A-Za-z][A-Za-z0-9_]+?)_([1-6])$")
                Dim geoGroups As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                For Each colName In allCols
                    If rangeMatchedCols.Contains(colName) Then Continue For
                    If skipCols.Contains(colName) Then Continue For
                    If colName.IndexOf("Limit", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                    If colName.IndexOf("Nominal", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                    Dim mG = rxGeo.Match(colName)
                    If mG.Success Then
                        Dim pfx = mG.Groups(1).Value
                        Dim t = Integer.Parse(mG.Groups(2).Value)
                        If Not geoGroups.ContainsKey(pfx) Then geoGroups(pfx) = 0
                        geoGroups(pfx) = Math.Max(geoGroups(pfx), t)
                    End If
                Next
                ' Add geo groups in column-order
                Dim orderedGeo = allCols _
                    .Select(Function(c) If(rxGeo.Match(c).Success AndAlso geoGroups.ContainsKey(rxGeo.Match(c).Groups(1).Value) AndAlso Not rangeMatchedCols.Contains(c), rxGeo.Match(c).Groups(1).Value, Nothing)) _
                    .Where(Function(p) p IsNot Nothing) _
                    .Distinct(StringComparer.OrdinalIgnoreCase) _
                    .ToList()
                For Each pfx In orderedGeo
                    Dim grp As New ObsGroupConfig()
                    grp.DbObsPrefix = pfx & "_"
                    grp.TrialCount = geoGroups(pfx)
                    grp.GroupHeader = GetObsGroupHeaderText(pfx)
                    config.ObsGroups.Add(grp)
                Next

                Return config
            End If

            ' ── PASS 2: Sequential _Obs_N grouping ──────────────────────────
            ' Matches: Go_Obs_1, NoGo_Obs_1, Obs_1, Minor_Dia_Obs_1, Wide_Obs_1 etc.
            Dim rxObs As New Regex("^(.+)_Obs_(\d+)$", RegexOptions.IgnoreCase)
            ' Also match bare "Obs_N" (plain_ring_gauge etc.)
            Dim rxBarObs As New Regex("^Obs_(\d+)$", RegexOptions.IgnoreCase)

            Dim obsGroups2 As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) ' prefix → maxTrial

            For Each colName In allCols
                Dim mBar = rxBarObs.Match(colName)
                If mBar.Success Then
                    Dim t = Integer.Parse(mBar.Groups(1).Value)
                    If Not obsGroups2.ContainsKey("Obs") Then obsGroups2("Obs") = 0
                    obsGroups2("Obs") = Math.Max(obsGroups2("Obs"), t)
                    Continue For
                End If
                Dim m2 = rxObs.Match(colName)
                If m2.Success Then
                    Dim pfx2 = m2.Groups(1).Value
                    Dim t = Integer.Parse(m2.Groups(2).Value)
                    If Not obsGroups2.ContainsKey(pfx2) Then obsGroups2(pfx2) = 0
                    obsGroups2(pfx2) = Math.Max(obsGroups2(pfx2), t)
                End If
            Next

            If obsGroups2.Count > 0 Then
                config.IsSequentialObs = True
                config.TabLabel = formName
                ' Sort groups by first appearance in column list
                Dim orderedPrefixes = allCols _
                    .Select(Function(c)
                                Dim mB = rxBarObs.Match(c)
                                If mB.Success Then Return "Obs"
                                Dim m2 = rxObs.Match(c)
                                If m2.Success Then Return m2.Groups(1).Value
                                Return Nothing
                            End Function) _
                    .Where(Function(p) p IsNot Nothing AndAlso obsGroups2.ContainsKey(p)) _
                    .Distinct(StringComparer.OrdinalIgnoreCase) _
                    .ToList()

                For Each grpPfx In orderedPrefixes
                    Dim grp As New ObsGroupConfig()
                    grp.DbObsPrefix = grpPfx & "_Obs_"    ' e.g. "Go_Obs_", "Obs_"
                    If grpPfx.Equals("Obs", StringComparison.OrdinalIgnoreCase) Then
                        grp.DbObsPrefix = "Obs_"
                    End If
                    grp.TrialCount = obsGroups2(grpPfx)
                    grp.GroupHeader = GetObsGroupHeaderText(grpPfx)

                    ' Determine limit/nominal/permerr columns
                    Dim basePfx = If(grpPfx.EndsWith("_Obs", StringComparison.OrdinalIgnoreCase),
                                     grpPfx.Substring(0, grpPfx.Length - 4),
                                     If(grpPfx.Equals("Obs", StringComparison.OrdinalIgnoreCase), "", grpPfx))
                    Dim minCol = If(basePfx = "", "Min_Limit", basePfx & "_Min_Limit")
                    Dim maxCol = If(basePfx = "", "Max_Limit", basePfx & "_Max_Limit")
                    Dim nomCol = If(basePfx = "", "Nominal", basePfx & "_Nominal")
                    Dim errCol = If(basePfx = "", "Permissible_Error", basePfx & "_Permissible_Error")

                    If allCols.Any(Function(c) c.Equals(minCol, StringComparison.OrdinalIgnoreCase)) Then grp.MinLimitCol = minCol
                    If allCols.Any(Function(c) c.Equals(maxCol, StringComparison.OrdinalIgnoreCase)) Then grp.MaxLimitCol = maxCol
                    If allCols.Any(Function(c) c.Equals(nomCol, StringComparison.OrdinalIgnoreCase)) Then grp.NominalCol = nomCol
                    If allCols.Any(Function(c) c.Equals(errCol, StringComparison.OrdinalIgnoreCase)) Then grp.PermErrCol = errCol

                    config.ObsGroups.Add(grp)
                Next
                Return config
            End If

            ' ── PASS 3: Generic _N grouping (bore_gauge, dial_gauge, passameter) ─
            ' Matches columns ending with _1.._6 that are not meta/limit columns
            Dim rxGen As New Regex("^([A-Za-z][A-Za-z0-9_]+?)_([1-6])$")
            Dim excludedStems As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "id", "RowType", "ControlNo", "CycleName", "Date", "Time", "Type",
                "Size", "LC", "Color", "Location", "Temperature", "Humidity", "TMU", "MU",
                "TimeIn", "TimeOut", "TotalTime", "Status", "Remark", "uploaded_doc",
                "CreatedAt", "Parameter", "Location_Obs", "Measurement_Range",
                "Taper_Angle_Degree", "Go_Size", "NoGo_Size", "DepthError"
            }

            Dim obsGroups3 As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each colName In allCols
                If excludedStems.Contains(colName) Then Continue For
                If colName.IndexOf("Limit", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                If colName.IndexOf("Nominal", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                If colName.IndexOf("Permissible", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                If colName.EndsWith("_LL", StringComparison.OrdinalIgnoreCase) OrElse
                   colName.EndsWith("_UL", StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim m3 = rxGen.Match(colName)
                If m3.Success Then
                    Dim pfx3 = m3.Groups(1).Value
                    Dim t = Integer.Parse(m3.Groups(2).Value)
                    If Not obsGroups3.ContainsKey(pfx3) Then obsGroups3(pfx3) = 0
                    obsGroups3(pfx3) = Math.Max(obsGroups3(pfx3), t)
                End If
            Next

            If obsGroups3.Count > 0 Then
                config.IsSequentialObs = True
                config.TabLabel = formName
                Dim orderedGen = allCols _
                    .Select(Function(c)
                                Dim m3 = rxGen.Match(c)
                                If m3.Success AndAlso obsGroups3.ContainsKey(m3.Groups(1).Value) Then Return m3.Groups(1).Value
                                Return Nothing
                            End Function) _
                    .Where(Function(p) p IsNot Nothing) _
                    .Distinct(StringComparer.OrdinalIgnoreCase) _
                    .ToList()

                For Each grpPfx In orderedGen
                    Dim grp As New ObsGroupConfig()
                    grp.DbObsPrefix = grpPfx & "_"
                    grp.TrialCount = obsGroups3(grpPfx)
                    grp.GroupHeader = GetObsGroupHeaderText(grpPfx)
                    ' No Min/Max limits for generic groups (bore gauge, dial gauge, passameter)
                    config.ObsGroups.Add(grp)
                Next
                Return config
            End If

            ' Fallback: no observations detected
            config.TabLabel = formName

        Catch ex As Exception
            config.TabLabel = formName
            Console.WriteLine("AutoDetectConfig Error: " & ex.Message)
        End Try

        Return config
    End Function





    ' ─── MAIN: Discover form_names and create one tab per form ────────────
    Private Sub LoadAllTabs()
        Try
            ' 1. Find all form_names mapped for this instrument
            Dim dtForms = _mySql.ReadDatatable(
                "SELECT DISTINCT form_name FROM calibrationmapping WHERE type_name = @name",
                {New MySqlParameter("@name", _instrumentName)})

            Dim totalCount As Integer = 0
            Dim diagLog As New System.Text.StringBuilder()
            diagLog.AppendLine($"Instrument: '{_instrumentName}'")
            diagLog.AppendLine($"Cycle: '{_cycleName}'")
            diagLog.AppendLine($"calibrationmapping rows: {If(dtForms Is Nothing, "null", dtForms.Rows.Count.ToString())}")

            If dtForms Is Nothing OrElse dtForms.Rows.Count = 0 Then
                Dim baseTbl = MySQLClass.TypeNameToTableName(_instrumentName)
                diagLog.AppendLine($"baseTbl: '{baseTbl}'")
                If Not String.IsNullOrEmpty(baseTbl) Then
                    ' Require the table to have a RowType column — inventory tables don't have one.
                    Dim exactCheck = _mySql.ReadDatatable(
                        "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS " &
                        "WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = @tbl AND COLUMN_NAME = 'RowType'",
                        {New MySqlParameter("@tbl", baseTbl)})
                    Dim exactExists = (exactCheck IsNot Nothing AndAlso exactCheck.Rows.Count > 0 AndAlso
                                       Convert.ToInt32(exactCheck.Rows(0)(0)) > 0)
                    diagLog.AppendLine($"exactExists: {exactExists}")

                    If exactExists Then
                        totalCount += AddTabForForm(_instrumentName, baseTbl, diagLog)
                    Else
                        Dim patternDt = _mySql.ReadDatatable(
                            "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS " &
                            "WHERE TABLE_SCHEMA = DATABASE() " &
                            "AND TABLE_NAME LIKE @pattern " &
                            "AND COLUMN_NAME = 'RowType' " &
                            "GROUP BY TABLE_NAME ORDER BY TABLE_NAME",
                            {New MySqlParameter("@pattern", baseTbl & "_%")})
                        diagLog.AppendLine($"patternDt rows: {If(patternDt Is Nothing, "null", patternDt.Rows.Count.ToString())}")
                        If patternDt IsNot Nothing Then
                            For Each patRow As DataRow In patternDt.Rows
                                Dim foundTbl = patRow("TABLE_NAME").ToString()
                                Dim displayName = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(
                                    foundTbl.Replace("_"c, " "c))
                                diagLog.AppendLine($"  pattern table: '{foundTbl}'")
                                totalCount += AddTabForForm(displayName, foundTbl, diagLog)
                            Next
                        End If
                    End If
                End If
            Else
                For Each formRow As DataRow In dtForms.Rows
                    Dim formName = formRow("form_name").ToString()
                    Dim tblDetail = MySQLClass.TypeNameToTableName(formName)
                    diagLog.AppendLine($"mapping: form='{formName}' tbl='{tblDetail}'")
                    If String.IsNullOrEmpty(tblDetail) Then Continue For
                    totalCount += AddTabForForm(formName, tblDetail, diagLog)
                Next
            End If

            TxtTotalDone.Text = "Total Calibration Done: " & totalCount

            ' ── TEMPORARY DIAGNOSTIC – remove after confirming data shows ──
            If totalCount = 0 Then
                MessageBox.Show(diagLog.ToString(), "Summary Diagnostic", MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading summary: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Creates a tab for a given form using config-driven rendering. Returns count of records found.
    ''' </summary>
    Private Function AddTabForForm(formName As String, tblDetail As String, Optional diagLog As System.Text.StringBuilder = Nothing) As Integer
        ' First try with cycle filter. If the cycle name stored in the calibration table
        ' differs from the combo selection (e.g. trimming/case/format difference), fall back
        ' to showing all RECORD rows for this instrument so the summary is never blank.
        Dim query = $"SELECT * FROM `{tblDetail}` " &
                    "WHERE RowType = 'RECORD' AND CycleName = @cycle " &
                    "ORDER BY ControlNo ASC"

        Dim params As MySqlParameter() = {
            New MySqlParameter("@cycle", _cycleName)
        }

        Dim dtRaw = _mySql.ReadDatatable(query, params)

        Dim diagStr = If(diagLog IsNot Nothing, diagLog, New System.Text.StringBuilder())
        diagStr.AppendLine($"  [AddTab] tbl='{tblDetail}' cycle='{_cycleName}' filteredRows={If(dtRaw Is Nothing, "null", dtRaw.Rows.Count.ToString())}")

        ' Fallback: if no rows match the cycle filter, show all RECORD rows
        If dtRaw Is Nothing OrElse dtRaw.Rows.Count = 0 Then
            Dim fallbackQuery = $"SELECT * FROM `{tblDetail}` " &
                                "WHERE RowType = 'RECORD' " &
                                "ORDER BY ControlNo ASC"
            dtRaw = _mySql.ReadDatatable(fallbackQuery)
        End If
        diagStr.AppendLine($"  [AddTab] after fallback rows={If(dtRaw Is Nothing, "null", dtRaw.Rows.Count.ToString())}")
        If dtRaw Is Nothing OrElse dtRaw.Rows.Count = 0 Then Return 0

        ' Auto-detect tab type from actual DB table columns
        Dim config = AutoDetectConfig(tblDetail, formName)

        Dim tabItem As New TabItem()
        tabItem.Header = config.TabLabel
        tabItem.FontSize = 13
        tabItem.FontWeight = FontWeights.SemiBold
        tabItem.Tag = config ' Store config for export

        Dim dtFinal As DataTable = Nothing
        tabItem.Content = BuildTabContent(dtRaw, config, dtFinal)
        tabItem.Tag = New TabData With {.Config = config, .Table = dtFinal}
        TabSummary.Items.Add(tabItem)
        If TabSummary.SelectedItem Is Nothing Then TabSummary.SelectedIndex = 0

        Return dtRaw.Rows.Count
    End Function

    ' ═══════════════════════════════════════════════════════════════════════
    '  GENERIC TAB CONTENT BUILDER  (replaces BuildTab300Content / BuildTab600Content)
    ' ═══════════════════════════════════════════════════════════════════════
    Private Function BuildTabContent(dtRaw As DataTable, config As CalibrationTabConfig, ByRef dtOut As DataTable) As UIElement
        ' ── Build DataTable ────────────────────────────────────────────────────
        Dim dtFinal As New DataTable()
        dtFinal.Columns.Add("sno")
        dtFinal.Columns.Add("ControlNo")
        dtFinal.Columns.Add("Size")
        dtFinal.Columns.Add("LC")
        dtFinal.Columns.Add("Color")
        dtFinal.Columns.Add("Location")
        dtFinal.Columns.Add("Date", GetType(DateTime))
        dtFinal.Columns.Add("Temperature")
        dtFinal.Columns.Add("Humidity")
        dtFinal.Columns.Add("TMU")

        If config.IsSequentialObs Then
            ' Sequential (gauges): 1 obs col per group, plus limit/nominal cols on the first trial row
            For gi As Integer = 0 To config.ObsGroups.Count - 1
                Dim grp = config.ObsGroups(gi)
                If grp.MinLimitCol <> "" Then dtFinal.Columns.Add($"SG{gi}_Min")
                If grp.MaxLimitCol <> "" Then dtFinal.Columns.Add($"SG{gi}_Max")
                If grp.NominalCol  <> "" Then dtFinal.Columns.Add($"SG{gi}_Nom")
                If grp.PermErrCol  <> "" Then dtFinal.Columns.Add($"SG{gi}_Err")
                dtFinal.Columns.Add($"SG{gi}")    ' single obs column (value changes per trial row)
            Next
        Else
            ' Range-based (instruments): 1 col per geo group, 1 col per range point
            For gi As Integer = 0 To config.ObsGroups.Count - 1
                dtFinal.Columns.Add($"GG{gi}")
            Next
            For Each r In config.Ranges
                Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                              r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                Dim safeColName = "Col_Obs_" & rStr.Replace(".", "_")
                If config.HasExtInt Then
                    dtFinal.Columns.Add("Ext_" & safeColName)
                    dtFinal.Columns.Add("Int_" & safeColName)
                Else
                    dtFinal.Columns.Add(safeColName)
                End If
            Next
            dtFinal.Columns.Add("DepthError")
        End If

        dtFinal.Columns.Add("TimeIn")
        dtFinal.Columns.Add("TimeOut")
        dtFinal.Columns.Add("TotalTime")
        dtFinal.Columns.Add("Status")
        dtFinal.Columns.Add("Remark")
        dtFinal.Columns.Add("RowColor")
        dtFinal.Columns.Add("RowHeight", GetType(Double))
        dtFinal.Columns.Add("IsSpacer",  GetType(Boolean))

        Dim colorWhite = "White"
        Dim colorGrey  = "#F0F0F0"

        ' Helper: blank all non-meta string cols on non-trial-1 rows
        Dim ClearInfoCols As Action(Of DataRow) =
            Sub(r2)
                r2("sno")         = ""
                r2("ControlNo")   = ""
                r2("Size")        = ""
                r2("LC")          = ""
                r2("Color")       = ""
                r2("Location")    = ""
                r2("Date")        = DBNull.Value
                r2("Temperature") = ""
                r2("Humidity")    = ""
                r2("TMU")         = ""
                r2("TimeIn")      = ""
                r2("TimeOut")     = ""
                r2("TotalTime")   = ""
                r2("Status")      = ""
                r2("Remark")      = ""
            End Sub

        For i As Integer = 0 To dtRaw.Rows.Count - 1
            Dim rawRow      = dtRaw.Rows(i)
            Dim currentColor = If(i Mod 2 = 0, colorWhite, colorGrey)

            If config.IsSequentialObs Then
                ' max trials across all groups
                Dim maxTrials = If(config.ObsGroups.Count > 0, config.ObsGroups.Max(Function(g) g.TrialCount), 1)

                For trial As Integer = 1 To maxTrials
                    Dim fr = dtFinal.NewRow()
                    fr("RowColor")    = currentColor
                    fr("RowHeight")   = 35
                    fr("IsSpacer")    = False

                    If trial = 1 Then
                        fr("sno")         = i + 1
                        fr("ControlNo")   = SafeCol(rawRow, "ControlNo")
                        fr("Size")        = SafeCol(rawRow, "Size")
                        fr("LC")          = SafeCol(rawRow, "LC")
                        fr("Color")       = SafeCol(rawRow, "Color")
                        fr("Location")    = SafeCol(rawRow, "Location")
                        fr("Date")        = If(rawRow.Table.Columns.Contains("Date"), rawRow("Date"), DBNull.Value)
                        fr("Temperature") = SafeCol(rawRow, "Temperature")
                        fr("Humidity")    = SafeCol(rawRow, "Humidity")
                        fr("TMU")         = SafeCol(rawRow, "TMU")
                        fr("TimeIn")      = SafeCol(rawRow, "TimeIn")
                        fr("TimeOut")     = SafeCol(rawRow, "TimeOut")
                        fr("TotalTime")   = SafeCol(rawRow, "TotalTime")
                        fr("Status")      = SafeCol(rawRow, "Status")
                        fr("Remark")      = SafeCol(rawRow, "Remark")
                        ' Limits/nominal/permerr (shown on row 1 only)
                        For gi As Integer = 0 To config.ObsGroups.Count - 1
                            Dim grp = config.ObsGroups(gi)
                            If grp.MinLimitCol <> "" Then
                                Dim v = GetColValue(rawRow, grp.MinLimitCol)
                                fr($"SG{gi}_Min") = If(v IsNot Nothing AndAlso Not IsDBNull(v), v.ToString(), "")
                            End If
                            If grp.MaxLimitCol <> "" Then
                                Dim v = GetColValue(rawRow, grp.MaxLimitCol)
                                fr($"SG{gi}_Max") = If(v IsNot Nothing AndAlso Not IsDBNull(v), v.ToString(), "")
                            End If
                            If grp.NominalCol <> "" Then
                                Dim v = GetColValue(rawRow, grp.NominalCol)
                                fr($"SG{gi}_Nom") = If(v IsNot Nothing AndAlso Not IsDBNull(v), v.ToString(), "")
                            End If
                            If grp.PermErrCol <> "" Then
                                Dim v = GetColValue(rawRow, grp.PermErrCol)
                                fr($"SG{gi}_Err") = If(v IsNot Nothing AndAlso Not IsDBNull(v), v.ToString(), "")
                            End If
                        Next
                    Else
                        ClearInfoCols(fr)
                    End If

                    ' Per-group obs value for this trial
                    For gi As Integer = 0 To config.ObsGroups.Count - 1
                        Dim grp = config.ObsGroups(gi)
                        If trial <= grp.TrialCount Then
                            Dim obsVal = GetColValue(rawRow, grp.DbObsPrefix & trial.ToString())
                            Dim obsStr = If(obsVal IsNot Nothing AndAlso Not IsDBNull(obsVal), obsVal.ToString().Trim(), "")
                            fr($"SG{gi}") = If(obsStr = "", "-", obsStr)
                        Else
                            fr($"SG{gi}") = "-"
                        End If
                    Next
                    dtFinal.Rows.Add(fr)
                Next

            Else   ' ── Range-based ────────────────────────────────────────────
                For trial As Integer = 1 To config.TrialCount
                    Dim fr = dtFinal.NewRow()
                    fr("RowColor")   = currentColor
                    fr("RowHeight")  = 35
                    fr("IsSpacer")   = False

                    If trial = 1 Then
                        fr("sno")         = i + 1
                        fr("ControlNo")   = SafeCol(rawRow, "ControlNo")
                        fr("Size")        = SafeCol(rawRow, "Size")
                        fr("LC")          = SafeCol(rawRow, "LC")
                        fr("Color")       = SafeCol(rawRow, "Color")
                        fr("Location")    = SafeCol(rawRow, "Location")
                        fr("Date")        = If(rawRow.Table.Columns.Contains("Date"), rawRow("Date"), DBNull.Value)
                        fr("Temperature") = SafeCol(rawRow, "Temperature")
                        fr("Humidity")    = SafeCol(rawRow, "Humidity")
                        fr("TMU")         = SafeCol(rawRow, "TMU")
                        fr("TimeIn")      = SafeCol(rawRow, "TimeIn")
                        fr("TimeOut")     = SafeCol(rawRow, "TimeOut")
                        fr("TotalTime")   = SafeCol(rawRow, "TotalTime")
                        fr("Status")      = SafeCol(rawRow, "Status")
                        fr("Remark")      = SafeCol(rawRow, "Remark")
                        Dim deVal = If(rawRow.Table.Columns.Contains("DepthError"), rawRow("DepthError"), DBNull.Value)
                        fr("DepthError") = If(IsDBNull(deVal) OrElse String.IsNullOrWhiteSpace(deVal.ToString()), "-", deVal.ToString())
                    Else
                        ClearInfoCols(fr)
                        fr("DepthError") = ""
                    End If

                    ' Geometric obs groups — 1 col per group, value from this trial
                    For gi As Integer = 0 To config.ObsGroups.Count - 1
                        Dim grp = config.ObsGroups(gi)
                        If trial <= grp.TrialCount Then
                            Dim v = GetColValue(rawRow, grp.DbObsPrefix & trial.ToString())
                            fr($"GG{gi}") = If(v IsNot Nothing AndAlso Not IsDBNull(v), v.ToString(), "-")
                        Else
                            fr($"GG{gi}") = "-"
                        End If
                    Next

                    ' Range obs — 1 col per range, value from this trial
                    For Each r In config.Ranges
                        Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                                      r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                        Dim safeColName = "Col_Obs_" & rStr.Replace(".", "_")
                        If config.HasExtInt Then
                            Dim extVal = GetColValue(rawRow, config.DbExtPrefix & rStr & "_" & trial)
                            Dim intVal = GetColValue(rawRow, config.DbIntPrefix & rStr & "_" & trial)
                            fr("Ext_" & safeColName) = If(extVal IsNot Nothing AndAlso Not IsDBNull(extVal), extVal.ToString(), "-")
                            fr("Int_" & safeColName) = If(intVal IsNot Nothing AndAlso Not IsDBNull(intVal), intVal.ToString(), "-")
                        Else
                            Dim dbVal = GetColValue(rawRow, config.DbObsPrefix & rStr & "_" & trial)
                            fr(safeColName) = If(dbVal IsNot Nothing AndAlso Not IsDBNull(dbVal), dbVal.ToString(), "-")
                        End If
                    Next
                    dtFinal.Rows.Add(fr)
                Next
            End If

            AddGenericSpacerRow(dtFinal, config)
        Next

        ' ── Build UI ───────────────────────────────────────────────────────────
        Dim scroll As New ScrollViewer()
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.VerticalScrollBarVisibility   = ScrollBarVisibility.Auto

        Dim dock As New DockPanel()

        Dim ghGrid = BuildGroupHeader(config)
        DockPanel.SetDock(ghGrid, System.Windows.Controls.Dock.Top)
        dock.Children.Add(ghGrid)

        Dim dg = BuildDataGrid()
        AddColumns(dg, config)
        dg.ItemsSource = dtFinal.DefaultView
        dock.Children.Add(dg)

        AddHandler dg.LayoutUpdated, Sub()
                                         If dg.Columns.Count = ghGrid.ColumnDefinitions.Count Then
                                             For idx = 0 To dg.Columns.Count - 1
                                                 Dim w As New GridLength(dg.Columns(idx).ActualWidth)
                                                 ghGrid.ColumnDefinitions(idx).Width = w
                                             Next
                                         End If
                                     End Sub

        scroll.Content = dock
        dtOut = dtFinal
        Return scroll
    End Function


    ' ═══════════════════════════════════════════════════════════════════════
    '  GENERIC UI BUILDER HELPERS
    ' ═══════════════════════════════════════════════════════════════════════






    Private Function BuildDataGrid() As DataGrid
        Dim dg As New DataGrid()
        dg.AutoGenerateColumns = False
        dg.CanUserAddRows = False
        dg.IsReadOnly = True
        dg.CanUserReorderColumns = False
        dg.CanUserSortColumns = False
        dg.BorderThickness = New Thickness(0)
        dg.Background = Brushes.White
        dg.GridLinesVisibility = DataGridGridLinesVisibility.None
        dg.HeadersVisibility = DataGridHeadersVisibility.Column
        dg.ColumnHeaderStyle = CType(FindResource("SummaryHeaderStyle"), Style)
        dg.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        dg.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
        dg.CellStyle = CType(FindResource("SummaryCellStyle"), Style)

        ' Row style
        Dim rowStyle As New Style(GetType(DataGridRow))
        rowStyle.Setters.Add(New Setter(DataGridRow.BackgroundProperty, New Binding("RowColor")))
        rowStyle.Setters.Add(New Setter(DataGridRow.HeightProperty, New Binding("RowHeight")))
        dg.Resources.Add(GetType(DataGridRow), rowStyle)

        Return dg
    End Function

    ''' <summary>
    ''' Builds the group header row above the DataGrid columns. Layout adapts to config.
    ''' </summary>
    Private Function BuildGroupHeader(config As CalibrationTabConfig) As Grid
        Dim g As New Grid()

        ' Both paths: 10 info cols (S.No, ControlNo, Size, LC, Colour, Location, Date, Temp, Hum, TMU)
        Dim infoWidths = {50, 100, 90, 50, 70, 100, 100, 100, 100, 70}
        For Each w In infoWidths
            g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(w)})
        Next
        Dim infoColCount = 10

        If config.IsSequentialObs Then
            ' Sequential (gauges): limit cols + 1 obs col per group
            For Each grp In config.ObsGroups
                If grp.NominalCol  <> "" Then g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(80)})
                If grp.PermErrCol  <> "" Then g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(80)})
                If grp.MinLimitCol <> "" Then g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(90)})
                If grp.MaxLimitCol <> "" Then g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(90)})
                g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(65)})  ' 1 obs col
            Next
            ' Result cols: TimeIn, TimeOut, TotalTime, Status, Remark
            For Each w In {70, 75, 80, 75, 130}
                g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(w)})
            Next
            ' Group header spans
            AddGroupHeaderCell(g, 0, infoColCount, "")
            Dim colIdx = infoColCount
            For Each grp In config.ObsGroups
                Dim span = If(grp.NominalCol <> "", 1, 0) + If(grp.PermErrCol <> "", 1, 0) +
                           If(grp.MinLimitCol <> "", 1, 0) + If(grp.MaxLimitCol <> "", 1, 0) + 1  ' +1 obs col
                AddGroupHeaderCell(g, colIdx, span, grp.GroupHeader)
                colIdx += span
            Next
            AddGroupHeaderCell(g, colIdx, 5, "")
        Else
            ' Range-based: 1 col per geo group + 1 col per range (Ext side) + 1 col per range (Int side)
            For gi = 0 To config.ObsGroups.Count - 1
                g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(65)})  ' 1 col per geo group
            Next
            If config.HasExtInt Then
                For i = 0 To config.Ranges.Length - 1
                    g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(55)})
                Next
                For i = 0 To config.Ranges.Length - 1
                    g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(55)})
                Next
            Else
                For i = 0 To config.Ranges.Length - 1
                    g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(55)})
                Next
            End If
            ' Result cols: DepthError, TimeIn, TimeOut, TotalTime, Status, Remark
            For Each w In {75, 70, 70, 80, 75, 130}
                g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(w)})
            Next
            ' Draw GG group headers (1 span each)
            AddGroupHeaderCell(g, 0, infoColCount, "")
            Dim ggColIdx = infoColCount
            For gi = 0 To config.ObsGroups.Count - 1
                AddGroupHeaderCell(g, ggColIdx, 1, config.ObsGroups(gi).GroupHeader)
                ggColIdx += 1
            Next
            ' Draw range obs group headers
            Dim extSpan = config.Ranges.Length  ' 1 col per range
            If config.HasExtInt Then
                If config.Ranges.Length > 0 Then
                    AddGroupHeaderCell(g, ggColIdx, extSpan, config.ExtGroupHeader)
                    AddGroupHeaderCell(g, ggColIdx + extSpan, extSpan, config.IntGroupHeader)
                End If
                AddGroupHeaderCell(g, ggColIdx + extSpan * 2, 6, "")
            Else
                If config.Ranges.Length > 0 Then
                    AddGroupHeaderCell(g, ggColIdx, extSpan, config.GroupHeaderText)
                End If
                AddGroupHeaderCell(g, ggColIdx + extSpan, 6, "")
            End If
        End If

        Return g
    End Function

    Private Sub AddGroupHeaderCell(grid As Grid, col As Integer, span As Integer, text As String)
        Dim border As New Border()
        border.Background = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#E8E8E8"), Color))
        border.BorderBrush = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#999999"), Color))
        border.BorderThickness = New Thickness(0, 0, 1, 1)
        Grid.SetColumn(border, col)
        Grid.SetColumnSpan(border, span)

        Dim tb As New TextBlock()
        tb.Text = text
        tb.FontWeight = FontWeights.SemiBold
        tb.FontSize = 10
        tb.HorizontalAlignment = HorizontalAlignment.Center
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.Padding = New Thickness(4)
        border.Child = tb

        grid.Children.Add(border)
    End Sub

    ''' <summary>
    ''' Second header row for range-based instruments showing each range value as a 1-col cell.
    ''' </summary>
    Private Function BuildRangeLabelHeader(config As CalibrationTabConfig) As Grid
        Dim g As New Grid()
        ' 10 info cols
        Dim infoWidths = {50, 100, 90, 50, 70, 100, 100, 100, 100, 70}
        For Each w In infoWidths
            g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(w)})
        Next
        ' 1 col per geo group
        For gi = 0 To config.ObsGroups.Count - 1
            g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(65)})
        Next
        ' 1 col per range (Ext side)
        For i = 0 To config.Ranges.Length - 1
            g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(55)})
        Next
        ' 1 col per range (Int side)
        If config.HasExtInt Then
            For i = 0 To config.Ranges.Length - 1
                g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(55)})
            Next
        End If
        ' Result cols
        For Each w In {75, 70, 70, 80, 75, 130}
            g.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(w)})
        Next

        Dim ggSpanTotal = config.ObsGroups.Count  ' 1 col per geo group
        AddGroupHeaderCell(g, 0, 10 + ggSpanTotal, "")
        Dim colIdx = 10 + ggSpanTotal
        ' Ext range labels — 1 col each
        For Each r In config.Ranges
            AddGroupHeaderCell(g, colIdx, 1, r.ToString("0.##") & " mm")
            colIdx += 1
        Next
        ' Int range labels
        If config.HasExtInt Then
            For Each r In config.Ranges
                AddGroupHeaderCell(g, colIdx, 1, r.ToString("0.##") & " mm")
                colIdx += 1
            Next
        End If
        ' Blank result cols
        AddGroupHeaderCell(g, colIdx, 6, "")
        Return g
    End Function

    ''' <summary>
    ''' Adds DataGrid columns based on config. Works for any calibration template.
    ''' </summary>
    Private Sub AddColumns(dg As DataGrid, config As CalibrationTabConfig)
        ' Fixed info columns (10 total)
        dg.Columns.Add(MakeCol("S.No",       "sno",        50))
        dg.Columns.Add(MakeCol("Control No.", "ControlNo",  100))
        dg.Columns.Add(MakeCol("Range/size",  "Size",       90))
        dg.Columns.Add(MakeCol("L.C",         "LC",         50))
        dg.Columns.Add(MakeCol("Colour",      "Color",      70))
        dg.Columns.Add(MakeCol("Location",    "Location",   100))
        dg.Columns.Add(MakeDateCol("Calibration Date",     "Date",        100))
        dg.Columns.Add(MakeCol("Temperature [20±2°C]",     "Temperature", 100))
        dg.Columns.Add(MakeCol("Humidity [40%-60%] Rh",    "Humidity",    100))
        dg.Columns.Add(MakeCol("T.M.U [± mm]",             "TMU",         70))

        If config.IsSequentialObs Then
            ' Sequential (gauges): 1 obs col per group
            For gi As Integer = 0 To config.ObsGroups.Count - 1
                Dim grp = config.ObsGroups(gi)
                If grp.NominalCol <> "" Then dg.Columns.Add(MakeCol("Nominal",    $"SG{gi}_Nom", 80))
                If grp.PermErrCol <> "" Then dg.Columns.Add(MakeCol("Perm. Err.", $"SG{gi}_Err", 80))
                If grp.MinLimitCol <> "" Then dg.Columns.Add(MakeCol("Min Limit",  $"SG{gi}_Min", 90))
                If grp.MaxLimitCol <> "" Then dg.Columns.Add(MakeCol("Max Limit",  $"SG{gi}_Max", 90))
                dg.Columns.Add(MakeCol(grp.GroupHeader, $"SG{gi}", 65, False))  ' group name as col header
            Next
            dg.Columns.Add(MakeCol("Time In (Hrs)",    "TimeIn",    70))
            dg.Columns.Add(MakeCol("Time Out (Hrs)",   "TimeOut",   75))
            dg.Columns.Add(MakeCol("Total Time (Hrs)", "TotalTime", 80))
            dg.Columns.Add(MakeCol("Status (OK/NG)",   "Status",    75))
            dg.Columns.Add(MakeCol("Remark",           "Remark",   130))
        Else
            ' Range-based: 1 col per geo group, then Ext cols, then Int cols
            For gi As Integer = 0 To config.ObsGroups.Count - 1
                dg.Columns.Add(MakeCol(config.ObsGroups(gi).GroupHeader, $"GG{gi}", 65, False))
            Next
            If config.HasExtInt Then
                ' All External columns first
                For Each r In config.Ranges
                    Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                                  r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                    Dim safeColName = "Col_Obs_" & rStr.Replace(".", "_")
                    dg.Columns.Add(MakeCol(r.ToString("0.##") & " mm", "Ext_" & safeColName, 65, False))
                Next
                ' All Internal columns second
                For Each r In config.Ranges
                    Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                                  r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                    Dim safeColName = "Col_Obs_" & rStr.Replace(".", "_")
                    dg.Columns.Add(MakeCol(r.ToString("0.##") & " mm", "Int_" & safeColName, 65, False))
                Next
            Else
                For Each r In config.Ranges
                    Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                                  r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                    Dim safeColName = "Col_Obs_" & rStr.Replace(".", "_")
                    dg.Columns.Add(MakeCol(r.ToString("0.##") & " mm", safeColName, 65, False))
                Next
            End If
            dg.Columns.Add(MakeCol("Depth Error (mm)", "DepthError", 75))
            dg.Columns.Add(MakeCol("Time In (Hrs)",    "TimeIn",    70))
            dg.Columns.Add(MakeCol("Time Out (Hrs)",   "TimeOut",   70))
            dg.Columns.Add(MakeCol("Total Time (Hrs)", "TotalTime", 80))
            dg.Columns.Add(MakeCol("Status (OK/NG)",   "Status",    75))
            dg.Columns.Add(MakeCol("Remark",           "Remark",   130))
        End If
    End Sub

    Private Function MakeCol(header As String, binding As String, width As Integer, Optional canResize As Boolean = True) As DataGridTextColumn
        Dim col As New DataGridTextColumn()
        col.Header = header
        col.Binding = New Binding(binding)
        col.Width = New DataGridLength(width)
        col.CanUserResize = canResize
        Return col
    End Function

    Private Function MakeDateCol(header As String, binding As String, width As Integer) As DataGridTextColumn
        Dim col As New DataGridTextColumn()
        col.Header = header
        Dim b As New Binding(binding)
        b.StringFormat = "{0:dd/MM/yyyy}"
        col.Binding = b
        col.Width = New DataGridLength(width)
        Return col
    End Function

    ' ═══════════════════════════════════════════════════════════════════════
    '  GENERIC SPACER ROW
    ' ═══════════════════════════════════════════════════════════════════════
    Private Sub AddGenericSpacerRow(dtFinal As DataTable, config As CalibrationTabConfig)
        Dim spacerRow = dtFinal.NewRow()
        spacerRow("RowColor")    = "White"
        spacerRow("RowHeight")   = 15
        spacerRow("IsSpacer")    = True
        spacerRow("sno")         = ""
        spacerRow("ControlNo")   = ""
        spacerRow("Size")        = ""
        spacerRow("LC")          = ""
        spacerRow("Color")       = ""
        spacerRow("Location")    = ""
        spacerRow("Date")        = DBNull.Value
        spacerRow("Temperature") = ""
        spacerRow("Humidity")    = ""
        spacerRow("TMU")         = ""
        spacerRow("TimeIn")      = ""
        spacerRow("TimeOut")     = ""
        spacerRow("TotalTime")   = ""
        spacerRow("Status")      = ""
        spacerRow("Remark")      = ""

        If config.IsSequentialObs Then
            For gi As Integer = 0 To config.ObsGroups.Count - 1
                Dim grp = config.ObsGroups(gi)
                If grp.MinLimitCol <> "" Then spacerRow($"SG{gi}_Min") = ""
                If grp.MaxLimitCol <> "" Then spacerRow($"SG{gi}_Max") = ""
                If grp.NominalCol  <> "" Then spacerRow($"SG{gi}_Nom") = ""
                If grp.PermErrCol  <> "" Then spacerRow($"SG{gi}_Err") = ""
                spacerRow($"SG{gi}") = ""  ' single obs col
            Next
        Else
            For gi As Integer = 0 To config.ObsGroups.Count - 1
                spacerRow($"GG{gi}") = ""  ' 1 col per geo group
            Next
            For Each r In config.Ranges
                Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                              r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                Dim safeColName = "Col_Obs_" & rStr.Replace(".", "_")
                If config.HasExtInt Then
                    spacerRow("Ext_" & safeColName) = ""
                    spacerRow("Int_" & safeColName) = ""
                Else
                    spacerRow(safeColName) = ""
                End If
            Next
            spacerRow("DepthError") = ""
        End If
        dtFinal.Rows.Add(spacerRow)
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════
    '  DATA HELPER METHODS
    ' ═══════════════════════════════════════════════════════════════════════

    Private Function SafeCol(row As DataRow, colName As String) As String
        If row.Table.Columns.Contains(colName) AndAlso Not IsDBNull(row(colName)) Then
            Return row(colName).ToString()
        End If
        Return ""
    End Function

    Private Function GetColValue(row As DataRow, colName As String) As Object
        For Each col As DataColumn In row.Table.Columns
            If col.ColumnName.Equals(colName, StringComparison.OrdinalIgnoreCase) Then
                Return row(col)
            End If
        Next
        Return Nothing
    End Function

    ' ═══════════════════════════════════════════════════════════════════════
    '  GENERIC EXPORT TO CSV (reads from currently selected tab's config)
    ' ═══════════════════════════════════════════════════════════════════════
    ' ═══════════════════════════════════════════════════════════════════════
    '  EXPORT TO EXCEL  — mirrors on-screen layout exactly
    ' ═══════════════════════════════════════════════════════════════════════
    Private Sub BtnExportExcel_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim dlg As New Microsoft.Win32.SaveFileDialog()
            dlg.Filter   = "Excel Workbook (*.xlsx)|*.xlsx"
            dlg.Title    = "Export Calibration Summary to Excel"
            dlg.FileName = $"{_instrumentName.Replace(" ", "_")}_CalibrationSummary_{DateTime.Now:yyyyMMdd}.xlsx"
            If dlg.ShowDialog() <> True Then Return

            Using wb As New XLWorkbook()
                For Each item As TabItem In TabSummary.Items
                    Dim td = TryCast(item.Tag, TabData)
                    If td Is Nothing OrElse td.Table Is Nothing Then Continue For
                    ExportTabToExcel(wb, item.Header.ToString(), td.Table, td.Config)
                Next
                wb.SaveAs(dlg.FileName)
            End Using

            Dim result = MessageBox.Show("Exported successfully! Open file now?",
                                        "Export to Excel", MessageBoxButton.YesNo, MessageBoxImage.Information)
            If result = MessageBoxResult.Yes Then
                Process.Start(New ProcessStartInfo(dlg.FileName) With {.UseShellExecute = True})
            End If
        Catch ex As Exception
            MessageBox.Show("Export error: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Writes one worksheet per tab that mirrors the on-screen layout:
    ''' Row 1 = group header spans, Row 2 = column names, Row 3+ = data.
    ''' </summary>
    Private Sub ExportTabToExcel(wb As XLWorkbook, sheetName As String, dtFinal As DataTable, config As CalibrationTabConfig)
        ' Sheet names max 31 chars
        Dim safeName = If(sheetName.Length > 31, sheetName.Substring(0, 31), sheetName)
        Dim ws = wb.Worksheets.Add(safeName)

        ' ── Build ordered column list matching AddColumns ────────────────────
        ' Each entry: (Excel column header, DataTable column name)
        Dim cols As New List(Of (hdr As String, dtCol As String))
        cols.Add(("S.No",                     "sno"))
        cols.Add(("Control No.",              "ControlNo"))
        cols.Add(("Range/Size",               "Size"))
        cols.Add(("L.C",                      "LC"))
        cols.Add(("Colour",                   "Color"))
        cols.Add(("Location",                 "Location"))
        cols.Add(("Calibration Date",         "Date"))
        cols.Add(("Temperature [20±2°C]",     "Temperature"))
        cols.Add(("Humidity [40%-60%] Rh",    "Humidity"))
        cols.Add(("T.M.U [± mm]",             "TMU"))

        ' Track group spans for row 1: (label, startCol 1-based, endCol 1-based)
        Dim spans As New List(Of (lbl As String, c1 As Integer, c2 As Integer))
        spans.Add(("", 1, 10))   ' info cols blank
        Dim ci = 11              ' next available 1-based col index

        If config.IsSequentialObs Then
            For gi = 0 To config.ObsGroups.Count - 1
                Dim grp = config.ObsGroups(gi)
                Dim spanStart = ci
                If grp.NominalCol  <> "" Then cols.Add(("Nominal",    $"SG{gi}_Nom")) : ci += 1
                If grp.PermErrCol  <> "" Then cols.Add(("Perm. Err.", $"SG{gi}_Err")) : ci += 1
                If grp.MinLimitCol <> "" Then cols.Add(("Min Limit",  $"SG{gi}_Min")) : ci += 1
                If grp.MaxLimitCol <> "" Then cols.Add(("Max Limit",  $"SG{gi}_Max")) : ci += 1
                cols.Add((grp.GroupHeader, $"SG{gi}"))
                spans.Add((grp.GroupHeader, spanStart, ci))
                ci += 1
            Next
            Dim rs = ci
            cols.Add(("Time In (Hrs)",    "TimeIn"))
            cols.Add(("Time Out (Hrs)",   "TimeOut"))
            cols.Add(("Total Time (Hrs)", "TotalTime"))
            cols.Add(("Status (OK/NG)",   "Status"))
            cols.Add(("Remark",           "Remark"))
            spans.Add(("", rs, rs + 4))
        Else
            ' Geo groups
            For gi = 0 To config.ObsGroups.Count - 1
                cols.Add((config.ObsGroups(gi).GroupHeader, $"GG{gi}"))
                spans.Add((config.ObsGroups(gi).GroupHeader, ci, ci))
                ci += 1
            Next
            If config.HasExtInt Then
                Dim extStart = ci
                For Each r In config.Ranges
                    Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                                  r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                    cols.Add((r.ToString("0.##") & " mm", "Ext_Col_Obs_" & rStr.Replace(".", "_")))
                    ci += 1
                Next
                spans.Add((config.ExtGroupHeader, extStart, ci - 1))
                Dim intStart = ci
                For Each r In config.Ranges
                    Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                                  r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                    cols.Add((r.ToString("0.##") & " mm", "Int_Col_Obs_" & rStr.Replace(".", "_")))
                    ci += 1
                Next
                spans.Add((config.IntGroupHeader, intStart, ci - 1))
            Else
                Dim obsStart = ci
                For Each r In config.Ranges
                    Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r),
                                  r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_").TrimEnd("0"c).TrimEnd("_"c))
                    cols.Add((r.ToString("0.##") & " mm", "Col_Obs_" & rStr.Replace(".", "_")))
                    ci += 1
                Next
                If config.Ranges.Length > 0 Then spans.Add((config.GroupHeaderText, obsStart, ci - 1))
            End If
            ' Result cols
            Dim res = ci
            cols.Add(("Depth Error (mm)", "DepthError"))
            cols.Add(("Time In (Hrs)",    "TimeIn"))
            cols.Add(("Time Out (Hrs)",   "TimeOut"))
            cols.Add(("Total Time (Hrs)", "TotalTime"))
            cols.Add(("Status (OK/NG)",   "Status"))
            cols.Add(("Remark",           "Remark"))
            spans.Add(("", res, res + 5))
        End If

        ' ── Row 1: Group header spans ───────────────────────────────────────
        Dim headerBg = XLColor.FromHtml("#E8E8E8")
        For Each sp In spans
            Dim r1 = ws.Range(1, sp.c1, 1, sp.c2)
            If sp.c1 <> sp.c2 Then r1.Merge()
            r1.FirstCell().Value = sp.lbl
            r1.Style.Fill.BackgroundColor      = headerBg
            r1.Style.Font.Bold                 = True
            r1.Style.Font.FontSize             = 10
            r1.Style.Alignment.Horizontal      = XLAlignmentHorizontalValues.Center
            r1.Style.Alignment.Vertical        = XLAlignmentVerticalValues.Center
            r1.Style.Alignment.WrapText        = True
            r1.Style.Border.OutsideBorder      = XLBorderStyleValues.Thin
            r1.Style.Border.OutsideBorderColor = XLColor.FromHtml("#999999")
        Next

        ' ── Row 2: Column names ─────────────────────────────────────────────
        For c = 0 To cols.Count - 1
            Dim cell = ws.Cell(2, c + 1)
            cell.Value = cols(c).hdr
            cell.Style.Fill.BackgroundColor      = headerBg
            cell.Style.Font.Bold                 = True
            cell.Style.Font.FontSize             = 10
            cell.Style.Alignment.Horizontal      = XLAlignmentHorizontalValues.Center
            cell.Style.Alignment.Vertical        = XLAlignmentVerticalValues.Center
            cell.Style.Alignment.WrapText        = True
            cell.Style.Border.OutsideBorder      = XLBorderStyleValues.Thin
            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#999999")
        Next

        ' ── Row height for headers ───────────────────────────────────────────
        ws.Row(1).Height = 30
        ws.Row(2).Height = 40

        ' ── Data rows ────────────────────────────────────────────────────────
        Dim excelRow = 3
        For Each dr As DataRow In dtFinal.Rows
            ' Spacer rows: write a short blank row and skip
            Dim isSpacer As Boolean = False
            If dtFinal.Columns.Contains("IsSpacer") AndAlso Not IsDBNull(dr("IsSpacer")) Then
                isSpacer = CBool(dr("IsSpacer"))
            End If
            If isSpacer Then
                ws.Row(excelRow).Height = 8
                excelRow += 1
                Continue For
            End If

            ' Row background from RowColor
            Dim rowBg = XLColor.White
            If dtFinal.Columns.Contains("RowColor") AndAlso Not IsDBNull(dr("RowColor")) Then
                Dim colorStr = dr("RowColor").ToString().Trim()
                Try
                    rowBg = XLColor.FromHtml(If(colorStr.Equals("White", StringComparison.OrdinalIgnoreCase), "#FFFFFF", colorStr))
                Catch
                End Try
            End If

            For c = 0 To cols.Count - 1
                Dim dtColName = cols(c).dtCol
                Dim cell = ws.Cell(excelRow, c + 1)
                If dtFinal.Columns.Contains(dtColName) AndAlso Not IsDBNull(dr(dtColName)) Then
                    Dim v = dr(dtColName)
                    If TypeOf v Is DateTime Then
                        cell.Value = CDate(v)
                        cell.Style.DateFormat.Format = "dd/MM/yyyy"
                    ElseIf IsNumeric(v) AndAlso Not String.IsNullOrWhiteSpace(v.ToString()) Then
                        Dim d As Double
                        If Double.TryParse(v.ToString(), System.Globalization.NumberStyles.Any,
                                          System.Globalization.CultureInfo.InvariantCulture, d) Then
                            cell.Value = d
                        Else
                            cell.Value = v.ToString()
                        End If
                    Else
                        cell.Value = v.ToString()
                    End If
                End If
                cell.Style.Fill.BackgroundColor      = rowBg
                cell.Style.Alignment.Horizontal      = XLAlignmentHorizontalValues.Center
                cell.Style.Alignment.Vertical        = XLAlignmentVerticalValues.Center
                cell.Style.Border.OutsideBorder      = XLBorderStyleValues.Thin
                cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#BBBBBB")
            Next
            ws.Row(excelRow).Height = 20
            excelRow += 1
        Next

        ' ── Freeze top 2 header rows ─────────────────────────────────────────
        ws.SheetView.FreezeRows(2)

        ' ── Auto-fit columns (cap at 50) ─────────────────────────────────────
        ws.Columns().AdjustToContents()
        For Each col As IXLColumn In ws.Columns()
            If col.Width > 50 Then col.Width = 50
        Next
    End Sub





    ''' <summary>
    ''' Recursively find a DataGrid inside a visual tree.
    ''' </summary>
    Private Function FindDataGridInVisual(parent As DependencyObject) As DataGrid
        If parent Is Nothing Then Return Nothing
        If TypeOf parent Is DataGrid Then Return DirectCast(parent, DataGrid)

        Dim childCount = VisualTreeHelper.GetChildrenCount(parent)
        For i = 0 To childCount - 1
            Dim result = FindDataGridInVisual(VisualTreeHelper.GetChild(parent, i))
            If result IsNot Nothing Then Return result
        Next

        ' Check logical children too (for ContentPresenter, ScrollViewer etc.)
        For Each child In LogicalTreeHelper.GetChildren(parent)
            If TypeOf child Is DependencyObject Then
                Dim result = FindDataGridInVisual(DirectCast(child, DependencyObject))
                If result IsNot Nothing Then Return result
            End If
        Next

        Return Nothing
    End Function
End Class
