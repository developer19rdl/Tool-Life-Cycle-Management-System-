Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Media
Imports System.Linq
Imports System.Text.RegularExpressions

Public Class CalibrationHistoryDetailWindow
    Private _mySql As New MySQLClass()
    Private _controlNo As String
    Private _cycleName As String

    Public Sub New(controlNo As String, cycleName As String)
        InitializeComponent()
        _controlNo = controlNo
        _cycleName = cycleName
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            ' 1. Fetch record from specific instrument table
            Dim dt = _mySql.GetCalibrationDetail(_controlNo, _cycleName)
            If dt.Rows.Count = 0 Then
                MessageBox.Show("No detailed record found for this calibration.", "Information", MessageBoxButton.OK, MessageBoxImage.Information)
                Me.Close()
                Return
            End If

            Dim row = dt.Rows(0)
            
            ' 2. Fetch Category and Type Name from result_list for better display naming
            Dim queryRes = $"SELECT category, type_name FROM result_list WHERE control_no = '{_controlNo.Replace("'", "''")}' AND CycleName = '{_cycleName.Replace("'", "''")}' LIMIT 1"
            Dim dtRes = _mySql.ReadDatatable(queryRes)
            
            ' Populate UI Headers
            TxtControlNo.Text = _controlNo
            TxtCycle.Text = _cycleName
            TxtCalDate.Text = SafeGet(row, "Date", True)
            
            If dtRes.Rows.Count > 0 Then
                TxtCategory.Text = SafeGet(dtRes.Rows(0), "category")
                TxtTypeName.Text = SafeGet(dtRes.Rows(0), "type_name")
            Else
                TxtCategory.Text = SafeGet(row, "Type") ' Fallback
                TxtTypeName.Text = "-"
            End If
            
            TxtRange.Text = SafeGet(row, "Size")
            TxtLC.Text = SafeGet(row, "LC")
            TxtColor.Text = SafeGet(row, "Color")
            TxtLocation.Text = SafeGet(row, "Location")
            
            TxtTemp.Text = SafeGet(row, "Temperature")
            TxtHum.Text = SafeGet(row, "Humidity")
            TxtTMU.Text = SafeGet(row, "TMU")

            ' Populate Footer
            TxtDepthError.Text = SafeGet(row, "DepthError")
            TxtTimeIn.Text = SafeGet(row, "TimeIn")
            TxtTimeOut.Text = SafeGet(row, "TimeOut")
            TxtTotalTime.Text = SafeGet(row, "TotalTime")
            TxtRemark.Text = SafeGet(row, "Remark")
            TxtStatus.Text = SafeGet(row, "Status")
            
            If TxtStatus.Text = "OK" Then
                TxtStatus.Foreground = Brushes.Black
            ElseIf TxtStatus.Text = "NG" Then
                TxtStatus.Foreground = Brushes.Black
            Else
                TxtStatus.Foreground = Brushes.Black
            End If

            ' Process and build the measurements grid dynamically
            Dim config = AutoDetectConfigFromDataRow(row, TxtTypeName.Text)
            If config.Ranges.Length > 0 Then
                BuildMeasurementsGrid(row, config)
            Else
                Dim txtInfo As New TextBlock() With {
                    .Text = "No measurement data points found for this record.",
                    .Margin = New Thickness(20),
                    .Foreground = Brushes.Gray
                }
                ObservationsPanel.Children.Add(txtInfo)
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading detail: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' ─── CONFIG CLASS & AUTO-DETECTION ──────────────────────

    Private Class ParameterConfig
        Public Property DbBaseName As String = ""
        Public Property DisplayLabel As String = ""
    End Class

    ' ─── CONFIG: defines how each calibration template tab looks ──────
    Private Class CalibrationTabConfig
        Public Property TabLabel As String = ""
        Public Property Ranges As Decimal() = {}
        Public Property TrialCount As Integer = 3
        ''' <summary>True for 300mm (External + Internal columns), False for single Observation columns.</summary>
        Public Property HasExtInt As Boolean = False
        ''' <summary>DB column prefix for single-observation types (e.g. "v" for 600, "Obs" for Low Force).</summary>
        Public Property DbObsPrefix As String = ""
        ''' <summary>DB column prefix for external measurements (only when HasExtInt=True).</summary>
        Public Property DbExtPrefix As String = "Ext"
        ''' <summary>DB column prefix for internal measurements (only when HasExtInt=True).</summary>
        Public Property DbIntPrefix As String = "Int"
        ''' <summary>Group header text for single-observation types.</summary>
        Public Property GroupHeaderText As String = "Observation (mm) at Different Ranges*"
        ''' <summary>Group header text for external columns (only when HasExtInt=True).</summary>
        Public Property ExtGroupHeader As String = "External error(t) (mm)  at Different Ranges*"
        ''' <summary>Group header text for internal columns (only when HasExtInt=True).</summary>
        Public Property IntGroupHeader As String = "Internal error(t) (mm)  at Different Ranges*"
        ''' <summary>Maps a Decimal Range to its original string identifier (e.g. 15.0 -> "15_0").</summary>
        Public Property RangeToDbId As New Dictionary(Of Decimal, String)
        
        Public Property IsParameterBased As Boolean = False
        Public Property Parameters As New List(Of ParameterConfig)()
    End Class

    ' ─── Auto-detect config by inspecting the DataRow columns ──────
    Private Function AutoDetectConfigFromDataRow(row As DataRow, formName As String) As CalibrationTabConfig
        Dim config As New CalibrationTabConfig()

        ' Check for Bore Gauge parameter-based table
        If row.Table.Columns.Contains("Wide_Range_1") Then
            config.IsParameterBased = True
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Wide_Range", .DisplayLabel = "Wide Range Error"})
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Adjacent_Error", .DisplayLabel = "Adjacent Error"})
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Repeatability", .DisplayLabel = "Repeatability"})
            config.TrialCount = 3
            Return config
        End If

        ' Check for Dial Gauge / Dial Test Indicator parameter-based table
        If row.Table.Columns.Contains("Obs_1_10_Rev_1") Then
            config.IsParameterBased = True
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Obs_1_10_Rev", .DisplayLabel = "1/10 Rev"})
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Obs_1_2_Rev", .DisplayLabel = "1/2 Rev"})
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Obs_1_Rev", .DisplayLabel = "1 Rev"})
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Obs_Whole_Range", .DisplayLabel = "Whole Range"})
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Obs_Hysteresis", .DisplayLabel = "Hysteresis"})
            config.Parameters.Add(New ParameterConfig() With {.DbBaseName = "Obs_Repeatability", .DisplayLabel = "Repeatability"})
            config.TrialCount = 3
            Return config
        End If
        
        ' Regex to match calibration measurement columns. 
        ' Support for decimal ranges via underscores (e.g. Obs2_5_1)
        Dim rx As New Regex("^(Ext|Int|Obs|v)([\d_.]+)_(\d+)$", RegexOptions.IgnoreCase)
        
        Dim hasExt As Boolean = False
        Dim hasInt As Boolean = False
        Dim obsPrefix As String = ""
        Dim detectedExtPrefix As String = "Ext"
        Dim detectedIntPrefix As String = "Int"
        Dim ranges As New SortedSet(Of Decimal)
        Dim maxTrial As Integer = 1

        ' Inspect each column name in the result table
        For Each col As DataColumn In row.Table.Columns
            Dim m = rx.Match(col.ColumnName)
            If m.Success Then
                Dim prefix = m.Groups(1).Value
                Dim rangeStr = m.Groups(2).Value
                
                ' Parse range: handle 50, 2_5, etc.
                Dim rangeVal As Decimal
                If Decimal.TryParse(rangeStr.Replace("_", "."), System.Globalization.CultureInfo.InvariantCulture, rangeVal) Then
                    ranges.Add(rangeVal)
                    If Not config.RangeToDbId.ContainsKey(rangeVal) Then
                        config.RangeToDbId(rangeVal) = rangeStr
                    End If
                    
                    Dim trial = Integer.Parse(m.Groups(3).Value)
                    maxTrial = Math.Max(maxTrial, trial)

                    If prefix.Equals("Ext", StringComparison.OrdinalIgnoreCase) Then
                        hasExt = True
                        detectedExtPrefix = prefix
                    ElseIf prefix.Equals("Int", StringComparison.OrdinalIgnoreCase) Then
                        hasInt = True
                        detectedIntPrefix = prefix
                    Else
                        obsPrefix = prefix
                    End If
                End If
            End If
        Next

        ' Populate config
        config.Ranges = ranges.ToArray()
        config.TrialCount = maxTrial
        
        If hasExt AndAlso hasInt Then
            config.HasExtInt = True
            config.DbExtPrefix = detectedExtPrefix
            config.DbIntPrefix = detectedIntPrefix
        Else
            config.HasExtInt = False
            config.DbObsPrefix = If(obsPrefix <> "", obsPrefix, "Obs")
        End If

        Return config
    End Function

    ' ─── GENERIC GRID BUILDER ────────────────────────────────

    Private Sub BuildMeasurementsGrid(row As DataRow, config As CalibrationTabConfig)
        ' Prepare DataTable for Display
        Dim dtDisplay As New DataTable()
        dtDisplay.Columns.Add("Observation")
        
        If config.IsParameterBased Then
            For Each p In config.Parameters
                dtDisplay.Columns.Add("Param_" & p.DbBaseName)
            Next
        Else
            For Each r In config.Ranges
                Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r), r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_"))
                If config.HasExtInt Then
                    dtDisplay.Columns.Add("Ext_" & rStr)
                    dtDisplay.Columns.Add("Int_" & rStr)
                Else
                    dtDisplay.Columns.Add("Obs_" & rStr)
                End If
            Next
        End If

        ' Fill Data rows
        For trial As Integer = 1 To config.TrialCount
            Dim displayRow = dtDisplay.NewRow()
            displayRow("Observation") = trial.ToString() & If(trial = 1, "st", If(trial = 2, "nd", If(trial = 3, "rd", "th"))) & " Obs"
            
            If config.IsParameterBased Then
                For Each p In config.Parameters
                    Dim val = GetColValue(row, p.DbBaseName & "_" & trial)
                    displayRow("Param_" & p.DbBaseName) = If(val IsNot Nothing AndAlso Not IsDBNull(val) AndAlso val.ToString() <> "", val.ToString(), "-")
                Next
            Else
                For Each r In config.Ranges
                    Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r), r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_"))
                    
                    If config.HasExtInt Then
                        Dim extVal = GetColValue(row, config.DbExtPrefix & rStr & "_" & trial)
                        displayRow("Ext_" & rStr) = If(extVal IsNot Nothing AndAlso Not IsDBNull(extVal) AndAlso extVal.ToString() <> "", extVal.ToString(), "-")

                        Dim intVal = GetColValue(row, config.DbIntPrefix & rStr & "_" & trial)
                        displayRow("Int_" & rStr) = If(intVal IsNot Nothing AndAlso Not IsDBNull(intVal) AndAlso intVal.ToString() <> "", intVal.ToString(), "-")
                    Else
                        Dim obsVal = GetColValue(row, config.DbObsPrefix & rStr & "_" & trial)
                        displayRow("Obs_" & rStr) = If(obsVal IsNot Nothing AndAlso Not IsDBNull(obsVal) AndAlso obsVal.ToString() <> "", obsVal.ToString(), "-")
                    End If
                Next
            End If
            dtDisplay.Rows.Add(displayRow)
        Next

        Dim dock As New DockPanel()

        ' Group Header Grid (Stacked above columns)
        Dim ghGrid As New Grid()
        ghGrid.HorizontalAlignment = HorizontalAlignment.Left
        ghGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(100)}) ' First column "Observation"
        
        If config.IsParameterBased Then
            For i = 0 To config.Parameters.Count - 1
                ghGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(130)})
            Next
            AddSummaryGroupHeaderCel(ghGrid, 0, 1, "")
            AddSummaryGroupHeaderCel(ghGrid, 1, config.Parameters.Count, "Error observations (mm) for parameters*")
        ElseIf config.HasExtInt Then
            ' 300mm layout: External + Internal
            For i = 0 To config.Ranges.Length * 2 - 1
                ghGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(45)})
            Next
            AddSummaryGroupHeaderCel(ghGrid, 0, 1, "")
            AddSummaryGroupHeaderCel(ghGrid, 1, config.Ranges.Length, config.ExtGroupHeader)
            AddSummaryGroupHeaderCel(ghGrid, 1 + config.Ranges.Length, config.Ranges.Length, config.IntGroupHeader)
        Else
            ' Single observation layout
            For i = 0 To config.Ranges.Length - 1
                ghGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(45)})
            Next
            AddSummaryGroupHeaderCel(ghGrid, 0, 1, "")
            AddSummaryGroupHeaderCel(ghGrid, 1, config.Ranges.Length, config.GroupHeaderText)
        End If

        DockPanel.SetDock(ghGrid, System.Windows.Controls.Dock.Top)
        dock.Children.Add(ghGrid)

        ' Main DataGrid
        Dim dg As New DataGrid()
        dg.HorizontalAlignment = HorizontalAlignment.Left
        dg.AutoGenerateColumns = False
        dg.CanUserAddRows = False
        dg.IsReadOnly = True
        dg.CanUserSortColumns = False
        dg.CanUserReorderColumns = False
        dg.BorderThickness = New Thickness(0)
        dg.Background = Brushes.White
        dg.GridLinesVisibility = DataGridGridLinesVisibility.None
        dg.HeadersVisibility = DataGridHeadersVisibility.Column
        dg.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        dg.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled

        ' Stylings from App Resources
        Try
            dg.ColumnHeaderStyle = DirectCast(FindResource("SummaryHeaderStyle"), Style)
            dg.CellStyle = DirectCast(FindResource("SummaryCellStyle"), Style)
        Catch ex As Exception
        End Try

        Dim firstColStyle As Style = Nothing
        Try
            firstColStyle = DirectCast(FindResource("SummaryFirstColumnCellStyle"), Style)
        Catch ex As Exception
        End Try

        ' Add "Observation" first column
        dg.Columns.Add(MakeSummaryCol("Observation", "Observation", 100, firstColStyle))
        
        ' Add data point columns
        If config.IsParameterBased Then
            For Each p In config.Parameters
                dg.Columns.Add(MakeSummaryCol(p.DisplayLabel, "Param_" & p.DbBaseName, 130, Nothing))
            Next
        Else
            For Each r In config.Ranges
                Dim rStr = If(config.RangeToDbId.ContainsKey(r), config.RangeToDbId(r), r.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture).Replace(".", "_"))
                If config.HasExtInt Then
                    dg.Columns.Add(MakeSummaryCol(r.ToString(), "Ext_" & rStr, 45, Nothing))
                    dg.Columns.Add(MakeSummaryCol(r.ToString(), "Int_" & rStr, 45, Nothing))
                Else
                    dg.Columns.Add(MakeSummaryCol(r.ToString(), "Obs_" & rStr, 45, Nothing))
                End If
            Next
        End If

        dg.ItemsSource = dtDisplay.DefaultView
        dock.Children.Add(dg)

        ' Outer border for the integrated look
        Dim outerBorder As New Border()
        outerBorder.BorderThickness = New Thickness(1, 1, 0, 0)
        outerBorder.BorderBrush = New SolidColorBrush(Color.FromRgb(153, 153, 153))
        outerBorder.HorizontalAlignment = HorizontalAlignment.Left
        outerBorder.Child = dock

        Dim scroll As New ScrollViewer()
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll.Content = outerBorder
        scroll.Margin = New Thickness(0, 5, 0, 20)

        ' Synchronize column widths between Header Grid and DataGrid
        AddHandler dg.LayoutUpdated, Sub()
            If dg.Columns.Count = ghGrid.ColumnDefinitions.Count Then
                For idx = 0 To dg.Columns.Count - 1
                    ghGrid.ColumnDefinitions(idx).Width = New GridLength(dg.Columns(idx).ActualWidth)
                Next
            End If
        End Sub

        ObservationsPanel.Children.Add(scroll)
    End Sub

    Private Sub AddSummaryGroupHeaderCel(grid As Grid, col As Integer, span As Integer, text As String)
        Dim border As New Border()
        border.Background = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#E8E8E8"), Color))
        border.BorderBrush = New SolidColorBrush(DirectCast(ColorConverter.ConvertFromString("#999999"), Color))
        border.BorderThickness = New Thickness(0, 0, 1, 1)
        System.Windows.Controls.Grid.SetColumn(border, col)
        System.Windows.Controls.Grid.SetColumnSpan(border, span)

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

    Private Function MakeSummaryCol(header As String, binding As String, width As Integer, cellStyle As Style) As DataGridTextColumn
        Dim col As New DataGridTextColumn()
        col.Header = header
        col.Binding = New Binding(binding)
        col.Width = New DataGridLength(width)
        col.CanUserResize = False
        If cellStyle IsNot Nothing Then
            col.CellStyle = cellStyle
        End If
        Return col
    End Function

    Private Function GetColValue(row As DataRow, colName As String) As Object
        For Each col As DataColumn In row.Table.Columns
            If col.ColumnName.Equals(colName, StringComparison.OrdinalIgnoreCase) Then
                Return row(col)
            End If
        Next
        Return Nothing
    End Function

    Private Function SafeGet(row As DataRow, colName As String, Optional isDate As Boolean = False) As String
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(colName) AndAlso Not IsDBNull(row(colName)) Then
            If isDate Then
                Return Convert.ToDateTime(row(colName)).ToString("dd/MM/yyyy")
            End If
            Return row(colName).ToString()
        End If
        Return "-"
    End Function
    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub
End Class
