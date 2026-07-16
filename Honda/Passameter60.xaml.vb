Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports MySql.Data.MySqlClient
Imports System.Linq

Public Class Passameter60
    Private MySqlCls As New MySQLClass()
    Private _masterLimits As DataRow = Nothing
    Private _instrumentName As String = "Passameter60"
    Private _category As String = "Instrument"
    Private _selectedMastersObservation As New List(Of MasterSelectorItem)
    Private _selectedMastersGeometric As New List(Of MasterSelectorItem)



    Private _timerRealTime As System.Windows.Threading.DispatcherTimer
    Private _isTimeOutManual As Boolean = False
    Private _isPopulating As Boolean = False

    Public Property IsEditMode As Boolean = False
    Public Property TargetCycle As String = ""

    Private ReadOnly observationPoints As String() = {"10", "20", "30", "40", "50", "60"}
    Private ReadOnly geometricFactors As String() = {"Anvil", "Spindle", "Parallel"}

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ComboBox1.Text = _instrumentName
        
        LoadDynamicLC()

        ' Add handlers for Plus observations
        For Each pt In observationPoints
            For i = 1 To 3
                Dim tb = DirectCast(Me.FindName($"txtPlus_{pt}_{i}"), TextBox)
                If tb IsNot Nothing Then
                    AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                    AddHandler tb.TextChanged, AddressOf TrialTextChanged
                    AddHandler tb.LostFocus, AddressOf ObservationLostFocus
                End If
            Next
        Next

        ' Add handlers for Minus observations
        For Each pt In observationPoints
            For i = 1 To 3
                Dim tb = DirectCast(Me.FindName($"txtMinus_{pt}_{i}"), TextBox)
                If tb IsNot Nothing Then
                    AddHandler tb.KeyDown, AddressOf ObservationKeyDown
                    AddHandler tb.TextChanged, AddressOf TrialTextChanged
                    AddHandler tb.LostFocus, AddressOf ObservationLostFocus
                End If
            Next
        Next

        ' Add handlers for Geometric trials
        For Each gf In geometricFactors
            For i = 1 To 3
                Dim tb = DirectCast(Me.FindName($"txt{gf}_{i}"), TextBox)
                If tb IsNot Nothing Then
                    AddHandler tb.TextChanged, AddressOf GeometricTextChanged
                    AddHandler tb.KeyDown, AddressOf GeometricKeyDown
                    AddHandler tb.LostFocus, AddressOf ObservationLostFocus
                End If
            Next
        Next

        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            Try
                LoadInstrumentDetails(TextBox1.Text.Trim())
            Catch ex As Exception
            End Try
        End If

        If Not IsEditMode Then
            TxtCalibrationDate.SelectedDate = DateTime.Today
        End If

        ' Add handlers for Duration
        AddHandler TextBox24.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.TextChanged, AddressOf UpdateDuration
        AddHandler TextBox26.GotFocus, Sub() _isTimeOutManual = True
        AddHandler TextBox26.PreviewKeyDown, Sub() _isTimeOutManual = True
        AddHandler TextBox21.SelectionChanged, AddressOf TextBox21_SelectionChanged
        AddHandler TextBox21.DropDownClosed, Sub() AdjustVisibilityBySize(TextBox21.Text)
        AddHandler TextBox21.LostFocus, Sub() AdjustVisibilityBySize(TextBox21.Text)
        AddHandler ComboBox5.SelectionChanged, Sub() AdjustVisibilityBySize(TextBox21.Text)

        _timerRealTime = New System.Windows.Threading.DispatcherTimer()
        _timerRealTime.Interval = TimeSpan.FromSeconds(5)
        AddHandler _timerRealTime.Tick, AddressOf TimerRealTime_Tick
        _timerRealTime.Start()

        If IsEditMode Then
            Try
                PopulateCalibrationData()
            Catch ex As Exception
            End Try
        End If

        ' Add handlers for Environmental conditions
        AddHandler TextBox30.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox28.TextChanged, AddressOf EnvironmentTextChanged
        AddHandler TextBox20.TextChanged, AddressOf EnvironmentTextChanged
        AdjustVisibilityBySize(TextBox21.Text)
    End Sub

    Private Sub TextBox21_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If TextBox21.SelectedItem IsNot Nothing Then
            Dim selectedSize As String = ""
            If TypeOf TextBox21.SelectedItem Is ComboBoxItem Then
                selectedSize = DirectCast(TextBox21.SelectedItem, ComboBoxItem).Content.ToString()
            Else
                selectedSize = TextBox21.SelectedItem.ToString()
            End If
            _masterLimits = ResolvePassameter60MasterLimits(selectedSize)
            AdjustVisibilityBySize(selectedSize)
            TriggerAllValidations()
        End If
    End Sub

    ''' <summary>
    ''' Finds the master limits row for the given instrument size string.
    ''' First tries an exact match; if not found, parses the upper bound of the
    ''' instrument size and returns the first master whose upper bound covers it.
    ''' E.g. instrument '0-25mm' → master '0-50mm' (25 ≤ 50).
    ''' </summary>
    Private Function ResolvePassameter60MasterLimits(instrumentSizeStr As String) As DataRow
        ' 1. Try exact match first
        Dim exact = MySqlCls.GetPassameter60MasterLimits(instrumentSizeStr)
        If exact IsNot Nothing Then Return exact

        ' 2. Parse the upper bound of the instrument size
        Dim instrMax As Decimal = 9999
        Dim clean = instrumentSizeStr.Replace("mm", "").Trim()
        If clean.Contains("-") Then
            Dim p = clean.Split("-"c)
            If p.Length > 1 Then Decimal.TryParse(p(1).Trim(), instrMax)
        Else
            Decimal.TryParse(clean, instrMax)
        End If

        ' 3. Load all master rows ordered ascending and pick first whose upper bound >= instrMax
        Try
            Dim dt = MySqlCls.ReadDatatable("SELECT * FROM passameter_60 WHERE RowType='MASTER' ORDER BY ID ASC")
            For Each masterRow As DataRow In dt.Rows
                Dim masterSizeStr = masterRow("Size").ToString().Replace("mm", "").Trim()
                Dim masterMax As Decimal = 9999
                If masterSizeStr.Contains("-") Then
                    Dim p = masterSizeStr.Split("-"c)
                    If p.Length > 1 Then Decimal.TryParse(p(1).Trim(), masterMax)
                Else
                    Decimal.TryParse(masterSizeStr, masterMax)
                End If
                If instrMax <= masterMax Then Return masterRow
            Next
        Catch ex As Exception
            Console.WriteLine("ResolvePassameter60MasterLimits Error: " & ex.Message)
        End Try

        Return Nothing
    End Function

    Private Sub AdjustVisibilityBySize(sizeStr As String)
        Try
            Dim activeMinRange As Decimal = 0
            Dim activeMaxRange As Decimal = 60
            If Not String.IsNullOrEmpty(sizeStr) Then
                If sizeStr.Contains("-") Then
                    Dim parts = sizeStr.Split("-"c)
                    If parts.Length > 1 Then
                        Decimal.TryParse(parts(0).Replace("mm", "").Trim(), activeMinRange)
                        Decimal.TryParse(parts(1).Replace("mm", "").Trim(), activeMaxRange)
                    End If
                Else
                    activeMinRange = 0
                    Decimal.TryParse(sizeStr.Replace("mm", "").Trim(), activeMaxRange)
                End If
            End If

            For Each pt In observationPoints
                Dim ptVal As Decimal = Decimal.Parse(pt)
                Dim inRange = (ptVal >= activeMinRange AndAlso ptVal <= activeMaxRange)
                
                For i = 1 To 3
                    Dim tbPlus = DirectCast(Me.FindName($"txtPlus_{pt}_{i}"), TextBox)
                    If tbPlus IsNot Nothing Then
                        If Not inRange Then
                            tbPlus.IsEnabled = False
                            tbPlus.Text = ""
                            tbPlus.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                        Else
                            If ComboBox5.SelectedItem IsNot Nothing Then
                                tbPlus.IsEnabled = True
                                If String.IsNullOrWhiteSpace(tbPlus.Text) Then tbPlus.Background = Brushes.White
                            Else
                                tbPlus.IsEnabled = False
                                tbPlus.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                            End If
                        End If
                    End If

                    Dim tbMinus = DirectCast(Me.FindName($"txtMinus_{pt}_{i}"), TextBox)
                    If tbMinus IsNot Nothing Then
                        If Not inRange Then
                            tbMinus.IsEnabled = False
                            tbMinus.Text = ""
                            tbMinus.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                        Else
                            If ComboBox5.SelectedItem IsNot Nothing Then
                                tbMinus.IsEnabled = True
                                If String.IsNullOrWhiteSpace(tbMinus.Text) Then tbMinus.Background = Brushes.White
                            Else
                                tbMinus.IsEnabled = False
                                tbMinus.Background = New SolidColorBrush(Color.FromRgb(241, 245, 249))
                            End If
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TxtCalibrationDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles TxtCalibrationDate.SelectedDateChanged
        If TxtCalibrationDate.SelectedDate.HasValue Then
            Dim selectedDateStr = TxtCalibrationDate.SelectedDate.Value.ToShortDateString()
            Dim env = EnvironmentCache.LoadEnvironment(Me.GetType().Name, selectedDateStr)
            If env IsNot Nothing Then
                TextBox30.Text = env.Temperature
                TextBox28.Text = env.Humidity
                TextBox20.Text = env.DepthError
                TextBox24.Text = env.TimeIn
            Else
                TextBox30.Text = "" : TextBox28.Text = "" : TextBox20.Text = "" : TextBox24.Text = ""
            End If
        End If
    End Sub

    Private Sub PopulateCalibrationData()
        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse String.IsNullOrWhiteSpace(TargetCycle) Then Return

        Try
            Dim dt = MySqlCls.GetPassameter60Data(TextBox1.Text.Trim(), TargetCycle)
            If dt.Rows.Count > 0 Then
                Dim dr = dt.Rows(0)
                _isPopulating = True

                ComboBox2.Text = dr("Location").ToString()
                
                ' Set Size
                Dim sizeVal = dr("Size").ToString()
                Dim sizeMatched = False
                For Each item As Object In TextBox21.Items
                    If TypeOf item Is ComboBoxItem Then
                        If DirectCast(item, ComboBoxItem).Content.ToString() = sizeVal Then
                            TextBox21.SelectedItem = item
                            sizeMatched = True
                            Exit For
                        End If
                    Else
                        If item.ToString() = sizeVal Then
                            TextBox21.SelectedItem = item
                            sizeMatched = True
                            Exit For
                        End If
                    End If
                Next
                If Not sizeMatched AndAlso Not String.IsNullOrEmpty(sizeVal) Then
                    TextBox21.Text = sizeVal
                End If
                ' Resolve master limits for the selected size
                _masterLimits = ResolvePassameter60MasterLimits(TextBox21.Text)
                
                ComboBox5.Text = dr("LC").ToString()
                TextBoxColor.Text = dr("Color").ToString()
                TextBox30.Text = dr("Temperature").ToString()
                TextBox28.Text = dr("Humidity").ToString()
                TextBox29.Text = dr("TMU").ToString()
                TextBox19.Text = dr("Remark").ToString()
                TextBox20.Text = dr("DepthError").ToString()

                If Not IsDBNull(dr("Date")) Then TxtCalibrationDate.SelectedDate = Convert.ToDateTime(dr("Date"))
                TextBox24.Text = dr("TimeIn").ToString()
                TextBox26.Text = dr("TimeOut").ToString()
                TextBox25.Text = dr("TotalTime").ToString()
                
                If dr("Status").ToString() = "NG" Then
                    ComboBox4.SelectedIndex = 1
                Else
                    ComboBox4.SelectedIndex = 0
                End If

                ' Geometric Factors
                Dim anvil = {dr("Flatness_Anvil_1"), dr("Flatness_Anvil_2"), dr("Flatness_Anvil_3")}
                Dim spindle = {dr("Flatness_Spindle_1"), dr("Flatness_Spindle_2"), dr("Flatness_Spindle_3")}
                Dim parallel = {dr("Parallelism_1"), dr("Parallelism_2"), dr("Parallelism_3")}
                PopulateGeometricFields("Anvil", anvil)
                PopulateGeometricFields("Spindle", spindle)
                PopulateGeometricFields("Parallel", parallel)

                ' Plus Observations
                For Each pt In observationPoints
                    For i = 1 To 3
                        Dim tb = DirectCast(Me.FindName($"txtPlus_{pt}_{i}"), TextBox)
                        If tb IsNot Nothing Then tb.Text = If(IsDBNull(dr($"Plus_{pt}_{i}")), "", dr($"Plus_{pt}_{i}").ToString())
                    Next
                Next

                ' Minus Observations
                For Each pt In observationPoints
                    For i = 1 To 3
                        Dim tb = DirectCast(Me.FindName($"txtMinus_{pt}_{i}"), TextBox)
                        If tb IsNot Nothing Then tb.Text = If(IsDBNull(dr($"Minus_{pt}_{i}")), "", dr($"Minus_{pt}_{i}").ToString())
                    Next
                Next

                ' Restore cached calibration masters
                Dim cached = CalibrationMasterCache.LoadMastersDual(TextBox1.Text.Trim(), TargetCycle)
                If cached.Observation.Count > 0 Then
                    _selectedMastersObservation = MySqlCls.GetMastersByDescriptions(cached.Observation)
                    UpdateMasterUIObservation()
                End If
                If cached.Geometric.Count > 0 Then
                    _selectedMastersGeometric = MySqlCls.GetMastersByDescriptions(cached.Geometric)
                    UpdateMasterUIGeometric()
                End If

                _isPopulating = False
                AdjustVisibilityBySize(TextBox21.Text)
                TriggerAllValidations()
                CalculateTMU()
            End If
        Catch ex As Exception
            Console.WriteLine("PopulateCalibrationData Error: " & ex.Message)
            _isPopulating = False
        End Try
    End Sub

    Private Sub TriggerAllValidations()
        Try
            ' Environment
            EnvironmentTextChanged(TextBox30, Nothing)
            EnvironmentTextChanged(TextBox28, Nothing)
            EnvironmentTextChanged(TextBox20, Nothing)

            ' Geometric
            For Each gf In geometricFactors
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txt{gf}_{i}"), TextBox)
                    If tb IsNot Nothing Then GeometricTextChanged(tb, Nothing)
                Next
            Next

            ' Plus Observations
            For Each pt In observationPoints
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txtPlus_{pt}_{i}"), TextBox)
                    If tb IsNot Nothing Then TrialTextChanged(tb, Nothing)
                Next
            Next

            ' Minus Observations
            For Each pt In observationPoints
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txtMinus_{pt}_{i}"), TextBox)
                    If tb IsNot Nothing Then TrialTextChanged(tb, Nothing)
                Next
            Next

            UpdateOverallStatus()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PopulateGeometricFields(factor As String, vals As Object())
        For i As Integer = 1 To 3
            Dim tb = DirectCast(Me.FindName($"txt{factor}_{i}"), TextBox)
            If tb IsNot Nothing Then
                tb.Text = If(IsDBNull(vals(i - 1)), "", vals(i - 1).ToString())
            End If
        Next
    End Sub

    Private Sub GeometricKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)
            Dim nameParts = tb.Name.Split("_"c) ' txtAnvil_1
            If nameParts.Length = 2 Then
                Dim factor = nameParts(0).Replace("txt", "")
                Dim trialNum = Integer.Parse(nameParts(1))
                If trialNum < 3 Then
                    Dim nextTb = DirectCast(Me.FindName($"txt{factor}_{trialNum + 1}"), TextBox)
                    If nextTb IsNot Nothing Then nextTb.Focus()
                Else
                    Dim factorIdx = Array.IndexOf(geometricFactors, factor)
                    If factorIdx < geometricFactors.Length - 1 Then
                        Dim nextFactor = geometricFactors(factorIdx + 1)
                        Dim nextTb = DirectCast(Me.FindName($"txt{nextFactor}_1"), TextBox)
                        If nextTb IsNot Nothing Then nextTb.Focus()
                    Else
                        Dim nextTb = DirectCast(Me.FindName($"txtPlus_10_1"), TextBox)
                        If nextTb IsNot Nothing Then nextTb.Focus()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObservationKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim tb = DirectCast(sender, TextBox)
            Dim parts = tb.Name.Split("_"c)
            If parts.Length = 3 Then
                Dim dir = parts(0).Replace("txt", "") ' Plus or Minus
                Dim pt = parts(1)
                Dim trialNum = Integer.Parse(parts(2))
                If trialNum < 3 Then
                    Dim nextTb = DirectCast(Me.FindName($"txt{dir}_{pt}_{trialNum + 1}"), TextBox)
                    If nextTb IsNot Nothing Then nextTb.Focus()
                Else
                    Dim ptIdx = Array.IndexOf(observationPoints, pt)
                    If ptIdx < observationPoints.Length - 1 Then
                        Dim nextTb = DirectCast(Me.FindName($"txt{dir}_{observationPoints(ptIdx + 1)}_1"), TextBox)
                        If nextTb IsNot Nothing Then nextTb.Focus()
                    Else
                        If dir = "Plus" Then
                            Dim nextTb = DirectCast(Me.FindName($"txtMinus_10_1"), TextBox)
                            If nextTb IsNot Nothing Then nextTb.Focus()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObservationLostFocus(sender As Object, e As RoutedEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        If Not String.IsNullOrWhiteSpace(tb.Text) Then
            Dim val As Decimal
            If Decimal.TryParse(tb.Text, val) Then
                tb.Text = val.ToString("0.0000")
            End If
        End If
    End Sub

    Private Sub TrialTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        If Not txtBox.IsEnabled Then Return
        Try
            CalculateTMU()

            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If
            
            Dim parts = txtBox.Name.Split("_"c)
            If parts.Length = 3 AndAlso _masterLimits IsNot Nothing Then
                Dim dir = parts(0).Replace("txt", "") ' Plus or Minus
                Dim pt = parts(1)
                
                Dim ulCol = $"{dir}_{pt}_UL"
                Dim llCol = $"{dir}_{pt}_LL"
                
                If _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    Dim ul = Convert.ToDecimal(_masterLimits(ulCol))
                    Dim ll = Convert.ToDecimal(_masterLimits(llCol))
                    Dim val As Decimal
                    If Decimal.TryParse(txtBox.Text, val) Then
                        If val > ul OrElse val < ll Then
                            txtBox.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226))
                            txtBox.Foreground = Brushes.Red
                        Else
                            txtBox.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
                            txtBox.Foreground = Brushes.Black
                        End If
                    End If
                End If
            End If
            UpdateOverallStatus()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub GeometricTextChanged(sender As Object, e As TextChangedEventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        Try
            CalculateTMU()
            
            If String.IsNullOrWhiteSpace(txtBox.Text) Then
                txtBox.Background = Brushes.White
                Return
            End If
            
            Dim parts = txtBox.Name.Split("_"c)
            If parts.Length = 2 AndAlso _masterLimits IsNot Nothing Then
                Dim factor = parts(0).Replace("txt", "")
                Dim ulCol = ""
                Dim llCol = ""
                
                If factor = "Anvil" Then
                    ulCol = "Flatness_Anvil_UL"
                    llCol = "Flatness_Anvil_LL"
                ElseIf factor = "Spindle" Then
                    ulCol = "Flatness_Spindle_UL"
                    llCol = "Flatness_Spindle_LL"
                ElseIf factor = "Parallel" Then
                    ulCol = "Parallelism_UL"
                    llCol = "Parallelism_LL"
                End If
                
                If Not String.IsNullOrEmpty(ulCol) AndAlso _masterLimits.Table.Columns.Contains(ulCol) AndAlso _masterLimits.Table.Columns.Contains(llCol) Then
                    Dim ul = Convert.ToDecimal(_masterLimits(ulCol))
                    Dim ll = Convert.ToDecimal(_masterLimits(llCol))
                    Dim val As Decimal
                    If Decimal.TryParse(txtBox.Text, val) Then
                        If val > ul OrElse val < ll Then
                            txtBox.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226))
                            txtBox.Foreground = Brushes.Red
                        Else
                            txtBox.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
                            txtBox.Foreground = Brushes.Black
                        End If
                    End If
                End If
            End If
            UpdateOverallStatus()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BtnBrowseMasterObservation_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMastersObservation.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMastersObservation = selector.SelectedMasters
            UpdateMasterUIObservation()
            CalculateTMU()
        End If
    End Sub

    Private Sub BtnBrowseMasterGeometric_Click(sender As Object, e As RoutedEventArgs)
        Dim currentSelection = _selectedMastersGeometric.Select(Function(m) m.Description).ToList()
        Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
        Dim selector As New MasterSelector(currentSelection, calibDate)
        selector.Owner = Me
        If selector.ShowDialog() = True Then
            _selectedMastersGeometric = selector.SelectedMasters
            UpdateMasterUIGeometric()
            CalculateTMU()
        End If
    End Sub

    Private Sub UpdateMasterUIObservation()
        ItemsControlSelectedMastersObservation.ItemsSource = Nothing
        ItemsControlSelectedMastersObservation.ItemsSource = _selectedMastersObservation
        MasterGroupBoxObservation.Visibility = If(_selectedMastersObservation.Count > 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub UpdateMasterUIGeometric()
        ItemsControlSelectedMastersGeometric.ItemsSource = Nothing
        ItemsControlSelectedMastersGeometric.ItemsSource = _selectedMastersGeometric
        MasterGroupBoxGeometric.Visibility = If(_selectedMastersGeometric.Count > 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub TimerRealTime_Tick(sender As Object, e As EventArgs)
        If Not _isTimeOutManual AndAlso Not String.IsNullOrWhiteSpace(TextBox24.Text) Then
            TextBox26.Text = DateTime.Now.ToString("HH:mm")
        End If
    End Sub

    Private Sub CalculateTMU()
        If _isPopulating OrElse TextBox29 Is Nothing Then Return
        Try
            Dim masterU_Geometric As Decimal = 0D
            If _selectedMastersGeometric.Count > 0 Then
                ' RSS-combine all selected geometric masters: u = sqrt(sum(ui^2))
                masterU_Geometric = Math.Sqrt(_selectedMastersGeometric.Sum(Function(m) m.MasterUncertainty ^ 2))
            Else
                masterU_Geometric = 0.000072D ' Fallback (Gauge Block Set)
            End If

            Dim masterU_Observation As Decimal = 0D
            If _selectedMastersObservation.Count > 0 Then
                ' RSS-combine all selected observation masters: u = sqrt(sum(ui^2))
                masterU_Observation = Math.Sqrt(_selectedMastersObservation.Sum(Function(m) m.MasterUncertainty ^ 2))
            Else
                masterU_Observation = 0.00012D ' Fallback (Gauge Block Set / Optical Flat)
            End If

            Dim totalSumSq = 0D

            ' Geometric Repeatability
            For Each gf In geometricFactors
                Dim vals As New List(Of Decimal)
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txt{gf}_{i}"), TextBox)
                    If tb IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
                        Dim val As Decimal
                        If Decimal.TryParse(tb.Text, val) Then vals.Add(val)
                    End If
                Next
                If vals.Count > 0 Then
                    Dim sigmaB1 = calcStdev(vals.ToArray()) / Math.Sqrt(3)
                    totalSumSq += (sigmaB1 ^ 2) + (masterU_Geometric ^ 2)
                End If
            Next

            ' Observations Repeatability
            Dim dirArray = {"Plus", "Minus"}
            For Each dirName In dirArray
                For Each pt In observationPoints
                    Dim vals As New List(Of Decimal)
                    For i = 1 To 3
                        Dim tb = DirectCast(Me.FindName($"txt{dirName}_{pt}_{i}"), TextBox)
                        If tb IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tb.Text) Then
                            Dim val As Decimal
                            If Decimal.TryParse(tb.Text, val) Then vals.Add(val)
                        End If
                    Next
                    If vals.Count > 0 Then
                        Dim sigmaB1 = calcStdev(vals.ToArray()) / Math.Sqrt(3)
                        totalSumSq += (sigmaB1 ^ 2) + (masterU_Observation ^ 2)
                    End If
                Next
            Next

            TextBox29.Text = (2 * Math.Sqrt(totalSumSq)).ToString("F9")
            TextBox29.Background = Brushes.White
        Catch ex As Exception
            TextBox29.Text = "Error"
        End Try
    End Sub

    Private Function calcStdev(arr() As Decimal) As Decimal
        If arr.Length <= 1 Then Return 0
        Dim avg = arr.Average()
        Dim sumSq = arr.Sum(Function(x) (x - avg) ^ 2)
        Return Math.Sqrt(sumSq / (arr.Length - 1))
    End Function

    Private Sub Button2_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MessageBox.Show("Control No cannot be empty.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Validation: Check if all observations or geometric triads are partially completed (1 or 2 trials filled)
        For Each pt In observationPoints
            Dim tbsPlus = {DirectCast(Me.FindName($"txtPlus_{pt}_1"), TextBox), DirectCast(Me.FindName($"txtPlus_{pt}_2"), TextBox), DirectCast(Me.FindName($"txtPlus_{pt}_3"), TextBox)}
            Dim countPlus = tbsPlus.Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.Text)).Count()
            If countPlus > 0 AndAlso countPlus < 3 Then
                MessageBox.Show($"Please complete all 3 trials for Plus {pt} mm (Observations).", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim tbsMinus = {DirectCast(Me.FindName($"txtMinus_{pt}_1"), TextBox), DirectCast(Me.FindName($"txtMinus_{pt}_2"), TextBox), DirectCast(Me.FindName($"txtMinus_{pt}_3"), TextBox)}
            Dim countMinus = tbsMinus.Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.Text)).Count()
            If countMinus > 0 AndAlso countMinus < 3 Then
                MessageBox.Show($"Please complete all 3 trials for Minus {pt} mm (Observations).", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
        Next

        For Each gf In geometricFactors
            Dim tbsGf = {DirectCast(Me.FindName($"txt{gf}_1"), TextBox), DirectCast(Me.FindName($"txt{gf}_2"), TextBox), DirectCast(Me.FindName($"txt{gf}_3"), TextBox)}
            Dim countGf = tbsGf.Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.Text)).Count()
            If countGf > 0 AndAlso countGf < 3 Then
                MessageBox.Show($"Please complete all 3 trials for {gf} flatness/parallelism.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
        Next

        If ComboBox4.SelectedItem Is Nothing Then
            MessageBox.Show("Please select OK/NG Status.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Duration Validation
        Dim timeInStr = TextBox24.Text.Trim()
        Dim timeOutStr = TextBox26.Text.Trim()

        If String.IsNullOrEmpty(timeOutStr) Then
            timeOutStr = DateTime.Now.ToString("HH:mm")
            TextBox26.Text = timeOutStr
        End If

        Dim duration = CalculateDuration(timeInStr, timeOutStr)
        If duration < 120 Then ' Less than 2 hours (120 minutes)
            MessageBox.Show("Calibration duration must be at least 2 hours.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim freshDetails = GetFreshInstrumentDetails(TextBox1.Text.Trim())
        Dim finalCat = freshDetails.Category
        Dim finalName = If(String.IsNullOrEmpty(freshDetails.Name), _instrumentName, freshDetails.Name)

        Dim controlNo = TextBox1.Text.Trim()
        Dim cycleName = MySqlCls.GetActiveCycleName()
        Dim size = ""
        If TextBox21.SelectedItem IsNot Nothing Then
            If TypeOf TextBox21.SelectedItem Is ComboBoxItem Then
                size = DirectCast(TextBox21.SelectedItem, ComboBoxItem).Content.ToString()
            Else
                size = TextBox21.SelectedItem.ToString()
            End If
        Else
            size = TextBox21.Text
        End If
        Dim lc = ComboBox5.Text
        Dim color = TextBoxColor.Text
        Dim location = ComboBox2.Text
        Dim temp = TextBox30.Text
        Dim humidity = TextBox28.Text
        Dim tmu = TextBox29.Text
        Dim remark = TextBox19.Text
        Dim depthError = TextBox20.Text
        Dim timeIn = TextBox24.Text
        Dim timeOut = TextBox26.Text
        Dim totalTime = TextBox25.Text
        Dim status = DirectCast(ComboBox4.SelectedItem, ComboBoxItem).Content.ToString()

        Dim dateStr = If(TxtCalibrationDate.SelectedDate, DateTime.Today).ToShortDateString()
        EnvironmentCache.SaveEnvironment(Me.GetType().Name, dateStr, TextBox30.Text, TextBox28.Text, TextBox24.Text, TextBox20.Text)

        Dim anvilFlatness = {DirectCast(Me.FindName("txtAnvil_1"), TextBox).Text, DirectCast(Me.FindName("txtAnvil_2"), TextBox).Text, DirectCast(Me.FindName("txtAnvil_3"), TextBox).Text}
        Dim spindleFlatness = {DirectCast(Me.FindName("txtSpindle_1"), TextBox).Text, DirectCast(Me.FindName("txtSpindle_2"), TextBox).Text, DirectCast(Me.FindName("txtSpindle_3"), TextBox).Text}
        Dim parallelism = {DirectCast(Me.FindName("txtParallel_1"), TextBox).Text, DirectCast(Me.FindName("txtParallel_2"), TextBox).Text, DirectCast(Me.FindName("txtParallel_3"), TextBox).Text}

        Dim plusObs As New Dictionary(Of String, Object())
        For Each pt In observationPoints
            plusObs.Add(pt, {DirectCast(Me.FindName($"txtPlus_{pt}_1"), TextBox).Text, DirectCast(Me.FindName($"txtPlus_{pt}_2"), TextBox).Text, DirectCast(Me.FindName($"txtPlus_{pt}_3"), TextBox).Text})
        Next

        Dim minusObs As New Dictionary(Of String, Object())
        For Each pt In observationPoints
            minusObs.Add(pt, {DirectCast(Me.FindName($"txtMinus_{pt}_1"), TextBox).Text, DirectCast(Me.FindName($"txtMinus_{pt}_2"), TextBox).Text, DirectCast(Me.FindName($"txtMinus_{pt}_3"), TextBox).Text})
        Next

        Dim success As Boolean = False
        If IsEditMode Then
            success = MySqlCls.UpdatePassameter60(finalName, controlNo, TargetCycle, size, lc, color, location, temp, humidity, tmu, anvilFlatness, spindleFlatness, parallelism, plusObs, minusObs, depthError, timeIn, timeOut, totalTime, status, remark, TxtCalibrationDate.SelectedDate)
        Else
            success = MySqlCls.InsertPassameter60(finalName, controlNo, cycleName, TxtCalibrationDate.SelectedDate, size, lc, color, location, temp, humidity, tmu, anvilFlatness, spindleFlatness, parallelism, plusObs, minusObs, depthError, timeIn, timeOut, totalTime, status, remark)
        End If

        If success Then
            Dim calibDate = If(TxtCalibrationDate.SelectedDate, DateTime.Today)
            Dim targetCyc = If(IsEditMode, TargetCycle, cycleName)
            MySqlCls.UpdateRegularCalibrationStatus(controlNo, calibDate, status, targetCyc)
            MySqlCls.InsertResultRecord(controlNo, _category, _instrumentName, targetCyc)

            CalibrationMasterCache.SaveMastersDual(controlNo, targetCyc, _selectedMastersObservation, _selectedMastersGeometric)

            If status = "NG" Then
                ProcessNGFlow(controlNo, targetCyc, calibDate, status)
            End If

            MessageBox.Show("Record saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.Close()
        Else
            MessageBox.Show("Failed to save record.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub ProcessNGFlow(ctrlNo As String, cycleName As String, calibDate As DateTime, statusVal As String)
        Dim dtDept = MySqlCls.ReadDatatable($"SELECT * FROM department_list WHERE `Control No` = '{ctrlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")
        If dtDept.Rows.Count > 0 Then
            Dim row = dtDept.Rows(0)
            Dim dept = row("Department").ToString(), instName = row("InstrumentName").ToString()
            Dim color = row("Color").ToString(), size = row("SizeandRange").ToString()
            Dim ngReason = If(Not String.IsNullOrWhiteSpace(TextBox19.Text), TextBox19.Text.Trim(), "Calibration NG")
            Dim wopReason = "Calibration NG" & If(Not String.IsNullOrWhiteSpace(TextBox19.Text), ": " & TextBox19.Text.Trim(), "")
            Dim dueDate = calibDate.AddYears(1).AddDays(-1)

            MySqlCls.InsertNGRecord(ctrlNo, instName, dept, cycleName, ngReason, calibDate, dueDate, statusVal)
            MySqlCls.InsertWOPRecord(ctrlNo, cycleName, dept, instName, color, wopReason)
            MySqlCls.InsertInterchangeRecord(cycleName, ctrlNo, dept, instName, size, color, "WOP", wopReason)
            MySqlCls.ClearRFIDTag(ctrlNo)
        End If
    End Sub

    Private Sub UpdateDuration(sender As Object, e As TextChangedEventArgs)
        Dim timeInStr = TextBox24.Text.Trim(), timeOutStr = TextBox26.Text.Trim()
        Dim dummyTimeOut = If(String.IsNullOrEmpty(timeOutStr), DateTime.Now.ToString("HH:mm"), timeOutStr)
        Dim totalMinutes = CalculateDuration(timeInStr, dummyTimeOut)
        If totalMinutes >= 0 Then
            Dim hours = totalMinutes \ 60, mins = totalMinutes Mod 60
            TextBox25.Text = $"{hours:D2}:{mins:D2}"
            If totalMinutes >= 120 Then
                TextBox25.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' light green
                TextBox25.Foreground = Brushes.Black
            Else
                TextBox25.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' light red
                TextBox25.Foreground = Brushes.Red
            End If
        Else
            TextBox25.Text = ""
            TextBox25.Background = Brushes.White
            TextBox25.Foreground = Brushes.Black
        End If
    End Sub

    Private Function CalculateDuration(timeInStr As String, timeOutStr As String) As Integer
        Try
            Dim timeIn, timeOut As DateTime
            Dim formats As String() = {"HH:mm", "H:mm"}
            If DateTime.TryParseExact(timeInStr.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, timeIn) AndAlso
               DateTime.TryParseExact(timeOutStr.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, timeOut) Then
                If timeOut < timeIn Then timeOut = timeOut.AddDays(1)
                Return CInt((timeOut - timeIn).TotalMinutes)
            End If
        Catch ex As Exception
        End Try
        Return -1
    End Function

    Private Sub EnvironmentTextChanged(sender As Object, e As TextChangedEventArgs)
        If _isPopulating OrElse _masterLimits Is Nothing Then Return
        Dim tb = DirectCast(sender, TextBox)
        Dim val As Decimal
        If Decimal.TryParse(tb.Text, val) Then
            Dim isOK = True
            If tb.Name = "TextBox30" Then ' Temp
                If _masterLimits.Table.Columns.Contains("Env_Temp_LL") AndAlso _masterLimits.Table.Columns.Contains("Env_Temp_UL") Then
                    Dim ll = Convert.ToDecimal(_masterLimits("Env_Temp_LL"))
                    Dim ul = Convert.ToDecimal(_masterLimits("Env_Temp_UL"))
                    If val < ll OrElse val > ul Then isOK = False
                End If
            ElseIf tb.Name = "TextBox28" Then ' Humidity
                If _masterLimits.Table.Columns.Contains("Env_Hum_LL") AndAlso _masterLimits.Table.Columns.Contains("Env_Hum_UL") Then
                    Dim ll = Convert.ToDecimal(_masterLimits("Env_Hum_LL"))
                    Dim ul = Convert.ToDecimal(_masterLimits("Env_Hum_UL"))
                    If val < ll OrElse val > ul Then isOK = False
                End If
            ElseIf tb.Name = "TextBox20" Then ' Depth Error
                If _masterLimits.Table.Columns.Contains("Depth_LL") AndAlso _masterLimits.Table.Columns.Contains("Depth_UL") Then
                    Dim ll = Convert.ToDecimal(_masterLimits("Depth_LL"))
                    Dim ul = Convert.ToDecimal(_masterLimits("Depth_UL"))
                    If val < ll OrElse val > ul Then isOK = False
                End If
            End If

            If Not isOK Then
                tb.Background = New SolidColorBrush(Color.FromRgb(254, 226, 226)) ' Light Pink
                tb.Foreground = Brushes.Red
            Else
                tb.Background = New SolidColorBrush(Color.FromRgb(220, 252, 231)) ' Light Green
                tb.Foreground = Brushes.Black
            End If
        Else
            tb.Background = Brushes.White : tb.Foreground = Brushes.Black
        End If
        UpdateOverallStatus()
    End Sub

    Private Sub LoadInstrumentDetails(controlNo As String)
        Try
            Dim dtSettings = MySqlCls.ReadDatatable("SELECT TypeName, Category FROM type_details")
            Dim found As Boolean = False

            For Each r As DataRow In dtSettings.Rows
                Dim tbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString())
                Dim query = $"SELECT * FROM `{tbl}` WHERE ControlNo = '{controlNo.Replace("'", "''")}' LIMIT 1"
                Dim dt = MySqlCls.ReadDatatable(query)
                If dt.Rows.Count > 0 Then
                    Dim row = dt.Rows(0)

                    Dim cycleName = MySqlCls.GetActiveCycleName()
                    Dim dtDept = MySqlCls.ReadDatatable($"SELECT Department, Color, SizeandRange FROM department_list WHERE `Control No` = '{controlNo.Replace("'", "''")}' AND CycleName = '{cycleName.Replace("'", "''")}' LIMIT 1")

                    _category = r("Category").ToString()
                    Dim prefix = MySQLClass.ExtractPrefix(controlNo)
                    Dim sizes = MySqlCls.GetSizesForGroup(prefix, _category)
                    TextBox21.ItemsSource = sizes

                    ' Size Matching Logic: Try to match size from database with combobox items
                    Dim matchedSize As String = ""
                    If row.Table.Columns.Contains("Size") AndAlso Not IsDBNull(row("Size")) Then
                        matchedSize = TryMatchSize(row("Size").ToString(), sizes)
                    End If

                    If Not String.IsNullOrEmpty(matchedSize) Then
                        TextBox21.Text = matchedSize
                    ElseIf sizes.Count = 1 Then
                        TextBox21.Text = sizes(0)
                    Else
                        TextBox21.Text = ""
                    End If
                    _masterLimits = ResolvePassameter60MasterLimits(TextBox21.Text)
                    AdjustVisibilityBySize(TextBox21.Text)

                    If dtDept.Rows.Count > 0 Then
                        Dim deptRow = dtDept.Rows(0)
                        ComboBox2.Text = deptRow("Department").ToString()
                        TextBoxColor.Text = deptRow("Color").ToString()
                    Else
                        ' Fallback to source table if not in department_list
                        ComboBox2.Text = row("Location").ToString()
                        TextBoxColor.Text = row("Color").ToString()
                    End If

                    _category = r("Category").ToString()
                    _instrumentName = If(_category.Equals("Instrument", StringComparison.OrdinalIgnoreCase), row("InstrumentName").ToString(), row("GaugeName").ToString())

                    found = True
                    Exit For
                End If
            Next

            If Not found Then
                _instrumentName = "Passameter60"
                ClearForm()
                AdjustVisibilityBySize(TextBox21.Text)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function TryMatchSize(inputSize As String, sizes As List(Of String)) As String
        If String.IsNullOrWhiteSpace(inputSize) Then Return ""
        Dim cleanInput = inputSize.ToLower().Replace("mm", "").Replace(" ", "").Trim()
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanInput = cleanS Then Return s
        Next
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanS.Contains("-") Then
                Dim parts = cleanS.Split("-"c)
                If parts.Length > 1 AndAlso parts(1).Trim() = cleanInput Then
                    Return s
                End If
            End If
        Next
        For Each s In sizes
            Dim cleanS = s.ToLower().Replace("mm", "").Replace(" ", "").Trim()
            If cleanS.Contains(cleanInput) OrElse cleanInput.Contains(cleanS) Then
                Return s
            End If
        Next
        Return ""
    End Function

    Private Sub ClearForm()
        ComboBox2.Text = ""
        TextBoxColor.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub LoadDynamicLC()
        Try
            Dim dt = MySqlCls.ReadDatatable("SELECT DISTINCT LC FROM passameter_60 WHERE RowType='MASTER' ORDER BY LC ASC")
            ComboBox5.Items.Clear()
            For Each row As DataRow In dt.Rows
                ComboBox5.Items.Add(row("LC").ToString())
            Next
            If ComboBox5.Items.Count > 0 Then
                ComboBox5.SelectedIndex = 0
            Else
                ComboBox5.Items.Add("0.001")
                ComboBox5.Items.Add("0.01")
                ComboBox5.SelectedIndex = 0
            End If
        Catch ex As Exception
            Console.WriteLine("LoadDynamicLC Error: " & ex.Message)
        End Try
    End Sub

    Private Sub UpdateOverallStatus()
        Dim isNG As Boolean = False

        ' 1. Check observations
        For Each pt In observationPoints
            For i = 1 To 3
                Dim tbPlus = DirectCast(Me.FindName($"txtPlus_{pt}_{i}"), TextBox)
                If tbPlus IsNot Nothing AndAlso tbPlus.Background IsNot Nothing AndAlso TypeOf tbPlus.Background Is SolidColorBrush Then
                    Dim c = DirectCast(tbPlus.Background, SolidColorBrush).Color
                    If c = Color.FromRgb(254, 226, 226) OrElse c = Color.FromRgb(255, 182, 193) Then
                        isNG = True
                        Exit For
                    End If
                End If

                Dim tbMinus = DirectCast(Me.FindName($"txtMinus_{pt}_{i}"), TextBox)
                If tbMinus IsNot Nothing AndAlso tbMinus.Background IsNot Nothing AndAlso TypeOf tbMinus.Background Is SolidColorBrush Then
                    Dim c = DirectCast(tbMinus.Background, SolidColorBrush).Color
                    If c = Color.FromRgb(254, 226, 226) OrElse c = Color.FromRgb(255, 182, 193) Then
                        isNG = True
                        Exit For
                    End If
                End If
            Next
            If isNG Then Exit For
        Next

        ' 2. Check geometric factors
        If Not isNG Then
            For Each gf In geometricFactors
                For i = 1 To 3
                    Dim tb = DirectCast(Me.FindName($"txt{gf}_{i}"), TextBox)
                    If tb IsNot Nothing AndAlso tb.Background IsNot Nothing AndAlso TypeOf tb.Background Is SolidColorBrush Then
                        Dim c = DirectCast(tb.Background, SolidColorBrush).Color
                        If c = Color.FromRgb(254, 226, 226) OrElse c = Color.FromRgb(255, 182, 193) Then
                            isNG = True
                            Exit For
                        End If
                    End If
                Next
                If isNG Then Exit For
            Next
        End If

        ' 3. Check environmental fields
        If Not isNG Then
            Dim envTbs = {TextBox30, TextBox28, TextBox20}
            For Each tb In envTbs
                If tb.Background IsNot Nothing AndAlso TypeOf tb.Background Is SolidColorBrush Then
                    Dim c = DirectCast(tb.Background, SolidColorBrush).Color
                    If c = Color.FromRgb(254, 226, 226) OrElse c = Color.FromRgb(255, 182, 193) Then
                        isNG = True
                        Exit For
                    End If
                End If
            Next
        End If

        If isNG Then
            ComboBox4.SelectedIndex = 1 ' NG
        Else
            ComboBox4.SelectedIndex = 0 ' OK
        End If
    End Sub

    Private Function GetFreshInstrumentDetails(controlNo As String) As (Category As String, Name As String)
        Try
            Dim dtSettings = MySqlCls.ReadDatatable("SELECT TypeName, Category FROM type_details")
            For Each r As DataRow In dtSettings.Rows
                Dim tbl = MySQLClass.TypeNameToTableName(r("TypeName").ToString())
                Dim query = $"SELECT * FROM `{tbl}` WHERE ControlNo = '{controlNo.Replace("'", "''")}' LIMIT 1"
                Dim dt = MySqlCls.ReadDatatable(query)
                If dt.Rows.Count > 0 Then
                    Dim row = dt.Rows(0)
                    Dim cat = r("Category").ToString()
                    Dim name = If(cat.Equals("Instrument", StringComparison.OrdinalIgnoreCase), row("InstrumentName").ToString(), row("GaugeName").ToString())
                    Return (cat, name)
                End If
            Next
        Catch ex As Exception
        End Try
        Return ("Instrument", "")
    End Function
End Class
