Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows.Data
Imports System.Data

Public Class BulkScanSummaryWindow
    Public Class IssuedItem
        Implements INotifyPropertyChanged

        Public Property ControlNo As String
        Public Property Name As String
        Public Property Department As String
        Public Property RFID As String

        Private _isChecked As Boolean = True
        Public Property IsChecked As Boolean
            Get
                Return _isChecked
            End Get
            Set(value As Boolean)
                If _isChecked <> value Then
                    _isChecked = value
                    OnPropertyChanged(NameOf(IsChecked))
                End If
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Protected Sub OnPropertyChanged(name As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
        End Sub
    End Class

    Public Class IgnoredItem
        Public Property RFID As String
        Public Property Reason As String
        Public Property ControlNo As String
        Public Property DisplayColor As String
    End Class

    Public Enum ScannerType
        USB
        RFID_Gun
    End Enum

    Private _issued As ObservableCollection(Of IssuedItem)
    Private _ignored As ObservableCollection(Of IgnoredItem)
    Private _mySql As MySQLClass
    Private _cycleName As String
    Private _isApplied As Boolean = False

    Private _parentPage As InterchangeablePage
    Private _scanType As ScannerType

    Public Sub New(issued As ObservableCollection(Of IssuedItem), ignored As ObservableCollection(Of IgnoredItem), mySql As MySQLClass, cycle As String, scanType As ScannerType, parent As InterchangeablePage)
        InitializeComponent()
        _issued = issued
        _ignored = ignored
        _mySql = mySql
        _cycleName = cycle
        _scanType = scanType
        _parentPage = parent

        If scanType = ScannerType.USB Then
            TxtTitle.Text = "Bulk Scan Summary | USB Scanner"
            TxtSubtitle.Text = "Review the results of the USB scan session."
        Else
            TxtTitle.Text = "Bulk Scan Summary | RFID Gun Scanner"
            TxtSubtitle.Text = "Review the results of the RFID gun scan session."
        End If

        ListIssued.ItemsSource = _issued
        ListIgnored.ItemsSource = _ignored

        AddHandler _issued.CollectionChanged, AddressOf OnCollectionChanged
        AddHandler _ignored.CollectionChanged, AddressOf OnCollectionChanged

        LoadDepartments()

        UpdateCounts()
        UpdateScanButtonStates(True)
    End Sub

    Private Sub LoadDepartments()
        Dim dt = _mySql.ReadDatatable("SELECT DepartmentName FROM departmentmaster ORDER BY DepartmentName")
        Dim departments As New List(Of String)
        departments.Add("All")
        For Each row As DataRow In dt.Rows
            departments.Add(row("DepartmentName").ToString())
        Next
        CmbDepartment.ItemsSource = departments
        CmbDepartment.SelectedIndex = 0
    End Sub

    Private Sub CmbDepartment_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim view As ICollectionView = CollectionViewSource.GetDefaultView(ListIssued.ItemsSource)
        If view IsNot Nothing AndAlso CmbDepartment.SelectedItem IsNot Nothing Then
            Dim selectedDept As String = CmbDepartment.SelectedItem.ToString()
            If selectedDept = "All" Then
                view.Filter = Nothing
            Else
                view.Filter = Function(item)
                                  Dim issuedItem = TryCast(item, IssuedItem)
                                  Return issuedItem IsNot Nothing AndAlso issuedItem.Department = selectedDept
                              End Function
            End If
            UpdateCounts()
        End If
    End Sub

    Private Sub OnCollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
        UpdateCounts()
    End Sub

    Private Sub UpdateCounts()
        Dim visibleCount As Integer = 0
        Dim view As ICollectionView = CollectionViewSource.GetDefaultView(ListIssued.ItemsSource)
        If view IsNot Nothing Then
            For Each item In view
                visibleCount += 1
            Next
        End If

        TxtIssuedCount.Text = $"{visibleCount} Items"
        TxtIgnoredCount.Text = $"{_ignored.Count} Tags"
        
        ' Enable/Disable Apply button based on findings
        BtnApply.IsEnabled = (visibleCount > 0)
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As RoutedEventArgs)
        Dim itemsToIssue As New List(Of IssuedItem)
        Dim view As ICollectionView = CollectionViewSource.GetDefaultView(ListIssued.ItemsSource)
        If view IsNot Nothing Then
            For Each item In view
                Dim issuedItem = TryCast(item, IssuedItem)
                If issuedItem IsNot Nothing AndAlso issuedItem.IsChecked Then
                    itemsToIssue.Add(issuedItem)
                End If
            Next
        End If

        If itemsToIssue.Count = 0 Then
            MessageBox.Show("No tools found to issue for the selected criteria.", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        If MessageBox.Show($"Are you sure you want to issue {itemsToIssue.Count} tools?", "Confirm Issue", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
            Try
                Dim successCount = 0
                
                For Each item In itemsToIssue
                    ' 1. Fetch details from interchangeability since item only has limited info
                    Dim dtDetails = _mySql.ReadDatatable($"SELECT Department, InstrumentName, Color, SizeandRange FROM interchangeability WHERE ControlNo = '{item.ControlNo.Replace("'", "''")}' AND CycleName = '{_cycleName.Replace("'", "''")}' LIMIT 1")
                    
                    If dtDetails.Rows.Count > 0 Then
                        Dim dept = dtDetails.Rows(0)("Department").ToString()
                        Dim inst = dtDetails.Rows(0)("InstrumentName").ToString()
                        Dim color = dtDetails.Rows(0)("Color").ToString()
                        Dim size = dtDetails.Rows(0)("SizeandRange").ToString()
                        
                        ' 2. Use the centralized InsertInterchangeRecord which now automatically handles:
                        ' - Duplicate check (UPSERT)
                        ' - Interchangeability update
                        ' - Cycle History log
                        ' - Transaction table population (Issue/Receive)
                        If _mySql.InsertInterchangeRecord(_cycleName, item.ControlNo, dept, inst, size, color, "Issued", "Bulk Issued via Scan") Then
                            ' 3. Clear RFID from all tables
                            _mySql.ClearRFIDTag(item.ControlNo)
                            successCount += 1
                            _issued.Remove(item)
                        End If
                    End If
                Next

                MessageBox.Show($"Successfully issued {successCount} tools.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                UpdateCounts()
                
                ' We don't close automatically so user can see the final summary
            Catch ex As Exception
                MessageBox.Show("Error applying changes: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As RoutedEventArgs)
        _parentPage.ClearProcessedTags()
        _issued.Clear()
        _ignored.Clear()
        UpdateCounts()
    End Sub

    Private Sub BtnStop_Click(sender As Object, e As RoutedEventArgs)
        _parentPage.StopBulkScan()
        UpdateScanButtonStates(False)
    End Sub

    Private Sub BtnStartScan_Click(sender As Object, e As RoutedEventArgs)
        If _scanType = ScannerType.USB Then
            _parentPage.StartUsbPolling()
        Else
            _parentPage.StartGunScan()
        End If
        UpdateScanButtonStates(True)
    End Sub

    Private Sub UpdateScanButtonStates(isScanning As Boolean)
        BtnStartScan.IsEnabled = Not isScanning
        BtnStop.IsEnabled = isScanning
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub
End Class
