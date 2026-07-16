Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class DepartmentSelectionWindow
    Public Property SelectedDepartments As New List(Of String)
    Private _deptList As New ObservableCollection(Of DepartmentItem)

    Public Sub New(allDepartments As List(Of String), alreadySelected As List(Of String))
        InitializeComponent()

        For Each dept In allDepartments
            _deptList.Add(New DepartmentItem With {
                .Name = dept,
                .IsSelected = alreadySelected.Contains(dept)
            })
        Next

        DeptListBox.ItemsSource = _deptList
    End Sub

    Private Sub BtnSelectAll_Click(sender As Object, e As RoutedEventArgs)
        For Each item In _deptList
            item.IsSelected = True
        Next
    End Sub

    Private Sub BtnClearAll_Click(sender As Object, e As RoutedEventArgs)
        For Each item In _deptList
            item.IsSelected = False
        Next
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As RoutedEventArgs)
        SelectedDepartments = _deptList.Where(Function(d) d.IsSelected).Select(Function(d) d.Name).ToList()
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Public Class DepartmentItem
        Implements INotifyPropertyChanged

        Private _name As String
        Private _isSelected As Boolean

        Public Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
                OnPropertyChanged("Name")
            End Set
        End Property

        Public Property IsSelected As Boolean
            Get
                Return _isSelected
            End Get
            Set(value As Boolean)
                _isSelected = value
                OnPropertyChanged("IsSelected")
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Protected Sub OnPropertyChanged(name As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
        End Sub
    End Class
End Class
