Imports System.Data
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class TypeSelectionWindow
    Public Property SelectedTypes As New List(Of String)
    Private _instList As New ObservableCollection(Of TypeItem)
    Private _gaugeList As New ObservableCollection(Of TypeItem)

    Public Sub New(allTypes As DataTable, alreadySelected As List(Of String))
        InitializeComponent()

        For Each row As DataRow In allTypes.Rows
            Dim typeName = row("TypeName").ToString()
            Dim category = row("Category").ToString()
            
            Dim item As New TypeItem With {
                .Name = typeName,
                .IsSelected = alreadySelected.Contains(typeName)
            }

            If category.Equals("Instrument", StringComparison.OrdinalIgnoreCase) Then
                _instList.Add(item)
            Else
                _gaugeList.Add(item)
            End If
        Next

        InstListBox.ItemsSource = _instList
        GaugeListBox.ItemsSource = _gaugeList
    End Sub

    Private Sub BtnSelectAllInst_Click(sender As Object, e As RoutedEventArgs)
        For Each item In _instList
            item.IsSelected = True
        Next
    End Sub

    Private Sub BtnClearAllInst_Click(sender As Object, e As RoutedEventArgs)
        For Each item In _instList
            item.IsSelected = False
        Next
    End Sub

    Private Sub BtnSelectAllGauge_Click(sender As Object, e As RoutedEventArgs)
        For Each item In _gaugeList
            item.IsSelected = True
        Next
    End Sub

    Private Sub BtnClearAllGauge_Click(sender As Object, e As RoutedEventArgs)
        For Each item In _gaugeList
            item.IsSelected = False
        Next
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As RoutedEventArgs)
        SelectedTypes.AddRange(_instList.Where(Function(d) d.IsSelected).Select(Function(d) d.Name))
        SelectedTypes.AddRange(_gaugeList.Where(Function(d) d.IsSelected).Select(Function(d) d.Name))
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Public Class TypeItem
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
