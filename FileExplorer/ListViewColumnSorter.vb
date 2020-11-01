Imports System.Collections
Imports System.Windows.Forms

Public Class ListViewColumnSorter
    Implements IComparer
    Private ColumnToSort As Integer
    Private OrderOfSort As SortOrder
    Private ObjectCompare As CaseInsensitiveComparer

    Public Sub New()
        ColumnToSort = 0
        OrderOfSort = SortOrder.None
        ObjectCompare = New CaseInsensitiveComparer()
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim compareResult As Integer
        Dim listviewX, listviewY As ListViewItem
        listviewX = CType(x, ListViewItem)
        listviewY = CType(y, ListViewItem)
        compareResult = ObjectCompare.Compare(listviewX.SubItems(ColumnToSort).Text, listviewY.SubItems(ColumnToSort).Text)
        If OrderOfSort = SortOrder.Ascending Then
            Return compareResult
        ElseIf OrderOfSort = SortOrder.Descending Then
            Return (-compareResult)
        Else
            Return 0
        End If
    End Function

    Public Property SortColumn As Integer
        Set(ByVal value As Integer)
            ColumnToSort = value
        End Set
        Get
            Return ColumnToSort
        End Get
    End Property

    Public Property Order As SortOrder
        Set(ByVal value As SortOrder)
            OrderOfSort = value
        End Set
        Get
            Return OrderOfSort
        End Get
    End Property
End Class
