VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataGridClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public grid As MSFlexGrid
Public Event ChangeDate(data As DataTable)
Public Sub FillGrid(data As DataTable)
    If grid Is Nothing Then
        Exit Sub
    End If
    If data Is Nothing Then
        Exit Sub
    End If
    If data.ColumnItems.Count = 0 Then Exit Sub
    grid.Clear
    grid.Cols = data.ColumnItems.Count + 1
    grid.Rows = data.RowCount + 1
    grid.CellBackColor = grid.Parent.Gobalobj.BackColor
    Dim c As Variant
    Dim i As Integer
    Dim j As Long
    For Each c In data.ColumnItems.Keys
        grid.TextMatrix(0, i + 1) = c
        i = i + 1
    Next
    For j = 1 To data.RowCount
        i = 0
        For Each c In data.ColumnItems.Keys
            grid.TextMatrix(j, i + 1) = IIf(IsNull(data.Rows(j)(c)), "", data.Rows(j)(c))
            i = i + 1
        Next
        i = 0
    Next
    AdjustColWidth grid.Parent, grid
End Sub

'//自动调整Grid各列列宽为最合适的宽度
Public Sub AdjustColWidth(frmCur As Form, gridCur As Object, Optional bNullRow As Boolean = True, Optional dblIncWidth As Double = 0)
    '--------------------------------------------------------------------
    '功能:
    ' 自动调整Grid各列列宽为最合适的宽度
    '参数:
    ' [frmCur].........................................当前工作窗体
    ' [gridCur]........................................当前要调整的Grid
    '--------------------------------------------------------------------
    Dim i, j As Integer
    Dim dblWidth As Double
    
    With gridCur
    For i = 0 To .Cols - 1
        dblWidth = 0
        If .ColWidth(i) <> 0 Then
        For j = 0 To .Rows - 1
            If frmCur.TextWidth(.TextMatrix(j, i)) > dblWidth Then
                dblWidth = frmCur.TextWidth(.TextMatrix(j, i))
            End If
        Next
        .ColWidth(i) = dblWidth + dblIncWidth + 100
        End If
    Next
    End With
End Sub
