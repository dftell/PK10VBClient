VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_Columns As Dictionary
Dim m_Rows() As Dictionary
Public Property Get ColumnItems() As Dictionary
    Set ColumnItems = m_Columns
End Property

Public Property Get RowCount() As Long
    RowCount = UBound(m_Rows)
End Property



Public Property Get Rows(index As Long) As Dictionary
    If index > 0 And index <= UBound(m_Rows) Then
        Set Rows = m_Rows(index)
    Else
       Set Rows = Nothing
    End If
End Property


Public Property Get Columns(index As Integer) As Dictionary
     If Not m_Columns(index) Then
        Set Columns = m_Columns(index)
    Else
       Set Columns = Nothing
    End If
End Property

Public Function Fill(res As Recordset) As Boolean
    On Error Resume Next
    Set m_Columns = New Dictionary
    If Not res Then
        For i = 0 To res.Fields.Count - 1
            If Not m_Columns.Exists(res.Fields(i).Name) Then m_Columns.Add res.Fields(i).Name, res.Fields(i)
        Next
    End If
    Dim cnt As Long
    While Not res.EOF
        cnt = cnt + 1
        ReDim Preserve m_Rows(cnt)
        Set m_Rows(cnt) = New Dictionary
        Dim col As Integer
        Dim c As Variant
        For Each cccc In m_Columns.Keys
            If IsEmpty(cccc) Then
                test = ""
                'MsgBox c
                Exit Function
            End If
            If Not m_Rows(cnt).Exists(cccc) Then
                m_Rows(cnt).Add cccc, res(col).Value
                col = col + 1
            End If
        Next
        col = 0
        Err.Clear
        res.MoveNext
        If Err.Number <> 0 Then
            'MsgBox Err.Description
            test = ""
            Exit Function
        End If
    Wend
End Function

Public Function FillByArray(Properties As String, objarr, Optional StartIndex As Long, Optional excludeCols As String = "")
    Dim strArr() As String
    strArr = Split(Properties, ",")
    Set m_Columns = New Dictionary
    Dim i As Long
    For i = 0 To UBound(strArr)
        Dim str As String
        str = strArr(i)
        If Not m_Columns.Exists(str) Then m_Columns.Add str, str
    Next
    
    Dim excludeDic As New Dictionary
    Dim exclArr() As String
    If Len(Trim(excludeCols)) > 0 Then
        exclArr = Split(excludeCols, ",")
        For i = 0 To UBound(exclArr)
            If (Not excludeDic.Exists(exclArr(i))) Then
                excludeDic.Add exclArr(i), exclArr(i)
            End If
        Next
    End If
    If UBound(objarr) <= 0 Then Exit Function
    ReDim m_Rows(UBound(objarr))
    For i = StartIndex To UBound(objarr)
        Dim obj As Object
        Set obj = objarr(i)
        Dim c As Variant
        Set m_Rows(i) = New Dictionary
        For Each c In m_Columns.Keys
            If Not excludeDic.Exists(c) Then m_Rows(i).Add c, CallByName(obj, c, VbGet)
        Next
    Next
    
    
End Function




Private Sub Class_Initialize()
    Set m_Columns = New Dictionary
    ReDim m_Rows(0)
End Sub
