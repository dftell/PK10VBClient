VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpectData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public expect As String
Public OpenTime As Date
Public OpenCode As String
Public db As New DBClass


Public Property Get ExpectIndex() As String
    v = expect
    If expect >= 344978 Then 'Miss 34977
        v = v - 1
    End If
    If expect >= 356395 Then '2013/4/16 356375-356395 system no exist data
        v = v - 19
    End If
''    If Expect >= 676963 Then '2018/4/17 676957-676962 systm no exist data
''        v = v - 6
''    End If
    ExpectIndex = v
End Property
Public Sub DeleteDayData(day As Date)
    Dim sql As String
    sql = Replace("delete from Newestdata where [opentime]>='XX'", "XX", day)
    db.ExcelSql (sql)
End Sub
Public Function getDayData(day As Date, Optional lastExpect As String) As DataTable
    On Error Resume Next
    Dim sql As String
    Dim ret As DataTable
    
    sql = Replace("select * from Newestdata where [opentime]>'XX' order by Expect desc", "XX", day)
    If Len(lastExpect) > 0 Then
        sql = Replace(Replace("select * from Newestdata where [opentime]>'XX' and [Expect] <= 'YY' order by Expect desc", "XX", day), "YY", lastExpect)
    End If
    Set ret = db.getDataBySql(sql)
    Set getDayData = ret
End Function

Public Sub TableToArray(dt As DataTable, ByRef arr() As ExpectData)
    ReDim arr(0)
    If dt Is Nothing Then Exit Sub
    ReDim arr(dt.RowCount)
    Dim i As Long
    For i = 1 To dt.RowCount
        Dim dic As Dictionary
        Set dic = dt.Rows(i)
        If dic Is Nothing Then
            Exit Sub
        End If
        ReDim Preserve arr(i)
        Set arr(i) = New ExpectData
        arr(i).expect = dic("Expect")
        arr(i).OpenCode = dic("OpenCode")
        arr(i).OpenTime = dic("OpenTime")
    Next
    Exit Sub
End Sub

Public Function ArrayToTable(arr() As ExpectData) As DataTable
    Dim ret As New DataTable
    ret.FillByArray "Expect,OpenCode,OpenTime", arr, 1
    Set ArrayToTable = ret
End Function

Public Function SaveNewData(arr() As ExpectData) As Boolean
    Dim dt As DataTable
    Set dt = ArrayToTable(arr)
    db.AddNewDataBuySql "select * from NewestData", dt
End Function

Public Function SaveHistoryData(arr() As ExpectData) As Boolean
    Dim dt As DataTable
    Set dt = ArrayToTable(arr)
    db.AddNewDataBuySql "select * from HistoryData", dt
End Function

Public Function GetHistoryData(oldestExpect As String, Optional topN As String, Optional orderby As String = "desc") As DataTable
    Dim sql As String
    Dim ret As DataTable
    If Len(Trim(oldestExpect)) = 0 Then oldestExpect = "0"
    If Len(Trim(topN)) > 0 Then
        sql = Replace("select top " & topN & " * from Historydata where [Expect]> XX order by Expect " & orderby, "XX", oldestExpect)
    Else
        sql = Replace("select  * from Historydata where [Expect]> XX order by Expect " & orderby, "XX", oldestExpect)
    End If
    Set ret = db.getDataBySql(sql)
    Set GetHistoryData = ret
End Function
