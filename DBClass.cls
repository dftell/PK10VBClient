VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Const ConnString = "provider=microsoft.ace.oledb.12.0;data source=[XX]\data.accdb;Persist Security Info=False;"
'Data Source=190.190.200.100,1433;Network Library=DBMSSOCN;
'Initial Catalog=myDataBase;User ID=myUsername;Password=myPassword;
'Const ConnString = "provider=SQLNCLI10;server=iZwslososrxxsfZ\sqlexpress;Database=PK10DB;UId=f_chinazhou;PWD=bolts;"
'Const ConnString = "Server=47.95.222.142\SQLEXPRESS;Database=PK10DB;User Id=sa;Password =bolts;"
'Const ConnString = "data source=47.95.222.142\SQLEXPRESS;Initial Catalog=pk10db;User ID=sa;Password=bolts"
'Const ConnString = "Data Source=47.95.222.142\SQLEXPRESS;Initial Catalog=PK10DB;Persist Security Info=True;User ID=sa;Password=bolts;"
'
'Const ConnString = "Provider=SQLOLEDB.1;Server=47.95.222.142\SQLEXPRESS;Persist Security Info=True;User ID=sa;Password=bolts;Initial Catalog=pk10db" 'SQLOLEDB.1
'Const ConnString = "Provider=SQLNCLI.1;Server=47.95.222.142;User ID=sa;Password=bolts;Initial Catalog=pk10db" 'SQLOLEDB.1
'Const ConnString = "" '"DSN=pk10db;UID=sa;Pwd=bolts;"
Const ConnString = "Provider=SQLNCLI.1;Persist Security Info=False;Server=iZwslososrxxsfZ\SQLEXPRESS;User ID=sa;Password=bolts;Initial Catalog=pk10db;"
Public gobj As SystemClass
Public Function getDataBySql(strsql As String) As DataTable
    On Error Resume Next
    Dim conn As New ADODB.Connection
    conn.ConnectionString = Replace(ConnString, "[XX]", App.Path)
    Err.Clear
    conn.Open conn.ConnectionString
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        MsgBox Err.Description
        Exit Function
    End If
    Dim res As New Recordset
    Set res.ActiveConnection = conn
    res.Open strsql, conn, adOpenStatic, adLockReadOnly, adCmdText
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If
    Dim dt As New DataTable
    'gobj.LogObj.Log "Ìî³äµ½±í¸ñ"
    dt.Fill res
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If
    res.Close
    conn.Close
    Set getDataBySql = dt
End Function

Public Function ExcelSql(strsql As String, Optional ByRef cnt = -1) As Boolean
    On Error Resume Next
    Dim conn As New ADODB.Connection
    conn.ConnectionString = Replace(ConnString, "[XX]", App.Path)
    conn.Open conn.ConnectionString
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If

    Err.Clear
    Dim testcnt As Long

    conn.Execute strsql, testcnt
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If
    If cnt <> -1 Then
        cnt = testcnt
    End If
    conn.Close
    rec.Close
    Set conn = Nothing
    ExcelSql = True
End Function

Public Function UpdateDataBuySql(strsql As String, dt As DataTable, Optional UIDName As String) As Long
    On Error Resume Next
    Dim conn As New ADODB.Connection
    conn.ConnectionString = Replace(ConnString, "[XX]", App.Path)
    conn.Open conn.ConnectionString
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If
    Dim res As New Recordset

    res.Open strsql, conn, 1, 2
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If
    Dim i As Long
    Dim c As Variant
    Dim col As Integer
    If res.RecordCount = 0 Then
        For i = 1 To dt.RowCount
            res.AddNew
            col = 0
            For Each c In dt.Rows(i).Keys
                If c <> UIDName Then
                    res(c) = dt.Rows(i)(c)
                    col = col + 1
                Else
                    Dim test As String
                    test = "test"
                End If
            Next
            col = 0
            res.Update
            res.MoveNext
        Next
    Else
        i = 1
        While Not res.EOF
            col = 0
            For Each c In dt.Rows(i).Keys
                If c <> UIDName Then
                    res(c) = dt.Rows(i)(c)
                    col = col + 1
                End If
            Next
            i = i + 1
            res.Update
            res.MoveNext
        Wend
    End If
    'res.UpdateBatch adAffectCurrent
    res.Close
    Set res = Nothing
    conn.Close
    Set conn = Nothing
End Function

Public Function AddNewDataBuySql(strsql As String, dt As DataTable) As Long
    On Error Resume Next
    Dim conn As New ADODB.Connection
    conn.ConnectionString = Replace(ConnString, "[XX]", App.Path)
    conn.Open conn.ConnectionString
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If
    Dim res As New Recordset
    'Set res.ActiveConnection = conn
    'res.CacheSize = 1000
    'res.CursorType = adUseClient
    res.Open strsql, conn, 1, 3
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
        
    End If
    Dim i As Long
    For i = 1 To dt.RowCount
        Dim c As Variant
        res.AddNew
        Dim col As Integer
        For Each c In dt.Rows(i).Keys
            res(col) = dt.Rows(i)(c)
            col = col + 1
        Next
        col = 0
        res.Update
        res.MoveNext
    Next
    'res.UpdateBatch adAffectCurrent
    If Err.Number <> 0 Then
        gobj.LogObj.Log Err.Description
        Exit Function
    End If
    
    res.Close
    Set res = Nothing
    conn.Close
End Function
