VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const FileName = "\策略交易终端日志.log"
Const OutName = "c:\inetpub\wwwroot\record.txt"
Const ExpectName = "c:\inetpub\wwwroot\expertNo.txt"
Public Sub Log(txt As String)
    Path = App.Path & FileName
    Dim fobj As New FileSystemObject
    Dim txtstr As TextStream
    Set txtstr = fobj.OpenTextFile(Path, ForAppending, True)
    txtstr.WriteLine Now() & ":" & txt
    txtstr.Close
    Set fobj = Nothing
    Set txtstr = Nothing
End Sub


Public Sub OutRecord(expectNo As String, txt As String)
    Path = OutName
    Dim fobj As New FileSystemObject
    Dim txtstr As TextStream
    Set txtstr = fobj.OpenTextFile(Path, ForAppending, True)
    txtstr.WriteLine expectNo & ":" & txt
    txtstr.Close
    Set fobj = Nothing
    Set txtstr = Nothing
End Sub

Public Sub OutExpectNo(txt As String)
    Path = ExpectName
    Dim fobj As New FileSystemObject
    Dim txtstr As TextStream
    Set txtstr = fobj.OpenTextFile(Path, ForWriting, True)
    txtstr.WriteLine txt
    txtstr.Close
    Set fobj = Nothing
    Set txtstr = Nothing
End Sub




