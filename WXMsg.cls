VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WXMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FromUserName As String
Public ToUserName As String
Public Msg As String
Public Content As String
Public MsgType As String
Public Items As Dictionary
Public Speaker As String
Public Function LoadMsgsByText(str As String) As WXMsg()
    On Error Resume Next
    Dim strMsgs As String
    Dim Msg() As String
    Dim ret() As WXMsg
    str = Replace(str, Chr(13), "")
    str = Replace(str, Chr(10), "")
    Dim obj As Object
    Dim jc As New JsonClass
    Dim msgcnt As Long
    msgcnt = jc.GetJsonVal(str, "AddMsgCount")
    
    Set obj = jc.GetJsonVal(str, "")
    Dim msgobjs() As Object
    jc.ToArray obj, "AddMsgList", msgobjs
    
    For i = 1 To UBound(msgobjs)
        Dim mobj As New WXMsg
        mobj.LoadMsg msgobjs(i)
        ReDim Preserve ret(i)
       Set ret(i) = New WXMsg
       ret(i).FromUserName = mobj.FromUserName
       ret(i).ToUserName = mobj.ToUserName
       ret(i).Msg = mobj.Msg
       ret(i).MsgType = mobj.MsgType
    Next
    LoadMsgsByText = ret
End Function

Public Sub LoadMsg(uobj As Object)
    On Error Resume Next
    Dim jc As New JsonClass
    'FromUser = Trim(uobj.FromUsername)
    FromUserName = uobj.FromUserName
    ToUserName = Trim(uobj.ToUserName)
    Msg = jc.RemoveSpeaker(jc.RemoveHTML(uobj.Content))
    Content = uobj.Content
    MsgType = uobj.MsgType
    Set JsonObj = uobj
End Sub




