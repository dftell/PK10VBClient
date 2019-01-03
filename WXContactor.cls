VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WXContactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Info As Dictionary
Public MemberList As Dictionary
Public JsonObj As Object
Public Sub LoadDataByText(obj As Object)
    Dim wxu As New WXUser
    Dim str As String
    wxu.LoadDataByText str '±‡“Î ß∞‹
    Set JsonObj = wxu.JsonObj
End Sub

Public Function LoadListByText(str As String) As Dictionary
    On Error Resume Next
    Dim strList As String
    Dim ret As New Dictionary
    Dim js As New JsonClass
'    strList = Replace(strList, Chr(10), "")
'    strList = Replace(strList, Chr(13), "")
'    strList = js.RemoveHTML(strList)
    Dim users() As WXContactor
    Dim userobjs() As Object
    Dim UsersObj As Object
    Set UsersObj = js.GetJsonVal(str, "")
    js.ToArray UsersObj, "MemberList", userobjs
    Set FriendList = New Dictionary
    For i = 1 To UBound(userobjs)
        Dim wxct As New WXContactor
        wxct.LoadDataBuyJson userobjs(i)
        ReDim Preserve users(i)
        Set users(i) = New WXContactor
        users(i).JsonObj = wxct.JsonObj
        If Not ret.Exists(wxct.JsonObj.UserName) Then
            ret.Add wxct.JsonObj.UserName, wxct
        End If
        Set wxct = Nothing
    Next
    Set LoadListByText = ret
End Function

Public Sub LoadDataBuyJson(obj As Object)
    Set JsonObj = obj
    Dim js As New JsonClass
    JsonObj.NickName = js.RemoveHTML(JsonObj.NickName)
    JsonObj.UserName = Trim(JsonObj.UserName)
End Sub

