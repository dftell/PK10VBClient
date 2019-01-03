VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WXUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Uin As String
'Public UserName As String
'Public NikeName As String
'Public HeadImgUrl As String
'Public RemarkName As String
'Public PYInitial As String
'Public PYQuanPin As String
'Public RemarkPYQuanPin As String
'Public RemarkPYInitial As String
'Public HideInputBarFlag As String
'Public StarFriend As String
'Public Sex As String
'Public Signature As String
'Public AppAccountFlag As String
'Public VerifyFlag As String
'Public ContactFlag As String
'Public WebWxPluginSwitch As String
'Public HeadImgFlag As String
'Public SnsFlag As String
Public Info As Dictionary
Public JsonObj As Object

Public Sub LoadDataByText(str As String)
    On Error Resume Next
    Dim js As New JsonClass
    Set JsonObj = js.GetJsonVal(str, "")
End Sub
    
Public Sub LoadDataBuyJson(obj As Object)
    Set JsonObj = obj
End Sub

