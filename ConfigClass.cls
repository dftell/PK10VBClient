VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public theForm As Form
Public ConfigType As String
Public Name As String
Public Value As String
Dim frm As Frame
Const ctrlname = "txt_config_[AA]_[BB]"
Public Property Get Contrl() As TextBox
    On Error Resume Next
    Dim strname As String
    strname = Replace(Replace(ctrlname, "[AA]", ConfigType), "[BB]", Name)
    Err.Clear
    Set Contrl = CallByName(theForm, strname, VbGet)
    If Err.Number <> 0 Then
        'MsgBox strname
        Dim i As String
        i = "test"
    End If
End Property

