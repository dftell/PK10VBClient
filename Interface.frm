VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FormInterface 
   Caption         =   "接口"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9690
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9690
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7065
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   9285
      ExtentX         =   16378
      ExtentY         =   12462
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "FormInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public InData As String
Public outdata As String
Public RowCount As Integer
Public Ip As Integer
Public str__RequestVerificationToken As String
Public str__Cookie As String
Public parentForm As frmLogin
Public expectNo As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim NeedAsp As Boolean
Private Sub Form_Load()
    'WebBrowser1.Navigate "https://www.kcai773.com"
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pdisp As Object, url As Variant)
    WebBrowser1.Silent = True
    If pdisp Is WebBrowser1.object Then '加载完毕
        If InStr(WebBrowser1.LocationURL, "kcai") > 0 Then
            Dim obj As Object
            Dim htmlobj As New DHTMLPage
            Dim objs As IHTMLElementCollection
            Dim ReqVT As IHTMLElement
            str__Cookie = WebBrowser1.Document.cookie
            parentForm.cookie = str__Cookie 'Replace(str__Cookie, " ", "")
            Set objs = WebBrowser1.Document.getElementsByName("__RequestVerificationToken")
            If objs.Length = 1 Then
                Set ReqVT = objs(0)
                Me.str__RequestVerificationToken = ReqVT.Value
                parentForm.cookie = parentForm.cookie & ";__RequestVerificationToken=" & str__RequestVerificationToken
                WebBrowser1.Navigate "http://localhost/model.asp?Ip=" & CStr(Ip)
                'WebBrowser1.Refresh
            End If
        Else
            '响应
            'WebBrowser1.Navigate2 "http://localhost/model.asp?Ip=" & CStr(Ip), Null, Null, Null, Null
            'Sleep 1000
            'If NeedAsp = False Then
                parentForm.MousePointer = 1
                parentForm.cmdOK.Enabled = True
                parentForm.cmdOK_Click
            'Else
            '    Translate1
            '    NeedAsp = False
            'End If
        End If
    End If
End Sub
Public Sub Translate()
     NeedAsp = True
     Translate1
     NeedAsp = False
     'WebBrowser1.Navigate "http://localhost/model.asp?Ip=" & CStr(Ip), Null, Null, Null, Null
End Sub

Public Sub Translate1()
    On Error Resume Next
     'Sleep 1000
     If Len(Trim(InData)) > 0 Then
            Dim InCtrl As Object
            Dim RowCountCtrl As Object
            Dim OutCtrl As Object
            Dim BtnCtrl As Object
            Dim ctrl As Object
            
            For Each ctrl In WebBrowser1.Document.All
                If ctrl.tagName = "TEXTAREA" And CallByName(ctrl, "id", VbGet) = "InputData" Then
                    Set InCtrl = ctrl
                End If
                If ctrl.tagName = "INPUT" And CallByName(ctrl, "id", VbGet) = "RowCount" Then
                    Set RowCountCtrl = ctrl
                End If
                If ctrl.tagName = "TEXTAREA" And CallByName(ctrl, "id", VbGet) = "OutputData" Then
                    Set OutCtrl = ctrl
                End If
            Next
            If InCtrl Is Nothing Then Exit Sub
            CallByName InCtrl, "value", VbLet, InData
            CallByName RowCountCtrl, "value", VbLet, CStr(RowCount)
            WebBrowser1.Document.parentWindow.execScript "Translate()"
            outdata = OutCtrl.Value
    End If
End Sub

Public Sub TranslateInst()
     On Error Resume Next
     If Len(Trim(InData)) > 0 Then
            Dim InCtrl As Object
            Dim OutCtrl As Object
            Dim ReqVerCtrl As Object
            Dim ExpectCtrl As Object
            Dim RowCountCtrl As Object
            Dim ctrl As Object
            For Each ctrl In WebBrowser1.Document.All
                If ctrl.tagName = "INPUT" And CallByName(ctrl, "id", VbGet) = "__RequestVerificationToken" Then
                    Set ReqVerCtrl = ctrl
                End If
                If ctrl.tagName = "TEXTAREA" And CallByName(ctrl, "id", VbGet) = "InputData" Then
                    Set InCtrl = ctrl
                End If
                If ctrl.tagName = "TEXTAREA" And CallByName(ctrl, "id", VbGet) = "OutputData" Then
                    Set OutCtrl = ctrl
                End If
                If ctrl.tagName = "INPUT" And CallByName(ctrl, "id", VbGet) = "ExpertNo" Then
                    Set ExpectCtrl = ctrl
                End If
                If ctrl.tagName = "INPUT" And CallByName(ctrl, "id", VbGet) = "RowCount" Then
                    Set RowCountCtrl = ctrl
                End If
            Next
            CallByName ReqVerCtrl, "value", VbLet, str__RequestVerificationToken
            CallByName InCtrl, "value", VbLet, InData
            CallByName ExpectCtrl, "value", VbLet, Me.expectNo
            CallByName RowCountCtrl, "value", VbLet, RowCount
            WebBrowser1.Document.parentWindow.execScript "TranslateInst()"
            outdata = OutCtrl.Value
    End If
End Sub
