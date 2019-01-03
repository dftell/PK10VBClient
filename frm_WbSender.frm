VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm_WbSender 
   Caption         =   "Form4"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15135
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   15135
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   930
      Width           =   14955
      ExtentX         =   26379
      ExtentY         =   13996
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
   Begin VB.Label lbl_ExecStatus 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9075
   End
End
Attribute VB_Name = "frm_WbSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strurl  As String '= "https://www.kcai{host}.com"
Dim strHost As String '= "331"
Dim strHostList As String
''Const strurl = "http://mem1.naewno216.{host}.com:88"
''Const strHost = "zyxghl"
Public gobj As SystemClass
Public parentForm As Form1
Public doc As HTMLDocument
Public strJscript As String
Dim cm As New ClassModel
Dim buffScript As String
Private Sub Form_Load()
    
    strurl = gobj.LoginUrlModel
    strHost = gobj.LoginDefaultHost
    'Me.WebBrowser1.Silent = True
    Me.WebBrowser1.Navigate Replace(strurl, "{host}", strHost)
    Me.Caption = gobj.ClientUserName
    
End Sub


Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    
    If WebBrowser1.readyState < READYSTATE_COMPLETE Then
        Exit Sub
    End If
    cm.gobj = gobj
    Dim obj As IUnknown
    Set obj = IUnknown
    If (pDisp Is Me.WebBrowser1.object) Then
    Else
        'MsgBox pDisp
    End If

    WebBrowser1.RegisterAsBrowser = True
    If WebBrowser1.LocationURL Then
        Set doc = WebBrowser1.Document
    End If
    'End If
    If doc Is Nothing Then Exit Sub
    Dim txtUser As Object
    Set txtUser = doc.getElementById("loginName")
    If txtUser Is Nothing Then Exit Sub
    
    Dim strLogin As String
    'function Login(strUserName,strPassword)
    Dim Purl As String
    Purl = strurl
    Purl = Replace(Purl, "http://", "")
    Purl = Replace(Purl, "https://", "")
    Purl = Replace(Purl, "/", "")
    Purl = Replace(Purl, ":", "")
     Dim fobj As New FileSystemObject
    Dim stmText As TextStream
    Dim strJson As String
    Set stmText = fobj.OpenTextFile(App.Path & "\" & Purl & "_pure.js")
    strJscript = stmText.ReadAll()
    strLogin = "{0};Login('{1}','{2}')"
    strJsAll = Replace(strLogin, "{0}", strJscript)
    strJsAll = Replace(strJsAll, "{1}", gobj.ClientUserName)
    strJsAll = Replace(strJsAll, "{2}", gobj.ClientPassword)
    If InStr(1, strurl, "kcai") > 0 Then
        doc.parentWindow.execScript strJsAll
    End If
    If InStr(1, strurl, "mg") > 0 Then
        WebBrowser1.Silent = True
    End If
    Exit Sub
    
    Dim objname As IHTMLElement
    Dim objpwd As IHTMLElement
    Dim objbtn As IHTMLElement
    Set objname = doc.getElementById("txt_username")
    Set objpwd = doc.getElementById("txt_pwd")
    Set objbtn = doc.getElementById("login-submit-Button")
    If objname Is Nothing Or objpwd Is Nothing Then Exit Sub
    objname.setAttribute "value", gobj.ClientUserName
    objpwd.setAttribute "value", gobj.ClientPassword
    'objbtn.onclick
End Sub

Public Function SendMsg(Expect As String, msg As String, cnt As Integer, Optional ByRef amt As Currency = 0) As Boolean
     On Error Resume Next
     If doc Is Nothing Then
        Exit Function
     End If
         buffScript = ""
     Dim fobj As New FileSystemObject
    Dim stmText As TextStream
    Dim URL As String
    URL = strurl
    URL = Replace(URL, "http://", "")
    URL = Replace(URL, "https://", "")
    URL = Replace(URL, ":", "")
    URL = Replace(URL, "/", "")
    Set stmText = fobj.OpenTextFile(App.Path & "\" & URL & "_pure.js")
    strJscript = stmText.ReadAll()
    stmText.Close
    Dim c2i As New CCS2InstrClass
    c2i.cJsOdds = Me.gobj.Odds
    Dim strJson As String
    If gobj.LoginInstFillOrEnCode = 1 Then '填充模式
        strJson = cm.ToSerial(msg, Me.gobj.minChips)
    Else '编码模式
        strJson = c2i.InstrToJsonString(msg) 'Trim(Me.parentForm.cm.ToSerial(msg))
    End If
    Dim strJsAll As String
    strJsAll = Replace(Replace(Replace("{0};SendMsg('{1}','{2}')", "{0}", strJscript), "{1}", Expect), "{2}", Trim(strJson))
    'Me.WebBrowser1.RegisterAsBrowser = True
     
     'Set doc = Me.WebBrowser1.Document
    
     'MsgBox doc.body.innerHTML
    'Me.WebBrowser1.Document.body.innerHTML ' doc.parentWindow.Parent.Document.body.innerHTML
     'Me.gobj.LogObj.Log doc.body.innerHTML
     If gobj.LoginInFrame Then
        If doc.parentWindow.frames("mainFrame") Is Nothing Then
        Else
          Set doc = doc.parentWindow.frames("mainFrame").Document
        End If
     End If
         buffScript = strJsAll
     Me.WebBrowser1.Silent = True
     doc.parentWindow.execScript buffScript
     'Me.WebBrowser1.Silent = False
     Dim objPoint As IHTMLElement
     Set objPoint = doc.getElementById("userGamePointId")
     If objPoint Is Nothing Then Exit Function
     amt = objPoint.getAttribute("data")
     Me.Show
     
End Function


Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    
   Cancel = True
   doc.parentWindow.execScript buffScript
End Sub
