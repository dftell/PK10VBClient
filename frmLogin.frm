VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登录"
   ClientHeight    =   3930
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2321.974
   ScaleMode       =   0  'User
   ScaleWidth      =   9506.824
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "执行服务"
      Height          =   435
      Left            =   6930
      TabIndex        =   11
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton txt_TestInst 
      Caption         =   "测试指令"
      Height          =   465
      Left            =   6840
      TabIndex        =   10
      Top             =   1200
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3870
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   9
      Top             =   930
      Width           =   975
   End
   Begin VB.TextBox txtValidCode 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   930
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   1395
      Left            =   750
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "frmLogin.frx":0000
      Top             =   1860
      Width           =   9315
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Text            =   "user331"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   390
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   390
      Left            =   2250
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "abcdef123"
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "验证码"
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户名称(&U):"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码(&P):"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Public LoginSucceeded As Boolean
Public InterFace As FormInterface
Public cookie As String
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Public gobj As SystemClass
Private Type TGUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type
Public parentForm As Form1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function LoadPicture(ByVal strFileName As String) As Picture
   Dim IID As TGUID
   With IID
      .Data1 = &H7BF80980
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(2) = &H0
      .Data4(3) = &HAA
      .Data4(4) = &H0
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
   End With
   On Error GoTo LocalErr
   OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
   Exit Function
LocalErr:
   Set LoadPicture = VB.LoadPicture(strFileName)
   Err.Clear
End Function

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    LoginSucceeded = False
    Me.Hide
End Sub

Public Sub cmdOK_Click()
    '检查正确的密码
    Dim InData As String
    InData = "{""userName"":""[UN]"",""password"":""[PD]"",""valicode"":""[VC]""}"
    InData = Replace(InData, "[UN]", Me.txtUserName.Text)
    InData = Replace(InData, "[PD]", Me.txtPassword.Text)
    InData = Replace(InData, "[VC]", Me.txtValidCode.Text)
    InterFace.InData = InData
'    frm.WebBrowser1.Silent = True
    InterFace.Translate
    'MsgBox InterFace.outdata
    Dim httpreq As New XMLHTTP60
    httpreq.Open "Post", "https://www.kcai" & CStr(gobj.Ip) & ".com/Login/PostLogin", False
    'httpreq.Open "Post", "https://www.kcai773.com/Login/PostLogin"
    Dim httpheads As String
    'httpreq.setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    httpreq.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    'httpreq.setRequestHeader "Content-type", "application/json, text/javascript, */*; q=0.01"
    httpreq.setRequestHeader "Host", "www.kcai" & CStr(gobj.Ip) & ".com/home"
    httpreq.setRequestHeader "Referer", "https://www.kcai" & CStr(gobj.Ip) & ".com"
    httpreq.setRequestHeader "Origin", "https://www.kcai" & CStr(gobj.Ip) & ".com"
    httpreq.setRequestHeader "X-Requested-With", "XMLHttpRequest" '
    httpreq.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134"
    'method.addHeader("Content-type","application/json; charset=utf-8");
    '            method.setHeader("Accept", "application/json");
    httpreq.setRequestHeader "Cookie", Me.cookie & ";HTTPOnly" 'InterFace.str__Cookie
    'httpreq.setRequestHeader "Set-Cookie", Me.cookie & ";HTTPOnly" 'InterFace.str__Cookie
    'httpreq.setRequestHeader "__RequestVerificationToken", "tAo354u2Zr48wyYpIX2XBRe4Qr0IEiIbbU_NRyZqvVXwF1llRuyKmjB4Q0tHeZ2nTv4n9v2SY2XVw8laBsjg49DcooPyWFp7IC011JHUImLhAJVbDd"
    httpreq.send InterFace.outdata & "&__RequestVerificationToken=" & InterFace.str__RequestVerificationToken
    
    'cookie = httpreq.getResponseHeader("Set-Cookie")
    If httpreq.Status <> 200 Then
        Me.Text1.Text = InterFace.outdata & "&__RequestVerificationToken=" & InterFace.str__RequestVerificationToken & Chr(13) _
        & "[cookie]" & InterFace.str__Cookie & Chr(13) _
        & "[header]:" & httpheads & Chr(13) & httpreq.responseText
        
        Exit Sub
    End If
    httpheads = httpreq.getAllResponseHeaders()
    Me.Text1.Text = ""
    'MsgBox cookie
    If Len(Trim(cookie)) = 0 Then
        Me.Text1.Text = httpheads & Chr(13) & httpreq.responseText
    End If
    Dim strRet As String
    strRet = httpreq.responseText
    If strRet <> """suc""" Then
        Picture1_Click
        Exit Sub
    End If
     Me.Text1.Text = httpreq.responseText & "[cookie]" & InterFace.str__Cookie & Chr(13) _
        & "[header]:" & httpheads & Chr(13) & httpreq.responseText
        
    InterFace.str__Cookie = cookie
    'Me.Hide
    txt_TestInst_Click
    'Unload Me
    Exit Sub
    
    
    
    
    If txtPassword = "password" Then
        '将代码放在这里传递
        
        '成功到 calling 函数
        '设置全局变量时最容易的
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "无效的密码，请重试!", , "登录"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Command1_Click()
    Dim frm As New Form1
    frm.Show
    Set gobj = frm.Gobalobj
End Sub

Public Sub Form_Load()
    Dim defaultIp As String
    defaultIp = InputBox("输入网站", "WEB SITE", "773")
    If gobj Is Nothing Then
        Set gobj = New SystemClass
        gobj.Ip = CInt(defaultIp)
    Else
        If gobj.Ip = 0 Then gobj.Ip = CInt(defaultIp)
    End If
    Set InterFace = New FormInterface
    Me.Caption = "登录=>" & gobj.Ip
    Set InterFace.parentForm = Me
    InterFace.Ip = gobj.Ip
    Me.cmdOK.Enabled = False
    InterFace.WebBrowser1.Navigate "https://www.kcai" & CStr(gobj.Ip) & ".com"
    Sleep 2000
    Me.MousePointer = vbHourglass
    'Set Me.Picture1 = Me.LoadPicture("https://www.kcai" & CStr(gobj.Ip) & ".com/Login/ValidateCode?r=" & Math.Rnd())
End Sub

Private Sub Picture1_Click()
    Set Me.Picture1 = Me.LoadPicture("https://www.kcai" & CStr(gobj.Ip) & ".com/Login/ValidateCode?r=" & Math.Rnd())
End Sub

Private Sub txt_TestInst_Click()
    Dim frm As frm_Test
    If Me.parentForm.ExchangeForm Is Nothing Then
        Set frm = New frm_Test
    Else
        Set frm = Me.parentForm.ExchangeForm
    End If
    Set frm.InterFace = Me.InterFace
    Set frm.LoginForm = Me
    Set frm.parentForm = Me.parentForm
    Set frm.parentForm.ExchangeForm = frm
    frm.Caption = "交易=>" & gobj.Ip
    frm.cookie = Me.cookie
    If Me.gobj Is Nothing Then
    Else
        Set frm.gobj = Me.gobj
    End If
    frm.Show
End Sub
