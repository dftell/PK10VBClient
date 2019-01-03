VERSION 5.00
Begin VB.Form frm_Test 
   Caption         =   "Form4"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12000
   LinkTopic       =   "Form4"
   ScaleHeight     =   8295
   ScaleWidth      =   12000
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt_cookie 
      Height          =   1245
      Left            =   1260
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   6930
      Width           =   8265
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成Json"
      Height          =   495
      Left            =   9780
      TabIndex        =   11
      Top             =   7560
      Width           =   915
   End
   Begin VB.CommandButton btn_send 
      Caption         =   "发送"
      Height          =   465
      Left            =   10680
      TabIndex        =   10
      Top             =   2670
      Width           =   945
   End
   Begin VB.TextBox txtExpectNo 
      Height          =   315
      Left            =   9930
      TabIndex        =   9
      Text            =   "660992"
      Top             =   180
      Width           =   1665
   End
   Begin VB.CommandButton btn_Translate 
      Caption         =   "转换"
      Height          =   495
      Left            =   10830
      TabIndex        =   8
      Top             =   7560
      Width           =   915
   End
   Begin VB.TextBox txtResponse 
      Height          =   1755
      Left            =   1260
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   690
      Width           =   10335
   End
   Begin VB.TextBox txtEncrypt 
      Height          =   1455
      Left            =   1290
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   3840
      Width           =   8235
   End
   Begin VB.TextBox txtJson 
      Height          =   1365
      Left            =   1260
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   5430
      Width           =   8265
   End
   Begin VB.TextBox txtInsts 
      Height          =   345
      Left            =   1260
      TabIndex        =   1
      Top             =   150
      Width           =   8535
   End
   Begin VB.Label lbl_ExecStatus 
      Caption         =   "禁止"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lbl_RespStatus 
      Caption         =   "Label6"
      Height          =   345
      Left            =   1260
      TabIndex        =   15
      Top             =   2730
      Width           =   9045
   End
   Begin VB.Label Label5 
      Caption         =   "Cookie"
      Height          =   195
      Left            =   150
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl_status 
      Caption         =   "状态"
      Height          =   315
      Left            =   300
      TabIndex        =   12
      Top             =   6630
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label4 
      Caption         =   "返回"
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "加密"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Json"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   630
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "指令"
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frm_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public InterFace As New FormInterface

Public gobj As SystemClass
Public cookie As String
Public parentForm As Form1
Public LoginForm As frmLogin
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function SendMsg(Expect As String, msg As String, cnt As Integer, Optional ByRef amt As Currency = 0) As Boolean
    On Error Resume Next
    SendMsg = False
    Dim IpAdd As Integer
    IpAdd = gobj.Ip
    Me.txtInsts.Text = msg
    Dim c2i As New CCS2InstrClass
    c2i.cJsOdds = gobj.Odds
    Dim strJson As String
    Dim strEnc As String
    strJson = c2i.InstrToJsonString(Me.txtInsts.Text, 2)
     Me.InterFace.InData = strJson
    Me.InterFace.expectNo = Expect
    Me.InterFace.RowCount = cnt
    Me.InterFace.TranslateInst
    strEnc = Me.InterFace.outdata
    Me.txtEncrypt.Text = strEnc
    Me.txtJson.Text = strJson
    Me.txtExpectNo.Text = Expect
    
    
    
''
''    Dim creq As New Csharp.KCaiRequest
''    Me.txtResponse.Text = creq.GetResult("https://www.kcai773.com/Bet/CqcSubmit", Trim(Me.cookie), strEnc)
''
''
''
''
''
''
''    Exit Sub

    
    Dim httpreq As New XMLHTTP60 'WinHttp.WinHttpRequest    'XMLHTTP60 'MSXML2.ServerXMLHTTP60 '
    Set httpreq = New XMLHTTP60
    'Set httpreq = New XMLHTTP60
    'httpreq.onreadystatechange = httpOnReadyStateChange
    httpreq.Open "Post", "https://www.kcai" & CStr(IpAdd) & ".com/Bet/CqcSubmit"
    httpreq.setRequestHeader "Content-type", "application/x-www-form-urlencoded;charset=utf-8"
    
    'httpreq.setRequestHeader "Accept", "text/plain, text/html"
    'httpreq.Option(4) = 13056
    'httpreq.setProxy
    'httpreq.setRequestHeader "Content-type", "application/json, text/javascript, */*; q=0.01"
    'httpreq.setRequestHeader "Content-type", "application/json; charset=utf-8"
    '            method.setHeader("Accept", "application/json");
    Dim strEnd As String
    
    httpreq.setRequestHeader "Cookie", Me.cookie  '& strEnd
    'httpreq.setRequestHeader "Host", "www.kcai773.com"
'    httpreq.setRequestHeader "Referer", "https://www.kcai773.com/home"
'    httpreq.setRequestHeader "Origin", "https://www.kcai773.com/home"
'    httpreq.setRequestHeader "X-Requested-With", "XMLHttpRequest" '
'    httpreq.setRequestHeader "User-Agent", "Mozilla/6.0" '用户浏览器信息
'    httpreq.setRequestHeader "Access-Control-Allow-Origin", "https://www.kcai773.com"
    'httpreq.setRequestHeader "__RequestVerificationToken", "tAo354u2Zr48wyYpIX2XBRe4Qr0IEiIbbU_NRyZqvVXwF1llRuyKmjB4Q0tHeZ2nTv4n9v2SY2XVw8laBsjg49DcooPyWFp7IC011JHUImLhAJVbDd"
    Me.txtExpectNo.Text = Expect
    Err.Clear
    httpreq.send strEnc
    DoEvents
    stime = Now '获取当前时间
    While httpreq.readyState <> 4
        DoEvents
        ntime = Now '获取循环时间
        If DateDiff("s", stime, ntime) > 3 Then
            'getHtmlStr = "OutTime"
            Exit Function '判断超出3秒即超时退出过程
        End If
    Wend
    If (Err.Number <> 0) Then
        Exit Function
    End If
    Me.lbl_RespStatus.Caption = "已发送指令"
   
    'Me.Timer1.Enabled = True
    'Sleep 1000
    Me.txtResponse.Text = ""
    If 1 = 1 Then
        If httpreq.Status <> 200 Then
            'Me.Timer1.Enabled = False
            'MsgBox httpreq.getAllResponseHeaders()
            Me.txtResponse.Text = Me.txtInsts.Text & " 错误：" & httpreq.getAllResponseHeaders()
            Exit Function
        End If
        SendMsg = True
        Me.lbl_RespStatus.Caption = Me.txtInsts.Text & "=>" & Now()
        Set jsobj = New JsonClass
        Set obj = jsobj.GetJsonVal(httpreq.responseText, "")
        amt = obj.gamePoint
        If obj.Ok = 1 Then
            Me.txtResponse.Text = "结果：" & httpreq.responseText
            Me.txtInsts.Text = ""
        Else
            Me.txtResponse.Text = Me.txtInsts.Text & " 错误：" & obj.Tip & Chr(13) & Chr(10) & httpreq.getAllResponseHeaders()
        End If
        'Me.Timer1.Enabled = False
        Exit Function
    End If
End Function

Private Sub btn_send_Click()
        Dim retamt As Currency
        SendMsg Me.txtExpectNo.Text, Me.txtInsts.Text, 1, retamt
'''''    Dim httpheads As String
'''''    httpheads = httpreq.getAllResponseHeaders()
'''''    cookie = httpreq.getResponseHeader("Set-Cookie")
'''''    If httpreq.Status <> 200 Then
'''''        Me.txtResponse.Text = InterFace.OutData & "&__RequestVerificationToken=" & InterFace.str__RequestVerificationToken & Chr(13) _
'''''        & "[cookie]" & InterFace.str__Cookie & Chr(13) _
'''''        & "[header]:" & httpheads & Chr(13) & httpreq.responseText
'''''
'''''        Exit Sub
'''''    End If
'''''    'MsgBox cookie
'''''    If Len(Trim(cookie)) = 0 Then
'''''        Me.txtResponse.Text = httpheads & Chr(13) & httpreq.responseText
'''''    End If
'''''    Dim strRet As String
'''''    strRet = httpreq.responseText
'''''    If strRet <> """suc""" Then
'''''        'Picture1_Click
'''''        Exit Sub
'''''    End If
End Sub

Function httpOnReadyStateChange()
    Me.lbl_status.Caption = httpreq.readyState
End Function
    
Private Sub btn_Translate_Click()
    Dim cnt As Integer
    Me.InterFace.InData = Me.txtJson.Text
    Me.InterFace.expectNo = gobj.lastExpect
    Me.InterFace.RowCount = cnt
    Me.InterFace.TranslateInst
    Me.txtEncrypt.Text = Me.InterFace.outdata
End Sub

Private Sub Command1_Click()
    Dim c2i As New CCS2InstrClass
    c2i.cJsOdds = gobj.Odds
    Dim cnt As Integer
    
    Me.txtJson.Text = c2i.InstrToJsonString(Me.txtInsts.Text, 2)
End Sub

Private Sub Form_Load()
    If gobj Is Nothing Then Exit Sub
    Me.txtExpectNo.Text = gobj.lastExpect
    Me.txt_cookie.Text = Me.cookie
End Sub

Private Sub Timer1_Timer()
'''        Me.txtResponse.Text = ""
'''        If httpreq.readyState = 4 Then
'''            If httpreq.Status <> 200 Then
'''                Me.Timer1.Enabled = False
'''                Me.txtResponse.Text = Me.txtInsts.Text & " 错误：" & httpreq.statusText & Chr(13) & Chr(10)
'''                Exit Sub
'''            End If
'''            Me.lbl_RespStatus.Caption = Me.txtInsts.Text & "=>" & Now()
'''            Me.txtResponse.Text = httpreq.responseText
'''            Me.txtInsts.Text = ""
'''            Me.Timer1.Enabled = False
'''            Exit Sub
'''        End If
End Sub

Public Function SendInst(strInst As String)
    Me.txtInsts.Text = ""
    If Len(Trim(strInst)) = 0 Then Exit Function
    Me.txtInsts.Text = strInst
    btn_send_Click
End Function



Private Sub txtResponse_DblClick()
    txtResponse.SetFocus
    txtResponse.SelStart = 0
    txtResponse.SelLength = Len(txtResponse.Text)
End Sub
