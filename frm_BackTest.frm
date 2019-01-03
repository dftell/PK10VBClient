VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_BackTest 
   Caption         =   "历史回测"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11775
   LinkTopic       =   "Form4"
   ScaleHeight     =   8070
   ScaleWidth      =   11775
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6375
      Left            =   270
      TabIndex        =   15
      Top             =   1410
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   11245
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox txt_ReadDataRecordCnt 
      Height          =   285
      Left            =   7590
      TabIndex        =   14
      Text            =   "10000"
      Top             =   870
      Width           =   1605
   End
   Begin VB.TextBox txt_CheckStep 
      Height          =   270
      Left            =   4980
      TabIndex        =   12
      Text            =   "5"
      Top             =   870
      Width           =   765
   End
   Begin VB.TextBox txt_MinCheckCount 
      Height          =   285
      Left            =   1650
      TabIndex        =   10
      Text            =   "10"
      Top             =   870
      Width           =   1545
   End
   Begin VB.TextBox txt_ChipCount 
      Height          =   285
      Left            =   10140
      TabIndex        =   7
      Text            =   "6"
      Top             =   300
      Width           =   1395
   End
   Begin VB.ComboBox ddl_type 
      Height          =   300
      ItemData        =   "frm_BackTest.frx":0000
      Left            =   7560
      List            =   "frm_BackTest.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   270
      Width           =   1695
   End
   Begin VB.CommandButton btn_start 
      Caption         =   "开始"
      Height          =   375
      Left            =   10200
      TabIndex        =   4
      Top             =   870
      Width           =   1275
   End
   Begin VB.TextBox txt_ToExpect 
      Height          =   285
      Left            =   4170
      TabIndex        =   3
      Top             =   270
      Width           =   1605
   End
   Begin VB.TextBox txt_FromExpect 
      Height          =   285
      Left            =   1650
      TabIndex        =   1
      Top             =   270
      Width           =   1545
   End
   Begin VB.Label Label7 
      Caption         =   "单词获取数据条数"
      Height          =   285
      Left            =   5910
      TabIndex        =   13
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "数据检查步长"
      Height          =   255
      Left            =   3420
      TabIndex        =   11
      Top             =   870
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "最小检测次数"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "注长度"
      Height          =   225
      Left            =   9420
      TabIndex        =   8
      Top             =   330
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "回测类型"
      Height          =   285
      Left            =   5970
      TabIndex        =   6
      Top             =   300
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   285
      Left            =   3450
      TabIndex        =   2
      Top             =   300
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "开始期号"
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   735
   End
End
Attribute VB_Name = "frm_BackTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gobj As SystemClass

Private Sub btn_start_Click()
    'On Error Resume Next
    Dim btc As New BackTestClass
    Set btc.gobj = New SystemClass
    Set btc.gobj.SysParams = gobj.SysParams
    btc.CheckType = Me.ddl_type.ListIndex
    btc.ChipCount = CInt(Me.txt_ChipCount.Text)
    btc.FromExpect = Me.txt_FromExpect.Text
    btc.MinCheckTimes = CInt(Me.txt_MinCheckCount.Text)
    btc.ReadDataCount = CLng(Me.txt_ReadDataRecordCnt.Text)
    btc.StepLen = CInt(Me.txt_CheckStep.Text)
    Dim dicRepeatCodes As New Dictionary
    btc.ExecTest dicRepeatCodes
    If dicRepeatCodes Is Nothing Then
    Else
        With Me.TreeView1
            .Nodes.Clear
            Dim cnt As Integer
            cnt = 0
            .Nodes.Add , , "Root", "重复号码"
            For Each key In dicRepeatCodes.Keys
                Dim OpenCodeNode As node
                Dim strArr() As String
                cnt = cnt + 1
                Dim strKey As String
                strKey = "K" & CStr(cnt)
                Set OpenCodeNode = .Nodes.Add("Root", tvwChild, strKey, key)
                OpenCodeNode.Sorted = True
                OpenCodeNode.Expanded = True
                strArr = dicRepeatCodes(key)
                Dim i As Integer
                For i = 1 To UBound(strArr)
                    Dim subNode As node
                    Set subNode = .Nodes.Add(strKey, tvwChild, strKey & "_" & strArr(i), strArr(i))
                    'Set subNode.Parent = OpenCodeNode
                Next
            Next
        End With
    End If
End Sub

Private Sub Form_Load()
    Me.txt_FromExpect.Text = gobj.ValidOldestHistoryExpect
End Sub
