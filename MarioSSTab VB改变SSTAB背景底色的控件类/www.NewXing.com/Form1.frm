VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SubClassing SSTab "
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SSTab13 
      Caption         =   "Buttons"
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   8040
      Width           =   1815
   End
   Begin MSComCtl2.MonthView SSTab11 
      Height          =   2370
      Left            =   6120
      TabIndex        =   20
      Top             =   6360
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19791873
      CurrentDate     =   38376
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OFF"
      Height          =   495
      Left            =   2640
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton SSTab10 
      Caption         =   "Supports"
      Height          =   615
      Left            =   4080
      TabIndex        =   18
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton SSTab9 
      Caption         =   "Now"
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   1815
   End
   Begin MSComctlLib.Slider SSTab7 
      Height          =   675
      Left            =   240
      TabIndex        =   15
      Top             =   6360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1191
      _Version        =   393216
   End
   Begin Project1.CSubclass CSubclass1 
      Left            =   9000
      Top             =   120
      _extentx        =   2355
      _extenty        =   1296
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apply SubClassing"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   8040
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   4260
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   16777215
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mario Alberto Flores Gonzalez"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   2100
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   16777215
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":05A6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":05C2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":05DE
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Image Image4 
         Height          =   720
         Left            =   120
         Picture         =   "Form1.frx":05FA
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image Image5 
         Height          =   720
         Left            =   -74760
         Picture         =   "Form1.frx":14C4
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   -74760
         Picture         =   "Form1.frx":238E
         Top             =   1560
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab4 
      Height          =   2415
      Left            =   4920
      TabIndex        =   8
      Top             =   3360
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   4260
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   65535
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":3258
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":3274
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mario Alberto Flores Gonzalez"
         Height          =   195
         Left            =   -74760
         TabIndex        =   9
         Top             =   1920
         Width           =   2100
      End
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   2655
      Left            =   4920
      TabIndex        =   10
      Top             =   360
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":3290
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":32AC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":32C8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin TabDlg.SSTab SSTab5 
      Height          =   2655
      Left            =   9960
      TabIndex        =   13
      Top             =   360
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   4683
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":32E4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":3300
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":331C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin TabDlg.SSTab SSTab6 
      Height          =   2655
      Left            =   9960
      TabIndex        =   14
      Top             =   3120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   4683
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":3338
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":3354
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":3370
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSComctlLib.Slider SSTab8 
      Height          =   675
      Left            =   240
      TabIndex        =   16
      Top             =   7200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1191
      _Version        =   393216
   End
   Begin MSComCtl2.MonthView SSTab12 
      Height          =   2370
      Left            =   9120
      TabIndex        =   21
      Top             =   6360
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19791873
      CurrentDate     =   38376
   End
   Begin VB.Image Image6 
      Height          =   3000
      Left            =   11280
      Picture         =   "Form1.frx":338C
      Tag             =   "YUST FOR DEMO "
      Top             =   8280
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ssStylePropertPage"
      Height          =   195
      Left            =   7560
      TabIndex        =   12
      Top             =   3120
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ssStyleTabbedDialog"
      Height          =   195
      Left            =   7560
      TabIndex        =   11
      Top             =   5880
      Width           =   1500
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "SSTab Subclassing By Mario Flores G"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2685
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ssStyleTabbedDialog"
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   5880
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ssStylePropertPage"
      Height          =   195
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   1380
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   11760
      Picture         =   "Form1.frx":7803
      Tag             =   "YUST FOR DEMO "
      Top             =   7440
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   11400
      Picture         =   "Form1.frx":A82E
      Tag             =   "YUST FOR DEMO "
      Top             =   6600
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CSubclass1.Subclass_StopAll
Command3.Enabled = True
Command1.Visible = False
Me.Hide
Me.Show
End Sub

'SUBCLASSING THE SSTab Control  By Mario Alberto Flores Gonzalez
'version 2.1
'January 10, 2005
'Feel free to use this source code as you wish in your projects

'Download by http://www.NewXing.com
Private Sub Command3_Click()
     
       
    CSubclass1.SubClassMe SSTab1.hwnd, 1, Image1 '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab2.hwnd, 1, Image2 '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab3.hwnd, 2, , , 0, &HC0&, &H80FFFF   '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab4.hwnd, 2, , , 1, &HC56A31, vbButtonFace '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab5.hwnd, 0, , &HFFC0FF    '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab6.hwnd, 0, , &H80FF80    '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab7.hwnd, 1, Image6 '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab8.hwnd, 2, , , 0, vbBlue, vbRed '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab9.hwnd, 1, Image6 '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab10.hwnd, 2, , , 0, vbBlue, vbRed '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab11.hwnd, 1, Image1 '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab12.hwnd, 1, Image2 '//--- Begin SubClassing
    CSubclass1.SubClassMe SSTab13.hwnd, 2, , , 0, vbBlue, vbButtonFace '//--- Begin SubClassing
    Command3.Enabled = False
    Command1.Visible = True

End Sub


