VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Login"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4575
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4575
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   465
      Left            =   1740
      TabIndex        =   4
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2310
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "abcdef123"
      Top             =   1170
      Width           =   1395
   End
   Begin VB.TextBox txt_user 
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Text            =   "dftell"
      Top             =   810
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
      Height          =   225
      Left            =   1050
      TabIndex        =   2
      Top             =   1230
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "用户"
      Height          =   255
      Left            =   1050
      TabIndex        =   0
      Top             =   810
      Width           =   885
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
