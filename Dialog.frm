VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8175
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4965
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8758
      _Version        =   393216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FrmType As String
Public Params As String
Public Gobalobj As SystemClass
Private Sub Form_Load()
    Select Case FrmType
    Case "Chance":
        LoadChances CDate(Params)
    Case Else:
    End Select
End Sub


Sub LoadChances(day As Date)
    Dim cc As New ChanceClass
    Dim dt As DataTable
    Set dt = cc.getSpecDayData(day)
   Dim dgc  As New DataGridClass
   Set dgc.grid = Me.MSFlexGrid1
   dgc.FillGrid dt
End Sub
