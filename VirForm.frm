VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000011&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "策略交易终端"
   ClientHeight    =   10320
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   20115
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "VirForm.frx":0000
   ScaleHeight     =   10320
   ScaleWidth      =   20115
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu menu_1 
      Caption         =   "系统"
   End
   Begin VB.Menu menu_2 
      Caption         =   "账户"
   End
   Begin VB.Menu menu_3 
      Caption         =   "策略管理"
   End
   Begin VB.Menu menu_4 
      Caption         =   "策略资产"
   End
   Begin VB.Menu menu_5 
      Caption         =   "交易"
   End
   Begin VB.Menu menu_6 
      Caption         =   "日终处理"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frm As New Form1
Private Const MIM_BACKGROUND = &H2
Private Const MIM_APPLYTOSUBMENUS = &H80000000
   
Private Type MENUINFO
    cbSize   As Long
    fMask   As Long
    dwStyle   As Long
    cyMax   As Long
    hbrBack   As Long
    dwContextHelpID   As Long
    dwMenuData   As Long
End Type
   
Private Declare Function GetMenuInfo _
                Lib "user32" (ByVal hMenu As Long, _
                              mi As MENUINFO) As Long
Private Declare Function SetMenuInfo _
                Lib "user32" (ByVal hMenu As Long, _
                              mi As MENUINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Sub Form_DblClick()
    If frm Is Nothing Then
        Set frm = New Form1
        frm.Visible = False
    End If
    If frm.Visible = False Then
        frm.Visible = True
        frm.Show
    Else
        frm.Visible = False
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        MsgBox "该程序已经运行，请退出！"
        End
    End If
    Dim MyMenu As MENUINFO
    MyMenu.cbSize = Len(MyMenu)
    MyMenu.fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
    MyMenu.hbrBack = CreateSolidBrush(RGB(205, 201, 201))
    SetMenuInfo GetMenu(Me.hwnd), MyMenu
End Sub
