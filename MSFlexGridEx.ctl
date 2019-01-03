VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl MSFlexGridEditor 
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   ScaleHeight     =   3585
   ScaleWidth      =   5370
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6165
      _Version        =   393216
   End
End
Attribute VB_Name = "MSFlexGridEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const ASC_ENTER = 13 '回车

Dim gRow As Integer

Dim gCol As Integer

Public Grid As MSFlexGrid


Private Sub Grid1_KeyPress(KeyAscii As Integer)
    ' Move the text box to the current grid cell:

    Text1.Top = Grid1.CellTop + Grid1.Top
    
    Text1.Left = Grid1.CellLeft + Grid1.Left
    
    ' Save the position of the grids Row and Col for later:
    
    gRow = Grid1.row
    
    gCol = Grid1.col
    
    ' Make text box same size as current grid cell:
    
    Text1.Width = Grid1.CellWidth - 2 * Screen.TwipsPerPixelX
    
    Text1.Height = Grid1.CellHeight - 2 * Screen.TwipsPerPixelY
    
    ' Transfer the grid cell text:
    
    Text1.Text = Grid1.Text
    
    ' Show the text box:
    
    Text1.Visible = True
    
    Text1.ZOrder 0 ' 把 Text1 放到最前面！
    
    Text1.SetFocus
    
    ' Redirect this KeyPress event to the text box:
    
    If KeyAscii <> ASC_ENTER Then
    
    SendKeys Chr$(KeyAscii)
    
    End If
End Sub





Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = ASC_ENTER Then
    
    Grid1.SetFocus ' Set focus back to grid, see Text_LostFocus.
    
    KeyAscii = 0 ' Ignore this KeyPress.

End If
End Sub

Private Sub Text1_LostFocus()
    
    Dim tmpRow As Integer
    
    Dim tmpCol As Integer
    
    ' Save current settings of Grid Row and col. This is needed only if
    
    ' the focus is set somewhere else in the Grid.
    
    tmpRow = Grid1.row
    
    tmpCol = Grid1.col
    
    ' Set Row and Col back to what they were before Text1_LostFocus:
    
    Grid1.row = gRow
    
    Grid1.col = gCol
    
    Grid1.Text = Text1.Text ' Transfer text back to grid.
    
    Text1.SelStart = 0 ' Return caret to beginning.
    
    Text1.Visible = False ' Disable text box.
    
    ' Return row and Col contents:
    
    Grid1.row = tmpRow
    Grid1.col = tmpCol
End Sub

Private Sub UserControl_Initialize()
    Set Grid = Grid1
End Sub
