VERSION 5.00
Begin VB.UserControl CSubclass 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   2445
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cool SSTab  2.0"
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   930
   End
End
Attribute VB_Name = "CSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Download by http://www.NewXing.com
'SUBCLASSING THE SSTab Control  By Mario Alberto Flores Gonzalez
'version 2.0
'February 1, 2005
'Feel free to use this source code as you wish in your projects

'                        sistec_de_juarez@hotmail.com

'Revision 2.0 Changed Subclass Method, safe whitout vb crash!!

'SStab code By Mario Flores.
'Subclass patch by Paul Caton

'==================================================================================================
'Subclasser declarations


Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData


'=====================================================
'THE POINTAPI STRUCTURE
Private Type POINTAPI
    X As Long                       ' The POINTAPI structure defines the x- and y-coordinates of a point.
    Y As Long
End Type
'=====================================================


'=====================================================
'THE RECT STRUCTURE
Private Type RECT
    Left   As Long
    Top    As Long                  ' The RECT structure defines the coordinates of the
    Right  As Long                  ' upper-left and lower-right corners of a rectangle.
    Bottom As Long
End Type
'=====================================================

' *********************************************************************************
'  API Declarations...
' *********************************************************************************

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long



' *********************************************************************************
'  Const Declarations...
' *********************************************************************************

Private Const WM_PAINT          As Long = &HF
Private Const WM_DESTROY        As Long = &H2
Private Const WM_TIMER          As Long = &H113
Private Const WM_ENABLE         As Long = &HA
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_KEYUP          As Long = &H101
Private Const WM_LBUTTONDOWN    As Long = &H201        '//--- Subclass Messages
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_MOUSEMOVE      As Long = &H200


'------------------------------------
Private DestDC      As Long
Private MaskDC      As Long
Private MemDC       As Long
Private OrigDC      As Long
Private MaskPic     As Long              'Temporary DC
Private MemPic      As Long
Private TempPic     As Long
Private OrigPic     As Long
Private TempDC      As Long
'-------------------------------------

Private origBrush As Long
Private TempBrush As Long
Private origColor As Long 'BackColor

'---------------------------
Private gColor1   As Long 'Gradient Color Start
Private gColor2   As Long 'Gradient Color End
Private gDir      As Long 'Gradient Dir


'=====================================================
'MATH CONSTANTS
Private Const TwoPower16 = 2 ^ 16
'=====================================================

Private iHw   As Integer
Private iLW   As Integer
Private bDown As Boolean
Private bOver As Boolean




'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data

    Dim m_ItemRect As RECT
    Dim m_Width    As Long
    Dim m_Height   As Long


  If Ambient.UserMode = False Then Exit Sub
  
  
  Select Case uMsg
   
    
    Case WM_PAINT, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_KEYDOWN, WM_KEYUP, WM_TIMER, WM_MOUSEMOVE

        
        If uMsg <> WM_PAINT Then
            Select Case ThisWindowClassName(lng_hWnd)
               
               Case "ThunderCommandButton", "ThunderRT6CommandButton", "Button"
           
                    If uMsg = WM_TIMER Then
                        If InsideArea(lng_hWnd) = False Then
                            KillTimer lng_hWnd, 1
                            bOver = False
                            RedrawWindow lng_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                        Else
                            bOver = True
                        End If
                        Exit Sub
                    End If
               
                    If uMsg = WM_MOUSEMOVE Then
                        If InsideArea(lng_hWnd) = True Then
                            If bOver = False Then
                                bOver = True
                                RedrawWindow lng_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                                SetTimer lng_hWnd, 1, 1, 0
                            End If
                        Else
                            If bOver Then
                                bOver = False
                            End If
                        End If
                        Exit Sub
                    End If
                              
                    If uMsg = WM_KEYDOWN Then
                        If Not bDown Then
                            bDown = True
                            LongInt2Int wParam, iHw, iLW
                            Select Case (iLW)
                                   Case vbKeySpace
                                        '-Pass
                                   Case Else
                                        Exit Sub
                            End Select
                        End If
                    End If
                   
                    If uMsg = WM_KEYDOWN Then
                        If bDown Then
                            bDown = False
                            LongInt2Int wParam, iHw, iLW
                            Select Case (iLW)
                                   Case vbKeySpace
                                        '-Pass
                                   Case Else
                                        Exit Sub
                            End Select
                        End If
                    End If
                
               Case Else
                    Exit Sub
               End Select
        
        End If
         
        '---------------------------------------------------

        '//--- Get the SSTab's dimensions
        DestDC = GetDC(lng_hWnd)

        GetWindowRect lng_hWnd, m_ItemRect
                m_Width = m_ItemRect.Right - m_ItemRect.Left
                m_Height = m_ItemRect.Bottom - m_ItemRect.Top
       '---------------------------------------------------


      '---------------------------------------------------

        '//--- Select The Parameters (SSTab New Style)
        Select Case GetStyleParams(lng_hWnd)

         Case 0
              GetSolidColor lng_hWnd
         Case 1
              GetPictureParams lng_hWnd
         Case 2
              GetGradientColor1 lng_hWnd
              GetGradientColor2 lng_hWnd
              GetGradientDir lng_hWnd
         Case Else
               Debug.Print "Invalid Style"
         End Select
       '---------------------------------------------------

       '----------------------------------------------------------------------------
       '//--- To Work With a Cleaner and Less Flicker Screen Create the Temporary DC
       CreateNewDCWorkArea m_Width, m_Height
       '----------------------------------------------------------------------------

       '---------------------------------------------------------------------------
        Call SelectBitmap '//-- Selected Image
       '---------------------------------------------------------------------------
        
       
       '---------------------------------------------------------------------------
        CallWindowProc sc_aSubData(zIdx(lng_hWnd)).nAddrOrig, lng_hWnd, WM_PAINT, OrigDC, lParam  ' PAINT SSTab in TEMPORARY DC
       '---------------------------------------------------------------------------
        '---------------------------------------------------------------------------
        Call CreateBackMask(m_Width, m_Height)  '//-- A Mask For RasterOperations
        '---------------------------------------------------------------------------



        'The PatBlt function paints the given rectangle using the brush that is currently
        'selected into the specified device context.
        'The brush color and the surface color(s) are combined by using the given raster operation.

        '-----------------------------------------------------------------------------------------------------
        origBrush = SelectObject(TempDC, TempBrush)

        If GetStyleParams(lng_hWnd) = 2 Then
            DrawGradient gColor1, gColor2, 0, 0, m_Width, m_Height, TempDC, IIf(gDir = 1, False, True)
        Else
            PatBlt TempDC, 0, 0, m_Width, m_Height, vbPatCopy
        End If

        SelectObject TempDC, origBrush
        '------------------------------------------------------------------------------------------------------

        Call DOBitBlt(m_Width, m_Height) '//--- Do RasterOperations
        Call CleanDCs                    '//--- Free Memory <--Prevent Leaks


        '-----------------------------------------------------------------------
        SetBkColor DestDC, origColor
        ReleaseDC lng_hWnd, DestDC '//-- Free The DC FROM GetDC API ..AND RETURN THE COLOR BACK TO NORMAL
        ValidateRect lng_hWnd, 0
        '-----------------------------------------------------------------------
   
    Case WM_DESTROY
        KillTimer lng_hWnd, 0
        DeleteObject TempBrush
       
   End Select


End Sub




Public Sub SubClassMe( _
           ByVal Window As Long, _
           ByVal Style As Long, _
           Optional Picture As StdPicture, _
           Optional SolidColor As Long = 0, _
           Optional GradientDir As Long = 0, _
           Optional GradientColor1 As Long = 0, _
           Optional GradientColor2 As Long = 0)

               If GetSubClassedTag(Window) = 1 Then
                    Subclass_Stop (Window)
                    Debug.Print "HandleWindow : ", Window, " UnSubClassed"
               End If
               
               Call Subclass_Start(Window)
               Call Subclass_AddMsg(Window, WM_PAINT, MSG_BEFORE)
               Call Subclass_AddMsg(Window, WM_TIMER, MSG_BEFORE)
               Call Subclass_AddMsg(Window, WM_DESTROY, MSG_BEFORE)
               Call Subclass_AddMsg(Window, WM_ENABLE, MSG_BEFORE)
               Call Subclass_AddMsg(Window, WM_KEYDOWN, MSG_BEFORE_AND_AFTER)
               Call Subclass_AddMsg(Window, WM_KEYUP, MSG_BEFORE_AND_AFTER)
               Call Subclass_AddMsg(Window, WM_LBUTTONDOWN, MSG_BEFORE_AND_AFTER)
               Call Subclass_AddMsg(Window, WM_LBUTTONUP, MSG_BEFORE_AND_AFTER)
               Call Subclass_AddMsg(Window, WM_MOUSEMOVE, MSG_BEFORE)
               
               SetStyle Window, Style
               SetGradientDir Window, GradientDir
               If Not Picture Is Nothing Then SetPicture Window, Picture.Width, Picture.Height, Picture
               If SolidColor <> 0 Then SetSolidColor Window, SolidColor
               If GradientColor1 <> 0 Then SetGradientColor1 Window, GradientColor1
               If GradientColor2 <> 0 Then SetGradientColor2 Window, GradientColor2
               
               RedrawWindow Window, ByVal 0&, ByVal 0&, &H1
               
               SetSubClassedTag Window, 1
               Debug.Print "HandleWindow : ", Window, " SubClassed"
               
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
  On Error GoTo Catch
  If Ambient.UserMode Then
    'Stop all subclassing - either that or call Subclass_Stop for each individual hWnd that's being subclassed
    Call Subclass_StopAll
    Exit Sub
  End If
 
Catch:

Subclass_StopAll

End Sub


'Stop all subclassing
Public Sub Subclass_StopAll()
On Error Resume Next
  
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        SetSubClassedTag .hwnd, 0
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    
    I = I - 1                                                                           'Next element
  Loop
End Sub

'Refresh All the Subclassed Windows.
Public Sub Refresh_All()
   On Error Resume Next
   If Ambient.UserMode Then
       Dim I As Long
     
       I = UBound(sc_aSubData())
       Do While I >= 0
       With sc_aSubData(I)
          If .hwnd <> 0 Then
            RedrawWindow .hwnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
          End If
       End With
       I = I - 1
       Loop
   End If
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error Resume Next
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    KillTimer lng_hWnd, 1
    Call SetWindowLong(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'On Error Resume Next
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
 
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLong(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
    
  End With
End Function

'=======================================================================================================================
' SELECT THE CURRENT IMAGE
'=======================================================================================================================

Private Sub SelectBitmap()
Dim cHandle As Long

       cHandle = SelectObject(MaskDC, MaskPic)
       DeleteObject cHandle
       cHandle = SelectObject(MemDC, MemPic)
       DeleteObject cHandle
       cHandle = SelectObject(TempDC, TempPic)
       DeleteObject cHandle
       cHandle = SelectObject(OrigDC, OrigPic)
       DeleteObject cHandle
       
End Sub

'=======================================================================================================================
' CREATE A MASK COLOR BACKGROUND
'=======================================================================================================================

Private Sub CreateBackMask(ByVal m_Width As Long, ByVal m_Height As Long)
        
        origColor = SetBkColor(DestDC, GetSysColor(15))
        SetBkColor OrigDC, GetSysColor(15)
        BitBlt MaskDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcCopy
       
End Sub


'=======================================================================================================================
' CREATE THE NEW TEMP WORK AREA
'=======================================================================================================================

Private Sub CreateNewDCWorkArea(ByVal m_Width As Long, ByVal m_Height As Long)
        
        MaskDC = CreateCompatibleDC(DestDC)
        MaskPic = CreateBitmap(m_Width, m_Height, 1, 1, ByVal 0&)
        MemDC = CreateCompatibleDC(DestDC)
        MemPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        TempDC = CreateCompatibleDC(DestDC)
        TempPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        OrigDC = CreateCompatibleDC(DestDC)
        OrigPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)

End Sub


'=======================================================================================================================
' BITBLT  RasterOperations
'=======================================================================================================================

Private Sub DOBitBlt(ByVal m_Width As Long, ByVal m_Height As Long)
        
        BitBlt MemDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbSrcCopy
        BitBlt MemDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcPaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbMergePaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, MemDC, 0, 0, vbSrcAnd
        BitBlt DestDC, 0, 0, m_Width, m_Height, TempDC, 0, 0, vbSrcCopy

End Sub

'=======================================================================================================================
' CLEAN UP MEMORY
'=======================================================================================================================

Private Sub CleanDCs()
        
        DeleteDC TempDC
        DeleteObject TempPic
        DeleteDC MaskDC
        DeleteObject MaskPic
        DeleteDC MemDC
        DeleteObject MemPic
        DeleteDC OrigDC
        DeleteObject OrigPic
        DeleteObject TempBrush

End Sub

'=======================================================================================================================
' MY GRADIENT FUNCTION
'=======================================================================================================================
Public Sub DrawGradient(lEndColor As Long, lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hdc As Long, Optional bH As Boolean)
    ''Draw a Vertical Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    lEndColor = GetLngColor(lEndColor)
    lStartcolor = GetLngColor(lStartcolor)
    
    'lh = Height
    'lw = Width
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    sR = (sR - eR) / IIf(bH, X2, Y2)
    sG = (sG - eG) / IIf(bH, X2, Y2)
    sB = (sB - eB) / IIf(bH, X2, Y2)
    
        
    For ni = 0 To IIf(bH, X2, Y2)
        
        If bH Then
            DrawLine X + ni, Y, X + ni, Y2, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Else
            DrawLine X, Y + ni, X2, Y + ni, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        End If
        
    Next ni
End Sub

'======================================================================
'DRAWS A LINE WITH A DEFINED COLOR
Private Sub DrawLine( _
           ByVal X As Long, _
           ByVal Y As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal cHdc As Long, _
           ByVal Color As Long)

    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim Outline As Long
    Dim POS     As POINTAPI

    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    
        MoveToEx cHdc, X, Y, POS
        LineTo cHdc, Width, Height
          
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1

End Sub
'======================================================================

'=======================================================================================================================
' FADE COLOR GRADIENT FUNCTION
'=======================================================================================================================

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long

Dim R As Long
Dim G As Long
Dim B As Long

      R = (Color And &HFF) + Value
      G = ((Color \ &H100) Mod &H100) + Value
      B = ((Color \ &H10000) Mod &H100)
      B = B + ((B * Value) \ &HC0)
      
    If Value > 0 Then
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
    ElseIf Value < 0 Then
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
    End If

    ShiftColor = R + 256& * G + 65536 * B

End Function

Private Function GetLngColor(Color As Long) As Long

    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

Private Sub SetSubClassedTag(ByVal hwnd As Long, State As Long)
           SetProp hwnd, "ImSubClassed", State
End Sub

Private Function GetSubClassedTag(ByVal hwnd As Long) As Long
           GetSubClassedTag = GetProp(hwnd, "ImSubClassed")
End Function

Private Sub SetStyle(ByVal hwnd As Long, ByRef Style As Long)
           SetProp hwnd, "MyStyle", Style
End Sub

Private Function GetStyleParams(ByVal hwnd As Long) As Long
           GetStyleParams = GetProp(hwnd, "MyStyle")
End Function

Private Sub SetGradientDir(ByVal hwnd As Long, ByVal Style As Long)
           SetProp hwnd, "MyGradientDir", Style
End Sub

Private Sub GetGradientDir(ByVal hwnd As Long)
           gDir = GetProp(hwnd, "MyGradientDir")
End Sub

Private Sub SetSolidColor(ByVal hwnd As Long, ByVal Color As Long)
           SetProp hwnd, "MySolidColor", GetLngColor(Color)
End Sub

Private Sub SetGradientColor1(ByVal hwnd As Long, ByVal Color As Long)
           SetProp hwnd, "MyGradientColor1", GetLngColor(Color)
End Sub

Private Sub SetGradientColor2(ByVal hwnd As Long, ByVal Color As Long)
           SetProp hwnd, "MyGradientColor2", GetLngColor(Color)
End Sub

Private Sub GetSolidColor(ByVal hwnd As Long)
     TempBrush = CreateSolidBrush(GetProp(hwnd, "MySolidColor"))
End Sub

Private Sub GetGradientColor1(ByVal hwnd As Long)
     gColor1 = GetProp(hwnd, "MyGradientColor1")
End Sub

Private Sub GetGradientColor2(ByVal hwnd As Long)
     gColor2 = GetProp(hwnd, "MyGradientColor2")
End Sub

Private Sub SetPicture(ByVal hwnd As Long, ByVal Width As Long, ByVal Height As Long, ByRef cPicture As StdPicture)

           SetProp hwnd, "MyPicture", cPicture.Handle
           SetProp hwnd, "MyPictureWidth", Width
           SetProp hwnd, "MyPictureHeight", Height

End Sub

Private Sub GetPictureParams(ByVal hwnd As Long)

    TempBrush = CreatePatternBrush(GetProp(hwnd, "MyPicture"))

End Sub

'======================================================================
'GET'S THE NAME OF A CLASS
Private Function ThisWindowClassName(ByVal hwnd As Long) As String
Dim retVal As Long, lpClassName As String

    lpClassName = Space(255)
    retVal = GetClassName(hwnd, lpClassName, 255)
    ThisWindowClassName = VBA.Left$(lpClassName, retVal)

End Function
'======================================================================


Private Function LongInt2Int(ByVal lLongInt As Long, ByRef iHiWord As Integer, ByRef iLowWord As Integer) As Boolean

Dim tmpHW As Integer, tmpLW As Integer
    
    RtlMoveMemory tmpLW, lLongInt, Len(tmpLW)
    tmpHW = (lLongInt / TwoPower16)
    iHiWord = tmpHW
    iLowWord = tmpLW

End Function

Private Function InsideArea(cHandle As Long) As Boolean
Dim POS As POINTAPI
        
        GetCursorPos POS

        If (WindowFromPoint(POS.X, POS.Y) <> cHandle) Then
            InsideArea = False
            Else 'NOT (WINDOWFROMPOINT(POS.X,...
            InsideArea = True
        End If

End Function

