VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Rem *** Mouse Class ***
Rem *** Comments: VERY important in RTS games ;)

Enum enumButtonStates                               ' button state
  MS_Unpressed = 0
  MS_Up
  MS_Down
End Enum

Enum enumMouseMode
  MI_DI = 0
  MI_API
End Enum

Dim CP_OLD As POINTAPI                              ' old_Mouse_Pos ( to restore when we exit game )
Dim CP As POINTAPI
Dim x  As Long, y  As Long                          ' mouse_x&y
Dim ox As Long, oy As Long                          ' old_x&y
Dim cx As Integer, cy As Integer                    ' cursor size
Dim Sens As Integer
Dim bMouseMoved As Boolean
Dim LeftButton  As enumButtonStates                 ' leftbutton_states
Dim RightButton As enumButtonStates                 ' rightbutton_states
Dim InputMode   As enumMouseMode

' //////////////////////////////////////////////////////////
' //// On initialization
' //////////////////////////////////////////////////////////
Private Sub _
Class_Initialize()
 
 Call SaveCoords
End Sub

' //////////////////////////////////////////////////////////
' //// On destruction
' //////////////////////////////////////////////////////////
Private Sub _
Class_Terminate()
 '...
 Call SetOldCoords
End Sub

' //////////////////////////////////////////////////////////
' //// Get mouse position and button info
' //////////////////////////////////////////////////////////
Public Sub GetActions()                             ' get user interaction
 On Local Error Resume Next
 Dim lpDIPt As POINTAPI
 Dim bL As Boolean, bR As Boolean
  
  ox = x                                            ' save old position
  oy = y
  
  If InputMode = MI_DI Then                         ' get info via DirectInput
     
     Call DIGetMouse(lpDIPt, bL, bR)                 ' get DI mouse offsets and buttons status

     x = x + lpDIPt.x * Sens                        ' get position
     y = y + lpDIPt.y * Sens

  If bL Then                                        ' do checks for the left_button
     LeftButton = MS_Down                           ' is it's pressed then set down flag
  Else
     If LeftButton = MS_Down Then                   ' ok, now if it was down and now it's up set UP flag
        LeftButton = MS_Up
     Else                                           ' it was NOT down, so set normal unpressed flag
        LeftButton = MS_Unpressed
     End If
  End If
  
  If bR Then                                        ' the same goes for the right button
     RightButton = MS_Down
  Else
     If RightButton = MS_Down Then
        RightButton = MS_Up
     Else
        RightButton = MS_Unpressed
     End If
  End If
  
  
  ElseIf InputMode = MI_API Then    '*****          ' get info via API
     Call GetCursorPos(CP)
     Call ScreenToClient(frmMain.hwnd, CP)          ' we need to convert global mouse coords. to local window coords
     x = CInt(CP.x)
     y = CInt(CP.y)
          
  'If GetAsyncKeyState(WM_LBUTTONDOWN) <> 0 Then
  If GetAsyncKeyState(VK_LBUTTON) <> 0 Then
     LeftButton = MS_Down
  Else
     If LeftButton = MS_Down Then
        LeftButton = MS_Up
     Else
        LeftButton = MS_Unpressed
     End If
  End If
  
  'If GetAsyncKeyState(WM_RBUTTONDOWN) <> 0 Then
  If GetAsyncKeyState(VK_RBUTTON) <> 0 Then
     RightButton = MS_Down
  Else
     If RightButton = MS_Down Then
        RightButton = MS_Up
     Else
        RightButton = MS_Unpressed
     End If
  End If
  
  End If ' /**********/
  
  If (x <> ox Or y <> oy) Then                      ' set mouse_moved flag
     bMouseMoved = True
  Else
     bMouseMoved = False
  End If
  
  If (GetX < 0) Then x = 0
  If (GetX > MAX_CX) Then x = MAX_CX
  If (GetY < 0) Then y = 0
  If (GetY > MAX_CY) Then y = MAX_CY
  
End Sub

Public Sub Acquire()                                ' acquire mouse
 'On Local Error Resume Next
 If InputMode = MI_DI Then Call mDirectInput.DIAcquire(DI_MOUSE)
 If Err.Number <> 0 Then Call mDirectInput.DIGetErrorDesc(Err.Number)    'CDXErr.HandleError(Err.Number)
End Sub

Public Sub UnAcquire()
 If InputMode = MI_DI Then Call mDirectInput.DIUnAcquire(DI_MOUSE)
End Sub

Public Sub SaveCoords()                             ' save old windows coordinates
 'Call ShowCursor(False)                             ' hide cursor
 Call GetCursorPos(CP_OLD)
 Call ScreenToClient(frmMain.hwnd, CP_OLD)          ' convert to screen coords.
 x = CP_OLD.x                                       ' get coordinates
 y = CP_OLD.y
 Sens = 1                                           ' default sensitivity
End Sub

Public Sub SetOldCoords()                           ' when leaving set user's prev. mouse pos.
' Call ShowCursor(True)                              ' show back cursor
 Call ClientToScreen(frmMain.hwnd, CP_OLD)
 Call SetCursorPos(CP_OLD.x, CP_OLD.y)
End Sub

' //////////////////////////////////////////////////////////
' //// Properties setting and retrieving info
' //////////////////////////////////////////////////////////
Public Property Get MouseMoved() As Boolean
  If bMouseMoved = True Then MouseMoved = True
End Property

Public Property Let SetMouseInput(IMode As enumMouseMode) ' set input mode
  On Local Error Resume Next
  InputMode = IMode
  
  If InputMode = MI_DI Then
     Call Acquire
     If Err.Number <> 0 Then Call mDirectInput.DIGetErrorDesc(Err.Number)
  Else                                               ' release DI_Mouse if we're goin' to use API
     Call UnAcquire
  End If
End Property

Public Property Get GetRight() As enumButtonStates  ' get_right_button_state
  GetRight = RightButton
End Property

Public Property Get GetLeft() As enumButtonStates   ' get_left_button state
  GetLeft = LeftButton
End Property

Public Property Let SetSens(nSens As Integer)       ' set mouse sensitivity
  Sens = nSens
End Property

Public Property Get GetSens() As Integer            ' get mouse sensitiviy
  GetSens = Sens
End Property

Public Property Get GetX() As Integer               ' get mouse position
  GetX = CInt(x)
End Property

Public Property Get GetY() As Integer
  GetY = CInt(y)
End Property

Public Property Get GetCenterX() As Integer             ' get cursor-middle position
  GetCenterX = CInt(x) - cx \ 2
End Property

Public Property Get GetCenterY() As Integer
  GetCenterY = CInt(y) - cy \ 2
End Property

Public Property Let SetX(nVal As Integer)
  x = nVal - cx \ 2
End Property

Public Property Let SetY(nVal As Integer)
  y = nVal - cx \ 2
End Property

Public Property Let SetCursorWidth(nVal As Integer)
  cx = nVal
End Property

Public Property Let SetCursorHeight(nVal As Integer)
  cy = nVal
End Property

