Attribute VB_Name = "mDirectInput"
Option Explicit
'-------------------------------------
'--> DirectInput Engine
'--> by Peter "Pro-XeX" Petrov
'--> KenamicK Entertainment 1998-2002
'-------------------------------------

'Keycode constants ( pasted from directly from Lucky, so I'm lazy ;)
Global Const DIK_ESCAPE = 1
Global Const DIK_1 = 2
Global Const DIK_2 = 3
Global Const DIK_3 = 4
Global Const DIK_4 = 5
Global Const DIK_5 = 6
Global Const DIK_6 = 7
Global Const DIK_7 = 8
Global Const DIK_8 = 9
Global Const DIK_9 = 10
Global Const DIK_0 = 11
Global Const DIK_MINUS = 12
Global Const DIK_EQUALS = 13
Global Const DIK_BACKSPACE = 14
Global Const DIK_TAB = 15
Global Const DIK_Q = 16
Global Const DIK_W = 17
Global Const DIK_E = 18
Global Const DIK_R = 19
Global Const DIK_T = 20
Global Const DIK_Y = 21
Global Const DIK_U = 22
Global Const DIK_I = 23
Global Const DIK_O = 24
Global Const DIK_P = 25
Global Const DIK_LBRACKET = 26
Global Const DIK_RBRACKET = 27
Global Const DIK_RETURN = 28
Global Const DIK_LCONTROL = 29
Global Const DIK_A = 30
Global Const DIK_S = 31
Global Const DIK_D = 32
Global Const DIK_F = 33
Global Const DIK_G = 34
Global Const DIK_H = 35
Global Const DIK_J = 36
Global Const DIK_K = 37
Global Const DIK_L = 38
Global Const DIK_SEMICOLON = 39
Global Const DIK_APOSTROPHE = 40
Global Const DIK_GRAVE = 41  ' = `
Global Const DIK_LSHIFT = 42
Global Const DIK_BACKSLASH = 43
Global Const DIK_Z = 44
Global Const DIK_X = 45
Global Const DIK_C = 46
Global Const DIK_V = 47
Global Const DIK_B = 48
Global Const DIK_N = 49
Global Const DIK_M = 50
Global Const DIK_COMMA = 51
Global Const DIK_PERIOD = 52
Global Const DIK_SLASH = 53
Global Const DIK_RSHIFT = 54
Global Const DIK_MULTIPLY = 55
Global Const DIK_LALT = 56
Global Const DIK_SPACE = 57
Global Const DIK_CAPSLOCK = 58
Global Const DIK_F1 = 59
Global Const DIK_F2 = 60
Global Const DIK_F3 = 61
Global Const DIK_F4 = 62
Global Const DIK_F5 = 63
Global Const DIK_F6 = 64
Global Const DIK_F7 = 65
Global Const DIK_F8 = 66
Global Const DIK_F9 = 67
Global Const DIK_F10 = 68
Global Const DIK_NUMLOCK = 69
Global Const DIK_SCROLL = 70
Global Const DIK_NUMPAD7 = 71
Global Const DIK_NUMPAD8 = 72
Global Const DIK_NUMPAD9 = 73
Global Const DIK_SUBTRACT = 74
Global Const DIK_NUMPAD4 = 75
Global Const DIK_NUMPAD5 = 76
Global Const DIK_NUMPAD6 = 77
Global Const DIK_ADD = 78
Global Const DIK_NUMPAD1 = 79
Global Const DIK_NUMPAD2 = 80
Global Const DIK_NUMPAD3 = 81
Global Const DIK_NUMPAD0 = 82
Global Const DIK_DECIMAL = 83
Global Const DIK_F11 = 87
Global Const DIK_F12 = 88
Global Const DIK_NUMPADENTER = 156
Global Const DIK_RCONTROL = 157
Global Const DIK_DIVIDE = 181
Global Const DIK_RALT = 184
Global Const DIK_HOME = 199
Global Const DIK_UP = 200
Global Const DIK_PAGEUP = 201
Global Const DIK_LEFT = 203
Global Const DIK_RIGHT = 205
Global Const DIK_END = 207
Global Const DIK_DOWN = 208
Global Const DIK_PAGEDOWN = 209
Global Const DIK_INSERT = 210
Global Const DIK_DELETE = 211

Enum enumDIDevices
 DI_MOUSE = 1
 DI_KEYBOARD = 2
End Enum

Public Enum cnstKeyState
 KS_UNPRESSED = 0
 KS_KEYUP
 KS_KEYDOWN
End Enum

Private Const DI_MOUSE_GUID = "GUID_SYSMOUSE"
Private Const DI_KEYBOARD_GUID = "GUID_SYSKEYBOARD"
Private Const DIC = &H80

Private lpDI         As DirectInput             ' DirectInput object
Private lpDIMouse    As DirectInputDevice       ' Mouse Device
Private lpDIKeyboard As DirectInputDevice       ' Keyboard Device

' additional declares
Private bInitKeyb   As Boolean
Private bInitMouse  As Boolean
Private arKeys(211) As cnstKeyState


' //////////////////////////////////////////////////////////
' //// Initialize DirectInput object and devices
' //////////////////////////////////////////////////////////
Public Sub _
DIInit(hwnd As Long, lFlags As enumDIDevices)
On Local Error GoTo DIError
   
   Call AppendToLog(LOG_DASH)
   Call AppendToLog("Initializing DirectInput...")
   Set lpDI = lpDX.DirectInputCreate()
   
   If (lFlags And DI_KEYBOARD) Then
    Call AppendToLog("Setting up Keyboard...")
    Set lpDIKeyboard = lpDI.CreateDevice(DI_KEYBOARD_GUID)
    Call lpDIKeyboard.SetCommonDataFormat(DIFORMAT_KEYBOARD)
    Call lpDIKeyboard.SetCooperativeLevel(hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND)
    Call lpDIKeyboard.Acquire
    bInitKeyb = True
    Call AppendToLog("Keyboard is acquired.")
   End If
   
   If (lFlags And DI_MOUSE) Then
    Call AppendToLog("Setting up Mouse...")
    Set lpDIMouse = lpDI.CreateDevice(DI_MOUSE_GUID)
    Call lpDIMouse.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call lpDIMouse.SetCooperativeLevel(hwnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND)
    'Call lpDIMouse.Acquire
    'cMouse.Acquire
    bInitMouse = True
    Call AppendToLog("Mouse is acquired.")
   End If

Exit Sub

DIError: 'AppendToLog ("Error: Error initializing DirectInput!")
         'MakeError ("Error initializing DirectInput!")
         'Call CDXErr.HandleError(Err.Number)
         Call mDirectInput.DIGetErrorDesc(Err.Number)
End Sub


' //////////////////////////////////////////////////////////
' //// Convert DirectInput key constant to the ASCII eq.
' //// Remark: Supported only A-Z and 0-9 keys
' //// Beware of Japanese keyboards ;)
' //// BYTE bytKeyCode - key constant to convert
' //// BOOL bShiftDown -
' //// Returns: key ASCII code
' //////////////////////////////////////////////////////////
Public Function _
DIKeyToASCII(bytKeyCode As Byte, bUcase As Boolean) As Byte
 
 ' 65 = A, 97 = a
 ' 48 = 0, 57 = 9

 Dim bytShift As Byte
 
 ' shift letters & some chars.
 If (bUcase) Then bytShift = 0 _
  Else bytShift = 32
 
  ' check writings
 Select Case bytKeyCode
   
   ' various
   Case DIK_GRAVE: DIKeyToASCII = IIf(bytShift <> 0, 96, 126)
   Case DIK_BACKSLASH: DIKeyToASCII = IIf(bytShift <> 0, 92, 124)
   Case DIK_SLASH: DIKeyToASCII = IIf(bytShift <> 0, 47, 63)
   Case DIK_MINUS:  DIKeyToASCII = IIf(bytShift <> 0, 45, 95)
   Case DIK_EQUALS: DIKeyToASCII = IIf(bytShift <> 0, 46, 61)
   Case DIK_SUBTRACT: DIKeyToASCII = 45
   Case DIK_ADD: DIKeyToASCII = 46
   
   
   ' numeric
   Case DIK_1: DIKeyToASCII = 33 + (bytShift / 2)
   Case DIK_2: DIKeyToASCII = 50 '34 + (bytShift / 2)
   Case DIK_3: DIKeyToASCII = 35 + (bytShift / 2)
   Case DIK_4: DIKeyToASCII = 36 + (bytShift / 2)
   Case DIK_5: DIKeyToASCII = 37 + (bytShift / 23)
   Case DIK_6: DIKeyToASCII = 54 '38 + (bytShift / 2)
   Case DIK_7: DIKeyToASCII = 55 '39 + (bytShift / 2)
   Case DIK_8: DIKeyToASCII = 56 '40 + (bytShift / 2)
   Case DIK_9: DIKeyToASCII = 41 + (bytShift / 2) ' - CBool(bytShift / 2))
   Case DIK_0: DIKeyToASCII = 48 '32 + (bytShift / 2)
   
   ' charset
   Case DIK_A
     DIKeyToASCII = 65 + bytShift
   Case DIK_B
     DIKeyToASCII = 66 + bytShift
   Case DIK_C
     DIKeyToASCII = 67 + bytShift
   Case DIK_D
     DIKeyToASCII = 68 + bytShift
   Case DIK_E
     DIKeyToASCII = 69 + bytShift
   Case DIK_F
     DIKeyToASCII = 70 + bytShift
   Case DIK_G
     DIKeyToASCII = 71 + bytShift
   Case DIK_H
     DIKeyToASCII = 72 + bytShift
   Case DIK_I
     DIKeyToASCII = 73 + bytShift
   Case DIK_J
     DIKeyToASCII = 74 + bytShift
   Case DIK_K
     DIKeyToASCII = 75 + bytShift
   Case DIK_L
     DIKeyToASCII = 76 + bytShift
   Case DIK_M
     DIKeyToASCII = 77 + bytShift
   Case DIK_N
     DIKeyToASCII = 78 + bytShift
   Case DIK_O
     DIKeyToASCII = 79 + bytShift
   Case DIK_P
     DIKeyToASCII = 80 + bytShift
   Case DIK_Q
     DIKeyToASCII = 81 + bytShift
   Case DIK_R
     DIKeyToASCII = 82 + bytShift
   Case DIK_S
     DIKeyToASCII = 83 + bytShift
   Case DIK_T
     DIKeyToASCII = 84 + bytShift
   Case DIK_U
     DIKeyToASCII = 85 + bytShift
   Case DIK_V
     DIKeyToASCII = 86 + bytShift
   Case DIK_W
     DIKeyToASCII = 87 + bytShift
   Case DIK_X
     DIKeyToASCII = 88 + bytShift
   Case DIK_Y
     DIKeyToASCII = 89 + bytShift
   Case DIK_Z
     DIKeyToASCII = 90 + bytShift
   
   ' not found
   Case Else
     DIKeyToASCII = 0
 
 End Select
 
End Function

' //////////////////////////////////////////////////////////
' //// Exclusive proc. for special mentainance
' //// BYTE bytKeyCode - key constant to check
' //// Retruns: keystate constant
' //////////////////////////////////////////////////////////
Public Function _
DICheckKeyEx(bytKeyCode As Byte) As cnstKeyState

  Dim dk As DIKEYBOARDSTATE
  
  Call lpDIKeyboard.GetDeviceStateKeyboard(dk)
  If (Err.Number <> 0) Then Call mDirectInput.DIGetErrorDesc(Err.Number)  'Call CDXErr.HandleError(Err.Number)
  
   If (arKeys(bytKeyCode) = KS_KEYDOWN) Then
    arKeys(bytKeyCode) = KS_KEYUP
   Else
    arKeys(bytKeyCode) = KS_UNPRESSED
   End If
   
   If (dk.Key(bytKeyCode) <> 0) Then
    arKeys(bytKeyCode) = KS_KEYDOWN
   Else
    '...
   End If
  
  
  DICheckKeyEx = arKeys(bytKeyCode)
'  If (dk.Key(bytKeyCode) And &H80) Then DICheckKeyEx = True _
'  Else DICheckKeyEx = False
 
End Function

' //////////////////////////////////////////////////////////
' //// Check if a specific key is pressed
' //////////////////////////////////////////////////////////
Public Function _
DIKeyState(bytKeyCode As Byte) As cnstKeyState
 
 DIKeyState = arKeys(bytKeyCode)

End Function

' //////////////////////////////////////////////////////////
' //// Put keys states into an array
' //////////////////////////////////////////////////////////
Public Sub _
DICheckKeys()

  Dim dk As DIKEYBOARDSTATE
  Dim cn As Integer
  
  'Call lpDIKeyboard.GetDeviceState(Len(dk), dk)
  Call lpDIKeyboard.GetDeviceStateKeyboard(dk)
  If (Err.Number <> 0) Then Call mDirectInput.DIGetErrorDesc(Err.Number)
   
  For cn = 1 To 211
   
   If (arKeys(cn) = KS_KEYDOWN) Then
    arKeys(cn) = KS_KEYUP
   Else
    arKeys(cn) = KS_UNPRESSED
   End If
   
   If (dk.Key(cn) <> 0) Then
    arKeys(cn) = KS_KEYDOWN
   Else
    '...
   End If
   
  Next

End Sub

' //////////////////////////////////////////////////////////
' //// Gets mouse coords. and button status
' //////////////////////////////////////////////////////////
Public Function _
DIGetMouse(lpPt As POINTAPI, bLeft As Boolean, bRight As Boolean) As Boolean
 On Local Error GoTo DIError
 Dim dm As DIMOUSESTATE

  lpDIMouse.GetDeviceStateMouse dm
 'Call lpDIMouse.GetDeviceState(Len(dm), dm)
  
  lpPt.x = dm.x
  lpPt.y = dm.y
  bLeft = CBool(dm.buttons(0))
  bRight = CBool(dm.buttons(1))
  DIGetMouse = True

Exit Function

DIError: ' error handler
 lpPt.x = 0
 lpPt.y = 0
 bLeft = False
 bRight = False
 DIGetMouse = False
 Call mDirectInput.DIGetErrorDesc(Err.Number)
End Function
     
' //////////////////////////////////////////////////////////
' //// Acquire device
' //////////////////////////////////////////////////////////
Public Function _
DIAcquire(enumDIWhat As enumDIDevices)
 
 If ((enumDIWhat And DI_KEYBOARD) And bInitKeyb) Then
  lpDIKeyboard.Acquire
 End If
 
 If ((enumDIWhat And DI_MOUSE) And bInitMouse) Then
  lpDIMouse.Acquire
 End If
 
End Function

' //////////////////////////////////////////////////////////
' //// Unacquire device
' //////////////////////////////////////////////////////////
Public Function _
DIUnAcquire(enumDIWhat As enumDIDevices)
 
 If ((enumDIWhat And DI_KEYBOARD) And bInitKeyb) Then
  lpDIKeyboard.UnAcquire
 End If
 
 If ((enumDIWhat And DI_MOUSE) And bInitMouse) Then
  lpDIKeyboard.UnAcquire
 End If
 
End Function


' //////////////////////////////////////////////////////////
' //// DirectSound error descriptions
' //// lError - error number
' //////////////////////////////////////////////////////////
Public Function _
DIGetErrorDesc(lError As Long) As String

 Dim strMsg As String

 Select Case lError
   
   Case DIERR_INPUTLOST
     strMsg = "DI_ERROR: DIERR_INPUTLOST !"
   
   Case DIERR_NOTACQUIRED
     strMsg = "DI_ERROR: DIERR_NOTACQUIRED !"
   
   Case DIERR_INVALIDPARAM
     strMsg = "DI_ERROR: DIERR_INVALIDPARAM !"
   
   Case E_PENDING
     strMsg = "DI_ERROR: E_PENDING !" & Chr$(13) & _
     "Mouse data is not available yet!"
   
   Case DIERR_OTHERAPPHASPRIO
     strMsg = "DI_ERROR: DIERR_OTHERAPPHASPRIO !" & Chr$(13) & _
     "Can't acquire device in background!"
   
   Case DIERR_OUTOFMEMORY
     strMsg = "DI_ERROR: DIERR_OUTOFMEMORY !" & Chr$(13) & _
     "DirectInput couldn't allocate sufficient memory to complete the call. !"
   
   Case DIERR_NOINTERFACE
     strMsg = "DI_ERROR: DIERR_NOINTERFACE !" & Chr$(13) & _
     "The specified interface is not supported by the object !"

   Case Else
     strMsg = "Error description not found!"
 
 End Select

 ' return error message
 DIGetErrorDesc = strMsg
 AppendToLog (strMsg)
 
End Function

' //////////////////////////////////////////////////////////
' //// Release DirectInput objects
' //////////////////////////////////////////////////////////
Public Sub _
DIRelease()
 
 'AppendToLog (LOG_DASH)
 AppendToLog ("Closing DirectInput")
 Call DIUnAcquire(DI_KEYBOARD Or DI_MOUSE)
 Set lpDIKeyboard = Nothing
 Set lpDIMouse = Nothing
 Set lpDI = Nothing

End Sub

