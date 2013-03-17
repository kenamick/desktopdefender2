Attribute VB_Name = "mUtil"
'-----------------------------------------
'--- Utils Module
'-----------------------------------------
Rem *** Purpose: Some math, calculations, system info retrieving procs., _
                 Log File and etc.

Enum CONST_PRIORITIES
  P_HIGH = THREAD_PRIORITY_ABOVE_NORMAL
  P_NORMAL = THREAD_PRIORITY_NORMAL
End Enum


' logfile constants
Public Const LOG_FILENAME = "dd2log.txt"             ' logfile name
Public Const LOG_FILEHANDLE = 1                      ' logfile_handle
Public Const LOG_DASH As String = "-------------------------------------------------" & vbCrLf
' math constantss
Public Const PI = 3.1415926
Public Const PI360 = 6.2831853
Public Const DEGTORAD = PI / 180
Public Const RADTODEG = 180 / PI
Public Const MAX_INT = 32767
Public Const MIN_INT = -32768


' ////////////////////////////////////////////////////////////////
' //// Set program priority
' ////////////////////////////////////////////////////////////////
Public Function _
SetProgramPriority(lprior As CONST_PRIORITIES) As Boolean

 Dim curproc As Long

 curproc = GetCurrentThread()
 If (SetThreadPriority(curproc, lprior)) Then
  SetPriority = True
  AppendToLog ("Setting priority " & lprior & " successful.")
 Else
  SetPriority = False
  AppendToLog ("Setting priority " & lprior & " failed.")
 End If
 
 'curproc = GetCurrentProcess()
 'If (SetPriorityClass(curproc, lprior)) Then _
 ' SetPriority = True Else _
 ' SetPriority = False

End Function


' ////////////////////////////////////////////////////////////////
' //// Open up the Log File
' ////////////////////////////////////////////////////////////////
Public Sub _
OpenLog(lpPath As String)
  
 On Local Error GoTo OPENLOG_ERROR:
  
 Open lpPath & LOG_FILENAME For Append Access Write Lock Read As #LOG_FILEHANDLE

Exit Sub

OPENLOG_ERROR:
 Call ErrorMsg("Could not open Log file!")
End Sub


' ////////////////////////////////////////////////////////////////
' //// Stream data to log file
' ////////////////////////////////////////////////////////////////
Public Sub _
AppendToLog(lpStr As String)
   
 On Local Error Resume Next
  
  Print #LOG_FILEHANDLE, chGetTime & ": " & lpStr '   Format(Time$, "hh:mm:ss")

End Sub


' ////////////////////////////////////////////////////////////////
' //// Close the log file
' ////////////////////////////////////////////////////////////////
Public Sub _
CloseLog()
 
 On Local Error Resume Next
 
 AppendToLog (LOG_DASH)
 Print #LOG_FILEHANDLE, "Game Closed at: " & chGetTime
 Print #LOG_FILEHANDLE, LOG_DASH
 Close #LOG_FILEHANDLE
End Sub


' ////////////////////////////////////////////////////////////////
' //// Get velocity frictions (I used this before I get
' //// the GetAngle function)
' ////////////////////////////////////////////////////////////////
Public Sub CalcVelocityBound(ByVal x As Single, ByVal y As Single, _
                             ByVal dx As Single, ByVal dy As Single, _
                             xVel_Bound As Single, yVel_Bound As Single)
 ' Desc: procedure that returns velocity fixing values, so
 '       an object will move at specific vector
    
 Dim fVar1 As Single, fVar2 As Single
   
 xVel_Bound = 1                                      ' reset vars to 1 so division will give the same value
 yVel_Bound = 1
 fVar1 = Abs(x - dx)
 fVar2 = Abs(y - dy)
 If fVar1 <= 0 Then fVar1 = 1
 If fVar2 <= 0 Then fVar2 = 1
 'Call GetDist2P(x, dx, fVar1)                        ' get x distance
 'Call GetDist2P(y, dy, fVar2)                        ' get y distacne
 If fVar1 > fVar2 Then                               ' if xdist>ydist then
    yVel_Bound = fVar1 / fVar2                       ' assign xd/yd divison to y-velocity-boundary
 ElseIf fVar1 < fVar2 Then
    xVel_Bound = fVar2 / fVar1                       ' assign yd/xd divison to x-velocity-boundary
 End If
 
End Sub


' ////////////////////////////////////////////////////////////////
' //// Get The Angle Between Two Points
' ////////////////////////////////////////////////////////////////
Public Function _
GetAngle(ByVal x1, ByVal y1, ByVal x2, ByVal y2) As Single

 Dim val1 As Single, val2 As Single
 val1 = (x2 - x1)
 val2 = (y2 - y1)
 
 If (val1 > 0) Then
  GetAngle = Atn(val2 / val1)
 ElseIf (val1 < 0) Then
  GetAngle = Atn(val2 / val1) + PI
 Else
  GetAngle = 2 * Atn(Sgn(val2))
 End If

End Function


' ////////////////////////////////////////////////////////////////
' //// Checks if 2 rectangles interpolate
' ////////////////////////////////////////////////////////////////
Public Function _
Collide(rObject1 As RECT, rObject2 As RECT) As Boolean
 Dim rRect As RECT

 If (IntersectRect(rRect, rObject1, rObject2)) Then
  Collide = True
 Else
  Collide = False
 End If
 
End Function


Public Function _
InRange(ByVal SrcVar As Integer, ByVal MinVar As Integer, ByVal MaxVar As Integer) As Boolean

 If (SrcVar >= MinVar And SrcVar <= MaxVar) Then
  InRange = True
 Else
  InRange = False
 End If

End Function

' ////////////////////////////////////////////////////////////////
' //// checks if var 1 is bigger than var2 ( for both INT and
' //// FLOAT data types )
' ////////////////////////////////////////////////////////////////
Public Function _
max(ByVal Var1, ByVal Var2) As Boolean
 If Var1 >= Var2 Then
  max = True
 Else
  max = False
 End If
End Function


' ////////////////////////////////////////////////////////////////
' //// Get 1D distance ;-)
' ////////////////////////////////////////////////////////////////
Public Sub _
GetDist2P(ByVal Var1 As Single, ByVal Var2 As Single, ByVal Var3 As Single)
  ' Desc: Get the distance between 2 points
  'Var3 = Sqr((Var1 - Var2) ^ 2)
  Var3 = Abs(Var1 - Var2)
End Sub


' ////////////////////////////////////////////////////////////////
' //// Get the Distance between two points (Phytagorous)
' ////////////////////////////////////////////////////////////////
Public Function _
nGetDist2D(sx As Integer, sy As Integer, _
                           dx As Integer, dy As Integer) As Integer
 nGetDist2D = Sqr(((sx - dx) ^ 2) + ((sy - dy) ^ 2))
End Function


Public Function _
fGetDist2D(sx As Single, sy As Single, _
                           dx As Single, dy As Single) As Single
 fGetDist2D = Sqr(((sx - dx) ^ 2) + ((sy - dy) ^ 2))
End Function


' ////////////////////////////////////////////////////////////////
' //// get random INT value
' ////////////////////////////////////////////////////////////////
Public Function _
nGetRnd(nMin As Integer, nMax As Integer) As Integer
 nGetRnd = ((nMax - nMin) * Rnd) + nMin
End Function


' ////////////////////////////////////////////////////////////////
' //// get random FLOAT value
' ////////////////////////////////////////////////////////////////
Public Function _
fGetRnd(nMin As Single, nMax As Single) As Single
 fGetRnd = ((nMax - nMin) * Rnd) + nMin
End Function


'////////////////////////////////////////////////////////////////
'//// Convert unsigned to signed value
'////////////////////////////////////////////////////////////////
Public Function _
CSigned(lNum As Long) As Integer

 If (lNum < 32768) Then
  CSigned = CInt(lNum)
 Else
  CSigned = CInt(lNum - 65535)
 End If

End Function


'////////////////////////////////////////////////////////////////
'//// Convert signed to unsigned value
'////////////////////////////////////////////////////////////////
Public Function _
CUnSigned(nNum As Integer) As Long

 If (nNum >= 0) Then
  CUnSigned = nNum
 Else
  CUnSigned = (nNum + 65535)
 End If

End Function


'////////////////////////////////////////////////////////////////
'//// Return Windows tickcount
'////////////////////////////////////////////////////////////////
Public Function _
GetTicks() As Long
 
 GetTicks = GetTickCount()
End Function

'////////////////////////////////////////////////////////////////
'//// Return current local time (h:m:s:ms)
'////////////////////////////////////////////////////////////////
Public Function _
chGetTime() As String
 
 Dim lpSysTime As SYSTEMTIME
 Dim chs As String
 
 ' get local time settings
 GetLocalTime lpSysTime
 chs = Format$(lpSysTime.wSecond) ', "##")
 chGetTime = lpSysTime.wHour & ":" & lpSysTime.wMinute & ":" & _
             chs & "." & lpSysTime.wMilliseconds
 
End Function


'////////////////////////////////////////////////////////////////
'//// Calculates the Frame Rate per Second ( in 2 ways )
'////////////////////////////////////////////////////////////////
Public Function _
CalcFrameRate(npFPS As Integer) As Integer
 Static lTime As Long
 Static lFrameTime As Long
 Static nFrameCount As Long
 Static Frames As Long
 Static szAveFPS As String
  
 ' Method 1
 ' nFrameCount = nFrameCount + 1
 ' lTime = GetTicks - lFrameTime                  ' see if substraction is bigger then 1 sec.
 
 ' If lTime > 1000 Then                                 ' 1 sec. elapsed
 '   Frames = (nFrameCount * 1000) / lTime              ' calc FPS
 '   lFrameTime = GetTicks
 '   nFrameCount = 0                                    ' zero raw frame counter
 '   CalcFrameRate = Frames
 '   'szAveFPS = CStr(Frames)                            ' copy average FPS to a string
 ' End If
  
  'Method 2
  If GetTicks - lTime >= 1000 Then
        lTime = GetTicks
        'szAveFPS = CStr(Frames)
        npFPS = CInt(Frames)
        Frames = 0
  End If
  Frames = Frames + 1
 
 ' blit FPS info to backbuffer
 'lpBack.SetForeColor RGB(45, 255, 35)
 'lpBack.DrawText 10, 70, szAveFPS & " FPS", False

End Function


'////////////////////////////////////////////////////////////////
'//// Gets key value
'//// LONG   lKeyRoot - registry root   (HKEY_LOCAL_MACHINE)
'//// LPCSTR strKeyName - main key name (SOFTWARE\DD2\PLAYERS)
'//// LPCSTR strKeyName - key name      (NAME)
'//// LPCSTR strKeyVal - where the data goes to
'////////////////////////////////////////////////////////////////
Public Function _
RegLoadKey(lKeyRoot As Long, _
           strKeyName As String, _
           strSubKeyName As String, _
           ByRef strKeyVal As String) As Boolean

 Dim cn      As Long
 Dim ret     As Long
 Dim hKey    As Long
 Dim hDisp   As Long
 Dim lpAttr  As SECURITY_ATTRIBUTES
 Dim lplType As Long
 Dim strTemp As String
 Dim lSize   As Long
 
 ' open key reg
 ret = RegOpenKeyEx(lKeyRoot, strKeyName, 0, KEY_ALL_ACCESS, hKey)
 
 If (ret <> 0) Then GoTo REGERROR
 
 strTemp = Space$(1024)
 lSize = 1024
 
 ' retrieve key-value
 ret = RegQueryValueEx(hKey, strSubKeyName, 0, _
                       lplType, strTemp, lSize)
                       
 If (ret <> 0) Then GoTo REGERROR
 
 If (Asc(Mid(strTemp, lSize, 1)) = 0) Then
  strTemp = Left(strTemp, lSize - 1)
 Else
  strTemp = Left(strTemp, lSize)
 End If
 
 ' determine keytype and do convertions
 Select Case lplType
   
   Case REG_SZ
    strKeyVal = strTemp
   
   Case REG_DWORD
     For cn = Len(strTemp) To 1 Step -1                     ' Convert Each Bit
      strKeyVal = strKeyVal + Hex(Asc(Mid(strTemp, cn, 1)))  ' Build Value Char. By Char.
      'strKeyVal = strKeyVal + Chr$(Asc(Mid(strTemp, cn, 1))) ' convert to chars
     Next
     strKeyVal = Format$("&h" + strKeyVal)                  ' Convert Double Word To String
 
 End Select
 
  
 ' unlock
 ret = RegCloseKey(hKey)
 RegLoadKey = True

Exit Function

REGERROR:
 ret = RegCloseKey(hKey)
 RegLoadKey = False
End Function


'////////////////////////////////////////////////////////////////
'//// Writes key value
'//// LONG   lKeyRoot - registry root (HKEY_LOCAL_MACHINE)
'//// LPCSTR strKeyName - main key name (SOFTWARE\DD2\PLAYERS)
'//// LPCSTR strSubKeyName - key name
'//// LPCSTR strKeyVal - data to copy
'////////////////////////////////////////////////////////////////
Public Function _
RegSetKey(lKeyRoot As Long, strKeyName As String, _
          strSubKeyName As String, strSubKeyValue As String) As Boolean
                  
 Dim ret    As Long
 Dim hKey   As Long
 Dim hDisp  As Long
 Dim lpAttr As SECURITY_ATTRIBUTES
                  
 lpAttr.nLength = 50
 lpAttr.lpSecurityDescriptor = 0
 lpAttr.bInheritHandle = True
 
 ' create/open reg. key
 ret = RegCreateKeyEx(lKeyRoot, strKeyName, 0, REG_SZ, _
                      REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                      hKey, hDisp)
 ' incase of an error
 If (ret <> 0) Then GoTo REGERROR
                    
 ' create/modify key value
 If (strSubKeyName = "") Then strSubKeyName = " "  ' need space to work
 
 ret = RegSetValueEx(hKey, strSubKeyName, 0, _
                     REG_SZ, strSubKeyValue, Len(strSubKeyValue))
 
 If (ret <> 0) Then GoTo REGERROR
 
 ' close
 ret = RegCloseKey(hKey)
 RegSetKey = True
 
Exit Function

REGERROR:
 ret = RegCloseKey(hKey)
 RegSetKey = False
End Function


