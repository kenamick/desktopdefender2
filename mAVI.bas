Attribute VB_Name = "mAvi"
Option Explicit

'-------------------------------------
'--> MCI Avi Module
'--> by Peter "Pro-XeX" Petrov
'--> KenamicK Entertainment 1998-2002
'-------------------------------------

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Public Type RECT
'       Top As Long
'       Left As Long
'       Right As Long
'       Bottom As Long
'End Type

Public Enum enumAVIHowPlay
   Notify = 0
   FromStart
   FromEnd
End Enum


Private Const AVI_VIDEO = "avivideo"
Private Const OPEN_AVI_VIDEO = "open avivideo"
Private Const CLOSE_AVI_VIDEO = "close avivideo"
Private Const gNULL = vbNull


Public Sub InitAVI()
       Call mciSend(OPEN_AVI_VIDEO)                 ' initialize libraries
End Sub


Public Sub CloseAllAVI()                            ' closes opened avi and device type
       Call mciSend(CLOSE_AVI_VIDEO)
End Sub


Public Sub OpenAVI(hParent As Long, szFile As String, szAlias As String)
      ' hParent - Parent window that will support the AVI child
      ' szFile  - name of the AVI file
      ' szAlias - alias that mciSendString API will use to recognize the file
      
      Call mciSend("open " & szFile & " alias " & szAlias & _
                   " style child parent " & (hParent))
    
      Call mciSend("realize " & szAlias)            ' the AVI palette needs to be realized
End Sub


Public Sub PlayAVI(szAlias As String, enumHow As enumAVIHowPlay)
     
 If enumHow = Notify Then
    Call mciSend("play " & szAlias & " notify")
 ElseIf enumHow = FromStart Then
    Call mciSend("seek " & szAlias & " to start")   ' rewind
    Call mciSend("play " & szAlias & " notify")     ' play forwards
 ElseIf enumHow = FromEnd Then
    Call mciSend("seek " & szAlias & " to end")     ' play backward
    Call mciSend("play " & szAlias & " reverse notify")
 End If
 
End Sub


Public Sub FFwdAVI(szAlias As String)
    Call mciSend("seek " & szAlias & " to start")   ' rewind
End Sub


Public Sub RewindAVI(szAlias As String)
    Call mciSend("seek " & szAlias & " to end")
End Sub


Public Sub MoveAVIForward(szAlias As String, Optional Step As Integer = 1)
    Call mciSend("step " & szAlias & " by " & Step)
End Sub


Public Sub MoveAVIReverse(szAlias As String, Optional Step As Integer = 1)
    Call mciSend("step " & szAlias & " reverse by " & Step)
End Sub


Public Sub ShowAVI(szAlias As String)               ' display (ONLY) first FRAME to screen
    Call mciSend("window " & szAlias & " state show")
End Sub

Public Sub HideAVI(szAlias As String)               ' hide AVI window
    Call mciSend("window " & szAlias & " state hide")
End Sub

Public Sub PauseAVI(szAlias As String)              ' pause movie
    Call mciSend("pause " & szAlias)
End Sub


Public Sub MoveAVIWindow(szAlias As String, x As Integer, y As Integer, _
                         Optional Width As Integer = 0, Optional Height As Integer = 0)
  Dim rval As Long
  Dim hWin As Long
 ' Dim Width As Integer, Height As Integer
  
  hWin = GetAVIWindow(szAlias)                      ' get the AVI window
  If Width = 0 Then Width = CInt(GetAVIRect(szAlias).Right)    ' get the width
  If Height = 0 Then Height = CInt(GetAVIRect(szAlias).Bottom) ' get the height
  rval = MoveWindow(hWin, CLng(x), CLng(y), CLng(Width), CLng(Height), True)
     
End Sub


Public Function GetAVIRect(szAlias As String) As RECT
  Dim cn As Integer
  Dim Phase As Integer, TempInt As Integer, TempInt2 As Integer
  Dim rval As Long
  Dim szrAVI As String * 128
  Dim rAvi As RECT
                                                    ' get AVI window dimensions in a String Buffer
  rval = mciSendString("where " & szAlias & " source", szrAVI, Len(szrAVI), gNULL)
    
  Do While cn < Len(szrAVI)                         ' algorythm that gets AVI window dimensions from the Buffer
      cn = cn + 1                                   ' string position
      
      If (Asc(Mid$(szrAVI, cn, 1)) < 48 Or Asc(Mid$(szrAVI, cn, 1)) > 57) Then
         Phase = Phase + 1                          ' phase counter ( each phase represents a point of the rectangle )
         If Phase = 1 Then                          ' this is Top
             rAvi.Top = Val(Mid$(szrAVI, 1, cn - 1))
             TempInt = cn + 1                       ' remember position
         ElseIf Phase = 2 Then                      ' this is Left
             rAvi.Left = Val(Mid(szrAVI, TempInt, cn - TempInt))
             TempInt = cn + 1
         ElseIf Phase = 3 Then                      ' this is Right ( Width )
             rAvi.Right = Val(Mid$(szrAVI, TempInt, cn - TempInt))
             TempInt = cn + 1
         ElseIf Phase = 4 Then                      ' this is Bottom ( Height )
             rAvi.Bottom = Val(Mid$(szrAVI, TempInt, cn - TempInt))
             TempInt = cn + 1
         ElseIf Phase > 4 Then                      ' exit, because there's nothing more to get
             Exit Do                                ' actually only Width&Height are important,
         End If                                     ' Top&Left are usually zero
      End If
 Loop
                                         
 GetAVIRect = rAvi
End Function


Public Function GetAVIWindow(szAlias As String) As Long
   Dim rval As Long                                 ' retrive AVI hWnd address
   Dim rs As String * 127
   
   'Call mciSend("status " & szAlias & " window handle")
   rval = mciSendString("status " & szAlias & " window handle", rs, Len(rs), gNULL)
   If rval = 0 Then
     GetAVIWindow = Val(rs)
     Exit Function
   Else                                             ' oops...something went wrong ;-)
     Call mciGetErrorString(rval, rs, Len(rs))      ' find out what...
     MsgBox rs, vbExclamation                       ' tell user...
      GetAVIWindow = -1
   End If
End Function


Public Sub CloseAVI(szAlias As String)
      Call mciSend("close " & szAlias)
End Sub


Public Function mciSend(szCommand As String) As Long
 Dim rval As Long

 rval = mciSendString(szCommand, gNULL, gNULL, gNULL)
 
 mciSend = rval                                     ' get return value
End Function
