Attribute VB_Name = "SinusFlag"
Option Explicit

Rem *****************************
Rem * Visual Basic Sinus Effect *
Rem * written by Pro-XeX of     *
Rem * KenamicK Entertainment'01 *
Rem * send me an e-mail,if u're *
Rem * satisfied: bgPro_XeX@yahoo*
Rem *****************************
Rem CM: Now it's a little more realistic maybe!

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const PI = 3.141592
Public Const RAD = PI / 180
Public Const d128 = 720! / 128!         ' make 720 in 128
Public Const AMP = 10                   ' wave amplitude
Public Const WindSpeed = 2#
Public Const BASE_X = 0                ' base image coords
Public Const BASE_Y = 0

Public bRunning As Boolean              ' program running indicator
Public SinTable(128) As Single          ' single for smooth animation
Public DeltaX As Single, DeltaY As Single, DestX As Single, DestY As Single
Public Phase As Long

Public Sub _
CalcSinTable()                            ' calculate the sinus table
Dim i As Integer
 
 For i = 1 To 128
      SinTable(i) = (Sin(d128 * i * RAD) * AMP)
 Next

End Sub

Public Sub _
WaveFlag(ByRef s_pic As PictureBox, ByRef d_pic As PictureBox)

 Dim i As Long, j As Long                 ' local counters
 
 d_pic.Cls                                 ' cleanup old mess
  
 Rem *** It's all here ***
 Phase = Phase + WindSpeed               ' increment the phase
 
 For j = 0 To d_pic.ScaleHeight
  For i = 0 To d_pic.ScaleWidth
  
  DeltaX = SinTable((j + Phase) And 127) / 8
  DestX = i + DeltaX
  
  DeltaY = SinTable((i + Phase) And 127) / 2
  DestY = j + DeltaY
 Rem *********************
                                          ' blit to picture box
  BitBlt d_pic.hDC, BASE_X + DestX, BASE_Y + DestY, 2, 2, s_pic.hDC, i, j, vbSrcCopy
  
  Next
 Next
 
 d_pic.Refresh                             ' refresh new gr. data
 
End Sub



