VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShuttle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem *** Player_Shuttle Ship Class ***
Rem *** Purpose: Maintain player shuttle

Enum CONST_SHUTTLEMOVEMENT
  SM_NONE = 0
  SM_MOVELEFT = 1
  SM_MOVERIGHT = 2
  SM_MOVEUP = 4
  SM_MOVEDOWN = 8
End Enum

Private Const SHUTTLE_ACCELERATION      As Single = 0.05
Private Const SHUTTLE_MAXACCELERATION   As Single = 4.2
Private Const SHUTTLE_MAXVACCELERATION  As Single = SHUTTLE_MAXACCELERATION / 3#

Private m_x           As Single                        ' position
Private m_y           As Single
Private m_vx          As Single                        ' velocity
Private m_vy          As Single
Private m_ax          As Single                        ' acceleration
Private m_ay          As Single
Private m_HitPoints   As Long                          ' shuttle hitpoints

'///////////////////////////////////////////////////////////////////////////////////
'//// Name: ()
'//// Desc:
'///////////////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
 Call Me.Create(50, 50)
End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: Create()
'//// Desc: Setup the ship
'///////////////////////////////////////////////////////////////////////////////////
Public Sub _
Create(xpos As Integer, ypos As Integer)

 ' set starting position
 m_x = rScreen.Right / 2  'xpos
 m_y = rScreen.Bottom / 2 'ypos
 m_HitPoints = 100
 wx = 0
 wy = 0
End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: Update()
'//// Desc: Update ship movement and render it
'///////////////////////////////////////////////////////////////////////////////////
Public Sub _
Update()
 
 Static lTime     As Long
 Static bMoveShip As Boolean                            ' advance ship or world position
 Static shpDir    As enumDirection
 Static lFrame    As Long
  
 Dim dx           As Long
 Dim dy           As Long
 Dim lCurTime     As Long
 Dim shipmv       As CONST_SHUTTLEMOVEMENT
 
 shipmv = SM_NONE
 
 lCurTime = GetTicks()
 If (lTime < lCurTime) Then
  lTime = 100 + lCurTime
 
  Call mDirectInput.DICheckKeys
  If (DIKeyState(DIK_W) = KS_KEYDOWN) Then shipmv = shipmv Or SM_MOVEUP
  If (DIKeyState(DIK_S) = KS_KEYDOWN) Then shipmv = shipmv Or SM_MOVEDOWN
  If (DIKeyState(DIK_A) = KS_KEYDOWN) Then shipmv = shipmv Or SM_MOVELEFT
  If (DIKeyState(DIK_D) = KS_KEYDOWN) Then shipmv = shipmv Or SM_MOVERIGHT
 End If
 

 ' accelerate shuttle
 If (shipmv And SM_MOVELEFT) Then
  m_vx = -SHUTTLE_MAXACCELERATION
  lFrame = 0
 ElseIf (shipmv And SM_MOVERIGHT) Then
  m_vx = SHUTTLE_MAXACCELERATION
  lFrame = 6
 End If
 
 If (shipmv And SM_MOVEUP) Then
  m_vy = -SHUTTLE_MAXVACCELERATION
 ElseIf (shipmv And SM_MOVEDOWN) Then
  m_vy = SHUTTLE_MAXVACCELERATION
 End If
 
 ' advance vertical position
 m_y = m_y + m_vy
 
 ' check and advance either world or ship position
 If (Not bMoveShip) Then
  
  wx = wx + m_vx
  
  If (wx < 0) Then
   shpDir = SO_LEFT
   bMoveShip = True
  ElseIf (wx > arMS_Offsets(2)) Then
   shpDir = SO_RIGHT
   bMoveShip = True
  End If
  
 Else
 
  m_x = m_x + m_vx
  
  ' move shuttle left
  If (shpDir = SO_LEFT) Then
  
   If (m_x < 0) Then
    m_x = 0
    m_vx = 0
   ElseIf (m_x > rScreen.Right / 2) Then
    m_x = rScreen.Right / 2
    bMoveShip = False
   End If
  
  ' move shuttle right
  ElseIf (shpDir = SO_RIGHT) Then
   
   ' bound horizontal position
   If (m_x > MAX_CX - 50) Then
    m_x = MAX_CX - 50
    m_vx = 0
   ElseIf (m_x < rScreen.Right / 2) Then
    m_x = rScreen.Right / 2
    bMoveShip = False
   End If
  
  End If
  
   End If
 
 
 ' kill velocity
 If (m_vx < 0) Then
  m_vx = m_vx + SHUTTLE_ACCELERATION
 ElseIf (m_vx > 0) Then
  m_vx = m_vx - SHUTTLE_ACCELERATION
 End If
 If (m_vy < 0) Then
  m_vy = m_vy + SHUTTLE_ACCELERATION
 ElseIf (m_vy > 0) Then
  m_vy = m_vy - SHUTTLE_ACCELERATION
 End If
 
 ' bound coordinates
 If (m_y < 0) Then
  m_y = 0
  m_vy = 0
 ElseIf (m_y > VISIBLE_AREA_CY) Then
  m_y = VISIBLE_AREA_CY
  m_vy = 0
 End If
 
 ' get blitting coordinates
 dx = m_x
 dy = m_y
 
 ' render ship
 Call mGFX.BltFastGFX_HBM(dx, dy, g_Objects.Ship1(lFrame))
 
End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: GetX()()
'//// Desc: get horizontal position
'///////////////////////////////////////////////////////////////////////////////////
Public Property Get _
GetX() As Long

 GetX = wx + m_x

End Property


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: GetY()()
'//// Desc: get vertical position
'///////////////////////////////////////////////////////////////////////////////////
Public Property Get _
GetY() As Long

 GetY = wy + m_y

End Property



