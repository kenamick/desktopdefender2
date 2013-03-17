VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem *** Ship Class ***
Rem *** Purpose: Create and update ships, ship's weapons and etc.

Enum enumShipType
  ' z = 2
  ST_INTERCEPTOR1 = 0               ' interceptor: SMALL
  ST_LRSNEAKY                       ' missile-crusier: MEDIUM
  ST_SEVENTHFOX                     ' lasergun-crusier: MEDIUM
  ST_CARRIER1                       ' carrier: BIG
  ST_PARTICLEBEAST                  ' carrier: BIG
  ' z = 1
  ST_LRCLOSETERROR                  ' missile-cruiser: MEDIUM
  ' for AI purpose
  ST_RANDOM
End Enum

' represents current ship-state
Enum enumShipStates
  SS_WARP = 0
  SS_MOVE
  SS_ATTACKINMOTION
  SS_ATTACK
  SS_STOP
  SS_INTERCEPTBACK
End Enum

Enum enumDirection
  SO_RIGHT = 0
  SO_LEFT
  ' additional directions
  SO_LEFTUP
  SO_LEFTDOWN
  SO_RIGHTUP
  SO_RIGHTDOWN
End Enum

' interceptor rotation constants
Enum enumRotation
 RO_NONE = 0
 RO_LEFT
 RO_RIGHT
End Enum

' weapon kinds
Enum enumWeapon
  SW_LASER = 0
  SW_GREENPUS
  SW_PARTICLE
  SW_MISSILES
  SW_GREENLASER
  SW_INTERCEPTOR
End Enum

' current ship animation
Enum enumShipAnimtaions
  SA_WARP = 0
  SA_TURN
  SA_MOVE
  SA_ATTACK
  'SA_STOP
  SA_DESTROYED
End Enum

' animation loop type
Enum enumAnimationLoop
  AL_NONE = 0
  AL_PLAY_REWIND
  AL_CYCLE
End Enum

' 'start' and 'end' frames representing an animation
Private Type typeAnimation
 Start  As Long
 End    As Long
 speed  As Long
 Loop   As enumAnimationLoop
End Type

Private Const ANIMATIONS = 5 - 1

Dim m_x            As Single                      ' x position of ship
Dim m_y            As Single                      ' y position of ship
Dim m_z            As Byte                        ' distance from player
Dim m_fFloat       As Single                      ' floating value
Dim m_xVel         As Single                      ' x velocity
Dim m_yVel         As Single                      ' y velocity
Dim m_MaxVel       As Single                      ' maximum ship velocity
Dim m_Angle        As Single                      ' angle at which ship is currently moving
Dim m_DestAngle    As Single                      ' destination point angle
Dim m_Rotation     As enumRotation                ' interceptor rotation position
Dim m_cx           As Integer                     ' ship width
Dim m_cy           As Integer                     ' ship height
Dim m_dx           As Integer                     ' x destination
Dim m_dy           As Integer                     ' y destination
Dim m_Type         As enumShipType                ' ship kind
Dim m_Dir          As enumDirection               ' What is the direction the Ship Comes from...
Dim m_State        As enumShipStates              ' ship-state
' animation vars
Dim m_arAnims(ANIMATIONS) As typeAnimation        ' animations array
Dim m_CurFrame     As Long
Dim m_frInc        As Long                        ' animation frame and frame incrementer
Dim m_frFace       As Long                        ' frame num of face_animation (left/right)
Dim m_ShipAnim     As enumShipAnimtaions
Dim m_bUpdateFrame As Boolean                     ' should animations be updated
' ...
Dim m_Weapon       As enumWeapon
Dim m_hitpoints    As Byte
Dim m_bytScore     As Byte                        ' score that player will recieve
Dim m_lExperience  As Long                        ' experience to give player
Dim m_Visible      As Boolean                     ' is this ship currently in use
Dim m_myIndex      As Byte                        ' ship array index
' weaponary stuff
'Dim m_bytBeamSize As Byte                        ' width of weapon beam
Dim m_AttackDelay  As Integer                     ' delay between shots
Dim m_AttackRange  As Integer                     ' maximum attack range
Dim m_Damage       As Integer                     ' ship damage
Dim m_Target       As enumEnemyObjective          ' Who's gonna get it ? ;)
Dim m_EnemyIndex   As Byte                        ' bunker class index ( if bunker's the target )
' carrier-interceptor vars
Dim m_bytInterceptors           As Integer '{!}   ' interceptors warped for the carrier ship
Dim m_bytOwner                  As Byte           ' if this's an interceptor class then we should know our OWNER index
Dim m_arChild(MAX_INTERCEPTORS) As Byte           ' carrier keeps a list of its' childs indecies
Dim m_lDieTime                  As Long           ' wick of destruction counter

Dim bNoMove As Boolean                            ' see UpdateMovingAI procedure
Dim bWarped As Boolean
Dim lwgind  As Long                               ' warpgate animation index


' /////////////////////////////////////////////////////////////
' //// Init Ship
' /////////////////////////////////////////////////////////////
Public Sub _
CreateShip(pos_x As Integer, pos_y As Integer, _
                      ByVal eType As enumShipType) ',
                      'eWeapon As enumWeapon)
 Dim fVar1 As Single, fVar2 As Single
 
 bNoMove = False    ' hm...
 bWarped = False    ' !!!!!
 
 m_x = CSng(pos_x)
 m_y = CSng(pos_y)
 m_fFloat = 0
 m_Type = eType
 'm_lAnimSpeed = FPS_ANIMS
 
 Call FillKindInfo(m_Type)                                ' set additional ship info depedant on it's kind
' m_x = m_x / m_z                                          ' prevent ships in 2nd plane go out of sight_range
' m_y = m_y / m_z
 
 'm_Dir = eDir
 'm_Weapon = eWeapon
 ' damage, attackrange, velocity ... depends on the weapon or ship type
 m_MaxVel = arVelocity(m_Type)
 m_Damage = arDamages(m_Weapon)
 m_AttackRange = arAttackRange(m_Weapon)
 m_Visible = True
 Call GetTarget(m_dx, m_dy)
 Call SetupPhysics
 Call FixFace
  
  ' do not warp interceptors
 If (eType = ST_INTERCEPTOR1) Then
  m_State = SS_MOVE
  'm_ShipAnim = SA_MOVE
  m_CurFrame = m_arAnims(m_ShipAnim).Start
  ' do not precalc m_x and m_y position for intercaptors because they've
  ' been called by an inner ship class and coords.'ve are already been calibrated for
  ' the second_plane
 Else
  ' warp that ship
  m_State = SS_WARP
  m_ShipAnim = SA_WARP
  m_CurFrame = m_arAnims(m_ShipAnim).Start
  m_frInc = 1
  ' prevent ships in 2nd plane go out of sight_range
  m_x = m_x / m_z
  m_y = m_y / m_z
  'lwgind = CreateWarpGate(CInt(m_x), CInt(m_y) + m_cy / 2, m_z)
  lwgind = CreateWarpGate(GetX, GetY, m_z)
 End If
 
End Sub

' /////////////////////////////////////////////////////////////
' //// Play ship warping animation
' /////////////////////////////////////////////////////////////
Private Sub _
Warp()
  
 
 If (m_CurFrame >= m_arAnims(m_ShipAnim).End And _
     (Not g_WarpGate(lwgind).Visible)) Then
  ' set ship state
  m_State = SS_MOVE
  ' set ship animation
  m_ShipAnim = SA_MOVE
  ' set animation frame at the MOVE animation start
  m_CurFrame = m_arAnims(m_ShipAnim).Start
 End If
 
End Sub

' /////////////////////////////////////////////////////////////
' //// Ship attacks in motion
' /////////////////////////////////////////////////////////////
Private Sub _
Intercept()
 
 'Static bytTurn As Byte
 Dim bTurning As Boolean
 Dim nDist As Integer
 Dim lk As Integer
  
 If (m_Rotation = RO_RIGHT) Then
  m_ShipAnim = SA_TURN
  m_Angle = m_Angle - 0.05 '* m_MaxVel
  
  If m_Angle < m_DestAngle Then
   m_Rotation = RO_NONE                                ' set no_rotation flag
   m_Angle = m_DestAngle                               ' set current angle to the destination angle
   m_DestAngle = GetAngle(m_x, m_y, m_dx, m_dy)        ' precalculate the destination angle again (for precision)
   If (m_DestAngle > m_Angle) Then
    m_Rotation = RO_LEFT
   Else
    m_Rotation = RO_RIGHT
   End If
   Call FixFace
   'm_xVel = Sgn(m_xVel) * m_MaxVel
   'm_yVel = Sgn(m_yVel) * m_MaxVel

   bTurning = False
  Else
   bTurning = True
  End If

 ElseIf (m_Rotation = RO_LEFT) Then
  m_ShipAnim = SA_TURN
  m_Angle = m_Angle + 0.05 '* m_MaxVel

  If m_Angle > m_DestAngle Then
   m_Angle = m_DestAngle
   m_Rotation = RO_NONE
   m_xVel = m_MaxVel
   m_yVel = m_MaxVel
   m_DestAngle = GetAngle(m_x, m_y, m_dx, m_dy)
   If (m_DestAngle > m_Angle) Then
    m_Rotation = RO_LEFT
   Else
    m_Rotation = RO_RIGHT
   End If
   Call FixFace
   
   'm_xVel = Sgn(m_xVel) * m_MaxVel
   'm_yVel = Sgn(m_yVel) * m_MaxVel
  
   bTurning = False
  Else
   bTurning = True
  End If
  
 ElseIf m_Rotation = RO_NONE Then
  m_ShipAnim = SA_MOVE
 '...
 ' Debug.Print "NO GO!"
 End If
 
 'Debug.Print m_Angle
 'm_x = m_x + m_xVel '+ (Sgn(m_xVel) * Cos(fa)) '* 2
 'm_y = m_y + m_yVel '+ (Sgn(m_yVel) * Sin(fa)) '* 2
 m_x = m_x + Cos(m_Angle) * m_xVel
 m_y = m_y + Sin(m_Angle) * m_yVel
 'If (m_y < 0 Or m_y > VISIBLE_AREA_CY) Then
 ' Call GetTarget(m_dx, m_dy)
 'End If
' Stop
 'Debug.Print "angle: " & m_Angle
 'Debug.Print "x: " & m_x
 'Debug.Print "y: " & m_y
 'Debug.Print "dx: " & m_dx
 'Debug.Print "dy: " & m_dy
 'Debug.Print "xVel: " & m_xVel
 'Debug.Print "yVel: " & m_yVel
 
 
 nDist = nGetDist2D(CInt(m_x), CInt(m_y), m_dx, m_dy)
 If (Not bTurning) Then
  If (nDist <= m_AttackRange) Then
'  Stop
  'm_State = SS_ATTACK
   'm_CurFrame = m_arAnims(SA_ATTACK).Start  ' ako stoi tuk togawa za interceptora w Attack() nqma da dade bAttack = True
   Call Attack
   Call GetTarget(m_dx, m_dy)
   Call SetupPhysics
  End If
 End If
 
 ' else intercept in space
 'fAng = GetAngle(m_x, m_y, m_dx, m_dy)
 'fAng = fAng + fGetRnd(0.05, 0.1)
 'If fAng > 6.28 Then fAng = fAng - 6.28
'
' m_x = m_x + Cos(fAng) * (Sgn(m_xVel) * m_MaxVel)
' m_y = m_y + Sin(fAng) * (Sgn(m_yVel) * m_MaxVel)
 
 ' update animations
 'If m_bUpdateFrame Then
 ' m_CurFrame = m_CurFrame + 1
 ' If (m_CurFrame > m_arAnims(SA_MOVE).End) Then
 '  m_CurFrame = m_arAnims(SA_MOVE).Start
 ' End If
 'End If
  
End Sub

' /////////////////////////////////////////////////////////////
' //// Move Interceptor back to mother-ship
' /////////////////////////////////////////////////////////////
Private Sub _
InterceptBack()
 
 Dim nDist As Integer
 Dim fReduce As Single
 
 m_x = m_x + Cos(m_DestAngle) * m_xVel
 m_y = m_y + Sin(m_DestAngle) * m_yVel

 ' if no carrier around engage suicide mission
 If (Not CShip(m_bytOwner).GetVisible) Then
  Call GetTarget(m_dx, m_dy)
  Call SetupPhysics
  Call FixFace
  m_State = SS_MOVE
 End If
 
 ' check distance
 If (Collide(GetRect, CShip(m_bytOwner).GetRect)) Then
  ' simulate interceptor-recharge by destroying this one (new one will be created)
  SetVisible = False
  If (m_Type = ST_INTERCEPTOR1 And CShip(m_bytOwner).GetVisible And _
   CShip(m_bytOwner).GetInterceptors > 0) Then _
    Call CShip(m_bytOwner).KillInterceptor(m_myIndex)
 End If
 
End Sub

' /////////////////////////////////////////////////////////////
' //// Move ship to the target destination
' /////////////////////////////////////////////////////////////
Private Sub _
Move()
 
 Dim nDist As Integer
 Dim fReduce As Single
 
 'm_x = m_x + m_xVel
 'm_y = m_y + m_yVel
 m_x = m_x + Cos(m_DestAngle) * m_xVel
 m_y = m_y + Sin(m_DestAngle) * m_yVel
 
 If (m_Type = ST_PARTICLEBEAST) Then
  nDist = Sqr((m_dx - m_x) ^ 2)
 Else
  nDist = nGetDist2D(CInt(m_x), CInt(m_y), m_dx, m_dy)
 End If
 ' reduce ship speed if it's less than 10 pixels from the target spot
 fReduce = m_MaxVel / 10
 If (nDist <= 6) Then
  m_xVel = m_xVel - fReduce
  m_yVel = m_yVel - fReduce
  If (m_xVel <= 0) Then m_xVel = 0 ' m_xVel + fReduce
  If (m_yVel <= 0) Then m_yVel = 0 ' m_yVel + fReduce
 End If
 
 If (nDist <= m_AttackRange) Then
  ' for interceptors
  If m_Type = ST_INTERCEPTOR1 Then           ' FIX THIS ---------
   m_State = SS_ATTACKINMOTION
   m_Rotation = RO_NONE
   m_Angle = m_DestAngle
   m_xVel = m_MaxVel
   m_yVel = m_MaxVel
   Call GetTarget(m_dx, m_dy)
   Call SetupPhysics
  ' for other ship kinds
  Else
   m_State = SS_ATTACK
  End If
  m_ShipAnim = SA_ATTACK
  m_CurFrame = m_arAnims(m_ShipAnim).Start
  '...
 End If
 
 ' update animations
 'If (m_bUpdateFrame) Then
 ' m_CurFrame = m_CurFrame + 1
 ' If (m_CurFrame > m_arAnims(SA_MOVE).End) Then
 '  m_CurFrame = m_arAnims(SA_MOVE).Start
 ' End If
 'End If
 
End Sub

' /////////////////////////////////////////////////////////////
' //// Make ship float in space
' /////////////////////////////////////////////////////////////
Private Sub _
Float()
 
 Static fFloat As Single
 
 'If m_bUpdateFrame Then
 ' m_CurFrame = m_CurFrame + 1
 ' If m_CurFrame > m_arAnims(SA_STOP).End Then _
 '    m_CurFrame = m_arAnims(SA_STOP).Start
 'End If
 
 ' do floating calculations
 fFloat = fFloat + m_fFloat
 If (fFloat > 2) Then m_fFloat = -m_fFloat
 If (fFloat < -2) Then m_fFloat = Abs(m_fFloat)
 m_y = m_y + m_fFloat

End Sub

' /////////////////////////////////////////////////////////////
' //// Update Ship
' /////////////////////////////////////////////////////////////
Public Sub _
Render()
 
 Static lFPSTime As Long
 Dim lTicks      As Long
 Dim rShip       As RECT
 
 lTicks = GetTicks()
 If (lFPSTime < lTicks) Then
    lFPSTime = lTicks + m_arAnims(m_ShipAnim).speed
    m_bUpdateFrame = True
    m_CurFrame = m_CurFrame + m_frInc
 Else
    m_bUpdateFrame = False
 End If

 'm_State = SS_ATTACKINMOTION
 Select Case m_State
   Case SS_WARP
    m_ShipAnim = SA_WARP
    Call Warp
    Exit Sub
    
   Case SS_MOVE
    m_ShipAnim = SA_MOVE
    Call Move
    
   Case SS_ATTACK
    Call Float
    m_ShipAnim = SA_ATTACK
    Call Attack
   
   Case SS_ATTACKINMOTION
    ' see if ship's to be destroyed
    If (m_lDieTime > 0 And m_lDieTime < GetTicks) Then
     Call Destroy
    End If
    
    'm_ShipAnim = SA_TURN
    Call Intercept
    'Call Attack
   
   Case SS_INTERCEPTBACK
    m_ShipAnim = SA_MOVE
    Call InterceptBack
    
   Case SS_STOP
    Call Float
    
   Case Else: '...
 End Select
  

 ' increment animation frame
 If (m_bUpdateFrame) Then
  
  ' onwards
  If (m_CurFrame > m_arAnims(m_ShipAnim).End) Then
    
   If (m_arAnims(m_ShipAnim).Loop = AL_PLAY_REWIND) Then
    m_frInc = -1
    m_CurFrame = m_arAnims(m_ShipAnim).End
   ElseIf (m_arAnims(m_ShipAnim).Loop = AL_CYCLE) Then
    m_CurFrame = m_arAnims(m_ShipAnim).Start
   ElseIf (m_arAnims(m_ShipAnim).Loop = AL_NONE) Then
    m_CurFrame = m_arAnims(m_ShipAnim).End
   End If
  
  ' backwards
  ElseIf (m_CurFrame < m_arAnims(m_ShipAnim).Start) Then
   m_frInc = 1
   m_CurFrame = m_CurFrame + m_frInc
  End If
  
 End If
 
 Dim dx As Long, dy As Long
 Dim frame As Long
 dx = m_x - wx / m_z
 dy = m_y - wy / m_z
 
 ' set facing
 Call FixFace
 ' prepare frame
 frame = m_CurFrame + (m_Dir * m_frFace)
 
 Select Case m_Type
    ' far enemies
    Case ST_INTERCEPTOR1
     Call BltFastGFX_HBM(dx, dy, g_Objects.Ship1(frame))
    Case ST_LRSNEAKY
     Call BltFastGFX_HBM(dx, dy, g_Objects.Ship3(frame))
    Case ST_SEVENTHFOX
     Call BltFastGFX_HBM(dx, dy, g_Objects.Ship4(frame))
    
    ' close enemies
    Case ST_LRCLOSETERROR
     'Call BltFastGFX_HBM(dx, dy, g_Objects.Ship5(frame))
    Case ST_CARRIER1
     Call BltFastGFX_HBM(dx, dy, g_Objects.Ship2(frame))
    Case ST_PARTICLEBEAST
     Call BltFastGFX_HBM(dx, dy, g_Objects.Ship5(frame))
     
  Case Else: '...
  End Select
  
 
 ' destroy ship if out of screen-range
 'If ((GetX / GetZ) < -50 Or (GetX / GetZ) < SCREEN_PIXEL_WIDTH + 50) Then SetVisible = False
 If ((GetY / GetZ) < -35 Or (GetY / GetZ) > SCREEN_PIXEL_HEIGHT + 50) Then Call Me.Destroy

End Sub

' /////////////////////////////////////////////////////////////
' //// Set's additional data dependent on Ship Kind
' /////////////////////////////////////////////////////////////
Private Sub _
FillKindInfo(eShip As enumShipType)
 
 Dim cn As Long
 
 ' Ships in Z = 2 ( Far Ships )
 ' ------------------------------
 If (eShip = ST_INTERCEPTOR1) Then
    m_cx = g_Objects.Ship1(0).cx
    m_cy = g_Objects.Ship1(0).cy
    m_z = 2
    m_fFloat = 0.05
    m_AttackDelay = 350                            ' in ms.
    m_Weapon = SW_LASER
    m_hitpoints = 5
    m_lExperience = m_hitpoints * 5
    m_bytScore = m_hitpoints * 2                   ' score to give player
    m_lDieTime = 0                                 ' reset destruction timer

    m_frFace = ((UBound(g_Objects.Ship1) \ 2)) + 1 ' get facing_change
    ' set animations
    m_arAnims(SA_WARP).Start = 0
    m_arAnims(SA_WARP).End = 0
    m_arAnims(SA_WARP).speed = FPS_ANIMS
    m_arAnims(SA_TURN).Start = 1
    m_arAnims(SA_TURN).End = 5
    m_arAnims(SA_TURN).speed = FPS_ANIMS + FPS_ANIMS
    m_arAnims(SA_MOVE).Start = 0
    m_arAnims(SA_MOVE).End = 0
    m_arAnims(SA_MOVE).speed = FPS_ANIMS
    m_arAnims(SA_ATTACK).Start = 0
    m_arAnims(SA_ATTACK).End = 0
    m_arAnims(SA_ATTACK).speed = FPS_ANIMS
    m_arAnims(SA_DESTROYED).Start = 0
    m_arAnims(SA_DESTROYED).End = 0
    m_arAnims(SA_DESTROYED).speed = FPS_ANIMS
 
 ElseIf (eShip = ST_LRSNEAKY) Then                 ' missile ship
    m_cx = g_Objects.Ship3(0).cx
    m_cy = g_Objects.Ship3(0).cy
    m_z = 2
    m_fFloat = 0.09
    m_AttackDelay = 1325
    m_Weapon = SW_MISSILES
    m_hitpoints = 6
    m_lExperience = m_hitpoints * 5                ' assign experience
    m_bytScore = m_hitpoints * 2                   ' score to give player
    m_frFace = ((UBound(g_Objects.Ship3) \ 2)) + 1 ' get facing_change
    m_arAnims(SA_WARP).Start = 0
    m_arAnims(SA_WARP).End = 0
    m_arAnims(SA_WARP).speed = FPS_ANIMS
    m_arAnims(SA_TURN).Start = 0
    m_arAnims(SA_TURN).End = 0
    m_arAnims(SA_TURN).speed = FPS_ANIMS
    m_arAnims(SA_MOVE).Start = 0
    m_arAnims(SA_MOVE).End = 0
    m_arAnims(SA_MOVE).speed = FPS_ANIMS
    m_arAnims(SA_ATTACK).Start = 0
    m_arAnims(SA_ATTACK).End = 0
    m_arAnims(SA_ATTACK).speed = FPS_ANIMS
    m_arAnims(SA_DESTROYED).Start = 0
    m_arAnims(SA_DESTROYED).End = 0
    m_arAnims(SA_DESTROYED).speed = FPS_ANIMS
 
 ElseIf (eShip = ST_LRCLOSETERROR) Then            ' missile ship z = 1
    m_cx = g_Objects.Ship5(0).cx
    m_cy = g_Objects.Ship5(0).cy
    m_z = 1
    m_fFloat = 0.04
    m_AttackDelay = 2000
    m_Weapon = SW_MISSILES
    m_hitpoints = 6
    m_bytScore = m_hitpoints * 2                   ' score to give player
    
 ElseIf (eShip = ST_SEVENTHFOX) Then               ' direct-attack ship
    m_cx = g_Objects.Ship4(0).cx
    m_cy = g_Objects.Ship4(0).cy
    m_z = 2
    m_fFloat = 0.09
    m_AttackDelay = 1000
    m_Weapon = SW_GREENLASER
    m_hitpoints = 5
    m_lExperience = m_hitpoints * 5                ' assign experience
    m_bytScore = m_hitpoints * 2                   ' score to give player
    m_frFace = ((UBound(g_Objects.Ship4) \ 2)) + 1 ' get facing_change
    
    m_arAnims(SA_WARP).Start = 0
    m_arAnims(SA_WARP).End = 0
    m_arAnims(SA_WARP).speed = FPS_ANIMS
    m_arAnims(SA_TURN).Start = 0
    m_arAnims(SA_TURN).End = 0
    m_arAnims(SA_TURN).speed = FPS_ANIMS
    m_arAnims(SA_MOVE).Start = 0
    m_arAnims(SA_MOVE).End = 0
    m_arAnims(SA_MOVE).speed = FPS_ANIMS
    m_arAnims(SA_ATTACK).Start = 0
    m_arAnims(SA_ATTACK).End = 7
    m_arAnims(SA_ATTACK).speed = FPS_ANIMS
    m_arAnims(SA_ATTACK).Loop = AL_CYCLE
    m_arAnims(SA_DESTROYED).Start = 0
    m_arAnims(SA_DESTROYED).End = 0
    m_arAnims(SA_DESTROYED).speed = FPS_ANIMS
    
    ' --- Carriers ---------------
 ElseIf (eShip = ST_CARRIER1) Then
    m_cx = g_Objects.Ship2(0).cx
    m_cy = g_Objects.Ship2(0).cy
    m_z = 2
    m_fFloat = 0.02
    m_AttackDelay = 1650                           ' in ms.
    m_Weapon = SW_INTERCEPTOR
    m_hitpoints = 60
    m_lExperience = m_hitpoints * 5                ' assign experience
    m_bytScore = m_hitpoints * 2                   ' score to give player
    m_bytInterceptors = 0                          ' reset interceptors-warped counter
    For cn = 0 To MAX_INTERCEPTORS
     m_arChild(cn) = 0                             ' reset track-indecies
    Next
    m_frFace = ((UBound(g_Objects.Ship2) \ 2)) + 1 ' get facing_change
    ' // set animations
    m_arAnims(SA_WARP).Start = 0
    m_arAnims(SA_WARP).End = 0
    m_arAnims(SA_WARP).speed = FPS_ANIMS
    m_arAnims(SA_TURN).Start = 0
    m_arAnims(SA_TURN).End = 0
    m_arAnims(SA_TURN).speed = FPS_ANIMS
    m_arAnims(SA_MOVE).Start = 0
    m_arAnims(SA_MOVE).End = 0
    m_arAnims(SA_MOVE).speed = FPS_ANIMS
    m_arAnims(SA_DESTROYED).Start = 0
    m_arAnims(SA_DESTROYED).End = 0
    m_arAnims(SA_DESTROYED).speed = FPS_ANIMS

    m_arAnims(SA_ATTACK).Start = 1
    m_arAnims(SA_ATTACK).End = 13
    m_arAnims(SA_ATTACK).speed = FPS_ANIMS + FPS_ANIMS + FPS_ANIMS
    m_arAnims(SA_ATTACK).Loop = AL_PLAY_REWIND
 
 ElseIf eShip = ST_PARTICLEBEAST Then
    m_cx = g_Objects.Ship5(2).cx
    m_cy = g_Objects.Ship5(2).cy
    m_z = 2
    m_fFloat = 0.01
    m_AttackDelay = 3000                           ' in ms.
    m_Weapon = SW_PARTICLE
    m_hitpoints = 100
    m_bytScore = m_hitpoints * 2                   ' score to give player
    m_lExperience = m_hitpoints * 5                ' assign experience
    m_frFace = ((UBound(g_Objects.Ship5) \ 2)) + 1 ' get facing_change
     ' // set animations
    m_arAnims(SA_WARP).Start = 0
    m_arAnims(SA_WARP).End = 0
    m_arAnims(SA_WARP).speed = FPS_ANIMS
    m_arAnims(SA_TURN).Start = 0
    m_arAnims(SA_TURN).End = 0
    m_arAnims(SA_TURN).speed = FPS_ANIMS
    m_arAnims(SA_MOVE).Start = 0
    m_arAnims(SA_MOVE).End = 0
    m_arAnims(SA_MOVE).speed = FPS_ANIMS
    m_arAnims(SA_DESTROYED).Start = 0
    m_arAnims(SA_DESTROYED).End = 0
    m_arAnims(SA_DESTROYED).speed = FPS_ANIMS

    m_arAnims(SA_ATTACK).Start = 1
    m_arAnims(SA_ATTACK).End = 5
    m_arAnims(SA_ATTACK).speed = FPS_ANIMS + FPS_ANIMS + FPS_ANIMS
    m_arAnims(SA_ATTACK).Loop = AL_NONE
 
 'm_HitPoints = 15
 End If
 
End Sub

' /////////////////////////////////////////////////////////////
' //// Make ship attack
' /////////////////////////////////////////////////////////////
Private Sub _
Attack()
 
 Static lDelay As Long
 Dim sx As Integer, sy As Integer
 Dim dx As Integer, dy As Integer
 Dim bAttack As Boolean
 Dim cn      As Long
 Dim i       As Long
 
 ' update animation
 'If m_bUpdateFrame Then
 ' m_CurFrame = m_CurFrame + 1
 ' If (m_CurFrame > m_arAnims(SA_ATTACK).End) Then _
 '     m_CurFrame = m_arAnims(SA_ATTACK).Start
 '     bAttack = True
 'End If
 ' since the current animation is attack mode wait until it's done and attack
 If (m_CurFrame >= m_arAnims(m_ShipAnim).End) Then bAttack = True
 
 ' do damage dependent on the armament
 sx = GetX
 sy = GetY
 Select Case m_Weapon
 
   Case SW_LASER
    If (bAttack And _
       lDelay < GetTicks) Then
       
       ' set weapon delay
       lDelay = GetTicks + m_AttackDelay '(m_Weapon * 350) + 150
       bAttack = False
       
       ' set source
       dx = nGetRnd(10, 25)
       dy = nGetRnd(10, 25)
       dx = dx + GetX
       dy = dy + GetY
       Call CreateLaserCut(sx, sy, m_z, dx, dy, 15, 15, m_Weapon, m_myIndex)
       ' PLAY SFX
       Call DSPlaySound(g_dsSfx(SFX_NORMALASER), False, (GetX - wx), SFX_VOLUMEFAR)
    End If
   
   Case SW_GREENLASER
    If (bAttack And _
       lDelay < GetTicks) Then
       
       ' set weapon delay
       lDelay = GetTicks + m_AttackDelay '(m_Weapon * 350) + 150
       bAttack = False
       
       ' set source
       Call CreateLaserCut(sx, sy, m_z, m_dx, m_dy, 15, 15, m_Weapon, m_myIndex)
    
       ' PLAY SOUND
       Call DSPlaySound(g_dsSfx(nGetRnd(SFX_GREENLASER1, SFX_GREENLASER2)), False, (GetX - wx), SFX_VOLUMEFAR)
       
    End If
          
   Case SW_INTERCEPTOR
    If ( _
       lDelay < GetTicks) Then
       
     ' check if interceptor limit has been exceeded
     If (m_bytInterceptors >= MAX_INTERCEPTORS) Then Exit Sub
       
     ' set weapon delay
     lDelay = GetTicks + m_AttackDelay '(m_Weapon * 350) - nGetRnd(150, 800)
     bAttack = False
     
     ' check for empty enemy class
     Do While (cn < MAX_ENEMIES)
      
      If (Not CShip(cn).GetVisible) Then
       Call CShip(cn).CreateShip(Me.GetX, Me.GetY, ST_INTERCEPTOR1)
       
       For i = 0 To MAX_INTERCEPTORS
        If (m_arChild(i) = 0) Then
         m_arChild(i) = cn                          ' save child index
         Exit For
        End If
       Next
       
       m_bytInterceptors = m_bytInterceptors + 1    ' increment interc. created counter
       CShip(cn).SetOwner = m_myIndex               ' tell the child that we're its' mother (uhm, weird a?) ;)
       Exit Do
      End If
     
     cn = cn + 1
     Loop
     
     ' PLAY FLY SOUND
     Call DSPlaySound(g_dsSfx(nGetRnd(SFX_INTERCEPT1, SFX_INTERCEPT2)), False, (GetX - wx), SFX_VOLUMEFAR)
    End If
    
   Case SW_GREENPUS
    '...
   
   Case SW_PARTICLE
    If (bAttack And _
       lDelay < GetTicks) Then
    
     ' set weapon delay
     lDelay = GetTicks + m_AttackDelay 'nGetRnd(2000, 4000)
     bAttack = False
     
     Call CreateParticleExplosion(GetX() - 20, m_y + 199)
     Call CreateParticleLaser(GetX() - 36, GetY(), m_myIndex)
     '...
     ' PLAY BEAM PREPARATION SOUND
    End If
    
   Case SW_MISSILES
    If (bAttack And _
       lDelay < GetTicks) Then
       
       ' set weapon delay
       lDelay = GetTicks + m_AttackDelay '3500 '(m_Weapon * 350) + 150
       bAttack = False
       
       ' hit a bunker
       If (m_EnemyIndex <> 255) Then
        Call CreateMissile(sx, sy, m_z, m_dx, m_dy, 10, 10, CBunker(m_EnemyIndex).GetPos, m_myIndex)
       Else ' hit Earth
        Call CreateMissile(sx, sy, m_z, m_dx, m_dy, 20, 20, BP_VERYFAR, m_myIndex)
       End If
    End If
   
 Case Else: '...
 End Select
  
End Sub

' /////////////////////////////////////////////////////////////
' //// Set enemy target spot dependant on it's kind
' /////////////////////////////////////////////////////////////
Private Sub _
GetTarget(dx As Integer, dy As Integer)
 
 Dim cn As Long
 Dim fAngle As Single
 
 Select Case m_Type
  
   Case ST_INTERCEPTOR1, ST_LRSNEAKY, ST_SEVENTHFOX, _
        ST_CARRIER1
    
    If (CBattleStation.GetVisible And nGetRnd(0, 1000) > 700) Then
     dx = CBattleStation.GetX + nGetRnd(-15, 15)
     dy = CBattleStation.GetY + nGetRnd(-15, 15)
     m_Target = EO_BATTLESTATION
    Else
     dx = (nGetRnd(-10, g_cxEarth) + g_xEarth)
     dy = (nGetRnd(0, g_cyEarth - 100) + g_yEarth)
     m_Target = EO_EARTH
    End If
    
   Case ST_PARTICLEBEAST
    
    dx = nGetRnd(25, g_cxEarth - 25) + g_xEarth
    dy = (nGetRnd(90, g_cyEarth - 100) + g_yEarth)
    m_Target = EO_EARTH
    
   Case ST_LRCLOSETERROR
     ' get a random enemy bunker as a target
     m_EnemyIndex = nGetRnd(0, MAX_BUNKERS)
     
     ' if bunker is already dead then choose the first found
     If (Not CBunker(m_EnemyIndex).GetVisible) Then
      For cn = 0 To MAX_BUNKERS
       If (CBunker(cn).GetVisible) Then
        m_EnemyIndex = cn
        Exit For
       End If
        m_EnemyIndex = 255
      Next
     End If
        
     ' if no-operating bunker was found '{!}NOT COMPLETED
     If (m_EnemyIndex = 255) Then
      dx = (nGetRnd(50, g_cxEarth - 50) + g_xEarth)
      dy = (nGetRnd(50, g_cyEarth - 50) + g_yEarth)
      m_Target = EO_EARTH
     
     Else
      dx = CBunker(m_EnemyIndex).GetX
      dy = CBunker(m_EnemyIndex).GetY
      m_Target = EO_MOONBUNKER
     End If
         
   Case Else
    Debug.Print "Error in ShipClass: GetTarget function"
 
 End Select
 
 ' set animation frame
 m_CurFrame = m_arAnims(SA_MOVE).Start
End Sub

' /////////////////////////////////////////////////////////////
' //// setup ship physics
' /////////////////////////////////////////////////////////////
Private Sub _
SetupPhysics()

 Dim fAngle As Single

 ' calculate velocity
 fAngle = GetAngle(m_x, m_y, m_dx, m_dy)
 m_DestAngle = fAngle
 m_xVel = m_MaxVel
 m_yVel = m_MaxVel
 If (m_Type = ST_PARTICLEBEAST) Then m_yVel = 0
 'm_yVel = m_MaxVel
 
 ' setup rotation for interceptors
 If (m_DestAngle > m_Angle) Then
  m_Rotation = RO_LEFT
  'm_Angle = m_Angle '- PI
 Else
  m_Rotation = RO_RIGHT
  'm_Angle = m_Angle '+ PI
 End If

End Sub


' /////////////////////////////////////////////////////////////
' //// Orientate ship
' /////////////////////////////////////////////////////////////
Private Sub _
FixFace()
 'm_xVel = Cos(fAngle) * m_MaxVel
 'm_yVel = Sin(fAngle) * m_MaxVel
 
 ' make ship faces the target
 If (max(CInt(m_x), CInt(m_dx))) Then
  m_Dir = SO_RIGHT
  'm_xVel = -m_xVel
 Else
  m_Dir = SO_LEFT
 End If
 
 ' fix y-velocity
 'If (Max(CInt(m_y), CInt(m_dy))) Then m_yVel = -m_yVel

End Sub


' /////////////////////////////////////////////////////////////
' //// Kill an interceptor
' /////////////////////////////////////////////////////////////
Public Sub _
KillInterceptor(bytIndex As Byte)
 
 Dim cn As Long
 
 ' clear index
 For cn = 0 To MAX_INTERCEPTORS
  If (m_arChild(cn) = bytIndex) Then
   m_arChild(cn) = 0
   Exit For
  End If
 Next
 ' decrement counter
 m_bytInterceptors = m_bytInterceptors - 1
End Sub


' /////////////////////////////////////////////////////////////
' //// Take hitpoints from enemy
' /////////////////////////////////////////////////////////////
Public Sub _
DoEnemyDamage()
 
 Select Case m_Target
  Case EO_EARTH
   g_hpEarth = g_hpEarth - m_Damage
  
  Case EO_MOONBUNKER
    ' get a new target if current's been destroyed
    'If Not cBunker(m_EnemyIndex).GetVisible Then
    '   Call GetTarget(m_dx, m_dy)
    '   bNoMove = False
    'Else
      CBunker(m_EnemyIndex).DoDamage = m_Damage              ' do damage to current bunker target
    'End If
    If (Not CBunker(m_EnemyIndex).GetVisible) Then
       Call GetTarget(m_dx, m_dy)
       Call SetupPhysics
       bNoMove = False
    End If
    
  Case EO_BATTLESTATION
    CBattleStation.DoDamage = m_Damage
    
  Case EO_PLAYERBUNKER
   ' ...
   
  Case Else
   Debug.Print "ERROR: Ship has no target!"
 End Select
 
 
End Sub


' /////////////////////////////////////////////////////////////
' //// Destroy this ship class
' /////////////////////////////////////////////////////////////
Public Sub _
Destroy()
 
 m_hitpoints = 0
 CreateExplosion GetX, GetY, m_z, ET_BIG + 3 * (m_z - 1)
 SetVisible = False                                  ' destroy the ship
 m_lDieTime = 0
 m_bytOwner = 0
 
End Sub

Public Function GetRect() As RECT                            ' get ship_blitting rect.
  Dim rOut As RECT
  
  Call SetRect(rOut, m_x, m_y, m_x + m_cx, m_y + m_cy)
  Call CopyRect(GetRect, rOut)
  'Call SetRect(GetRect, m_x \ GetZ, m_y \ GetZ, m_x \ GetZ + m_cx, m_y \ GetZ + m_cy)
End Function

Public Function GetRectOnPlane(z As Byte) As RECT            ' get coords. on passed plane
  Call SetRect(GetRect, m_x \ z, m_y \ z, m_x \ z + m_cx, m_y \ z + m_cy)
  'Call SetRect(GetRect, GetX \ GetZ, GetY \ GetZ, GetX \ GetZ + m_cx, GetY \ GetZ + m_cy)
End Function

' Info Setting props.
'-------------------------------------------------------------------

Public Property Let SetOwner(bytOwner As Byte)          ' set child-ship onwer (interceptors)
 m_bytOwner = bytOwner
End Property


Public Property Let DoDamage(nDamage As Integer)
 
 Dim nHitPoints As Integer                              ' temp hitpoints
 Dim cn         As Long                                 ' local counter
 
 nHitPoints = CInt(m_hitpoints)
 ' do damage to ship
 nHitPoints = nHitPoints - nDamage
 
 If (nHitPoints <= 0) Then
    m_hitpoints = 0
    'CreateExplosion GetX, GetY, m_z, ET_BIG Or m_z      ' create en exlposion
    ' score player
    g_Player.score = g_Player.score + GetScore()
    g_Player.Exp = g_Player.Exp + m_lExperience
    ' explode
    Call CreateExplosion(GetX, GetY, m_z, ET_BIG + 3 * (m_z - 1))
    ' PLAY die SOUND and add Kills to player {!}
    Select Case (m_Type)
      ' --- big ships
      Case ST_CARRIER1
        Call DSPlaySound(g_dsSfx(SFX_BIGBLAST2), False, (GetX - wx))
        g_Player.kills = g_Player.kills + GetInterceptors()
        ' carrier type ships always give bonus
        Call CreateBonus(GetX + nGetRnd(-5, 5), GetY + nGetRnd(-5, 5), nGetRnd(1, 5))
        
      Case ST_PARTICLEBEAST
        Call DSPlaySound(g_dsSfx(SFX_BIGBLAST2), False, (GetX - wx))
        g_Player.kills = g_Player.kills + 1
        Call CreateBonus(GetX + nGetRnd(-5, 5), GetY + nGetRnd(-5, 5), nGetRnd(1, 2))
        
      ' --- small ships
      Case Else
        Call DSPlaySound(g_dsSfx(SFX_BIGBLAST1), False, (GetX - wx), SFX_VOLUMEFAR)
        g_Player.kills = g_Player.kills + 1
        ' 1:5 chance for a small ship to give bonus
        If (nGetRnd(0, 1000) < 200) Then Call CreateBonus(GetX + nGetRnd(-5, 5), GetY + nGetRnd(-5, 5), nGetRnd(3, 5))
    
    End Select
    ' kill
    SetVisible = False                                  ' destroy the ship
    
    ' if this is an interceptor-ship then the mothership must be informed of
    ' our destruction (if it hasn't been destroyed, yet AND we're not wicked)
    If (m_Type = ST_INTERCEPTOR1 And CShip(m_bytOwner).GetVisible And (Not GetWick)) Then
     Call CShip(m_bytOwner).KillInterceptor(m_myIndex)
    ' tell(/destroy) our childs that(/'cos) we've been destroyed
    ElseIf (m_Type = ST_CARRIER1) Then
     For cn = 0 To m_bytInterceptors
      ' instigate destruction wick
      CShip(m_arChild(cn)).SetWick = True
     Next
    End If
 
 ' go back to mother-ship (or any-ship around, if there's space) if hitpoints are low
 ElseIf (nHitPoints <= 4 And m_Type = ST_INTERCEPTOR1 And _
         CShip(m_bytOwner).GetVisible And CShip(m_bytOwner).GetInterceptors <= MAX_INTERCEPTORS) Then
    m_dx = CShip(m_bytOwner).GetX
    m_dy = CShip(m_bytOwner).GetY
    ' get angle
    m_DestAngle = GetAngle(GetX, GetY, m_dx, m_dy)
    ' set face
    Call FixFace
    ' change states
    m_State = SS_INTERCEPTBACK
 
 Else
    m_hitpoints = CByte(nHitPoints)
 End If

End Property

Public Property Let SetWick(bWick As Boolean)
 m_lDieTime = GetTicks + nGetRnd(1000, 2200)
End Property

Public Property Let SetVisible(bVisible As Boolean)
 m_Visible = bVisible
End Property

Public Property Let SetIndex(bytIndex As Byte)
 m_myIndex = bytIndex
End Property

' Info retrieving props.
'-------------------------------------------------------------------

Public Property Get GetScore() As Byte
 GetScore = m_bytScore
End Property

Public Property Get GetInterceptors() As Integer
 GetInterceptors = m_bytInterceptors
End Property

Public Property Get GetWick() As Boolean
 If (m_lDieTime > 0) Then GetWick = True
End Property

Public Property Get GetOwner()                           ' get child-ship onwer (interceptors)
 GetOwner = m_bytOwner
End Property

Public Property Get GetX() As Integer
 GetX = m_x + m_cx / 2                                   ' get middle x
End Property

Public Property Get GetY() As Integer
 GetY = m_y + m_cy / 2                                   ' get middle y
End Property

Public Property Get GetZ() As Byte
 GetZ = m_z
End Property

Public Property Get GetVisible() As Boolean
 GetVisible = m_Visible
End Property

Public Property Get GetIndex() As Byte                   ' get class array index
 SetIndex = m_myIndex
End Property

Public Property Get GetHitPoints() As Integer
 GetHitPoints = CInt(m_hitpoints)

End Property

