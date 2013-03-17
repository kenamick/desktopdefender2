VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBattleStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' battlestation missile-object
Private Type typeMissile
  x          As Single
  y          As Single
  dx         As Single
  dy         As Single
  ang        As Single
  enemyIndex As Long
  enemyType  As enumBunkerTarget
  bActive    As Boolean
End Type

' constants
Private Const BSMISSILES_MAX         As Long = 19
Private Const BSMISSILES_SPEED       As Single = 1.2
Private Const BSMISSILES_DAMAGE      As Single = 0.4
Private Const BSMISSILES_RANGE       As Single = 800#


' vars
Private m_x                          As Single
Private m_y                          As Single
Private m_Frame                      As Long
Private m_bVisible                   As Boolean
Private m_bActive                    As Boolean
Private m_arMissile(BSMISSILES_MAX)  As typeMissile
Private m_hitpoints                  As Integer



Private Sub Class_Initialize()
 m_hitpoints = 100
End Sub

'///////////////////////////////////////////////////////////////////////////////////
'//// Name: Create()
'//// Desc: Initalize Battlestation
'///////////////////////////////////////////////////////////////////////////////////
Public Sub _
Create(Optional bActive As Boolean = False)

 ' adjust position
 m_x = g_xEarth + g_cxEarth + 52
 m_y = g_yEarth + 18
 m_hitpoints = 100
 m_bVisible = True
 m_bActive = bActive
 Erase m_arMissile()
 'm_Frame = 0
 
End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: UpdateMissiles()
'//// Desc: update missile position and check collisions
'///////////////////////////////////////////////////////////////////////////////////
Private Sub _
UpdateMissiles()

 
 Dim cn            As Long
 Dim enemyIndex    As Long
 Dim benemyVisible As Boolean
 Dim fdistance     As Single
 Dim fex           As Single
 Dim fey           As Single
 Dim rSrc          As RECT
 Dim rDest         As RECT
 Dim bltx          As Long
 Dim blty          As Long
 
 For cn = 0 To BSMISSILES_MAX
 
  With m_arMissile(cn)
   If (m_arMissile(cn).bActive) Then
  
    enemyIndex = m_arMissile(cn).enemyIndex
  
    ' get target coordinates
    Select Case (m_arMissile(cn).enemyType)
  
     Case BTARGET_SHIP
       fex = CShip(enemyIndex).GetX()
       fey = CShip(enemyIndex).GetY()
       benemyVisible = CShip(enemyIndex).GetVisible()
       Call CopyRect(rDest, CShip(enemyIndex).GetRect)
          
     Case BTARGET_METEOR
       fex = g_Meteor(enemyIndex).x
       fey = g_Meteor(enemyIndex).y
       benemyVisible = g_Meteor(enemyIndex).Visible
       Call SetRect(rDest, fex, fey, fex + g_Meteor(enemyIndex).cx, fey + g_Meteor(enemyIndex).cy)
       
     Case BTARGET_MISSILE
       fex = g_Missile(enemyIndex).x
       fey = g_Missile(enemyIndex).y
       benemyVisible = g_Missile(enemyIndex).Visible
       Call SetRect(rDest, fex, fey, fex + 5, fey + 5)
    
    End Select
    
    
    ' move missile
    .x = .x + Cos(.ang) * BSMISSILES_SPEED
    .y = .y + Sin(.ang) * BSMISSILES_SPEED
       
    ' blit missile
    bltx = .x - wx / PLANE_FAR
    blty = .y - wy / PLANE_FAR
    Call BltFastGFX_HBM(bltx, blty, g_Objects.bs_missile)
       
    If (benemyVisible) Then
    
     ' if target has moved then change the destination angle and save position
     If (fex <> .dx Or fey <> .dy) Then
      .ang = GetAngle(.x, .y, .dx, .dy)
      .dx = fex
      .dy = fey
     End If
     
      ' set rocket rect
      Call SetRect(rSrc, .x, .y, .x + 4, .y + 4)
      
      ' check for collisions and do damage is so...
      If (Collide(rDest, rSrc)) Then
      
       If (.enemyType = BTARGET_METEOR) Then
        g_Meteor(enemyIndex).HP = g_Meteor(enemyIndex).HP - 4
        .bActive = False
        Call CreateExplosion(CInt(.x), CInt(.y), PLANE_FAR, ET_SMALL_FAR)
        
       ElseIf (.enemyType = BTARGET_SHIP) Then
        CShip(enemyIndex).DoDamage = 1
        .bActive = False
        Call CreateExplosion(CInt(.x), CInt(.y), PLANE_FAR, ET_SMALL_FAR)
        
       ElseIf (.enemyType = BTARGET_MISSILE) Then
        g_Missile(enemyIndex).Visible = False
        .bActive = False
        Call CreateExplosion(CInt(.x), CInt(.y), PLANE_FAR, ET_SMALL_FAR)
        
        '...
       End If
      
      End If
    
    Else
     .bActive = False
     Call CreateExplosion(CInt(.x), CInt(.y), PLANE_FAR, ET_SMALLBLUE_FAR)
    End If
  
   
   End If ' end bActive check
  End With
  
 Next cn ' end loop

End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: FireMissile()
'//// Desc: Fire-up a guided missile
'///////////////////////////////////////////////////////////////////////////////////
Private Sub _
FireMissile(enemyType As enumBunkerTarget, enemyIndex As Long)

 Dim cn  As Long
 Dim fex As Single
 Dim fey As Single
 
 ' get target position
 If (m_arMissile(cn).enemyType = BTARGET_SHIP) Then
   fex = CShip(enemyIndex).GetX()
   fey = CShip(enemyIndex).GetY()
 ElseIf (m_arMissile(cn).enemyType = BTARGET_METEOR) Then
   fex = g_Meteor(enemyIndex).x
   fey = g_Meteor(enemyIndex).y
 ElseIf (m_arMissile(cn).enemyType = BTARGET_MISSILE) Then
   fex = g_Missile(enemyIndex).x
   fey = g_Missile(enemyIndex).y
 End If
   
 
 ' look for empty missile slot
 cn = 0
 Do While (cn < BSMISSILES_MAX)
   
   ' find and create....
   If (Not m_arMissile(cn).bActive) Then
    m_arMissile(cn).x = GetX()
    m_arMissile(cn).y = GetY()
    m_arMissile(cn).dx = fex
    m_arMissile(cn).dy = fey
    m_arMissile(cn).ang = GetAngle(m_arMissile(cn).x, m_arMissile(cn).y, fex, fey)
    m_arMissile(cn).enemyType = enemyType
    m_arMissile(cn).enemyIndex = enemyIndex
    m_arMissile(cn).bActive = True
    Exit Do
   End If
  
   cn = cn + 1
 Loop
  
 
End Sub



'///////////////////////////////////////////////////////////////////////////////////
'//// Name: GetTarget()
'//// Desc: Get a target in sight
'///////////////////////////////////////////////////////////////////////////////////
Private Sub _
GetTarget()

 Dim cn         As Long
 Dim lTicks     As Long
 Static lTime   As Long
 
 lTicks = GetTicks()
 If (lTime > lTicks) Then Exit Sub
 
 ' fire up missile every ???? seconds
 lTime = GetTicks() + 1500
 
 
 ' check for an enemy spaceship
 For cn = 0 To MAX_ENEMIES
  If (CShip(cn).GetVisible) Then
   
   If (fGetDist2D(CShip(cn).GetX, CShip(cn).GetY, GetX, GetY) < BSMISSILES_RANGE) Then
    Call FireMissile(BTARGET_SHIP, cn)
   End If
  
  End If
 Next cn
 
 
 ' look for a meteor threat
 For cn = 0 To MAX_METEORS
   If (g_Meteor(cn).Visible And (g_Meteor(cn).Data And MC_FAR)) Then
    
    If (nGetDist2D(CInt(g_Meteor(cn).x), CInt(g_Meteor(cn).y), GetX, GetY) < BSMISSILES_RANGE) Then
      Call FireMissile(BTARGET_METEOR, cn)
    End If
   
   End If
 Next
 
 
 ' look for enemy missile
 For cn = 0 To MAX_MISSILES
  If (g_Missile(cn).Visible) Then
    
    If (nGetDist2D(CInt(g_Missile(cn).x), CInt(g_Missile(cn).y), GetX, GetY) < BSMISSILES_RANGE) Then
     Call FireMissile(BTARGET_MISSILE, cn)
    End If
  
  End If
 Next


End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: Update()
'//// Desc: update object actions and animations
'///////////////////////////////////////////////////////////////////////////////////
Public Sub _
Update()

 If (Not m_bVisible) Then Exit Sub

 Dim dx          As Long
 Dim dy          As Long
 Dim lTicks      As Long
 Static lTime    As Long
 Static lHPTime  As Long
 
 If (m_bActive) Then
  ' do station engage-AI
  Call GetTarget
  Call UpdateMissiles
 End If

 ' repair hp
 lTicks = GetTicks()
 If (lHPTime < lTicks) Then
  lHPTime = lTicks + 800
  If (m_hitpoints < 100) Then m_hitpoints = m_hitpoints + 1
  If (m_hitpoints <= 0) Then Call Destroy
 End If
 
 ' compute position
 dx = m_x - wx / PLANE_FAR
 dy = m_y - wy / PLANE_FAR
  
 ' work out animation frame
 If (lTime < lTicks) Then
  m_Frame = m_Frame + 1
  If (m_Frame > 14) Then m_Frame = 0
  lTime = lTicks + FPS_ANIMS + FPS_ANIMS
 End If
 
 ' blit battlestation
 Call BltFastGFX_HBM(dx, dy, g_Objects.BattleStation(m_Frame))

End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: Destroy()
'//// Desc: destroy the battlestation
'///////////////////////////////////////////////////////////////////////////////////
Public Sub _
Destroy()

 If (Not m_bVisible) Then Exit Sub
 
 Call CreateExplosion(GetX(), GetY(), PLANE_FAR, ET_BIG_FAR)
 m_bVisible = False

End Sub


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: DoDamage()
'//// Desc: do damage to the battlestatoin
'///////////////////////////////////////////////////////////////////////////////////
Public Property Let _
DoDamage(nDamage As Integer)

 If (Not m_bVisible) Then Exit Property
 
 m_hitpoints = m_hitpoints - nDamage
 If (m_hitpoints < 0) Then Call Destroy

End Property


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: GetHitPoints()
'//// Desc: return hitpoints
'///////////////////////////////////////////////////////////////////////////////////
Public Property Get _
GetHitPoints() As Long

 GetHitPoints = m_hitpoints

End Property


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: GetX()
'//// Desc: get horizontal position
'///////////////////////////////////////////////////////////////////////////////////
Public Property Get _
GetX() As Long

 GetX = m_x + g_Objects.BattleStation(m_Frame).cx / 2

End Property


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: GetY()
'//// Desc: get vertical position
'///////////////////////////////////////////////////////////////////////////////////
Public Property Get _
GetY() As Long

 GetY = m_y + g_Objects.BattleStation(m_Frame).cy / 2

End Property


'///////////////////////////////////////////////////////////////////////////////////
'//// Name: GetVisible()
'//// Desc: is the bs active
'///////////////////////////////////////////////////////////////////////////////////
Public Property Get _
GetVisible() As Boolean

 GetVisible = m_bVisible

End Property



Private Sub Class_Terminate()
 
 Erase m_arMissile()
 
End Sub


