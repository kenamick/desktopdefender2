VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBunker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem *** Bunker Class
Rem *** Purpose: Inits, draws handles bunkers actions

Enum enumBunkerTarget
 BTARGET_SHIP = 0
 BTARGET_METEOR
 BTARGET_MISSILE
End Enum

Enum enumBunkerPosition
  BP_CLOSE = 0
  BP_FAR
  BP_VERYFAR
End Enum

Enum enumBunkerType
  BT_WEAK = 1
  BT_STRONG
  BT_STONE 'COLD
End Enum

Dim m_x  As Integer
Dim m_y  As Integer
Dim m_cx As Integer
Dim m_cy As Integer
Dim m_BP As enumBunkerPosition                      ' distance
Dim m_BT As enumBunkerType
Dim m_hitpoints As Byte                             ' hitpoints left
Dim m_Rotation As enumDirection                     ' set rotation animation
Dim m_bFiring As Boolean                            ' enable firing animation
Dim m_Frame As Byte                                 ' animation frame
Dim m_Visible As Boolean                            ' Deceased or not

' enemy containers
Dim m_Target As enumBunkerTarget                    ' what will the bunker fire at?
Dim m_TR As Integer                                 ' maximum target range
'Dim w
Dim m_EnemyIndex As Byte
Dim m_bEnemyEngaged As Boolean


Public Sub CreateBunker(x As Integer, y As Integer, _
                        eBP As enumBunkerPosition, _
                        eBT As enumBunkerType)
 m_x = x
 m_y = y
 m_BP = eBP
 m_BT = eBT
 m_Frame = 0
 m_hitpoints = eBT * 10
 
 ' set target range
 If m_BP = BP_CLOSE Then
  m_cx = 180
  m_cy = 140
  m_TR = 350
 ElseIf m_BP = BP_FAR Then
  m_cx = 100
  m_cy = 60
  m_TR = 350
 ElseIf m_BP = BP_VERYFAR Then
  m_cx = 60
  m_cy = 40
  m_TR = 350
 End If
 
 m_Visible = True
End Sub

Public Sub GetTarget()                                  ' choose a random target in it's sight range
 ' Desc: Get an enemy target for the bunker
 If m_bEnemyEngaged Or m_Visible = False Then Exit Sub  ' exit if an enemy's been already engaged
 Dim cn As Long                                         ' local counter
 
 ' Check for a meteor in the sight-range
 If (Not m_bEnemyEngaged) Then
  For cn = 0 To MAX_METEORS
   If g_Meteor(cn).Visible And _
     (nGetDist2D(CInt(g_Meteor(cn).x), CInt(g_Meteor(cn).y), _
                 GetX, GetY) < m_TR) Then
      
      m_bEnemyEngaged = True
      m_Target = BTARGET_METEOR
      m_EnemyIndex = CByte(cn)
      SetFire = True
   End If
  Next
 End If
 
 ' check for an enemy spaceship
 For cn = 0 To MAX_ENEMIES
  If CShip(cn).GetVisible Then
   If fGetDist2D(CShip(cn).GetX, CShip(cn).GetY, _
                 CSng(GetX), CSng(GetY)) < m_TR Then
   
     m_bEnemyEngaged = True                              ' set "busy attacking" flag
     m_Target = BTARGET_SHIP
     m_EnemyIndex = CByte(cn)
     SetFire = True
   End If
  End If
 Next
 
 ' Check for an enemy missiles
 If (Not m_bEnemyEngaged) Then
  For cn = 0 To MAX_MISSILES
   If g_Missile(cn).Visible And _
     (nGetDist2D(CInt(g_Missile(cn).x), CInt(g_Missile(cn).y), _
                 GetX, GetY) < m_TR) Then
   
      m_bEnemyEngaged = True
      m_Target = BTARGET_MISSILE
      m_EnemyIndex = CByte(cn)
      SetFire = True
   End If
  Next
 End If
    
 ' rotate bunker at the enemy's face
  'If cShip(cn).GetX > GetX Then
  ' SetRotation = SO_RIGHT
  'Else
  ' SetRotation = SO_LEFT
  'End If
    
  'm_bEnemyEngaged = True                              ' set "busy attacking" flag
  'm_EnemyIndex = CByte(cn)
  'm_Target = BTARGET_SHIP                             ' identify taget ( Sounds familiar a? ;)
  
  'Debug.Print "Targeted ship n:" & cn
End Sub

Public Sub FireWeapon()
 ' Desc: Update bunker fire, do range checkings and position tracking
 If (Not m_Visible) Then Exit Sub
 
 Dim fEnemyX As Single
 Dim fEnemyY As Single
 Dim bytEnemyZ As Byte
 Dim benemyVisible As Boolean
 Dim eExplosion As enumExplosionType
 
 ' fill target info
 If m_bEnemyEngaged Then                                     ' if an enemy's been engaged
  
  Select Case m_Target
   Case BTARGET_SHIP
     fEnemyX = CShip(m_EnemyIndex).GetX
     fEnemyY = CShip(m_EnemyIndex).GetY
     bytEnemyZ = CShip(m_EnemyIndex).GetZ
     benemyVisible = CShip(m_EnemyIndex).GetVisible
     eExplosion = (ET_SMALLBLUE + 3 * (bytEnemyZ - 1))
   
   Case BTARGET_METEOR
     fEnemyX = g_Meteor(m_EnemyIndex).x
     fEnemyY = g_Meteor(m_EnemyIndex).y
     bytEnemyZ = g_Meteor(m_EnemyIndex).z
     benemyVisible = g_Meteor(m_EnemyIndex).Visible
     eExplosion = (ET_SMALLBLUE + 3 * (bytEnemyZ - 1))
       
   Case BTARGET_MISSILE
     fEnemyX = g_Missile(m_EnemyIndex).x
     fEnemyY = g_Missile(m_EnemyIndex).y
     bytEnemyZ = g_Missile(m_EnemyIndex).z
     benemyVisible = g_Missile(m_EnemyIndex).Visible
     eExplosion = (ET_SMALLBLUE + 3 * (bytEnemyZ - 1))
   
   Case Else: '...
  End Select
  
    
 ' if target's already been destroyed or got out of range then release bunker and exit procedure
 If (Not benemyVisible Or _
     nGetDist2D(CInt(fEnemyX), CInt(fEnemyY), GetX, GetY) > m_TR) Then
     m_bEnemyEngaged = False
  Exit Sub
 End If
  
  If (nGetRnd(0, 2000) < 100) Then                           ' determine if it should fire
    
    If (fEnemyX > GetX) Then                                 ' set appropriate rotation relevant to ship x-position
     SetRotation = SO_RIGHT
    Else
     SetRotation = SO_LEFT
    End If
   
   ' take damage from encounterd target
   If (m_Target = BTARGET_SHIP) Then
    CShip(m_EnemyIndex).DoDamage = 1
   
   ElseIf (m_Target = BTARGET_MISSILE) Then
    'g_Missile(m_EnemyIndex).HP = g_Missile(m_EnemyIndex).HP - 1
    g_Missile(m_EnemyIndex).Visible = False
   
   ElseIf (m_Target = BTARGET_METEOR) Then
    g_Meteor(m_EnemyIndex).HP = g_Meteor(m_EnemyIndex).HP - 2 '{!}
   End If
   
   Call CreateExplosion(CInt(fEnemyX + nGetRnd(-20, 20)), _
                        CInt(fEnemyY + nGetRnd(-2, 2)), _
                        bytEnemyZ, eExplosion)
   SetFire = True                                             ' enable fire animation
  Else
   SetFire = False
  End If
 
 End If

End Sub

Public Sub Render()
 Dim rBunker As RECT
 Static lRTime As Long                                  ' rotation time
 
 If (m_Visible) Then
  ' update rotation animation frame
  Call Rotate(m_Rotation)
  Call SetRect(rBunker, 0, 0, m_cx, m_cy)
  Call BltFastGFX_HBM(m_x - wx, m_y - wy, g_Objects.Bunker1(m_BP, m_Frame))
 Else
  ' update destroyed anim.
  Call SetRect(rBunker, 0, 0, m_cx, m_cy)
  Call BltFastGFX_HBM(m_x - wx, m_y - wy, g_Objects.Bunker1Dead(m_BP))
 End If
 
End Sub

Private Sub Rotate(eDir As enumDirection)
 Static lFPSTime As Long
 Dim bUpdateFrame As Boolean
 
 If lFPSTime < GetTicks Then
  lFPSTime = GetTicks + FPS_ANIMS
  bUpdateFrame = True
 End If
 
 If bUpdateFrame Then
  ' reset fire
 
  If m_Frame = 5 And m_bFiring = True Then
   m_Frame = 2
   m_bFiring = False
  End If
  If m_Frame = 3 And m_bFiring = True Then
   m_Frame = 0
   m_bFiring = False
  End If
  
  ' do rotating
  Select Case eDir                                     ' determine rotation
    Case SO_LEFT ' 3-5 ' firing
     If m_bFiring Then                                 ' check for fire flag
      If m_Frame > 3 Then m_Frame = m_Frame - 1
     Else
      If m_Frame > 0 Then m_Frame = m_Frame - 1
     End If
     
    Case SO_RIGHT
     If m_bFiring Then
      If m_Frame < 5 Then m_Frame = m_Frame + 1
     Else
      If m_Frame < 2 Then m_Frame = m_Frame + 1
     End If
  Case Else
  End Select
 
 End If
 
 
End Sub

' Info setting procs.
'-------------------------------------------------------------------

Public Property Let DoDamage(nDamage As Integer)
 ' Desc: Take hitpoints
 If Not m_Visible Then Exit Property                        ' if bunker's already dead then exit
 
 Dim nHP As Integer
 
 nHP = CInt(m_hitpoints)
 nHP = nHP - nDamage
 ' bunker's about to die -> do app. settings
 If nHP <= 0 Then
  SetVisible = False
  m_hitpoints = 0
  Call CreateExplosion(GetX, GetY, 1, ET_BIG)
  
  ' play destruction sound
  If (m_BP = BP_CLOSE) Then Call DSPlaySound(g_dsSfx(SFX_BUNKEREXPLODE), False, (GetX() - wx), SFX_VOLUMENORMAL)
  If (m_BP = BP_FAR) Then Call DSPlaySound(g_dsSfx(SFX_BUNKEREXPLODE), False, (GetX() - wx), SFX_VOLUMECLOSE)
  If (m_BP = BP_VERYFAR) Then Call DSPlaySound(g_dsSfx(SFX_BUNKEREXPLODE), False, (GetX() - wx), SFX_VOLUMEFAR)
 Else
  m_hitpoints = nHP
 End If
 
End Property

Public Property Let SetRotation(eDir As enumDirection)
 m_Rotation = eDir
End Property

Public Property Let SetFire(bSF As Boolean)
 If (m_bFiring = True And bSF = True) Or _
    (m_bFiring = False And bSF = False) Then Exit Property
 m_bFiring = bSF
 
 ' adjust animation frame
 If m_bFiring = True Then
  m_Frame = (m_Frame + 3) ' Mod 5) + 3 ' ' move right 3 frames
  If m_Frame > 5 Then m_Frame = 5
 Else
  If m_Frame >= 3 Then
   m_Frame = (m_Frame - 3)               ' move left 3 frames to disable weapons fire
  End If
 End If
 
 ' play fire sound
 'getx = 2080 ' right
 'wx = 0
 'fl = (GetX - wx)  '/ 2100
 'Call DSPlaySound(g_dsCannon(m_BP + 2), False, (GetX - wx))
 If (m_BP = BP_CLOSE) Then Call DSPlaySound(g_dsCannon(SFX_CLOSEBUNKER), False, (GetX - wx), SFX_VOLUMECLOSE)
 If (m_BP = BP_FAR) Then Call DSPlaySound(g_dsCannon(SFX_FARBUNKER), False, (GetX - wx), SFX_VOLUMEFAR)
 If (m_BP = BP_VERYFAR) Then Call DSPlaySound(g_dsCannon(SFX_VERYFARBUNKER), False, (GetX - wx), SFX_VOLUMEVERYFAR)
 
End Property

Public Property Let SetVisible(bVisible As Boolean)
 m_hitpoints = m_BT * 10    '{!}
 m_Visible = bVisible
End Property

' Info retrieving procs.
'-------------------------------------------------------------------

Public Property Get GetPos() As enumBunkerPosition
 GetPos = m_BP
End Property


Public Property Get GetX() As Integer
 GetX = m_x + m_cx / 2
End Property

Public Property Get GetY() As Integer
 GetY = m_y + m_cy / 2
End Property

Public Property Get GetZ() As enumBunkerPosition
 GetZ = m_BP
End Property

Public Property Get GetVisible() As Boolean
 GetVisible = m_Visible
End Property
