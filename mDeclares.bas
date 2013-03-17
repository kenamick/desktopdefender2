Attribute VB_Name = "mDeclares"
Option Explicit

' global variables
' -------------------------------------------------------------

Public bRunning          As Boolean                 ' is app. running
Public bWindowed         As Boolean                 ' is app.windowed
'Public bMainMenu         As Boolean                 ' is game. in main menu state
Public wx                As Integer                 ' world x position
Public wy                As Integer                 ' world y position
Public w_ScrollRate      As Integer
Public g_xEarth  As Integer, g_yEarth  As Integer   ' global earth position
Public g_cxEarth As Integer, g_cyEarth As Integer   ' global earth size
Public g_hpEarth         As Integer
Public xBg1 As Single, yBg1 As Single               ' scrolling background coords.
Public xBg2 As Single, yBg2 As Single
Public xBg3 As Single, yBg3 As Single
Public cpx  As Integer, cpy As Integer              ' global cockpit position
Public PlayerWeapon      As enumPlayerWeapon        ' current player weapon
Public g_PlayerDmgBonus  As Byte
Public g_PlayerFireDelay As Long
Public nFPS              As Integer                 ' game frame-rate
Public g_bDebug          As Boolean                 ' debuggin flag
Public g_bNotRetrace     As Boolean                 ' shoud I retrace ? ( I dunno..., should you? :)
Public g_nMusicVol       As Long                    ' global music volume
Public g_Player          As stPlayer
Public g_nPlayerIndex    As Integer                 ' player-character-array-index
Public g_Gates           As cnstGatesState          ' gates' state
Public g_smq()           As stSMQ                   ' scrolling-message-queue
Public g_numsmq          As Integer                 ' number of SMQs
Public g_GameState       As cnstGameState           ' current game state
Public g_bmusPlaying     As Boolean                 ' is background music playing
Public g_Language        As cnstLanguage            ' game language mode
Public g_showFPS         As Boolean                 ' show FPS
Public g_loadprogress    As Integer                 ' loading progressbar width
Public g_bsolidbars      As Boolean                 ' transluent health bars

Public bools(6)          As Boolean '{!}
 
' sfx
Public g_dsCannon(4)       As stSoundBuffer         ' player-bunker fire
Public g_dsbexplode        As stSoundBuffer
Public g_dsSfx(SFX_SOUNDS) As stSoundBuffer
'Public g_dsShip(5)   As stSoundBuffer

 
' global arrays
' -------------------------------------------------------------
Public arMS_Offsets(4)        As Integer
Public arDamages(WEAPONS)     As Integer            ' table of damages dependent on weapon type
Public arAttackRange(WEAPONS) As Integer            ' table of attack ranges dependant on weapon type
Public arVelocity(SHIPS)      As Single             ' table of velocities dependant on ship type
'Public arClr_ParticleBeam(9)  As Long               ' Particle Beam Color table
'Public arClr_LaserBeam(9)     As Long               ' Laser Beam - || - || - || -
'Public arClr_GreenPus(9)      As Long               ' GreenPus Beam - || - || - || -
'Public arBunkPos(19) As typeStar
Public arCreditsList(8)       As String             ' array of strings containing authors info
Public arText(TEXTS_NUM)      As String           'String             ' all menu texts
Public g_arEpilogue()         As String
Public g_arIntro()            As String

' global structures
' -------------------------------------------------------------
Public rScreen                   As RECT            ' screen rectangle
Public rEmpty                    As RECT
Public g_Objects                 As typeGr_Objects  ' all the gfx
Public Star(MAX_STARS)           As typeStar        ' some stars
Public g_WarpGate(MAX_WARPGATES) As typeWarpGate
Public g_PExp(MAX_EXPLOSIONS)    As typePExplosion  ' particle explosion structures
Public g_CPixel(MAX_CPIXELS)     As typeCPixel      ' chilling pixels structures
Public g_Missile(MAX_MISSILES)   As typeWeapon
Public g_LaserCut(MAX_LASERS)    As typeWeapon
Public g_PLaser(MAX_LASERS)      As typeWeapon      ' particle_beam_laser
Public g_StarTrip(MAX_STARS)     As typeParticle
Public g_Credits(2)              As typeCreditList  ' credits' list array structure
Public g_Explosion(MAX_EXPLOSIONS) As typeExplosion ' explosions array
Public g_Meteor(MAX_METEORS)     As typeMeteor      ' meteor threat
'Public g_Pl_Missile(MAX_MISSILES) As typePL_Weapon ' player-missiles
'Public g_Pl_Laser(MAX_LASERS) As typePL_Weapon
Public g_Pl_Weapon(MAX_SHOTS)    As typePL_Weapon
Public g_Bonus(BONUS_MAX)        As stBonus         ' bonus object uses chilling pixels' struct
Public g_PRemover(PREMOVER_MAX)  As typeParticle    ' particle remover - bonus object
Public g_MenuPos(12)             As POINTAPI

' global classes
' -------------------------------------------------------------
Public CMouse         As clsMouse
'Public cAI       As clsAI                          ' the AI class
Public CShip()        As clsShip                    ' Enemy Ship class
Public CBunker()      As clsBunker                  ' Bunker class
'Public CDXErr         As New clsDXErrors            ' DirectX errors class ( not that full but...)
Public CLevel         As New clsLevel               ' global mission class
Public CKdfGfx        As New clsKDF2                ' graphics packet
Public CKdfSfx        As New clsKDF2                ' sounds packet
Public CShuttle       As New clsShuttle
Public CBattleStation As New clsBattleStation

' Structures
' -------------------------------------------------------------

Public Type typeGr_Objects
  ' main menu and cursor stuff
  vbc            As typeGFX_HBM
  Cursor(1)      As typeGFX_HBM
  CockPit        As typeGFX_HBM
  buton          As typeGFX_HBM
  butoff         As typeGFX_HBM
  Title(20)      As typeGFX_HBM
  gate(1)        As typeGFX_HBM
  caline         As typeGFX_HBM
  seline         As typeGFX_HBM
  dialog         As typeGFX_HBM
  errdialog      As typeGFX_HBM
  backpaper(1)   As typeGFX_HBM                     ' briefing backpapers
  credits        As typeGFX_HBM
  ' moon surface(s) and static objects
  MoonSurf(4)    As typeGFX_HBM
  BackMoon(1)    As typeGFX_HBM
  'BackGround(2) As typeGFX_HBM                     ' background scrolling images
  Bonus(5, 6)    As typeGFX_HBM
  Earth          As typeGFX_HBM
  earthhp        As typeGFX_HBM                     ' earthhitpoints
  bshp           As typeGFX_HBM                     ' battlestation hitpoints
  Sun            As typeGFX_HBM
  Meteor1(21)    As typeGFX_HBM
  Meteor1_Shadow(21) As typeGFX_HBM
  Meteor2(21)    As typeGFX_HBM
  ' bunkers
  Bunker1(2, 5)  As typeGFX_HBM
  Bunker1Dead(2) As typeGFX_HBM
  ' battlestation
  BattleStation(14) As typeGFX_HBM
  bs_missile        As typeGFX_HBM
  ' ships & weapons
  Ship1(11)      As typeGFX_HBM                     ' SHIP_INTERCEPTOR
  Ship2(29)      As typeGFX_HBM                     ' SHIP_CARRIER
  Ship3(1)       As typeGFX_HBM                     ' SHIP_MISSILE_CRUISER
  Ship4(15)      As typeGFX_HBM                     ' SHIP_LASER_CRUISER
  Ship5(11)      As typeGFX_HBM                     ' SHIP_BEAM_CARRIER
  bclaz(7)       As typeGFX_HBM                     ' particle_lazer
  GM(16)         As typeGFX_HBM                     ' pl_missile
  LS(16)         As typeGFX_HBM                     ' pl_laser
  pr             As typeGFX_HBM                     ' particle remover
  RocketL(0)     As typeGFX_HBM                     ' enemy_missiles
  RocketR(0)     As typeGFX_HBM
  RocketL_Close(0)  As typeGFX_HBM
  RocketR_Close(0)  As typeGFX_HBM
  RocketL_VClose(0) As typeGFX_HBM
  RocketR_VClose(0) As typeGFX_HBM
  RedLaser(1)       As typeGFX_HBM                  ' enemy_lasers
  GreenLaser(1)     As typeGFX_HBM
  'Beam As typeGFX_HBM
  CannonLeft        As typeGFX_HBM
  CannonRight       As typeGFX_HBM
  ' Gfx ( warpgates, explosions ...
  'WG_Back           As DirectDrawSurface7
  'WarpGate(10)      As typeGFX_HBM
  WarpGate_Far(10)  As typeGFX_HBM
  Exp1(20)          As typeGFX_HBM
  Exp2(9)           As typeGFX_HBM
  Exp3(9)           As typeGFX_HBM
  Exp1Far(20)       As typeGFX_HBM
  Exp2Far(9)        As typeGFX_HBM
  Exp3Far(9)        As typeGFX_HBM
  'Exp2DC(9) As Long
  es1               As typeGFX_HBM
  es2               As typeGFX_HBM
  es3               As typeGFX_HBM
  Star1(0)          As typeGFX_HBM
  FireCles          As typeGFX_HBM                   ' fire-particles surface
  BlueCles          As typeGFX_HBM                   ' blue-particles
  ChillCles         As typeGFX_HBM                   ' chilling-particles
  starcles          As typeGFX_HBM                   ' starbitmap
End Type

Public Type typeCreditList                           ' author structure
  nReserved As Integer
  x As Integer                                       ' font position
  y As Integer
  xVel As Byte
  yVel As Byte
  ang As Integer                                     ' rotation angle
  fs As Integer                                      ' font size
  Heading As Byte
  lpszAuthor As String
  cr As Byte                                         ' saturation values
  Visible As Boolean
End Type

'Public Type stDialogBox
'  x         As Integer
'  y         As Integer
'  bCentered As Boolean
'  strTitle  As String
'  strInput  As String
'  Visible   As Boolean
'End Type

Public Type stSMQ
  strMsg       As String
  lduration    As Long
  lTimeCounter As Long
End Type

Public Type stBonus                                  ' bonus structure
  x      As Single
  y      As Single
  floata As Single
  floatb As Single
  kind   As cnstBonuses
  fxTime As Long
  frame  As Byte
  state  As cnstBonusState
End Type

Public Type typeCPixel
  x As Single
  y As Single
  z As Byte
  'cc As Byte                                         ' chilling color
  xVel As Single
  yVel As Single
  frame As Byte                                      ' animation frame
  Life As Byte
  Visible As Boolean
End Type

Public Type typeParticle
  x As Single
  y As Single
  xVel As Single
  yVel As Single
  yFriction As Single
  r As Byte
  G As Byte
  B As Byte
  cf As Integer
  Heading As Byte
  frame As Byte                                   ' animation frame
  Life As Integer
  Visible As Boolean
End Type

Public Type typePExplosion                        ' particle explosion UDT
  Visible As Boolean
  Particle(MAX_PARTICLES) As typeParticle
End Type

Public Type typePL_Weapon                         ' player-weapon UDT
  x As Single                                     ' coords.
  y As Single
  z As Byte
  xVel As Single
  yVel As Single
  '---- target info
  dx As Integer                                   ' destination position
  dy As Integer
  enemyIndex As Byte                              ' enemy index, needed for guided missiles
  TargetKind As enumBunkerTarget
  dDist As enumBunkerPosition
  '---- additional info
  Guided As Boolean                               ' automatically tracks the enemy
  kind As enumPlayerWeapon
  frame As Byte                                   ' anim. frame
  Visible As Boolean
End Type

Public Type typeWeapon                            ' missile object
  x As Single                                     ' position
  y As Single
  z As Byte
  nTempVar As Integer
  xVel As Single                                  ' velocity
  yVel As Single
  xVelB As Single
  yVelB As Single
  Possession As Byte                              ' who's this missle belong to
  '---- target info
  dx As Integer                                   ' destination position
  dy As Integer
  dcx As Integer                                  ' destination dimensions
  dcy As Integer
  Direction As enumDirection                      ' blitting directions
  '---- additional info
  kind As enumWeapon
  Distance As enumBunkerPosition
  ObjectiveDist As enumBunkerPosition
  frame As Byte                                   ' anim. frame
  Visible As Boolean                              ' in use, or not...
End Type

Public Type typeStar
  x     As Single
  y     As Single
  z     As Single                                     ' star z-distance
  frame As Byte
End Type

Public Type typeWarpGate
  x As Integer
  y As Integer
  z As Byte
  frame As Byte                                   ' animation frames
  bPlayBack As Boolean                            ' play animation backwards
  Visible As Boolean
End Type

Public Type typeExplosion                         ' explosion object
  x As Integer
  y As Integer
  z As Byte
  frame As Byte
  kind As enumExplosionType
  Visible As Boolean
End Type

Public Type typeMeteor                            ' meteor structure
  x       As Single
  y       As Single
  z       As Byte
  cx      As Integer
  cy      As Integer
  xVel    As Single
  yVel    As Single
  sy      As Single                               ' shadow - y
  HP      As Integer
  Exp     As Byte
  Data    As enumMeteorConsts                     ' additional data
  putshad As Boolean
  frame   As Byte
  Visible As Boolean
  ' target coords.
  dx As Integer
  dy As Integer
End Type

