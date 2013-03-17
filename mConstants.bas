Attribute VB_Name = "mConstants"
Option Explicit

' global constants
' -------------------------------------------

' gameplay stuff
'Public Const VIEWS = 5
Public Const SCREEN_PIXEL_WIDTH = (MAX_CX * 3) + 160
Public Const SCREEN_PIXEL_HEIGHT = MAX_CY
'Public Const VISIBLE_AREA_CX = 350
Public Const VISIBLE_AREA_CY = 285
Public Const VISIBLE_AREA_CY_2 = 285 \ 2
Public Const TARGET_AREA = 285 + 70
Public Const MAX_STARS As Long = 169
Public Const MAX_BUNKERS As Long = 5 '14
Public Const MAX_WARPGATES As Long = 20
Public Const MAX_PARTICLES As Long = 29
Public Const MAX_CPIXELS As Long = 24
Public Const MAX_EXPLOSIONS As Long = 29
Public Const MAX_ENEMIES As Long = 49
Public Const MAX_INTERCEPTORS = 7 '(7+0=8)
Public Const MAX_MISSILES As Long = 19
Public Const MAX_LASERS As Long = 20
Public Const MAX_SHOTS As Long = 30
Public Const MAX_METEORS As Long = 19
Public Const EARTH_HITPOINTS = 250
Public Const NO_POSSESSION = 255
Public Const SHIPS = 5
Public Const WEAPONS = 5
Public Const FPS_ANIMS = 1000 / 15
Public Const FPS_AI = 15
Public Const KVID As Long = &H4449564B

' detail constants
Public Const PLANE_CLOSE = 1                          ' planes to put the objects into
Public Const PLANE_FAR = 2
Public Const PLAYER_MISSILE_SPEED = 1.9
Public Const PLAYER_LASER_SPEED = 8

' experience constants
'Public Const EXPERIENCE_METEORFAR = 30
'Public Const EXPERIENCE_METEORCLOSE = 10

' bonus constants
Public Const BONUS_MAX = 6
Public Const BONUS_DURMEGADAMAGE = 15000                ' megadamage duration (in ms.)
Public Const BONUS_DURRAPIDFIRE = 12000                 ' rapid-fire duration (in ms.)
Public Const PREMOVER_MAX = 9                           ' max objects
Public Const PREMOVER_MAXTARGETS = 3                    ' max ships to hit
Public Const PREMOVER_DAMAGE = 25

' text constants
Public Const TEXTS_NUM = 58
Public Const MENU_START = 0
Public Const MENU_OPTIONS = 1
Public Const MENU_HALLOFFAME = 2
Public Const MENU_HELP = 3
Public Const MENU_EXIT = 4
Public Const MENU_BACK = 5
Public Const MENU_GAMMA = 6
Public Const MENU_GAMMAUNAVAILABLE = 7
Public Const MENU_SOUND = 8
Public Const MENU_MUSIC = 9
Public Const MENU_VSYNCON = 10
Public Const MENU_VSYNCOFF = 11
Public Const MENU_CREDITS = 12
Public Const MENU_ADDNEWPLAYER = 13
Public Const MENU_PLAYERNAMEQUERY = 14
Public Const MENU_PLAYERNAMEERROR = 15
Public Const MENU_PLAYERDELETE = 16
Public Const MENU_PLAYERDELETEERROR = 17
Public Const MENU_PASSWORDERROR = 18
Public Const MENU_PASSWORDENTER = 19
Public Const MENU_PASSWORDQUERY = 20
Public Const MENU_PASSWORDCONFIRM = 21
Public Const MENU_PLAYERNAME = 22               ' --- listings
Public Const MENU_PLAYERPASSWORD = 23
Public Const MENU_PLAYERSCORE = 24
Public Const MENU_PLAYERLEVEL = 25
Public Const MENU_PLAYERMISSION = 26
Public Const MENU_PLAYERTOTALSHOTS = 27
Public Const MENU_PLAYERKILLS = 28
Public Const MENU_PLAYEREXPERIENCE = 29
Public Const MENU_PLAYERSUCCESS = 30
Public Const MENU_INFOBOXCREDITS = 31
Public Const MENU_INFOBOXHELP = 32
Public Const MENU_INFOBOXHOF = 33
Public Const MENU_INFOBOXPLAY = 34
Public Const MENU_HELP1 = 35
Public Const MENU_HELP2 = 36
Public Const MENU_HELP3 = 37
Public Const MENU_HELP4 = 38
Public Const MENU_HELP5 = 39
Public Const MENU_HELP6 = 40
Public Const MENU_HELP7 = 41
Public Const MENU_HELP8 = 42
Public Const MENU_HELP9 = 43
Public Const MENU_HELP10 = 44
Public Const MENU_HELP11 = 45
Public Const MENU_HELP12 = 46
Public Const MENU_PLAYERHELP1 = 47
Public Const MENU_PLAYERHELP2 = 48
Public Const MENU_PLAYERHELP3 = 49
Public Const MENU_COCKPIT_EARTH = 50
Public Const MENU_COCKPIT_TIMELEFT = 51
Public Const MENU_COCKPIT_WEAPON = 52
Public Const MENU_COCKPIT_WEAPON_LASER = 53
Public Const MENU_COCKPIT_WEAPON_MISFAR = 54
Public Const MENU_COCKPIT_WEAPON_MISCLOSE = 55
Public Const MENU_INTRO = 56
Public Const MENU_SOLIDBARSON = 57
Public Const MENU_SOLIDBARSOFF = 58

' SFX constants
Public Const SFX_VOLUMEMAX = DSBVOLUME_MAX
Public Const SFX_VOLUMEMIN = DSBVOLUME_MIN
Public Const SFX_VOLUMENORMAL = 0
Public Const SFX_VOLUMECLOSE = -200
Public Const SFX_VOLUMEFAR = -500
Public Const SFX_VOLUMEVERYFAR = -1000
Public Const SFX_SOUNDS = 20                    ' num of sounds

 ' menu sounds
Public Const SFX_MENUCHOICE = 0
Public Const SFX_MENUSELECT = 1
Public Const SFX_MENUCALIBRATE = 2
Public Const SFX_MENUEXIT = 3
Public Const SFX_SPACEMYST1 = 4
Public Const SFX_OPENGATE = 5
Public Const SFX_CLOSEGATE = 6
Public Const SFX_BUNKEREXPLODE = 7
Public Const SFX_COCKPITSMQ = 8
Public Const SFX_WARPGATE = 9
Public Const SFX_PLAYERROCKETFIRE = 10
Public Const SFX_GREENLASER1 = 11
Public Const SFX_GREENLASER2 = 12
Public Const SFX_NORMALASER = 13
Public Const SFX_INTERCEPT1 = 14
Public Const SFX_INTERCEPT2 = 15
Public Const SFX_FAREXPLOSION1 = 16
Public Const SFX_BIGBLAST1 = 17
Public Const SFX_BIGBLAST2 = 18
Public Const SFX_METEORBLAST = 19
Public Const SFX_PARTICLEXPLOSION = 20

 ' bunker sounds
Public Const SFX_CANNONPLAYER1 = 0
Public Const SFX_CANNONPLAYER2 = 1
Public Const SFX_CLOSEBUNKER = 2
Public Const SFX_FARBUNKER = 3
Public Const SFX_VERYFARBUNKER = 4

' Music constants
Public Const MUSIC_VOLUMEMAX = 100
Public Const MUSIC_VOLUMEMIN = 0

' global enums
' -------------------------------------------

Public Enum cnstLanguage
 LANG_ENGLISH = 0
 LANG_BULGARIAN
End Enum

Public Enum cnstGameState
 GAMSTATE_LOGO = 0                                    ' logo playing
 GAMSTATE_MAINMENU = 1                                ' main menu mode
 GAMSTATE_BRIEFING                                    ' mission briefing
 GAMSTATE_PLAY                                        ' gamePlay
 GAMSTATE_EPILOGUE                                    ' gameEpilogue and ending animation
 GAMSTATE_QUIT                                        ' quiting state
 GAMSTATE_INTRO                                       ' intro story
End Enum

Public Enum cnstGatesState
  GS_NONE = 0                                         ' no state
  GS_OPEN                                             ' open gates
  GS_CLOSE                                            ' close gates
  GS_CLOSEOPEN                                        ' close and then open
End Enum

Public Enum cnstBonusState
  BSTATE_INACTIVE = 0
  BSTATE_FLOATING                                     ' bonus animation of floating around in space
  BSTATE_ACTIVE                                       ' bonus's currently active
End Enum

Public Enum cnstBonuses
  BONUS_ANNIHILATE = 0                                ' annihilate all visible enemies
  BONUS_RANGEKILL                                     ' kill all ships in 3x32 squares dist
  BONUS_PREMOVER                                      ' a little mech that does 10pts. of damage
  BONUS_RAPIDFIRE                                     ' shots delay decreases by 100ms
  BONUS_REVIVEBUNKER                                  ' resurrect destroyed bunker
  BONUS_MEGADAMAGE                                    ' 20 pts. of damage
  '...
End Enum

Public Enum enumEnemyObjective
  EO_EARTH = 0
  EO_MOONBUNKER = 50
  EO_PLAYERBUNKER = 100
  EO_BATTLESTATION = 150
End Enum

Public Enum enumPlayerWeapon
  PW_LASER = 0
  PW_MISSILE_CLOSERANGE
  PW_MISSILE_LONGRANGE
End Enum

Public Enum enumFontAnimationStates
  FA_INTRO = 0
  FA_MAINMENU
End Enum

Public Enum enumExplosionType
  ET_SMALL = 0
  ET_SMALLBLUE
  ET_BIG
'  ' distant
  ET_SMALL_FAR
  ET_SMALLBLUE_FAR
  ET_BIG_FAR
End Enum

Public Enum enumObjectDistance
  OD_FAR = 0
  OD_CLOSE
  OD_VERYCLOSE
End Enum

' meteor constants
Public Enum enumMeteorConsts
  ' distance
  MC_FAR = 1
  MC_CLOSE = 2
  MC_LEFT = &H4        ' 2 ^ 2
  MC_RIGHT = &H8       ' 2 ^ 3
  ' target
  MC_HITEARTH = &H10   ' 2 ^ 4
  MC_HITMOON = &H20    ' 2 ^ 5
  ' kind
  MC_TYPE1 = &H40      ' 2 ^ 6
    
End Enum


