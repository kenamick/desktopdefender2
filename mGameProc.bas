Attribute VB_Name = "mGameProc"
Option Explicit

' handles most of the inside-game-engine

' Blittings
' ---------------------------------------------------------------

Public Sub UpdateFrame()                        ' Updates All Animations and Graphics
 
 Select Case g_GameState
   
   ' play KenamicK's Logo
   Case GAMSTATE_LOGO
    Call mGFX.GFXClearBackBuffer
    Call mDirectDraw.DDBlitToPrim
    Call DoKKLogo
    
    
   Case GAMSTATE_INTRO
    Call UpdateStarTrip
    Call UpdateStory(True)
    
   ' --- main menu mode
   Case GAMSTATE_MAINMENU
    
    Call BltFastGFX_HBM(0, 0, g_Objects.backpaper(1))
    Call BltFastGFX_HBM(538, 434, g_Objects.vbc)
    Call UpdateStarTrip
    Call UpdateMainMenu
    Call UpdateLogo
    Call DoErrorBox(False)                        ' update error box messages
      
    If (g_showFPS) Then Call GFXTextOut(lpBack, 10, 470, "Running at: " & CStr(nFPS) & " FPS", 12, RGB(25, 25, 225))
   
   ' --- do briefing
   Case GAMSTATE_BRIEFING
   '...
    
    Call UpdateBriefing
    Call UpdateStarTrip
    
    If (g_showFPS) Then Call GFXTextOut(lpBack, 10, 470, "Running at: " & CStr(nFPS) & " FPS", 12, RGB(25, 25, 225))
      
   ' --- gamePlay mode
   Case GAMSTATE_PLAY
     
     Call GFXFadeInOut
     Call UpdateWorld                              ' update world position
     Call PlayerGetInput
     Call DoMoonQuake                              ' update Moon Quakes
     Call UpdateStarField
     ' update objects
     If Not bools(1) Then
     'Call UpdateBackGround                         ' update scrolling-background
     End If
     'Call UpdateStarField                          ' update persepective-star-field
     If Not bools(2) Then
     Call UpdateEarth                              ' update our homeworld ;)
     End If
     Call UpdateBonus                              ' draw all bonus objects
     Call UpdateParticleLasers                     ' update plaser
     Call UpdateFarEnemies                         ' update distant ships
     Call UpdateWarpGates
     '{!}
      If (bools(6)) Then Call CShuttle.Update
      Call CBattleStation.Update
     '{!}
     If Not bools(3) Then
     Call UpdateMoonSurface                        ' update close-range moon-surface
     End If
     Call UpdateMeteors                            ' update meteor threats
     Call UpdateExplosionsFar                      ' update distant explosions ONLY
     Call UpdateCloseEnemies                       ' update enemy ships classes ( closer )
     Call UpdateBunkers                            ' update friendly bunkers
     Call UpdateExplosions
     ' update weapons and gfx
     Call UpdateMissiles                           ' update fired-missiles
     Call UpdateLaserCuts
     Call UpdateParticleRemover
     Call UpdateChillingPixels                     ' update missile trails
     'Call UpdateExplosions                         ' update all explosions
     Call UpdateParticleExplosions
     ' handle all-player stuff
     If Not bools(4) Then
     Call UpdatePlayer
     End If
  
     ' update mission
     'g_Gates = GS_NONE  ' {!}
     If (g_Gates = GS_NONE) Then
      If (g_showFPS) Then Call GFXTextOut(lpBack, 610, 468, CStr(nFPS) & "fps", 12, RGB(15, 180, 15))
      
      If (Not CLevel.Update) Then
       ' increment mission counter
       g_Player.mission = g_Player.mission + 1
       Call StartBriefing
       ' copy scores
       g_arChar(g_nPlayerIndex) = g_Player
       ' precalc levels
       Dim cn As Long
       For cn = 0 To g_nChars
        g_arChar(cn).level = CHARExperienceToLevel(g_arChar(cn).Exp)
       Next
      End If
     End If
     
   
   ' --- Finished game
   Case GAMSTATE_EPILOGUE
     Call UpdateStarTrip
     Call UpdateStory(False)
     
   ' --- QuitGame mode
   Case GAMSTATE_QUIT
     '...
      Debug.Print "GAME LEVEL DONE"
      bRunning = False
   
   Case Else
    Debug.Print "Invalid Game State!"
   
 End Select
   
 ' gates need to be opened/closed here, before backbuffer's data is copied to the frontbuffer
 If (g_Gates = GS_OPEN) Then
  Call DoGates(False)
 ElseIf (g_Gates = GS_CLOSE Or g_Gates = GS_CLOSEOPEN) Then
  Call DoGates(True)
 End If
  
  
End Sub

'/////////////////////////////////////////////////////////////////
'//// Update player-input interaction
'/////////////////////////////////////////////////////////////////
Public Sub _
PlayerGetInput()
 
 
 Dim bPlayerFire As Boolean
 Dim cn          As Long
 
 ' get all mouse actions (left,right,pos)
 Call CMouse.GetActions
 Call mDirectInput.DICheckKeys
 
 ' change weapon
 If (CMouse.GetRight = MS_Up) Then
  PlayerWeapon = PlayerWeapon + 1
  If (PlayerWeapon > PW_MISSILE_LONGRANGE) Then PlayerWeapon = PW_LASER
 End If
  
 If CMouse.GetLeft = MS_Up Then
  bPlayerFire = True
 End If
 
 ' get keyboard interaction
 If (DIKeyState(DIK_ESCAPE) = KS_KEYUP) Then
  g_Gates = GS_CLOSEOPEN
  g_GameState = GAMSTATE_MAINMENU
  g_arChar(g_nPlayerIndex) = g_Player
  ' precalc levels
  For cn = 0 To g_nChars
   g_arChar(cn).level = CHARExperienceToLevel(g_arChar(cn).Exp)
  Next
 End If
 
 If (DIKeyState(DIK_LEFT) = KS_KEYDOWN Or DIKeyState(DIK_A)) Then UpdateWorld True
 If (DIKeyState(DIK_RIGHT) = KS_KEYDOWN Or DIKeyState(DIK_D)) Then UpdateWorld , True
 If (DIKeyState(DIK_1) = KS_KEYUP) Then PlayerWeapon = PW_LASER
 If (DIKeyState(DIK_2) = KS_KEYUP) Then PlayerWeapon = PW_MISSILE_CLOSERANGE
 If (DIKeyState(DIK_3) = KS_KEYUP) Then PlayerWeapon = PW_MISSILE_LONGRANGE
 If (DIKeyState(DIK_SPACE) = KS_KEYUP) Then bPlayerFire = True
 If (DIKeyState(DIK_F) = KS_KEYDOWN) Then
  If (g_showFPS) Then g_showFPS = False Else g_showFPS = True
 End If
 
 ' sound volume
 If (DIKeyState(DIK_MINUS) = KS_KEYDOWN) Then
  If (m_nGlobalVol > SFX_VOLUMEMIN) Then m_nGlobalVol = m_nGlobalVol - 5
 End If
 If (DIKeyState(DIK_EQUALS) = KS_KEYDOWN) Then
  If (m_nGlobalVol < SFX_VOLUMEMAX) Then m_nGlobalVol = m_nGlobalVol + 5
 End If
 
 ' play/stop music
 If (DIKeyState(DIK_M) = KS_KEYUP) Then
   If (g_bmusPlaying) Then Call mFMod.STOPmus Else _
    mFMod.PLAYmus True
 End If
 
 
 ' debug keys
 If (g_bDebug) Then
  If DIKeyState(DIK_F1) = KS_KEYUP Then bools(1) = True
  If DIKeyState(DIK_F2) = KS_KEYUP Then bools(2) = True
  If DIKeyState(DIK_F3) = KS_KEYUP Then bools(3) = True
  If DIKeyState(DIK_F4) = KS_KEYUP Then bools(4) = True
  If DIKeyState(DIK_F9) = KS_KEYUP Then bools(6) = True
  If DIKeyState(DIK_F5) = KS_KEYUP Then bools(1) = False
  If DIKeyState(DIK_F6) = KS_KEYUP Then bools(2) = False
  If DIKeyState(DIK_F7) = KS_KEYUP Then bools(3) = False
  If DIKeyState(DIK_F8) = KS_KEYUP Then bools(4) = False
  If DIKeyState(DIK_F10) = KS_KEYUP Then bools(6) = False
  If DIKeyState(DIK_F11) = KS_KEYUP Then CBunker(0).DoDamage = 100
  If DIKeyState(DIK_F12) = KS_KEYUP Then CBunker(3).DoDamage = 100
  If DIKeyState(DIK_R) = KS_KEYUP Then Call DDSetGamma(0, 0, 142)
  If DIKeyState(DIK_B) = KS_KEYUP Then CreateBonus (CMouse.GetX) + wx, (CMouse.GetY) + wy
  If (DIKeyState(DIK_H) = KS_KEYUP) Then Call GFXFadeInOut(2)
 End If
 
  
 ' Cheater keys !
 If (g_Player.cheater = 123) Then
  
  Dim objbns As stBonus
  
  If DIKeyState(DIK_NUMPAD7) = KS_KEYUP Then CreateBonus (CMouse.GetX) - wx, (CMouse.GetY) - wy, BONUS_PREMOVER
  If DIKeyState(DIK_NUMPAD8) = KS_KEYUP Then CreateBonus (CMouse.GetX) - wx, (CMouse.GetY) - wy, BONUS_RANGEKILL
  If DIKeyState(DIK_NUMPAD4) = KS_KEYUP Then
   objbns.kind = BONUS_REVIVEBUNKER
   Call GiveBonus(objbns)
  ElseIf DIKeyState(DIK_NUMPAD5) = KS_KEYUP Then
   objbns.kind = BONUS_RAPIDFIRE
   Call GiveBonus(objbns)
  ElseIf DIKeyState(DIK_NUMPAD6) Then
   objbns.kind = BONUS_MEGADAMAGE
   Call GiveBonus(objbns)
  ElseIf DIKeyState(DIK_NUMPAD1) Then
   objbns.kind = BONUS_ANNIHILATE
   Call GiveBonus(objbns)
  End If
 
 End If
  
 
 ' activate actions
 If (bPlayerFire And CMouse.GetY < TARGET_AREA) Then
  
  ' create missiles shot
  If (PlayerWeapon <> PW_LASER) Then
   Call Player_CreateMissile(CMouse.GetX, _
                             CMouse.GetY, 1, 1)
   
   g_Player.score = g_Player.score - 2  ' weaponary
   Call DSPlaySound(g_dsSfx(SFX_PLAYERROCKETFIRE), False)
   cpy = 6
  ' create laser shot
  Else
   
   Static lTimeLaser As Long
   If (lTimeLaser < GetTicks) Then
    lTimeLaser = 150 + GetTicks + g_PlayerFireDelay
   
    g_Player.score = g_Player.score - 1 ' weaponary
    Call Player_CreateLaser(CMouse.GetX, CMouse.GetY, 1, 1)
     ' do cockpit offset
    cpx = 4
    cpy = 10
   
    'DSPlaySound g_dsCannon(SFX_CANNONPLAYER), False
    DSPlaySound g_dsCannon(nGetRnd(SFX_CANNONPLAYER1, SFX_CANNONPLAYER2)), False
   End If
  End If
  
 End If
 

End Sub


Public Sub UpdatePlayer()
 ' Desc: Update player actions and stuff
 
 'Call Player_UpdateMissiles
 'Call Player_UpdateLasers
 Call Player_UpdateShots                        ' refresh all shots
 Call UpdateCockPit                             ' redraw cockpit
 Call UpdateSMQ                                 ' check and output scrolling-messages
 
 ' blit cursor
 If (CMouse.GetY > TARGET_AREA) Then
  Call BltFastGFX_HBM(CMouse.GetCenterX, CMouse.GetCenterY, g_Objects.Cursor(1))
 Else
  Call BltFastGFX_HBM(CMouse.GetCenterX, CMouse.GetCenterY, g_Objects.Cursor(0))
 End If
   
 ' draw info
  Dim strWeapon As String
  'Dim strAmmo   As String
  
  If (PlayerWeapon = PW_LASER) Then
   strWeapon = arText(MENU_COCKPIT_WEAPON_LASER) '"Laser"
  ' strAmmo = "Infinite"
  ElseIf (PlayerWeapon = PW_MISSILE_CLOSERANGE) Then
   strWeapon = arText(MENU_COCKPIT_WEAPON_MISFAR) '"Mis_Close"
  ' strAmmo = "10"
  ElseIf (PlayerWeapon = PW_MISSILE_LONGRANGE) Then
   strWeapon = arText(MENU_COCKPIT_WEAPON_MISCLOSE) '"Mis_Far"
  ' strAmmo = "10"
  End If
  
  Call DrawTextCP(0, arText(MENU_COCKPIT_WEAPON) & strWeapon)
  'Call DrawTextCP(0, "BS: " & CBattleStation.GetHitPoints)
  ' Call DrawTextCP(1, "Ammo: " & strAmmo)
  Call DrawTextCP(1, arText(MENU_COCKPIT_EARTH) & g_hpEarth)
  Call DrawTextCP(2, arText(MENU_PLAYERSCORE) & g_Player.score)
  Call DrawTextCP(3, arText(MENU_COCKPIT_TIMELEFT) & CLevel.GetTimeLeft())
  
  ' earth dies -> you die...
  If (g_hpEarth <= 0 Or (CBattleStation.GetHitPoints <= 0 And g_Player.mission = 4)) Then
   g_Gates = GS_CLOSEOPEN
   g_GameState = GAMSTATE_MAINMENU
   ' copy scores
   g_arChar(g_nPlayerIndex) = g_Player
   ' precalc levels
   Dim cn As Long
   For cn = 0 To g_nChars
   g_arChar(cn).level = CHARExperienceToLevel(g_arChar(cn).Exp)
   Next
  
  End If
   
End Sub


'///////////////////////////////////////////////////////////////
'//// Play Logo Animation
'///////////////////////////////////////////////////////////////
Public Sub _
DoKKLogo()

 On Local Error GoTo StartMenu:
 
 Dim lTimeLogo  As Long
 Dim xavi       As Integer
 Dim yavi       As Integer
 Dim sfilename  As String
 Dim ff         As Integer
 Dim header     As Long
  
 ' stop bgmusic
 'mFMod.STOPmus
   
 ' set movie path
 sfilename = App.Path & "\data\logo.avi"
 ff = FreeFile()
  
 Open (sfilename) For Binary Access Write Lock Read Write As #ff
  Put #ff, , &H46464952
 Close #ff
 
 Call mAvi.OpenAVI(frmMain.hwnd, sfilename, "KL")
 If (mAvi.GetAVIWindow("KL") = -1) Then
  'call errormsg("Could not locate logo file")
  AppendToLog ("Error loading logo movie.")
  GoTo StartMenu
 End If
 
 ' set timer
 lTimeLogo = 12900 + GetTicks()
 
 ' setup avi position
 xavi = (MAX_CX / 2) - 200 '(mAvi.GetAVIRect("KL").Right / 2)
 yavi = (MAX_CY / 2) - (mAvi.GetAVIRect("KL").Bottom / 2)
 Call mAvi.MoveAVIWindow("KL", xavi, yavi, 400)
 

 frmMain.BackColor = &H0
 'frmMain.AutoRedraw = True
 frmMain.Cls
 
 ' star movie
 Call mAvi.PlayAVI("KL", FromStart)
 
 Do While (lTimeLogo > GetTicks())
   
  Call DICheckKeys
  If (DIKeyState(DIK_ESCAPE) = KS_KEYUP Or _
      DIKeyState(DIK_SPACE) = KS_KEYUP Or _
      DIKeyState(DIK_RETURN) = KS_KEYUP) Then Exit Do
  
  'frmMain.Cls
  DoEvents
 Loop

 
StartMenu:
 
 ' set logo passed key in registry
 Call mUtil.RegSetKey(&H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "logo", "1")
 
 frmMain.AutoRedraw = False
 Call mAvi.CloseAVI("KL")
 Call mAvi.CloseAllAVI
 
 Open (sfilename) For Binary Access Write Lock Read Write As #ff
  Put #ff, , KVID
 Close #ff
 
 g_GameState = GAMSTATE_INTRO
 
 'AppendToLog ("Error playing logo!")
 ' {!} to remove
 'mDirectSound.m_nGlobalVol = -1000
 'DSPlaySound g_dsSfx(SFX_SPACEMYST1), True

End Sub


'///////////////////////////////////////////////////////////////
'//// Show Game Story - Intro/Epilogue
'///////////////////////////////////////////////////////////////
Public Sub _
UpdateStory(bIntro As Boolean)

 Static lTime       As Long
 'Static bmus_load   As Boolean
 Dim lClr           As Long
 Dim lClr1          As Long
 Dim x              As Integer
 Dim y              As Integer
 Dim msg            As String
 Dim cn             As Long
 
 
 ' get input
 Call DICheckKeys
 Call CMouse.GetActions
  

 If (lTime = 0 And bIntro) Then
  lTime = GetTicks() + 300000
 ElseIf (lTime = 0 And (Not bIntro)) Then
  lTime = GetTicks() + 60000
 End If
 
 ' on keypress
 If (DIKeyState(DIK_ESCAPE) = KS_KEYUP Or _
     DIKeyState(DIK_SPACE) = KS_KEYUP Or _
     DIKeyState(DIK_RETURN) = KS_KEYUP Or _
     CMouse.GetLeft = MS_Up Or _
     lTime < GetTicks()) Then
  
  lTime = 0
  g_GameState = GAMSTATE_MAINMENU
  
  
  ' start music after intro_story
  If (bIntro) Then
   Call mFMod.PLAYmus(True)
   g_bmusPlaying = True
   g_Gates = GS_NONE
  Else
   g_Gates = GS_CLOSEOPEN
  ' bmus_load = False
  ' mFMod.STOPmus
  ' mFMod.CLOSEmus
  ' ' load back main menu music
  ' If (Not mFMod.OPENmus(App.Path & "\data\ar.it")) Then
  '  Call ErrorMsg("Could not load music file!")
  ' Else
  '  AppendToLog ("FMod: loaded \ar.it")
  ' End If
   g_bmusPlaying = False
   
     
  End If
  
 Else
   lClr = RGB(0, 215, 15)
   lClr1 = RGB(75, 215, 15)
   Call BltFastGFX_HBM(0, 0, g_Objects.backpaper(0))
   
   If (Not bIntro) Then
   ' update epilogue
   
    'If (Not bmus_load) Then
   '
   '  mFMod.STOPmus
   '  mFMod.CLOSEmus
   '  If (Not mFMod.OPENmus(App.Path & "\data\wheni.it")) Then
   '   Call ErrorMsg("Could not load music file!")
   '  Else
   '   AppendToLog ("FMod: loaded \wheni.it")
   '  End If
   '
   '  mFMod.PLAYmus True
   '  bmus_load = True
   ' End If
    
   
    x = rScreen.Right / 2 - 4
    y = rScreen.Top + 50
    Call GFXTextOut(lpBack, x, y, g_arEpilogue(0), 24, lClr, , True)
    x = 50
    y = rScreen.Top + 90
    
    For cn = 1 To UBound(g_arEpilogue)
     Call GFXTextOut(lpBack, x, y + cn * 20, g_arEpilogue(cn), 20, lClr, , True)
    Next
   
   Else
    x = 5
    y = rScreen.Top + 65
    
    For cn = 0 To UBound(g_arIntro)
     Call GFXTextOut(lpBack, x, y + cn * 16, g_arIntro(cn), 18, lClr, , True)
    Next
   End If
   
 End If

  
End Sub


'///////////////////////////////////////////////////////////////
'//// Show mission briefing
'///////////////////////////////////////////////////////////////
Public Sub _
UpdateBriefing()

 Dim cn         As Integer
 Dim lvlDur     As Single
 Dim lClr       As Long
 Dim lClr1      As Long
 Dim strDur     As String
 
 ' get input
 Call DICheckKeys
 Call CMouse.GetActions
 ' user cancel
 If (DIKeyState(DIK_ESCAPE) = KS_KEYUP) Then
  g_GameState = GAMSTATE_MAINMENU
  g_Gates = GS_CLOSEOPEN
 End If
 
 If (CLevel.ElapsedBriefingTime() Or DIKeyState(DIK_SPACE) = KS_KEYUP Or CMouse.GetLeft = MS_Up) Then
  ' start game
  g_GameState = GAMSTATE_PLAY
  g_Gates = GS_CLOSEOPEN
   
  'bRunning = False
   
 ' update breifing
 Else
  ' draw briefing text
  lvlDur = (CLevel.m_lDuration / 60000)
  lClr = RGB(200, 15, 15)
  lClr1 = RGB(175, 15, 15)
  
  Call BltFastGFX_HBM(0, 0, g_Objects.backpaper(0))
  Call GFXTextOut(lpBack, 15, 50 + (cn * 20), CLevel.m_strName, 24, lClr, , True)
  
  
  strDur = Format$(lvlDur, "#.##")
  strDur = "Duration: " & strDur & " mins."
  Call GFXTextOut(lpBack, 15, 100, strDur, 20, lClr, , True)
  Call GFXTextOut(lpBack, 15, 130, "Briefing...", 20, lClr, , True)
  
  For cn = 0 To CLevel.GetBriefingLines
   Call GFXTextOut(lpBack, 15, 180 + (cn * 20), CLevel.GetBriefingLine(cn), 20, lClr1, , True)
  Next
  
 End If
 
End Sub


'///////////////////////////////////////////////////////////////
'//// Setup next game mission
'///////////////////////////////////////////////////////////////
Public Sub _
StartBriefing()

 Dim strMissionFile As String
  
 ' bound mission
 If (g_Player.mission < 1) Then g_Player.mission = 1
 If (g_Player.mission > 6) Then g_Player.mission = 6
 
 If (g_Language = LANG_BULGARIAN) Then strMissionFile = "cmp" _
  Else strMissionFile = "eng"
 
 Select Case (g_Player.mission)
 
   Case 1
    strMissionFile = strMissionFile & "1.ks"
   Case 2
    strMissionFile = strMissionFile & "2.ks"
   Case 3
    strMissionFile = strMissionFile & "3.ks"
   Case 4
    strMissionFile = strMissionFile & "4.ks"
   Case 5
    strMissionFile = strMissionFile & "5.ks"
    
   Case 6
    g_GameState = GAMSTATE_EPILOGUE
    g_Gates = GS_CLOSEOPEN
    g_Player.mission = 1
    g_arChar(g_nPlayerIndex).mission = g_Player.mission
    Exit Sub
    
   Case Else
    ' reset mission counter
    Debug.Print "Invalid mission!"
    g_Player.mission = 1
    g_arChar(g_nPlayerIndex).mission = g_Player.mission
    strMissionFile = "cmp1.ks"
    Exit Sub
    
 End Select
 
 
 ' load the mission script
 Call CLevel.Init
 If (Not CLevel.LoadScript(App.Path & "\script\" & strMissionFile)) Then
  AppendToLog ("Error loading level script! 'StartBreifing() Proc'")
  'Call mMain.ErrorMsg("Error loading level!")
  Call mMain.MakeError("Error loading mission! " & vbCr & "Please reinstall game!")
  Exit Sub
 End If

 g_GameState = GAMSTATE_BRIEFING
 
 ' close and open gates
 g_Gates = GS_CLOSEOPEN
 
 ' stop music
 If (g_bmusPlaying) Then
  g_bmusPlaying = False
  Call mFMod.STOPmus
 End If

 ' start space sound
 Call mDirectSound.DSPlaySound(g_dsSfx(SFX_SPACEMYST1), True)

 ' resetgame
 Call ResetGame

End Sub


'///////////////////////////////////////////////////////////////
'//// Refresh logo animation
'///////////////////////////////////////////////////////////////
Public Static Sub _
UpdateLogo()

 Dim lFPSTime     As Long
 Dim bytlogoframe As Byte
 Dim dx           As Long
 Dim dy           As Long
  
  If (lFPSTime < GetTicks) Then
     lFPSTime = GetTicks + (FPS_ANIMS + FPS_ANIMS)
      bytlogoframe = bytlogoframe + 1
     If (bytlogoframe > 20) Then bytlogoframe = 0
  End If
    
  dx = MAX_CX \ 2 - g_Objects.Title(bytlogoframe).cx \ 2
  dy = -20
  
   ' update menu
  Call BltFastGFX_HBM(dx, dy, g_Objects.Title(bytlogoframe))
  'Call BltFastGFX_HBM(0, 0, g_Objects.menu)

End Sub


'///////////////////////////////////////////////////////////////
'//// Refresh main menu gfx and actions
'///////////////////////////////////////////////////////////////
Public Sub _
UpdateMainMenu()

 Dim cn             As Integer
 Dim ltxtColor      As Long
 Dim lcnsttxtColor  As Long
 Dim ltxtselColor   As Long
 Dim ltxtbkColor    As Long
 Dim bytFontSize    As Byte
  
  ' animation vars.
 Dim bUpdateFrame    As Boolean
 Static lFPSTime     As Long
 Static bytlogoframe As Byte
 
 ' menu commands
 Static bytMenuState As Byte
 'Static nkeyX        As Integer                      ' selection position
 Static nkeySel      As Integer
 Static bSelected    As Boolean                      ' user pressed Enter or MouseKey
   
  ' get input only if dialog box's not on the screen
   Call CMouse.GetActions
   Call mDirectInput.DICheckKeys
  
  If (DIKeyState(DIK_F) = KS_KEYDOWN) Then
   If (g_showFPS) Then g_showFPS = False Else g_showFPS = True
  End If
  If (DIKeyState(DIK_ESCAPE) = KS_KEYUP) Then bytMenuState = 4 'bRunning = False
  If (DIKeyState(DIK_UP) = KS_KEYUP) Then
   ' move selection
   nkeySel = nkeySel - 1
   ' play sound
   Call DSPlaySound(g_dsSfx(SFX_MENUSELECT), False)
  End If
  If (DIKeyState(DIK_DOWN) = KS_KEYUP) Then
   ' move selection
   nkeySel = nkeySel + 1
   ' play sound
   Call DSPlaySound(g_dsSfx(SFX_MENUSELECT), False)
  End If
 
  ' do color var. calcs
  Static nrc As Integer
  Static nv  As Integer
  
  If (nrc >= 50) Then nv = -1
  If (nrc <= 0) Then nv = 1
  nrc = nrc + nv
    
  ' do precalcs
  ltxtColor = RGB(250 - (nrc * 4), 100 + (nrc * 2), 10) ' textcolor
  lcnsttxtColor = RGB(250, 50, 10)                   ' constant textcolor
  ltxtselColor = RGB(255, 255, 5)                    ' selection textcolor
  ltxtbkColor = RGB(50, 50, 225)                     ' selection bk text color
  bytFontSize = 28                                   ' letter font size
     
  'bytmenustate 0 - main menu, 1 - options, 2-hall of fame, 3-help , 4-exit, 5-play opts.
    
  
  ' start music
  If (Not g_bmusPlaying) Then
   ' stop space sound
   Call mDirectSound.DSStopSound(g_dsSfx(SFX_SPACEMYST1), 0)
   g_bmusPlaying = True
   Call mFMod.PLAYmus(True)
  End If
  
  
  ' check menu state
  Select Case bytMenuState
  
    ' main menu
    Case 0
      ' update right panel
      Call UpdateInfoBox(1)
      
      If (nkeySel > 5) Then nkeySel = 0
      If (nkeySel < 0) Then nkeySel = 5
      
      If (nkeySel = 0) Then  ' start new game
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_START), bytFontSize, ltxtselColor, ltxtbkColor)
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        bytMenuState = 5
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
       End If
      Else
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_START), bytFontSize, ltxtColor)
      End If
      If (nkeySel = 1) Then  ' options
       Call GFXTextOut(lpBack, g_MenuPos(1).x, g_MenuPos(1).y, arText(MENU_OPTIONS), bytFontSize, ltxtselColor, ltxtbkColor)
       ' selected
       If (DIKeyState(DIK_RETURN) = KS_KEYDOWN) Then
        bytMenuState = 1
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
       End If
      Else
       Call GFXTextOut(lpBack, g_MenuPos(1).x, g_MenuPos(1).y, arText(MENU_OPTIONS), bytFontSize, ltxtColor)
      End If
      If (nkeySel = 2) Then  ' hall of fame
       Call GFXTextOut(lpBack, g_MenuPos(2).x, g_MenuPos(2).y, arText(MENU_HALLOFFAME), bytFontSize, ltxtselColor, ltxtbkColor)
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
        bytMenuState = 2
       End If
      Else
       Call GFXTextOut(lpBack, g_MenuPos(2).x, g_MenuPos(2).y, arText(MENU_HALLOFFAME), bytFontSize, ltxtColor)
      End If
      If (nkeySel = 3) Then  ' help
       Call GFXTextOut(lpBack, g_MenuPos(3).x, g_MenuPos(3).y, arText(MENU_HELP), bytFontSize, ltxtselColor, ltxtbkColor)
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
        bytMenuState = 3
       End If
      Else
       Call GFXTextOut(lpBack, g_MenuPos(3).x, g_MenuPos(3).y, arText(MENU_HELP), bytFontSize, ltxtColor)
      End If
      If (nkeySel = 4) Then  ' replay intro
       Call GFXTextOut(lpBack, g_MenuPos(4).x, g_MenuPos(4).y, arText(MENU_INTRO), bytFontSize, ltxtselColor, ltxtbkColor)
       'If (bSelected) Then bytMenuState = 4
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        mFMod.STOPmus
        g_GameState = GAMSTATE_INTRO
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
       End If
      Else
       Call GFXTextOut(lpBack, g_MenuPos(4).x, g_MenuPos(4).y, arText(MENU_INTRO), bytFontSize, ltxtColor)
      End If
      If (nkeySel = 5) Then  ' exit
       Call GFXTextOut(lpBack, g_MenuPos(5).x, g_MenuPos(5).y, arText(MENU_EXIT), bytFontSize, ltxtselColor, ltxtbkColor)
       'If (bSelected) Then bytMenuState = 4
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        bytMenuState = 4
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
       End If
      Else
       Call GFXTextOut(lpBack, g_MenuPos(5).x, g_MenuPos(5).y, arText(MENU_EXIT), bytFontSize, ltxtColor)
      End If
      
       
    ' options
    Case 1
      ' update right panel
      Call UpdateInfoBox(1)
     
      If (nkeySel > 5) Then nkeySel = 0
      If (nkeySel < 0) Then nkeySel = 5
              
      If (nkeySel = 0) Then  ' calibrate gamma
       If ((Not bGamma)) Then
        nkeySel = nkeySel + 1
        Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_GAMMAUNAVAILABLE), bytFontSize, RGB(&HAA, &HAA, &HAA))
        'Exit Sub
       End If  ' else do normal gamma-text blit
      
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_GAMMA), bytFontSize, ltxtselColor, ltxtbkColor)
       
       ' gamma down
       If (DIKeyState(DIK_LEFT) = KS_KEYUP) Then
        m_gfxGammaRGB.r = m_gfxGammaRGB.r - 5
        If (m_gfxGammaRGB.r < -20) Then m_gfxGammaRGB.r = -20
        m_gfxGammaRGB.G = m_gfxGammaRGB.G - 5
        If (m_gfxGammaRGB.G < -20) Then m_gfxGammaRGB.G = -20
        m_gfxGammaRGB.B = m_gfxGammaRGB.B - 5
        If (m_gfxGammaRGB.B < -20) Then m_gfxGammaRGB.B = -20
        ' set new vals
        Call mDirectDraw.DDSetGamma(m_gfxGammaRGB.r, m_gfxGammaRGB.G, m_gfxGammaRGB.B)
        '' play calb. sound
        'Call DSPlaySound(g_dsSfx(SFX_MENUCALIBRATE), False)
       End If
       ' gamma up
       If (DIKeyState(DIK_RIGHT) = KS_KEYUP) Then
        m_gfxGammaRGB.r = m_gfxGammaRGB.r + 5
        If (m_gfxGammaRGB.r > 20) Then m_gfxGammaRGB.r = 20
        m_gfxGammaRGB.G = m_gfxGammaRGB.G + 5
        If (m_gfxGammaRGB.G > 20) Then m_gfxGammaRGB.G = 20
        m_gfxGammaRGB.B = m_gfxGammaRGB.B + 5
        If (m_gfxGammaRGB.B > 20) Then m_gfxGammaRGB.B = 20
        ' set new vals
        Call mDirectDraw.DDSetGamma(m_gfxGammaRGB.r, m_gfxGammaRGB.G, m_gfxGammaRGB.B)
        '' play calb. sound
        'Call DSPlaySound(g_dsSfx(SFX_MENUCALIBRATE), False)
       End If
       
      Else
       
       If ((Not bGamma)) Then
        Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_GAMMAUNAVAILABLE), bytFontSize, RGB(&HAA, &HAA, &HAA))
       Else
        Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_GAMMA), bytFontSize, ltxtColor)
       End If
      
      End If
      
      If (nkeySel = 1) Then  ' set sound vol.
       Call GFXTextOut(lpBack, g_MenuPos(1).x, g_MenuPos(1).y, arText(MENU_SOUND), bytFontSize, ltxtselColor, ltxtbkColor)
       
       ' volume down
       If (DIKeyState(DIK_LEFT) = KS_KEYDOWN) Then
        mDirectSound.m_nGlobalVol = mDirectSound.m_nGlobalVol - 20
        If (m_nGlobalVol < SFX_VOLUMEMIN) Then m_nGlobalVol = SFX_VOLUMEMIN
        Call DSPlaySound(g_dsSfx(SFX_MENUCALIBRATE), False)
       End If
       ' volume up
       If (DIKeyState(DIK_RIGHT) = KS_KEYDOWN) Then
        mDirectSound.m_nGlobalVol = mDirectSound.m_nGlobalVol + 20
        If (m_nGlobalVol > SFX_VOLUMEMAX) Then m_nGlobalVol = SFX_VOLUMEMAX
        Call DSPlaySound(g_dsSfx(SFX_MENUCALIBRATE), False)
       End If
       
      Else
       Call GFXTextOut(lpBack, g_MenuPos(1).x, g_MenuPos(1).y, arText(MENU_SOUND), bytFontSize, ltxtColor)
      End If
      
      If (nkeySel = 2) Then  ' set music vol.
       Call GFXTextOut(lpBack, g_MenuPos(2).x, g_MenuPos(2).y, arText(MENU_MUSIC), bytFontSize, ltxtselColor, ltxtbkColor)
      
       ' volume down
       If (DIKeyState(DIK_LEFT) = KS_KEYDOWN) Then
        g_nMusicVol = g_nMusicVol - 1
        If (g_nMusicVol < MUSIC_VOLUMEMIN) Then g_nMusicVol = MUSIC_VOLUMEMIN
         ' set FMOD music volume
         Call mFMod.VOLUMEmus(g_nMusicVol)
       End If
       ' volume up
       If (DIKeyState(DIK_RIGHT) = KS_KEYDOWN) Then
        g_nMusicVol = g_nMusicVol + 1
        If (g_nMusicVol > MUSIC_VOLUMEMAX) Then g_nMusicVol = MUSIC_VOLUMEMAX
        ' set FMOD music volume
        Call mFMod.VOLUMEmus(g_nMusicVol)
       End If
      
      Else
       Call GFXTextOut(lpBack, g_MenuPos(2).x, g_MenuPos(2).y, arText(MENU_MUSIC), bytFontSize, ltxtColor)
      End If
      
      If (nkeySel = 3) Then  ' VSync On/Off
       
       If (g_bNotRetrace) Then
        Call GFXTextOut(lpBack, g_MenuPos(3).x, g_MenuPos(3).y, arText(MENU_VSYNCOFF), bytFontSize, ltxtselColor, ltxtbkColor)
       Else
        Call GFXTextOut(lpBack, g_MenuPos(3).x, g_MenuPos(3).y, arText(MENU_VSYNCON), bytFontSize, ltxtselColor, ltxtbkColor)
       End If
       
       ' user pressed enter
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        If (g_bNotRetrace) Then g_bNotRetrace = False _
         Else g_bNotRetrace = True
        ' play on/off soud
        Call DSPlaySound(g_dsSfx(SFX_MENUCALIBRATE), False)
       End If
       
      Else
       
       If (g_bNotRetrace) Then
        Call GFXTextOut(lpBack, g_MenuPos(3).x, g_MenuPos(3).y, arText(MENU_VSYNCOFF), bytFontSize, ltxtColor)
       Else
        Call GFXTextOut(lpBack, g_MenuPos(3).x, g_MenuPos(3).y, arText(MENU_VSYNCON), bytFontSize, ltxtColor)
       End If
      
      End If
      
      If (nkeySel = 4) Then  ' Solid Health Bars
       
       If (g_bsolidbars) Then
        Call GFXTextOut(lpBack, g_MenuPos(4).x, g_MenuPos(4).y, arText(MENU_SOLIDBARSON), bytFontSize, ltxtselColor, ltxtbkColor)
       Else
        Call GFXTextOut(lpBack, g_MenuPos(4).x, g_MenuPos(4).y, arText(MENU_SOLIDBARSOFF), bytFontSize, ltxtselColor, ltxtbkColor)
       End If
       
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCALIBRATE), False)
        If (g_bsolidbars) Then g_bsolidbars = False Else g_bsolidbars = True
       End If
      
      Else
       
       If (g_bsolidbars) Then
        Call GFXTextOut(lpBack, g_MenuPos(4).x, g_MenuPos(4).y, arText(MENU_SOLIDBARSON), bytFontSize, ltxtColor)
       Else
        Call GFXTextOut(lpBack, g_MenuPos(4).x, g_MenuPos(4).y, arText(MENU_SOLIDBARSOFF), bytFontSize, ltxtColor)
       End If
      
      End If
     
      If (nkeySel = 5) Then  ' back
       Call GFXTextOut(lpBack, g_MenuPos(5).x, g_MenuPos(5).y, arText(MENU_BACK), bytFontSize, ltxtselColor, ltxtbkColor)
       
       ' return to main menu mode
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
        bytMenuState = 0
       End If
      
      Else
       Call GFXTextOut(lpBack, g_MenuPos(5).x, g_MenuPos(5).y, arText(MENU_BACK), bytFontSize, ltxtColor)
      End If
     
    ' blit status bars
    ' gamma ramp
    If (bGamma) Then
     Call BltFastGFX_HBM(g_MenuPos(0).x + 100, g_MenuPos(0).y + 10, g_Objects.caline)
     cn = ((m_gfxGammaRGB.r) + 50) '+ (Sgn(m_gfxGammaRGB.R) * 30)
     Call BltFastGFX_HBM(g_MenuPos(0).x + 100 + cn, g_MenuPos(0).y + 8, g_Objects.seline)
    End If
    ' sound fx
    Call BltFastGFX_HBM(g_MenuPos(1).x + 100, g_MenuPos(1).y + 10, g_Objects.caline)
    cn = 100 + ((mDirectSound.m_nGlobalVol) / 100)
    Call BltFastGFX_HBM(g_MenuPos(1).x + 100 + cn, g_MenuPos(1).y + 8, g_Objects.seline)
    ' music
    Call BltFastGFX_HBM(g_MenuPos(2).x + 100, g_MenuPos(2).y + 10, g_Objects.caline)
    cn = 100 + (g_nMusicVol - MUSIC_VOLUMEMAX)
    Call BltFastGFX_HBM(g_MenuPos(2).x + 100 + cn, g_MenuPos(2).y + 8, g_Objects.seline)
    
     
    ' hall of fame
    Case 2
     ' update infobox (right panel)
     Call UpdateInfoBox(2)
 
     If (nkeySel > 0 Or nkeySel < 0) Then nkeySel = 0
      
      If (nkeySel = 0) Then  ' back
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_BACK), bytFontSize, ltxtselColor, ltxtbkColor)
       
       ' return to main menu mode
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
        bytMenuState = 0
        ' sort by exp
        Call mCharacters.CHARSortCharacters
       End If
      
      Else
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_BACK), bytFontSize, ltxtColor)
      End If
    
    ' help
    Case 3
     
     ' update infobox (right panel)
     Call UpdateInfoBox(3)
     
     If (nkeySel > 0 Or nkeySel < 0) Then nkeySel = 0
      
      If (nkeySel = 0) Then  ' back
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_BACK), bytFontSize, ltxtselColor, ltxtbkColor)
       
       ' return to main menu mode
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
        bytMenuState = 0
       End If
      
      Else
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_BACK), bytFontSize, ltxtColor)
      End If
    
    ' exit
    Case 4
      g_Gates = GS_CLOSE
      mFMod.STOPmus
      bRunning = False
      
    ' Play/Char selection
    Case 5
      
      Dim strB       As String
      Dim strB1      As String
      Dim strB2      As String
      Dim bytlinePos As Byte
      ' chars. scrollbox value
      Static lOff    As Long
      Static lcnum   As Long
      Static lPage   As Long
      
      ' load characters
      lcnum = mCharacters.g_nChars
      
      ' page down
      If ((DIKeyState(DIK_PAGEDOWN) = KS_KEYUP Or _
           DIKeyState(DIK_NUMPAD3) = KS_KEYUP) And _
          g_nChars > ((lPage + 1) * 7)) Then lPage = lPage + 1
      ' page up
      If ((DIKeyState(DIK_PAGEUP) = KS_KEYUP Or _
           DIKeyState(DIK_NUMPAD9) = KS_KEYUP) And _
          lPage > 0) Then lPage = lPage - 1
      
      ' setup character page
      lOff = lPage * 7
      lcnum = g_nChars - lOff
      ' set char-list end
      If (lcnum > 7) Then
      ' more than 2 pages in list
       lcnum = 7 + lOff
      Else
      ' no more in list
       lcnum = lcnum + lOff
      End If
      
     ' bound selector
     If (nkeySel > (lcnum + 2)) Then nkeySel = lcnum + 2
     If (nkeySel < 0) Then nkeySel = 0 'lcnum + 2
     If (nkeySel < 2) Then
      ' draw char.menu help
      Call UpdateInfoBox(6)
     End If
     
     For cn = lOff To lcnum
       bytlinePos = cn - ((lOff / 7) * 7) + 2
       
       If (nkeySel = bytlinePos) Then  ' chars
        Call GFXTextOut(lpBack, g_MenuPos(bytlinePos).x, g_MenuPos(bytlinePos).y, g_arChar(cn).name, bytFontSize, ltxtselColor, ltxtbkColor)
         
        ' update infobox (right panel)
        Call UpdateInfoBox(5, CLng(cn))
       
        ' On selected character
        If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
         Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
         
           If (Len(Trim$(g_arChar(cn).pass)) < 1) Then
            '... select
             g_Player = g_arChar(cn)                ' set this as the current player
             g_nPlayerIndex = cn                    ' save index to apply scores later
             'Call DoGates(True)                     ' close gates
             Call StartBriefing                     ' send player to mission briefing
           Else
            ' bring up password dialog
            strB = ShowDialogBox(arText(MENU_PASSWORDQUERY), MAX_CX / 2, MAX_CY / 2, True, True)
            ' check password
            If ((Len(strB) < 1) Or (strB <> Trim$(g_arChar(cn).pass))) Then
             Call DoErrorBox(True, arText(MENU_PASSWORDERROR), MAX_CX / 2, MAX_CY / 2, True)
            Else
             '... select
             g_Player = g_arChar(cn)                ' set this as the current player
             g_nPlayerIndex = cn                    ' save index to apply scores later
             'Call DoGates(True)                     ' close gates
             'bMainMenu = False                      ' close main menu
             'g_Gates = GS_CLOSEOPEN
             Call StartBriefing
            End If
           End If
           
        End If ' end select char
       
        ' query delete character
        If (DIKeyState(DIK_DELETE) = KS_KEYUP) Then
         Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
         
         ' check for password player
         If (Len(Trim$(g_arChar(cn).pass)) < 1) Then
          ' no password, delete character
          Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
          If (Not mCharacters.CHARDelCharacter(cn)) Then Call DoErrorBox(True, arText(MENU_PLAYERDELETEERROR), MAX_CX / 2, MAX_CY / 2, True)
         Else
          ' query password
          strB = ShowDialogBox(arText(MENU_PASSWORDCONFIRM), MAX_CX / 2, MAX_CY / 2, True, False)
          ' password match
          If (strB = Trim$(g_arChar(cn).pass)) Then
           ' delete char
           Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
           If (Not mCharacters.CHARDelCharacter(cn)) Then Call DoErrorBox(True, arText(MENU_PLAYERDELETEERROR), MAX_CX / 2, MAX_CY / 2, True)
          Else
           ' invalid password
           Call DoErrorBox(True, arText(MENU_PASSWORDERROR), MAX_CX / 2, MAX_CY / 2, True)
          End If ' end password check
         End If ' end password query
         
        End If
      
       Else
        Call GFXTextOut(lpBack, g_MenuPos(bytlinePos).x, g_MenuPos(bytlinePos).y, g_arChar(cn).name, bytFontSize, lcnsttxtColor)
       End If
      Next
            
      
      If (nkeySel = 1) Then  ' Create New Player
       Call GFXTextOut(lpBack, g_MenuPos(1).x, g_MenuPos(1).y, arText(MENU_ADDNEWPLAYER), bytFontSize, ltxtselColor, ltxtbkColor)
       
       ' show create new player dialog
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
        
        ' bring name dialog
        strB = ShowDialogBox(arText(MENU_PLAYERNAMEQUERY), MAX_CX / 2, MAX_CY / 2, True, False)
        If (Len(strB) < 1) Then
         Call DoErrorBox(True, arText(MENU_PLAYERNAMEERROR), MAX_CX / 2, MAX_CY / 2, True)
        Else
         ' bring password dialog
         strB1 = ShowDialogBox(arText(MENU_PASSWORDENTER), MAX_CX / 2, MAX_CY / 2, True, True)
         If (Len(strB1) < 1) Then
          ' No Password {!}
           ' player ready
           '...
           Call mCharacters.CHARNewCharacter(strB, "")
           Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
          'Call DoErrorBox(True, arText(MENU_PASSWORDERROR), MAX_CX, MAX_CY, True)
         Else
          strB2 = ShowDialogBox(arText(MENU_PASSWORDCONFIRM), MAX_CX / 2, MAX_CY / 2, True, True)
          If (strB1 <> strB2) Then
           ' error password confirm
           Call DoErrorBox(True, arText(MENU_PASSWORDERROR), MAX_CX / 2, MAX_CY / 2, True)
          Else
           ' player ready
           Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
           Call mCharacters.CHARNewCharacter(strB, strB2)
          End If ' end confirm
         End If ' end password
        End If ' end player name
        
        'bytMenuState = 0
       End If
      
      Else
       Call GFXTextOut(lpBack, g_MenuPos(1).x, g_MenuPos(1).y, arText(MENU_ADDNEWPLAYER), bytFontSize, ltxtColor)
      End If
      
      If (nkeySel = 0) Then  ' back
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_BACK), bytFontSize, ltxtselColor, ltxtbkColor)
       
       ' return to main menu mode
       If (DIKeyState(DIK_RETURN) = KS_KEYUP) Then
        Call DSPlaySound(g_dsSfx(SFX_MENUCHOICE), False)
        bytMenuState = 0
       End If
      
      Else
       Call GFXTextOut(lpBack, g_MenuPos(0).x, g_MenuPos(0).y, arText(MENU_BACK), bytFontSize, ltxtColor)
      End If
 
      
  End Select
  
End Sub


'///////////////////////////////////////////////////////////////
'//// Refresh right main menu panel (the Info Box)
'//// BYTE bytMenuState - what is to be updated?
'///////////////////////////////////////////////////////////////
Public Sub _
UpdateInfoBox(bytMenuState As Byte, _
              Optional lExtra As Long = 0)
 
  'bytmenustate 0 - main menu, 1 - options, 2-hall of fame, 3-help , 4-exit, 5-play opts.
 ' 280 + 50 + 30 + 70, g_MenuPos(0).y, "ÊÐÅÄÈÒÈ", 24, RGB(25, 250, 25))
    
 Dim ltxtColor      As Long
 Dim ltxtTitleColor As Long
 Dim i              As Long
 Dim j              As Long
 Dim lsup           As Integer
 Dim tmpsz          As String
 Static ovc         As Byte
 Static ovcc        As Integer
 Static crdy        As Single
 
 ' prepare color
 If (ovc > 50) Then ovcc = -1
 If (ovc <= 0) Then ovcc = 1
 ovc = ovc + ovcc
 ltxtColor = RGB(50 + (ovc * 4), 20, 250 - (ovc * 3))
 ltxtTitleColor = RGB(225, 25, 25)
    
 Select Case bytMenuState
 
   ' draw credits
   Case 0, 1
   
     If (crdy = 0) Then
      crdy = rScreen.Bottom
     ElseIf (crdy < -700#) Then
      crdy = rScreen.Bottom
     End If
      
     Call BltFastGFX_HBM(360, crdy, g_Objects.credits)
     crdy = crdy - 0.3
     
   ' hall of fame
   Case 2
    Call GFXTextOut(lpBack, 370, g_MenuPos(0).y, _
                   arText(MENU_INFOBOXHOF), 20, ltxtTitleColor, , True)
    
    j = mCharacters.g_nChars
    If (j > 10) Then j = 10
    For i = 0 To j
     tmpsz = g_arChar(i).name & "  " & mCharacters.CHARLevelToString(g_arChar(i).level) & "  " & g_arChar(i).score
     
     Call GFXTextOut(lpBack, 340, g_MenuPos(i + 1).y, _
                     tmpsz, 20, ltxtColor, , True)
    Next
    
   
   ' help
   Case 3
    Call GFXTextOut(lpBack, 370, g_MenuPos(0).y, _
                   arText(MENU_INFOBOXHELP), 20, ltxtTitleColor, , True)
          
    ' put bonuses and their descriptions
    j = 0
    For i = MENU_HELP1 To MENU_HELP11
     j = j + 1
     Call GFXTextOut(lpBack, 340 + 32, g_MenuPos(j).y - lsup, arText(i), 20, ltxtTitleColor)
     lsup = lsup + 8 ' line y-distance offset
    Next
    lsup = 0
    For i = 0 To 5
     Call BltFastGFX_HBM(345, g_MenuPos(i + 5).y - lsup - 6, g_Objects.Bonus(i, 0))
     lsup = lsup + 9 ' line y-distance offset
    Next
    
   ' give selected char info
   Case 5
      ' title
      Call GFXTextOut(lpBack, 370, g_MenuPos(0).y, _
                     arText(MENU_INFOBOXPLAY), 20, ltxtTitleColor, , True)
      ' 280+50+30 = 340
      Call GFXTextOut(lpBack, 340, g_MenuPos(1).y, _
                     arText(MENU_PLAYERNAME) & g_arChar(lExtra).name, 20, ltxtColor, , True)
      
      Dim strMaskPass As String
      strMaskPass = String$(Len(g_arChar(lExtra).pass), "*")
      If (Len(Trim$(g_arChar(lExtra).pass)) < 1) Then strMaskPass = "none"
      
      Call GFXTextOut(lpBack, 340, g_MenuPos(2).y, _
                     arText(MENU_PLAYERPASSWORD) & strMaskPass, 20, ltxtColor, , True)
      Call GFXTextOut(lpBack, 340, g_MenuPos(3).y, _
                     arText(MENU_PLAYERSCORE) & g_arChar(lExtra).score, 20, ltxtColor, , True)
      Call GFXTextOut(lpBack, 340, g_MenuPos(4).y, _
                     arText(MENU_PLAYERLEVEL) & CHARLevelToString(g_arChar(lExtra).level), 20, ltxtColor, , True)
      Call GFXTextOut(lpBack, 340, g_MenuPos(5).y, _
                     arText(MENU_PLAYEREXPERIENCE) & g_arChar(lExtra).Exp, 20, ltxtColor, , True)
      Call GFXTextOut(lpBack, 340, g_MenuPos(6).y, _
                     arText(MENU_PLAYERMISSION) & g_arChar(lExtra).mission, 20, ltxtColor, , True)
      Call GFXTextOut(lpBack, 340, g_MenuPos(7).y, _
                     arText(MENU_PLAYERTOTALSHOTS) & g_arChar(lExtra).ts, 20, ltxtColor, , True)
      Call GFXTextOut(lpBack, 340, g_MenuPos(8).y, _
                     arText(MENU_PLAYERKILLS) & g_arChar(lExtra).kills, 20, ltxtColor, , True)
      ' get success percentage
      Dim fSuc As Single
      
      If (g_arChar(lExtra).ts <= 0) Then
       fSuc = 0#
      Else
       fSuc = ((g_arChar(lExtra).sit / g_arChar(lExtra).ts) * 100)  ', "##.##")
      End If
      
      Call GFXTextOut(lpBack, 340, g_MenuPos(9).y, _
                     arText(MENU_PLAYERSUCCESS) & fSuc & "%", 20, ltxtColor, , True)
   
   ' select character help
   Case 6
     Call GFXTextOut(lpBack, 340, g_MenuPos(3).y, arText(MENU_PLAYERHELP1), 20, ltxtTitleColor)
     Call GFXTextOut(lpBack, 340, g_MenuPos(4).y, arText(MENU_PLAYERHELP2), 20, ltxtTitleColor)
     Call GFXTextOut(lpBack, 340, g_MenuPos(5).y, arText(MENU_PLAYERHELP3), 20, ltxtTitleColor)
   
   ' on exit
   Case 4
   '...

 End Select
 

End Sub


'///////////////////////////////////////////////////////////////
'//// Add scrolling-message-queue
'//// STRING strMsg - message to add
'//// LONG   lDuration - time to scroll the message
'///////////////////////////////////////////////////////////////
Public Sub _
AddSMQ(strMsg As String, _
       lduration As Long)

 Dim cn As Integer
 
 ' no messages after MAX_INT
 If (g_numsmq > MAX_INT) Then
  Exit Sub
 ' trap error
 ElseIf (g_numsmq < -1) Then
  g_numsmq = -1
 '' add message
 'ElseIf (g_numsmq = 0) Then
 ''...
 'ElseIf (g_numsmq > ) Then
 End If
 
 g_numsmq = g_numsmq + 1
 
 ' widen array
 ReDim Preserve g_smq(g_numsmq)
 
 ' add
 g_smq(g_numsmq).strMsg = String$(14, " ") & strMsg
 g_smq(g_numsmq).lduration = lduration
 g_smq(g_numsmq).lTimeCounter = 0
 ' play SMQ recieved sound
 Call DSPlaySound(g_dsSfx(SFX_COCKPITSMQ), False)
   
End Sub

'///////////////////////////////////////////////////////////////
'//// Update scrolling-message-queue
'///////////////////////////////////////////////////////////////
Public Sub _
UpdateSMQ()
 
 ' see if there are any message in the queue
 If (g_numsmq < 0) Then Exit Sub
 
 Static strTxT   As String
 Static lTimeFPS As Long
 Static ntxtlen  As Integer
 Dim cn          As Integer
 
 ' is it time to update...
 If (lTimeFPS < GetTicks()) Then
  lTimeFPS = GetTicks() + (FPS_ANIMS + FPS_ANIMS)
  ' do scrolling
  ntxtlen = ntxtlen + 1
  strTxT = Mid$(g_smq(0).strMsg, ntxtlen, 13)
 End If
  
 ' output to cockpit
 Call DrawTextCP(4, strTxT)
  
 ' fill time counter (do it here...figure it out why ;)
 If (g_smq(0).lTimeCounter = 0) Then
  g_smq(0).lTimeCounter = GetTicks() + g_smq(0).lduration
 End If
  
 ' seek&destroy
 If (g_smq(0).lTimeCounter < GetTicks() Or _
     ntxtlen > Len(g_smq(0).strMsg)) Then
   
  ' step up
  For cn = 0 To (g_numsmq - 1)
   g_smq(cn) = g_smq(cn + 1)
  Next
  ' kill one
  g_numsmq = g_numsmq - 1
  ntxtlen = 0
 End If
 
End Sub

'///////////////////////////////////////////////////////////////
'//// Show modal dialog box
'//// BOOL   bNew - is this a new box or just refresh the old
'//// STRING strErrMsg - message to write out
'//// INT    x - horizontal position
'//// INT    y - vertival position
'//// BOOL   bCentered - center the box
'//// LONG   lTimeToStay - time before clearing the box
'///////////////////////////////////////////////////////////////
Public Static Sub _
DoErrorBox(bNew As Boolean, _
           Optional strErrMsg As String = "error code not found!", _
           Optional x As Integer = MAX_CX, _
           Optional y As Integer = MAX_CY, _
           Optional bCentered As Boolean = True, _
           Optional lTimeToStay As Long = 5000)

 ' destination coords
 Dim dx               As Integer
 Dim dy               As Integer
 Dim lTimePresence    As Long
 Dim bCenterit        As Boolean
 Dim strMsg           As String
 Dim rval             As Byte
 Dim rvalc            As Integer
 
 ' 7, 7   for title
 ' 7, 37  for user's text
 If (bNew) Then
  lTimePresence = GetTicks() + lTimeToStay
  
  strMsg = Left$(strErrMsg, 22)
  If (bCentered) Then bCenterit = True
  
  ' do coords settings
  If (bCenterit) Then
   dx = (x - (g_Objects.errdialog.cx / 2))
   dy = (y - (g_Objects.errdialog.cy / 2))
  Else
   dx = x
   dy = y
  End If
 End If
  
 ' do not refresh if no box acquired
 If (lTimePresence < GetTicks()) Then Exit Sub
 
 rval = rval + rvalc
 If (rval > 120) Then rvalc = -2
 If (rval < 15) Then rvalc = 2
 
 ' blit error box
 Call GFXTextOut(lpBack, dx + 7, dy + 7, strMsg, 20, RGB(80 + rval, 10, 10))
 Call BltFastGFX_HBM(dx, dy, g_Objects.errdialog)

End Sub

'///////////////////////////////////////////////////////////////
'//// Show modal dialog box
'//// STRING strTitle - title of the box
'//// INT    x - horizontal position
'//// INT    y - vertival position
'//// BOOL   bCentered - center the box
'//// Returns: String with user's input
'///////////////////////////////////////////////////////////////
Public Function _
ShowDialogBox(strTitle As String, _
              x As Integer, y As Integer, _
              bCentered As Boolean, _
              bMasked As Boolean) As String
 
 ' destination coords
 Dim dx         As Integer
 Dim dy         As Integer
 Dim cn         As Byte
 Dim nOff       As Integer
 Dim bEnd       As Boolean
 Dim bShift     As Boolean
 Dim strRead    As String
 Dim strMask    As String
 ' 7, 7   for title
 ' 7, 37  for user's text
 
 ' setup box
 strTitle = Left$(strTitle, 22)
 If (bCentered) Then
  dx = (x - (g_Objects.dialog.cx / 2))
  dy = (y - (g_Objects.dialog.cy / 2))
 Else
  dx = x
  dy = y
 End If
 
 Do While (Not bEnd)
   
  Call GFXClearBackBuffer
  ' refresh star travel
  Call UpdateStarTrip
  ' refresh logo animation
  Call UpdateLogo
  ' blit dialog box
  Call GFXTextOut(lpBack, dx + 7, dy + 7, strTitle, 20, RGB(180, 10, 10))
  Call BltFastGFX_HBM(dx, dy, g_Objects.dialog)
  ' blit user typed text
  If (Not bMasked) Then
   Call GFXTextOut(lpBack, dx + 7, dy + 37, strRead, 20, RGB(200, 30, 30))
  Else
   Call GFXTextOut(lpBack, dx + 7, dy + 37, strMask, 20, RGB(200, 30, 30))
  End If
  
  ' check keys
  If (DICheckKeyEx(DIK_RETURN) = KS_KEYUP) Then
   ShowDialogBox = strRead
   bEnd = True
  ElseIf (DICheckKeyEx(DIK_ESCAPE) = KS_KEYUP) Then
   ShowDialogBox = ""
   bEnd = True
  End If
  ' mentainance keys
  If (DICheckKeyEx(DIK_LSHIFT) = KS_UNPRESSED) Then
   bShift = False
  ElseIf (DICheckKeyEx(DIK_LSHIFT) = KS_KEYDOWN) Then
   bShift = True
  End If
  If (DICheckKeyEx(DIK_BACK) = KS_KEYUP) Then
   ' clear one char
   nOff = Len(strRead) - 1
   If (nOff < 0) Then nOff = 0
   strRead = Left$(strRead, nOff)
  ElseIf (DICheckKeyEx(DIK_SPACE) = KS_KEYUP) Then
   strRead = strRead & " "
  End If
  
  ' check numeric keys
  For cn = 2 To 11
   If (DICheckKeyEx(cn) = KS_KEYUP) Then
    strRead = strRead & Chr$(DIKeyToASCII(cn, bShift))
   End If
  Next
  ' check character keys
  For cn = DIK_A To DIK_L
   If (DICheckKeyEx(cn) = KS_KEYUP) Then
    strRead = strRead & Chr$(DIKeyToASCII(cn, bShift))
   End If
  Next
  For cn = DIK_Z To DIK_M
   If (DICheckKeyEx(cn) = KS_KEYUP) Then
    strRead = strRead & Chr$(DIKeyToASCII(cn, bShift))
   End If
  Next
  For cn = DIK_Q To DIK_P
   If (DICheckKeyEx(cn) = KS_KEYUP) Then
    strRead = strRead & Chr$(DIKeyToASCII(cn, bShift))
   End If
  Next
  If (DICheckKeyEx(DIK_MINUS) = KS_KEYUP) Then
   strRead = strRead & Chr$(DIKeyToASCII(DIK_MINUS, bShift))
  End If
    
   
  ' bound max string size
  If (Len(strRead) > 21) Then strRead = Left$(strRead, 21)
  ' check for mask
  If (bMasked) Then strMask = String$(Len(strRead), "*")
  
  ' do synchornization ( in other words - retracing )
  If (Not g_bNotRetrace) Then
   Do While (lpDD.WaitForVerticalBlank(DDWAITVB_BLOCKEND, 0) <> DD_OK)
   Loop
  End If
  
  ' swap buffers
  DoEvents
  Call DDBlitToPrim
 Loop

End Function


'///////////////////////////////////////////////////////////////
'//// Close Gates
'//// BOOL     bCloseGates - action to perform
'//// OPT.BOOL bCloseGame  - should it also close the game
'///////////////////////////////////////////////////////////////
Public Sub _
DoGates(bCloseGates As Boolean, _
        Optional bCloseGame As Boolean = False)
 
 On Local Error GoTo DOGATES_ERROR:
 
 Dim xOffset1     As Integer                    ' gate 1 - h.pos
 Dim xOffset2     As Integer                    ' gate 2 - h.pos
 Dim nspeed       As Integer                    ' speed of closing
 Dim lWaitTime    As Long                       ' time to wait before updateing
 Dim lpSurf       As DirectDrawSurface7
 Dim rSurf        As RECT
  
 If (bCloseGates) Then
  ' prepare open state
  If (g_Gates = GS_CLOSEOPEN) Then g_Gates = GS_OPEN _
   Else g_Gates = GS_NONE
  ' initial opened gates position
  xOffset1 = -320
  xOffset2 = 640
  nspeed = 4
  lWaitTime = 0
  ' play closing gates sound
  Call DSPlaySound(g_dsSfx(SFX_CLOSEGATE), False)
 Else
  ' initial closed gates position
  g_Gates = GS_NONE                             ' reset global state mode
  xOffset1 = 0
  xOffset2 = 320
  nspeed = -4
  lWaitTime = 0 'GetTicks() + 3000
  ' preserve backbuffer
  'SurfaceFromSurface lpBack, 640, 480, lpSurf, False
  Call SetRect(rSurf, 0, 0, MAX_CX, MAX_CY)
  Set lpSurf = CreateEmptySurface(640, 480)
  Call lpSurf.BltFast(0, 0, lpBack, rSurf, DDBLTFAST_WAIT)
  ' play openning gates sound
  Call DSPlaySound(g_dsSfx(SFX_OPENGATE), False)
 End If
 
 ' start gates loop
 Do While (1)
 
  'mDirectInput.DICheckKeys
  'If (mDirectInput.DIKeyState(DIK_ESCAPE)) Then Exit Do
 
  If (xOffset1 >= 0 And bCloseGates) Then
   nspeed = 0
   xOffset1 = 0
   xOffset2 = 320
   If (lWaitTime = 0) Then lWaitTime = 0 '1500 + GetTicks()
   If (lWaitTime < GetTicks()) Then Exit Do
  ElseIf (xOffset1 <= -320 And (Not bCloseGates)) Then
   nspeed = 0
   xOffset1 = -320
   xOffset2 = 640
   Exit Do
   'If (lWaitTime = 0) Then lWaitTime = 2500 + GetTicks()
   'If (lWaitTime < GetTicks()) Then Exit Do
  End If
  
  If (Not bCloseGates) Then
   ' clear with backbuffer's previous contents
   Call BltFast(0, 0, lpSurf, rSurf, False)
  End If
  
  ' blit gates
  Call BltFastGFX_HBM(xOffset1, 0, g_Objects.gate(0))
  Call BltFastGFX_HBM(xOffset2, 0, g_Objects.gate(1))
 
  ' increase speed
  xOffset1 = xOffset1 + nspeed
  xOffset2 = xOffset2 - nspeed
  
  ' increase closing speed
  'bytspeed = bytspeed + 1
  'If (bytspeed > 12) Then bytspeed = 12
  
  ' do synchornization ( in other words - retracing )
  If (Not g_bNotRetrace) Then
   Do While (lpDD.WaitForVerticalBlank(DDWAITVB_BLOCKEND, 0) <> DD_OK)
   Loop
  End If
  
  ' swap buffers
  DoEvents
  Call DDBlitToPrim
 Loop

 ' clear backbuffer
 Call GFXClearBackBuffer
 ' release temp surface
 Set lpSurf = Nothing
 ' check for game close
 If (bCloseGame) Then bRunning = False

Exit Sub


DOGATES_ERROR:
 Debug.Print "Error doing gates!"
 AppendToLog ("Error doing gates!")
End Sub


'///////////////////////////////////////////////////////////////
'//// Shake ground
'//// BOOL bNew - restart quake (kill previous)
'//// LONG lDureation - duration in ms.
'//// BYTE bytAmplitude - amplitude of the quake
'///////////////////////////////////////////////////////////////
Public Static Sub _
DoMoonQuake(Optional bNew As Boolean, _
            Optional lduration As Long, _
            Optional bytAmplitude As Byte = 1)
 
 Dim lTime As Long
 Dim bytAmp As Byte
 Dim nOffsetX As Integer, nOffsetY As Integer
  
  If (bNew = True) Then
   lTime = GetTicks + lduration
   bytAmp = bytAmplitude
  End If
  
  ' don't waste time processing if no quake is set
  If lTime < GetTicks Then
   UpdateWorld True, True, True, True
   Exit Sub
  End If
  
  'UpdateWorld True, True, True, True
  If lTime > GetTicks Then
   ' get random vibrations
   nOffsetX = nGetRnd(CInt(-bytAmp), CInt(bytAmp))
   nOffsetY = nGetRnd(CInt(-bytAmp), CInt(bytAmp))
   ' assign to world
   wx = wx + nOffsetX
   wy = wy + nOffsetY
   If wx < 0 Then wx = 0
   If wy > 0 Then wy = 0
   'If wy < -5 Then wy = -5
   If wx > SCREEN_PIXEL_WIDTH Then wx = SCREEN_PIXEL_WIDTH
   CMouse.SetX = CMouse.GetX + nOffsetX
   CMouse.SetY = CMouse.GetY + nOffsetY
  Else
   'UpdateWorld True, True, True, True
  End If
  
End Sub

Public Sub UpdateCockPit()
 
 'Static lFPSTime As Long
 Dim dx As Long, dy As Long
 Dim cn   As Integer
 
 ' work out cannon blitting coordinates
 dx = 0 - cpx * 4
 dy = cpy * 2 + 550 - g_Objects.CannonLeft.cy
 BltFastGFX_HBM dx, dy, g_Objects.CannonLeft
 dx = (MAX_CX - g_Objects.CannonRight.cx) + cpx * 4
 BltFastGFX_HBM dx, dy, g_Objects.CannonRight
 
 ' blit cockpit
 Call BltFastGFX_HBM(0, cpy + 180, g_Objects.CockPit)
 
 ' blit earth_hitpoints
 dx = 195 'MAX_CX / 2 - 250 / 2
 dy = 15
 g_Objects.earthhp.cx = g_hpEarth
 'Call BltFastGFX_HBM(dx, dy, g_Objects.earthhp)
  If (g_bsolidbars) Then
  Call BltFastGFX_HBM(dx, dy, g_Objects.earthhp)
 Else
  Call BltFxGFX_HBM(dx, dy, g_Objects.earthhp, R_OR)
 End If
' Stop
 
 ' blit battlestation hitpoints
 If (CBattleStation.GetVisible) Then
  dx = 270 'MAX_CX / 2 - 50
  dy = 30
  g_Objects.bshp.cx = CBattleStation.GetHitPoints
  If (g_bsolidbars) Then
    Call BltFastGFX_HBM(dx, dy, g_Objects.bshp)
  Else
    Call BltFxGFX_HBM(dx, dy, g_Objects.bshp, R_OR)
  End If
 End If
 
 ' blit selected weapons buttons
 Select Case PlayerWeapon
   Case PW_LASER
    Call BltFastGFX_HBM(300, 224 + 180 + cpy, g_Objects.buton)
    Call BltFastGFX_HBM(300, 224 + 180 + 20 + cpy, g_Objects.butoff)
    Call BltFastGFX_HBM(300, 224 + 180 + 40 + cpy, g_Objects.butoff)
   
   Case PW_MISSILE_CLOSERANGE
    Call BltFastGFX_HBM(300, 224 + 180 + cpy, g_Objects.butoff)
    Call BltFastGFX_HBM(300, 224 + 180 + 20 + cpy, g_Objects.buton)
    Call BltFastGFX_HBM(300, 224 + 180 + 40 + cpy, g_Objects.butoff)
   
   Case PW_MISSILE_LONGRANGE
    Call BltFastGFX_HBM(300, 224 + 180 + cpy, g_Objects.butoff)
    Call BltFastGFX_HBM(300, 224 + 180 + 20 + cpy, g_Objects.butoff)
    Call BltFastGFX_HBM(300, 224 + 180 + 40 + cpy, g_Objects.buton)
  
  End Select
 
 ' update enemies on radar
 'xy:237,235+180  size:50x35
 For cn = 0 To MAX_ENEMIES
  If (CShip(cn).GetVisible) Then
   
   Select Case CShip(cn).GetZ
    Case PLANE_CLOSE
     dx = (CShip(cn).GetX \ 50) + 237
     dy = ((CShip(cn).GetY \ 35) * 3) + 415 + cpy
    Case PLANE_FAR                           ' *2.3
     dx = ((CShip(cn).GetX \ PLANE_FAR \ 50) * 4) + 237
     dy = ((CShip(cn).GetY \ 35) * 3) + 415 + cpy
   End Select
   
   Call BltFastGFX_HBM(dx, dy, g_Objects.es1)
  
  End If
 Next
 
 For cn = 0 To MAX_METEORS
  If (g_Meteor(cn).Visible) Then
   If ((g_Meteor(cn).x \ g_Meteor(cn).z) > 0) Then
    
   Select Case g_Meteor(cn).z
    Case PLANE_CLOSE
     dx = (g_Meteor(cn).x \ 50) + 237
     dy = ((g_Meteor(cn).y \ 35) * 3) + 415 + cpy
    Case PLANE_FAR
     dx = ((g_Meteor(cn).x \ PLANE_FAR \ 50) * 2.3) + 237
     dy = ((g_Meteor(cn).y \ 35) * 3) + 415 + cpy
   End Select
    
   Call BltFastGFX_HBM(dx, dy, g_Objects.es2)
   
   End If
  End If
 Next
  
 ' do 'return to normal position' animation
 'If lFPSTime < GetTicks Then
  If cpx > 0 Then cpx = cpx - 1
  If cpy > 0 Then cpy = cpy - 1
' lFPSTime = GetTicks + FPS_ANIMS
 'End If
 
End Sub

Public Sub UpdateCloseEnemies()
 ' Desc: Render closer enemy ships
 Dim cn As Long
 
 For cn = 0 To MAX_ENEMIES
  If (CShip(cn).GetVisible And CShip(cn).GetZ = 1) Then
    CShip(cn).Render
  End If
 Next

End Sub

Public Sub UpdateFarEnemies()
 ' Desc: Render distant enemy ships
 Dim cn As Long
 
 For cn = 0 To MAX_ENEMIES
  If (CShip(cn).GetVisible And CShip(cn).GetZ = 2) Then
    CShip(cn).Render
  End If
 Next
  
End Sub

Public Sub UpdateBunkers()
 ' Desc: Render all bunkers animations
 Dim cn         As Long
 Dim lTicks     As Long
 Static lUBTime As Long
 
 For cn = 0 To MAX_BUNKERS
  'cBunker(cn).SetRotation = Rnd * 1
  'cBunker(cn).SetFire = True
  CBunker(cn).Render
 Next
 
 ' update bunkers 'lock and load' AI
 lTicks = GetTicks()
 If (lTicks > lUBTime) Then
  lUBTime = lTicks + 50
  For cn = 0 To MAX_BUNKERS
   CBunker(cn).GetTarget
   CBunker(cn).FireWeapon
  Next
 End If
 
End Sub

Public Sub UpdateMoonSurface()
 ' Desc: Redraws close moon surface

 ' blit farmoon
 'Call SetRect(rMS, 0, 0, 640, 25)
 BltFastGFX_HBM 0 - wx / 3, MAX_CY - 120 - wy, g_Objects.BackMoon(0)
 BltFastGFX_HBM 640 - wx / 3, MAX_CY - 120 - wy, g_Objects.BackMoon(1)
 
 ' blit surface 1
 BltFastGFX_HBM arMS_Offsets(0) - wx, MAX_CY - 130 - wy, g_Objects.MoonSurf(0)
 ' blit surface 2 to 4 are the light side ( play area )
 BltFastGFX_HBM arMS_Offsets(1) - wx, MAX_CY - 130 - wy, g_Objects.MoonSurf(1)
 BltFastGFX_HBM arMS_Offsets(2) - wx, MAX_CY - 130 - wy, g_Objects.MoonSurf(2)
 BltFastGFX_HBM arMS_Offsets(3) - wx, MAX_CY - 130 - wy, g_Objects.MoonSurf(3)
 ' blit surface 2
 'BltFast arMS_Offsets(4) - wx, 0 - wy, g_Objects.MoonSurf(0), rMS, True
  
End Sub

Public Sub UpdateBackGround()
 ' Desc: Redraws background images of our Galaxy ;)
 Dim TempX As Single, TempY As Single

 xBg1 = xBg1 - 0.5 '0.5
 xBg2 = xBg2 - 0.5 '0.5
 xBg3 = xBg3 - 0.5 '0.5
 yBg1 = -5
 yBg2 = -5
 yBg3 = -5
 
 If (xBg1 + MAX_CX) < 0 Then xBg1 = SCREEN_PIXEL_WIDTH - 720
 If (xBg2 + MAX_CX) < 0 Then xBg2 = xBg1 + MAX_CX  ' SCREEN_PIXEL_WIDTH \ 2
 If (xBg3 + MAX_CX) < 0 Then xBg3 = xBg2 + MAX_CX  ' SCREEN_PIXEL_WIDTH \ 2
 
 TempX = xBg1 - wx / PLANE_FAR
 TempY = yBg1 - wy / PLANE_FAR
  'Call BltFastGFX_HBM(TempX, TempY, g_Objects.BackGround(0))
 TempX = -1 + xBg2 - wx / PLANE_FAR
 TempY = yBg2 - wy / PLANE_FAR
  'Call BltFastGFX_HBM(TempX, TempY, g_Objects.BackGround(1))
 TempX = -2 + xBg3 - wx / PLANE_FAR
 TempY = yBg3 - wy / PLANE_FAR
  'Call BltFastGFX_HBM(TempX, TempY, g_Objects.BackGround(2))
End Sub

Public Sub UpdateEarth()
 ' Desc: Redraws our homeworld ;)
 Dim dx As Integer, dy As Integer
 Static lTimeHeal As Long
 'Static bytFrame As Byte
 'Static lFPSTime As Long                            ' update frame timer
  
  'Call SetRect(rSrc, 0, 0, 400, 360)
  dx = g_xEarth
  dy = g_yEarth
  dx = dx - wx / PLANE_FAR
  dy = dy - wy / PLANE_FAR

  ' blit sun
  'Call BltFastGFX_HBM(dx, dy - 65, g_Objects.Sun)
  Call BltFastGFX_HBM(g_xEarth - wx / 5, g_yEarth - wy / 5 - 65, g_Objects.Sun)
  Call BltFastGFX_HBM(dx, dy, g_Objects.Earth)

  ' heal 1 hp every 500 ms
  If (lTimeHeal < GetTicks) Then
   lTimeHeal = GetTicks + 500
   If (g_hpEarth < EARTH_HITPOINTS) Then g_hpEarth = g_hpEarth + 1
  End If
  
'  Stop
  'BltFastGFX_HBM dx, dy, g_Objects.Earth(bytFrame)
  
 'If lFPSTime + (FPS_ANIMS * 3) < GetTicks Then
 '   lFPSTime = GetTicks
 '   bytFrame = bytFrame + 1
 '   If bytFrame > 35 Then bytFrame = 0
 'End If
    
End Sub

Public Sub UpdateStarField()                        ' draw the star-field
 
 Dim cn As Long
 Dim dx As Integer, dy As Integer
 Dim rSrc  As RECT
 
 For cn = 0 To MAX_STARS
   With Star(cn)
     
     If (.x >= (wx / .z) And .x <= (wx / .z) + MAX_CX) Then
        dx = .x - (wx / .z)                         ' calculate star position
        dy = .y - (wy / .z)
        'Call BltFastGFX_HBM(dx, dy, g_Objects.Star1(0))
        rSrc.Left = .frame
        rSrc.Right = rSrc.Left + 1
        rSrc.Top = 0
        rSrc.Bottom = 1
        ' do blinking
        If (nGetRnd(0, 500) > 100) Then
         .frame = .frame + 1
         If (.frame > 149) Then .frame = 0
        End If
        
        Call BltFastW(dx, dy, g_Objects.BlueCles.dds, rSrc, False)
        
     End If
     
   End With
 Next
End Sub

Public Sub UpdateWarpGates()
 ' Desc: Updates Opened WarpGates animations
 Dim cn As Long                                     ' local counter
 Static lFPSTime As Long                            ' update frame timer
 Dim bUpdateFrame As Boolean                        ' this tells the loop if anim. frame should be increased
 Dim brval  As Boolean
 Dim lTicks As Long
 
   lTicks = GetTicks()
   If lFPSTime < lTicks Then
      lFPSTime = lTicks + FPS_ANIMS '/ 2
      bUpdateFrame = True
   End If
  
 For cn = 0 To MAX_WARPGATES
  With g_WarpGate(cn)
   If .Visible Then
    
    If .z = 1 Then
     ' try raster blit
     'brval = BltFX(.x - wx / .z, .y - wy / .z, g_Objects.WarpGate(.Frame), rWG, R_OR, True)
     'brval = BltFXHel(.x - wx / .z, .y - wy / .z, g_Objects.WarpGate(.Frame), rWG, R_OR)
     'If Not brval Then
     ' raster failed, so do no raster blit
      'Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.WarpGate(.frame))
     'End If
    
    ElseIf .z = 2 Then
     'Call SetRect(rWG, 0, 0, 60, 50)
     ' try raster blit
     'brval = BltFX(.x - wx / .z, .y - wy / .z, g_Objects.WarpGate_Far(.Frame), rWG, R_OR, True)
     ' brval = BltFXHel(.x - wx / .z, .y - wy / .z, g_Objects.WarpGate_Far(.Frame), rWG, R_OR)
     'If Not brval Then
     ' raster failed, so do no raster blit
      Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.WarpGate_Far(.frame))
     'End If
    End If
    'BltFX .x - wx / .z, .y - wy / .z, g_Objects.WarpGate(.Frame), rWG, R_OR, True
     
     If bUpdateFrame Then ' /---  see if anim.frame should be updated ---/
      If Not .bPlayBack Then
       .frame = .frame + 1
       If .frame > UBound(g_Objects.WarpGate_Far) Then
        .frame = .frame - 2
        .bPlayBack = True
       End If
      Else ' play backwards
       .frame = .frame - 1
       If .frame < 1 Then .Visible = False  ' kill WarpGate
      End If
      
     End If ' /--- end update frame ---/
    
   End If
  End With
 Next
 
End Sub

Public Sub UpdateParticleExplosions()
 
 Dim i As Long, j As Long
 Dim bDontKill    As Boolean
 Dim bUpdateFrame As Boolean
 Dim rBlueCle     As RECT
 Dim lTicks       As Long
 Static lFPSTime  As Long
 
  lTicks = GetTicks()
  If (lFPSTime < lTicks) Then
     lFPSTime = lTicks + FPS_ANIMS
     bUpdateFrame = True
  End If
  
 For j = 0 To MAX_EXPLOSIONS
  With g_PExp(j)
   If (.Visible) Then
     ' Debug.Print "par: " & j & " is visible."

    For i = 0 To MAX_PARTICLES
      With .Particle(i)
       If (.Visible) Then
          bDontKill = True
        'If bUpdateFrame Then
        
         If .Heading = 0 Then                             ' left-up
            .x = .x - .xVel
            .y = .y - .yVel
            ' do gravity
            .yVel = .yVel + .yFriction
            If .yVel > 0.3 Then .yFriction = .yFriction - .yFriction - .yFriction
         ElseIf .Heading = 1 Then                         ' left-down
            .x = .x - .xVel
            .y = .y + .yVel
         ElseIf .Heading = 2 Then                         ' right-up
            .x = .x + .xVel
            .y = .y - .yVel
            ' do gravity
            .yVel = .yVel + .yFriction
            If .yVel > 0.3 Then .yFriction = .yFriction - .yFriction - .yFriction
         ElseIf .Heading = 3 Then                         ' right-down
            .x = .x + .xVel
            .y = .y + .yVel
         End If
        
        'If bUpdateFrame = True Then
         .frame = .frame + 1
         If .frame > 150 Then .frame = 0
        'End If
        
         ' decrease life
         .Life = .Life - 1
         If (.Life <= 1) Then .Visible = False
        'End If
         ' blit particles
         Call SetRect(rBlueCle, .frame, 0, .frame + nGetRnd(1, 2), 1)
         'BltFastW .x - wx / 1, .y - wy / 1, g_Objects.BlueCles, rBlueCle, False
         Call BltFastW(.x - wx / 2, .y - wy / 2, g_Objects.BlueCles.dds, rBlueCle, False)
    
    ' do color decrease
    '   .r = .r + .cf / 3
    '   If .r > 230 Then .cf = -.cf
    '   If .r <= 6 Then .r = 6
       
    '   .g = .g + .cf / 2
    '   If .g > 230 Then .cf = -.cf
    '   If .g <= 6 Then .g = 6
       
    '   .b = .b + .cf
    '   If .b > 230 Then .cf = -.cf
    '   If .b <= 6 Then .b = 6
       'If InRange(.x, wx, wx + MAX_CX) Then
         'And InRange(.y, wy, wy + MAX_CY) Then
       '  lpBack.SetForeColor RGB(.r, .g, .b)
         'lpBack.DrawLine .x - wx / 2, .y - wy / 2, .x + 1 - wx / 2, .y - wy / 2
       ' lpBack.DrawLine .x - wx, .y - wy, .x + 1 - wx, .y - wy
       'End If
       End If
      End With
    Next
    
    ' kill explosion if all particles died
    If bDontKill = False Then .Visible = False
     
   End If
  End With
 Next
 
  ' End If

End Sub

Public Sub UpdateChillingPixels()
 ' Desc: Redraws created chilling trails
 Dim cn           As Long
 Dim rFireCle     As RECT
 Dim bUpdateFrame As Boolean
 Dim lTicks       As Long
 Static lFPSTime  As Long
 
 lTicks = GetTicks()
 If (lFPSTime < lTicks) Then
    lFPSTime = lTicks + FPS_ANIMS
    bUpdateFrame = True
 End If
 
 For cn = 0 To MAX_CPIXELS
  With g_CPixel(cn)
   If .Visible Then
     .x = .x + .xVel
     .y = .y + .yVel
     
     .Life = .Life - 1
     If .Life <= 1 Then
      .Visible = False
     End If
    
    'If bUpdateFrame Then
     .frame = .frame + 1
     If .frame > 150 Then .frame = 0
    'End If
     
     ' blit particles
     Call SetRect(rFireCle, .frame, 0, .frame + nGetRnd(1, 2), 1)
     Call BltFastW(.x - wx / .z, .y - wy / .z, g_Objects.ChillCles.dds, rFireCle, False)
      
   ' *** Primitives ARE SLOW ***
   ' If InRange(.x, wx, wx + MAX_CX) Then
      'lpBack.SetForeColor RGB(.r, .g, .b)
   '    .cc = .cc - 5
   '   If .cc < 15 Then .cc = 15
   '   lpBack.SetForeColor RGB(.cc, 60, 100)
   '   'lpBack.DrawLine .x - wx / .z, .y - wy / .z, _
   '                   .x - wx / .z + 1, .y - wy / .z
   '   'lpBack.DrawCircle .x - wx / .z, .y - wy / .z, 2
   '   lpBack.DrawEllipse .x - wx / .z, .y - wy / .z, _
                         .x - wx / .z + 2, .y - wy / .z + 2
   '  Dim hTemp As Long
   '  lpBack.restore
   '  hTemp = lpBack.GetDC()
   '   SetPixel hTemp, .x - wx / .z, .y - wy / .z, RGB(.cc, 80, 80)
   '  lpBack.ReleaseDC hTemp
   ' End If
   
   End If
  End With
 Next

End Sub

Public Sub UpdateMissiles()
 ' Desc: Updates all visible missiles
 Dim cn As Long
 Dim rM As RECT, rD As RECT
 Dim rRocket As RECT
 Dim nTemp   As Integer
 
 For cn = 0 To MAX_MISSILES
  With g_Missile(cn)
   If .Visible Then
     
     If (nGetRnd(0, 2000) < 1000) Then
       ' create chilling pixel trail
       Call CreateChillingPixel(CInt(.x), CInt(.y + 9), CByte(.z))
     End If
     
     .xVel = .xVel + 0.02
     .yVel = .yVel + 0.02
     If .xVel > 2.2 Then .xVel = 2.2 '3.5
     If .yVel > 2.2 Then .yVel = 2.2 '23.5
     If .Direction = SO_LEFT Then
      .x = .x - .xVel / .xVelB
     ElseIf .Direction = SO_RIGHT Then
      .x = .x + .xVel / .xVelB
     End If
     
     If (.y < .dy) Then
      .y = .y + .yVel / .yVelB
     Else
      .y = .y - .yVel / .yVelB
     End If
     
     ' setup some rects
     Call SetRect(rM, .x, .y, .x + 20, .y + 20)
     Call SetRect(rD, .dx, .dy, .dx + .dcx, .dy + .dcy)
'     Stop
     'Call SetRect(rRocket, 0, 0, 20, 18)
     
     ' check if destination is hit
     If Collide(rM, rD) Then
'      Stop
      .Visible = False
      Call CreateExplosion(CInt(.x), CInt(.y), .z, ET_SMALL)
      
      ' if an enemy ship fired it then tell class to handle the damage
      If .Possession <> NO_POSSESSION Then
       CShip(.Possession).DoEnemyDamage
      End If
      
      ' ...
     ElseIf (.x > SCREEN_PIXEL_WIDTH Or .y > SCREEN_PIXEL_HEIGHT) Then
       .Visible = False
     End If
       
     If .z = PLANE_CLOSE Then                          ' only close-plain missiles are to be resized
      ' determine which frame should be displayed ( based on distance from screen )
      nTemp = nGetDist2D(CInt(.x), CInt(.y), .dx, .dy) ' get distance from dest
      If (nTemp < .nTempVar) And _
         .ObjectiveDist = BP_CLOSE Then
       .Distance = BP_CLOSE
      ElseIf (nTemp < .nTempVar + .nTempVar) And _
         .ObjectiveDist = BP_FAR Then
       .Distance = BP_FAR
      ElseIf (nTemp < .nTempVar + .nTempVar + .nTempVar) And _
         .ObjectiveDist = BP_VERYFAR Then
       .Distance = BP_VERYFAR
      End If
     End If
     
       '.Frame = .Frame + 1
     If .Direction = SO_LEFT Then
      
      Select Case .Distance
       Case BP_VERYFAR
        'Call SetRect(rRocket, 0, 0, 15, 13)
        Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.RocketL(.frame))
       Case BP_FAR
        'Call SetRect(rRocket, 0, 0, 20, 18)
        Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.RocketL_Close(.frame))
       Case BP_CLOSE
        'Call SetRect(rRocket, 0, 0, 25, 23)
        Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.RocketL_VClose(.frame))
      End Select
      
     ElseIf .Direction = SO_RIGHT Then
      
      Select Case .Distance
       Case BP_VERYFAR
        'Call SetRect(rRocket, 0, 0, 15, 13)
        Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.RocketR(.frame))
       Case BP_FAR
        'Call SetRect(rRocket, 0, 0, 20, 18)
        Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.RocketR_Close(.frame))
       Case BP_CLOSE
        'Call SetRect(rRocket, 0, 0, 25, 23)
        Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.RocketR_VClose(.frame))
      End Select
      
     End If
        
   End If
  End With
 Next
 
End Sub

Public Sub UpdateLaserCuts()
 ' Desc: Update all fired lasercuts
 
 Dim cn     As Long
 Dim rLaser As RECT
 Dim rDest  As RECT, rSrc As RECT
 
 For cn = 0 To MAX_LASERS
  With g_LaserCut(cn)
   If .Visible Then
    .x = .x + .xVel '/ .xVelB
    .y = .y + .yVel '/ .yVelB
     
     ' check if destination is hit
    Call SetRect(rSrc, .x, .y, .x + 10, .y + 9)
    Call SetRect(rDest, .dx, .dy, .dx + .dcx, .dy + .dcy)
    Call SetRect(rLaser, 0, 0, 10, 9)
    If Collide(rSrc, rDest) Then
     .Visible = False
     
     If (.Possession <> NO_POSSESSION) Then
      CShip(.Possession).DoEnemyDamage
     End If
     
     '...
    ElseIf (.x > SCREEN_PIXEL_WIDTH Or .y > SCREEN_PIXEL_HEIGHT) Then
     .Visible = False
    End If
     
    ' blit laser(cheto)
    Select Case .kind
     Case SW_LASER
      BltFastGFX_HBM .x - wx / .z, .y - wy / .z, g_Objects.RedLaser(.Direction)
     Case SW_GREENLASER
      BltFastGFX_HBM .x - wx / .z, .y - wy / .z, g_Objects.GreenLaser(.Direction)
      'BltFast .x - wx / .z, .y - wy / .z, g_Objects.GreenLaser(.Direction - 2), rLaser, True
     Case Else: Debug.Print "Wrong weapon in 'Laser Cuts proc'!"
    End Select
    
    'If .Direction = SO_LEFTUP Then
    ' BltFast .x - wx / .z, .y - wy / .z, g_Objects.RedLaserUL, rLaser, True
    'ElseIf .Direction = SO_LEFTDOWN Then
    ' BltFast .x - wx / .z, .y - wy / .z, g_Objects.RedLaserDL, rLaser, True
    'ElseIf .Direction = SO_RIGHTUP Then
    ' BltFast .x - wx / .z, .y - wy / .z, g_Objects.RedLaserUR, rLaser, True
    'ElseIf .Direction = SO_RIGHTDOWN Then
    ' BltFast .x - wx / .z, .y - wy / .z, g_Objects.RedLaserDR, rLaser, True
    'End If
   End If
  End With
 Next
 
End Sub

Public Sub DrawTextCP(Row As Byte, lpStr As String, _
                      Optional lClr As Long = 8556930)
 ' Desc: OutPuts text to cockpit's monitor
 
 Dim x As Integer, y As Integer
 ' 358, 235 + +180
 x = 358 '+ cpx '290 + cpx
 y = 415 + (Row * 10) + cpy
 
 Call GFXTextOut(lpBack, x, y, lpStr, 14, lClr)
 'lpBack.SetForeColor lClr
 'lpBack.DrawText x, y, lpStr, False
 
End Sub

Public Sub UpdateStarTrip()
 ' Desc: Updates SpaceTrip Simulator
 
 Dim cn       As Long
 Dim lTicks   As Long
 Static lTime As Long
 Dim bUpdate  As Boolean
 
 lTicks = GetTickCount()
 If (lTime < lTicks) Then
  lTime = lTicks + FPS_ANIMS / 10
  bUpdate = True
 End If
 
 For cn = 0 To (MAX_STARS / 3)
  With g_StarTrip(cn)
   If (.Visible) Then
      
     If (bUpdate) Then
      Select Case .Heading
        Case 0                                          ' North-West
          '.xVel = .xVel - 0.5
          .x = .x - .xVel
          .y = .y - .yVel
        Case 1                                          ' South-West
          .x = .x - .xVel
          .y = .y + .yVel
        Case 2                                          ' South-East
          .x = .x + .xVel
          .y = .y + .yVel
        Case 3                                          ' North-East
          .x = .x + .xVel
          .y = .y - .yVel
        Case Else: Debug.Print "Error setting startrip star direction!"
      End Select
      
      '   If .x < 80 Or .x > MAX_CX - 80 Or _
            .y < 80 Or .y > MAX_CY - 80 Then
          
      '    .xVel = .xVel + 0.1
      '    .yVel = .yVel + 0.1
      '   End If
         
          '.r = .r + 4
          'If .r >= 250 Then .r = 250
          '.G = .G + 4
          'If .G >= 250 Then .G = 250
          '.B = .B + 5
          'If .B >= 250 Then .B = 250
          
          .frame = .frame + 1 '(.xVel + .yVel)
          If (.frame > 145) Then .frame = 146
     End If
     
          ' draw the star
          'lpBack.SetForeColor RGB(.r, .G, .B)
          'lpBack.DrawLine .x, .y, .x + 1, .y
          Dim rStar As RECT
          Call SetRect(rStar, .frame, 0, .frame + 1, 1)
          'Stop
          Call BltFastW(.x, .y, g_Objects.starcles.dds, rStar, False)
          
          
          If (.x < 0 Or .x > MAX_CX Or _
              .y < 0 Or .y > MAX_CY) Then
           .Visible = False
          End If
   Else   ' create new star
     .x = MAX_CX / 2
     .y = MAX_CY / 2
     .xVel = fGetRnd(0.7, 2.5)
     .yVel = fGetRnd(0.7, 2.5)
     .Heading = nGetRnd(0, 3)
     .r = 0
     .G = 0
     .B = 0
     .Life = 0
     .Visible = True
     .frame = 0
   End If
   
  End With
 Next
End Sub

Public Sub RotateText(dhDC As Long, x As Integer, y As Integer, _
                      lpStr As String, Size As Integer, Angle As Integer, _
                      lForeColor As Long, _
                      Optional bBold As Boolean = False, _
                      Optional lpszFontName As String = "Arial")
 Dim cn As Integer
 Dim hFont As Long
 Dim hOldFont As Long
 Dim Weight As Long
   
 If bBold Then
    Weight = FW_NORMAL
 Else
    Weight = FW_BOLD
 End If
   
 ' create the font
 hFont = CreateFont(Size, 0, Angle * 10, 0, Weight, _
                   False, False, False, _
                   DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
                   CLIP_DEFAULT_PRECIS, 0, _
                   DEFAULT_PITCH Or FF_DONTCARE, _
                   lpszFontName)
                    'PROOF_QUALITY,
 
 hOldFont = SelectObject(dhDC, hFont)
 Call SetBkMode(dhDC, TRANSPARENT)                      ' set font transparency
 Call SetTextColor(dhDC, lForeColor)                    ' set font color
 Call TextOut(dhDC, x, y, lpStr, Len(lpStr))            ' draw the text
 ' select back old font
 
 hFont = SelectObject(dhDC, hOldFont)
 ' remove the font from GDI heap
 If DeleteObject(hFont) = 0 Then
    Debug.Print "Error removing font from GDI heap!"
 End If
End Sub

Public Sub UpdateExplosions()
 ' Desc: Update all closer explosions animations
 Dim cn As Long
 Dim dx As Integer, dy As Integer                   ' explosion blitting destination
 Dim rExp As RECT
 Dim bUpdateFrame As Boolean                        ' this tells the loop if anim. frame should be increased
 Dim MaxFrames   As Byte
 Dim lTicks      As Long
 Static lFPSTime As Long                            ' update frame timer
 
   lTicks = GetTicks()
   If lFPSTime < lTicks Then
      lFPSTime = lTicks + FPS_ANIMS
      bUpdateFrame = True
   End If
 
 For cn = 0 To MAX_EXPLOSIONS
  With g_Explosion(cn)
   If .Visible And ( _
      .kind = ET_BIG Or _
      .kind = ET_SMALL Or _
      .kind = ET_SMALLBLUE) Then
    
    If .kind = ET_BIG Then                          ' determine explosion kind
     Call SetRect(rExp, 0, 0, 160, 120)             ' setup blitting rectangle
     dx = (.x - rExp.Right / 2) - wx / .z
     dy = (.y - rExp.Bottom / 2) - wy / .z
     Call BltFastGFX_HBM(dx, dy, g_Objects.Exp1(.frame))
     MaxFrames = UBound(g_Objects.Exp1)
     
     ' blit small explosion
    ElseIf .kind = ET_SMALL Then
     Call SetRect(rExp, 0, 0, 80, 60)
     dx = (.x - rExp.Right / 2) - wx / .z
     dy = (.y - rExp.Bottom / 2) - wy / .z
     If (Not mDirectDraw.bHardwareRasters) Then
      Call BltFastGFX_HBM(dx, dy, g_Objects.Exp2(.frame))
     Else      ' blit pseudo-transparent
      'BltFX dx, dy, g_Objects.Exp2(.frame), rExp, R_OR, True
     End If
     'BltFX .x - wx / .z, .y - wy / .z, g_Objects.Exp2(.Frame), rExp, R_OR, True
     'BltFXGDI .x - wx / .z, .y - wy / .z, g_Objects.Exp2DC(.Frame), rExp, R_OR
     MaxFrames = UBound(g_Objects.Exp2)
     
     ' blit small blue explosion
    ElseIf .kind = ET_SMALLBLUE Then
     Call SetRect(rExp, 0, 0, 80, 60)
     dx = (.x - rExp.Right / 2) - wx / .z
     dy = (.y - rExp.Bottom / 2) - wy / .z
     If Not mDirectDraw.bHardwareRasters Then
      Call BltFastGFX_HBM(dx, dy, g_Objects.Exp3(.frame))
     Else      ' blit pseudo-transparent
      'BltFX dx, dy, g_Objects.Exp3(.frame), rExp, R_OR, True
     End If

     MaxFrames = UBound(g_Objects.Exp3)
     
    End If
    
     If bUpdateFrame Then                               ' see if anim.frame should be updated
      .frame = .frame + 1
      If .frame > MaxFrames Then
       .Visible = False
      End If
     End If
    
   End If
  End With
 Next

End Sub

' //////////////////////////////////////////////////////////
' //// Update all Far explosions animations ( Needed since
' //// closer objects will block the view )
' //////////////////////////////////////////////////////////
Public Sub _
UpdateExplosionsFar()
 
 Dim cn As Long
 Dim dx As Integer, dy As Integer                   ' explosion blitting destination
 Dim rExp As RECT
 Dim bUpdateFrame As Boolean                        ' this tells the loop if anim. frame should be increased
 Dim MaxFrames   As Byte
 Dim lTicks      As Long
 Static lFPSTime As Long                            ' update frame timer
   
   lTicks = GetTicks()
   If lFPSTime < lTicks Then
      lFPSTime = lTicks + FPS_ANIMS
      bUpdateFrame = True
   End If
 
 For cn = 0 To MAX_EXPLOSIONS
  With g_Explosion(cn)
   If .Visible And ( _
      .kind = ET_BIG_FAR Or _
      .kind = ET_SMALL_FAR Or _
      .kind = ET_SMALLBLUE_FAR) Then
    
    If .kind = ET_BIG_FAR Then
     Call SetRect(rExp, 0, 0, 80, 60)
     dx = (.x - rExp.Right / 2) - wx / .z
     dy = (.y - rExp.Bottom / 2) - wy / .z
     Call BltFastGFX_HBM(dx, dy, g_Objects.Exp1Far(.frame))
     MaxFrames = UBound(g_Objects.Exp1Far)
     
     ' blit small explosion
    ElseIf .kind = ET_SMALL_FAR Then
     Call SetRect(rExp, 0, 0, 40, 30)
     dx = (.x - rExp.Right / 2) - wx / .z
     dy = (.y - rExp.Bottom / 2) - wy / .z
     'Debug.Print "Laser n:" & cn & " blit x: " & dx & " y: " & dy
     
     If (Not mDirectDraw.bHardwareRasters) Then
      Call BltFastGFX_HBM(dx, dy, g_Objects.Exp2Far(.frame))
     Else ' blit pseudo-transparent
      'BltFX dx, dy, g_Objects.Exp2Far(.frame), rExp, R_OR, True
     End If
     MaxFrames = UBound(g_Objects.Exp2Far)
     
     ' blit small blue explosion
    ElseIf .kind = ET_SMALLBLUE_FAR Then
     Call SetRect(rExp, 0, 0, 40, 30)
     dx = (.x - rExp.Right / 2) - wx / .z
     dy = (.y - rExp.Bottom / 2) - wy / .z
     If (Not mDirectDraw.bHardwareRasters) Then
      Call BltFastGFX_HBM(dx, dy, g_Objects.Exp3Far(.frame))
     Else ' blit pseudo-transparent
      'BltFX dx, dy, g_Objects.Exp3Far(.frame), rExp, R_OR, True
     End If
     MaxFrames = UBound(g_Objects.Exp3Far)
    End If
 
     If bUpdateFrame Then                               ' see if anim.frame should be updated
      .frame = .frame + 1
      If .frame > MaxFrames Then
       .Visible = False
      End If
     End If
    
   End If
  End With
 Next

End Sub

' //////////////////////////////////////////////////////////
' //// Update all meteors created
' //////////////////////////////////////////////////////////
Public Sub _
UpdateMeteors()
 
 Static lFPSTime  As Long
 Dim bUpdateFrame As Boolean
 Dim cn           As Long
 Dim i            As Long
 'Dim rMeteor      As RECT ', rTarget As RECT
 'Dim bPutShadow   As Boolean
 Dim dx           As Integer
 Dim dy           As Integer
 Dim lTicks       As Long
 'Dim lExperience  As Long
 
 lTicks = GetTicks()
 If (lFPSTime + FPS_ANIMS) < lTicks Then
  lFPSTime = lTicks
  bUpdateFrame = True
 End If
 
 For cn = 0 To MAX_METEORS
  With g_Meteor(cn)
   If (.Visible) Then
   
     .x = .x + .xVel
     .y = .y + .yVel
     
     'Debug.Print "Y is:"; .y
     
     If InRange(CInt(.x), .dx, .dx + 10) And _
        InRange(CInt(.y), .dy, .dy + 10) Then
      Call CreateExplosion(CInt(.x + .cx \ 2), CInt(.y + .cy \ 2), .z, ET_BIG + 3 * (.z - 1))
      
      If (.Data And MC_HITMOON) Then
       
       Call DoMoonQuake(True, 3000, 4)
       ' check for close bunker positions and
        For i = 0 To MAX_BUNKERS
         'If InRange((.x + 16) \ 2, cBunker(i).GetX - 50, cBunker(i).GetX + 50) Then
         If (nGetDist2D(CInt(.x), CInt(.y), _
             CBunker(i).GetX, CBunker(i).GetY)) < 100 Then
          CBunker(i).DoDamage = 1000
         End If
        Next
      
      ElseIf (.Data And MC_HITEARTH) Then
        
        g_hpEarth = g_hpEarth - 25 '{!}
        '...
      End If
       .Visible = False                             ' meteor has died :)
     End If
             
     ' see if animation frame should be updated
     If (bUpdateFrame) Then
      .frame = .frame + 1
      If .frame > UBound(g_Objects.Meteor1) Then .frame = 0
     End If
 
   
     ' blit to backbuffer
     If (.Data And MC_CLOSE) Then
      
      ' check if it's close enough to put shadow
      If (.y + .cy > 370) Then
       'bPutShadow = True
       .putshad = True
       .sy = .sy - .yVel '(.y / .dy) * .yVel
       If .sy < 0 Then .sy = 0
      End If
      
      'Call SetRect(rMeteor, 0, 0, .cx, .cy)
      If (.putshad) Then _
      Call BltFastGFX_HBM(.x - wx / .z, (.y + .sy) - wy / .z, g_Objects.Meteor1_Shadow(.frame))
      'BltFX .x - wx / .z, (.y + .sy) - wy / .z, g_Objects.Meteor1_Shadow(.Frame), rMeteor, R_AND, True
      Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.Meteor1(.frame))
      
     ElseIf (.Data And MC_FAR) Then
      'Call SetRect(rMeteor, 0, 0, .cx, .cy)
      Call BltFastGFX_HBM(.x - wx / .z, .y - wy / .z, g_Objects.Meteor2(.frame))
     End If
      
     ' check hitpoints
     If (.HP <= 0) Then
      'g_Player.score = g_Player.score + 96
      .Visible = False
      Call CreateExplosion(CInt(.x + .cx \ 2), CInt(.y + .cy \ 2), .z, ET_BIG + 3 * (.z - 1))
      
      '...PLAYSOUND
      If (.Data And MC_FAR) Then
       Call DSPlaySound(g_dsSfx(SFX_METEORBLAST), False, (.x - wx), SFX_VOLUMEFAR)
      ElseIf (.Data And MC_CLOSE) Then
       Call DSPlaySound(g_dsSfx(SFX_METEORBLAST), False, (.x - wx))
      End If
     End If
      
   End If
  End With
 Next
 
End Sub

' Calculations and Creations
' ---------------------------------------------------------------

Public Sub UpdateFontAnimation(eState As enumFontAnimationStates)
 
 Dim cn As Integer                                       ' local counter
 Static lFPSTime As Long
 Dim bUpdateFrame As Boolean
 Dim hTempDC As Long
 
 'If lFPSTime < GetTicks Then                       ' determine should animations be updated
 '   lFPSTime = GetTicks + FPS_ANIMS
    bUpdateFrame = True
 'End If
 
 If eState = FA_INTRO Then                               ' do intro font animations
  
  For cn = 0 To UBound(g_Credits)
   With g_Credits(cn)
    If .Visible Then                                    ' if it's been created
     If bUpdateFrame Then                               ' if calcs. should be updated
      Select Case .Heading
        Case 0                                          ' North-West
          .x = .x - .xVel
          .y = .y - .yVel
        Case 1                                          ' South-West
          .x = .x - .xVel
          .y = .y + .yVel
        Case 2                                          ' South-East
          .x = .x + .xVel
          .y = .y + .yVel
        Case 3                                          ' North-East
          .x = .x + .xVel
          .y = .y - .yVel
       Case Else: Debug.Print .lpszAuthor & " font heading failed!"
     End Select
         ' do color fading
         .cr = .cr + nGetRnd(2, 3)
         '.r = .r + nGetRnd(2, 3)
         '.g = .g + nGetRnd(2, 3)
         '.b = .b + nGetRnd(2, 3)
         'If .r > 242 Then .r = 242
         'If .g > 242 Then .g = 242
         'If .b > 242 Then .b = 242
         If .cr > 242 Then .cr = 242
         ' do font enlargement
         '.nReserved = .nReserved + 1
          .fs = .fs + 1
          If .fs > 18 Then .fs = 18
         ' do angle rotation
         .ang = .ang + .nReserved
         If .ang > 359 Then .ang = 0
         ' do clipping checking
         If .x < 0 Or .x > MAX_CX Or _
            .y < 0 Or .y > MAX_CY Then
            .Visible = False
            Call CreateAuthor(arCreditsList(nGetRnd(0, 7)))
         End If
     
     End If
         ' finally blit the text
         lpBack.restore
         hTempDC = lpBack.GetDC()
          Call RotateText(hTempDC, .x, .y, _
                          .lpszAuthor, .fs, .ang, RGB(.cr, .cr, .cr))
         lpBack.ReleaseDC hTempDC
         ' ok! we're done ;)
    End If
   End With
  Next
 
 ElseIf eState = FA_MAINMENU Then
  ' ...
  
 End If

End Sub

' //////////////////////////////////////////////////////////
' //// Create partivle remover object
' //////////////////////////////////////////////////////////
Public Sub _
UpdateParticleRemover()

 Dim i      As Long
 Dim j      As Long
 Dim rShip  As RECT
 Dim rPR    As RECT
 Dim nDist1 As Integer
 Dim nDist2 As Integer
 
 
 For i = 0 To PREMOVER_MAX
  
  If (g_PRemover(i).Visible) Then
   With g_PRemover(i)
    
    ' advance to hit target
    If (CShip(.Heading).GetX > .x) Then
     .x = .x + .xVel
    ElseIf (CShip(.Heading).GetX < .x) Then
     .x = .x - .xVel
    End If
    If (CShip(.Heading).GetY > .y) Then
     .y = .y + .yVel
    ElseIf (CShip(.Heading).GetY < .y) Then
     .y = .y - .yVel
    End If
       
    ' setup dest rects
    Call SetRect(rPR, .x, .y, .x + 5, .y + 5)
    ' check 4 interpolation
    If (Collide(rPR, CShip(.Heading).GetRect)) Then
     ' destroy ship ( 100 should do it )
     CShip(.Heading).DoDamage = 100
     ' create partivle explosion
     Call CreateParticleExplosion(CInt(.x), CInt(.y))
          
     ' get new target
     If (.frame < 3) Then
      
      .Heading = 255
      nDist1 = MAX_INT
   
      For j = 0 To MAX_ENEMIES
       ' get 1st target
       If (CShip(j).GetVisible And CShip(j).GetZ = PLANE_FAR) Then
        
        ' get enemy distance
        nDist2 = nGetDist2D(CShip(j).GetX, CShip(j).GetY, CInt(g_PRemover(i).x), CInt(g_PRemover(i).y))
        ' check if it's lower than the current closest target
        If (nDist2 < nDist1) Then
         nDist1 = nDist2
         g_PRemover(i).Heading = j
        End If
      
       End If
      Next
     
      If (g_PRemover(i).Heading = 255) Then
       .Visible = False
      Else
       .Visible = True
       .frame = .frame + 1
      End If
     
     Else ' 3 ships down
      
      .Visible = False
     End If
     
     '...
    End If
    
    ' blit it
    Call BltFastGFX_HBM(.x - wx / PLANE_FAR, .y - wy / PLANE_FAR, g_Objects.pr)
    
   End With
  End If
  
 Next
 

End Sub

' //////////////////////////////////////////////////////////
' //// Create partivle remover object
' //////////////////////////////////////////////////////////
Public Sub _
CreateParticleRemover(x As Integer, y As Integer)

 Dim i      As Long
 Dim j      As Long
 Dim nDist1 As Integer
 Dim nDist2 As Integer
  
 ' search for free object
 Do While (i < PREMOVER_MAX)
 
  ' if found -> create
  If (Not g_PRemover(i).Visible) Then
  
   g_PRemover(i).x = CSng(x)
   g_PRemover(i).y = CSng(y)
   g_PRemover(i).xVel = 5.8
   g_PRemover(i).yVel = 5.8
   g_PRemover(i).frame = 1             ' used as a targets hit counter
         
   g_PRemover(i).Heading = 255
   nDist1 = MAX_INT
   
   For j = 0 To MAX_ENEMIES
    ' get 1st target
    If (CShip(j).GetVisible And CShip(j).GetZ = PLANE_FAR) Then
     ' get enemy distance
     nDist2 = nGetDist2D(CShip(j).GetX, CShip(j).GetY, CInt(g_PRemover(i).x), CInt(g_PRemover(i).y))
     ' check if it's lower than the current closest target
     If (nDist2 < nDist1) Then
      nDist1 = nDist2
      g_PRemover(i).Heading = j
     End If
     
    End If
   Next
   
   ' if there's no ship visible then don't create premover object
   If (g_PRemover(i).Heading = 255) Then
    g_PRemover(i).Visible = False
   Else
    g_PRemover(i).Visible = True
   End If
   
   Exit Do
  End If
 
  i = i + 1
 Loop

End Sub


' //////////////////////////////////////////////////////////
' //// Update all bonus objects (effects and anims.)
' //////////////////////////////////////////////////////////
Public Sub _
UpdateBonus()
 
 Dim cn           As Integer
 Dim dx           As Integer
 Dim dy           As Integer
 Dim bUpdateFrame As Boolean
 Dim lTicks       As Long
 Static lFPSTime  As Long
  
 lTicks = GetTicks()
 ' animation 4 update check
 If (lFPSTime + FPS_ANIMS < lTicks) Then
  lFPSTime = lTicks
  bUpdateFrame = True
 End If
 
 
 For cn = 0 To BONUS_MAX
 
  If (g_Bonus(cn).state = BSTATE_FLOATING) Then
      
   ' make bonus float in space
   g_Bonus(cn).floatb = g_Bonus(cn).floatb + g_Bonus(cn).floata
   If (g_Bonus(cn).floatb > 2) Then g_Bonus(cn).floata = -g_Bonus(cn).floata
   If (g_Bonus(cn).floatb < -2) Then g_Bonus(cn).floata = Abs(g_Bonus(cn).floata)
   g_Bonus(cn).y = g_Bonus(cn).y + g_Bonus(cn).floata
  
   ' animation loop
   If (bUpdateFrame) Then
    g_Bonus(cn).frame = g_Bonus(cn).frame + 1
    If (g_Bonus(cn).frame > 6) Then g_Bonus(cn).frame = 0
   End If
   
   ' blit the bonus gr
   dx = g_Bonus(cn).x - wx / PLANE_FAR
   dy = g_Bonus(cn).y - wy / PLANE_FAR
   
   Call BltFastGFX_HBM(dx, dy, _
                       g_Objects.Bonus(g_Bonus(cn).kind, g_Bonus(cn).frame))
  
  ElseIf (g_Bonus(cn).state = BSTATE_ACTIVE) Then
  
   ' see if bonus-effect time has passed
   If (g_Bonus(cn).fxTime < GetTicks) Then
    ' deactivate bonus
    g_Bonus(cn).state = BSTATE_INACTIVE
    
    Select Case g_Bonus(cn).kind
         
      ' disable megadamage
      Case BONUS_MEGADAMAGE
       g_PlayerDmgBonus = 0
       
      ' disable rapid-fire
      Case BONUS_RAPIDFIRE
       g_PlayerFireDelay = 0
       
    End Select
   End If
  
  End If '// bonus check
   
 Next

End Sub

' //////////////////////////////////////////////////////////
' //// Give player bonus
' ////  objBonus - bonus type
' //////////////////////////////////////////////////////////
Public Sub _
GiveBonus(objBonus As stBonus)
 
 Dim cn As Long
 Dim i  As Long, j As Long
 
 Select Case objBonus.kind
   
   ' annihilate missiles bonus
   Case BONUS_ANNIHILATE
     
     Call GFXFadeInOut(2)
     ' destroy all ships
     For cn = 0 To MAX_ENEMIES
      If (CShip(cn).GetVisible) Then _
       CShip(cn).DoDamage = 100
     Next
     ' destroy all meteors
     For cn = 0 To MAX_METEORS
       g_Meteor(cn).Visible = False
     Next
     ' kill enemy missiles
     For cn = 0 To MAX_MISSILES
      g_Missile(cn).Visible = False
     Next
     ' kill bonus
     objBonus.state = BSTATE_INACTIVE
         
         
   ' kill all ships in 3x3 range
   Case BONUS_RANGEKILL
     
     Dim rDeathZone As RECT
     Dim rVictim    As RECT
     
     ' set zone of destruction rectangle
     Call SetRect(rDeathZone, objBonus.x - 48, objBonus.y - 48, _
                              objBonus.x + 48, objBonus.y + 48)
     
     For cn = 0 To MAX_ENEMIES
      If (CShip(cn).GetVisible) Then _
       If (Collide(CShip(cn).GetRect, rDeathZone)) Then _
        CShip(cn).DoDamage = 100
     Next
     ' destroy all meteors
     For cn = 0 To MAX_METEORS
      If (g_Meteor(cn).Visible And g_Meteor(cn).z = PLANE_FAR) Then
       
       Call SetRect(rVictim, g_Meteor(cn).x, g_Meteor(cn).y, _
                             g_Meteor(cn).x + g_Meteor(cn).cx, _
                             g_Meteor(cn).y + g_Meteor(cn).cy)
       ' check if it's to be destroyed
       If (Collide(rDeathZone, rVictim)) Then
        ' create explosion and kill the meteor
        Call CreateExplosion(CInt(g_Meteor(cn).x), CInt(g_Meteor(cn).y), g_Meteor(cn).z, ET_SMALL + 3 * (g_Meteor(cn).z - 1))
        g_Meteor(cn).Visible = False
       End If
      End If
     Next
     ' destroy enem missiles
     i = 0
     Do While (i < MAX_MISSILES)
      With g_Missile(i)
       If (.Visible) Then
        Call SetRect(rVictim, .x, .y, .x + 15, .y + 13)
        ' see if those collide
        If Collide(rDeathZone, rVictim) Then
         Call CreateExplosion(CInt(.x), CInt(.y), .z, ET_SMALL + 3 * (.z - 1))
         .Visible = False
         End If
        End If
       End With
       
       i = i + 1
     Loop
     ' kill bonus
     objBonus.state = BSTATE_INACTIVE
        
   ' unleash particle mech
   Case BONUS_PREMOVER
    Call CreateParticleRemover(CInt(objBonus.x), CInt(objBonus.y))
    ' kill bonus
    objBonus.state = BSTATE_INACTIVE
 
   ' decrease shot time
   Case BONUS_RAPIDFIRE
    g_PlayerFireDelay = 5
    ' activate bonus
    objBonus.state = BSTATE_ACTIVE
    objBonus.fxTime = BONUS_DURRAPIDFIRE + GetTicks

   ' resurrect a bunker
   Case BONUS_REVIVEBUNKER
    
    For cn = 0 To MAX_BUNKERS
     If (Not CBunker(cn).GetVisible) Then
      CBunker(cn).SetVisible = True
      '...
      Exit For
     End If
    Next
    ' kill bonus
    objBonus.state = BSTATE_INACTIVE

   ' resurrect a bunker
   Case BONUS_MEGADAMAGE
    g_PlayerDmgBonus = 5
    ' activate bonus
    objBonus.state = BSTATE_ACTIVE
    objBonus.fxTime = BONUS_DURMEGADAMAGE + GetTicks
   
 End Select
 
End Sub


' //////////////////////////////////////////////////////////
' //// Create bonus-object at place
' //// x,y - origin
' //////////////////////////////////////////////////////////
Public Sub _
CreateBonus(x As Integer, y As Integer, _
            Optional bytBonus As Byte = 255)
 
 Dim cn As Integer
 
 ' check for a free bouns struct
 Do While (cn < BONUS_MAX)
  
  If (g_Bonus(cn).state = BSTATE_INACTIVE) Then
   ' setup bonus
   g_Bonus(cn).x = x
   g_Bonus(cn).y = y
   ' set bonus type
   'g_Bonus(cn).kind = BONUS_PREMOVER
   If (bytBonus = 255) Then g_Bonus(cn).kind = nGetRnd(0, 5) _
    Else g_Bonus(cn).kind = bytBonus
   g_Bonus(cn).floata = 0.04          ' float a
   g_Bonus(cn).floatb = 0             ' float b
   g_Bonus(cn).fxTime = 0
   g_Bonus(cn).frame = 0
   g_Bonus(cn).state = BSTATE_FLOATING
   
   Exit Do
  End If
 
  cn = cn + 1
 Loop

End Sub


' //////////////////////////////////////////////////////////
' //// Gets ship's or whatever those coords. are relevant Z
' //////////////////////////////////////////////////////////
Public Function _
Player_GetTargetZ(x As Integer, y As Integer) As Byte
 
 Dim cn As Integer
 Dim rCoords As RECT
 Dim rTarget As RECT
 Dim ar As enumShipType
 
 Call SetRect(rCoords, x, y, x + 5, y + 5)
 
 ' priority 1 - Ships
 For cn = 0 To MAX_ENEMIES
   If CShip(cn).GetVisible Then
    
    If Collide(rCoords, CShip(cn).GetRect) Then
     Player_GetTargetZ = CShip(cn).GetZ
     Exit Function
    End If
  
  End If
 Next
 
 ' priority 2 - Meteors
 For cn = 0 To MAX_METEORS
  With g_Meteor(cn)
   Call SetRect(rTarget, .x, .y, .x + .cx, .y + .cy)
   
   If Collide(rCoords, rTarget) Then
    Player_GetTargetZ = .z
    Exit Function
   End If
  
  End With
 Next
  
 Player_GetTargetZ = PLANE_FAR
End Function

' //////////////////////////////////////////////////////////
' //// Get Ship, meteor, missile or etc. Index at spec.
' //// plane and location
' //////////////////////////////////////////////////////////
Public Sub _
Player_GetTargetIndex(x As Integer, y As Integer, z As Byte, _
                      bytIndex As Byte, TargetKind As enumBunkerTarget)
 Dim cn As Integer
 Dim rCoords As RECT
 Dim rTarget As RECT
 Dim ar As enumShipType
 
 Call SetRect(rCoords, x, y, x + 5, y + 5)
 
 ' priority 1 - Ships
 For cn = 0 To MAX_ENEMIES
   If (CShip(cn).GetVisible And CShip(cn).GetZ = z) Then
    
    If Collide(rCoords, CShip(cn).GetRect) Then
     bytIndex = cn
     TargetKind = BTARGET_SHIP
     Exit Sub
    End If
  
  End If
 Next
 
 ' priority 2 - Meteors
 For cn = 0 To MAX_METEORS
  With g_Meteor(cn)
   Call SetRect(rTarget, .x, .y, .x + .cx, .y + .cy)
   
   If (Collide(rCoords, rTarget) And .z = z) Then
    bytIndex = cn
    TargetKind = BTARGET_METEOR
    Exit Sub
   End If
  
  End With
 Next
  
 ' priority 3 - Missiles
 For cn = 0 To MAX_MISSILES
  With g_Missile(cn)
   Call SetRect(rTarget, .x, .y, .x + 5, .y + 5)
   
   If (Collide(rCoords, rTarget) And .z = z) Then
    bytIndex = cn
    TargetKind = BTARGET_MISSILE
    Exit Sub
   End If
  
  End With
 Next
  
 ' if there was nothing found then set pseudo nothing_found Flag
 bytIndex = 255
End Sub

' //////////////////////////////////////////////////////////
' //// Update fired missiles and lasers
' //// Both're checked in one proc. to save processing time
' //// We have ships, meteors and etc. to check you know,
' //// this ain't no C plus(*2) ;)
' //////////////////////////////////////////////////////////
Public Sub _
Player_UpdateShots()
 
 Dim cn       As Long
 Dim i        As Integer
 Dim rTarget  As RECT
 Dim rMissile As RECT
 Dim ncx As Integer, ncy As Integer
 Dim nDamage  As Integer
 'Dim nDist As Integer ', tcx As Integer, tcy As Integer
 'Dim nFrames As Integer
 
 For cn = 0 To MAX_SHOTS
  With g_Pl_Weapon(cn)
   If (.Visible) Then
    
     ' calculate position
     '.xVel = .xVel - (0.5 * Sgn(.xVel))
     '.yVel = .yVel + 0.5
     'If Abs(.xVel) < 1 Then .xVel = 1 * Sgn(.xVel)
     'If .yVel > -1 Then .yVel = -1
     
     '.xVel = .xVel + (.xVelInc * Sgn(.xVel))
     '.yVel = .yVel - .yVelInc
     'If Abs(.xVel) > PLAYER_MISSILE_SPEED Then .xVel = Sgn(.xVel) * PLAYER_MISSILE_SPEED
     'If .yVel < -PLAYER_MISSILE_SPEED Then .yVel = -PLAYER_MISSILE_SPEED
     '.x = .x + Cos(.Angle) * .xVel
     '.y = .y + Sin(.Angle) * .yVel
     If (.Guided) Then
      ' get target coordinates
      Select Case .TargetKind
       Case BTARGET_SHIP
        If (CShip(.enemyIndex).GetVisible) Then
         .dx = CShip(.enemyIndex).GetX
         .dy = CShip(.enemyIndex).GetY
        End If
        
       Case BTARGET_METEOR
        If (g_Meteor(.enemyIndex).Visible) Then
         .dx = Int(g_Meteor(.enemyIndex).x)
         .dy = Int(g_Meteor(.enemyIndex).y)
        End If
        
       Case BTARGET_SHIP
        If (g_Missile(.enemyIndex).Visible) Then
         .dx = Int(g_Missile(.enemyIndex).x)
         .dy = Int(g_Missile(.enemyIndex).y)
        End If
        
       Case Else
      End Select
      
      ' exec "smartmissile" ai ;-)
      If (.dx > .x) Then
       .x = .x + .xVel
      ElseIf (.dx < .x) Then
       .x = .x - .xVel
      End If
      If (.dy > .y) Then
       .y = .y + .yVel
      ElseIf (.dy < .y) Then
       .y = .y - .yVel
      End If
     
     Else ' do normal travel
      .x = .x + .xVel
      .y = .y + .yVel
     End If
     
    If (.kind <> PW_LASER) Then
     ' do shrinking...
     If (.dDist = BP_FAR) Then
      .frame = .frame + 1
      If .frame > (UBound(g_Objects.GM) - 1) Then .frame = .frame - 1
     ElseIf (.dDist = BP_VERYFAR) Then
      .frame = .frame + 1
      If .frame > UBound(g_Objects.GM) Then .frame = .frame - 1
     End If
     
     ' set object size
     ncx = g_Objects.GM(.frame).cx
     ncy = g_Objects.GM(.frame).cy
     Call SetRect(rMissile, .x, .y, .x + ncx, .y + ncy)
     ' set target rectangle
     Call SetRect(rTarget, .dx, .dy, .dx + 8, .dy + 8)
     
     nDamage = 2 + g_PlayerDmgBonus
    Else
     
     ' Check Laser
     'If bUpdateFrame Then .Frame = .Frame + 1
     .frame = .frame + 1
     If (.z = 2) Then
      If .frame > UBound(g_Objects.LS) Then .frame = .frame - 1
     Else
      If .frame > UBound(g_Objects.LS) - 1 Then .frame = .frame - 1
     End If
     
     ' set object size
     ncx = g_Objects.LS(.frame).cx
     ncy = g_Objects.LS(.frame).cy
     Call SetRect(rMissile, .x, .y, .x + ncx, .y + ncy)
     ' set target rectangle
     Call SetRect(rTarget, .dx, .dy, .dx + 8, .dy + 8)
     
     nDamage = 1 + (g_PlayerDmgBonus / 2) ' for each laser piece
    End If
    
    
    ' --- check if target is hit ---
    If Collide(rTarget, rMissile) Then
     ' ...
     .Visible = False
     
     ' check if it hit a ship
     For i = 0 To MAX_ENEMIES
      If (CShip(i).GetZ = .z And CShip(i).GetVisible) Then
       If Collide(rMissile, CShip(i).GetRect) Then
         
         Call CreateExplosion(CInt(.x), CInt(.y), .z, ET_SMALL + 3 * (.z - 1))
         CShip(i).DoDamage = nDamage
         g_Player.sit = g_Player.sit + 1
        
        Exit For
       End If
      End If
     Next
     
     ' check if it hit a meteor
     For i = 0 To MAX_METEORS
      If (g_Meteor(i).z = .z And g_Meteor(i).Visible) Then
        
        Call SetRect(rTarget, g_Meteor(i).x, g_Meteor(i).y, _
                     g_Meteor(i).x + g_Meteor(i).cx, g_Meteor(i).y + g_Meteor(i).cy)
       
       If (Collide(rMissile, rTarget)) Then
         Call CreateExplosion(CInt(.x), CInt(.y), .z, ET_SMALL + 3 * (.z - 1))
         g_Meteor(i).HP = g_Meteor(i).HP - nDamage
         g_Player.sit = g_Player.sit + 1
         
         ' add score,kills,experience
         If (g_Meteor(i).HP <= 0) Then
          g_Player.kills = g_Player.kills + 1
          g_Player.Exp = g_Player.Exp + g_Meteor(i).Exp
          g_Player.score = g_Player.score + 96
          'g_Player.score = g_Player.score + (g_Meteor(i).HP * 2)
         Else
          g_Player.Exp = g_Player.Exp + 1
          g_Player.score = g_Player.score + 1
         End If
         
        Exit For
       End If
      
      End If
     Next
     
     ' check bonus collision
     For i = 0 To BONUS_MAX
      
      If (g_Bonus(i).state = BSTATE_FLOATING) Then
       
       Call SetRect(rTarget, g_Bonus(i).x, g_Bonus(i).y, _
                            g_Bonus(i).x + 16, g_Bonus(i).y + 16)
       
       If Collide(rMissile, rTarget) Then
         ' explode
         Call CreateExplosion(CInt(.x), CInt(.y), .z, ET_SMALL + 3 * (.z - 1))
         '' kill bonus-object
         'g_Bonus(i).State = BSTATE_ACTIVE
         ' give the real bonus
         Call GiveBonus(g_Bonus(i))
         ' score player (5pts hitting a bonus) + 1 kill
         g_Player.score = g_Player.score + 2
         g_Player.sit = g_Player.sit + 1
         
         Exit For
       
       End If
      End If
     Next
    
    End If
    ' --- end target hit check ---
    
    ' check if this missiles collides with an enemy's one
    If (.kind <> PW_LASER) Then
     i = 0
     Do While i < MAX_MISSILES
      With g_Missile(i)
       If .Visible Then
        ' set enemy missile rect ( assuming it has the smallets size )
        '                        ( another NOT using typeGFX_HBM structure disadvantage!!! (have that in mind ;))
        Call SetRect(rTarget, .x, .y, .x + 15, .y + 13)
        
        ' see if those collide
        If Collide(rTarget, rMissile) Then
         Call CreateExplosion(CInt(.x), CInt(.y), .z, ET_SMALL + 3 * (.z - 1))
         .Visible = False
         g_Player.sit = g_Player.sit + 1
         ' score player (1pt for stopping enemy missile) + 1 kill
         'g_Player.score = g_Player.score + 1
         'g_Player.kills = g_Player.kills + 1
         'g_Player.Exp = g_Player.Exp + 1
        
         Exit Do                                             ' exit now,'cos our missile costs exactly on of the enemy's
        End If
       
       End If
      End With
    
     i = i + 1
     Loop
    End If
    
    ' check if object is out of viewport
     Call SetRect(rTarget, 0, 0, SCREEN_PIXEL_WIDTH, SCREEN_PIXEL_HEIGHT)
     If (Not Collide(rMissile, rTarget)) Then .Visible = False
    
    ' workout blitting position
     ncx = .x - wx \ .z
     ncy = .y - wy \ .z
     If (.x < 0 Or .x > SCREEN_PIXEL_WIDTH Or _
         .y < 0 Or .y > SCREEN_PIXEL_HEIGHT) Then .Visible = False
    
    ' blit it
     If .kind = PW_LASER Then
      Call BltFastGFX_HBM(ncx, ncy, g_Objects.LS(.frame))
     Else
      Call BltFastGFX_HBM(ncx, ncy, g_Objects.GM(.frame))
     End If
     
   End If
  End With
 Next
 
End Sub

' //////////////////////////////////////////////////////////
' //// Update player dual laser shot ( suspended )
' //////////////////////////////////////////////////////////
Public Sub _
Player_UpdateLasers()
 
 Static lTime As Long
 Dim bUpdateFrame As Boolean
 Dim cn As Integer, i As Integer                        ' local counters
 Dim rLaser As RECT, rTarget As RECT
 Dim ncx As Integer, ncy As Integer
 
 If (lTime < GetTicks) Then
  lTime = GetTicks + 10 'FPS_ANIMS
  bUpdateFrame = True
 End If
 
 For cn = 0 To MAX_LASERS
  With g_Pl_Weapon(cn)
   If .Visible Then
    '.x = .x + .xVel
    '.y = .y + .yVel
    .x = .x + .xVel ' Cos(.Angle) * .xVel
    .y = .y + .yVel ' Sin(.Angle) * .yVel
    'Stop
    
     If bUpdateFrame Then .frame = .frame + 1
     If .z = 2 Then
      If .frame > UBound(g_Objects.LS) Then .frame = .frame - 1
     Else
      If .frame > UBound(g_Objects.LS) - 3 Then .frame = .frame - 1
     End If
     
     ' set object size
     ncx = g_Objects.LS(.frame).cx
     ncy = g_Objects.LS(.frame).cy
     Call SetRect(rLaser, .x, .y, .x + ncx, .y + ncy)
     
    Call SetRect(rTarget, .dx, .dy, .dx + 6, .dy + 6)
    
    ' --- check if target is hit ---
    If Collide(rTarget, rLaser) Then
     .Visible = False                                   ' laser reached destination successfully
     Call CreateExplosion(CInt(.x), CInt(.y), .z, ET_SMALL + 3 * (.z - 1))
    
     ' check if it hit a ship
     For i = 0 To MAX_ENEMIES
      If (CShip(i).GetZ = .z And CShip(i).GetVisible) Then
       If Collide(rLaser, CShip(i).GetRect) Then
         CShip(i).DoDamage = 5
        Exit For
       End If
      End If
     Next
     
     ' check if it hit a meteor
     For i = 0 To MAX_METEORS
      If (g_Meteor(i).z = .z And g_Meteor(i).Visible) Then
        
        Call SetRect(rTarget, g_Meteor(i).x, g_Meteor(i).y, _
                     g_Meteor(i).x + g_Meteor(i).cx, _
                     g_Meteor(i).y + g_Meteor(i).cy)
       
       If Collide(rLaser, rTarget) Then
         g_Meteor(i).HP = g_Meteor(i).HP - 5
        Exit For
       End If
      
      End If
     Next
     
    End If
    ' --- end target hit check ---
    
    ' workout blitting position
    ncx = .x - wx \ .z
    ncy = .y - wy \ .z
    If ncx < 0 Or ncx > SCREEN_PIXEL_WIDTH Or _
       ncy < 0 Or ncy > SCREEN_PIXEL_HEIGHT Then .Visible = False
    
    ' blit laser
    Call BltFastGFX_HBM(ncx, ncy, g_Objects.LS(.frame))
   
   End If
  End With
 Next

End Sub

' //////////////////////////////////////////////////////////
' //// Create player dual laser shot
' //////////////////////////////////////////////////////////
Public Sub _
Player_CreateLaser(dx As Integer, dy As Integer, _
                   dcx As Integer, dcy As Integer)

  Dim cn     As Long
  Dim fAngle As Single
  Dim fTempX As Single, fTempY As Single
  Dim nDist  As Integer, nDist2 As Integer
  Dim sz     As Byte

  Do While (cn < MAX_LASERS - 1)
    ' two free laser structures found
   If (Not g_Pl_Weapon(cn).Visible And _
      Not g_Pl_Weapon(cn + 1).Visible) Then
        
    sz = Player_GetTargetZ(dx + wx, dy + wy)
    'sz = 2
    
    ' setup left laser
    With g_Pl_Weapon(cn)
     .kind = PW_LASER                                  ' set wapon kind
     .z = sz                                           ' set laser plane
     '.x = (wx / sz + MAX_CX / 2) - 280
     .x = (wx / sz + MAX_CX / 2) - 220
     .y = MAX_CY - 150
     .dDist = sz                                       ' set destination plane
     .frame = 0                                        ' reset animation frame
     .Visible = True                                   ' visualize laser
     .dx = dx + wx / .z                                ' set destination coordinates
     .dy = dy + wy / .z
     .Guided = False
     .enemyIndex = 255
     'If (.x > .dx) Then
     ' .xVel = -PLAYER_LASER_SPEED                      ' assign default velocity
     'Else
     ' .xVel = PLAYER_LASER_SPEED
     'End If
     '.yVel = -PLAYER_LASER_SPEED
     fAngle = GetAngle(.x, .y, .dx, .dy)
     .xVel = PLAYER_LASER_SPEED + g_PlayerFireDelay
     .yVel = PLAYER_LASER_SPEED + g_PlayerFireDelay
     .xVel = Cos(fAngle) * .xVel                       ' precalculate velocities ( for speed )
     .yVel = Sin(fAngle) * .yVel
     nDist = nGetDist2D(CInt(.x), CInt(.y), .dx, .dy)  ' get distance to destination
    End With
    
    With g_Pl_Weapon(cn + 1)
     .kind = PW_LASER
     .z = sz
     .x = (wx / sz + MAX_CX / 2) + 220 - 90
     .y = MAX_CY - 150
     .dDist = sz
     .frame = 0
     .Visible = True
     .dx = dx + wx / .z
     .dy = dy + wy / .z
     .Guided = False
     .enemyIndex = 255
     'If (.x > .dx) Then
     ' .xVel = -PLAYER_LASER_SPEED                       ' assign default velocity
     'Else
     ' .xVel = PLAYER_LASER_SPEED
     'End If
     '.yVel = -PLAYER_LASER_SPEED
     .xVel = PLAYER_LASER_SPEED + g_PlayerFireDelay
     .yVel = PLAYER_LASER_SPEED + g_PlayerFireDelay
     fAngle = GetAngle(.x, .y, .dx, .dy)
     .xVel = Cos(fAngle) * .xVel                        ' precalculate velocities
     .yVel = Sin(fAngle) * .yVel
     nDist2 = nGetDist2D(CInt(.x), CInt(.y), .dx, .dy)
    End With
    
    ' balance lasers
    If nDist > nDist2 Then
     fTempX = nDist / nDist2
     'g_Pl_Weapon(cn).xVel = g_Pl_Weapon(cn).xVel * fTempX
     'g_Pl_Weapon(cn).yVel = g_Pl_Weapon(cn).yVel * fTempX
     g_Pl_Weapon(cn + 1).xVel = g_Pl_Weapon(cn + 1).xVel / fTempX
     g_Pl_Weapon(cn + 1).yVel = g_Pl_Weapon(cn + 1).yVel / fTempX
    ElseIf nDist2 > nDist Then
     fTempY = nDist2 / nDist
     g_Pl_Weapon(cn).xVel = g_Pl_Weapon(cn).xVel / fTempY
     g_Pl_Weapon(cn).yVel = g_Pl_Weapon(cn).yVel / fTempY
     'g_Pl_Weapon(cn + 1).xVel = g_Pl_Weapon(cn + 1).xVel * fTempY
     'g_Pl_Weapon(cn + 1).yVel = g_Pl_Weapon(cn + 1).yVel * fTempY
    End If
    
    ' increment player shot var
    g_Player.ts = g_Player.ts + 1
    
    Exit Do                                            ' dual laser shot created, so we exit
   End If
   cn = cn + 1
  Loop
  
End Sub
                   
' //////////////////////////////////////////////////////////
' //// Create player missile shot
' //////////////////////////////////////////////////////////
Public Sub _
Player_CreateMissile(dx As Integer, dy As Integer, _
                     dcx As Integer, dcy As Integer)
 
 If (dy > TARGET_AREA) Then Exit Sub
 
 Dim cn     As Long                              ' local counter
 Dim fAngle As Single
 Dim fTempX As Single, fTempY As Single

  Do While cn < MAX_MISSILES
   With g_Pl_Weapon(cn)
    If Not .Visible Then                         ' create missile
      
     '.Kind = PW_MISSILE_LONGRANGE
     .kind = PlayerWeapon
     'If .kind = PW_MISSILE_CLOSERANGE Then .z = 1 Else .z = 2
     .z = Player_GetTargetZ(dx + wx, dy + wy) '{!}
     If (.z = 1) Then PlayerWeapon = PW_MISSILE_CLOSERANGE _
      Else PlayerWeapon = PW_MISSILE_LONGRANGE
      
      .x = (wx / .z + MAX_CX / 2)
      .y = MAX_CY - 125
      '.z = sz                                         ' assign blitting plane
      .dDist = .z                                     ' set destination plane
      .frame = 0                                      ' reset animation frame
      .Visible = True                                 ' visualize
      
      ' fill target info
      .dx = dx + wx / .z
      .dy = dy + wy / .z
      
      ' fill velocity info
      'If Max(.dx, CInt(.x)) Then
      ' .xVel = 1 ' 3 + nTemp
      'Else
      ' .xVel = -1 ' -3 - nTemp
      'End If
      
      ' calculate velocity boundaries
      'fTempX = 1
      'fTempY = 1
      'Call CalcVelocityBound(.x, .y, CSng(.dx), CSng(.dy), fTempX, fTempY)
      '.xVel = .xVel * PLAYER_MISSILE_SPEED                   ' assign default velocity
      '.yVel = -PLAYER_MISSILE_SPEED
      .xVel = PLAYER_MISSILE_SPEED + g_PlayerFireDelay
      .yVel = PLAYER_MISSILE_SPEED + g_PlayerFireDelay
      '.xVel = .xVel / fTempX                                ' calculate frictions
      '.yVel = .yVel / fTempY
      Call Player_GetTargetIndex(.dx, .dy, .z, .enemyIndex, .TargetKind)
      
      If (.enemyIndex = 255) Then                            ' no target found (guided missile off)
       .Guided = False
       fAngle = GetAngle(.x, .y, .dx, .dy)
       .xVel = .xVel * Cos(fAngle)                           ' precalculate velocity vectors
       .yVel = .yVel * Sin(fAngle)
      Else
       .Guided = True
      End If
      
      ' increment player shot var
      g_Player.ts = g_Player.ts + 1
 
       'If bLaserLeft Then
       ' Call Player_CreateMissile(dx, dy, dcx, dcy, Not bLaserLeft, nDist2, nLR_Index)
        ' izchisli kolko po-golqmo e rastoqnoeto za 1niq laser do Dest-a, ot towa na 2riq
        'bCalc = True
       '
       ' If nDist = 0 Then
       '  nDist = 1
       '  bCalc = Not bCalc
       ' ElseIf nDist2 = 0 Then
       '  nDist2 = 1
       '  bCalc = Not bCalc
       ' End If
       '  If (nDist > nDist2) And bCalc Then
       '   'fTempX = (nDist / (nDist2 + 1)) * 1
       '   fTempX = nDist / nDist2
       '   .veladd = fTempX
       '  Else
       '   fTempX = nDist2 / nDist
       '   g_Pl_Weapon(nLR_Index).veladd = fTempX
       '  End If
       '  .xVel = -.xVel
       'Else
       ' nIndex = cn
       ' nLLDist = nDist
       'End If
      'End If ' end bLaser IF
         
     
     Exit Do
    End If   ' end Visible IF
   End With
 
  cn = cn + 1
  Loop
End Sub

' //////////////////////////////////////////////////////////
' //// Create a meteor to hit Earth or Moon
' //////////////////////////////////////////////////////////
Public Sub _
CreateMeteor(mData As enumMeteorConsts)
 
 Dim cn     As Long
 Dim Sign   As Integer
 Dim fTempX As Single, fTempY As Single
 Dim fAngle As Single
 
 Do While (cn < MAX_METEORS)
  With g_Meteor(cn)
   If Not .Visible Then
    
     .putshad = False
     ' determine meteor appear
     If mData And MC_RIGHT Then
      .x = nGetRnd(SCREEN_PIXEL_WIDTH + 10, SCREEN_PIXEL_WIDTH + 100)
      .x = SCREEN_PIXEL_WIDTH - 200
      Sign = -1
     ElseIf mData And MC_LEFT Then
      .x = fGetRnd(-300, -100)
      '.x = 150
      Sign = 1
     End If
    
      .y = nGetRnd(100, 150) ' VISIBLE_AREA_CY - 100)
      
     ' determine distance
     If (mData And MC_CLOSE) Then
      '.xVel = Sign * fGetRnd(0.7, 1)
      .HP = 34
      .Exp = 15                                    ' assign experience
      .xVel = fGetRnd(0.7, 1)
      .yVel = Abs(.xVel)
      .z = 1
      .cx = 40
      .cy = 40
      .sy = 70
     ElseIf (mData And MC_FAR) Then
      .HP = 40
      .Exp = 30
      .z = 2
      '.xVel = Sign * fGetRnd(0.4, 0.6)
      .xVel = fGetRnd(0.4, 0.6)
      .yVel = Abs(.xVel)
      .cx = 16
      .cy = 16
     End If
      
      If mData And MC_HITEARTH Then             ' calculate earth-hit vector
       .dx = g_xEarth + (g_cxEarth / 2) + nGetRnd(-50, 50)
       .dy = g_yEarth + (g_cyEarth / 2) + nGetRnd(-50, 50)
      
      ElseIf mData And MC_HITMOON Then          ' calculate moon-hit vector
       
       If mData And MC_LEFT Then
        .dx = nGetRnd(50, SCREEN_PIXEL_WIDTH - 50)
        .dy = nGetRnd(VISIBLE_AREA_CY + 45, VISIBLE_AREA_CY + 100)
       ElseIf mData And MC_RIGHT Then
        .dx = nGetRnd(50, SCREEN_PIXEL_WIDTH - 50)
        .dy = nGetRnd(VISIBLE_AREA_CY + 45, VISIBLE_AREA_CY + 100)
       End If
       
      End If
      
      'Call CalcVelocityBound(.x, .y, CSng(.dx), CSng(.dy), fTempX, fTempY)
      '.xVel = .xVel / fTempX
      '.yVel = .yVel / fTempY
      fAngle = GetAngle(.x, .y, .dx, .dy)
      .xVel = Cos(fAngle) * .xVel
      .yVel = Sin(fAngle) * .yVel
      
      'Debug.Print "ftempX is " & fTempX
      'Debug.Print "ftempY is " & fTempY
      'Debug.Print .xVel
      'Debug.Print .yVel
     '.HP = 100 ' .......
     .Data = mData
     .frame = 0
     .Visible = True
    Exit Do
   End If
  End With
  
  cn = cn + 1
 Loop

End Sub

Public Sub CreateExplosion(x As Integer, y As Integer, z As Byte, eExplosion As enumExplosionType)
 ' Desc: Create an explosion at place
 Dim cn As Integer
 
 Do While cn < MAX_EXPLOSIONS
  With g_Explosion(cn)
   If Not .Visible Then
    .x = x
    .y = y
    .z = z
    .kind = eExplosion
    
    ' play explosion sound
    'If (.kind > ET_SMALL) Then
     Call DSPlaySound(g_dsSfx(SFX_FAREXPLOSION1), False, (.x - wx))
    'End If
    
    .Visible = True
    .frame = 0
    'Exit Do
    Exit Sub
   End If
  End With
 
 
 cn = cn + 1
 Loop
 
 'Debug.Print "No exp. available!"
 
End Sub


Public Sub _
UpdateParticleLasers()

 Dim cn As Long
 Static lFPSTime As Long                            ' update frame timer
 Dim bUpdateFrame As Boolean                        ' this tells the loop if anim. frame should be increased
 Dim lTicks As Long
 Dim dx As Long, dy As Long
 
   lTicks = GetTicks()
   If lFPSTime < lTicks Then
      lFPSTime = lTicks + FPS_ANIMS
      bUpdateFrame = True
   End If

 
 For cn = 0 To MAX_LASERS
 
  With g_PLaser(cn)
   If (.Visible) Then
       
    ' render
    dx = .x - wx / PLANE_FAR
    dy = .y - wy / PLANE_FAR
    mGFX.BltFastGFX_HBM dx, dy, g_Objects.bclaz(.frame)
    
    If (bUpdateFrame) Then
     .frame = .frame + 1
     If (.frame >= 8) Then
      ' end animation and set ship to attack
      CShip(.Possession).DoEnemyDamage
      .Visible = False
     End If
    End If
       
       
   End If
  
  End With
 
 Next
 
 
End Sub


' desc: particle laser creation
Public Sub _
CreateParticleLaser(sx As Integer, sy As Integer, _
                     Optional bytPossession As Byte = NO_POSSESSION)
                     
                     
 Dim cn As Long
 
                      
 Do While (cn < MAX_LASERS)
  
  With g_PLaser(cn)
   If (Not .Visible) Then
    .x = sx
    .y = sy
    .frame = 0
    .Possession = bytPossession
    .Visible = True
    
    Exit Do
   End If
  End With
  
  cn = cn + 1
 Loop
                     
                     
End Sub


Public Sub CreateLaserCut(sx As Integer, sy As Integer, sz As Byte, _
                          dx As Integer, dy As Integer, _
                          dcx As Integer, dcy As Integer, _
                          ekind As enumWeapon, _
                          Optional bytPossession As Byte = NO_POSSESSION)
 
 ' Desc: Inits a cutbeam weapon
 Dim cn As Long
 Dim v1 As Single, v2 As Single
 Dim fAngle As Single
 
 Do While cn < MAX_LASERS
  With g_LaserCut(cn)
   If Not .Visible Then
     .kind = ekind
     .x = sx
     .y = sy
     .z = sz                                          ' set laser plane
     .dx = dx
     .dy = dy
     .dcx = dcx
     .dcy = dcy
     If .kind = SW_LASER Then
      .xVel = 9.5
      .yVel = 9.5
     ElseIf .kind = SW_GREENLASER Then
      .xVel = 6
      .yVel = 6
     End If
     'v1 = 1
     'v2 = 1
     .Possession = bytPossession
     .frame = 0
     ' calculate velocity vectors
     'Call CalcVelocityBound(.x, .y, CSng(.dx), CSng(.dy), v1, v2)
     '.xVel = .xVel / v1
     '.yVel = .yVel / v2
     fAngle = GetAngle(.x, .y, .dx, .dy)
     .xVel = Cos(fAngle) * .xVel
     .yVel = Sin(fAngle) * .yVel
     
     ' set laser face and x_velocity sign
     If max(.x, .dx) Then
      .Direction = SO_LEFT
      '.xVel = -.xVel
       ' set face
      ' If Max(.y, .dy) Then
      '  '.Direction = SO_LEFTUP
      '  .yVel = -.yVel
      ' Else
        '.Direction = SO_LEFTDOWN
      ' End If
     Else
       ' set face
       .Direction = SO_RIGHT
      ' If Max(.y, .dy) Then
      '  '.Direction = SO_RIGHTUP
      '  .yVel = -.yVel
      ' Else
      '  '.Direction = SO_RIGHTDOWN
      ' End If
     End If
     
     .Visible = True
     
     Exit Do
   End If
  End With
 cn = cn + 1
 Loop

End Sub

Public Sub CreateMissile(sx As Integer, sy As Integer, sz As Byte, _
                         dx As Integer, dy As Integer, _
                         dcx As Integer, dcy As Integer, _
                         od As enumBunkerPosition, _
                         Optional bytPossession As Byte = NO_POSSESSION)
 
 ' Desc: Creates and inits a missile
 Dim cn As Long                                  ' local counter
 Dim v1 As Single, v2 As Single
 
 Do While cn < MAX_MISSILES
  With g_Missile(cn)
   If Not .Visible Then
     
     .x = sx
     .y = sy
     .z = sz
     .xVel = 0#
     .yVel = 0#
     .Distance = BP_VERYFAR
     .ObjectiveDist = od
     .Possession = bytPossession
     .dx = dx
     .dy = dy
     .dcx = dcx
     .dcy = dcy
     .xVelB = 1
     .yVelB = 1
     .nTempVar = nGetDist2D(CInt(.x), CInt(.y), .dx, .dy) \ 3
     ' upper value holds the distance a missile must travel
     ' before its' animation gets close(+f) to user's screen
     'v1 = Abs(.x - .dx)
     'v2 = Abs(.y - .dy)
     'If v1 = 0 Then v1 = 1
     'If v2 = 0 Then v2 = 1
     'If v1 > v2 Then
     '   .yVelB = v1 / v2 '(v2 + 0.001)
     'ElseIf v2 > v1 Then
     '   .xVelB = v2 / v1 '(v1 + 0.001)
     'End If
     Call CalcVelocityBound(.x, .y, .dx, .dy, .xVelB, .yVelB)
     
     .frame = 0
     If max(.x, .dx) Then                             ' setup blitting direction
        .Direction = SO_LEFT
     Else
        .Direction = SO_RIGHT
     End If
     .Visible = True                                  ' missile is ready to go
   
     Exit Do                                          ' missile's been created, so we exit!
   End If
  
  End With
 cn = cn + 1
 Loop
 
End Sub

Public Sub CreateChillingPixel(x As Integer, y As Integer, z As Integer)
 Dim cn As Integer                                   ' local counter
 
 Do While cn < MAX_CPIXELS
  With g_CPixel(cn)
   
   If Not .Visible Then
    .Visible = True
    .x = x
    .y = y
    .z = CByte(z)
    .xVel = fGetRnd(-0.3, 0.3)
    .yVel = fGetRnd(-0.3, 0.3)
    .frame = 0
    '.cc = 255
    '.r = 25
    '.g = 25
    '.b = 250
    '.cf = 0
    .Life = 80 ' nGetRnd(190, 250)
    Exit Do
   End If
  
  End With
 cn = cn + 1
 Loop
 
End Sub

Public Sub CreateParticleExplosion(x As Integer, y As Integer)
 
 Dim i As Long, j As Long                            ' local counter
 
  Do While (j < MAX_EXPLOSIONS)
   With g_PExp(j)
      
     If .Visible = False Then
      .Visible = True
      
      For i = 0 To MAX_PARTICLES
       With .Particle(i)
        .x = x
        .y = y
        .xVel = fGetRnd(0.1, 0.6)
        .yVel = fGetRnd(0.01, 0.2)
        .yFriction = fGetRnd(0.01, 0.07)
        .Heading = CByte(nGetRnd(0, 3))
        '.r = 0
        '.g = 0
        '.b = 0
        '.cf = 6
        .frame = 0
        .Visible = True
        .Life = nGetRnd(140, 180)
       End With
      Next
      ' PLAY SOUND
      Call DSPlaySound(g_dsSfx(SFX_PARTICLEXPLOSION), False, (x - wx))
      
      Exit Do
     End If
   
   End With
 j = j + 1
 Loop

End Sub

Public Function CreateWarpGate(x As Integer, y As Integer, z As Byte) As Long
 ' Desc: Creates an enemy WarpGate at place
 Dim cn As Long
                                                    ' create new warpgate
 For cn = 0 To MAX_WARPGATES
  With g_WarpGate(cn)
   If Not .Visible Then
     
     .x = x - 60                                    ' make middle placement calcs.
     .y = y - 40
     .z = z
     .frame = 0
     .bPlayBack = False
     .Visible = True
      ' PLAY WARP SOUND
      Call DSPlaySound(g_dsSfx(SFX_WARPGATE), False, (.x - wx), SFX_VOLUMEFAR)
      CreateWarpGate = cn                           ' set return index
     Exit For                                       ' gate's created now exit
   
   End If
  End With
 Next

End Function


Public Sub SetupStars()                             ' setup some stars ( 4d viewer's pleasure )
 Dim cn As Integer                                  ' local counter
 
 For cn = 0 To MAX_STARS
   With Star(cn)
     '.x = nGetRnd(-640, SCREEN_PIXEL_WIDTH)
     .x = nGetRnd(0, SCREEN_PIXEL_WIDTH)
     .y = nGetRnd(0, VISIBLE_AREA_CY)
     .z = nGetRnd(3, 6)
     .frame = nGetRnd(0, 148)
   End With
 Next
 
End Sub


Public Sub SetupBunkers()
 ' Desc: Setup All Bunkers' classes and positions
 
 Dim i        As Long, j      As Long, cn    As Long
 Dim nTemp1   As Long, nTemp2 As Long, nNext As Integer
 Dim cTBunker As New clsBunker
  
  For cn = 0 To MAX_BUNKERS
   Set CBunker(cn) = New clsBunker
  Next
  nTemp1 = MAX_BUNKERS / 3
  nTemp2 = MAX_BUNKERS / 2
  ' 180,100,60
  nNext = 300
  For cn = 0 To nTemp1
   CBunker(cn).CreateBunker nGetRnd(nNext - 100, nNext), _
                            nGetRnd(329, 359), BP_CLOSE, BT_WEAK
  ' cBunker(cn).CreateBunker nGetRnd(100, SCREEN_PIXEL_WIDTH - 200), _
                            nGetRnd(329, 359), BP_CLOSE, BT_WEAK
   
   nNext = nNext + (SCREEN_PIXEL_WIDTH / (nTemp1 + 1))
  Next
  
  nNext = 300
  For cn = nTemp1 + 1 To nTemp2
    CBunker(cn).CreateBunker nGetRnd(nNext - 100, nNext), _
                             nGetRnd(340, 359), BP_FAR, BT_WEAK
     
   nNext = nNext + (SCREEN_PIXEL_WIDTH / (nTemp2 + 1))
  Next
  
  nNext = 300
  For cn = nTemp2 + 1 To MAX_BUNKERS
   CBunker(cn).CreateBunker nGetRnd(nNext - 100, nNext), _
                            nGetRnd(347, 359), BP_VERYFAR, BT_WEAK
    
   nNext = nNext + (SCREEN_PIXEL_WIDTH / (nTemp1 + nTemp2))
  Next
  
  ' --------- arrange bunkers by distance from player ( for easier blitting )
  For i = 0 To MAX_BUNKERS - 1
   Set cTBunker = CBunker(i)
   nTemp1 = i
    
    For j = nTemp1 + 1 To MAX_BUNKERS
     If cTBunker.GetZ < CBunker(j).GetZ Then
       Set cTBunker = CBunker(j)
       nTemp1 = j
     End If
    Next
     
    Set CBunker(nTemp1) = CBunker(i)
    Set CBunker(i) = cTBunker
  Next

End Sub


Public Sub Do_PreCalcs()                                ' do precalculations to make game faster
  
  Dim eWeapon As enumWeapon
  Dim eShip   As enumShipType
  Dim cn      As Long
  
  ' --- Prepare unit classes  --- {!}
  
  ReDim CShip(MAX_ENEMIES)
  ReDim CBunker(MAX_BUNKERS)
  
  For cn = 0 To MAX_ENEMIES
   Set CShip(cn) = New clsShip
   CShip(cn).SetIndex = cn
  Next
  
  Call Randomize(GetTicks)
  Call SetupStars
  Call SetupBunkers
  
  g_numsmq = -1             ' reset message counter
  ' ------------------
  Call LoadCadets
  
   ' do main menu position calcs
   For cn = 0 To UBound(g_MenuPos)
    g_MenuPos(cn).x = 35
    g_MenuPos(cn).y = 125 + (cn * 30)
   Next
  
  Call LoadTextData
    
  ' precalculate background positions
   xBg1 = 0: yBg1 = 0
   xBg2 = xBg1 + MAX_CX: yBg2 = 0
   xBg3 = xBg2 + MAX_CX: yBg3 = 0
   'xBg4 = xBg3 + MAX_CX: yBg4 = 0
  
  ' precalcutlate moon separations
  arMS_Offsets(0) = 0
  arMS_Offsets(1) = MAX_CX
  arMS_Offsets(2) = MAX_CX * 2
  arMS_Offsets(3) = (MAX_CX * 3)
  'arMS_Offsets(4) = MAX_CX * 4
   'arMS_Offsets(5) = MAX_CX * 4
  ' precalculate earth position
  g_xEarth = (SCREEN_PIXEL_WIDTH \ 2 - g_cxEarth \ 2) - MAX_CX + 50
  'g_xEarth = (MAX_CX - g_cxEarth / 2)
  g_yEarth = (MAX_CY / 2 - g_cyEarth / 2) - 160
  ' fill weapons damage table
  'eWeapon = SW_LASER
  arDamages(SW_LASER) = 1
  arDamages(SW_PARTICLE) = 25
  arDamages(SW_GREENPUS) = 4
  arDamages(SW_MISSILES) = 3
  arDamages(SW_GREENLASER) = 4
  arDamages(SW_INTERCEPTOR) = 0                 ' mother ship does not do damage
  ' fill weapons ranges table ( in pixels )
  arAttackRange(SW_LASER) = 30
  arAttackRange(SW_PARTICLE) = 10
  arAttackRange(SW_GREENPUS) = 150
  arAttackRange(SW_MISSILES) = 500
  arAttackRange(SW_GREENLASER) = 200
  arAttackRange(SW_INTERCEPTOR) = 550
  ' fill velocity table in PPF (pixels/per/frame)
  'eShip = ST_INTERCEPTOR1
  arVelocity(ST_INTERCEPTOR1) = 2.6
  arVelocity(ST_CARRIER1) = 0.9
  arVelocity(ST_PARTICLEBEAST) = 0.5
  arVelocity(ST_LRSNEAKY) = 1.1
  arVelocity(ST_LRCLOSETERROR) = 1.4
  arVelocity(ST_SEVENTHFOX) = 1.3
  ' load weapons colors
  'For cn = 1 To 10
  '  If cn >= 7 Then
  '   arClr_LaserBeam(cn - 1) = RGB(226 - cn * 4, 116 - cn * 4, 80)
  '   arClr_ParticleBeam(cn - 1) = RGB(120, 160 - cn * 4, 240 - cn * 4)
  '   arClr_GreenPus(cn - 1) = RGB(42 - cn * 4, 200 - cn * 4, 80 - cn * 4)
  '  Else
  '   arClr_LaserBeam(cn - 1) = RGB(178 + cn * 4, 68 + cn * 4, 80)
  '   arClr_ParticleBeam(cn - 1) = RGB(120, 112 + cn * 4, 192 + cn * 4)
  '   arClr_GreenPus(cn - 1) = RGB(2 + cn * 4, 160 + cn * 4, 80 + cn * 4)
  '  End If
  'Next
  
  ' load credits list
  'arCreditsList(0) = "Game Idea:  EraZor & Pro-XeX"
  'arCreditsList(1) = "Game Design:  everyone... ;)"
  'arCreditsList(2) = "Programming:  Pro-XeX (arrgh!)"
  'arCreditsList(3) = "Lead Artist:  ErazoR"
  'arCreditsList(4) = "2D Artists:  EraZoR & PrankMaster"
  'arCreditsList(5) = "2D Animations:  EraZoR, LorD_VoRTeX, PrankMaster"
  'arCreditsList(6) = "3D Artists:  EraZoR, LorD_VoRTeX, Colt"
  'arCreditsList(7) = "SFX:  Grupa 'mUcHasHti'"
  'arCreditsList(8) = "Music:  Pro-XeX (yeah,I know u can't hear it...yet! :)"
  
  'Call CreateAuthor(arCreditsList(nGetRnd(0, 7)))
  'Call CreateAuthor(arCreditsList(nGetRnd(0, 7)))
  
End Sub


' load all cadets
Public Sub _
LoadCadets()
  Dim cn As Long
  
  ' load all characters
  'mCharacters.g_strCharFileName = App.Path & "\data\chars"
  If (Not mCharacters.CHARLoadCharacters(App.Path & "\cadets")) Then
    ' error loading characters file -> create default one
    Call CHARNewCharacter("BabaTU", "", 1000, , , 1000, 1000, 1000)
    If (Not CHARSaveCharacters()) Then
     AppendToLog ("Error saving characters! '!Precreation!'")
     Call MakeError("Error saving characters!")
    End If
    If (Not CHARLoadCharacters(App.Path & "\cadets")) Then
     AppendToLog ("Error loading characters! '!Precreation!'")
     Call MakeError("Error loading characters!")
    End If
  End If
  ' get levels
  For cn = 0 To g_nChars
   g_arChar(cn).level = CHARExperienceToLevel(g_arChar(cn).Exp)
  Next
  
End Sub


' load textdata files
Public Function _
LoadTextDataFile(strFileName As String, lPos As Long, strBuffer() As String) As Boolean

 On Local Error GoTo ERR_LOADTDF:

 Dim ff         As Integer
 Dim buffer     As String
 Dim strbegin   As String
 Dim strend     As String
 Dim cn         As Long
 
 strbegin = "{TEXTDATA}"
 strend = "{ENDTEXTDATA}"
 cn = 0
 
 ff = FreeFile()
 
 Open (strFileName) For Input Access Read Lock Write As #ff
  
  ' set file position
  Seek #ff, lPos
  
  Line Input #ff, buffer
  ' check header
  If (Left$(buffer, Len(strbegin)) = strbegin) Then
  
   Line Input #ff, buffer
   Do While (Left$(buffer, Len(strend)) <> strend)
    ReDim Preserve strBuffer(cn)
    strBuffer(cn) = buffer
    cn = cn + 1
    Line Input #ff, buffer
   Loop
  
  End If
   
 Close #ff

 LoadTextDataFile = True
Exit Function

ERR_LOADTDF:
 LoadTextDataFile = False
End Function


' load all text interface
Public Sub _
LoadTextData()

  On Local Error GoTo ERR_LOADTD:

  Dim tempPath   As String
  Dim buffer     As String
  Dim strbegin   As String
  Dim strend     As String
  Dim cn         As Long
  Dim menu_txt() As String
  Dim prefix     As String
  
  strbegin = "{TEXTDATA}"
  strend = "{ENDTEXTDATA}"
  cn = 0
      
  tempPath = App.Path & "\interf\"
  
  ' load epilogue
  If (g_Language = LANG_ENGLISH) Then
   prefix = "eng"
  ElseIf (g_Language = LANG_BULGARIAN) Then
   prefix = "bg"
  End If
  
  ' load intro text
  If (Not LoadTextDataFile(tempPath & prefix & "intro.ktd", 1, g_arIntro)) Then GoTo ERR_LOADTD:
  ' load epligoue text
  If (Not LoadTextDataFile(tempPath & prefix & "eplg.ktd", 1, g_arEpilogue)) Then GoTo ERR_LOADTD:
  ' load rankings
  If (Not LoadTextDataFile(tempPath & prefix & "rank.ktd", 1, mCharacters.g_arRanks())) Then GoTo ERR_LOADTD:
  ' load menu texts
  If (Not LoadTextDataFile(tempPath & prefix & "menu.ktd", 1, menu_txt)) Then GoTo ERR_LOADTD:
     
   arText(MENU_START) = menu_txt(0)
   arText(MENU_OPTIONS) = menu_txt(1)
   arText(MENU_HALLOFFAME) = menu_txt(2)
   arText(MENU_HELP) = menu_txt(3)
   arText(MENU_EXIT) = menu_txt(4)
   arText(MENU_BACK) = menu_txt(5)
   arText(MENU_GAMMA) = menu_txt(6)
   arText(MENU_GAMMAUNAVAILABLE) = menu_txt(7)
   arText(MENU_SOUND) = menu_txt(8)
   arText(MENU_MUSIC) = menu_txt(9)
   arText(MENU_VSYNCON) = menu_txt(10)
   arText(MENU_VSYNCOFF) = menu_txt(11)
   arText(MENU_CREDITS) = menu_txt(12)
   arText(MENU_ADDNEWPLAYER) = menu_txt(13)
   arText(MENU_PLAYERNAMEQUERY) = menu_txt(14)
   arText(MENU_PLAYERNAMEERROR) = menu_txt(15)
   arText(MENU_PLAYERDELETE) = menu_txt(16)
   arText(MENU_PLAYERDELETEERROR) = menu_txt(17)
   arText(MENU_PASSWORDERROR) = menu_txt(18)
   arText(MENU_PASSWORDENTER) = menu_txt(19)
   arText(MENU_PASSWORDQUERY) = menu_txt(20)
   arText(MENU_PASSWORDCONFIRM) = menu_txt(21)
   arText(MENU_PLAYERNAME) = menu_txt(22)
   arText(MENU_PLAYERPASSWORD) = menu_txt(23)
   arText(MENU_PLAYERSCORE) = menu_txt(24)
   arText(MENU_PLAYERLEVEL) = menu_txt(25)
   arText(MENU_PLAYERMISSION) = menu_txt(26)
   arText(MENU_PLAYERTOTALSHOTS) = menu_txt(27)
   arText(MENU_PLAYERSUCCESS) = menu_txt(28)
   arText(MENU_PLAYERKILLS) = menu_txt(29)
   arText(MENU_PLAYEREXPERIENCE) = menu_txt(30)
   arText(MENU_PLAYERHELP1) = menu_txt(31)
   arText(MENU_PLAYERHELP2) = menu_txt(32)
   arText(MENU_PLAYERHELP3) = menu_txt(33)
   arText(MENU_INFOBOXCREDITS) = menu_txt(34)
   arText(MENU_INFOBOXHOF) = menu_txt(35)
   arText(MENU_INFOBOXHELP) = menu_txt(36)
   arText(MENU_HELP1) = menu_txt(37)
   arText(MENU_HELP2) = menu_txt(38)
   arText(MENU_HELP3) = menu_txt(39)
   arText(MENU_HELP4) = menu_txt(40)
   arText(MENU_HELP5) = menu_txt(41)
   arText(MENU_HELP6) = menu_txt(42)
   arText(MENU_HELP7) = menu_txt(43)
   arText(MENU_HELP8) = menu_txt(44)
   arText(MENU_HELP9) = menu_txt(45)
   arText(MENU_HELP10) = menu_txt(46)
   arText(MENU_HELP11) = menu_txt(47)
   arText(MENU_HELP12) = "" 'menu_txt(48)
   arText(MENU_COCKPIT_EARTH) = menu_txt(49)
   arText(MENU_COCKPIT_TIMELEFT) = menu_txt(50)
   arText(MENU_COCKPIT_WEAPON) = menu_txt(51)
   arText(MENU_COCKPIT_WEAPON_LASER) = menu_txt(52)
   arText(MENU_COCKPIT_WEAPON_MISFAR) = menu_txt(53)
   arText(MENU_COCKPIT_WEAPON_MISCLOSE) = menu_txt(54)
   arText(MENU_INTRO) = menu_txt(55)
   arText(MENU_SOLIDBARSON) = menu_txt(56)
   arText(MENU_SOLIDBARSOFF) = menu_txt(57)
  
  Erase menu_txt()

Exit Sub

ERR_LOADTD:
 Call MakeError("Error Loading Text Interface!" & vbCr & "Please, reinstall the game!")
End Sub


Public Sub _
ResetGame()
 ' Desc: Reset game vars (must be done before level start)

 Dim cn As Long
 
 g_hpEarth = 250

 Erase g_WarpGate()
 Erase g_PExp()
 Erase g_CPixel()
 Erase g_Missile()
 Erase g_LaserCut()
 Erase g_StarTrip()
 Erase g_Explosion()
 Erase g_Meteor()
 Erase g_Pl_Weapon()
 Erase g_Bonus()
 Erase g_PRemover()

 Erase CShip()
 ReDim CShip(MAX_ENEMIES)
 'Erase CBunker()
 For cn = 0 To MAX_BUNKERS
  CBunker(cn).SetVisible = True
 Next
 
 ' reset enemy ships
 For cn = 0 To MAX_ENEMIES
  Set CShip(cn) = New clsShip
  CShip(cn).SetIndex = cn
 Next
 
 ' reset battlestation
 Set CBattleStation = New clsBattleStation
 
End Sub


Public Sub CreateAuthor(lpszNameAndOccup As String)
 ' Desc: Adds an author to the credit's list :-)
 Dim cn As Integer
 
 Do While cn < UBound(g_Credits)
  With g_Credits(cn)
   If Not .Visible Then
    .x = MAX_CX / 2
    .y = MAX_CY / 2
    .xVel = CByte(nGetRnd(1, 2))
    .yVel = CByte(nGetRnd(1, 2))
    .Heading = CByte(nGetRnd(0, 3))
    .nReserved = nGetRnd(-2, 2)
    .fs = 0
    .ang = nGetRnd(0, 359)
    .lpszAuthor = lpszNameAndOccup
    '.lpszOccupation = "none"
    .cr = 50
    .Visible = True
    Exit Do
   End If
  End With
 cn = cn + 1
 Loop
 
End Sub


Public Sub GetCommandLine()
 ' Desc: set command line options
 Dim szCL As String
 
 szCL = Command()
 
 ' check for windowed state
 If InStr(1, szCL, "-w") Then
  bWindowed = True
 End If
 ' if main menu is not to be displayed
 If InStr(1, szCL, "-nomm") Then
  'bMainMenu = False
  
 End If
 ' what kind of memory are we going to use for graphics
 'If InStr(1, szCL, "-vram") Then
 ' mDirectDraw.lMemMethod = DDSCAPS_VIDEOMEMORY
 'ElseIf InStr(1, szCL, "-sram") Then
 ' mDirectDraw.lMemMethod = DDSCAPS_SYSTEMMEMORY
 'End If
 
 g_bNotRetrace = True
 ' rectracing status
 If InStr(1, szCL, "-ret") Then
  g_bNotRetrace = False
 End If
 If InStr(1, szCL, "-Debug") Then
  g_bDebug = True
 Else
  g_bDebug = False
 End If
 
 ' load preferences file
 Dim ff      As Integer
 Dim buffer  As String
 Dim lbuffer As Long
  
  On Local Error Resume Next
    
  ff = FreeFile()
    
  Open (App.Path & "\pref") For Binary Access Read Lock Write As #ff
    
    Get #ff, , lbuffer
    If (lbuffer) Then
     g_Language = LANG_BULGARIAN
    Else
     g_Language = LANG_ENGLISH
    End If
        
    Get #ff, , lbuffer
    If (lbuffer) Then
     g_bNotRetrace = False
    Else
     g_bNotRetrace = True
    End If
    
    Get #ff, , lbuffer
    buffer = Space$(lbuffer)
    Get #ff, , buffer
    mDirectDraw.sDDrawDriver = buffer
    
  Close #ff


End Sub


Public Sub UpdateWorld(Optional g_L As Boolean, Optional g_R As Boolean, _
                       Optional g_U As Boolean, Optional g_D As Boolean)
                                                    ' do world scrolling if required
Dim goLeft As Boolean, goRight As Boolean
Dim goUp As Boolean, goDown As Boolean
 
  If CMouse.GetX > rScreen.Right - 10 Then goRight = True
  If CMouse.GetX < 15 Then goLeft = True
  If CMouse.GetY > rScreen.Bottom - 10 Then goDown = True
  If CMouse.GetY < 10 Then goUp = True
  
  If goRight Or g_R Then                            ' scroll_left
     wx = wx + w_ScrollRate
     If wx >= SCREEN_PIXEL_WIDTH - rScreen.Right Then
        wx = SCREEN_PIXEL_WIDTH - rScreen.Right
        'wx = arMS_Offsets(0)
     End If
  End If
  
  If goLeft Or g_L Then                             ' scroll_right
     wx = wx - w_ScrollRate
     If wx < arMS_Offsets(0) Then
        wx = arMS_Offsets(0)
        'wx = arMS_Offsets(3)
     End If
  End If
    
  If goDown Or g_D Then                              ' scroll_down
      wy = wy + w_ScrollRate
      If wy >= SCREEN_PIXEL_HEIGHT - rScreen.Bottom Then wy = SCREEN_PIXEL_HEIGHT - rScreen.Bottom
  End If
  
  If goUp Or g_U Then                               ' scroll_up
     wy = wy - w_ScrollRate
     If wy < 5 Then wy = 0
  End If
End Sub


'////////////////////////////////////////////////////////////////
'//// Load all game Sounds
'////////////////////////////////////////////////////////////////
Public Sub _
LoadSounds()

 Dim cn         As Long
 Dim tempPath   As String
 Dim LoadSrc As cnstLOADSOURCE
 
 If (g_bDebug) Then
  tempPath = App.Path & "\sfx\"
  LoadSrc = LS_FROMFILE
 Else
  tempPath = ""
  LoadSrc = LS_FROMBINRES
 End If
 
 frmMain.lblStatus.Caption = "Loading sounds..."
 Call DSCreateSound(g_dsCannon(SFX_CANNONPLAYER1), 2, LoadSrc, tempPath & "plcan1.wav")
 Call DSCreateSound(g_dsCannon(SFX_CANNONPLAYER2), 2, LoadSrc, tempPath & "plcan2.wav")
 Call DSCreateSound(g_dsCannon(SFX_CLOSEBUNKER), 3, LoadSrc, tempPath & "bunlas1.wav")
 Call DSCreateSound(g_dsCannon(SFX_FARBUNKER), 3, LoadSrc, tempPath & "bunlas2.wav")
 Call DSCreateSound(g_dsCannon(SFX_VERYFARBUNKER), 3, LoadSrc, tempPath & "blaster1.wav")
 Call DSCreateSound(g_dsSfx(SFX_BUNKEREXPLODE), 2, LoadSrc, tempPath & "bexp.wav")
 
 Call DSCreateSound(g_dsSfx(SFX_MENUCHOICE), 1, LoadSrc, tempPath & "mnu_c.wav")
 Call DSCreateSound(g_dsSfx(SFX_MENUSELECT), 2, LoadSrc, tempPath & "mnu_sel.wav")
 Call DSCreateSound(g_dsSfx(SFX_MENUCALIBRATE), 2, LoadSrc, tempPath & "mnu_snd.wav")
 Call DSCreateSound(g_dsSfx(SFX_MENUEXIT), 1, LoadSrc, tempPath & "mnu_q.wav")
 Call DSCreateSound(g_dsSfx(SFX_SPACEMYST1), 1, LoadSrc, tempPath & "SPACE4a.wav")
 Call DSCreateSound(g_dsSfx(SFX_OPENGATE), 1, LoadSrc, tempPath & "mnu_go.wav")
 Call DSCreateSound(g_dsSfx(SFX_CLOSEGATE), 1, LoadSrc, tempPath & "mnu_gc.wav")
 Call DSCreateSound(g_dsSfx(SFX_COCKPITSMQ), 1, LoadSrc, tempPath & "cp_smq.wav")
 Call DSCreateSound(g_dsSfx(SFX_PLAYERROCKETFIRE), 3, LoadSrc, tempPath & "plr1.wav")
 Call DSCreateSound(g_dsSfx(SFX_WARPGATE), 3, LoadSrc, tempPath & "warpgate.wav")
 Call DSCreateSound(g_dsSfx(SFX_GREENLASER1), 2, LoadSrc, tempPath & "grnlas1.wav")
 Call DSCreateSound(g_dsSfx(SFX_GREENLASER2), 2, LoadSrc, tempPath & "grnlas2.wav")
 Call DSCreateSound(g_dsSfx(SFX_NORMALASER), 4, LoadSrc, tempPath & "grnlas2.wav")
 Call DSCreateSound(g_dsSfx(SFX_INTERCEPT1), 2, LoadSrc, tempPath & "int1.wav")
 Call DSCreateSound(g_dsSfx(SFX_INTERCEPT2), 2, LoadSrc, tempPath & "int2.wav")
 Call DSCreateSound(g_dsSfx(SFX_FAREXPLOSION1), 6, LoadSrc, tempPath & "farexp1.wav")
 Call DSCreateSound(g_dsSfx(SFX_BIGBLAST1), 2, LoadSrc, tempPath & "heavyb3.wav")
 Call DSCreateSound(g_dsSfx(SFX_BIGBLAST2), 2, LoadSrc, tempPath & "heavyb1.wav")
 Call DSCreateSound(g_dsSfx(SFX_METEORBLAST), 2, LoadSrc, tempPath & "heavyb4.wav")
 Call DSCreateSound(g_dsSfx(SFX_PARTICLEXPLOSION), 2, LoadSrc, tempPath & "prexp.wav")
 
 ' set volumes
 'Call DSSetSoundVolume(g_dsCannon(SFX_VERYFARBUNKER), SFX_VOLUMEVERYFAR)
 'Call DSSetSoundVolume(g_dsCannon(SFX_FARBUNKER), SFX_VOLUMEFAR)
 'Call DSSetSoundVolume(g_dsCannon(SFX_CLOSEBUNKER), SFX_VOLUMECLOSE)
 
 'bDSOn = False
  frmMain.UpdatePBar 80

End Sub

'////////////////////////////////////////////////////////////////
'//// Load all game graphics
'////////////////////////////////////////////////////////////////
Public Sub _
LoadGraphics()
 
 Dim cn       As Long
 Dim tempPath As String
 Dim LoadSr   As cnstLOADSOURCE
 
  
 If (g_bDebug) Then
  LoadSr = LS_FROMFILE
 Else
  LoadSr = LS_FROMBINRES
  tempPath = ""
 End If

 With g_Objects
     frmMain.lblStatus.Caption = "Loading menu graphics..."
     ' load menu graphics
     If (g_bDebug) Then tempPath = App.Path & "\gfx\menu\"
     .Cursor(0) = CreateGFX_HBM(tempPath & "cursor.bmp", 0, 0, True, 0&, LoadSr)
     .Cursor(1) = CreateGFX_HBM(tempPath & "cursored.bmp", 0, 0, True, 0&, LoadSr)
     .vbc = CreateGFX_HBM(tempPath & "vbc.bmp", 0, 0, False, , LoadSr)
     '.Cursor = CreateGFX_HBM(App.Path & "\gfx\gfx.kdf", cKdf.GetEntryPositionFromName("cursor.bmp"), 25, 25, True)
     'Set .Cursor = DDLoadBitmapFromBinRes(TempPath & "cursor.bmp", 1, 25, 25, True, 0)
'     Stop
      
     .credits = CreateGFX_HBM(tempPath & "credits.bmp", 0, 0, True, 0&, LoadSr)
     'Set .WG_Back = DDLoadSurfaceFromFile(TempPath & "wgate.bmp", 320, 240, False)
     '.menu = CreateGFX_HBM(TempPath & "menu.bmp", 0, 0, True, 0&, LoadSr)
     .seline = CreateGFX_HBM(tempPath & "seline.bmp", 0, 0, False, , LoadSr)
     .caline = CreateGFX_HBM(tempPath & "caline.bmp", 0, 0, False, , LoadSr)
     .gate(0) = CreateGFX_HBM(tempPath & "gate1.bmp", 0, 0, False, , LoadSr)
     .gate(1) = CreateGFX_HBM(tempPath & "gate2.bmp", 0, 0, False, , LoadSr)
     .dialog = CreateGFX_HBM(tempPath & "dlg.bmp", 0, 0, True, 0&, LoadSr)
     .errdialog = CreateGFX_HBM(tempPath & "errdlg.bmp", 0, 0, True, 0&, LoadSr)
     ' load backpapaers
     '.backpaper(0) = CreateGFX_HBM(App.Path & "\kdf1.kdf", 0, 0, True, 0, LS_FROMBINRES, cKdf.GetPackedFilePosition(1))
     .backpaper(0) = CreateGFX_HBM(tempPath & "bp1.bmp", 0, 0, True, 0&, LoadSr)
     .backpaper(1) = CreateGFX_HBM(tempPath & "bp2.bmp", 0, 0, False, 0&, LoadSr)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\menu\logo\"
     For cn = 0 To UBound(g_Objects.Title)
      .Title(cn) = CreateGFX_HBM(tempPath & "logo" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
        
   ' Load environment
     frmMain.UpdatePBar 20
     frmMain.lblStatus.Caption = "Loading menu environment..."
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\"
     .CockPit = CreateGFX_HBM(tempPath & "cockpit.bmp", 0, 0, True, 0&, LoadSr)
     .buton = CreateGFX_HBM(tempPath & "buton.bmp", 0, 0, True, , LoadSr)
     .butoff = CreateGFX_HBM(tempPath & "butoff.bmp", 0, 0, True, , LoadSr)
     .es1 = CreateGFX_HBM(tempPath & "es1.bmp", 0, 0, False, , LoadSr)
     .es2 = CreateGFX_HBM(tempPath & "es2.bmp", 0, 0, False, , LoadSr)
     .es3 = CreateGFX_HBM(tempPath & "es3.bmp", 0, 0, False, , LoadSr)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\bonus\"
     ' load bonus-pix
     For cn = 0 To 6
      .Bonus(0, cn) = CreateGFX_HBM(tempPath & "bs1" & (cn + 1) & ".bmp", 0, 0, True, , LoadSr)
      .Bonus(1, cn) = CreateGFX_HBM(tempPath & "bs2" & (cn + 1) & ".bmp", 0, 0, True, , LoadSr)
      .Bonus(2, cn) = CreateGFX_HBM(tempPath & "bs3" & (cn + 1) & ".bmp", 0, 0, True, , LoadSr)
      .Bonus(3, cn) = CreateGFX_HBM(tempPath & "bs4" & (cn + 1) & ".bmp", 0, 0, True, , LoadSr)
      .Bonus(4, cn) = CreateGFX_HBM(tempPath & "bs5" & (cn + 1) & ".bmp", 0, 0, True, , LoadSr)
      .Bonus(5, cn) = CreateGFX_HBM(tempPath & "bs6" & (cn + 1) & ".bmp", 0, 0, True, , LoadSr)
     Next
     '{!}
     'If (g_GameState = GAMSTATE_MAINMENU) Then GoTo mainhere
    
     ' earth info
         ' 360, 320
     g_cxEarth = 264
     g_cyEarth = 262
     g_hpEarth = EARTH_HITPOINTS
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\"
     '.Star1(0) = CreateGFX_HBM(TempPath & "star1.bmp", 0, 0, True, 0, LoadSr)
     .Earth = CreateGFX_HBM(tempPath & "earth.bmp", 0, 0, True, , LoadSr)
     .earthhp = CreateGFX_HBM(tempPath & "ehp.bmp", 0, 0, False, , LoadSr, , SML_VIDEO)
     .bshp = CreateGFX_HBM(tempPath & "bshp.bmp", 0, 0, False, , LoadSr, , SML_VIDEO)
     .Sun = CreateGFX_HBM(tempPath & "sunlack.bmp", 0, 0, True, 0&, LoadSr)
     
     'For cn = 0 To 35
     '.Earth(cn) = CreateGFX_HBM(TempPath & "\earth\earthprv" & cn & ".bmp", 380, 380, True, 0)
     'Next
     '.BackGround(0) = CreateGFX_HBM(TempPath & "b1.bmp", 640, 370, False)
     '.BackGround(1) = CreateGFX_HBM(TempPath & "b2.bmp", 640, 370, False)
     '.BackGround(2) = CreateGFX_HBM(TempPath & "b3.bmp", 640, 370, False)
     .BackMoon(0) = CreateGFX_HBM(tempPath & "fm8x6401.bmp", 0, 0, True, 0&, LoadSr)
     .BackMoon(1) = CreateGFX_HBM(tempPath & "fm8x6402.bmp", 0, 0, True, 0&, LoadSr)
     
     .MoonSurf(0) = CreateGFX_HBM(tempPath & "m1x8.bmp", 0, 0, True, 0&, LoadSr)
     .MoonSurf(1) = CreateGFX_HBM(tempPath & "m2x8.bmp", 0, 0, True, 0&, LoadSr)
     .MoonSurf(2) = CreateGFX_HBM(tempPath & "m3x8.bmp", 0, 0, True, 0&, LoadSr)
     .MoonSurf(3) = CreateGFX_HBM(tempPath & "m4x8.bmp", 0, 0, True, 0&, LoadSr)
     
     ' load meteors
     For cn = 0 To UBound(.Meteor1)
      If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\astro1\"
      .Meteor1(cn) = CreateGFX_HBM(tempPath & "mete2" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
      .Meteor1_Shadow(cn) = CreateGFX_HBM(tempPath & "mete2" & cn & ".bmp", 0, 0&, True, 0&, LoadSr)
      Call CreateShadow(.Meteor1_Shadow(cn).dds)
      
      If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\astro2\"
      .Meteor2(cn) = CreateGFX_HBM(tempPath & "metr2" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
     
     ' load warpgate(s)
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\wgv\"
     For cn = 0 To UBound(.WarpGate_Far)
      '.WarpGate(cn) = CreateGFX_HBM(tempPath & "j" & cn & ".bmp", 120, 100, True, 0&, LoadSr)
      .WarpGate_Far(cn) = CreateGFX_HBM(tempPath & "j" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
      'Set .WarpGate(cn) = DDLoadSurfaceFromFile(TempPath & "wg000" & cn + 1 & ".bmp", 120, 100, True, 0)
      'Set .WarpGate_Far(cn) = DDLoadSurfaceFromFile(TempPath & "wg000" & cn + 1 & ".bmp", 60, 50, True, 0)
     Next
     
     ' load explosion(s)
     For cn = 0 To UBound(.Exp1)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\expbig\"
      .Exp1(cn) = CreateGFX_HBM(tempPath & "expl" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\expbig2\"
      .Exp1Far(cn) = CreateGFX_HBM(tempPath & "exbs" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     
     Next
     
     
     For cn = 0 To UBound(.Exp2)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\expsm\"
      .Exp2(cn) = CreateGFX_HBM(tempPath & "exsm00" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\expsm2\"
      .Exp2Far(cn) = CreateGFX_HBM(tempPath & "exsms0" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
      'Call CreateImageDC(.Exp2DC(cn), TempPath & "exsm00" & cn & ".bmp", 80, 60)
     Next
     
     
     For cn = 0 To UBound(.Exp3)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\expbl\"
      .Exp3(cn) = CreateGFX_HBM(tempPath & "exbl000" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\expbl2\"
      .Exp3Far(cn) = CreateGFX_HBM(tempPath & "exbls00" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     
     Next
     ' load particles
     If (g_bDebug) Then tempPath = App.Path & "\gfx\environ\"
     .FireCles = CreateGFX_HBM(tempPath & "firecles.bmp", 0, 0, False, , LoadSr)
     .BlueCles = CreateGFX_HBM(tempPath & "bluecles.bmp", 0, 0, False, , LoadSr)
     .ChillCles = CreateGFX_HBM(tempPath & "chills.bmp", 0, 0, False, , LoadSr)
     .starcles = CreateGFX_HBM(tempPath & "starcles.bmp", 0, 0, False, , LoadSr)
   
   ' --- Load ships
     frmMain.UpdatePBar 20
     frmMain.lblStatus.Caption = "Loading ships graphics..."
     
     ' // load carrier
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\carrier\"
     .Ship2(0) = CreateGFX_HBM(tempPath & "movel.bmp", 0, 0, True, 0&, LoadSr)
     For cn = 0 To 13 ' right side
     .Ship2(cn + 1) = CreateGFX_HBM(tempPath & "cl00" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next             ' left side
     .Ship2(15) = CreateGFX_HBM(tempPath & "mover.bmp", 0, 0, True, 0&, LoadSr)
     For cn = 0 To 13
     .Ship2(cn + 16) = CreateGFX_HBM(tempPath & "cr00" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
     
     ' // interceptors
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\intcpt\"
     For cn = 0 To 5
      ' left face
      .Ship1(cn) = CreateGFX_HBM(tempPath & "l1000" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
      ' right face
      .Ship1(cn + 6) = CreateGFX_HBM(tempPath & "r1000" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
     
     ' // beam_carrier
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bc\"
     For cn = 0 To 5
      ' left face
      .Ship5(cn) = CreateGFX_HBM(tempPath & "bc" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
      ' right face
      .Ship5(cn + 6) = CreateGFX_HBM(tempPath & "bcr" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
     
     ' // laser_cruiser
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\med_t\"
     For cn = 0 To 7
      .Ship4(cn) = CreateGFX_HBM(tempPath & "l2000" & cn & ".bmp", 0, 0, True, 0&, LoadSr)   ' fox
      .Ship4(cn + 8) = CreateGFX_HBM(tempPath & "1000" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
     
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\"
     ' // misile cruiser
     .Ship3(0) = CreateGFX_HBM(tempPath & "miscrl.bmp", 0, 0, True, 0&, LoadSr)  ' mis1
     .Ship3(1) = CreateGFX_HBM(tempPath & "miscrr.bmp", 0, 0, True, 0&, LoadSr)
     
     frmMain.UpdatePBar 20
     frmMain.lblStatus.Caption = "Loading weapons graphics..."
  ' --- Load weapons
     .RocketL(0) = CreateGFX_HBM(tempPath & "rocketl.bmp", 0, 0, True, 0&, LoadSr)
     .RocketR(0) = CreateGFX_HBM(tempPath & "rocketr.bmp", 0, 0, True, 0&, LoadSr)
     '.RocketL_Close(0) = CreateGFX_HBM(tempPath & "misl1.bmp", 20, 18, True, 0&, LoadSr)
     '.RocketR_Close(0) = CreateGFX_HBM(tempPath & "misr1.bmp", 20, 18, True, 0&, LoadSr)
     '.RocketL_VClose(0) = CreateGFX_HBM(tempPath & "misl1.bmp", 25, 23, True, 0&, LoadSr)
     '.RocketR_VClose(0) = CreateGFX_HBM(tempPath & "misr1.bmp", 25, 23, True, 0&, LoadSr)
     
     .CannonLeft = CreateGFX_HBM(tempPath & "ncr.bmp", 0, 0, True, 0&, LoadSr)
     .CannonRight = CreateGFX_HBM(tempPath & "ncl.bmp", 0, 0, True, 0&, LoadSr)
     ' particle remover
     .pr = CreateGFX_HBM(tempPath & "pr1.bmp", 0, 0, True, 0&, LoadSr)
     
     'TempPath = App.Path & "\gfx\ships\wpns\"
     .RedLaser(0) = CreateGFX_HBM(tempPath & "redl.bmp", 0, 0, True, 0&, LoadSr)
     .RedLaser(1) = CreateGFX_HBM(tempPath & "redr.bmp", 0, 0, True, 0&, LoadSr)
     .GreenLaser(0) = CreateGFX_HBM(tempPath & "laser5.bmp", 0, 0, True, 0&, LoadSr)
     .GreenLaser(1) = CreateGFX_HBM(tempPath & "laser5l.bmp", 0, 0, True, 0&, LoadSr)
    
     ' load particle laser
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bc\"
     For cn = 0 To UBound(.bclaz)
      .bclaz(cn) = CreateGFX_HBM(tempPath & "bclaz" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
    
     ' load player weapons
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\gm\"
     For cn = 0 To UBound(.GM)
      .GM(cn) = CreateGFX_HBM(tempPath & "r" & (cn + 1) & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\ls\"
     For cn = 0 To UBound(.LS)
      .LS(cn) = CreateGFX_HBM(tempPath & "ls" & (cn + 1) & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
       
     frmMain.UpdatePBar 20
     frmMain.lblStatus.Caption = "Loading surface models..."
     
     ' load BattleStation
     For cn = 0 To 14
      If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bs\"
      .BattleStation(cn) = CreateGFX_HBM(tempPath & "bsprv" & cn & ".bmp", 0, 0, True, 0&, LoadSr)
     Next
     
     .bs_missile = CreateGFX_HBM(tempPath & "bsmis.bmp", 0, 0, True, 0&, LoadSr)
     
     ' load Bunkers
     For cn = 1 To 6
      If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bunk4\"
      .Bunker1(0, cn - 1) = CreateGFX_HBM(tempPath & "4b4" & cn & ".bmp", 0, 0, True, , LoadSr)
      If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bunk2\"
      .Bunker1(1, cn - 1) = CreateGFX_HBM(tempPath & "2b2" & cn & ".bmp", 0, 0, True, , LoadSr)
      If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bunk3\"
      .Bunker1(2, cn - 1) = CreateGFX_HBM(tempPath & "3b3" & cn & ".bmp", 0, 0, True, , LoadSr)
     Next
     
     ' Load bunker_dead art
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bunk4\"
     .Bunker1Dead(0) = CreateGFX_HBM(tempPath & "4dead.bmp", 0, 0, True, , LoadSr)
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bunk2\"
     .Bunker1Dead(1) = CreateGFX_HBM(tempPath & "2dead.bmp", 0, 0, True, , LoadSr)
     If (g_bDebug) Then tempPath = App.Path & "\gfx\ships\bunk3\"
     .Bunker1Dead(2) = CreateGFX_HBM(tempPath & "3dead.bmp", 0, 0, True, , LoadSr)
 
     ' ------------------- LOAD FROM BINARY RESOURCE --------------
     GoTo mainhere
  End With
  
  
mainhere:
 
 

End Sub


' Release All Graphics
' ---------------------------------------------------------------
Public Sub _
ReleaseGame()
 
 Dim cn As Long
 Dim i  As Long
 
 ' save settings in registry
 mUtil.RegSetKey &H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "sndvol", CStr(mDirectSound.m_nGlobalVol)
 mUtil.RegSetKey &H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "musvol", CStr(g_nMusicVol)
 mUtil.RegSetKey &H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "path", App.Path
 mUtil.RegSetKey &H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "transbars", CStr(g_bsolidbars)
 
 ' close FMOD
 mFMod.STOPmus
 mFMod.CLOSEmus
 mFMod.closeFMOD
 AppendToLog ("FMod: closed.")
 
 ShowCursor True
 If (Not CHARSaveCharacters()) Then
  Call ErrorMsg("Error saving characters!")
  AppendToLog ("Error saving characters! '!Precreation!'")
 End If
  
 ' unload classes
 For cn = 0 To MAX_ENEMIES
  Set CShip(cn) = Nothing
 Next
 For cn = 0 To MAX_BUNKERS
  Set CBunker(cn) = Nothing
 Next
 Set CMouse = Nothing
 'Set cAI = Nothing
  
 ' unload sounds
 For cn = 0 To 4
  For i = 0 To UBound(g_dsCannon(cn).m_lpDSBuffer)
   Set g_dsCannon(cn).m_lpDSBuffer(i) = Nothing
  Next
 Next
 
 For cn = 0 To SFX_SOUNDS
  For i = 0 To UBound(g_dsSfx(cn).m_lpDSBuffer)
   Set g_dsSfx(cn).m_lpDSBuffer(i) = Nothing
  Next
 Next
   
 ' unload graphics
 With g_Objects
  Set .vbc.dds = Nothing
  Set .Earth.dds = Nothing
  Set .earthhp.dds = Nothing
  Set .bshp.dds = Nothing
  Set .Sun.dds = Nothing
  Set .Cursor(0).dds = Nothing
  Set .Cursor(1).dds = Nothing
  Set .CockPit.dds = Nothing
  'Set .WG_Back = Nothing
  Set .FireCles.dds = Nothing
  Set .BlueCles.dds = Nothing
  Set .ChillCles.dds = Nothing
  Set .buton.dds = Nothing
  Set .butoff.dds = Nothing
  Set .es1.dds = Nothing
  Set .es2.dds = Nothing
  Set .es3.dds = Nothing
  ' release menu
  Set .gate(0).dds = Nothing
  Set .gate(1).dds = Nothing
  'Set .menu.dds = Nothing
  Set .caline.dds = Nothing
  Set .seline.dds = Nothing
  Set .dialog.dds = Nothing
  Set .errdialog.dds = Nothing
  ' release earth's graphics
  'For cn = 0 To UBound(.Earth)
  ' Set .Earth(cn).dds = Nothing
  'Next
  
  ' release background graphics
  'For cn = 0 To UBound(.BackGround)
  ' Set .BackGround(cn).dds = Nothing
  'Next
   Set .Title(0).dds = Nothing
  
  ' release backpapers
  For cn = 0 To UBound(.backpaper)
   Set .backpaper(cn).dds = Nothing
  Next
  
  ' release moon graphpics
  For cn = 0 To UBound(.MoonSurf)
   Set .MoonSurf(cn).dds = Nothing
  Next
  Set .BackMoon(0).dds = Nothing
  Set .BackMoon(1).dds = Nothing
  
  ' release fx graphics
  For cn = 0 To UBound(.WarpGate_Far)
   'Set .WarpGate(cn).dds = Nothing
   Set .WarpGate_Far(cn).dds = Nothing
  Next
  For cn = 0 To UBound(.Exp1)
   Set .Exp1(cn).dds = Nothing
   Set .Exp1Far(cn).dds = Nothing
  Next
  For cn = 0 To UBound(.Exp2)
   Set .Exp2(cn).dds = Nothing
   Set .Exp2Far(cn).dds = Nothing
   'Call DeleteDC(.Exp2DC(cn))
  Next
  For cn = 0 To UBound(.Exp3)
   Set .Exp3(cn).dds = Nothing
   Set .Exp3Far(cn).dds = Nothing
  Next
  
  ' release ships graphics
  For cn = 0 To UBound(.Ship2)
   Set .Ship2(cn).dds = Nothing
  Next
  For cn = 0 To UBound(.Ship3)
   Set .Ship3(cn).dds = Nothing
  Next
  For cn = 0 To UBound(.Ship4)
   Set .Ship4(cn).dds = Nothing
  Next
  For cn = 0 To UBound(.Ship1)
   Set .Ship1(cn).dds = Nothing
  Next
  For cn = 0 To UBound(.Ship1)
   Set .Ship5(cn).dds = Nothing
  Next
  ' release rocket graphics
   Set .RocketL(0).dds = Nothing
   Set .RocketR(0).dds = Nothing
  
  ' release weapons graphics
  For cn = 0 To UBound(.RedLaser)
   Set .RedLaser(cn).dds = Nothing
   Set .GreenLaser(cn).dds = Nothing
  Next
  For cn = 0 To UBound(.bclaz)
   Set .bclaz(cn).dds = Nothing
  Next
  
  ' release player_weapons
  For cn = 0 To UBound(.GM)
   Set .GM(cn).dds = Nothing
   Set .LS(cn).dds = Nothing
  Next
   Set .CannonLeft.dds = Nothing
   Set .CannonRight.dds = Nothing
   Set .pr.dds = Nothing
  
  ' release meteors graphics
  For cn = 0 To UBound(.Meteor1)
   Set .Meteor1(cn).dds = Nothing
   Set .Meteor2(cn).dds = Nothing
   Set .Meteor1_Shadow(cn).dds = Nothing
  Next
  
  ' release battlestation graphics
  For cn = 0 To UBound(.BattleStation)
   Set .BattleStation(cn).dds = Nothing
  Next
  Set .bs_missile.dds = Nothing
  
  ' release bunker graphics
  For cn = 0 To UBound(.Bunker1)
   Set .Bunker1(0, cn).dds = Nothing
   Set .Bunker1(1, cn).dds = Nothing
   Set .Bunker1(2, cn).dds = Nothing
  Next
   Set .Bunker1Dead(0).dds = Nothing
   Set .Bunker1Dead(1).dds = Nothing
   Set .Bunker1Dead(2).dds = Nothing
  
 End With

End Sub

