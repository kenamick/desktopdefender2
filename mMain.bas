Attribute VB_Name = "mMain"
Option Explicit


Rem *** Desktop Defender® II ***
Rem *** Engine Started: 4:50 pm. on 29.XII.2001
Rem     Alpha Release:      12.IX.2002
Rem     Developer Release:  15.III.2003
Rem     Beta Release:       20.IV.2003
Rem     Official Release:   11.V.2003 / Fixed Release: 24.V.2003


Rem *** Main Module
Rem Purpose: Startup initializations, Input Handling, Game Loop...

Private Const GAME_ID = "Desktop Defender II - Battle for Existance"
Public Const MAX_CX = 640                           ' screen resolution
Public Const MAX_CY = 480
Public Const BPP = 16                               ' color depth



Sub Main()

Call LoadGame

bRunning = True                                     ' app. is ready to go

Do While (bRunning)                                 ' start main game loop

 Call GFXClearBackBuffer                            ' clear backbuffer content
 Call CalcFrameRate(nFPS)                           ' count FPS
 Call UpdateFrame                                   ' update game
  
  ' do synchornization ( in other words - retracing )
  If (Not g_bNotRetrace) Then
   Do While lpDD.WaitForVerticalBlank(DDWAITVB_BLOCKEND, 0) <> DD_OK
   Loop
  End If
  
 DoEvents                                           ' let windowZ do stuff
 Call DDBlitToPrim
Loop
 

Call ReleaseAll                                     ' release all mem and close game
End Sub


Public Sub LoadGame()                               ' everything to load and initialize is Here

 Dim tempPath As String
 Dim cn As Integer
 
 
 Call GetCommandLine                                ' set options from the command line
 'g_bDebug = True
 'bWindowed = True
 
 If (bWindowed) Then
    Call MoveWindow(frmMain.hwnd, 0, 0, MAX_CX + 50, MAX_CY + 50, True)
 End If
 DoEvents
 
 Call SetRect(rScreen, 0, 0, MAX_CX, MAX_CY)        ' setup screen rectangle
 w_ScrollRate = 15                                  ' set screen-scroll speed

 ' setup form
 ShowCursor False
 With frmMain
  .Caption = GAME_ID                          ' set game_ID
  .BackColor = &H0
  .Show
  DoEvents
 End With
 
 Call OpenLog(App.Path & "\")                       ' open log file
 AppendToLog (" *** DESKTOP DEFENDER II LOG FILE *** " & vbCrLf)
 AppendToLog ("Copyright © 2001-2003 KenamicK Entertainment®" & vbCrLf)
 AppendToLog ("Version " & App.Major & "." & App.Minor)
 AppendToLog ("Opened at: " & chGetTime)
 AppendToLog ("Opened on: " & Format(Date, "dddd, mmmm yyyy") & vbCrLf)
 
 ' setting high-priority ( vb games need that )
 AppendToLog (LOG_DASH)
 Call SetProgramPriority(P_HIGH)
  
 ' init DirectX
 AppendToLog (LOG_DASH)
 AppendToLog ("Initializing DirectX...")
 frmMain.lblStatus.Caption = "Initializing DirectX..."
 Call DXInit
 Call DDInit(frmMain.hwnd, MAX_CX, MAX_CY, BPP)
 frmMain.Picture = frmMain.picl.Picture
 Call DIInit(frmMain.hwnd, DI_KEYBOARD Or DI_MOUSE)
 Call DSInit(frmMain.hwnd)
 Call DDInitGamma
  
 ' init maintenance classes
 AppendToLog (LOG_DASH)
 AppendToLog ("Initializing Maintenance classes...")
 frmMain.UpdatePBar 80
 frmMain.lblStatus.Caption = "Initializing Maintenance classes..."
 
 Set CMouse = New clsMouse
 Call CMouse.SaveCoords
 CMouse.SetMouseInput = MI_API
 CMouse.SetCursorHeight = 25
 CMouse.SetCursorWidth = 25
 Call CMouse.Acquire

 ' do precalculations & reset game
 Call Do_PreCalcs
 
 frmMain.UpdatePBar 80
 frmMain.lblStatus.Caption = "Opening data paks.."
 ' prepare packets
 If (Not CKdfGfx.LoadTag(App.Path & "\data\gfx.ktf")) Then
  Call MakeError("Could not open " & App.Path & "\data\gfx.ktf" & vbCr & "Please, reinstall game!")
 Else
  If (Not CKdfGfx.LoadPacket(App.Path & "\data\gfx.kdf")) Then _
   Call MakeError("Could not open " & App.Path & "\data\gfx.kdf" & vbCr & "Please, reinstall game!")
 End If
 If (Not CKdfSfx.LoadTag(App.Path & "\data\sound.ktf")) Then
  Call MakeError("Could not open " & App.Path & "\data\sound.ktf" & vbCr & "Please, reinstall game!")
 Else
  If (Not CKdfSfx.LoadPacket(App.Path & "\data\sound.kdf")) Then _
   Call MakeError("Could not open " & App.Path & "\data\sound.kdf" & vbCr & "Please, reinstall game!")
 End If
 
 ' load all game_data
 AppendToLog (LOG_DASH)
 'mDirectSound.m_bDSOn = False
 'g_bDebug = True                                ' debug mode
 frmMain.UpdatePBar 80
 frmMain.lblStatus.Caption = "Loading graphics..."
 AppendToLog ("Loading graphics...")
 Call mGameProc.LoadSounds
 frmMain.lblStatus.Caption = "Loading sounds..."
 AppendToLog ("Loading sounds...")
 Call mGameProc.LoadGraphics
    
 Call mDirectDraw.DDFreeMemToLog
 AppendToLog ("Game started...")
 AppendToLog (LOG_DASH)
 
 ' clear main theme
 frmMain.Picture = LoadPicture()
 
 ' setup FMOD and load music
 frmMain.lblStatus.Caption = "Loading music..."
 mFMod.initFMOD
 AppendToLog ("FMod: initialized.")
 If (Not mFMod.OPENmus(App.Path & "\data\ar.it")) Then
  Call ErrorMsg("Could not load music file!")
 Else
  AppendToLog ("FMod: loaded \ar.it")
 End If
 frmMain.UpdatePBar 80
 
 Dim sval As String
  
  ' load music&sound volume
 mUtil.RegLoadKey &H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "sndvol", sval
 If (Val(sval) <> 0) Then mDirectSound.m_nGlobalVol = Val(sval)
 mUtil.RegLoadKey &H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "musvol", sval
 If (Val(sval) <> 0) Then g_nMusicVol = Val(sval)
    
 ' init
 g_Gates = GS_NONE
  
 ' check registry
 Call mUtil.RegLoadKey(&H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "logo", sval)
 g_GameState = GAMSTATE_LOGO
 Call mUtil.RegLoadKey(&H80000002, "SOFTWARE\KenamicK Entertainment\DD2\", "transbars", CStr(g_bsolidbars))
 
  ' hide loading navs
 frmMain.lblStatus.Visible = False
 frmMain.picl.Visible = False
 frmMain.shpload.Visible = False
 frmMain.shpoutline.Visible = False
 DoEvents

 Call DoKKLogo
 
 ' should (#logo#) intro be shown
 If (Val(sval) = 0) Then
  g_GameState = GAMSTATE_INTRO
 Else
  ' play background music
  Call mFMod.PLAYmus(True)
  g_bmusPlaying = True
  g_nMusicVol = MUSIC_VOLUMEMAX
  g_GameState = GAMSTATE_MAINMENU
 End If
 
 
 'g_GameState = GAMSTATE_BRIEFING
 'Call StartBriefing
 'g_GameState = GAMSTATE_MAINMENU

End Sub


Public Sub MakeError(lpStr As String)               ' error hadling procedure
  Call MsgBox(lpStr, vbExclamation, "What the...")
  Call ReleaseAll
End Sub


Public Function ConfirmMsg(lpStr As String) As Boolean ' confirm a message
 If MsgBox(lpStr, vbOKCancel) = vbOK Then ConfirmMsg = True _
  Else Call ReleaseAll
End Function


Public Sub ErrorMsg(lpStr As String)
 Call MsgBox(lpStr, vbCritical)
End Sub


Public Sub CheckIfTasked()                          ' check if game form has loset focus
 Dim ddrval As Long
 
 ddrval = lpDD.TestCooperativeLevel()
  
 If (ddrval <> DD_OK) Then
  DIUnAcquire (DI_KEYBOARD Or DI_MOUSE)             ' unacqure DI devices
  Call SetProgramPriority(P_NORMAL)                 ' return priorities
  
   Do While (ddrval <> DD_OK)
    ddrval = lpDD.TestCooperativeLevel()
    DoEvents
   Loop
  
  lpDD.RestoreAllSurfaces
  DIAcquire (DI_KEYBOARD Or DI_MOUSE)               ' acqure DI devices
  Call SetProgramPriority(P_HIGH)
 End If

  'If frmMain.WindowState = vbMinimized Then
     'DIUnAcquire (DI_KEYBOARD Or DI_MOUSE)          ' unacqure DI devices
     'Do
     'If frmMain.WindowState <> vbMinimized Then Exit Do
     '   ' ...
     'DoEvents
     'Loop
  'End If
End Sub


Public Sub ReleaseAll()                             ' release everything and close
 On Local Error Resume Next

 
 Call SetProgramPriority(P_NORMAL)
 AppendToLog (LOG_DASH)
 AppendToLog ("Unloading Game...")

 Call DDRestoreModes(frmMain.hwnd)                  ' restore Display and Coop. modes
 Call ReleaseGame                                   ' unload graphics, classes & etc.
 Call DDRelease                                     ' unload DirectDraw objects
 Call DIRelease                                     ' unload DirectInput objects
 Call DXRelease
 AppendToLog ("Closing DirectX")
 
 'cMouse.SetOldCoords                                ' show back cursor and restore previous position
 ' kill classes
 Set CMouse = Nothing
 Set CKdfGfx = Nothing
 Set CKdfSfx = Nothing
 

 Call CloseLog                                      ' close log file
 
 End
 'Call PostQuitMessage(0)                            ' get use out'f here...NOW!!!
End Sub

