VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLevelState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum cnstEnemyWarpPosition                     ' enemy warping position constants
  EWP_LEFT = 0
  EWP_RIGHT
  EWP_RANDOM = 255
End Enum

' -- Level State
Private Type stCommand
  lparam1  As Long
  lparam2  As Long
  lparam3  As Long
  lparam4  As Long
  lparam5  As Long
  lparam6  As Long
  lparam7  As String
  lTimer   As Long
  bExpired As Boolean
  type     As cnstCommands
End Type

' list of available script-state commands
Enum cnstCommands
  COM_WARPSHIP = 0
  COM_TIMEWARPSHIP
  COM_WARPMETEOR
  COM_TIMEWARPMETEOR
  COM_ADDSMQ
  COM_CREATEBONUS
  COM_GIVEBONUS
  COM_DESTROYBUNKER
  COM_DESTROYALLBUNKERS
  COM_CREATEBATTLESTATION
  COM_DESTROYBATTLESTATION
  '...
  
End Enum

Private Const PARAM_RANDOM = 255            ' parameter random ID

Private m_lDuration      As Long            ' state duration
Private m_lNumber        As Long            ' state execution turn
Private m_Command()      As stCommand       ' array of state commands
Private m_lCommands      As Long            ' num of commands
Private m_bExpired       As Boolean         ' has this state expired
Private m_lTimeRefresh   As Long            ' execution time counter

'//////////////////////////////////////////////////////////////////
'//// Create State
'//// LONG lDuration - time to last this duration
'//// LONG lnumber   - serial execution number
'//////////////////////////////////////////////////////////////////
Public Sub _
Init(lnumber As Long, lduration As Long)

 m_lDuration = lduration                     ' set state duration
 m_lNumber = lnumber                         ' set state execation turn
 m_lCommands = -1                            ' init number of commands into the dynamic array
 m_bExpired = False                          ' reset expired flag
 m_lTimeRefresh = 0                          ' reset counter timer
 
End Sub

'//////////////////////////////////////////////////////////////////
'//// Add command to the state execution list
'//// cnstCommand tpyCommand - command kind
'//// LONG   lparam1 - paramter
'//// LONG   lparam2 - paramter
'//// LONG   lparam3 - paramter
'//// LONG   lparam4 - paramter
'//// LONG   lparam5 - paramter
'//// LONG   lparam6 - paramter
'//// STRING lparam7 - parameter
'//////////////////////////////////////////////////////////////////
Public Sub _
AddCommand(cnstCommand As cnstCommands, _
           lparam1 As Long, _
           Optional lparam2 As Long = 0, _
           Optional lparam3 As Long = 0, Optional lparam4 As Long = 0, _
           Optional lparam5 As Long = 0, Optional lparam6 As Long = 0, _
           Optional lparam7 As String = "")
  
 ' increment commands counter
 m_lCommands = m_lCommands + 1
 ' make place for a new command
 ReDim Preserve m_Command(m_lCommands)
 ' assign command
 m_Command(m_lCommands).type = cnstCommand
 ' assign paramters
 m_Command(m_lCommands).lparam1 = lparam1
 m_Command(m_lCommands).lparam2 = lparam2
 m_Command(m_lCommands).lparam3 = lparam3
 m_Command(m_lCommands).lparam4 = lparam4
 m_Command(m_lCommands).lparam5 = lparam5
 m_Command(m_lCommands).lparam6 = lparam6
 m_Command(m_lCommands).lparam7 = lparam7           ' string parameter
 m_Command(m_lCommands).lTimer = 0
 m_Command(m_lCommands).bExpired = False
 
End Sub

'//////////////////////////////////////////////////////////////////
'//// Execute this state commands
'//////////////////////////////////////////////////////////////////
Public Sub _
Execute()

 ' exit if this state has expired
 If (m_bExpired) Then Exit Sub
 
 Dim cn              As Long
 
 ' no timer has been set yet
 If (m_lTimeRefresh = 0) Then
  m_lTimeRefresh = m_lDuration + GetTicks()
 ' if state has expired
 ElseIf (m_lTimeRefresh < GetTicks()) Then
  m_bExpired = True
  ' reset the static counter
  m_lTimeRefresh = 0
  Exit Sub
 End If
 
 ' execute state command-logic
 For cn = 0 To m_lCommands
  ' proceed with command
  Call ExecuteCommand(m_Command(cn))
 Next

 '...
End Sub

'//////////////////////////////////////////////////////////////////
'//// Execute command
'//// stCommand objCommand - command object to execute
'//////////////////////////////////////////////////////////////////
Private Sub _
ExecuteCommand(objCommand As stCommand)

 ' exit if this command has expired
 If (objCommand.bExpired) Then Exit Sub
 
 Dim cn   As Integer
 Dim val1 As Long
 Dim val2 As Long
 Dim val3 As Long
 
 Select Case objCommand.type
   
   ' --- warpship
   Case COM_WARPSHIP
     Call CreateEnemy(objCommand.lparam1, CByte(objCommand.lparam2), CByte(objCommand.lparam3))
     objCommand.bExpired = True
     
   ' --- warpship@time
   Case COM_TIMEWARPSHIP
     With objCommand
       ' reset time counter
       'If (.lTimer = 0) Then
       ' .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
       If (.lTimer < GetTicks()) Then
        ' setup timer
        .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
        ' create new ship
        Call CreateEnemy(.lparam3, .lparam4, CByte(.lparam5))
       End If
     End With
   
   ' --- warmeteor
   Case COM_WARPMETEOR
    With objCommand
      ' check paramteres
      If (.lparam1 = PARAM_RANDOM) Then .lparam1 = nGetRnd(1, 2)
      If (.lparam2 = PARAM_RANDOM) Then .lparam2 = MC_LEFT * nGetRnd(1, 2)
      'If (.lparam3 = PARAM_RANDOM) Then .lparam3 = MC_HITEARTH * nGetRnd(1, 2)
      If (.lparam1 = MC_CLOSE) Then .lparam3 = MC_HITMOON _
       Else .lparam3 = MC_HITEARTH
      
      Call CreateMeteor(.lparam1 Or .lparam2 Or .lparam3)
     ' kill command
     .bExpired = True
    End With
   
   ' --- warpmeteor@time
   Case COM_TIMEWARPMETEOR
     With objCommand
       'If (.lTimer = 0) Then
       ' .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
       If (.lTimer < GetTicks) Then
        ' setup timer
        .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
        ' check paramteres
        If (.lparam3 = PARAM_RANDOM) Then .lparam3 = nGetRnd(1, 2)
        If (.lparam4 = PARAM_RANDOM) Then .lparam4 = MC_LEFT * nGetRnd(1, 2)
        'If (.lparam3 = PARAM_RANDOM) Then .lparam3 = MC_HITEARTH * nGetRnd(1, 2)
        If (.lparam3 = MC_CLOSE) Then .lparam5 = MC_HITMOON _
         Else .lparam5 = MC_HITEARTH
        
        ' create meteor
        Call CreateMeteor(.lparam3 Or .lparam4 Or .lparam5)
       End If
     End With
    
   ' --- add scrolling-message to the cockpit
   Case COM_ADDSMQ
     With objCommand
       Call AddSMQ(.lparam7, .lparam1)
      ' Kill it
       .bExpired = True
     End With
     
   ' --- create bonus object@position
   Case COM_CREATEBONUS
     Dim bytType As Byte
     
     ' get bonus type
     If (objCommand.lparam3 = PARAM_RANDOM) Then
      bytType = 255
     ElseIf (objCommand.lparam3 < 0 Or objCommand.lparam3 > BONUS_MAX) Then
      Exit Sub
     Else
      bytType = CByte(objCommand.lparam3)
     End If
     
     ' create the bonus
     Call CreateBonus(CInt(objCommand.lparam1), CInt(objCommand.lparam2), bytType)
     ' kill it
     objCommand.bExpired = True
     
   ' --- create bonus object@position
   Case COM_GIVEBONUS
     
     Dim objBonus As stBonus
     
     ' get bonus type
     If (objCommand.lparam1 = PARAM_RANDOM) Then
      objCommand.lparam1 = 255
     ElseIf (objCommand.lparam1 < 0 Or objCommand.lparam1 > BONUS_MAX) Then
      Exit Sub
     End If
     
     objBonus.kind = CByte(objCommand.lparam1)
     ' give the bonus
     Call GiveBonus(objBonus)
     ' kill it
     objCommand.bExpired = True
     
   ' --- destroy a bunker
   Case COM_DESTROYBUNKER
    ' check for random value
    If (objCommand.lparam1 = PARAM_RANDOM) Then
     CBunker(nGetRnd(0, MAX_BUNKERS)).DoDamage = 100
    Else
     ' check if it's a valid bunker
     If (objCommand.lparam1 > MAX_BUNKERS) Then
      objCommand.lparam1 = MAX_BUNKERS
     ElseIf (objCommand.lparam1 < 0) Then
      objCommand.lparam1 = 0
     End If
      
     CBunker(objCommand.lparam1).DoDamage = 100
    End If
    ' kill command
    objCommand.bExpired = True
     
   ' --- destroy all the bunkers
   Case COM_DESTROYALLBUNKERS
    
    For cn = 0 To MAX_BUNKERS
     CBunker(cn).DoDamage = 100
    Next
    ' kill command
    objCommand.bExpired = True
     
   ' --- create battlestation
   Case COM_CREATEBATTLESTATION
    If (objCommand.lparam1 = 1) Then
     CBattleStation.Create True
    Else
     CBattleStation.Create False
    End If
    objCommand.bExpired = True
    
   ' --- destroy battlestation
   Case COM_DESTROYBATTLESTATION
    CBattleStation.Destroy
    objCommand.bExpired = True
    
   ' --- invalid command
   Case Else
    objCommand.lparam1 = objCommand.lparam1
    
    Debug.Print "STATE:Unknown command detected!"
 
 End Select
 
End Sub

' //////////////////////////////////////////////////////////
' //// Creates random enemy
' //////////////////////////////////////////////////////////
Private Sub _
CreateEnemy(WarpPosition As cnstEnemyWarpPosition, _
            lEnemy As Long, _
            bytNum As Byte)
  
 Dim cn        As Integer
 Dim i         As Integer
 Dim nDirX     As Integer
 Dim xpos      As Integer
 Dim ypos      As Integer
 Static y_temp As Integer
  
 ' bound enemy qunatity
 If (bytNum > 15) Then bytNum = 15
 ' enemy value check
 If (lEnemy < 0) Then lEnemy = 0
 If (lEnemy > SHIPS) Then lEnemy = SHIPS
  
 For i = 0 To (bytNum - 1) ' /* start master loop */
  
  Select Case WarpPosition
  
    ' --- warp from left side (of Earth)
    Case EWP_LEFT
      nDirX = nGetRnd(20, 200)
     
    ' --- warp from right side
    Case EWP_RIGHT
      nDirX = nGetRnd(SCREEN_PIXEL_WIDTH - 100, SCREEN_PIXEL_WIDTH)
     
    ' --- random warp side
    Case EWP_RANDOM
    
      ' get random appearance position
      If (nGetRnd(0, 9) > 4) Then
       nDirX = nGetRnd(20, 200) 'real
      Else
       nDirX = nGetRnd(SCREEN_PIXEL_WIDTH - 100, SCREEN_PIXEL_WIDTH)
      'nDirX = nGetRnd(2160, 2560)
      End If
    
    Case Else
      Debug.Print "Error in LEVELSTATE CLASS, CreateEnemy() function!"
  End Select
  
  ' set y - position
  If (y_temp >= 0) Then
   y_temp = nGetRnd(-165, -45)
  Else
   y_temp = nGetRnd(45, 165)
  End If
  ypos = VISIBLE_AREA_CY_2 + y_temp
  
  ' bound starting position
  If (ypos < 0) Then
   ypos = ypos - y_temp
  ElseIf (ypos > VISIBLE_AREA_CY) Then
   ypos = ypos + y_temp
  End If
  
  ' check for empty enemy class
  Do While (cn < MAX_ENEMIES)
   If (Not CShip(cn).GetVisible) Then
    
    'cShip(cn).CreateShip nDirX, yPos, ST_CARRIER1   '_
    If (lEnemy = 4) Then
     ypos = 25 + nGetRnd(0, 15)
     Call CShip(cn).CreateShip(nDirX, ypos, lEnemy)
    ElseIf (lEnemy <> PARAM_RANDOM) Then
     Call CShip(cn).CreateShip(nDirX, ypos, lEnemy)
    Else
     ' create random ship
     Call CShip(cn).CreateShip(nDirX, ypos, nGetRnd(0, SHIPS))
    End If
    
    Exit Do                                           ' exit here for creation has finished
   End If
 
  cn = cn + 1
  Loop

 Next ' /* end master loop */
 
' Debug.Print "ship n:" & cn & "  pos. app: " & nDirX
End Sub

'//////////////////////////////////////////////////////////////////
'//// Propery to return state's state
'//////////////////////////////////////////////////////////////////
Public Property Get Expired() As Boolean
 Expired = m_bExpired
End Property

'//////////////////////////////////////////////////////////////////
'//// Propery to return state's execution turn
'//////////////////////////////////////////////////////////////////
Public Property Get GetTurn() As Long
 GetTurn = m_lNumber
End Property


'Public Function GetTicks() As Long

' GetTicks = GetTickCount()

'End Function


