VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem *** The FSM-AI class
Rem *** Purpose: Create enemies, determine ships, _
                 bunkers and other objects behavior


Private m_lAIUpdateTime As Long                          ' update AI delay
Private lCSTime1      As Long                            ' create enemy delay
Private lCSTime2      As Long
Private lCSTime3      As Long
Private lCSTime4      As Long
Private bFlag1        As Boolean
Private bytLevelState As Byte
Private m_bytlevel As Byte

' //////////////////////////////////////////////////////////
' //// main AI refreshing
' //////////////////////////////////////////////////////////
Public Sub _
Update()
 
 Static lUBTime       As Long                             ' update bunkers counter
 Dim cn               As Integer
 
 If (m_lAIUpdateTime < GetTicks) Then
     m_lAIUpdateTime = GetTicks + FPS_AI
    
    
    Select Case m_bytlevel
    
      ' --- level 0 --- the Demo level ------------
      Case 0
       ' see if enemy1 should be created
       'If (GetTicks > lCSTime1) Then
        'lCSTime1 = GetTicks + nGetRnd(2500, 8000)
       ' lCSTime1 = GetTicks + nGetRnd(3500, 5000)
       ' Call CreateEnemy(ST_LRSNEAKY)
       'End If
       'If (GetTicks > lCSTime2) Then
       ' lCSTime2 = GetTicks + nGetRnd(3000, 4500)
       ' Call CreateEnemy(ST_SEVENTHFOX)
       'End If
       If (GetTicks > lCSTime2) Then
        lCSTime2 = GetTicks + nGetRnd(1200, 5500)
        Call CreateEnemy(ST_PARTICLEBEAST)
       End If
      
    
      ' --- level 1 --- the fast strike -----------
      Case 1
           
        Select Case bytLevelState
             
          ' phase 1 - the first encounter
          Case 1
           
           If (GetTicks > lCSTime1) Then
            
            For cn = 0 To 3
             Call CreateEnemy(ST_LRSNEAKY)
             Call CreateEnemy(ST_SEVENTHFOX)
            Next
               
            ' activate phase 2
            bytLevelState = 2
          End If
          
         
         ' phase 2 - random slayers
         Case 2
         
          ' see if enemy1 should be created
          If (GetTicks > lCSTime2) Then
           lCSTime2 = GetTicks + nGetRnd(3800, 4100)
           Call CreateEnemy(ST_LRSNEAKY)
          End If
          ' see if enemy2 should be created
          If (GetTicks > lCSTime3) Then
           lCSTime3 = GetTicks + nGetRnd(2500, 2800)
           Call CreateEnemy(ST_SEVENTHFOX)
          End If
        
        End Select
        
      ' --- level 2 --- meteor shower -------------
      Case 2
      '...
      
      ' --- level 3 --- big homeys ----------------
      Case 3
      '...
      
      ' --- level 4 --- defend the station --------
      Case 4
      '...
      
      ' --- level 5 --- fight for existance -------
      Case 5
      '...
      
      Case Else
       Exit Sub
    
    End Select
    
 End If
 
End Sub

' //////////////////////////////////////////////////////////
' //// Creates random enemy
' //////////////////////////////////////////////////////////
Public Sub _
SetupLevel(bytGameLevel As Byte)
 
 ' reset vars
 m_lAIUpdateTime = 0
 lCSTime1 = 0
 lCSTime2 = 0
 lCSTime3 = 0
 lCSTime4 = 0
 ' set current mission
 m_bytlevel = bytGameLevel
 ' set default levelstate
 bytLevelState = 1
  
 Select Case m_bytlevel
   
   Case 0
   '...
   
   Case 1
    lCSTime1 = 5000 + GetTicks
    bytLevelState = 1
     
   Case 2
   '...
   
   Case 3
   '...
   
   Case 4
   '...
   
   Case 5
   '...
 
 End Select
 

End Sub


' //////////////////////////////////////////////////////////
' //// Creates random enemy
' //////////////////////////////////////////////////////////
Private Sub _
CreateEnemy(bytEnemy As Byte)
  
 ' --------------!
 'Static ar As Integer
 'If ar > 2 Then Exit Sub
 'ar = ar + 1
 ' !--------------
 Dim cn As Integer
 Dim nDirX As Integer
 Dim xPos As Integer, yPos As Integer
 Dim nRandMin As Integer, nRandMax As Integer
 
  ' get random appearance position
  If nGetRnd(0, 1) = 0 Then
   nDirX = nGetRnd(20, 200) 'real
  Else
   nDirX = nGetRnd(SCREEN_PIXEL_WIDTH - 100, SCREEN_PIXEL_WIDTH)
   'nDirX = nGetRnd(2160, 2560)
  End If
  yPos = nGetRnd(0, VISIBLE_AREA_CY - 150)
  
  ' get type of ship to appear
  If nGetRnd(0, 100) < 25 Then   ' create close - ship
   nRandMin = 3
   nRandMax = SHIPS
  Else                            ' create far ship
   nRandMin = 0
   nRandMax = 2
 End If
  
  ' check for empty enemy class
 Do While cn < MAX_ENEMIES
  If (Not cShip(cn).GetVisible) Then
   
   'cShip(cn).CreateShip nDirX, yPos, ST_CARRIER1   '_
   If (bytEnemy <> 255) Then
    Call cShip(cn).CreateShip(nDirX, yPos, bytEnemy)
   Else
    ' create random ship
    Call cShip(cn).CreateShip(nDirX, yPos, nGetRnd(0, SHIPS))
   End If
   
   Exit Do                                           ' exit here for creation has finished
  End If
  
 cn = cn + 1
 Loop
 
' Debug.Print "ship n:" & cn & "  pos. app: " & nDirX
End Sub



