VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLevelState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

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
  '...
  COM_BATTLESTATION
End Enum

Private Const PARAM_RANDOM = -9999          ' parameter random ID

Private m_lDuration      As Long            ' state duration
Private m_lNumber        As Long            ' state execution turn
Private m_Command()      As stCommand       ' array of state commands
Private m_lCommands      As Long            ' num of commands
Private m_bExpired       As Boolean         ' has this state expired
    
'//////////////////////////////////////////////////////////////////
'//// Create State
'//// LONG lDuration - time to last this duration
'//// LONG lnumber   - serial execution number
'//////////////////////////////////////////////////////////////////
Public Sub _
Init(lnumber As Long, lduration As Long)

 m_lDuration = lduration
 m_lNumber = lnumber
 m_lCommands = -1
 m_bExpired = False

End Sub

'//////////////////////////////////////////////////////////////////
'//// Add command to the state execution list
'//// cnstCommand tpyCommand - command kind
'//// LONG lparam1 - paramter
'//// LONG lparam2 - paramter
'//// LONG lparam3 - paramter
'//// LONG lparam4 - paramter
'//// LONG lparam5 - paramter
'//// LONG lparam6 - paramter
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
 
 Static lTimeRefresh As Long
 Dim cn              As Long
 
 ' no timer has been set yet
 If (lTimeRefresh = 0) Then
  lTimeRefresh = m_lDuration + GetTicks()
 ' if state has expired
 ElseIf (lTimeRefresh < GetTicks()) Then
  m_bExpired = True
  ' reset the static counter
  lTimeRefresh = 0
 ' execute state command-logic
 ElseIf (lTimeRefresh > GetTicks()) Then
  For cn = 0 To m_lCommands
   ' proceed with command
   Call ExecuteCommand(m_Command(cn))
  Next
  
 End If
  

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

 Select Case objCommand.type
   
   ' --- warpship
   Case COM_WARPSHIP
     frmMain.lblAction = "CREATing ship with params: " & vbCrLf & _
                              objCommand.lparam1 & " " & objCommand.lparam2 & " " & vbCrLf & _
                              objCommand.lparam3 & " " & objCommand.lparam4 & " " & vbCrLf & _
                              objCommand.lparam5 & " " & objCommand.lparam6
     objCommand.bExpired = True
     
   ' --- warpship@time
   Case COM_TIMEWARPSHIP
     With objCommand
       If (.lTimer = 0) Then
        .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
       ElseIf (.lTimer < GetTicks) Then
        .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
        frmMain.lblAction.Caption = "CREATing TIME ship with params: " & vbCrLf & _
                              objCommand.lparam1 & " " & objCommand.lparam2 & " " & vbCrLf & _
                              objCommand.lparam3 & " " & objCommand.lparam4 & " " & vbCrLf & _
                              objCommand.lparam5 & " " & objCommand.lparam6 & " " & .lTimer
       End If
     End With
   
   ' --- warmeteor
   Case COM_WARPMETEOR
    With objCommand
     ' Call createmeteor(.lparam1 Or .lparam2 Or .lparam3)
     ' kill command
     .bExpired = True
    End With
   
   ' --- warpmeteor@time
   Case COM_TIMEWARPMETEOR
     With objCommand
       If (.lTimer = 0) Then
        .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
       ElseIf (.lTimer < GetTicks) Then
        .lTimer = GetTicks + (((.lparam2 - .lparam1) * Rnd) + .lparam1)
        frmMain.lblAction.Caption = "CREATing TIME METEO with params: " & vbCrLf & _
                              objCommand.lparam1 & " " & objCommand.lparam2 & " " & vbCrLf & _
                              objCommand.lparam3 & " " & objCommand.lparam4 & " " & vbCrLf & _
                              objCommand.lparam5 & " " & objCommand.lparam6 & " " & .lTimer
       End If
     End With
    
   ' --- add scrolling-message to the cockpit
   Case COM_ADDSMQ
     
     With objCommand
      ' time it
      If (.lTimer = 0) Then
       .lTimer = GetTicks() + .lparam1
      ' show it
      ElseIf (.lTimer > GetTicks()) Then
       frmMain.lblAction = .lparam7
      ' kill it
      ElseIf (.lTimer < GetTicks()) Then
       .bExpired = True
      End If
     End With
     
   ' --- create bonus object@position
   Case COM_CREATEBONUS
     Dim bytType As Byte
     ' get bonus type
     If (objCommand.lparam3 = PARAM_RANDOM) Then bytType = 255 _
      Else bytType = CByte(objCommand.lparam3)
     ' create the bonus
     'Call createbonus(objCommand.lparam1, objCommand.lparam2, byttype)
     ' kill it
     objCommand.bExpired = True
     
   ' --- create bonus object@position
   Case COM_GIVEBONUS
     Dim bytType1  As Byte
     'dim objBonus as stbonus
     
     ' get bonus type
     'If (objCommand.lparam1 = PARAM_RANDOM) Then bytType = 255 _
      Else bytType = CByte(objCommand.lparam1)
     
     'objbonus.kind = byttype
     ' give the bonus
     'call givebonus(objbonus)
     ' kill it
     objCommand.bExpired = True
     
   ' --- destroy a bunker
   Case COM_DESTROYBUNKER
    ' check for random value
    If (objCommand.lparam1 = PARAM_RANDOM) Then
    ' cbunker(ngetrnd(0,max_bunkers)).dodamage = 100
    Else
    ' cbunker(objcommand.lparam1).dodamage = 100
    End If
    ' kill command
    objCommand.bExpired = True
     
   ' --- destroy all the bunkers
   Case COM_DESTROYALLBUNKERS
    
    objCommand.lparam1 = objCommand.lparam1
    'Dim cn As Integer
   
    'For cn = 0 To max_bunkers
    ' cbunker(cn).dodamage = 100
    'Next
    ' kill command
    objCommand.bExpired = True
   
     
   Case Else
    objCommand.lparam1 = objCommand.lparam1
    Stop
    
    Debug.Print "STATE:Unknown command detected!"
 
 End Select
 
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


Public Function GetTicks() As Long

 GetTicks = GetTickCount()

End Function


