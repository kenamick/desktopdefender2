VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

' --- Load level script

' constants
Private Const SCRIPTSEP1 = ","                           ' paramters separator
Private Const SCRIPTSEP2 = "@"                           ' paramters separator
Private Const SCRIPTCOMMENTS = "//"
Private Const SCRIPTTAG_OPEN = "{"
Private Const SCRIPTTAG_CLOSE = "}"
Private Const SCRIPTAUTHOR = "#authorinfo"
Private Const SCRIPTAUTHOR_NAME = ".name"
Private Const SCRIPTAUTHOR_DATE = ".date"
Private Const SCRIPTLEVEL_NAME = "#levelname"
Private Const SCRIPTLEVEL_NUM = "#levelnum"
Private Const SCRIPTLEVEL_DURATION = "#levelduration"
Private Const SCRIPTLEVEL_DESC = "#levelbrief"
Private Const SCRIPTLEVEL_DESCBKPIC = "#leveldescbkpic"
Private Const SCRIPTLEVEL_DESCDURATION = "#leveldescduration"
Private Const SCRIPTLEVEL_STATE = "#levelstate"
Private Const SCRIPTSTATE_WARPSHIP = 0
Private Const SCRIPTSTATE_TIMEWARPSHIP = 1
Private Const SCRIPTSTATE_WARPMETEOR = 2
Private Const SCRIPTSTATE_TIMEWARPMETEOR = 3
Private Const SCRIPTSTATE_ADDSMQ = 4
Private Const SCRIPTSTATE_CREATEBONUS = 5
Private Const SCRIPTSTATE_GIVEBONUS = 6
Private Const SCRIPTSTATE_DESTROYBUNKER = 7
Private Const SCRIPTSTATE_DESTROYALLBUNKERS = 8

' variables
Public m_strAuthor           As String            ' author name
Public m_strDate             As String            ' date of creation
Public m_strName             As String            ' mission name
Public m_bytID               As Byte              ' mission num
Public m_lDuration           As Long              ' mission duration
Public m_strDesc             As String            ' mission briefing
' privates
Private m_strdescbkpic       As String            ' mission desc.background picture
Private m_ldescduration      As Long              ' description duration
Private m_lTimeBrief         As Long              ' briefing time counter
Private m_arSC(8)            As String
Private m_arData()           As String            ' script file array
Private m_clsState()         As cLevelState       ' level states
Private m_lStates            As Long              ' num of states

'//////////////////////////////////////////////////////////////////
'//// Init stuff
'//////////////////////////////////////////////////////////////////
Public Sub _
Init()

 m_lStates = -1
 ' init state command on the fly (easier checking)
 m_arSC(SCRIPTSTATE_WARPSHIP) = "#warpship"
 m_arSC(SCRIPTSTATE_TIMEWARPSHIP) = "#timewarpship"
 m_arSC(SCRIPTSTATE_WARPMETEOR) = "#warpmeteor"
 m_arSC(SCRIPTSTATE_TIMEWARPMETEOR) = "#timewarpmeteor"
 m_arSC(SCRIPTSTATE_ADDSMQ) = "#addsmq"
 m_arSC(SCRIPTSTATE_CREATEBONUS) = "#createbonus"
 m_arSC(SCRIPTSTATE_GIVEBONUS) = "#givebonus"
 m_arSC(SCRIPTSTATE_DESTROYBUNKER) = "#destroybunker"
 m_arSC(SCRIPTSTATE_DESTROYALLBUNKERS) = "#destroyallbunkers"

 ' reset briefing time counter
 m_lTimeBrief = 0
End Sub

'//////////////////////////////////////////////////////////////////
'//// Load Script map
'//// STRING strScriptName - path&name of the file
'//////////////////////////////////////////////////////////////////
Public Function _
LoadScript(strScriptName As String) As Boolean
 
 ' set error trap
 'On Local Error GoTo LEVELERROR
 
 Dim i         As Long
 Dim j         As Long
 Dim cn        As Long
 Dim lnumLines As Long
 Dim bCmdFound As Boolean                           ' command found flag
 
 ' send all script data to a private array
 If (Not ScriptToArray(strScriptName)) Then GoTo LEVELERROR
 
 ' start decoding
 lnumLines = UBound(m_arData())
 
 ' check4/get author info
 If (InStr(m_arData(0), SCRIPTAUTHOR)) Then
  m_strAuthor = Mid$(m_arData(1), InStr(m_arData(1), SCRIPTAUTHOR_NAME) + Len(SCRIPTAUTHOR_NAME) + 1)
  m_strDate = Mid$(m_arData(2), InStr(m_arData(2), SCRIPTAUTHOR_DATE) + Len(SCRIPTAUTHOR_DATE) + 1)
 End If
 
 i = 0
 Do While (i < lnumLines)
 
  ' check for command (don't check empty or commented lines
  If (Len(m_arData(i)) > 1 And (Left$(m_arData(i), 2) <> SCRIPTCOMMENTS)) Then
   
   ' --- get mission name ---
   If (Left$(m_arData(i), Len(SCRIPTLEVEL_NAME)) = SCRIPTLEVEL_NAME) Then
    m_strName = Mid$(m_arData(i), InStr(m_arData(i), "(") + 1, InStr(m_arData(i), ")") - Len(SCRIPTLEVEL_NAME) - 2)
   End If
   ' --- get mission number ---
   If (Left$(m_arData(i), Len(SCRIPTLEVEL_NUM)) = SCRIPTLEVEL_NUM) Then
    m_bytID = Val(Mid$(m_arData(i), InStr(m_arData(i), "(") + 1, InStr(m_arData(i), ")") - Len(SCRIPTLEVEL_NUM) - 2))
   End If
   ' --- get mission duration ---
   If (Left$(m_arData(i), Len(SCRIPTLEVEL_DURATION)) = SCRIPTLEVEL_DURATION) Then
    m_lDuration = Val(Mid$(m_arData(i), InStr(m_arData(i), "(") + 1, InStr(m_arData(i), ")") - Len(SCRIPTLEVEL_DURATION) - 2))
   End If
   
   ' --- get briefing background picture
   If (Left$(m_arData(i), Len(SCRIPTLEVEL_DESCBKPIC)) = SCRIPTLEVEL_DESCBKPIC) Then
    m_strdescbkpic = Mid$(m_arData(i), InStr(m_arData(i), "(") + 1, InStr(m_arData(i), ")") - Len(SCRIPTLEVEL_DESCBKPIC) - 2)
   End If
   
   ' --- get briefing duration
   If (Left$(m_arData(i), Len(SCRIPTLEVEL_DESCDURATION)) = SCRIPTLEVEL_DESCDURATION) Then
    m_ldescduration = Val(Mid$(m_arData(i), InStr(m_arData(i), "(") + 1, InStr(m_arData(i), ")") - Len(SCRIPTLEVEL_DESCDURATION) - 2))
   End If
   
   ' --- get mission description ---
   If (Left$(m_arData(i), Len(SCRIPTLEVEL_DESC)) = SCRIPTLEVEL_DESC) Then
    ' get strings
    For j = (i + 1) To lnumLines
     If (Len(m_arData(j)) > 1 And (Left$(m_arData(j), 2) <> SCRIPTCOMMENTS)) Then
      m_strDesc = m_strDesc & m_arData(j) & vbCrLf
     End If
     ' check for closing bracket
     If (Left$(m_arData(j), 1) = SCRIPTTAG_CLOSE) Then Exit For
    Next
    ' set new line count
    i = j
   End If
   
   ' --- get state ---
   If (Left$(m_arData(i), Len(SCRIPTLEVEL_STATE)) = SCRIPTLEVEL_STATE) Then
    
    Dim lTemp1 As Long
    Dim lTemp2 As Long
    
    ' get state serial num
    lTemp1 = Val(Mid$(m_arData(i), InStr(m_arData(i), "(") + 1, InStr(m_arData(i), ",") - Len(SCRIPTLEVEL_STATE) - 2))
    ' get state duration
    lTemp2 = Val(Mid$(m_arData(i), InStr(m_arData(i), ",") + 1, InStr(m_arData(i), ")") - Len(SCRIPTLEVEL_STATE) - 2))
    ' create new state
    m_lStates = m_lStates + 1
    ' make place for the new state
    ReDim Preserve m_clsState(m_lStates)
    Set m_clsState(m_lStates) = New cLevelState
    ' init the new state
    Call m_clsState(m_lStates).Init(lTemp1, lTemp2)
    
    ' get state's commands
    For j = (i + 1) To lnumLines
     ' check for comments or empty lines
     If (Len(m_arData(j)) > 1 And (Left$(m_arData(j), 2) <> SCRIPTCOMMENTS)) Then
      
      ' loop trough all known state commands
      For cn = 0 To UBound(m_arSC())
       ' check for a command
       If (Left$(m_arData(j), Len(m_arSC(cn))) = m_arSC(cn)) Then
               
        bCmdFound = True                          ' found command
        
        Dim lparam(6)  As String                  ' parameters
        Dim ncsize     As Integer                 ' command string size
        Dim strTemp    As String
        Dim noffset    As Integer
        Dim cn2        As Integer
        Dim cn3        As Integer
        Dim nLastComma As Integer
         
        ' reset array
        For cn2 = 0 To 6
         lparam(cn2) = ""
        Next
        
        ' get command string size (+1 for the open bracket)
        ncsize = Len(m_arSC(cn)) + 1
        nLastComma = ncsize
        ' reset parameters counter
        cn3 = 0
        
        ' extract params
        For cn2 = ncsize To Len(m_arData(j))
         ' get next char
         strTemp = Mid$(m_arData(j), cn2, 1)
         ' check for a value border
         If (strTemp = SCRIPTSEP1 Or strTemp = ")" Or strTemp = SCRIPTSEP2) Then
          ' extract value
          lparam(cn3) = Mid$(m_arData(j), nLastComma + 1, cn2 - nLastComma - 1)
          ' increment params counter
          cn3 = cn3 + 1
          ' bound
          If (cn3 > 6) Then Exit For
          ' save last found comma position
          nLastComma = cn2
         End If
        
        Next
        
        ' offset to the next comma
        'noffset = InStr(m_arData(j), ",") + 1
        '' only one param found
        'If (noffset = 1) Then
        ' lparam(0) = Val(Mid$(m_arData(j), InStr(m_arData(j), "(") + 1, InStr(m_arData(j), ")")))
        ' Exit For
        'End If
        ' extract paramters
        'lparam(0) = (Mid$(m_arData(j), InStr(m_arData(j), "(") + 1, noffset - ncsize - 2 - 1))
        'lparam(1) = (Mid$(m_arData(j), noffset, InStr(noffset, m_arData(j), ",") - noffset))
        'noffset = InStr(noffset + 1, m_arData(j), ",") + 1 ' offset to the next comma (last comma pos + 1 )
        'lparam(2) = (Mid$(m_arData(j), noffset, InStr(noffset, m_arData(j), ",") - noffset))
        'noffset = InStr(noffset + 1, m_arData(j), ",") + 1
        'lparam(3) = (Mid$(m_arData(j), noffset, InStr(noffset, m_arData(j), ",") - noffset))
        'noffset = InStr(noffset + 1, m_arData(j), ",") + 1
        'lparam(4) = (Mid$(m_arData(j), noffset, InStr(noffset, m_arData(j), ",") - noffset))
        'noffset = InStr(noffset + 1, m_arData(j), ",") + 1
        'lparam(5) = (Mid$(m_arData(j), noffset, InStr(noffset, m_arData(j), ")") - noffset))
        
        ' check for randoms
        For cn2 = 0 To 6
         If (lparam(cn2) = "random") Then lparam(cn2) = "-9999"
        Next
        ' add the new command
        Call m_clsState(m_lStates).AddCommand(cn, _
                                   Val(lparam(0)), Val(lparam(1)), _
                                   Val(lparam(2)), Val(lparam(3)), _
                                   Val(lparam(4)), Val(lparam(5)), _
                                   lparam(6))
        
       End If ' end commands check
      Next ' end commands loop
      
      ' not a valid command(line)
      If (Not bCmdFound) Then
       MsgBox "Invalid command at line: " & (j + 1)
       '...
      Else
       ' reset flag
       bCmdFound = False
      End If
      
     End If ' end comments and empty lines check
     ' check for closing bracket
     If (Left$(m_arData(j), 1) = SCRIPTTAG_CLOSE) Then Exit For
    Next
    ' set new line count
    i = j
   End If
   
   
  End If
 
  i = i + 1
 Loop
 
 ' sort classes
 Call SortStates
 
 ' wow, success
 LoadScript = True
Exit Function

' incase of some nasty error
LEVELERROR:
 LoadScript = False
End Function

'//////////////////////////////////////////////////////////////////
'//// Put script data into an array
'//// STRING strScriptName - path&name of the file
'//////////////////////////////////////////////////////////////////
Private Function _
ScriptToArray(strScriptName As String) As Boolean
 
 On Local Error GoTo LEVELERROR
 
 Dim strBuffer As String
 Dim lLines    As Integer
 Dim nfn       As Integer
  
 ' get free file handle
 nfn = FreeFile
  
 ' open script file
 Open (strScriptName) For Input Access Read Lock Write As #nfn
 
  Do While Not EOF(nfn)
   ' widen data array
   ReDim Preserve m_arData(lLines)
   ' read data
   Line Input #1, strBuffer
   ' copy data
   m_arData(lLines) = strBuffer
   ' increment line counter
   lLines = lLines + 1
  Loop
 
 Close #nfn

 ScriptToArray = True
Exit Function

LEVELERROR:
 ScriptToArray = False
End Function

'//////////////////////////////////////////////////////////////////
'//// Sort States
'//////////////////////////////////////////////////////////////////
Private Sub _
SortStates()
 
 Dim i           As Long
 Dim j           As Long
 Dim clstmpState As New cLevelState

 ' bubble sort states by their 'turn' values
 For i = 1 To m_lStates
  For j = 0 To (m_lStates - 1)
 
   If (m_clsState(i).GetTurn < m_clsState(j).GetTurn) Then
    ' swap classes
    Set clstmpState = m_clsState(i)
    Set m_clsState(i) = m_clsState(j)
    Set m_clsState(j) = clstmpState
    
   End If
   
  Next
 Next

 ' free temp state class
 Set clstmpState = Nothing

End Sub

'//////////////////////////////////////////////////////////////////
'//// Execute level system
'//////////////////////////////////////////////////////////////////
Public Sub _
Update()

  Dim cn             As Long
  Static lTimeBrief  As Long
    
  ' mission briefing
  If (lTimeBrief = 0) Then
   lTimeBrief = GetTicks() + m_ldescduration
   Exit Sub
  ElseIf (lTimeBrief > GetTicks()) Then
   '...
   frmMain.lblStateN = "-1" '{!}
   Exit Sub
  End If
    
  ' execute loaded state
  For cn = 0 To m_lStates
   If (Not m_clsState(cn).Expired) Then
    frmMain.lblStateN = cn '{1}
    Call m_clsState(cn).Execute
    Exit For
   End If
  Next

End Sub


Public Function GetTicks() As Long

 GetTicks = GetTickCount()

End Function

