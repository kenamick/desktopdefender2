Attribute VB_Name = "mCharacters"
Option Explicit

Public Type stPlayer
  name      As String * 11                           ' id
  pass      As String * 11                           ' char. passwod
  score     As Long
  level     As Byte                                  ' player level
  mission   As Byte                                  ' mission reached
  kills     As Long                                  ' total kills
  success   As Single                                ' %
  nReserved As Integer
End Type

Private Const HS_HEADER = "KKCHAR"
Private Const FILE_NAME = "scores.bin"
Private Const CHARS_MAX = 127

Public g_arChar(CHARS_MAX) As stPlayer
Public g_nChars            As Integer
Public g_strCharFileName   As String

'//////////////////////////////////////////////////////////////
'//// Save all characters
'//////////////////////////////////////////////////////////////
Public Function _
CHARSaveCharacters() As Boolean

 On Local Error GoTo CHARERROR
  
 Dim cn As Integer
 
 ' enrypt all characters
 For cn = 0 To g_nChars
  Call CHAREncryptCharacter(g_arChar(cn))
 Next
 
 ' save all
 If (Len(g_strCharFileName) < 1) Then g_strCharFileName = FILE_NAME
 Open (g_strCharFileName) For Binary Access Write Lock Read As #1
  
  If (g_nChars > CHARS_MAX) Then g_nChars = CHARS_MAX
  
  ' save header
  Put #1, , HS_HEADER
  
  ' save character number
  Put #1, , g_nChars
   
  ' load all characters
  For cn = 0 To g_nChars
   Put #1, , g_arChar(cn)
  Next
  
 Close #1
 
 CHARSaveCharacters = True
Exit Function

CHARERROR:
 CHARSaveCharacters = False
End Function

'//////////////////////////////////////////////////////////////
'//// Loads all characters
'//////////////////////////////////////////////////////////////
Public Function _
CHARLoadCharacters() As Boolean
 
 On Local Error GoTo CHARERROR
 
 Dim i         As Integer
 Dim cn        As Integer
 Dim nSize     As Integer
 Dim strTemp   As String
 Dim strHeader As String
  
 If (Len(g_strCharFileName) < 1) Then g_strCharFileName = FILE_NAME
 Open (g_strCharFileName) For Binary Access Read Lock Write As #1
  
  ' get header
  strHeader = Space$(Len(HS_HEADER))
  Get #1, , strHeader
  If (strHeader <> HS_HEADER) Then GoTo CHARERROR
  
  ' get character number
  Get #1, , g_nChars
  If (g_nChars > CHARS_MAX) Then GoTo CHARERROR

  ' load all characters
  For cn = 0 To g_nChars
   Get #1, , g_arChar(cn)
  Next
  
 Close #1
   
 ' decrypt all characters' data
 For cn = 0 To g_nChars
  ' clear temp string
  strTemp = ""
  ' start character decrypting
  With g_arChar(cn)
   '.name = .name
   ' decrypt password ( only if specified )
   strTemp = Trim$(.pass)
   nSize = Len(strTemp)
   strTemp = ""
   
   If (nSize > 0) Then
    For i = 1 To nSize
     strTemp = strTemp & Chr$(Asc(Mid$(.pass, i, 1)) Xor (.nReserved + i))
    Next
    .pass = strTemp
   End If
   
   ' encrypt tech data
   .score = .score Xor (.nReserved + 10)
   .level = .level Xor (.nReserved + 20)
   .mission = .mission Xor (.nReserved + 30)
   .kills = .kills Xor (.nReserved + 40)
   .success = .success Xor (.nReserved + 50)
  End With
 Next
   
 CHARLoadCharacters = True
Exit Function

CHARERROR:
 '...
 CHARLoadCharacters = False
End Function

'//////////////////////////////////////////////////////////////
'//// Creates a new character
'//// stPlayer objChar - char. object
'//////////////////////////////////////////////////////////////
Public Sub _
CHAREncryptCharacter(objChar As stPlayer)

 Dim cn        As Integer
 Dim nSize     As Integer
 Dim strTemp   As String
 
 ' create/encrypt new player
 With objChar
  ' get random mask
  '.nReserved = Int((Rnd * 10) + 1)
  '.name = Left$(strName, 11)
  ' encrypt password ( only if it's specified )
  nSize = Len(Trim$(.pass))
  If (nSize > 0) Then
   For cn = 1 To nSize
    strTemp = strTemp & Chr$(Asc(Mid$(.pass, cn, 1)) Xor (.nReserved + cn))
   Next
  .pass = strTemp
  End If
  ' encrypt tech data
  .score = .score Xor (.nReserved + 10)
  .level = .level Xor (.nReserved + 20)
  .mission = .mission Xor (.nReserved + 30)
  .kills = .kills Xor (.nReserved + 40)
  .success = .success Xor (.nReserved + 50)
 End With
  
End Sub

'//////////////////////////////////////////////////////////////
'//// Deletes character
'//// INT nChar - char. id to delete
'//////////////////////////////////////////////////////////////
Public Function _
CHARDelCharacter(nChar As Integer) As Boolean

 Dim cn As Integer
 
 ' check number
 If (nChar < 0 Or nChar > CHARS_MAX) Then
  CHARDelCharacter = False
 End If
 
 ' delete and shift
 For cn = nChar To g_nChars
  g_arChar(cn) = g_arChar(cn + 1)
 Next
 
 ' decrease char number
 g_nChars = g_nChars - 1
  
 
 CHARDelCharacter = True
End Function

'//////////////////////////////////////////////////////////////
'//// Creates a new character
'//// STRING strName - char. nick/name
'//// STRING strPass - char. password
'//////////////////////////////////////////////////////////////
Public Function _
CHARNewCharacter(strName As String, strPass As String, _
                 Optional lscore As Long = 0, _
                 Optional bytLevel As Byte = 1, _
                 Optional bytMission As Byte = 1, _
                 Optional lkills As Long = 0, _
                 Optional fSuccess As Single = 0#) As Boolean

 On Local Error GoTo CHARERROR
 
 ' check character number
 If ((g_nChars + 1) > CHARS_MAX) Then GoTo CHARERROR
 ' increase character number
 g_nChars = g_nChars + 1
 
 ' fill data
 With g_arChar(g_nChars)
  ' get random mask
  .nReserved = Int((Rnd * 10) + 1)
  .name = Trim$(strName)
  .pass = Trim$(strPass)
  ' add optionals
  If (lscore <> 0) Then .score = lscore
  If (bytLevel <> 1) Then .level = bytLevel
  If (bytMission <> 1) Then .mission = bytMission
  If (lkills <> 0) Then .kills = lkills
  If (fSuccess <> 0#) Then .success = fSuccess
 End With
 
 ' encrypt char. data
 'Call EncryptCharacter(g_arChar(g_nChars))
 
 '' save new character
 'Open (FNAME) For Binary Access Write Lock Read As #1
 
 ' Put #1, , g_nChars
  
 ' Seek #1, LOF(1) + 1
 ' Put #1, , objPlayer
 
 'Close #1
  
 CHARNewCharacter = True
Exit Function

CHARERROR:
 CHARNewCharacter = False
End Function



