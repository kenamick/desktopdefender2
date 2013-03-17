Attribute VB_Name = "mCharacters"
Option Explicit

'-------------------------------------
'--> Characters Info Module
'-------------------------------------

Public Type stPlayer
  name      As String * 18                           ' id
  pass      As String * 18                           ' char. passwod
  score     As Long
  level     As Byte                                  ' player level
  mission   As Byte                                  ' mission reached
  kills     As Long                                  ' total kills
  ts        As Long                                  ' total shots
  Exp       As Long                                  ' experience
  'success   As Single                                ' %
  nReserved As Integer
  sit       As Long                                  ' shots in target
  cheater   As Byte
End Type
' size = 64 = 2^6 - power of 2 works faster with Ram
 

Private Const HS_HEADER = "KKCHAR"
Private Const FILE_NAME = "scores.bin"
Private Const CHARS_MAX = 127
Private strCharsPath       As String                 ' folder that contains characters

Public g_arChar(CHARS_MAX) As stPlayer
Public g_nChars            As Integer
Public g_arRanks()         As String
'Public g_strCharFileName   As String


'//////////////////////////////////////////////////////////////
'//// Convert char experience to equal level
'//// LONG lExperience - char. experience
'//////////////////////////////////////////////////////////////
Public Function _
CHARExperienceToLevel(lExperience As Long) As Byte

  CHARExperienceToLevel = CByte((lExperience / 2000))

End Function


'//////////////////////////////////////////////////////////////
'//// Convert char levelrang to the string equivalent
'//// BYTE bytLevelRank - player rank
'//////////////////////////////////////////////////////////////
Public Function _
CHARLevelToString(bytLevelRank As Byte) As String

  
 CHARLevelToString = g_arRanks(bytLevelRank)
 'Dim i As Byte, j As Byte

 'For i = 0 To UBound(g_arRanks())
 ' For j = i To i + 3
 '  If (bytLevelRank = j) Then
 '   CHARLevelToString = g_arRanks(i)
 '   Exit Function
 '  End If
 ' Next
 'Next

 'CHARLevelToString = UBound(g_arRanks())

 'Select Case (bytLevelRank)

 '   Case 0, 1, 2: CHARLevelToString = "Mama's boy"
 '   Case 3, 4: CHARLevelToString = "Lurid Marine"
 '   Case 5, 6: CHARLevelToString = "Private Buttlick"
 '   Case 7, 8: CHARLevelToString = "Corporal Metalsuck"
 '   Case 9, 10: CHARLevelToString = "Leiutenant BallsKeeper"
 '   Case 11, 12: CHARLevelToString = "Captain Commando"
 '   Case 13, 14: CHARLevelToString = "Captain Destuctou'"
 '   Case 15, 16: CHARLevelToString = "Commander Shitonem"
 '   Case 17, 18: CHARLevelToString = "Bunker Lurid"
 '   Case 19, 20: CHARLevelToString = "Space Ace"
 '   Case 21, 22: CHARLevelToString = "Hidden Death"
 '   Case 23, 24: CHARLevelToString = "Death lives!"
 '
 '   Case Else
 '     CHARLevelToString = "!!!GOD!!!"
 '
 'End Select

End Function


'//////////////////////////////////////////////////////////////
'//// Sort characters by score
'//////////////////////////////////////////////////////////////
Public Sub _
CHARSortCharacters()

 Dim cn      As Integer
 Dim bloop   As Boolean
 Dim tmpChar As stPlayer
 
 bloop = True
 Do While (bloop)
 
  bloop = False
 
  For cn = 0 To g_nChars - 1
    
    If (g_arChar(cn + 1).Exp > g_arChar(cn).Exp And (g_arChar(cn + 1).cheater <> 123)) Then
     'swap
     tmpChar = g_arChar(cn)
     g_arChar(cn) = g_arChar(cn + 1)
     g_arChar(cn + 1) = tmpChar
    End If
    
  Next
 
 Loop


End Sub


'//////////////////////////////////////////////////////////////
'//// Save all characters
'//////////////////////////////////////////////////////////////
Public Function _
CHARSaveCharacters() As Boolean

 On Local Error GoTo CHARERROR
  
 Dim cn    As Integer
 Dim hFile As Integer
 
 ' sort characters
 Call mCharacters.CHARSortCharacters
 
 ' get free file handle
 hFile = FreeFile()
 
 ' enrypt all characters
 For cn = 0 To g_nChars
  Call CHAREncryptCharacter(g_arChar(cn))
 Next
 
 ' save all
 'If (Len(g_strCharFileName) < 1) Then g_strCharFileName = FILE_NAME
 'Open (g_strCharFileName) For Binary Access Write Lock Read As #hFile
  
  If (g_nChars > CHARS_MAX) Then g_nChars = CHARS_MAX
  
 ' ' save header
 ' Put #hFile, , HS_HEADER
  
 ' ' save character number
 ' Put #hFile, , g_nChars
   
 ' ' save all characters
 ' For cn = 0 To g_nChars
 '  Put #hFile, , g_arChar(cn)
 ' Next
  
 'Close #hFile
 
 ' save all char files
 For cn = 0 To g_nChars
  Open (strCharsPath & "\" & Trim$(g_arChar(cn).name) & ".ddc") For Binary Access Write As #hFile
   
   ' save header
   Put #hFile, , HS_HEADER
   
   ' save data
   Put #hFile, , g_arChar(cn)
   
  Close #hFile
 Next
 

 CHARSaveCharacters = True
Exit Function

CHARERROR:
 CHARSaveCharacters = False
End Function


'//////////////////////////////////////////////////////////////
'//// Loads all characters
'//// STRING strPath - path to search for characters' files
'//////////////////////////////////////////////////////////////
Public Function _
CHARLoadCharacters(strPath As String) As Boolean
 
 On Local Error GoTo CHARERROR
 
 Dim i         As Integer
 Dim cn        As Integer
 Dim nSize     As Integer
 Dim hFile     As Integer
 Dim strTemp   As String
 Dim strNext   As String
 Dim strHeader As String
   
 ' reset char counter
 g_nChars = -1
   
 ' get free file handle
 hFile = FreeFile()
  
 ' assign to local char-folder
 strCharsPath = strPath
  
 ' get first entry
 strNext = Dir(strCharsPath & "\*.ddc")
 
 ' loop untill no files
 Do While (strNext <> vbNullString)
 
 '/* old method 1 mother file
 'If (Len(g_strCharFileName) < 1) Then g_strCharFileName = FILE_NAME
 'Open (g_strCharFileName) For Binary Access Read Lock Write As #hFile
 '*/
   ' get header
   'strHeader = Space$(Len(HS_HEADER))
   'Get #hFile, , strHeader
   'If (strHeader <> HS_HEADER) Then GoTo CHARERROR
  
   '/* get character number
   'Get #hFile, , g_nChars
   'If (g_nChars > CHARS_MAX) Then GoTo CHARERROR
   '*/
   
   ' load all characters
   'For cn = 0 To g_nChars
   ' Get #hFile, , g_arChar(cn)
   'Next
  
  ' increment characters counter
  g_nChars = g_nChars + 1
  If (g_nChars > CHARS_MAX) Then g_nChars = CHARS_MAX 'GoTo CHARERROR
  
  ' open the file
  Open (strCharsPath & "\" & strNext) For Binary Access Read Lock Write As #hFile
    
   ' compare header
   strHeader = Space$(Len(HS_HEADER))
   Get #hFile, , strHeader
   If (strHeader <> HS_HEADER) Then GoTo CHARERROR

   ' get raw data
   Get #hFile, , g_arChar(g_nChars)
   
  Close #hFile
   
  ' get next file
  strNext = Dir()
  
 Loop
   
 ' no characters found -> raise error
 If (g_nChars < 0) Then GoTo CHARERROR
   
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
   
   ' decrypt tech data
   .score = .score Xor (.nReserved + 10)
   .level = .level Xor (.nReserved + 20)
   .mission = .mission Xor (.nReserved + 30)
   .ts = .ts Xor (.nReserved + 40)
   .kills = .kills Xor (.nReserved + 50)
   '.success = .success Xor (.nReserved + 60)
   .Exp = .Exp Xor (.nReserved + 60)
   .sit = .sit Xor (.nReserved + 70)
   .cheater = .cheater Xor (.nReserved + 80)
  
   ' check for cheaters and reset their level
   If (.cheater = 123) Then .level = 0
  End With
 Next
   
  
 ' sort characters
 Call mCharacters.CHARSortCharacters
 
 CHARLoadCharacters = True
Exit Function

CHARERROR:
 Close #hFile
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
  .ts = .ts Xor (.nReserved + 40)
  .kills = .kills Xor (.nReserved + 50)
  .Exp = .Exp Xor (.nReserved + 60)
  .sit = .sit Xor (.nReserved + 70)
  .cheater = .cheater Xor (.nReserved + 80)
  '.success = .success Xor (.nReserved + 60)
 End With
  
End Sub


'//////////////////////////////////////////////////////////////
'//// Deletes character
'//// INT nChar - char. id to delete
'//////////////////////////////////////////////////////////////
Public Function _
CHARDelCharacter(nChar As Integer) As Boolean

 Dim cn       As Integer
 Dim strfName As String
 
 ' check number
 If (nChar < 0 Or nChar > CHARS_MAX Or _
     (g_nChars - 1) < 0) Then
  CHARDelCharacter = False
  Exit Function
 End If
 
 ' delete file
 strfName = strCharsPath & "\" & Trim$(g_arChar(nChar).name) & ".ddc"
 
 If (DeleteFile(strfName) = 0) Then
  CHARDelCharacter = False
'  Stop
  Exit Function
 End If
 
 ' remove from array and rearrange
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
                 Optional bytMission As Byte = 0, _
                 Optional lTotalShots As Long = 0, _
                 Optional lkills As Long = 0, _
                 Optional lExperience As Long = 0) As Boolean
                 'Optional fSuccess As Single = 0#) As Boolean

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
  If (lTotalShots <> 0) Then .ts = lTotalShots
  If (lkills <> 0) Then .kills = lkills
  If (lExperience <> 0) Then .Exp = lExperience
  'If (fSuccess <> 0#) Then .success = fSuccess
  
  ' check for cheater
  If (Trim$(.name) = "cheater") Then
   .cheater = 123  ' this player's a cheater
  Else
   .cheater = 321  ' no cheating
  End If
   
 End With
 
 ' encrypt char. data
 'Call EncryptCharacter(g_arChar(g_nChars))
 
 '' save new character
 'Open (FNAME) For Binary Access Write Lock Read As #1
 
 ' Put #1, , g_nChars
  
 ' Seek #1, LOF(1) + 1
 ' Put #1, , objPlayer
 
 'Close #1
 Call CHARSaveCharacters
 Call CHARLoadCharacters(App.Path & "\cadets")
 
 CHARNewCharacter = True
Exit Function

CHARERROR:
 CHARNewCharacter = False
End Function



