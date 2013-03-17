VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKDF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------
'--> KenamicK Data File Pack 2
'--> by Peter "Pro-XeX" Petrov
'--> KenamicK Entertainment 1998-2002
'-------------------------------------

Private Type stKDFile
  strName   As String                           ' file name
  strInfo   As String                           ' extended file info
  lPosition As Long                             ' position in packet
End Type

Private Const KDF_IDENTITY = "KDF"              ' packet identity
Private Const KDF_TAGIDENTITY = "KDF_TAG"       ' packet identity
Private Const KDF_VERSION = 20                  ' packet version (2.0)

Dim m_arKDFile()     As stKDFile                ' array of all the files
Dim m_arKDFilePath() As String                  ' file path ( not part of the main data structure )
Dim m_lnumFiles      As Long                    ' number of files in packet
Dim m_strPacketName  As String                  ' packet file name
Dim m_strLog         As String                  ' log messages
Dim m_bTagOpened     As Boolean                 ' tag opened flag

'////////////////////////////////////////////////////////////////
'//// On initialization
'////////////////////////////////////////////////////////////////
Private Sub _
Class_Initialize()
 
 AppendToLog ("Packeger Class Ready.")
 Call Reset
   
End Sub


'////////////////////////////////////////////////////////////////
'//// On termination
'////////////////////////////////////////////////////////////////
Private Sub _
Class_Terminate()
 
 Call Reset
 
End Sub


'////////////////////////////////////////////////////////////////
'//// List all entires
'//// LISTBOX  lstBox - a valid listbox object
'////////////////////////////////////////////////////////////////
Public Sub _
ListAll(lstBox As ListBox)

 Dim cn As Long
 
 ' clear previous contents
 lstBox.Clear
 ' list entries
 For cn = 0 To m_lnumFiles
  lstBox.AddItem (m_arKDFile(cn).strName)
 Next

End Sub

'////////////////////////////////////////////////////////////////
'//// Clear all arrays
'////////////////////////////////////////////////////////////////
Private Sub _
Reset()
 
 AppendToLog ("Reseting arrays and vars.")
 
 ' reset file counter
 m_lnumFiles = -1
 ' reset tag flag
 m_bTagOpened = False
 
 ' clear array
 Erase m_arKDFile()
 Erase m_arKDFilePath()

End Sub

'////////////////////////////////////////////////////////////////
'//// Add message to local log
'//// STRING strMsg - text message to add
'////////////////////////////////////////////////////////////////
Private Sub _
AppendToLog(strMsg As String)
 
 m_strLog = m_strLog & vbCrLf & strMsg

End Sub

'////////////////////////////////////////////////////////////////
'//// Load Tag
'//// STRING strFileName - path&filename to load
'////////////////////////////////////////////////////////////////
Public Function _
LoadTag(strFileName As String) As Boolean

 ' tarp local errors
 On Local Error GoTo KDFTAGLOADERROR
 
 Dim cn          As Integer
 Dim hFile       As Integer
 Dim strIdentity As String
 
 ' reset arrays and counter
 Call Reset
 
 ' get free file handle
 hFile = freefile()
  
 AppendToLog ("Opening tag " & strFileName)
 
 Open (strFileName) For Binary Access Read As #hFile
  
  ' get file identity
  strIdentity = Space$(Len(KDF_TAGIDENTITY))
  Get #hFile, , strIdentity
  If (strIdentity <> KDF_TAGIDENTITY) Then
   AppendToLog ("Invalid tag file!")
   GoTo KDFTAGLOADERROR
  End If
  
  ' get num of files
  Get #hFile, , m_lnumFiles
  
  ' resize array
  ReDim m_arKDFile(m_lnumFiles)
  
  ' load all entries
  For cn = 0 To m_lnumFiles
   Get #hFile, , m_arKDFile(cn)
  Next
 Close #hFile
 
 m_bTagOpened = True
 LoadTag = True
 AppendToLog ("Tag Opened...")
Exit Function

KDFTAGLOADERROR:
 LoadTag = False
 AppendToLog ("Error opening tag " & strFileName)
 m_bTagOpened = False
End Function


'////////////////////////////////////////////////////////////////
'//// Load packet
'//// STRING strFileName - path&filename to load
'////////////////////////////////////////////////////////////////
Public Function _
LoadPacket(strFileName As String) As Boolean

 ' trap local errors
 On Local Error GoTo KDFLOADERROR
 
 Dim hFile       As Integer
 Dim hFile2      As Integer
 Dim strIdentity As String
 Dim nVersion    As Integer
 
 ' see if tag is opened
 If (Not m_bTagOpened) Then
  AppendToLog ("Tag file is not opened!")
  GoTo KDFLOADERROR
 End If
 
 AppendToLog ("Opening packet " & strFileName)
 
 hFile = freefile()
 
 ' open packet
 Open (strFileName) For Binary Access Read As #hFile
  
  ' get identity
  strIdentity = Space$(Len(KDF_IDENTITY))
  Get #hFile, , strIdentity
  ' check identity
  If (strIdentity <> KDF_IDENTITY) Then
   AppendToLog ("Not a valid KDF file!")
   GoTo KDFLOADERROR
  End If
  ' get version
  Get #hFile, , nVersion
  If (nVersion > KDF_VERSION) Then
   AppendToLog ("KDF Packet version is higher than the supported!")
   GoTo KDFLOADERROR
  End If
    
  '...
  
 Close #hFile
 
 ' set packet name
 m_strPacketName = strFileName
 ' success
 LoadPacket = True
 AppendToLog ("Loading packet successful...")

Exit Function

KDFLOADERROR:
 '...
 AppendToLog ("Error loading packet " & strFileName)
 LoadPacket = False
End Function

'////////////////////////////////////////////////////////////////
'//// Save Tag
'//// STRING strFileName - path&filename to save
'////////////////////////////////////////////////////////////////
Public Function _
SaveTag(strFileName As String) As Boolean

 ' tarp local errors
 On Local Error GoTo KDFTAGSAVERROR
 
 Dim cn     As Integer
 Dim hFile  As Integer
 
 ' check enties count
 If (m_lnumFiles < 0) Then
  AppendToLog ("No files in packet!")
  Exit Function
 End If
 
 ' get free file handle
 hFile = freefile()
 
 AppendToLog ("Saving tag " & strFileName)
 
 ' check for extension
 If (Not CBool(InStr(Right$(strFileName, 4), "."))) Then
  strFileName = strFileName & ".kdf"
 End If

 Open (strFileName) For Binary Access Write As #hFile
  
  ' save tag
  Put #hFile, , KDF_TAGIDENTITY
  
  ' save num of entires
  Put #hFile, , m_lnumFiles
  
  ' save all entries
  For cn = 0 To m_lnumFiles
   Put #hFile, , m_arKDFile(cn)
  Next
 Close #hFile
 
 SaveTag = True
 AppendToLog ("Saving tag successful...")
Exit Function

KDFTAGSAVERROR:
 SaveTag = False
 AppendToLog ("Error saving tag " & strFileName)
End Function

'////////////////////////////////////////////////////////////////
'//// Save packet
'//// STRING strFileName - path&filename to save
'////////////////////////////////////////////////////////////////
Public Function _
SavePacket(strFileName As String) As Boolean

 ' trap local errors
 On Local Error GoTo KDFSAVERROR
 
 Dim cn       As Long
 Dim hFile    As Integer
 Dim hFile2   As Integer
 Dim arData() As Byte
  
 ' check enties count
 If (m_lnumFiles < 0) Then
  AppendToLog ("No files in packet!")
  Exit Function
 End If
 
 ' get free file handles
 hFile = freefile()
 hFile2 = freefile() + 1
 
 AppendToLog ("Saving packet " & strFileName)
 
 ' check for extension
 If (Not CBool(InStr(Right$(strFileName, 4), "."))) Then
  strFileName = strFileName & ".kdf"
 End If
 
 Open (strFileName) For Binary Access Write As #hFile
   
  ' save header
  Put #hFile, , KDF_IDENTITY
  ' save version
  Put #hFile, , KDF_VERSION
  
  For cn = 0 To m_lnumFiles
   
   ' open KDFile entry
   Open (m_arKDFilePath(cn)) For Binary Access Read As #hFile2
    ' resize data array
    ReDim arData((LOF(hFile2) - 1))
    ' get data
    Get #hFile2, , arData()
   Close #hFile2
      
   ' save position in packet
   m_arKDFile(cn).lPosition = Seek(hFile)
   ' save data
   Put #hFile, , arData()
   
  Next cn
      
 Close #hFile
 
 SavePacket = True
 AppendToLog ("Packet saved successfully...")
 
Exit Function

KDFSAVERROR:
 '...
 SavePacket = False
 AppendToLog ("Error saving packet " & strFileName)
End Function

'////////////////////////////////////////////////////////////////
'//// Extarct File from packet
'//// STRING strFileName - path&filename to load
'//// STRING strDestPath - destination path&
'///// ...not finished
'////////////////////////////////////////////////////////////////
Public Function _
ExtractFile(strFileName As String, strDestPath As String) As Boolean

 On Local Error GoTo EXFILERROR
 
 '...
 
EXFILERROR:
 '...
 ExtractFile = False
End Function


'////////////////////////////////////////////////////////////////
'//// Delete file from packet
'//// STRING strFileName - path&filename to delete from packet
'////////////////////////////////////////////////////////////////
Public Function _
DeleteFile(strFileName As String)
   
 Dim i      As Long
 Dim j      As Long
 Dim bFound As Boolean
 
  ' check enties count
 If (m_lnumFiles < 0) Then
  AppendToLog ("No files in packet!")
  Exit Function
 End If
 
 AppendToLog ("Removing " & strFileName)
  
 i = 0
 j = 0
 
 ' loop trough all entries
 Do While (i <= m_lnumFiles)
  
  ' look for file
  If (m_arKDFile(i).strName = strFileName) Then
   '' exchangle files & paths
   'm_arKDFile(i) = m_arKDFile(m_lnumFiles)
   'm_arKDFilePath(i) = m_arKDFilePath(m_lnumFiles)
   bFound = True
   
   ' kill this entry
   For j = i To (m_lnumFiles - 1)
    m_arKDFile(i) = m_arKDFile(i + 1)
    m_arKDFilePath(i) = m_arKDFilePath(i + 1)
   Next
      
   ' decrement file counter
   m_lnumFiles = m_lnumFiles - 1
   ' reset if less than 0
   If (m_lnumFiles < 0) Then
    Call Reset
   End If
   
   ' resize arrays
   ReDim Preserve m_arKDFile(m_lnumFiles)
   ReDim Preserve m_arKDFilePath(m_lnumFiles)
   ' job done, exit stage...left ;)
   Exit Do
  End If
  
  ' increment counter
  i = i + 1
 Loop

 If (bFound) Then
  AppendToLog (strFileName & " was removed sucessfully.")
 Else
  AppendToLog ("Could not find " & strFileName & " in archive!")
 End If
End Function

'////////////////////////////////////////////////////////////////
'//// Add file to packet
'//// STRING strFileName - path&filename to add to packet
'//// STRING strInfo     - filename info
'//// Returns: index in the kdf array
'////////////////////////////////////////////////////////////////
Public Function _
AddFile(strFileName As String, _
        Optional strInfo As String = "") As Long
 
 ' incremenet files counter
 m_lnumFiles = m_lnumFiles + 1
 ' resize array
 ReDim Preserve m_arKDFile(m_lnumFiles)
 ReDim Preserve m_arKDFilePath(m_lnumFiles)
 
 ' fill info
 Dim llbl     As Long                            ' last backlash located@
 Dim strFile  As String
 Dim cn       As Integer
 
 AppendToLog ("Adding " & strFileName)
 
 ' extract file name from file path
 Do While (cn < Len(strFileName))
  
  cn = llbl + 1
  llbl = InStr(cn, strFileName, "\")
  
  ' no more backlashes located
  If (llbl = 0) Then
   strFile = Mid$(strFileName, cn)
   Exit Do
  End If
 
 Loop
 
 m_arKDFile(m_lnumFiles).strName = strFile
 m_arKDFile(m_lnumFiles).strInfo = strInfo
 m_arKDFilePath(m_lnumFiles) = strFileName
 
 ' return
 AddFile = m_lnumFiles
 AppendToLog (strFileName & " was added successfully.")
 
End Function

' *_*_-_*----- proprties -----*_-_*_*

'////////////////////////////////////////////////////////////////
'//// Retrieve loaded packet name
'////////////////////////////////////////////////////////////////
Public Property Get _
GetPacketName() As String

 GetPacketName = m_strPacketName

End Property

'////////////////////////////////////////////////////////////////
'//// Get entry position from it's name
'//// STRING strFileName - valid file name
'////////////////////////////////////////////////////////////////
Public Property Get _
GetEntryPositionFromName(strFileName As String) As Long
 
 Dim cn     As Integer
 Dim bFound As Boolean
 
 ' check for opened tag
 If (Not m_bTagOpened Or _
     m_lnumFiles < 0) Then Exit Property
 
 For cn = 0 To m_lnumFiles
  
  'Stop
  
  If (m_arKDFile(cn).strName = strFileName) Then
   bFound = True
   GetEntryPositionFromName = m_arKDFile(cn).lPosition
   Exit For
  End If
 
 Next cn
 
 ' check for file found
 If (Not bFound) Then GetEntryPositionFromName = -1
 
End Property

'////////////////////////////////////////////////////////////////
'//// Get entry position from it's index
'//// LONG lNum - valid index
'////////////////////////////////////////////////////////////////
Public Property Get _
GetEntryPositionFromIndex(lNum As Long) As Long
 
 If (lNum < 0 Or lNum > m_lnumFiles) Then
  GetEntryPositionFromIndex = -1
  Exit Property
 End If
 
 GetEntryPositionFromIndex = m_arKDFile(lNum).lPosition

End Property

'////////////////////////////////////////////////////////////////
'//// Get entry info from it's index
'//// LONG lNum - valid index
'////////////////////////////////////////////////////////////////
Public Property Get _
GetEntryInfo(lNum As Long) As String
 
 If (lNum < 0 Or lNum > m_lnumFiles) Then
  GetEntryInfo = ""
  Exit Property
 End If
 
 GetEntryInfo = m_arKDFile(lNum).strInfo

End Property

'////////////////////////////////////////////////////////////////
'//// Get entry name from it's index
'//// LONG lNum - valid index
'////////////////////////////////////////////////////////////////
Public Property Get _
GetEntryName(lNum As Long) As String
 
 If (lNum < 0 Or lNum > m_lnumFiles) Then
  GetEntryName = ""
  Exit Property
 End If
 
 GetEntryName = m_arKDFile(lNum).strName

End Property

'////////////////////////////////////////////////////////////////
'//// Get class status/log
'////////////////////////////////////////////////////////////////
Public Property Get _
GetLog() As String

 GetLog = m_strLog

End Property

'////////////////////////////////////////////////////////////////
'//// Get version
'////////////////////////////////////////////////////////////////
Public Property Get _
GetVersion() As String

 Dim strTemp As String
 
 strTemp = CStr(KDF_VERSION)
 strTemp = Left$(strTemp, 1) & "." & Mid$(strTemp, 2)
 GetVersion = strTemp
  
End Property

