Attribute VB_Name = "mDirectSound"
Option Explicit
'-------------------------------------
'--> DirectSound Engine
'--> by Peter "Pro-XeX" Petrov
'--> KenamicK Entertainment 1998-2002
'-------------------------------------

'File format structure for wave files
Private Type WAVEFORMATEX1
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
End Type

'File header structure for wave files.
Private Type WAVEFILEHEADER
    dwRiff As Long
    dwFileSize As Long
    dwWave As Long
    dwFormat As Long
    dwFormatLength As Long
End Type

Public Type stSoundBuffer
 m_lpDSBuffer() As DirectSoundBuffer              ' the sound buffer(s)
 nNumBuffers  As Integer                        ' number of buffers
 nCurrent     As Integer                        ' current/last buffer playing
 nVolume      As Integer                        ' current volume
End Type

Enum cnstLOADSOURCE
 LS_FROMFILE = 0
 LS_FROMRESOURCE
 LS_FROMBINRES
End Enum

Private Const RIFF = &H46464952                 ' RIFF header
Private Const WAVE = &H45564157                 ' WAVE header

Private m_lpDS           As DirectSound         ' DirectSound object
Private m_bDSInitialized As Boolean             ' is DirectSound intialized
Public m_bDSOn           As Boolean             ' sound on/off
Public m_nGlobalVol      As Integer             ' global volume ( 0 - MAX )

'//////////////////////////////////////////////////////////////////
'//// Initialize DirectSound
'//// hWnd - handle to the game canvas
'//////////////////////////////////////////////////////////////////
Public Sub _
DSInit(hwnd As Long)

 On Local Error GoTo DSERROR
 
 Call AppendToLog(LOG_DASH)
 Call AppendToLog("Initializing DirectSound...")
 ' create object
 Set m_lpDS = lpDX.DirectSoundCreate("")
 ' set coop. level
 AppendToLog ("Setting cooperativelevel...")
 Call m_lpDS.SetCooperativeLevel(hwnd, DSSCL_PRIORITY)
 Call AppendToLog("DirectSound was initialized successfully.")
 
 ' set DS initialized flag
 m_bDSInitialized = True
 
 ' enable play by default
 m_bDSOn = True

 ' set default volume {!}
 m_nGlobalVol = -2003

Exit Sub

DSERROR:
 Call AppendToLog(DSGetErrorDesc(Err.Number))
 MsgBox DSGetErrorDesc(Err.Number), vbExclamation
End Sub

'//////////////////////////////////////////////////////////////////
'//// Create Sound Buffer From file
'//// strFileName - soundfile name and location
'//// m_lpDSBuffer - pointer to the DS buffer
'//////////////////////////////////////////////////////////////////
Public Sub _
DSLoadSoundFromFile(strFileName As String, m_lpDSBuffer As DirectSoundBuffer)
 
 ' exit if no initialization
 If (Not m_bDSInitialized) Then Exit Sub
 
 On Local Error GoTo DSERRLSFF
 
 Dim bd As DSBUFFERDESC
 Dim wf As WAVEFORMATEX

 ' set flags - STATIC means to be loaded into hardware mem.
 bd.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or _
             DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC

 wf.nSize = Len(wf)
 wf.nFormatTag = WAVE_FORMAT_PCM
 wf.lSamplesPerSec = 22050
 wf.nChannels = 2
 wf.nBitsPerSample = 16
 wf.nBlockAlign = wf.nBitsPerSample / 8 * wf.nChannels
 wf.lAvgBytesPerSec = wf.lSamplesPerSec * wf.nBlockAlign
 
 ' initialize DS buffer
 Set m_lpDSBuffer = m_lpDS.CreateSoundBufferFromFile(strFileName, bd, wf)
 
Exit Sub

DSERRLSFF:
 Call AppendToLog(DSGetErrorDesc(Err.Number))
 MsgBox DSGetErrorDesc(Err.Number), vbExclamation
 Call ReleaseAll
End Sub

' /////////////////////////////////////////////////////////////////
' //// Loads a WAVE file from a given position in a binary data
' //// packet
' //// lOffset - the position in the binary data packet
' //// m_lpDSBuffer - is the buffer where the data will be copied to
' //// (Note: The buffer must NOT have been already created )
' /////////////////////////////////////////////////////////////////
Public Function _
DSLoadSoundFromBin(lpszFileName As String, lOffset As Long, _
                   ByRef m_lpDSBuffer As DirectSoundBuffer) As Boolean

 ' exit if no initialization
 If (Not m_bDSInitialized) Then Exit Function
 
 On Local Error GoTo DSERRLSB
 
 Dim cn As Long                               ' local counter
 Dim nFF As Integer                           ' file handle
 Dim byTemp As Byte                           ' one temp byte
 Dim lchunkSize As Long                       ' size of wave chunk
 Dim szTemp As String * 4                     ' where the "data" string will go
 Dim bFound As Boolean                        ' was the "data" string found
 Dim lpbyData() As Byte                       ' pure wave data
 ' wave vars
 Dim wavHeader As WAVEFILEHEADER              ' waveFile header
 Dim wavFormat As WAVEFORMATEX1               ' waveFormat header
 Dim uFmt As WAVEFORMATEX                     ' DirectSound format
 Dim bd As DSBUFFERDESC
 Dim dsb As DirectSoundBuffer
 
 nFF = FreeFile()
 
 ' --- Open File ---
 Open lpszFileName For Binary Access Read Lock Write As nFF
  
  ' get header
  Get nFF, lOffset, wavHeader
  ' check for RIFF_WAVEfmt
  If (wavHeader.dwRiff <> RIFF Or _
      wavHeader.dwWave <> WAVE) Then
   DSLoadSoundFromBin = False
   Exit Function
  End If
  
  ' check for header size (only > 16 are supported)
  If (wavHeader.dwFormatLength < 16) Then
   DSLoadSoundFromBin = False
   Exit Function
  End If
  
  ' get wave_format_header
  Get nFF, , wavFormat
  
  ' get rid of extra bytes
  bFound = False
  For cn = Seek(nFF) To LOF(nFF)
   Get nFF, cn, szTemp
   If szTemp = "data" Then
    bFound = True
    Exit For
   End If
  Next
  ' "data" string was not found ( always expect the unexpectable ;)
  If (Not bFound) Then
   DSLoadSoundFromBin = False
   Exit Function
  End If
  
  ' get size of pure wav data ( the chunk maaan, the chunk )
  Get nFF, , lchunkSize
  
  ' resize byte-array
  ReDim lpbyData(lchunkSize)
  ' ok, let me have it
  Get nFF, , lpbyData
  
 ' --- Close File ---
 Close nFF
 
 ' fill format data
 With uFmt
  .nSize = LenB(uFmt)
  .lExtra = 0
  .nFormatTag = wavFormat.wFormatTag
  .nChannels = wavFormat.nChannels
  .nBitsPerSample = wavFormat.wBitsPerSample
  .lSamplesPerSec = wavFormat.nSamplesPerSec
  .nBlockAlign = wavFormat.nBlockAlign
  .lAvgBytesPerSec = wavFormat.nAvgBytesPerSec
 End With
  
 ' set size of data chunk
 bd.lBufferBytes = lchunkSize
 ' flags should be applied here
 bd.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or _
             DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC

 Set m_lpDSBuffer = m_lpDS.CreateSoundBuffer(bd, uFmt)
 ' write byte data to the buffer
 Call m_lpDSBuffer.WriteBuffer(0, lchunkSize, _
                             lpbyData(0), DSBLOCK_ENTIREBUFFER)
 DSLoadSoundFromBin = True
Exit Function

DSERRLSB:
 Call AppendToLog(DSGetErrorDesc(Err.Number))
 MsgBox DSGetErrorDesc(Err.Number), vbExclamation
 DSLoadSoundFromBin = False
End Function

'//////////////////////////////////////////////////////////////////
'//// Create a sound strucutre
'//// lpSound - pointer to the Sound structure
'//// nNumBuffers - a number of sound buffers
'//// strFileName - filename of the sound file or bin file
'//// lOffset - pos/index of the sound into the bin/res file
'//////////////////////////////////////////////////////////////////
Public Sub _
DSCreateSound(lpSound As stSoundBuffer, nNumBuffers As Integer, _
              cnstLR As cnstLOADSOURCE, _
              strFileName As String, _
              Optional lOffset As Long = 0)

 ' exit if no initialization
 If (Not m_bDSInitialized) Then Exit Sub
 
 On Error GoTo 0
 
 Dim cn As Long
 
 ' set buffers
 If (nNumBuffers < 1) Then nNumBuffers = 1
 ' make zero floor..err, wha?
 nNumBuffers = nNumBuffers - 1
 lpSound.nNumBuffers = nNumBuffers
 lpSound.nCurrent = 0
 lpSound.nVolume = DSBVOLUME_MAX
 
 ' resize dsbuffer array
 ReDim lpSound.m_lpDSBuffer(nNumBuffers)
  
 ' load primary sound buffer
 If (cnstLR = LS_FROMFILE) Then
  Call DSLoadSoundFromFile(strFileName, lpSound.m_lpDSBuffer(0))
 ElseIf (cnstLR = LS_FROMRESOURCE) Then
  '...
 ElseIf (cnstLR = LS_FROMBINRES) Then
  'Call DSLoadSoundFromBin(strFileName, lOffset, lpSound.m_lpDSBuffer(0))
  Call DSLoadSoundFromBin(CKdfSfx.GetPacketName(), _
                          CKdfSfx.GetEntryPositionFromName(strFileName), _
                          lpSound.m_lpDSBuffer(0))
 End If

 ' fill other buffers, if any
 If (nNumBuffers < 1) Then Exit Sub
 
 For cn = 1 To nNumBuffers
  
  ' duplicate buffer
  Set lpSound.m_lpDSBuffer(cn) = m_lpDS.DuplicateSoundBuffer(lpSound.m_lpDSBuffer(0))
  
  ' check for error and apply alternative method
  If (Err.Number <> 0) Then
   Call AppendToLog(DSGetErrorDesc(Err.Number))
   
   If (cnstLR = LS_FROMFILE) Then
    Call DSLoadSoundFromFile(strFileName, lpSound.m_lpDSBuffer(cn))
   ElseIf (cnstLR = LS_FROMRESOURCE) Then
    '...
   ElseIf (cnstLR = LS_FROMBINRES) Then
    Call DSLoadSoundFromBin(strFileName, lOffset, lpSound.m_lpDSBuffer(cn))
   End If
  
  End If '/error/
 
 Next

End Sub

'//////////////////////////////////////////////////////////////////
'//// Play a sound struct
'//// m_lpDSBuffer - pointer to the DS buffer
'//// bLoop - looping flag
'//////////////////////////////////////////////////////////////////
Public Sub _
DSPlaySound(lpSound As stSoundBuffer, bloop As Boolean, _
            Optional fPos As Long = 0, _
            Optional nVolume As Long = DSBVOLUME_MAX)
  
 On Local Error GoTo DSERRORPS
 
 ' exit if no initialization or sound if off
 If ((Not m_bDSOn) Or (Not m_bDSInitialized)) Then Exit Sub
 
 Dim cn        As Long
 Dim lPlayFlag As Long
  
 ' set play flag
 If (bloop) Then lPlayFlag = DSBPLAY_LOOPING Else lPlayFlag = DSBPLAY_DEFAULT
 
 For cn = 0 To lpSound.nNumBuffers
 
  Select Case lpSound.m_lpDSBuffer(cn).GetStatus
    
    Case DSBSTATUS_PLAYING
    '...
    
    Case DSBSTATUS_LOOPING
    '...
    
    Case DSBSTATUS_BUFFERLOST
     lpSound.m_lpDSBuffer(cn).restore
     '...
    
    ' this buffer is free
    Case 0
     'If (bLoop) Then Call lpSound.m_lpDSBuffer(cn).Play(DSBPLAY_LOOPING) Else _
     ' Call lpSound.m_lpDSBuffer(cn).Play(DSBPLAY_DEFAULT)
     'If (fPos <= 1# And fPos >= -1#) Then lpSound.m_lpDSBuffer(cn).SetPan (CLng(fPos * 10000&))
     Call lpSound.m_lpDSBuffer(cn).SetVolume(nVolume + m_nGlobalVol)
     Call lpSound.m_lpDSBuffer(cn).SetPan(fPos)
     Call lpSound.m_lpDSBuffer(cn).Play(lPlayFlag)
     lpSound.nCurrent = cn
     lpSound.nCurrent = lpSound.nCurrent + 1
     If (lpSound.nCurrent > lpSound.nNumBuffers) Then lpSound.nCurrent = 0
     Exit Sub
 
  End Select
 Next
 
 ' we haven't found a free buffer, so will stop one currently played
 ' and have it play the new sound
 If (lpSound.nCurrent > lpSound.nNumBuffers) Then lpSound.nCurrent = 0
 If (fPos <= 1# And fPos >= -1#) Then lpSound.m_lpDSBuffer(lpSound.nCurrent).SetPan (CLng(fPos * 10000&))
 Call lpSound.m_lpDSBuffer(lpSound.nCurrent).SetVolume(nVolume + m_nGlobalVol)
 Call lpSound.m_lpDSBuffer(lpSound.nCurrent).SetPan(fPos)
 Call lpSound.m_lpDSBuffer(lpSound.nCurrent).Stop
 Call lpSound.m_lpDSBuffer(lpSound.nCurrent).SetCurrentPosition(0)
 Call lpSound.m_lpDSBuffer(lpSound.nCurrent).Play(lPlayFlag)
 lpSound.nCurrent = lpSound.nCurrent + 1
 
Exit Sub

DSERRORPS:
 Call AppendToLog(DSGetErrorDesc(Err.Number))
 MsgBox DSGetErrorDesc(Err.Number), vbExclamation
End Sub

'//////////////////////////////////////////////////////////////////
'//// Stop playing sound structure buffer
'//// lpSound - pointer to the sound structure
'//// nBuffer - buffer to be stopped
'//////////////////////////////////////////////////////////////////
Public Sub _
DSStopSound(lpSound As stSoundBuffer, nBuffer As Integer)

 ' exit if no initialization or sound if off
 If ((Not m_bDSOn) Or (Not m_bDSInitialized)) Then Exit Sub
 
 Call DSStop(lpSound.m_lpDSBuffer(nBuffer))
End Sub

'//////////////////////////////////////////////////////////////////
'//// Set sound-structure volume
'//// lpSound - pointer to the sound structure
'//////////////////////////////////////////////////////////////////
Public Sub _
DSSetSoundVolume(lpSound As stSoundBuffer, lVolume As Long)
 
 ' exit if no initialization or sound if off
 If ((Not m_bDSOn) Or (Not m_bDSInitialized)) Then Exit Sub
 
 Dim cn As Integer
 
 ' loop trough all buffers and apply volume
 For cn = 0 To lpSound.nNumBuffers
  Call DSSetVolume(lpSound.m_lpDSBuffer(cn), lVolume)
 Next
 
End Sub

'//////////////////////////////////////////////////////////////////
'//// Play panned buffer
'//// m_lpDSBuffer - pointer to the DS buffer
'//// bLoop - loop the sound
'//// fPos - pan position between -1.0 and 1.0
'//////////////////////////////////////////////////////////////////
Public Sub _
DSPlayPan(m_lpDSBuffer As DirectSoundBuffer, _
          bloop As Boolean, _
          fPos As Single)
  
  On Local Error GoTo DSERRORPP
  
  ' check if exists or initialized or sound on/off flag
  If ((m_lpDSBuffer Is Nothing) Or (Not m_bDSOn) Or (Not m_bDSInitialized)) Then Exit Sub
  ' set panning
  Call m_lpDSBuffer.SetPan(CLng((fPos * 10000&)))
  ' set volume
  m_lpDSBuffer.SetVolume (m_lpDSBuffer.GetVolume + m_nGlobalVol)
  ' play sound
  If (bloop) Then Call m_lpDSBuffer.Play(DSBPLAY_LOOPING) Else _
   Call m_lpDSBuffer.Play(DSBPLAY_DEFAULT)

Exit Sub

DSERRORPP:
 Call AppendToLog(DSGetErrorDesc(Err.Number))
 MsgBox DSGetErrorDesc(Err.Number), vbExclamation
End Sub

'//////////////////////////////////////////////////////////////////
'//// Stop playing buffer
'//// m_lpDSBuffer - pointer to the DS buffer
'//////////////////////////////////////////////////////////////////
Public Sub _
DSStop(m_lpDSBuffer As DirectSoundBuffer)
 
 ' check if exists or initialized or sound on/off flag
 If ((m_lpDSBuffer Is Nothing) Or (Not m_bDSOn) Or (Not m_bDSInitialized)) Then Exit Sub
 ' stop playing
 m_lpDSBuffer.Stop
 ' reset position
 Call m_lpDSBuffer.SetCurrentPosition(0)
 
End Sub

'//////////////////////////////////////////////////////////////////
'//// Set sound-buffer volume
'//// m_lpDSBuffer - pointer to the DS buffer
'//// lVolume - volume num between -10000 and 0
'//////////////////////////////////////////////////////////////////
Public Sub _
DSSetVolume(m_lpDSBuffer As DirectSoundBuffer, lVolume As Long)

 ' check if exists or initialized or sound on/off flag
 If ((m_lpDSBuffer Is Nothing) Or (Not m_bDSOn) Or (Not m_bDSInitialized)) Then Exit Sub

 ' bound volume
 'If (lVolume < 0) Then lVolume = 0
 'If (lVolume > 10000) Then lVolume = 10000
 ' set
 Call m_lpDSBuffer.SetVolume(lVolume)
 
End Sub

'//////////////////////////////////////////////////////////////////
'//// DirectSound error descriptions
'//// lError - error number
'//////////////////////////////////////////////////////////////////
Public Function _
DSGetErrorDesc(lError As Long) As String

 Dim strMsg As String

 Select Case lError
   
   Case DSERR_ALLOCATED
    strMsg = "DSERR_ALLOCATED:The request failed because resources, such as a priority level, were already in use by another caller."
   
   Case DSERR_ALREADYINITIALIZED
    strMsg = "DSERR_ALREADYINITIALIZED:The object is already initialized."
   
   Case DSERR_BADFORMAT
    strMsg = "DSERR_BADFORMAT:The specified wave format is not supported."
   
   Case DSERR_BUFFERLOST
    strMsg = "DSERR_BUFFERLOST:The buffer memory has been lost and must be restored."
   
   Case DSERR_CONTROLUNAVAIL
    strMsg = "DSERR_CONTROLUNAVAIL:The control (volume, pan, and so forth) requested by the caller is not available."
   
   Case DSERR_GENERIC
    strMsg = "DSERR_GENERIC:An undetermined error occurred inside the DirectSound subsystem."
   
   Case DSERR_INVALIDCALL
    strMsg = "DSERR_INVALIDCALL:This function is not valid for the current state of this object."
   
   Case DSERR_INVALIDPARAM
    strMsg = "DSERR_INVALIDPARAM:An invalid parameter was passed to the returning function."
   
   'Case DSERR_NOAGGREGATION
   ' strMsg = "The object does not support aggregation."
   
   Case DSERR_NODRIVER
    strMsg = "DSERR_NODRIVER:No sound driver is available for use."
   
   Case DSERR_NOINTERFACE
    strMsg = "DSERR_NOINTERFACE:The requested COM interface is not available."
   
   Case DSERR_OTHERAPPHASPRIO
    strMsg = "DSERR_OTHERAPPHASPRIO:Another application has a higher priority level, preventing this call from succeeding."
   
   Case DSERR_OUTOFMEMORY
    strMsg = "DSERR_OUTOFMEMORY:The DirectSound subsystem could not allocate sufficient memory to complete the caller's request."
   
   Case DSERR_PRIOLEVELNEEDED
    strMsg = "DSERR_PRIOLEVELNEEDED:The caller does not have the priority level required for the function to succeed."
   
   Case DSERR_UNINITIALIZED
    strMsg = "Case DSERR_UNINITIALIZED:The DirectSound device has not been initialized."
   
   Case DSERR_UNSUPPORTED
    strMsg = "DSERR_UNSUPPORTED:The function called is not supported at this time."
   
   Case Else
    strMsg = "DSERROR: Error description not found!"
 
 End Select

 DSGetErrorDesc = strMsg
 AppendToLog (strMsg)

End Function

'//////////////////////////////////////////////////////////////////
'//// Unload DirectSound
'//////////////////////////////////////////////////////////////////
Public Sub _
DSRelease()
 
 ' exit if no initialization
 If (Not m_bDSInitialized) Then Exit Sub
 
 Call AppendToLog("Closing DirectSound")
 Set m_lpDS = Nothing

End Sub


