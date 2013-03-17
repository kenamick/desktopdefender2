Attribute VB_Name = "mFModUtil"
'-------------------------------------
'--> DirectSound Engine
'--> by Ilian "Firesword"
'--> KenamicK Entertainment 1998-2002
'-------------------------------------

'MODULE FE (FMOD EASY)
Public sarray(2, 32) As Long '1- stream , 2 - channel
Public channel2 As Long
Public stream2 As Long


'**********************************
'Some SUBS for INIT & CLOSE FMOD!!!
'----------=IMOPRTANT=-------------
'Some SUBS for INIT & CLOSE FMOD!!!
'**********************************

Public Sub initFMOD()
If FSOUND_Init(44100, 32, 0) = 0 Then
    'Error
    MsgBox "An error occured initializing fmod" & vbCrLf & FSOUND_GetErrorString(FSOUND_GetError)
    End
End If
End Sub

Public Sub closeFMOD()
'Stop the song that was playing
For i = 1 To 32
If sarray(2, i) <> 0 Then
    FSOUND_Stream_Stop sarray(1, i)
    sarray(2, i) = 0
End If
Next i
'Close any opened file
For i = 1 To 32
If sarray(1, i) <> 0 Then
    FSOUND_Stream_Close sarray(1, i)
    sarray(1, i) = 0
End If
Next i

If stream2 <> 0 Then
    FMUSIC_StopSong stream2
    FMUSIC_FreeSong stream2
    stream2 = 0
End If
'Make sure you close FMOD on exiting
'(If you forget this, Visual Basic will crash upon exiting the app in debug mode)
FSOUND_Close
End Sub


'*****************************************************
'SUBs FOR MP3 (Also works with other formats even MID)
'SUBs FOR MP3 (Also works with other formats even MID)
'SUBs FOR MP3 (Also works with other formats even MID)
'*****************************************************
'you don't have to use fmusic_ for MIDs
'all are                       ****file

Public Sub OPENfile(file As String, VariableBitRate As Boolean, channel As Byte)
'ako kanala raboti da spre i da se osvobodi
If sarray(2, channel) <> 0 Then
    FSOUND_Stream_Stop sarray(1, channel)
    sarray(2, channel) = 0
End If
'stream-a da se zatvori vednagi4eski
If sarray(1, channel) <> 0 Then
    FSOUND_Stream_Close sarray(1, channel)
    sarray(1, channel) = 0
End If
'
If VariableBitRate Then
        sarray(1, channel) = FSOUND_Stream_OpenFile(file, FSOUND_MPEGACCURATE Or FSOUND_STEREO Or FSOUND_2D Or FSOUND_16BITS Or FSOUND_LOOP_NORMAL, 0)
    Else
        sarray(1, channel) = FSOUND_Stream_OpenFile(file, FSOUND_STEREO Or FSOUND_2D Or FSOUND_16BITS Or FSOUND_LOOP_NORMAL, 0)
End If

End Sub

Public Sub PLAYfile(channel As Byte)
If sarray(2, channel) <> 0 Then
    FMod.FSOUND_Stream_Stop sarray(1, channel)
    sarray(2, channel) = 0
End If
sarray(2, channel) = FSOUND_Stream_Play(FSOUND_FREE, sarray(1, channel))
'MsgBox (FMod.FSOUND_GetLoopMode(sarray(2, channel)))
'bloop = FSOUND_SetLoopMode(sarray(2, channel), bloop)
End Sub

Public Sub STOPfile(channel As Byte)
'stop stream
    FSOUND_Stream_Stop sarray(1, channel)
    sarray(2, channel) = 0 '- sled stop kanala se osvobojdava i za tova e redno promenlivata da e 0
End Sub

Public Sub PAUSEfile(channel As Byte)
'no error if no stream loaded or playing !! !!!!
FSOUND_SetPaused sarray(2, channel), Not CBool(FSOUND_GetPaused(sarray(2, channel)))
End Sub

Public Sub CLOSEfile(channel As Byte)
'VERY interesting!!! no error if a stream is playing!!! !!!!

    FSOUND_Stream_Close sarray(1, channel)
    sarray(1, channel) = 0
    sarray(2, channel) = 0

End Sub
 
Public Sub volume(channel As Byte, volume As Byte)
FMod.FSOUND_SetVolume sarray(2, channel), volume
End Sub

'********************************
'SUBs FOR  -=MOD,S3M,XM,IT,MIDI=-
'********************************
'using FMUSIC_
'all are             ****mus


Public Sub OPENmus(file As String)
If stream2 <> 0 Then FMUSIC_StopSong stream2
stream2 = 0
stream2 = FMUSIC_LoadSong(file)
End Sub

Public Sub PLAYmus(bloop As Boolean)
'if a song is playing it restarts automatically
'here it returns an error for sure
If stream2 <> 0 Then
FMod.FMUSIC_PlaySong (stream2)
FMod.FMUSIC_SetLooping stream2, bloop
End If
End Sub

Public Sub PAUSEmus()
'it doesn't make an error if no music is loaded
FMod.FMUSIC_SetPaused stream2, Not CBool(FMod.FMUSIC_GetPaused(stream2))
End Sub

Public Sub STOPmus()
'stops a song, doesn't make an error if the music is already stopped
FMod.FMUSIC_StopSong stream2
End Sub

Public Sub CLOSEmus()
'strange function doesn't close the music ???!!!!
FMod.FMUSIC_FreeSong stream2
stream2 = 0
End Sub
