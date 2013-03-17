VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFile 
      Height          =   5130
      ItemData        =   "frmMain.frx":030A
      Left            =   3840
      List            =   "frmMain.frx":030C
      TabIndex        =   8
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status Box"
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   3615
      Begin VB.TextBox txtLog 
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      X1              =   7200
      X2              =   7200
      Y1              =   5400
      Y2              =   6600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      X1              =   3720
      X2              =   7200
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label lblContainer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lblIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4560
      TabIndex        =   10
      Top             =   5640
      Width           =   585
   End
   Begin VB.Label lblIndex1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3840
      TabIndex        =   9
      Top             =   5640
      Width           =   630
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      X1              =   0
      X2              =   3720
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   3720
      X2              =   7200
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   3720
      X2              =   3720
      Y1              =   120
      Y2              =   6600
   End
   Begin VB.Label lblPosition 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4800
      TabIndex        =   7
      Top             =   6120
      Width           =   585
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4440
      TabIndex        =   6
      Top             =   5880
      Width           =   585
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   5400
      Width           =   585
   End
   Begin VB.Label lblPosition1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3840
      TabIndex        =   4
      Top             =   6120
      Width           =   915
   End
   Begin VB.Label lblInfo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3840
      TabIndex        =   3
      Top             =   5880
      Width           =   390
   End
   Begin VB.Label lblName1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3840
      TabIndex        =   2
      Top             =   5400
      Width           =   690
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOpenTag 
         Caption         =   "Open &Tag File"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open &Data File"
      End
      Begin VB.Menu mnu_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnu_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add File"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAddDir 
         Caption         =   "Add &Dir"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove File"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "&Extarct File"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuExtractAll 
         Caption         =   "E&xtract All"
      End
      Begin VB.Menu mnu_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddNfo 
         Caption         =   "Add &Info on Files"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpR 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnu_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' vars
Private cKDF2      As clsKDF2
Private bTagOpened As Boolean
Private bNew       As Boolean
Private bInfoOn    As Boolean
Private strTagName As String
Private strKDFName As String

' functions

Private Sub Form_Load()
  
  ' new KDF
  bNew = True
  strKDFName = ""
  strTagName = ""
  lstFile.Clear
  
  ' on startup
  Set cKDF2 = Nothing
  Set cKDF2 = New clsKDF2
  
  frmMain.Caption = "KenamicK Data File Packeger " & cKDF2.GetVersion
  bTagOpened = False

  ' setup menus
  mnuOpen.Enabled = False
  mnuAdd.Enabled = True
  mnuRemove.Enabled = True
  mnuExtract.Enabled = False
  mnuExtractAll.Enabled = False
  
  ' setup lables
  lblName = ""
  lblInfo = ""
  lblPosition = ""
  lblIndex = ""
  lblContainer = ""
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 ' on end
 Set cKDF2 = Nothing
 End
 
End Sub


Private Sub _
AddDir(strPath As String)
 
 ' desc: Add directory files to packet

  Dim strTemp  As String
  'Dim strPath  As String
  
  ' add lash
  If (Right$(strPath, 1) <> "\") Then strPath = strPath & "\"
  
  strTemp = Dir(strPath)
  
  ' save all files in dir
  Do While (strTemp <> "")
   Call cKDF2.AddFile(strPath & strTemp)
   strTemp = Dir()
  Loop
    
  'cKDF2.SavePacket ("d:\projects\dd2\gfx\myk1.kdf")
  'cKDF2.SaveTag ("d:\projects\dd2\gfx\tag1.ktf")
     
  Call RefreshList
  Call RefStatus
  
End Sub

Private Sub _
RefreshList()
 ' desc: refresh file list
 
 Call cKDF2.ListAll(lstFile)

End Sub

Private Sub _
RefStatus()
 ' desc: refrseh status box
 On Local Error Resume Next
 
 txtLog.Text = ""
 txtLog.Text = cKDF2.GetLog

End Sub

Private Sub lstFile_Click()
 ' get file info
  
 lblName.Caption = cKDF2.GetEntryName(lstFile.ListIndex)
 lblInfo.Caption = cKDF2.GetEntryInfo(lstFile.ListIndex)
 lblPosition.Caption = cKDF2.GetEntryPositionFromIndex(lstFile.ListIndex)
 lblIndex.Caption = lstFile.ListIndex

End Sub

Private Sub mnuAddNfo_Click()
 ' desc: add info option
 If (bInfoOn) Then
   bInfoOn = False
   mnuAddNfo.Checked = False
 Else
   bInfoOn = True
   mnuAddNfo.Checked = True
 End If
End Sub

' --- Menus ---

Private Sub mnuExit_Click()
 ' exit program
 Set cKDF2 = Nothing
 End
End Sub

Private Sub mnuExtract_Click()
 ' desc: extract file
 
 Call MsgBox("Command not supported!")
End Sub

Private Sub mnuExtractAll_Click()
  ' desc: extract all files
 
 Call MsgBox("Command not supported!")
End Sub


Private Sub mnuRemove_Click()
 ' desc: remove file from packet
 
 If (bNew) Then
  
  Call cKDF2.DeleteFile(lstFile.List(lstFile.ListIndex))
  ' refresh list
  Call RefreshList
  Call RefStatus
 End If
 
End Sub

Private Sub mnuAddDir_Click()
 ' desc: add directory
   
 ' show folder dialog
 frmDir.Show vbModal, Me
 
 ' add the whole dir
 If (frmDir.strDir <> "") Then
  Call AddDir(frmDir.strDir)
 End If
 
End Sub


Private Sub mnuAdd_Click()
 ' desc: add file to packet
 
 Dim strFile As String
 Dim strInf  As String
 
 ' only if this is a new packet
 If (bNew) Then
  
  Call OpenFile(strFile, "*.*", , True, "Select File to Add")
  If (strFile = "none") Then Exit Sub
  ' check for info on
  If (bInfoOn) Then
   strInf = InputBox("Enter this file Extended Information:", "Add Info")
  End If
  ' add the file
  Call cKDF2.AddFile(strFile, strInf)
  ' refresh list
  Call RefreshList
 Else
  '...
 End If
 
End Sub


Private Sub mnuNew_Click()
 ' desc: reset class
 Call Form_Load
  
End Sub

Private Sub mnuOpen_Click()
 ' desc: open saved pack
 
 Dim strPackName As String
 
 Call OpenFile(strPackName, "*.KDF", , True)
 
 ' exit if nothing selected
 If (strPackName = "none") Then
  GoTo LOCERROR
 ' attempt to open packet
 Else
  If (Not cKDF2.LoadPacket(strPackName)) Then GoTo LOCERROR
 End If
  
Exit Sub

LOCERROR:
 Call RefStatus
End Sub


Private Sub mnuOpenTag_Click()
 ' desc: open tag
  
 Dim strPackName As String
 
 Call OpenFile(strPackName, "*.ktf", , True)
 
 ' exit if nothing selected
 If (strPackName = "none") Then
  GoTo LOCERROR
 ' attempt to open packet
 Else
  If (Not cKDF2.LoadTag(strPackName)) Then GoTo LOCERROR
  mnuOpen.Enabled = True
  'list files
  Call cKDF2.ListAll(lstFile)
  ' disable new & adds
  bNew = False
  mnuAdd.Enabled = False
  mnuRemove.Enabled = False
  mnuExtract.Enabled = True
  mnuExtractAll.Enabled = True
 End If
  
Exit Sub

LOCERROR:
 Call RefStatus
End Sub


Private Sub mnuSave_Click()
 ' desc: save kdf

 ' check file names
 If (strKDFName = "" Or strTagName = "") Then
  Call mnuSaveAs_Click
 Else
  ' save data & tag
  Call cKDF2.SavePacket(strKDFName)
  Call cKDF2.SaveTag(strTagName)
 End If
 
  ' disable new & adds
  bNew = False
  mnuAdd.Enabled = False
  mnuRemove.Enabled = False
  mnuExtract.Enabled = True
  mnuExtractAll.Enabled = True
 
 Call RefStatus
 
End Sub

Private Sub mnuSaveAs_Click()
 ' desc: save as option
 
 Dim strPackName As String
 
 Call OpenFile(strPackName, "*.ktf", False, , "Enter Tag Name")
 ' exit if nothing selected
 If (strPackName = "none") Then
  GoTo LOCERROR
 Else
  strTagName = strPackName
 End If
 
 Call OpenFile(strPackName, "*.kdf", , False, "Enter Data File Name")
 ' exit if nothing selected
 If (strPackName = "none") Then
  GoTo LOCERROR
 Else
  strKDFName = strPackName
 End If
  
 Call mnuSave_Click
  
Exit Sub

LOCERROR:
 Call RefStatus
End Sub

Private Sub Timer1_Timer()
 ' desc: timer to refrsh status
 Call RefStatus
End Sub

Private Sub mnuHelpR_Click()
 ' desc: help, heh
 
 Dim strMsg As String
 
 strMsg = "Really maaan, it's simple!" & vbCrLf & _
          "Don't make me explain, my eyeballs are to fall" & vbCrLf & _
          "any moment now....heeeelp!" & vbCrLf & vbCrLf & _
          "Who needs help now...."

 Call MsgBox(strMsg, vbCritical)
End Sub

Private Sub mnuAbout_Click()
 ' desc: uhhaha
 
 Dim strMsg As String

 strMsg = "KenamicK Data File Packeger®" & vbCrLf & _
          "----------------------------------" & vbCrLf & _
          "by Pro-XeX" & vbCrLf & vbCrLf & _
          "made in a hurry for our game Desktop Defender II" & vbCrLf & _
          "The old KDF didn't work right, so I had to write this " & vbCrLf & _
          "new one....from scratch (yikes)!" & vbCrLf & vbCrLf & _
          "e-mail: bgPro_XeX@yahoo.com" & vbCrLf & _
          "http://kenamick.hit.bg" & vbCrLf
 
 MsgBox strMsg, vbInformation

End Sub

'Private Sub Picture1_Click()
'
' Static lc As Integer
'
' LRLoadBitmap Picture1.hdc, "D:\Projects\DD2\Tools\mygfx.kdf", cKDF2.GetEntryPositionFromName("Mete2" & lc & ".bmp")
' lc = lc + 1
'
'End Sub
