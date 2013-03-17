VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Editor"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   427
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSaveAs 
      Caption         =   "Save &As"
      Height          =   615
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   2640
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2640
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton btnDelChar 
      Caption         =   "&Del Char"
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FF0000&
      Height          =   3765
      ItemData        =   "frmMain.frx":0000
      Left            =   3000
      List            =   "frmMain.frx":0002
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtCInfo 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add &New Char"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton btnSaveAll 
      Caption         =   "&Save "
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   136
      X2              =   424
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   192
      X2              =   192
      Y1              =   64
      Y2              =   336
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ready."
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   5040
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   1830
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub _
CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Form_Load()
 ' init
 
' btnLoadAll.Enabled = False
 'btnDelChar.Enabled = False
 g_nChars = -1
 
End Sub

Private Sub btnNew_Click()
 ' desc: start new list
 
 Dim cn As Integer
 
 For cn = 0 To g_nChars
  Call CopyMemory(g_arChar(cn), 0, Len(g_arChar(cn)))
 Next
 
 g_nChars = -1
 Call ListChars
 
 
End Sub


Private Sub ListChars()
 ' desc: put all into listbox
 
 List1.Clear
 
 lblStatus.ForeColor = RGB(&H12, 12, 255)
 lblStatus.Caption = "Listing..."

 Dim cn As Integer
 
 For cn = 0 To mCharacters.g_nChars
  List1.AddItem g_arChar(cn).name, cn
 Next
 
 lblStatus.ForeColor = RGB(&H2, 255, 2)
 lblStatus.Caption = "Listing done..."

End Sub

Private Sub btnDelChar_Click()
 ' desc: delete character
 
 Call mCharacters.CHARDelCharacter(List1.ListIndex)
 Call ListChars
 
End Sub

Private Sub btnAdd_Click()
 ' desc: add new character
 
 On Local Error GoTo ADDERR
 
 Dim objCharTemp As stPlayer
 
 lblStatus.ForeColor = RGB(12, 12, &HFF)
 lblStatus.Caption = "Adding character..."
 
 ' get name
 objCharTemp.name = InputBox("Enter character name:", , "Champ")
 If (Len(objCharTemp.name) < 1) Then GoTo ADDERR
 objCharTemp.pass = InputBox("Enter character password:" & vbCrLf & "Leave blank for NONE...", , "newpass")
 objCharTemp.score = Val(InputBox("Enter score", , "0"))
 objCharTemp.level = Val(InputBox("Enter character level", , "1"))
 objCharTemp.mission = Val(InputBox("Enter character mission", , "1"))
 objCharTemp.kills = Val(InputBox("Enter character kills", , "0"))
 objCharTemp.success = Val(InputBox("Enter character shoot success", , "0"))
 
 lblStatus.ForeColor = &H2FF02
 lblStatus.Caption = "Character added..."
 
 ' add new character to the global array
 Call CHARNewCharacter(objCharTemp.name, objCharTemp.pass, objCharTemp.score, _
                       objCharTemp.level, objCharTemp.mission, objCharTemp.kills, _
                       objCharTemp.success)
 Call ListChars
 
Exit Sub

ADDERR:
 Call MsgBox("Wrong value added!", vbExclamation)
 
 lblStatus.ForeColor = RGB(&HFF, 12, 12)
 lblStatus.Caption = "Character not added..."
End Sub

Private Sub btnLoadAll_Click()
 ' desc: load all characters
 
 If (mCharacters.CHARLoadCharacters) Then
  lblStatus.ForeColor = RGB(&H12, 255, 12)
  lblStatus.Caption = "Load successful!"
 Else
  lblStatus.ForeColor = RGB(&HFF, 12, 12)
  lblStatus.Caption = "Error loading..."
  Exit Sub
 End If
  
 Call ListChars
  
End Sub

Private Sub btnSaveAll_Click()
 
 If (Len(g_strCharFileName) < 1) Then Call btnSaveAs_Click
 
 If (mCharacters.CHARSaveCharacters) Then
  lblStatus.ForeColor = RGB(&H12, 255, 12)
  lblStatus.Caption = "Save successful!"
 Else
  lblStatus.ForeColor = RGB(&HFF, 12, 12)
  lblStatus.Caption = "Error saveing..."
 End If

End Sub

Private Sub btnSaveAs_Click()
 ' desc: perform 'save as'
 
 g_strCharFileName = InputBox("Enter file PATH and file name: ", , CStr(App.Path) & "\")
 
 If (Len(g_strCharFileName) < 1) Then
  lblStatus.ForeColor = RGB(&HFF, 12, 12)
  lblStatus.Caption = "Invalid filename..."
  Exit Sub
 End If
 
 ' perfrom save all after filename is changed
 Call btnSaveAll_Click
    
End Sub

Private Sub Command1_Click()
' Stop
 g_nChars = -1
 
 '
 Call CHARNewCharacter("12309840923232pwer", "")
 Call CHARNewCharacter("peter", "sekta")
 Call CHARNewCharacter("ratae", "police")

End Sub

Private Sub List1_Click()
 ' show info
 Dim strTemp As String
 
 With g_arChar(List1.ListIndex)
  strTemp = "Name: " & .name & vbCrLf
  strTemp = strTemp & "Pass: " & .pass & vbCrLf
  strTemp = strTemp & "Score: " & .score & vbCrLf
  strTemp = strTemp & "Level: " & .level & vbCrLf
  strTemp = strTemp & "Mission: " & .mission & vbCrLf
  strTemp = strTemp & "Kills: " & .kills & vbCrLf
  strTemp = strTemp & "Success: " & .success & vbCrLf
  strTemp = strTemp & "Mask: " & .nReserved & vbCrLf
 End With
 
 txtCInfo.Text = strTemp
 
End Sub


Private Sub File1_DblClick()
 ' desc: load characters
  
 mCharacters.g_strCharFileName = File1.Path & "\" & File1.FileName
  
 If (mCharacters.CHARLoadCharacters) Then
  lblStatus.ForeColor = RGB(&H12, 255, 12)
  lblStatus.Caption = "Load successful!"
 Else
  lblStatus.ForeColor = RGB(&HFF, 12, 12)
  lblStatus.Caption = "Error loading..."
  Exit Sub
 End If
  
 Call ListChars
 
End Sub

Private Sub Drive1_Change()
 On Local Error Resume Next
 Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
End Sub

