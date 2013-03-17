VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4800
   Icon            =   "DD2Setup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DD2Setup.frx":0ECA
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVsync 
      BackColor       =   &H00000000&
      Caption         =   "VSync  (recommended)"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      MaskColor       =   &H80000005&
      TabIndex        =   11
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   795
      Left            =   1800
      Picture         =   "DD2Setup.frx":2DFD
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   10
      Top             =   4200
      Width           =   1110
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   795
      Left            =   240
      Picture         =   "DD2Setup.frx":3563
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   9
      Top             =   4200
      Width           =   1110
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H00FFFF00&
      Caption         =   "Run Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   9
      Left            =   1560
      Top             =   2280
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Choose Language"
      ForeColor       =   &H00FFFF00&
      Height          =   2775
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      Begin VB.PictureBox picENG 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   480
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   70
         TabIndex        =   6
         ToolTipText     =   "English texts"
         Top             =   1800
         Width           =   1080
      End
      Begin VB.PictureBox picBG 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   480
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   70
         TabIndex        =   5
         ToolTipText     =   "Bulgarian texts"
         Top             =   480
         Width           =   1080
      End
      Begin VB.OptionButton optENG 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton optBG 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Display Adapter"
      ForeColor       =   &H00FFFF00&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox cbAdapter 
         Height          =   315
         ItemData        =   "DD2Setup.frx":44B3
         Left            =   120
         List            =   "DD2Setup.frx":44B5
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://kenamick.hit.bg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   2760
      MousePointer    =   10  'Up Arrow
      TabIndex        =   12
      Top             =   3000
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Type ADAPTER_
  strDesc As String
  strGuid As String
End Type

Dim m_dx          As DirectX7
Dim arAdapter()   As ADAPTER_
Dim nAdapterCount As Long


Private Sub cmdExit_Click()
 Call SavePrefs
 End
End Sub

' run the game
Private Sub cmdRun_Click()
 Call SavePrefs
 Call ShellExecute(0&, "open", "dd2.exe", vbNullString, App.Path, 1)
 End
End Sub

Private Sub Form_Load()
 On Local Error GoTo DXERROR
  
 ' create directX object
 Set m_dx = New DirectX7
 
 Me.Caption = "DD2 Battle For Existance - Setup"
 'optENG.Value = True
 Label1.Caption = "http://www.kenamick.com"
 
 Call GetDisplayModes
 Call SinusFlag.CalcSinTable
  SinusFlag.WaveFlag pic1, picBG
 SinusFlag.WaveFlag pic2, picENG
   Call OpenPrefs

Exit Sub

DXERROR:
 MsgBox "init() error or DirectX7 not found!", vbExclamation
 End
End Sub

Private Sub _
GetDisplayModes()
 
 On Local Error Resume Next
 
 Dim ddEnum As DirectDrawEnum
 'Dim strGuid As String
 Dim i      As Long
 
 Set ddEnum = m_dx.GetDDEnum()
  
 ' resize array
 nAdapterCount = ddEnum.GetCount()
 ReDim arAdapter(nAdapterCount - 1)
 
 ' enumerate
 For i = 1 To ddEnum.GetCount()
   arAdapter(i - 1).strGuid = ddEnum.GetGuid(i)
   arAdapter(i - 1).strDesc = ddEnum.GetDescription(i)
   ' add to combo
   cbAdapter.AddItem arAdapter(i - 1).strDesc
   
 Next
 
 ' set first device selected
 cbAdapter.ListIndex = 0
 
End Sub

Private Sub _
OpenPrefs()

 Dim ff      As Integer
 Dim buffer  As String
 Dim lbuffer As Long
  
  On Local Error Resume Next
    
  ff = FreeFile()
    
  Open (App.Path & "\pref") For Binary Access Read Lock Write As #ff
    
    Get #ff, , lbuffer
    If (lbuffer) Then
     optBG.Value = True
     optENG.Value = False
    Else
     optBG.Value = False
     optENG.Value = True
    End If
        
    Get #ff, , lbuffer
    If (lbuffer) Then
     chkVsync.Value = vbChecked
    Else
     chkVsync.Value = vbUnchecked
    End If
    
    Get #ff, , lbuffer
    buffer = Space$(lbuffer)
    Get #ff, , buffer
    
  Close #ff
End Sub


Private Sub _
SavePrefs()

 Dim ff As Integer
 Dim buffer  As String
 Dim lbuffer As Long
  
 On Local Error Resume Next
    
  Kill App.Path & "\pref"
    
  ff = FreeFile()
    
  Open (App.Path & "\pref") For Binary Access Write Lock Read As #ff
    

    If (optBG.Value = True And optENG.Value = False) Then
     lbuffer = 1
    Else
     lbuffer = 0
    End If
    Put #ff, , lbuffer
    
    If (chkVsync.Value = vbChecked) Then
     lbuffer = 1
    Else
     lbuffer = 0
    End If
    Put #ff, , lbuffer

    lbuffer = Len(arAdapter(cbAdapter.ListIndex).strGuid)
    Put #ff, , lbuffer
    Put #ff, , arAdapter(cbAdapter.ListIndex).strGuid

  Close #ff
End Sub

' web
Private Sub Label1_Click()
 On Local Error Resume Next
 
 Call ShellExecute(0&, vbNullString, Trim$(Label1.Caption), vbNullString, "C:\", 1)
End Sub

Private Sub picBG_Click()
 optBG.Value = True
End Sub

Private Sub picENG_Click()
 optENG.Value = True
End Sub

' waveflags
Private Sub Timer1_Timer()
 
 SinusFlag.WaveFlag pic1, picBG
 SinusFlag.WaveFlag pic2, picENG

End Sub

