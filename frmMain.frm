VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picl 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   2400
      Picture         =   "frmMain.frx":0ECA
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   9600
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BETA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   555
         Left            =   1800
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Label lblStatus 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   585
   End
   Begin VB.Shape shpload 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   720
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpoutline 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   720
      Top             =   2520
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private lastload As Long

Public Sub _
UpdatePBar(indicator As Integer)
  
  Dim pbar As Integer
  
  ' update progress bar
  lastload = lastload + indicator
  pbar = lastload * Screen.TwipsPerPixelX
  If (pbar > shpoutline.Width) Then pbar = shpoutline.Width
  
  shpload.Width = pbar

End Sub


Private Sub Form_Activate()
 
 lastload = 0
 Dim fwidth As Long
 Dim fheight As Long
 
 fwidth = MAX_CX * Screen.TwipsPerPixelX
 fheight = MAX_CY * Screen.TwipsPerPixelY
 
 ' setup loading stuff and splash
 With Me
  .picl.Left = fwidth / 2 - .picl.ScaleWidth / 2
  .picl.Top = (fheight / 4 - .picl.ScaleHeight / 2)
  .shpoutline.Width = 500 * Screen.TwipsPerPixelX
  .shpoutline.Left = fwidth / 2 - .shpoutline.Width / 2
  .shpoutline.Top = fwidth - fwidth / 2
  .shpload.Left = .shpoutline.Left
  .shpload.Top = .shpoutline.Top
  .shpload.Width = 0
  .lblStatus.Left = .shpoutline.Left
  .lblStatus.Top = .shpoutline.Top + .shpoutline.Height + 10 * Screen.TwipsPerPixelY
 End With
 DoEvents
 
 ' hide beta stuff
 Label1.Visible = False
 
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call ReleaseAll                                 ' release game stuff on exit
End Sub

