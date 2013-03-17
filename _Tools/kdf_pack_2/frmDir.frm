VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Directory"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Folder to Add"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1440
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strDir As String

Private Sub btnCancel_Click()
 strDir = ""
 
 Me.Hide
End Sub

Private Sub Form_Activate()
 If (strDir <> "") Then Dir1.Path = strDir
End Sub

Private Sub btnOk_Click()

 strDir = Dir1.Path
 
 Me.Hide
End Sub

Private Sub Drive1_Change()
 ' select new drive
 On Local Error GoTo DRIVERROR
 
 Dir1.Path = Drive1.Drive

Exit Sub

DRIVERROR:
 Call MsgBox("Drive not responding, dude!", vbExclamation, "Whoopsyyy!")
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call btnCancel_Click
End Sub
