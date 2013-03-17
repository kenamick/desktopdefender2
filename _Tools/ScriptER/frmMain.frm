VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   7200
      Top             =   5040
   End
   Begin VB.TextBox txtInfo 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      TabIndex        =   7
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Left            =   2760
      TabIndex        =   6
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label lblAction 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5280
      TabIndex        =   5
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblStateN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   45
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State N:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cs As New cLevel

Private Sub Command1_Click()
 
 Timer1.Enabled = False
 
 lblStatus.Caption = "Loading script..."
 
 Call cs.Init
 cs.LoadScript ("d:\map1.txt")
 
 ' do info1
 Dim strTemp As String
  
 strTemp = "Author: " & cs.m_strAuthor & vbCrLf
 strTemp = strTemp & "Date: " & cs.m_strDate & vbCrLf
 strTemp = strTemp & "Map name: " & cs.m_strName & vbCrLf
 strTemp = strTemp & "Mission: " & cs.m_bytID & vbCrLf
 strTemp = strTemp & "Desc: " & cs.m_strDesc & vbCrLf
 strTemp = strTemp & "Duration: " & cs.m_lDuration & vbCrLf
 
 txtInfo.Text = strTemp
 ' Set Picture1.Picture = LoadPicture("d:\projects\dd2\gfx\environ\" & cs.m_strdescbkpic)
 
 
 ' do info2 & execute
 lblStatus.Caption = "Script Loaded..."
 Timer1.Enabled = True
 lblStatus.Caption = "Executing..."
End Sub

Private Sub Timer1_Timer()
  
  Call cs.Update
End Sub
