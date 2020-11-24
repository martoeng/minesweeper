VERSION 5.00
Begin VB.Form frmBestList 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Die schnellsten MineSweeper"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBestList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtName 
      Alignment       =   1  'Rechts
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Text            =   "Anonym"
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "(Kein Eintrag)"
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   2500
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "(Kein Eintrag)"
      Height          =   210
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   480
      Width           =   2500
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "(Kein Eintrag)"
      Height          =   210
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   2500
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   210
      Index           =   3
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   315
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   210
      Index           =   2
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   315
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   210
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lblProfis 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profis:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   585
   End
   Begin VB.Label lblFortgeschrittene 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fortgeschrittene:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label lblAnfänger 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anfänger:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmBestList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    'Prüfen, ob neue Bestzeit
    If txtName.Visible Then
        Select Case txtName.Top
            Case lblName(1).Top 'Anfänger
                SaveSetting "Martoeng\MineSweeper", "Toplist", "StarterName", txtName.Text
                SaveSetting "Martoeng\MineSweeper", "Toplist", "StarterTime", lblPoints(1).Caption
            Case lblName(2).Top 'Fortgeschrittene
                SaveSetting "Martoeng\MineSweeper", "Toplist", "AdvancedName", txtName.Text
                SaveSetting "Martoeng\MineSweeper", "Toplist", "AdvancedTime", lblPoints(2).Caption
            Case lblName(3).Top 'Profis
                SaveSetting "Martoeng\MineSweeper", "Toplist", "ProfiName", txtName.Text
                SaveSetting "Martoeng\MineSweeper", "Toplist", "ProfiTime", lblPoints(3).Caption
        End Select
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    lblPoints(1).Caption = GetSetting("Martoeng\MineSweeper", "Toplist", "StarterTime", lblPoints(1).Caption)
    lblPoints(2).Caption = GetSetting("Martoeng\MineSweeper", "Toplist", "AdvancedTime", lblPoints(2).Caption)
    lblPoints(3).Caption = GetSetting("Martoeng\MineSweeper", "Toplist", "ProfiTime", lblPoints(3).Caption)
    lblName(1).Caption = GetSetting("Martoeng\MineSweeper", "Toplist", "StarterName", "(Kein Eintrag)")
    lblName(2).Caption = GetSetting("Martoeng\MineSweeper", "Toplist", "AdvancedName", "(Kein Eintrag)")
    lblName(3).Caption = GetSetting("Martoeng\MineSweeper", "Toplist", "ProfiName", "(Kein Eintrag)")
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0: txtName.SelLength = Len(txtName.Text)
End Sub
