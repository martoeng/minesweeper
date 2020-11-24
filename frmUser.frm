VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Benutzerdefiniertes Feld"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUser.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   104
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtMines 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Minen:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label lblWidth 
      Caption         =   "Breite:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblHeight 
      Caption         =   "Höhe:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   615
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dieser Sourcecode stammt von http://www.martoeng.de
'Alle Rechte liegen beim Autor Martin Walter.
'Für eine Veröffentlichung benötigen Sie die Zustimmung des Autors.
'Sie dürfen diesen Code für private Zwecke gebrauchen, nicht _
jedoch verkaufen oder für andere finanzielle Zwecke verwenden.

'Abbrechen
Private Sub cmdCancel_Click()
    mPlay.UserCancel = True
    Unload Me
End Sub

'Bestätigen und erstellen
Private Sub cmdOK_Click()
    If IsNumeric(txtWidth.Text) Then
        If IsNumeric(txtHeight.Text) Then
            If IsNumeric(txtMines.Text) Then
                If CInt(txtWidth.Text) < 9 Then txtWidth.Text = "9"
                If CInt(txtHeight.Text) < 9 Then txtHeight.Text = "9"
                If CInt(txtMines.Text) < 10 Then txtMines.Text = "10"
                If CInt(txtWidth.Text) > 30 Then txtWidth.Text = "30"
                If CInt(txtHeight.Text) > 24 Then txtHeight.Text = "24"
                If CInt(txtMines.Text) > (CInt(txtWidth.Text) - 1) * (CInt(txtHeight.Text) - 1) Then
                    txtMines.Text = (CInt(txtWidth.Text) - 1) * (CInt(txtHeight.Text) - 1)
                End If
                
                mPlay.InitArea CInt(txtWidth.Text), CInt(txtHeight.Text), CInt(txtMines.Text)
                mDraw.ClearFields
                mDraw.DrawArea
                mDraw.DrawMineCount
                mDraw.DrawTime
                mDraw.InitMines
                frmMain.tmrTime.Enabled = False
                mPlay.Seconds = 0
                mPlay.Clicked = 0
                mPlay.FlagCount = 0
                
                'Werte in die Registrierung schreiben
                SaveSetting "Martoeng\MineSweeper", "Options", "UserX", txtWidth.Text
                SaveSetting "Martoeng\MineSweeper", "Options", "UserY", txtHeight.Text
                SaveSetting "Martoeng\MineSweeper", "Options", "UserMines", txtMines.Text
                
                Unload Me
            Else
                MsgBox "Kein gültiger numerischer Ausdruck.", vbExclamation, "Fehler"
                txtMines.SetFocus
            End If
        Else
            MsgBox "Kein gültiger numerischer Ausdruck.", vbExclamation, "Fehler"
            txtHeight.SetFocus
        End If
    Else
        MsgBox "Kein gültiger numerischer Ausdruck.", vbExclamation, "Fehler"
        txtWidth.SetFocus
    End If
End Sub

'Mit Startwerten initialisieren
Private Sub Form_Load()
    'Werte aus der Registrierung holen
    txtWidth.Text = GetSetting("Martoeng\MineSweeper", "Options", "UserX", Fields_X)
    txtHeight.Text = GetSetting("Martoeng\MineSweeper", "Options", "UserX", Fields_Y)
    txtMines.Text = GetSetting("Martoeng\MineSweeper", "Options", "UserMines", MineCount)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UserCancel = True
End Sub
