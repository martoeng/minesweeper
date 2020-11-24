VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MineSweeper"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   240
   End
   Begin VB.PictureBox picSeconds 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3360
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdNew 
      DownPicture     =   "frmMain.frx":0CCA
      Height          =   615
      Left            =   1763
      Picture         =   "frmMain.frx":0FD4
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3150
      Left            =   1680
      Picture         =   "frmMain.frx":1416
      ScaleHeight     =   210
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox picMines 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox picMine 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3360
      Left            =   2040
      Picture         =   "frmMain.frx":31E0
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPlay 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Zentriert
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Spiel"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&Neu"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameStarter 
         Caption         =   "&Anfänger"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGameAdvanced 
         Caption         =   "&Fortgeschrittene"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuGameProfessional 
         Caption         =   "&Profis"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuGameUser 
         Caption         =   "Benutzer&definiert..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuGameBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameQuestion 
         Caption         =   "&Merker (?)"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameSound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGameBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameBestTimes 
         Caption         =   "Best&zeiten..."
      End
      Begin VB.Menu mnuGameBreak4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameQuit 
         Caption         =   "&Beenden"
      End
   End
End
Attribute VB_Name = "frmMain"
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

'Neues Spiel
Private Sub cmdNew_Click()
    tmrTime.Enabled = False
    'Anzahl der gesetzten Fahnen auf 0
    FlagCount = 0
    'Anzahl der geklickten Felder auf 0
    Clicked = 0
    'Alle Felder löschen
    mDraw.ClearFields
    'Minen intialisieren
    mDraw.InitMines
    mDraw.DrawArea
    
    'Zeichnen
    DrawMineCount
    DrawTime
    picPlay.Refresh
    picPlay.Enabled = True
    'Sekunden auf 0
    Seconds = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyP Or KeyCode = vbKeyPause Then
        'Pause ein/aus
        If picPlay.Visible = False Then
            picPlay.Visible = True
            tmrTime.Enabled = IIf(Clicked > 0, True, False)
        Else
            tmrTime.Enabled = False
            picPlay.Visible = False
        End If
    End If
End Sub

'Formular wird geladen
Private Sub Form_Load()
    'Zufallsgenerator initialisieren, dasss nicht immer die gleichen Zahlen kommen
    Randomize
    
    'Werte aus der Registrierung holen
    mnuGameSound.Checked = IIf(GetSetting("Martoeng\MineSweeper", "Options", "Sound", "Yes") = "Yes", True, False)
    mnuGameQuestion.Checked = IIf(GetSetting("Martoeng\MineSweeper", "Options", "Marker", "Yes") = "Yes", True, False)
    
    'Spielmodus auswählen
    Select Case GetSetting("Martoeng\MineSweeper", "Options", "Level", "1")
        Case "1" 'Anfänger
            mnuGameStarter_Click
        Case "2" 'Fortgeschritten
            mnuGameAdvanced_Click
        Case "3" 'Profi
            mnuGameProfessional_Click
        Case "4" 'Benutzerdefiniert
            If IsNumeric(GetSetting("Martoeng\MineSweeper", "Options", "UserX", "")) And IsNumeric(GetSetting("Martoeng\MineSweeper", "Options", "UserY", "")) And IsNumeric(GetSetting("Martoeng\MineSweeper", "Options", "UserMines", "")) Then
                mPlay.InitArea GetSetting("Martoeng\MineSweeper", "Options", "UserX", "9"), GetSetting("Martoeng\MineSweeper", "Options", "UserY", 9), GetSetting("Martoeng\MineSweeper", "Options", "UserMines", 10)
                mDraw.InitMines
                mDraw.DrawArea
                mDraw.DrawMineCount
                mDraw.DrawTime
                mnuGameUser.Checked = True
            Else
                mnuGameStarter_Click
            End If
    End Select
            
End Sub

'Formular wird entfernt
Private Sub Form_Terminate()
    picPlay.AutoRedraw = False
    Set picPlay.Picture = Nothing
    picMine.AutoRedraw = False
    Set picMine.Picture = Nothing
    picTime.AutoRedraw = False
    Set picTime.Picture = Nothing
    picMines.AutoRedraw = False
    Set picMines.Picture = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Werte in die Registrierung schreiben
    SaveSetting "Martoeng\MineSweeper", "Options", "Marker", IIf(mnuGameQuestion.Checked, "Yes", "No")
    SaveSetting "Martoeng\MineSweeper", "Options", "Sound", IIf(mnuGameSound.Checked, "Yes", "No")
    If mnuGameStarter.Checked Then
        SaveSetting "Martoeng\MineSweeper", "Options", "Level", "1"
    ElseIf mnuGameAdvanced.Checked Then
        SaveSetting "Martoeng\MineSweeper", "Options", "Level", "2"
    ElseIf mnuGameAdvanced.Checked Then
        SaveSetting "Martoeng\MineSweeper", "Options", "Level", "3"
    Else
        SaveSetting "Martoeng\MineSweeper", "Options", "Level", "4"
    End If
End Sub

'Fortgeschrittene
Private Sub mnuGameAdvanced_Click()
    mnuGameAdvanced.Checked = True
    mnuGameStarter.Checked = False: mnuGameProfessional.Checked = False: mnuGameUser.Checked = False
    mPlay.InitArea 16, 16, 40
    mDraw.InitMines
    mDraw.DrawArea
    mDraw.DrawTime
    mDraw.DrawMineCount
    
    'Neues Spiel
    cmdNew_Click
End Sub

Private Sub mnuGameBestTimes_Click()
    frmBestList.Show vbModal
End Sub

'Neues Spiel
Private Sub mnuGameNew_Click()
    cmdNew_Click
End Sub

'Professionell
Private Sub mnuGameProfessional_Click()
    mnuGameProfessional.Checked = True
    mnuGameAdvanced.Checked = False: mnuGameStarter.Checked = False: mnuGameUser.Checked = False
    mPlay.InitArea 30, 16, 99
    mDraw.InitMines
    mDraw.DrawArea
    mDraw.DrawTime
    mDraw.DrawMineCount
    
    'Neues Spiel
    cmdNew_Click
End Sub

Private Sub mnuGameQuestion_Click()
    mnuGameQuestion.Checked = Not mnuGameQuestion.Checked
End Sub

'Spiel verlassen
Private Sub mnuGameQuit_Click()
    Unload Me
End Sub

'Sound an/aus
Private Sub mnuGameSound_Click()
    mnuGameSound.Checked = Not mnuGameSound.Checked
End Sub

'Anfänger
Private Sub mnuGameStarter_Click()
    mnuGameStarter.Checked = True
    mnuGameAdvanced.Checked = False: mnuGameProfessional.Checked = False: mnuGameUser.Checked = False
    mPlay.InitArea 9, 9, 10
    mDraw.InitMines
    mDraw.DrawArea
    mDraw.DrawTime
    mDraw.DrawMineCount
    
    'Neues Spiel
    cmdNew_Click
End Sub

'Benutzerdefiniert
Private Sub mnuGameUser_Click()
Dim b As Boolean
    b = tmrTime.Enabled
    tmrTime.Enabled = False
    mPlay.ShowUserDefined
    If mPlay.UserCancel = False Then
        mnuGameUser.Checked = True
        mnuGameAdvanced.Checked = False: mnuGameProfessional.Checked = False: mnuGameStarter.Checked = False
    Else
        tmrTime.Enabled = b
    End If
End Sub

'Die Maus wird gedrückt
Private Sub picPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fld As Integer
    If tmrTime.Enabled = False Then tmrTime.Enabled = True
    fld = Int(X / 16) + (Int(Y / 16) * Fields_X)
    'Sound abspielen, falls aktiviert
    If mnuGameSound.Checked Then PlaySound sound_Click
    If Button = 1 Then
        'Feld klicken
        mDraw.ClickField (fld), False
        'Ansicht erneuern
        picPlay.Refresh
    ElseIf Button = 2 Then
        'Überprüfen, welches Symbol angezeigt wird (Standard/Fahne/Fragezeichen)
        Select Case GetGraphic(fld)
            Case gc_Default
                If FlagCount < MineCount Then
                    SetGraphic fld, gc_Flag
                    FlagCount = FlagCount + 1
                    DrawMineCount
                End If
            Case gc_Flag
                If mnuGameQuestion.Checked Then
                    SetGraphic fld, gc_Question
                Else
                    SetGraphic fld, gc_Default
                End If
                FlagCount = FlagCount - 1
                DrawMineCount
            Case gc_Question
                SetGraphic fld, gc_Default
        End Select
    End If
End Sub

'Wenn das Bildfeld in der Größe verändert wird
Private Sub picPlay_Resize()
    Me.Width = picPlay.Width * Screen.TwipsPerPixelX + picPlay.Left * Screen.TwipsPerPixelX * 3
    Me.Height = picPlay.Height * Screen.TwipsPerPixelY + picPlay.Top * Screen.TwipsPerPixelY * 2
    picSeconds.Left = picPlay.Left + picPlay.Width - picSeconds.Width
    cmdNew.Left = Me.ScaleWidth / 2 - cmdNew.Width / 2
    picPlay.Refresh
    lblPause.Left = picPlay.Left + picPlay.Width / 2 - lblPause.Width / 2
    lblPause.Top = picPlay.Top + picPlay.Height / 2 - lblPause.Height / 2
End Sub

'Eine Sekunde addieren
Private Sub tmrTime_Timer()
    'Maximaler zeitraum sind 999 Sekunden (sollte eigentlich reichen)
    If Seconds < 999 Then
        Seconds = Seconds + 1
    Else
        Lose "Das Zeitlimit von 999 Sekunden wurde überschritten. Sie haben leider verloren."
    End If
    'Zeit als Digitalzahlen anzeigen
    DrawTime
End Sub
