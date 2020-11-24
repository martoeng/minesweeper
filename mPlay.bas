Attribute VB_Name = "mPlay"
Option Explicit
'Dieser Sourcecode stammt von http://www.martoeng.de
'Alle Rechte liegen beim Autor Martin Walter.
'Für eine Veröffentlichung benötigen Sie die Zustimmung des Autors.
'Sie dürfen diesen Code für private Zwecke gebrauchen, nicht _
jedoch verkaufen oder für andere finanzielle Zwecke verwenden.

'Funktionen zum Spielverlauf
'*************************************************************************
Public Fields_X As Integer  'Wieviele Kästchen breit?
Public Fields_Y As Integer  'Wieviele Kästchen hoch?
Public MineCount As Integer 'Wieviele Minen gibt es?
Public FlagCount As Integer 'Wieviele Flaggen wurden gesetzt?
Public Clicked As Integer   'Wieviele Kästchen wurden bereits geklickt?
Public Seconds As Integer   'Wieviele Sekunden wurden schon gespielt?

Public UserCancel As Boolean 'Benutzerdefiniert-Dialog abgebrochen?

'Spielfeld initialisieren (noch keine Werte setzen)
Public Sub InitArea(ByVal Hor_Fields As Integer, ByVal Ver_Fields As Integer, ByVal Mines As Integer)
    'Breite und Höhe
    Fields_X = Hor_Fields
    Fields_Y = Ver_Fields
    'Neu dimensionieren
    mDraw.ReDimArray (Hor_Fields * Ver_Fields)
    'Breite des Bildfeldes
    frmMain.picPlay.Width = Hor_Fields * 16
    frmMain.picPlay.Height = Ver_Fields * 16
    'Minen
    MineCount = Mines
    'Zeit
    Seconds = 0
    frmMain.picPlay.Visible = True
End Sub

'Die Reihe eines Feldes
Public Function GetRow(Index As Integer) As Integer
    GetRow = Int(Index / Fields_X)
End Function

'Die Spalte eines Feldes
Public Function GetCol(Index As Integer) As Integer
    GetCol = Index - (GetRow(Index) * Fields_X)
End Function

'"Benutzerdefiniert"-Dialog
Public Sub ShowUserDefined()
    UserCancel = False
    frmUser.Show vbModal, frmMain
End Sub

'Sound abspielen
Public Sub PlaySound(ByVal SoundType As SOUND_CONSTANTS)
    Dim sSoundFile As String
    
    'Dateinamen mit Pfad beginnen
    sSoundFile = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    
    'Sound auswählen
    Select Case SoundType
        Case sound_Click
            sSoundFile = sSoundFile & "click.wav"
        Case sound_Won
            sSoundFile = sSoundFile & "won.wav"
        Case sound_Lost
            sSoundFile = sSoundFile & "lost.wav"
    End Select
    
    'Asynchron abspielen; keinen Piepton, falls nicht vorhanden
    sndPlaySound sSoundFile, SND_FILENAME + SND_ASYNC + SND_NODEFAULT
End Sub
