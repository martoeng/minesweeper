Attribute VB_Name = "mDeclarations"
Option Explicit
'Dieser Sourcecode stammt von http://www.martoeng.de
'Alle Rechte liegen beim Autor Martin Walter.
'Für eine Veröffentlichung benötigen Sie die Zustimmung des Autors.
'Sie dürfen diesen Code für private Zwecke gebrauchen, nicht _
jedoch verkaufen oder für andere finanzielle Zwecke verwenden.

'Win32-API-Deklarationen
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_NODEFAULT = &H2   ' Don't use default sound
Public Const SND_FILENAME = &H20000
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Eigene Enumerationen
Public Enum GRAPHIC_CONSTANTS
    gc_Default = 0  'Normales Feld
    gc_Free = 1     'Bereits geklicktes aber freies Feld
    gc_Flag = 2     'Feld mit Fahne
    gc_Question = 3 'Feld mit Fragezeichen
    gc_Mine = 4     'Mine
    gc_RedMine = 5  'Mine mit rotem Hintergrund
    gc_One = 6      'Feld mit 1
    gc_Two = 7      'Feld mit 2
    gc_Three = 8    'Feld mit 3
    gc_Four = 9     'Feld mit 4
    gc_Five = 10    'Feld mit 5
    gc_Six = 11     'Feld mit 6
    gc_Seven = 12   'Feld mit 7
    gc_Eight = 13   'Feld mit 8
End Enum

Public Enum VALUE_CONSTANTS
    vc_Free = 1
    vc_Mine = 4
    vc_RedMine = 5
    vc_One = 6
    vc_Two = 7
    vc_Three = 8
    vc_Four = 9
    vc_Five = 10    'Feld mit 5
    vc_Six = 11     'Feld mit 6
    vc_Seven = 12   'Feld mit 7
    vc_Eight = 13   'Feld mit 8
End Enum

Public Enum SOUND_CONSTANTS
    sound_Click = 0
    sound_Won = 1
    sound_Lost = 2
End Enum

'Eigene Typendeklaration
Public Type FIELD
    Graphic As GRAPHIC_CONSTANTS    'Welche Grafik wird für das Feld angezeigt
    Value As VALUE_CONSTANTS        'Welchen (verborgenen) Wert hat das Feld
    Clicked As Boolean              'Wurde schon geklickt?
End Type
