Attribute VB_Name = "mDraw"
Option Explicit
'Dieser Sourcecode stammt von http://www.martoeng.de
'Alle Rechte liegen beim Autor Martin Walter.
'Für eine Veröffentlichung benötigen Sie die Zustimmung des Autors.
'Sie dürfen diesen Code für private Zwecke gebrauchen, nicht _
jedoch verkaufen oder für andere finanzielle Zwecke verwenden.

Dim m_Fields() As FIELD 'Array der Felder

'Array neu dimensionieren
Public Sub ReDimArray(NewLength As Integer)
    ReDim m_Fields(0 To (NewLength - 1))
End Sub

'Den Wert eines Feldes setzen
Public Sub SetValue(Index As Integer, New_Value As VALUE_CONSTANTS)
    m_Fields(Index).Value = New_Value
End Sub

'Geklickt setzen
Public Sub SetClicked(Index As Integer, Optional ByVal New_Clicked As Boolean = True)
    If m_Fields(Index).Clicked = False Then
        m_Fields(Index).Clicked = New_Clicked
        If New_Clicked Then mPlay.Clicked = mPlay.Clicked + 1
    Else
        m_Fields(Index).Clicked = New_Clicked
    End If
    If Clicked = Fields_X * Fields_Y - MineCount Then
        Win
    End If
End Sub

'Grafik ermitteln
Public Function GetGraphic(Index As Integer) As GRAPHIC_CONSTANTS
    GetGraphic = m_Fields(Index).Graphic
End Function

'Grafik setzen (für Flag und Fragezeichen)
Public Sub SetGraphic(Index As Integer, NewGraphic As GRAPHIC_CONSTANTS, Optional ByVal bReDraw As Boolean = True)
    m_Fields(Index).Graphic = NewGraphic
    DrawField Index, bReDraw
End Sub

'Ein Feld wird gedrückt
Public Sub ClickField(ByVal Index As Integer, Optional ByVal bReDraw As Boolean = True)
    If m_Fields(Index).Clicked = False Then
        If m_Fields(Index).Value = vc_Mine Then m_Fields(Index).Value = vc_RedMine
        m_Fields(Index).Graphic = m_Fields(Index).Value
        DrawField Index, bReDraw
        SetClicked Index, True
        If m_Fields(Index).Value = vc_RedMine Then
            Lose
            Exit Sub
        End If
        
        'Nachbarfelder
        'Oben links
        If m_Fields(Index).Value = vc_Free Then
Dim fld As Integer
            fld = GetTopLeftField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
            'Oben
            fld = GetTopField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
            'Oben rechts
            fld = GetTopRightField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
            'Links
            fld = GetLeftField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
            'Rechts
            fld = GetRightField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
            'Unten links
            fld = GetBottomLeftField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
            'Unten
            fld = GetBottomField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
            'Unten rechts
            fld = GetBottomRightField(Index)
            If fld <> -1 Then
                If Not m_Fields(fld).Graphic = gc_Flag Then
                    If m_Fields(fld).Value = vc_Free Then
                        ClickField fld, False
                    ElseIf m_Fields(fld).Value >= vc_One Then
                        SetClicked fld, True
                        m_Fields(fld).Graphic = m_Fields(fld).Value
                        DrawField fld, False
                    End If
                End If
            End If
        End If
    End If
End Sub

'Bildfeld neu zeichnen
Public Sub ReDraw()
    frmMain.picPlay.Refresh
End Sub

'Ein Feld zeichnen
Public Sub DrawField(Index As Integer, Optional ByVal bReDraw As Boolean = True)
    BitBlt frmMain.picPlay.hDC, GetCol(Index) * 16, GetRow(Index) * 16, 16, 16, frmMain.picMine.hDC, 0, m_Fields(Index).Graphic * 16, vbSrcCopy
    If bReDraw Then frmMain.picPlay.Refresh
End Sub

'Alle Felder zeichnen
Public Sub DrawArea()
    Dim X As Integer, Y As Integer, produkt As Integer, fld As Integer
    produkt = Fields_X * Fields_Y
    For Y = 0 To Fields_Y
        For X = 0 To Fields_X
            fld = Y * Fields_X + X
            If fld < produkt Then mDraw.DrawField fld, False
        Next X
    Next Y
    frmMain.picPlay.Refresh
End Sub

'Verteilt die Minen
Public Sub InitMines()
    Dim n As Integer, fld As Integer
    For n = 1 To MineCount
        Do
            fld = Int(Fields_X * Fields_Y * Rnd)
        Loop While m_Fields(fld).Value = vc_Mine
        
        m_Fields(fld).Value = vc_Mine
    Next n
    
    For n = 0 To Fields_X * Fields_Y - 1
        If m_Fields(n).Value = vc_Mine Then
            'Nachbarfelder setzen
            IncreaseField GetTopLeftField(n)
            IncreaseField GetTopField(n)
            IncreaseField GetTopRightField(n)
            IncreaseField GetLeftField(n)
            IncreaseField GetRightField(n)
            IncreaseField GetBottomLeftField(n)
            IncreaseField GetBottomField(n)
            IncreaseField GetBottomRightField(n)
        ElseIf m_Fields(n).Value = 0 Then
            m_Fields(n).Value = vc_Free
        End If
    Next n
End Sub

'Wenn keine Mine um eins erhöhen oder beim ersten Eintrag auf vc_One setzen
Public Sub IncreaseField(fld As Integer)
    If fld <> -1 Then
        If m_Fields(fld).Value <> vc_Mine Then
            If m_Fields(fld).Value >= vc_One Then
                m_Fields(fld).Value = m_Fields(fld).Value + 1
            Else
                m_Fields(fld).Value = vc_One
            End If
        End If
    End If
End Sub

'Die verschiedenen Nachbarfelder
Public Function GetLeftField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index - 1
    If GetRow(fld) < GetRow(Index) Then
        GetLeftField = -1
    Else
        GetLeftField = fld
    End If
End Function

Public Function GetRightField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index + 1
    If GetRow(fld) <> GetRow(Index) Or fld >= Fields_X * Fields_Y Or GetCol(fld) <> GetCol(Index) + 1 Then
        GetRightField = -1
    Else
        GetRightField = fld
    End If
End Function

Public Function GetTopField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index - Fields_X
    If fld < 0 Then
        GetTopField = -1
    Else
        GetTopField = fld
    End If
End Function

Public Function GetBottomField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index + Fields_X
    If fld > Fields_X * Fields_Y - 1 Or GetCol(fld) <> GetCol(Index) Then
        GetBottomField = -1
    Else
        GetBottomField = fld
    End If
End Function

Public Function GetTopLeftField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index - Fields_X - 1
    If fld < 0 Or GetRow(fld) <= GetRow(Index) - 2 Then
        GetTopLeftField = -1
    Else
        GetTopLeftField = fld
    End If
End Function

Public Function GetTopRightField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index - Fields_X + 1
    If (GetCol(fld) <> (GetCol(Index) + 1)) Or (fld < 0) Then
        GetTopRightField = -1
    Else
        GetTopRightField = fld
    End If
End Function

Public Function GetBottomRightField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index + Fields_X + 1
    If fld > Fields_X * Fields_Y - 1 Or GetCol(Index) <> GetCol(fld) - 1 Then
        GetBottomRightField = -1
    Else
        GetBottomRightField = fld
    End If
End Function

Public Function GetBottomLeftField(Index As Integer) As Integer
Dim fld As Integer
    fld = Index + Fields_X - 1
    If fld > Fields_X * Fields_Y - 1 Or GetCol(Index) <> GetCol(fld) + 1 Then
        GetBottomLeftField = -1
    Else
        GetBottomLeftField = fld
    End If
End Function

'Alle Felder leeren
Public Sub ClearFields()
Dim n As Integer
    For n = 0 To Fields_X * Fields_Y - 1
        m_Fields(n).Clicked = False: m_Fields(n).Graphic = gc_Default: m_Fields(n).Value = vc_Free
        DrawField n, False
    Next n
End Sub

'Minenanzahl zeichnen
Public Sub DrawMineCount()
    Dim s As String
    s = CStr(MineCount - FlagCount)
    Select Case Len(s)
        Case 1
            BitBlt frmMain.picMines.hDC, 1, 1, 11, 21, frmMain.picTime.hDC, 0, 0, vbSrcCopy
            BitBlt frmMain.picMines.hDC, 15, 1, 11, 21, frmMain.picTime.hDC, 0, 0, vbSrcCopy
            BitBlt frmMain.picMines.hDC, 29, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(s) * 21, vbSrcCopy
        Case 2
            BitBlt frmMain.picMines.hDC, 1, 1, 11, 21, frmMain.picTime.hDC, 0, 0, vbSrcCopy
            BitBlt frmMain.picMines.hDC, 15, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 1, 1)) * 21, vbSrcCopy
            BitBlt frmMain.picMines.hDC, 29, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 2, 1)) * 21, vbSrcCopy
        Case 3
            BitBlt frmMain.picMines.hDC, 1, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 1, 1)) * 21, vbSrcCopy
            BitBlt frmMain.picMines.hDC, 15, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 2, 1)) * 21, vbSrcCopy
            BitBlt frmMain.picMines.hDC, 29, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 3, 1)) * 21, vbSrcCopy
    End Select
    frmMain.picMines.Refresh
End Sub

'Zeit zeichnen
Public Sub DrawTime()
    Dim s As String
    s = CStr(Seconds)
    Select Case Len(s)
        Case 1
            BitBlt frmMain.picSeconds.hDC, 1, 1, 11, 21, frmMain.picTime.hDC, 0, 0, vbSrcCopy
            BitBlt frmMain.picSeconds.hDC, 15, 1, 11, 21, frmMain.picTime.hDC, 0, 0, vbSrcCopy
            BitBlt frmMain.picSeconds.hDC, 29, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(s) * 21, vbSrcCopy
        Case 2
            BitBlt frmMain.picSeconds.hDC, 1, 1, 11, 21, frmMain.picTime.hDC, 0, 0, vbSrcCopy
            BitBlt frmMain.picSeconds.hDC, 15, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 1, 1)) * 21, vbSrcCopy
            BitBlt frmMain.picSeconds.hDC, 29, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 2, 1)) * 21, vbSrcCopy
        Case 3
            BitBlt frmMain.picSeconds.hDC, 1, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 1, 1)) * 21, vbSrcCopy
            BitBlt frmMain.picSeconds.hDC, 15, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 2, 1)) * 21, vbSrcCopy
            BitBlt frmMain.picSeconds.hDC, 29, 1, 11, 21, frmMain.picTime.hDC, 0, CInt(Mid(s, 3, 1)) * 21, vbSrcCopy
    End Select
    frmMain.picSeconds.Refresh
End Sub

'Verloren
Public Sub Lose(Optional ByVal sMsgText As String)
    Dim n As Integer
    For n = 0 To Fields_X * Fields_Y - 1
        If m_Fields(n).Value = vc_Mine Then
            m_Fields(n).Graphic = gc_Flag
            DrawField n, False
        End If
        If m_Fields(n).Graphic = gc_Default Then
            m_Fields(n).Graphic = m_Fields(n).Value
            DrawField n, False
        End If
    Next n
    ReDraw
    frmMain.tmrTime.Enabled = False
    frmMain.picPlay.Enabled = False
    If frmMain.mnuGameSound.Checked Then PlaySound sound_Lost
    If sMsgText <> "" Then MsgBox sMsgText, vbExclamation, "Pech gehabt"
End Sub

'Gewinnen
Public Sub Win()
    Dim n As Integer
    For n = 0 To Fields_X * Fields_Y - 1
        If m_Fields(n).Value = vc_Mine Then
            m_Fields(n).Graphic = gc_Flag
            DrawField n, False
        End If
        If m_Fields(n).Graphic = gc_Default Then
            m_Fields(n).Graphic = m_Fields(n).Value
            DrawField n, False
        End If
    Next n
    ReDraw
    frmMain.tmrTime.Enabled = False
    frmMain.picPlay.Enabled = False
    
    'Sound abspielen falls gewünscht
    If frmMain.mnuGameSound.Checked Then PlaySound sound_Won
    MsgBox "Sie haben gewonnen.", vbInformation, "Glückwunsch"
    
    'Bestliste anzeigen, falls nicht "Benutzerdefiniert"
    If frmMain.mnuGameUser.Checked = False Then
        Load frmBestList
        With frmBestList
            If frmMain.mnuGameStarter.Checked Then
                .txtName.Visible = True
                .lblPoints(1).Caption = Seconds
            ElseIf frmMain.mnuGameAdvanced.Checked Then
                .txtName.Top = .lblName(2).Top
                .lblPoints(2).Caption = Seconds
                .txtName.Visible = True
            ElseIf frmMain.mnuGameProfessional.Checked Then
                .txtName.Top = .lblName(3).Top
                .lblPoints(3).Caption = Seconds
                .txtName.Visible = True
            End If
        End With
        frmBestList.Show vbModal
    End If
End Sub


