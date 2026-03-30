Attribute VB_Name = "MMakros"
Option Explicit
Dim UndoRedo As UndoRedoCell

Sub FormatFormula()
    If UndoRedo Is Nothing Then Set UndoRedo = New UndoRedoCell
    UndoRedo.Save Excel.ActiveCell
    ReplaceGreekChars
    ReplaceRootChar
    ReplaceIndices
    ReplaceExponent
    ReplaceMathOps
    ReplaceGTEQULTEQU
    Application.OnUndo "Undo Makro FormatFormula", "UndoFormatFormula"
End Sub

Public Sub ShowFrmClipBoard()
    UserForm1.Show vbModeless
End Sub

Private Sub UndoFormatFormula()
    'With UndoRedo: If .CanUndo Then .Undo: End With
    If UndoRedo.CanUndo Then
        UndoRedo.Undo
        'geht leider nicht:
        'Application.OnUndo "Redo Makro FormatFormula", "RedoFormatFormula"
    End If
End Sub
'Private Sub RedoFormatFormula()
'    If UndoRedo.CanRedo Then
'        UndoRedo.Redo
'        Application.OnUndo "Undo Makro FormatFormula", "UndoFormatFormula"
'    End If
'End Sub

Private Sub ReplaceGreekChars()
    Replace4 "alpha", "alp", "a", "Alpha", "Alp", "A"
    Replace4 "beta", "bet", "b", "Beta", "Bet", "B"
    Replace4 "chi", "", "c", "Chi", "", "C"
    Replace4 "gamma", "gam", "g", "Gamma", "Gam", "G"
    Replace4 "delta", "del", "d", "Delta", "Del", "D"
    Replace4 "epsilon", "eps", "e", "Epsilon", "Eps", "E"
    Replace4 "phi", "", "j", "Phi", "", "F"
    Replace4 "eta", "", "h", "Eta", "", "H"
    Replace4 "theta", "tht", "J", "Theta", "Tht", "Q"
    Replace4 "iot", "iota", "i", "Iot", "Iota", "I"
    Replace4 "kappa", "kap", "k", "Kappa", "Kap", "K"
    Replace4 "lambda", "lam", "l", "Lambda", "Lam", "L"
    Replace4 "mue", "", "m", "Mue", "", "M"
    Replace4 "nue", "", "n", "Nue", "", "N"
    Replace4 "omega", "omg", "w", "Omega", "Omg", "W"
    Replace4 "omikron", "omi", "o", "Omikron", "Omi", "O"
    Replace4 "pi", "", "p", "Pi", "", "P"
    Replace4 "rho", "", "r", "Rho", "", "R"
    Replace4 "sigma", "sig", "s", "Sigma", "Sig", "S"
    Replace4 "tau", "", "t", "Tau", "", "T"
    Replace4 "xi", "", "x", "Xi", "", "X"
    Replace4 "yi", "", "y", "Yi", "", "Y"
    Replace4 "ypsilon", "yps", "u", "Ypsilon", "Yps", "U"
    Replace4 "zeta", "zet", "z", "Zeta", "Zet", "Z"
End Sub

Private Sub Replace4(var11 As String, var12 As String, repby1 As String, var21 As String, var22 As String, repby2 As String)
    Replace2 var11, var12, repby1
    Replace2 var21, var22, repby2
End Sub

Private Sub Replace2(var1 As String, var2 As String, repby As String)
    If Len(var1) Then CellReplace ActiveCell, var1, repby, "Symbol", vbBinaryCompare
    If Len(var2) Then CellReplace ActiveCell, var2, repby, "Symbol", vbBinaryCompare
End Sub

Private Sub CellReplace(Cell As Excel.Range, ByVal sFind As String, ByVal sRepl As String, Optional FontName As String = "", Optional ByVal comp As VbCompareMethod = vbTextCompare)
    'Wenn die Länge des zu findenden Textes sFind ungleich der Länge des Ersatz-Textes ist dann
    '   Wenn die Länge des zu findenden Textes sFind kleiner als die Länge des Ersatz-Textes ist dann
    '       müssen nach sFind zusätzliche Character eingefügt werden, öha wie soll das gehen?
    '   Sonst
    '       müssen Teile von sFind gelöscht werden, OK schon gemacht ist einfach
    '   Ende Wenn
    'Ende Wenn
    Dim c As Characters
    Dim i As Long, n As Long: n = ActiveCell.Characters.Count
    Dim c1 As String: c1 = Left(sFind, 1)
    Dim bFound As Boolean
    Do
        i = i + 1: If i > n Then Exit Do
        Set c = ActiveCell.Characters(i, 1)
        If StrComp(c.Text, c1, comp) = 0 Then
            bFound = True ' possibly, probably
            'OK jetzt ab hier vergleichen mit sFind
            If n < (i + Len(sFind)) Then Exit Sub
            Dim ii As Long: ii = i
            Dim j As Long, cc As String
            For j = 2 To Len(sFind)
                cc = Mid(sFind, j, 1)
                ii = ii + 1: If ii > n Then Exit Do
                Set c = ActiveCell.Characters(ii, 1)
                If StrComp(c.Text, cc, comp) <> 0 Then
                    bFound = False
                    Exit For
                End If
            Next
            If bFound Then
                i = i - 1
                'jetzt die Ersetzung vornehmen
                Dim m As Long: m = Min(Len(sFind), Len(sRepl))
                'zuerst die minimale Anzahl
                For j = 1 To m
                    i = i + 1: If i > n Then Exit Do
                    Set c = ActiveCell.Characters(i, 1)
                    c.Text = Mid(sRepl, j, 1)
                    If Len(FontName) Then
                        c.Font.Name = FontName
                    End If
                Next
                'jetzt den Rest
                Dim d As Long: d = Len(sFind) - Len(sRepl)
                If d > 0 Then
                    'löschen
                    i = i + 1
                    For j = 1 To Abs(d)
                        Set c = ActiveCell.Characters(i, 1)
                        c.Text = ""
                    Next
                    n = ActiveCell.Characters.Count
                ElseIf d < 0 Then
                    'einfügen
                    For j = m To m + Abs(d)
                        'cc = Mid(sRepl, j, 1)
                        'i = i + 1
                        ActiveCell.Characters(i, 1).Insert Mid(sRepl, j, 1)
                    Next
                End If
            End If
        End If
        If i = n Then Exit Do
    Loop
End Sub
Private Function Min(v1, v2)
    If v1 < v2 Then Min = v1 Else Min = v2
End Function
Private Sub ReplaceGTEQULTEQU()
    CellReplace ActiveCell, "<=", ChrW(8804), , vbBinaryCompare
    CellReplace ActiveCell, ">=", ChrW(8805), , vbBinaryCompare
End Sub
Private Sub ReplaceMathOps()
    CellReplace ActiveCell, "*", ChrW(183), , vbBinaryCompare
    CellReplace ActiveCell, "+-", ChrW(177), , vbBinaryCompare
    '· × ±
    '<, =, >, =
    '?
    '?
    '8
    '?, ?,
    '?, ?
    '?
    '?, ?, ?
End Sub
Private Sub ReplaceRootChar()
    
    Dim i As Long, n As Long: n = ActiveCell.Characters.Count
    Dim c As Characters, iw As Long
    Do
        i = i + 1: If i > n Then Exit Do
        Set c = ActiveCell.Characters(i, 1)
        If StrComp(c.Text, "W", vbTextCompare) = 0 Then
            iw = i
            i = i + 1: If i > n Then Exit Do
            Set c = ActiveCell.Characters(i, 1)
            If StrComp(c.Text, "u", vbTextCompare) = 0 Then
                i = i + 1: If i > n Then Exit Do
                Set c = ActiveCell.Characters(i, 1)
                If StrComp(c.Text, "r", vbTextCompare) = 0 Then
                    i = i + 1: If i > n Then Exit Do
                    Set c = ActiveCell.Characters(i, 1)
                    If StrComp(c.Text, "z", vbTextCompare) = 0 Then
                        i = i + 1: If i > n Then Exit Do
                        Set c = ActiveCell.Characters(i, 1)
                        If StrComp(c.Text, "e", vbTextCompare) = 0 Then
                            i = i + 1: If i > n Then Exit Do
                            Set c = ActiveCell.Characters(i, 1)
                            If StrComp(c.Text, "l", vbTextCompare) = 0 Then
                                c.Text = "": i = i - 1
                                Set c = ActiveCell.Characters(i, 1): c.Text = "": i = i - 1
                                Set c = ActiveCell.Characters(i, 1): c.Text = "": i = i - 1
                                Set c = ActiveCell.Characters(i, 1): c.Text = "": i = i - 1
                                Set c = ActiveCell.Characters(i, 1): c.Text = "": i = i - 1
                                Set c = ActiveCell.Characters(i, 1): c.Text = "Ö"
                                ActiveCell.Characters(i, 1).Font.Name = "Symbol"
                                n = ActiveCell.Characters.Count
                                Dim ip As Long: ip = iw - 1
                                If ip > 0 Then
                                    Set c = ActiveCell.Characters(ip, 1)
                                    If StrComp(c.Text, ".", vbTextCompare) = 0 Then
                                        ip = ip - 1
                                        If ip > 0 Then
                                            Set c = ActiveCell.Characters(ip, 1)
                                            If InStr(1, "0123456789", c.Text) > 0 Then
                                                c.Font.Superscript = True
                                                ip = ip + 1
                                                Set c = ActiveCell.Characters(ip, 1)
                                                c.Text = "": i = i - 1
                                                n = ActiveCell.Characters.Count
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If i = n Then Exit Do
    Loop

End Sub
Private Sub ReplaceIndices()
    Dim i As Long, n As Long: n = ActiveCell.Characters.Count
    Dim c As Characters
    Do
        i = i + 1: If i > n Then Exit Do
        Set c = ActiveCell.Characters(i, 1)
        If StrComp(c.Text, "_", vbBinaryCompare) = 0 Then
            c.Text = "": i = i - 1
            n = ActiveCell.Characters.Count
            Do
                'jetzt so lange bis kein alphanum/Komma
                i = i + 1: If i > n Then Exit Do
                Set c = ActiveCell.Characters(i, 1)
                If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ0123456789.,", c.Text, vbTextCompare) > 0 Then
                    c.Font.Subscript = True
                Else
                    Exit Do
                End If
                If i = n Then Exit Do
            Loop
        End If
        If i = n Then Exit Do
    Loop
End Sub

Private Sub ReplaceExponent()
    Dim i As Long, n As Long: n = ActiveCell.Characters.Count
    Dim c As Characters
    Do
        i = i + 1: If i > n Then Exit Do
        Set c = ActiveCell.Characters(i, 1)
        If StrComp(c.Text, "^", vbBinaryCompare) = 0 Then
            'OK ab hier Text nach oben, bis " " oder ")+-*/"
            c.Text = "": i = i - 1
            Do
                i = i + 1: If i > n Then Exit Do
                Set c = ActiveCell.Characters(i, 1)
                If StrComp(c.Text, "(", vbTextCompare) = 0 Then
                    c.Text = "": i = i - 1
                    'jetzt alle nach oben bis ")"
                    'c.Font.Superscript = True
                    Do
                        i = i + 1: If i > n Then Exit Do
                        Set c = ActiveCell.Characters(i, 1)
                        c.Font.Superscript = True
                        If StrComp(c.Text, ")", vbTextCompare) = 0 Then
                            c.Text = "": i = i - 1
                            Exit Do
                        End If
                        If i = n Then Exit Do
                    Loop
                Else
                    If InStr(1, " +-*/,()", c.Text, vbTextCompare) > 0 Then
                        Exit Do
                    Else
                        c.Font.Superscript = True
                    End If
                End If
                If i = n Then Exit Do
            Loop
        End If
        If i = n Then Exit Do
    Loop

End Sub
