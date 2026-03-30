Option Explicit
'Private Pi As Variant

Sub AddCRLFToClipB()
    Dim s As String
    'entweder so:
'    Dim ClipBoard As Object
'    Set ClipBoard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    'oder so:
    Dim ClipBoard As DataObject
    Set ClipBoard = New DataObject

    ClipBoard.GetFromClipboard
    s = ClipBoard.GetText
    
    Dim sa() As String: sa = Split(s, vbCrLf)
    Dim C As Long: C = UBound(sa) + 1
    ReDim sa_out(0 To 2 * C - 1) As String
    Dim i As Long, j As Long
    For i = 0 To UBound(sa)
        sa_out(j) = sa(i)
        j = j + 1
        'sa_out(j) = vbCrLf
        j = j + 1
    Next
    s = Join(sa_out, vbCrLf)
    ClipBoard.Clear
    ClipBoard.SetText s
    ClipBoard.PutInClipboard
End Sub

Sub Test()
    
    'Verweis auf: "Microsoft XML, v6.0" Pfad: "C:\Windows\System32\msxml6.dll"
    Dim xd As MSXML2.IXMLDOMDocument2
    Set xd = New MSXML2.DOMDocument60
    Dim at As MSXML2.IXMLDOMAttribute
    
    If xd.Load("C:\<File>.kml") Then
        
        Debug.Print xd.BaseName & " " & xd.nodeName
        If Not xd.Attributes Is Nothing Then
            Set at = xd.Attributes
            Debug.Print at.Name
        End If
        
    End If
    
    
'    Dim pos As Long
'    Dim ch As String, cl As Long
'    Dim sHex As String
'    pos = InStr(1, s, "\u")
'    ch = Mid(s, pos - 1, 1)
'    cl = AscW(ch)
'    If cl = 92 Then
'        'nö
'    Else
'        sHex = Mid(s, pos + 2, 4)
'        If IsHex(sHex) Then
'            cl = CLng("&H" & sHex)
'            ch = ChrW(cl)
'            s = Replace(s, "\u" & sHex, ch, pos, 6)
'        End If
'    End If
    
    Dim Shp As String: Shp = "<YourSharepointDirectoryHere>"
    Dim Flt As String: Flt = "PDF (*.pdf)|*.pdf"
    Dim Ttl As String: Ttl = "Als pdf speichern"
    Dim FNm 'As String
    FNm = Application.GetSaveAsFilename(Shp, Flt, 1, Ttl)
    If FNm <> False Then
        ActiveSheet.ExportAsFixedFormat xlTypePDF, FNm
        
    End If

'    Dim d As Double
'
'    d = 0.00000000000123
'
'
'
'
'
'
'
'    'Application.Volatile
'
'    Dim Pi:   Pi = CDec("3,14159265358979323846264338327")
'    '                    3,1415926535897932384626433833
'    Debug.Print Pi
'    Dim Pi2: Pi2 = CDec(CDec(4#) * Atn(CDec(1#)))
'
'    MsgBox "Pi : " & Pi & _
'           "Pi2: " & Pi2 & _
'           " (Pi == Pi2) = " & (Pi = Pi2)
'
'
'    'VB und die Kommentare
'    'in anderen Programmiersprachen sind mehrzeileige Kommentare in mehrzeiligen Anweisungen möglich
'    'in VB ist eine Zeile eine Anweisung, wenn sich eine Anweisung auf mehrere zeilen verteilen muss
'    'dann nur mit einem Zeilen-Fortsetzungszeichen am Ende der Zeile _
'    'danach ist allerdings kein Kommentar mehr möglich
'
'    'in VB sind mehrzeilige Anweisungen möglich mit "_"
'    If i And 1 Or _
'       i And 2 Or _
'       i And 3 Or _
'       i And 4 Then
'        MsgBox "OK"
'    End If
'
'    'in VB sind mehrzeilige Kommentare möglich mit "_" das "'"-Zeichen am Anfang kann in den darauffolgenden Zeilen entfallen
'    'Kommentar1 _
'     Kommentar2 _
'     Kommentar3 _
'     Kommentar4
'    MsgBox "OK"
'
'    'dies ist in VB bisher nicht möglich:
''    If i And 1 Or _   ' Kommentar1
''       i And 2 Or _   ' Kommentar2
''       i And 3 Or _   ' Kommentar3
''       i And 4 Then   ' Kommentar4
'
'        MsgBox "nopo"
'    End If
    
End Sub

Sub LinksZusammenfassen()
    'Dim wks As Worksheet: Set wks = ActiveWorkbook.ActiveSheet
    Dim r As Range: Set r = Excel.Selection
    Dim C As Range, s As String
    Dim Row As Long, Col As Long
    For Row = 1 To r.Rows.Count
        'von rechts
        For Col = r.Columns.Count To 1 Step -1
            Set C = r.Cells(Row, Col)
            s = C.Value & " " & s
            C.Value = ""
        Next
        If Left(s, 1) = "=" Then s = "'" & s
        C.Value = s
        s = ""
    Next
End Sub

Sub FormatNachweisRotGrün()
'
' FormatNachweisRotGrün Makro
' Vergleicht den Wert in der links danebenliegenden Zelle mit 1 (bzw. 1.03) und fügt "OK" oder "Achtung" ein
' Macht 2*bedingte Formatierung wenn Zelle "OK" enthält dann Grün, wenn Zelle "Achtung" enthält dann Rot.
'
    Dim C As Range: Set C = Selection
    
    C.FormulaR1C1 = "=IF(RC[-1]<=1.03, ""OK"", ""Achtung!"")"
    'Range("G43").Select

    C.FormatConditions.Add Type:=xlTextString, String:="OK", TextOperator:=xlContains
    C.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With C.FormatConditions(1).Font
        .color = -16752384
        .TintAndShade = 0
    End With
    With C.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = 13561798
        .TintAndShade = 0
    End With
    C.FormatConditions(1).StopIfTrue = False
        
    C.FormatConditions.Add Type:=xlTextString, String:="Achtung", TextOperator:=xlContains
    C.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With C.FormatConditions(1).Font
        .color = -16383844
        .TintAndShade = 0
    End With
    With C.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = 13551615
        .TintAndShade = 0
    End With
    C.FormatConditions(1).StopIfTrue = False
'
End Sub

Sub LinIPol_H()
'
' LinIPol_H Makro
' Lineare Interpolation Horizontal
'
' Tastenkombination: Strg+Umschalt+I
'

'LinIPol_H
'Formel in C4
'=WENN(B3 = D3; B4; B4 + (D4 - B4) / (D3 - B3) * (C3 - B3))

    Dim C As Range: Set C = Excel.Selection
    
    C.FormulaR1C1 = "=IF(R[-1]C[-1] = R[-1]C[+1], RC[-1], RC[-1] + (RC[+1] - RC[-1]) / (R[-1]C[+1] - R[-1]C[-1]) * (R[-1]C - R[-1]C[-1]))"
    
End Sub

Sub LinIPol_V()
'
' LinIPol_V Makro
' Lineare Interpolation Vertikal
'
' Tastenkombination: Strg+Umschalt+V
'

'LinIPol_V
'Formel in I3
'=WENN(H2 = H4; I2; I2 + (I4 - I2) / (H4 - H2) * (H3 - H2))
    
    Dim C As Range: Set C = Excel.Selection
    
    C.FormulaR1C1 = "=IF(R[-1]C[-1] = R[+1]C[-1], R[-1]C, R[-1]C + (R[+1]C - R[-1]C) / (R[+1]C[-1] - R[-1]C[-1]) * (RC[-1] - R[-1]C[-1]))"
    
End Sub

Sub BiIPol()
    
'    'c ist J16 bzw Spalte=10, Zeile=16
'    Dim c  As Range:      Set c = Excel.Selection 'die Zelle die den zu berechnenden Wert empfangen soll
'    Dim ws As Worksheet: Set ws = c.Worksheet
'    Dim co As Range:     Set co = ws.Cells(c.Row - 1, c.Column) ' die Zelle oberhalb  von c
'    Dim cl As Range:     Set cl = ws.Cells(c.Row, c.Column - 1) ' die Zelle links     von c
'    Dim cu As Range:     Set cu = ws.Cells(c.Row + 1, c.Column) ' die Zelle unterhalb von c
'    Dim cr As Range:     Set cr = ws.Cells(c.Row, c.Column + 1) ' die Zelle rechts    von c
'
'    '               =WENN(I14=K14;I15; I15+(K15-I15)/(K14-I14)*(J14-I14))
'    co.FormulaR1C1 = "=IF()"
    
    Debug.Print BilIPol(0.5, 10, 0.6, 20, 0.9, 70, 0.91, 0.875, 0.97, 0.96)
    Debug.Print BilIPol2(0.5, 0.6, 0.9, 10, 20, 70, 0.91, 0.875, 0.97, 0.96)
    
End Sub

'     |  x1  |  x  |  x2
'-------------------------
'  y1 | f11  |  ?  | f12
'  y  |  ?   | ??? |  ?
'  y2 | f12  |  ?  | f22
'------------------------

Function BilIPol(ByVal x1 As Double, ByVal y1 As Double, ByVal x As Double, ByVal y As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal f11 As Double, ByVal f21 As Double, ByVal f12 As Double, ByVal f22 As Double) As Double
    Dim R1 As Double: R1 = (x2 - x) / (x2 - x1) * f11 + (x - x1) / (x2 - x1) * f21
    Dim R2 As Double: R2 = (x2 - x) / (x2 - x1) * f12 + (x - x1) / (x2 - x1) * f22
    BilIPol = (y2 - y) / (y2 - y1) * R1 + (y - y1) / (y2 - y1) * R2
End Function

'     |  x1  |  x   |  x2
'-------------------------
' y1  | f11     ?     f21
' y   |  ?     ???     ?
' y2  | f12     ?     f22
'-------------------------

Function BilIPol2(ByVal x1 As Double, ByVal x As Double, ByVal x2 As Double, ByVal y1 As Double, ByVal y As Double, ByVal y2 As Double, ByVal f11 As Double, ByVal f21 As Double, ByVal f12 As Double, ByVal f22 As Double) As Double
    'https://de.wikipedia.org/wiki/Bilineare_Filterung
    Dim x2Minx1 As Double: x2Minx1 = x2 - x1
    Dim x2MinxDivx2Minx1 As Double: x2MinxDivx2Minx1 = (x2 - x) / x2Minx1
    Dim xMinx1Divx2Minx1 As Double: xMinx1Divx2Minx1 = (x - x1) / x2Minx1
    Dim R1 As Double: R1 = x2MinxDivx2Minx1 * f11 + xMinx1Divx2Minx1 * f21
    Dim R2 As Double: R2 = x2MinxDivx2Minx1 * f12 + xMinx1Divx2Minx1 * f22
    BilIPol2 = (y2 - y) / (y2 - y1) * R1 + (y - y1) / (y2 - y1) * R2
End Function

Sub SwapAdjacentCells()
    
    Dim r As Range:         Set r = Excel.Selection
    Dim i As Long, ui As Long: ui = r.Rows.Count
    Dim j As Long, uj As Long: uj = r.Columns.Count
    
    If uj < 2 Then
        MsgBox "Select 2 columns"
        Exit Sub
    End If
    If uj > 2 Then
        MsgBox "Select maximum 2 columns"
        Exit Sub
    End If
    
    Dim c1 As Range, c2 As Range
    For i = 1 To ui
        Set c1 = r.Cells(i, 1)
        Set c2 = r.Cells(i, 2)
        SwapCells c1, c2
    Next
End Sub

Sub SwapCells(c1 As Range, c2 As Range)
    
    Dim v: v = c1.Value
    Dim F: F = c1.Formula
    
    c1.Value = c2.Value
    c1.Formula = c2.Formula
    
    c2.Value = v
    c2.Formula = F
    
End Sub

