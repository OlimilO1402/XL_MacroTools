VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function CountClipboardFormats Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GetClipboardFormatNameA Lib "user32" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDst As LongPtr, ByVal pSrc As LongPtr, ByVal bLen As Long)

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
'https://msdn.microsoft.com/de-de/library/windows/desktop/aa366574(v=vs.85).aspx
'GlobalAlloc: Allocates the specified number of bytes from the heap.
'Note:
'The global functions have greater overhead and provide fewer features than other memory management functions.
'New applications should use the heap functions unless documentation states that a global function should be used.
'For more information, see Global and Local Functions.
'
'Global and Local Functions:
'https://msdn.microsoft.com/de-de/library/windows/desktop/aa366596(v=vs.85).aspx
'For example, some Windows functions allocate memory that must be freed with LocalFree, and the global
'functions are still used with Dynamic Data Exchange (DDE), the clipboard functions, and OLE data objects.


'Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Const GMEM_MOVEABLE As Long = &H2&
Public Enum ClipboardFormat
                           ' Das Objekt in der Zwischenablage ist ein . . .
    CF_TEXT = 1                     ' Handle auf einen unformatierten Ansi-Text (UTF-8).
    CF_BITMAP = 2                   ' Handle auf eine Windows Bitmap-Grafik
    CF_METAFILEPICT = 3             ' Handle auf eine Windows Metafile-Grafik
    CF_SYLK = 4                     ' Handle auf einen Microsoft Symbolic-Link
    CF_DIF = 5                      ' "Software Arts' Data Interchange Format"
    CF_TIFF = 6                     ' Handle zu einem Tiff-Bitmap
    CF_OEMTEXT = 7                  ' Handle zu einem OEM-String
    CF_DIB = 8                      ' Handle zu einer %BITMAPINFO%-Struktur
    CF_BOTTOMUP_DIB = CF_DIB        '
    CF_PALETTE = 9                  ' Handle zu einer Palette
    CF_PENDATA = 10                 ' sind Daten zu einem Microsoft Pen Extensions
    CF_RIFF = 11                    ' Handle zu einer Audiodatei
    CF_WAVE = 12                    ' Handle zu Wavedatei
    CF_UNICODETEXT = 13             ' Handle zu einem Unicode-String (UTF-16)
    CF_ENHMETAFILE = 14             ' Handle zu einer Enhanced Metadatei
    CF_HDROP = 15                   ' Liste von Dateihandles im Zusammenhang mit Drag-And-Drop
    CF_LOCALE = 16                  ' Sprach-ID, die für Text-Strings in der Zwischenablage benutzt wurde
    CF_DIBV5 = 17                   ' Handle zu einer %BITMAPV5HEADER%-Struktur (Win 2000/XP)
    
    CF_JPEG = 19
    CF_TOPDOWN_DIB = 20
    
    CF_MULTI_TIFF = 22
    
    CF_OWNERDISPLAY = &H80&         '   128 ' benutzerdefinierter Anzeigetyp
    CF_DSPTEXT = &H81&              '   129 ' Text, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_DSPBITMAP = &H82&            '   130 ' Bitmap, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_DSPMETAFILEPICT = &H83&      '   131 ' Zwischendatei, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_PRIVATEFIRST = &H200&        '   512 ' privates Handle
    CF_PRIVATELAST = &H2FF&         '   767 ' privates Handle
    CF_GDIOBJFIRST = &H300&         '   768 ' Dient in der Zwischenablage dem Setzen von privaten Datenformate
    CF_GDIOBJLAST = &H3FF&          '  1023 ' Dient in der Zwischenablage dem Setzen von privaten Datenformaten
    CF_FileName = &HC006&           ' 49158 ' Dateiname
    CF_FileNameW = &HC007&          ' 49159 ' Dateiname
    CF_DataObject = &HC009&         ' 49161 ' DataObject
    'von mir selber hinzugefügt, herausgefunden durch Excel-Zelle in Zwischenablage:
    CF_RTF = &HC09A&                ' 49306 ' Richt Text Format
    
    CF_HTML_xls1 = &HC108&               ' 49416 ' HTML Format
    CF_HTML_xls2 = &HC12F&               ' 49455 ' HTML Format
    'da will mich bei Microsoft wohl jemand ärgern!!
    'jetzt hat sich plötzlich die Konstante geändert schon sehr merkwürdig
    'man braucht eine Funktion die zum Text HTML Format die Konstante raussucht!
    
    CF_PICTURE = &HC20A&            ' 49674 ' Handle auf ein Objekt vom Datentyp Picture
    CF_OBJECT = &HC215&             ' 49685 ' Handle auf ein beliebiges Objekt
    CF_XML = &HC308&                ' 49928 ' XML Spreadheet
' da gibt es noch mehr Konstanten...
End Enum

' BitBlt dwRop-Konstante
Private Const SRCCOPY            As Long = &HCC0020

Private m_CBFormats() As Long
Private m_hWnd As LongPtr

Private Sub Class_Initialize()
    ReDim m_CBFormats(0)
End Sub

Public Sub Clear()
    OpenClipboard m_hWnd
    EmptyClipboard
    CloseClipboard
    ReDim m_CBFormats(0)
End Sub
Public Property Get IsClear() As Boolean
    IsClear = UBound(m_CBFormats) = 0
End Property
Public Sub ClearCBFormats()
    ReDim m_CBFormats(0)
End Sub

Public Property Get hWnd() As LongPtr
    hWnd = m_hWnd
End Property
Public Property Let hWnd(aHWnd As LongPtr)
    m_hWnd = aHWnd
End Property

Public Sub ReadClipboardFormats()
    If OpenClipboard(m_hWnd) Then
        Dim n As Long: n = CountClipboardFormats
        ReDim m_CBFormats(0 To n)
        Dim cbf As Long
        Dim i As Long
        Do
            i = i + 1
            ' Das nächste Format aus der Zwischenablage ermitteln
            cbf = EnumClipboardFormats(cbf)
            If cbf = 0 Then Exit Do
            m_CBFormats(i) = cbf
        Loop
        CloseClipboard
    End If
End Sub

Public Property Get CBFormats() As Long()
    If UBound(m_CBFormats) < 1 Then ReadClipboardFormats
    CBFormats = m_CBFormats
End Property

Public Function GetCBFormatForName(aName As String) As Long
    If UBound(m_CBFormats) < 1 Then ReadClipboardFormats
    Dim i As Long
    For i = 0 To UBound(m_CBFormats)
        Dim n1 As String: n1 = CBFormat_ToStr(m_CBFormats(i))
        Dim n2 As String: n2 = "*" & aName & "*"
        If n1 Like n2 Then
            GetCBFormatForName = m_CBFormats(i)
            Exit Function
        End If
    Next
End Function

Public Property Get HasFormat(ByVal aCBFormat As ClipboardFormat) As Boolean
    HasFormat = IsClipboardFormatAvailable(aCBFormat) '<> 0
    'Dim i As Long
    'For i = 0 To UBound(m_CBFormats)
    '    HasCBFormat = m_CBFormats(i) = aCBFormat
    '    If HasFormat Then
    '        Exit Property
    '    End If
    'Next
End Property

Public Property Get StrData(ByVal aCBFormat As Long) As String
    If OpenClipboard(m_hWnd) Then
        Dim pMem As LongPtr:         pMem = GetClipboardData(aCBFormat)
        If pMem Then
            Dim bLen As Long:        bLen = GlobalSize(pMem)
            If bLen Then
                Dim pLok As LongPtr: pLok = GlobalLock(pMem)
                If pLok Then
                    If aCBFormat = CF_LOCALE Then
                        Dim localeID As Long
                        RtlMoveMemory VarPtr(localeID), pLok, bLen
                        StrData = CStr(localeID) 'StrConv(StrData, vbUnicode)
                    Else
                        Dim slen As Long: slen = IIf(aCBFormat = CF_UNICODETEXT, (bLen \ 2) - 1, bLen)
                        StrData = Space$(slen)
                        RtlMoveMemory StrPtr(StrData), pLok, bLen
                        If aCBFormat <> CF_UNICODETEXT Then
                            StrData = StrConv(StrData, vbUnicode)
                        End If
                    End If
                    GlobalUnlock pMem
                End If
            End If
        End If
        CloseClipboard
    End If
End Property

Public Property Let StrData(ByVal aCBFormat As Long, Value As String)
    If OpenClipboard(m_hWnd) Then
        Dim bLen As Long:            bLen = LenB(Value) + 2
        If bLen > 2 Then
            Dim pMem As LongPtr:     pMem = GlobalAlloc(GMEM_MOVEABLE, bLen)
            If pMem Then
                Dim pLok As LongPtr: pLok = GlobalLock(pMem)
                If pLok Then
                    RtlMoveMemory pLok, StrPtr(Value), bLen
                    GlobalUnlock pMem
                    SetClipboardData aCBFormat, pMem
                End If
            End If
        End If
        CloseClipboard
    End If
End Property

Public Function CBFormat_ToStr(aCBFormat As Long) As String
    Dim s  As String: s = CLng(aCBFormat) & ", &H" & Hex(aCBFormat)
    Dim s2 As String: s2 = Space(256)
    Dim rv As Long: rv = GetClipboardFormatNameA(aCBFormat, s2, 256)
    s2 = Trim(s2)
    If Len(s2) Then
        s = s & "(api): " & s2
    Else
        Select Case aCBFormat
        Case 0:
        Case CF_TEXT:         s = s & ": Handle zu String"
        Case CF_BITMAP:       s = s & ": Handle eines Bitmaps"
        Case CF_METAFILEPICT: s = s & ": Handle zu einem Metafile Bild"
        Case CF_SYLK:         s = s & ": CF_SYLK"
        Case CF_DIF:          s = s & ": Software Arts' Data Interchange Format"
        Case CF_TIFF:         s = s & ": Handle zu einem Tiff-Bitmap"
        Case CF_OEMTEXT:      s = s & ": Handle zu einem OEM-String"
        Case CF_DIB:          s = s & ": Handle zu einer %BITMAPINFO%-Struktur"
        Case CF_BOTTOMUP_DIB: s = s & ": CF_DIB"
        Case CF_PALETTE:      s = s & ": Handle zu einer Palette"
        Case CF_PENDATA:      s = s & ": sind Daten zu einem Microsoft Pen Extensions"
        Case CF_RIFF:         s = s & ": Handle zu einer Audiodatei"
        Case CF_WAVE:         s = s & ": Handle zu Wavedatei"
        Case CF_UNICODETEXT:  s = s & ": Handle zu einem Unicode-String"
        Case CF_ENHMETAFILE:  s = s & ": Handle zu einer Enhanced Metadatei"
        Case CF_HDROP:        s = s & ": Liste von Dateihandles"
        Case CF_LOCALE:       s = s & ": Sprach-ID, die für Text-Strings in der Zwischenablage benutzt wurde"
        Case CF_DIBV5:        s = s & ": Handle zu einer %BITMAPV5HEADER%-Struktur (Win 2000/XP)"
    
        Case CF_JPEG:         s = s & ": CF_JPEG"
        Case CF_TOPDOWN_DIB:  s = s & ": CF_TOPDOWN_DIB"
    
        Case CF_MULTI_TIFF:   s = s & ": CF_MULTI_TIFF"
    
        Case CF_OWNERDISPLAY: s = s & ": benutzerdefinierter Anzeigetyp"
        Case CF_PRIVATEFIRST: s = s & ": CF_PRIVATEFIRST privates Handle"
        Case CF_PRIVATELAST:  s = s & ": CF_PRIVATELAST privates Handle"
        Case CF_GDIOBJFIRST:  s = s & ": CF_GDIOBJFIRST"
        Case CF_GDIOBJLAST:   s = s & ": CF_GDIOBJLAST"
    
        Case CF_DataObject:   s = s & ": CF_DataObject DataObject"
        Case CF_FileName:     s = s & ": CF_FileName   Dateiname"
        Case CF_FileNameW:    s = s & ": CF_FileNameW  Dateiname"
        Case Else:            s = s & ": unbekanntes Format"
        End Select
    End If
    CBFormat_ToStr = s
End Function

