VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UndoRedoCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type CRange
    RngCell As Excel.Range
    SerHTML As String
End Type
Private Type StackOfCRange
    Arr() As CRange
    Count As Long
End Type

Private m_Undo As StackOfCRange
Private m_Redo As StackOfCRange
Private Const CF_HTML             As Long = &HC108&   ' 49416 HTML Format
'

Private Sub Class_Initialize()
    Call Clear
End Sub
Public Sub Clear()
    Call Stack_Create(m_Undo)
    Call Stack_Create(m_Redo)
End Sub

Public Sub Save(aCell As Excel.Range)
    Call Stack_Push(m_Undo, New_CRange(aCell, SerialiseCell(aCell)))
    Call Stack_Clear(m_Redo)
End Sub

Private Function SerialiseCell(aCell As Range) As String
    'den Bereich in die Zwischenablage kopieren dabei wird eine HTML-Serialisierung durchgeführt
    Dim cb As New cClipBoard
    aCell.Copy
    Dim s As String
    If cb.HasFormat(ClipboardFormat.CF_HTML_xls1) Then
        s = cb.StrData(ClipboardFormat.CF_HTML_xls1)
    ElseIf cb.HasFormat(ClipboardFormat.CF_HTML_xls2) Then
        s = cb.StrData(ClipboardFormat.CF_HTML_xls2)
    Else
        Dim cbf As ClipboardFormat
        cbf = cb.GetCBFormatForName("HTML")
        If cb.HasFormat(cbf) Then
            s = cb.StrData(cbf)
        End If
    End If
    SerialiseCell = ExtractHTML(s)
End Function

Private Function ExtractHTML(ByVal scbHTML As String) As String
    Dim s As String
    Dim pos As Long
    Dim StartHTML As Long, EndHTML   As Long
    If Len(scbHTML) = 0 Then Exit Function
    pos = InStr(1, scbHTML, "StartHTML:", vbTextCompare)
    If pos > 0 Then
        s = Mid(scbHTML, pos + 10, 10)
        StartHTML = CLng(s)
        pos = InStr(1, scbHTML, "EndHTML:", vbTextCompare)
        If pos > 0 Then
            s = Mid(scbHTML, pos + 8, 10)
            EndHTML = CLng(s)
            ExtractHTML = Trim(Mid(scbHTML, StartHTML + 2, EndHTML - StartHTML - 2))
        End If
    End If
End Function

Public Property Get CanUndo() As Boolean
    'über CanUndo und CanRedo die Icons de-/aktivieren (nicht hier)
    CanUndo = 0 < m_Undo.Count
End Property
Public Property Get CanRedo() As Boolean
    CanRedo = 0 < m_Redo.Count
End Property

Public Sub Undo()
    'Zelle komplett deserialisieren
    Dim cr As CRange: cr = Stack_Pop(m_Undo)
    Call Stack_Push(m_Redo, cr)
    Paste cr
End Sub

Public Sub Redo()
    'Zelle komplett deserialisieren
    Dim cr As CRange: cr = Stack_Pop(m_Redo)
    Call Stack_Push(m_Undo, cr)
    Paste cr
End Sub

Private Sub Paste(cr As CRange)
    With cr
        Dim cb As New cClipBoard
        cb.StrData(CF_UNICODETEXT) = .SerHTML
        .RngCell.PasteSpecial
    End With
End Sub

Private Function New_CRange(aCell As Excel.Range, s As String) As CRange
    With New_CRange
        Set .RngCell = aCell
        .SerHTML = s
    End With
End Function

Private Sub Stack_Create(this As StackOfCRange)
    this.Count = 0
End Sub
Private Sub Stack_Clear(this As StackOfCRange)
    With this
        Erase .Arr
        .Count = 0
    End With
End Sub

Private Function Stack_Peek(this As StackOfCRange) As CRange
    'das aktuelle (letzte) Element auf dem Stack zurückgeben
    With this
        If .Count > 0 Then Stack_Peek = .Arr(.Count)
    End With
End Function

Private Sub Stack_Push(this As StackOfCRange, Value As CRange)
    'ein neues Element auf den Stack legen
    With this
        If .Count <= 0 Then
            ReDim .Arr(0)
        Else
            Dim u As Long: u = UBound(.Arr)
            If u < .Count Then
                ReDim Preserve .Arr(0 To .Count)
            End If
        End If
        .Arr(.Count) = Value
        .Count = .Count + 1
    End With
End Sub
Private Function Stack_Pop(this As StackOfCRange) As CRange
    'das letzte Element zurückgeben und den Index um 1 zurücksetzen
    With this
        Stack_Pop = .Arr(.Count - 1)
        .Count = .Count - 1
    End With
End Function


'Function Stack_UBound(this As StackOfCRange) As Long
'    Stack_UBound = UBound(this.Arr)
'End Function
'Function Stack_Count(this As StackOfCRange) As Long
'    Stack_Count = Stack_UBound(this) + 1
'End Function
'Function Stack_Item(ByVal Index As Long) As String
'    Stack_Item = m_Arr(Index)
'End Function
'Private Sub Stack_Clear(this As StackOfCRange)
'    empty this.Arr
'    With this
'        ReDim .Arr(-1 To -1)
'    End With
'End Sub
'Function Stack_IsEmpty() As Boolean
'    Stack_IsEmpty = Stack_UBound < 0
'End Function
'Function Stack_Add(Value As String)
'    Dim u As Long: u = MyList_UBound + 1
'    If Stack_IsEmpty Then ReDim m_Arr(0 To u) Else ReDim Preserve m_Arr(0 To u)
'    m_Arr(u) = Value
'End Function

