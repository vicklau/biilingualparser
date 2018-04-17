Attribute VB_Name = "SplitChiEng"
'=========== To split Chinese and English texts (chiefly addresses) into two adjacent cells
'========== September 2017  (c) Victor Lau   Licence: GNU GPL v3.0

Option Explicit

Public Sub SplitChiEng()
Dim TargetArea() As Variant, TargetCell As Range, TargetChar As String * 1
Dim EngPasteRange As Range, ChiPasteRange As Range
Dim AreaCount As Long, CellCount As Long, CharCount As Long
Dim CheckDigit As String * 1, CheckDigit3 As String * 3
Dim EngContent As String, ChiContent As String
Dim n         As Integer
If Selection Is Nothing Then
    Set Selection = Application.InputBox("Select the content to be parsed.", "Bilingual Parse", ActiveCell, Type:=8)
    AreaCount = 1
Else:
AreaCount = Selection.Areas.Count
End If
For n = 1 To AreaCount
ReDim TargetArea(1 To AreaCount) As Variant
Set TargetArea(n) = Selection.Areas(n)

CellCount = TargetArea(n).Cells.Count

Dim m: For m = 1 To CellCount
    Set TargetCell = TargetArea(n).Cells(m)
    Debug.Print TargetCell.address
    With TargetCell
        If Len(.Offset(0, 1).Text) > 0 Then
            If MsgBox("There is data in the cells to be replaced.", vbOKCancel, "Alert") = vbCancel Then Stop: Exit Sub
        End If
        CharCount = .Cells(m).Characters.Count
        CheckDigit3 = "---"
        Set EngPasteRange = .EntireRow.Cells(TargetArea(n).Columns.Count): Set ChiPasteRange = EngPasteRange.Offset(0, 1)
        EngContent = "": ChiContent = ""
        Dim l: For l = 1 To CharCount
            TargetChar = Mid$(EngPasteRange.Text, l, 1)
            Debug.Print TargetChar & ", " & WorksheetFunction.Unicode(TargetChar) & ", " & CheckDigit3
            Select Case WorksheetFunction.Unicode(TargetChar)
                    'English
                Case 38, 40 To 41, 44 To 47, 64 To 90, 97 To 122, 224 To 253
                    EngContent = EngContent & TargetChar
                    CheckDigit = "E"
                    'Chinese
                Case 11904 To 12245, 13312 To 19893, 19968 To 69999
                    ChiContent = ChiContent & TargetChar
                    CheckDigit = "C"
                    'Numerals and Punctuation
                Case 32, 48 To 57, 9312 To 9371, 10102 To 10131
                    CheckDigit = "N"
                    Select Case True
                        Case CheckDigit3 Like "??E", CheckDigit3 Like "--N", CheckDigit3 Like "-NN", CheckDigit3 Like "---"
                            EngContent = EngContent & TargetChar
                        Case CheckDigit3 Like "??C", CheckDigit3 Like "?CN", CheckDigit3 Like "CNN"
                            ChiContent = ChiContent & TargetChar
                        Case Else
                    End Select
                Case Else
                    Select Case MsgBox("Excel cannot identify this: " & TargetChar & " in '" & TargetCell.address & "'. If it is English, click YES; Chinese, click NO.", _
                                       vbYesNo + vbCritical, "Error")
                        Case vbYes
                            EngContent = EngContent & TargetChar
                            CheckDigit = "E"
                        Case vbNo
                            ChiContent = ChiContent & TargetChar
                            CheckDigit = "C"
                    End Select
            End Select
            CheckDigit3 = Right$(CheckDigit3 & CheckDigit, 3)
        Wait Now + TimeValue("00:00:01"): Next l
        EngPasteRange.Value = EngContent
        ChiPasteRange.Value = ChiContent
    End With
EngContent = "":  ChiContent = ""
Next m: Next n
On Error Resume Next
Erase TargetArea()
Set TargetCell = Nothing
TargetChar = vbNull
End Sub
