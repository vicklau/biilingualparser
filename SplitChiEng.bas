Attribute VB_Name = "SplitChiEng"
'=========== To split Chinese and English texts (chiefly addresses) into two adjacent cells
'========== September 2017  (c) Victor Lau   Licence: GNU GPL v3.0

Option Explicit

Public Sub SplitChiEng()
Dim TargetArea() As Variant, TargetRange As Range, TargetChar As String * 1
Dim EngPasteRange As Range, ChiPasteRange As Range
Dim AreaCount As Long, RangeCount As Long, CellCount As Long, CharCount As Long
Dim CheckDigit As String * 1, CheckDigit3 As String * 4
Dim EngContent As String, ChiContent As String
    
If Selection Is Nothing Then
    Set TargetRange = Application.InputBox("Select the content to be parsed.", "Bilingual Parse", ActiveCell, Type:=8)
End If

AreaCount = Selection.Areas.Count
Dim n: For n = 1 To AreaCount
    ReDim TargetArea(1 To n)
    Set TargetArea(n) = Selection.Areas(n)
    RangeCount = TargetArea(n).Areas.Count
    
Dim m: For m = 1 To RangeCount
    Set TargetRange = TargetArea(n).Cells(m)
    With TargetRange
        CellCount = .Columns.Count * .Rows.Count
        If Len(.Offset(0, 1).Text) > 0 Then
            If MsgBox("There is data in the cells to be replaced.", vbOKCancel, "Alert") = vbCancel Then Stop: Exit Sub
        End If
Dim k: For k = 1 To CellCount
        CharCount = .Cells(k).Characters.Count
        CheckDigit3 = "---"
        Set EngPasteRange = .Cells(k): Set ChiPasteRange = EngPasteRange.Offset(0, 1)
        
Dim l: For l = 1 To CharCount
        TargetChar = Mid$(EngPasteRange.Text, l, 1)
        Select Case WorksheetFunction.Unicode(TargetChar)
            'English
            Case 38, 40 - 41, 44 - 47, 64 - 90, 97 - 122, 224 - 253
                EngContent = EngContent & TargetChar
                CheckDigit = "E"
            'Chinese
            Case 11904 - 12245, 13312 - 19893, 19968 - 40959
                ChiContent = ChiContent & TargetChar
                CheckDigit = "C"
            'Numerals and Punctuation
            Case 32, 48 - 57, 9312 - 9371, 10102 - 10131
                CheckDigit = "N"
                Select Case True
                    Case CheckDigit3 Like "??E", CheckDigit3 Like "--N", CheckDigit3 Like "-NN", CheckDigit3 Like "---"
                        EngContent = EngContent & TargetChar
                    Case CheckDigit3 Like "??C", CheckDigit3 Like "?CN", CheckDigit3 Like "CNN"
                        ChiContent = ChiContent & TargetChar
                    Case Else
                End Select
            Case Else
                Select Case MsgBox("Excel cannot identify this: " & TargetChar & " in '" & TargetRange.Cells(k).Text & "'. If it is English, click YES; Chinese, click NO.", _
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
        Next l
        EngPasteRange.Value = EngContent
        ChiPasteRange.Value = ChiContent
        Next k
End With
        Next m: Next n
End Sub

'========
'REALIGN : Rearrange non-empty cells into one flat column
'============

Public Sub Realign()
Dim rScope As Range, rCell As Range, cellOutput As Range, n As Long
Set rScope = Selection
Set cellOutput = Range(InputBox("Where do you want to place the output?", , rScope.address))
n = 0
For Each rCell In rScope
    If Not IsEmpty(rCell.Text) Then
        cellOutput.Offset(n, 0).Value2 = rCell.Text
        n = n + 1
    End If
Next rCell
End Sub

