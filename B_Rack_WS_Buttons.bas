Attribute VB_Name = "B_Rack_WS_Buttons"
Sub Negative_QC()

    If ActiveSheet.Negative.Value = True Then
        Range("D6").Interior.Color = RGB(0, 255, 0)
        ActiveSheet.Negative.BackColor = RGB(0, 255, 0)
    ElseIf ActiveSheet.Negative.Value = False Then
        Range("D6").Interior.Color = RGB(255, 0, 1)
        ActiveSheet.Negative.BackColor = RGB(255, 0, 1)
    End If
End Sub
Sub Positive_QC()

    If ActiveSheet.Positive.Value = True Then
        Range("C6").Interior.Color = RGB(0, 255, 0)
        ActiveSheet.Positive.BackColor = RGB(0, 255, 0)
    ElseIf ActiveSheet.Positive.Value = False Then
        Range("C6").Interior.Color = RGB(255, 0, 2)
        ActiveSheet.Positive.BackColor = RGB(255, 0, 2)
    End If
End Sub
Sub Positive_Result()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(255, 0, 0)
Selection.Font.Color = RGB(255, 255, 255)
End Sub
Sub Positive_Cluster()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(179, 179, 179)
Selection.Font.Color = RGB(0, 0, 0)
End Sub
Sub N_Pos()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(221, 221, 255)
Selection.Font.Color = RGB(0, 0, 0)
End Sub
Sub S_Pos()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(255, 219, 167)
Selection.Font.Color = RGB(0, 0, 0)
End Sub
Sub ORF_Pos()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(255, 217, 236)
Selection.Font.Color = RGB(0, 0, 0)
End Sub
Sub MS2()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(204, 255, 255)
Selection.Font.Color = RGB(0, 0, 0)
End Sub
Sub Analytical_Recheck()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(255, 255, 102)
Selection.Font.Color = RGB(0, 0, 0)
End Sub
Sub Rerack()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Interior.Color = RGB(51, 204, 255)
Selection.Font.Color = RGB(255, 255, 255)
End Sub
Sub Reject_Result()
If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub

Dim destRng As Range, cRng As Range, lCell As String, cRngS As String
Set cRng = Selection.Parent.Cells(Selection.Row, "B")
cRngS = Right(Selection.Parent.Cells(Selection.Row, "B").Value, 1)
lCell = Left(Selection.Parent.Cells(5, Selection.Column).Value, 2)
Set destRng = ActiveSheet.Cells(Rows.Count, "L").End(xlUp).Offset(1, 0)

A = InputBox("List of Rejections:" & vbNewLine & vbNewLine & "1. Quantity Not Sufficient (QNS)" & vbNewLine & "2. Contaminated Specimen (CS)" & vbNewLine & "3. Mismatched Specimen (MS)" & vbNewLine & "4. Missing Specimen Swab (MSS)" & vbNewLine & "5. Specimen Too Old (STO)" & vbNewLine & "6. Unapproved Media Type (UMT)" & vbNewLine & "7. Unapproved Specimen Type (UST)" & vbNewLine & "8. Unlabeled Specimen (US)" & vbNewLine & "9. Dry Swab (DS)" & vbNewLine & vbNewLine & "Enter number 1-9 to record rejection." & vbNewLine)
If A = 0 Or A = "" Or A > 9 Then Exit Sub
        With destRng
            .Font.Size = 14
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
        End With
If A = 1 Then
    destRng.Value = cRngS & lCell & " - " & "Quantity Not Sufficient (QNS)"
ElseIf A = 2 Then
    destRng.Value = cRngS & lCell & " - " & "Contaminated Specimen (CS)"
ElseIf A = 3 Then
    destRng.Value = cRngS & lCell & " - " & "Mismatched Specimen (MS)"
ElseIf A = 4 Then
    destRng.Value = cRngS & lCell & " - " & "Missing Specimen Swab (MSS)"
ElseIf A = 5 Then
    destRng.Value = cRngS & lCell & " - " & "Specimen too old (STO)"
ElseIf A = 6 Then
    destRng.Value = cRngS & lCell & " - " & "Unapproved Media Type (UMT)"
ElseIf A = 7 Then
    destRng.Value = cRngS & lCell & " - " & "Unapproved Specimen Type (UST)"
ElseIf A = 8 Then
    destRng.Value = cRngS & lCell & " - " & "Unlabeled Specimen (US)"
ElseIf A = 9 Then
    destRng.Value = cRngS & lCell & " - " & "Dry Swab (DS)"
End If
Selection.Interior.Color = RGB(0, 0, 2)
Selection.Font.Color = RGB(255, 255, 255)
End Sub
Sub No_Fill_Result()
If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub

    If Selection.Interior.Color = RGB(0, 0, 2) Then
        Call OptimizeCode_Begin
        
        Dim fRng As Range, sRng As Range, cRng As Range, cRngS As String
        Dim lCell As String, c
        Set sRng = Cells(16, 12)
        Set fRng = Cells(Rows.Count, "L").End(xlUp)
        Set cRng = Selection.Parent.Cells(Selection.Row, "B")
        cRngS = Right(Selection.Parent.Cells(Selection.Row, "B").Value, 1)
        lCell = Left(Selection.Parent.Cells(5, Selection.Column).Value, 2)
        c = cRngS & lCell
        Range(sRng, fRng).Find(c).Delete
        
            Call OptimizeCode_End
    End If

Selection.Interior.Color = xlNone
Selection.Font.Color = RGB(0, 0, 0)
End Sub
Sub Add_RR_Border()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Borders.Weight = xlThick
Selection.Borders.Color = RGB(0, 0, 192)
End Sub
Sub Remove_RR_Border()

If (Selection.Column > 14) Or (Selection.Row > 13) Or (Selection.Row < 5) Then Exit Sub
Selection.Borders.LineStyle = xlThin
Selection.Borders.Color = RGB(0, 0, 0)
End Sub
Function CountCellsByColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Interior.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByColor = cntRes
End Function
