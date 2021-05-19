Attribute VB_Name = "Module1"
Sub Call_Result_Information()

Call OptimizeCode_Begin

Dim FileToOpen As Variant, xRet As Boolean, Name As String
Dim QSResultFileWB As Workbook, QSResultFileWS As Worksheet, FormattingWS As Worksheet
Dim SampleNameLastRow As Long, TargetNameLastRow As Long, CrtLastRow As Long, crtsdLastRow As Long, dLastRow As Long, FormattingWBCrtLastRow As Long  'find last row of significant columns
Dim SampleName As Range, TargetName As Range, crt As Range, crtSD As Range            'find column headers to extract data
Dim SampleID As Range, TargetID As Range, CrtCell As Range, CrtSDCell As Range
Dim SampleIDDest As Range, TargetIDDest As Range, crtDest As Range, crtSDDest As Range  'column data to transfer
Dim CrtAverage As Range, CrtThresholdCutoff As Integer, FinalResult As Range, FirstTarget As Range, SecondTarget As Range
Dim FirstTargetCrtValue As Range, SecondTargetCrtValue As Range, CrtSDValue As Range


Dim myPath As String
    myPath = "X:\Jacob\Macros\Open Array\UTM Open Array\UTM Format For Ligo Exports\OpenArray Validation Files"
    ChDir myPath
    

CrtThresholdCutoff = 38

 FileToOpen = Application.GetOpenFilename(Title:="", FileFilter:="Excel Files (*.xls*),*xls*")       'if file types change to csv or something else, this needs changed
                            
                            Dim StartTime As Double
                            Dim MinutesElapsed As String
                            
                            'Remember time when macro starts
                              StartTime = Timer
  
        If FileToOpen <> False Then
        
        Name = CStr(FileToOpen)
        xRet = IsWorkBookOpenNow(Name)
        
        Set QSResultFileWB = Application.Workbooks.Open(FileToOpen)
        Set QSResultFileWS = QSResultFileWB.Sheets(1)
        Set FormattingWS = ThisWorkbook.Sheets(1)
        
        With QSResultFileWS
            Set SampleName = .Range("A1:Q50").Find("Sample Name")           'finds sample name column header
            Set TargetName = .Range(.Cells(SampleName.Row, 1), .Cells(SampleName.Row, .Columns.Count).End(xlToLeft)).Find("Target Name")        'finds target name column header
            Set crt = .Range(.Cells(SampleName.Row, 1), .Cells(SampleName.Row, .Columns.Count).End(xlToLeft)).Find("Crt")       'finds crt column header
            Set crtSD = .Range(.Cells(SampleName.Row, 1), .Cells(SampleName.Row, .Columns.Count).End(xlToLeft)).Find("Crt SD")  'finds crt SD column header
            SampleNameLastRow = .Cells(Rows.Count, SampleName.Column).End(xlUp).Row
            TargetNameLastRow = .Cells(Rows.Count, TargetName.Column).End(xlUp).Row
            CrtLastRow = .Cells(Rows.Count, crt.Column).End(xlUp).Row
            crtsdLastRow = .Cells(Rows.Count, crtSD.Column).End(xlUp).Row
        End With
        
        Set SampleIDDest = FormattingWS.Range("D10")    'setting data destinantion for sample name
        Set TargetIDDest = FormattingWS.Range("E10")    'setting data destination for target name
        Set crtDest = FormattingWS.Range("M10")         'setting data destination for crt values
        Set crtSDDest = FormattingWS.Range("G10")       'setting data destination for crtSD values
        
        FormattingWS.Range("F10").value = "Crt Average"
        FormattingWS.Range("H10").value = "Final Result"
        
            'import Sample Name Column
        For Each SampleID In QSResultFileWS.Range(QSResultFileWS.Cells(SampleName.Row, SampleName.Column), QSResultFileWS.Cells(SampleNameLastRow, SampleName.Column)).Cells
            If SampleID.value = "" Then
                SampleIDDest.value = "Blank"
            Else
                SampleIDDest.value = SampleID.value
            End If
            
            Set SampleIDDest = SampleIDDest.Offset(1, 0)
        Next SampleID
        
            'import Target Name Column
         For Each TargetID In QSResultFileWS.Range(QSResultFileWS.Cells(TargetName.Row, TargetName.Column), QSResultFileWS.Cells(TargetNameLastRow, TargetName.Column)).Cells
            TargetIDDest.value = TargetID.value
         
            Set TargetIDDest = TargetIDDest.Offset(1, 0)
         Next TargetID
         
            'import Crt Value Column
         For Each CrtCell In QSResultFileWS.Range(QSResultFileWS.Cells(crt.Row, crt.Column), QSResultFileWS.Cells(CrtLastRow, crt.Column)).Cells
            crtDest.value = CrtCell.value
            With crtDest
                .NumberFormat = "0.00"  'round to 2 decimal places
            End With
            Set crtDest = crtDest.Offset(1, 0)
         Next CrtCell
                
            'import CrtSD Column
         For Each CrtSDCell In QSResultFileWS.Range(QSResultFileWS.Cells(crtSD.Row, crtSD.Column), QSResultFileWS.Cells(crtsdLastRow, crtSD.Column)).Cells
            crtSDDest.value = CrtSDCell.value
                With crtSDDest
                    .NumberFormat = "0.00"  'round to 2 decimal places
                End With
            Set crtSDDest = crtSDDest.Offset(1, 0)
         Next CrtSDCell
        
        With FormattingWS
            dLastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
            FormattingWBCrtLastRow = .Cells(.Rows.Count, "M").End(xlUp).Row
        End With

'Dim i As Integer, mergeC As Range
        'merge Crt Average cells
    'For i = 11 To FormattingWBCrtLastRow Step 2
    'Set mergeC = FormattingWS.Cells(i, 4)
    '    FormattingWS.Range(mergeC, mergeC.Offset(1, 0)).Merge
    'Next i

        For Each CrtAverage In FormattingWS.Range("F11:F" & FormattingWBCrtLastRow).Cells       'for every cell in Crt Average Column
           
           Set FirstTarget = CrtAverage.Offset(0, -1)
           Set SecondTarget = CrtAverage.Offset(1, -1)
           Set FirstTargetCrtValue = CrtAverage.Offset(0, 7)
           Set SecondTargetCrtValue = CrtAverage.Offset(1, 7)
           Set FinalResult = CrtAverage.Offset(0, 2)
           Set CrtSDValue = CrtAverage.Offset(0, 1)
           
            If FirstTarget.value = SecondTarget Then            'check 2 columns to the left, if this target and the target directly below are the same then
                If FirstTargetCrtValue.value = "Undetermined" And SecondTargetCrtValue.value = "Undetermined" Then         'if both Crt value = "Undetermined" then
                    With CrtAverage
                        .value = ""                     'leave Crt average column blank if both crt values = Undetermined
                    End With
                    With FinalResult
                        .value = "Not Detected"         'final result = Not Detected
                    End With
                    GoTo NextIteration
                    
                ElseIf IsNumeric(FirstTargetCrtValue.value) = True And IsNumeric(SecondTargetCrtValue.value) = True Then        'if both target crt values are numbers then
                    With CrtAverage
                        .value = Application.Average(Range(FirstTargetCrtValue, SecondTargetCrtValue))                      'find the crt average
                        .NumberFormat = "0.00"
                    End With
                        If (CrtAverage.value - CrtSDValue) < CrtThresholdCutoff Then                'if (crt average - crt std dev) < crt cutoff then
                            With FinalResult
                                .value = "Detected"                                                 'final result = detected
                                .Interior.Color = RGB(0, 255, 0)
                            End With
                        ElseIf (CrtAverage.value - CrtSDValue) > CrtThresholdCutoff Then            'if (crt average - crt sd dev) > crt cutoff then
                            With FinalResult
                                .value = "Inconclusive"                                        'final result = inconclusive
                                .Interior.Color = RGB(255, 255, 0)
                            End With
                        End If
                ElseIf IsNumeric(FirstTargetCrtValue.value) = True And IsNumeric(SecondTargetCrtValue) = False Then     '1st = true 2nd = false
                    With CrtAverage
                        .value = FirstTargetCrtValue.value
                        .NumberFormat = "0.00"
                    End With
                    With FinalResult
                        .value = "Inconclusive"
                        .Interior.Color = RGB(255, 255, 0)
                    End With
                ElseIf IsNumeric(FirstTargetCrtValue.value) = False And IsNumeric(SecondTargetCrtValue.value) = True Then   '1st = false 2nd = true
                    With CrtAverage
                        .value = SecondTargetCrtValue.value
                        .NumberFormat = "0.00"
                    End With
                    With FinalResult
                        .value = "Inconclusive"
                        .Interior.Color = RGB(255, 255, 0)
                    End With
                End If
            Else: GoTo NextIteration                            'if the target directly below does not match the target above, then go to the next cell in Crt Average column
            End If
        
NextIteration:         Next CrtAverage
       
        With FormattingWS.Range("A10:F10")
            .HorizontalAlignment = xlHAlignCenter
            .Font.Size = 14
            .Font.Bold = True
        End With
       
        With FormattingWS.Range("A10:F" & SampleNameLastRow)
            .HorizontalAlignment = xlHAlignCenter
            .Columns.AutoFit
        End With
        
        
            If xRet <> True Then
                QSResultFileWB.Close False
            End If
            Else: Exit Sub
        End If
        
        
Call OptimizeCode_End
        
       MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

'Notify user in seconds
  MsgBox "This ran fast as shit! " & MinutesElapsed & " minutes", vbInformation
   
End Sub


Sub Create_Ligo_Exports_File()

'Application.Union(Range("A" & i), Range("B" & i), Range("D" & i), Range("E" & i), Range("F" & i))
Call OptimizeCode_Begin
                            Dim StartTime As Double
                            Dim MinutesElapsed As String
                            
                            'Remember time when macro starts
                              StartTime = Timer

Dim ws As Worksheet
Dim LigoExportsWS As Worksheet
Dim c As Range, RangeValue As Range, i As Integer
Dim LigoRanges As Range
Dim LigoExportsDest As Range, SameRowResult As Range, WSLastRow As Long, LigoExportsLastRow As Long

Set ws = ThisWorkbook.Sheets(1)
Set LigoExportsWS = ThisWorkbook.Sheets("LigoExport")

LigoExportsWS.Range("A1:E7000").Clear

WSLastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
LigoExportsLastRow = LigoExportsWS.Cells(Rows.Count, "A").End(xlUp).Row
Set LigoExportsDest = LigoExportsWS.Cells(Rows.Count, "A").End(xlUp).Offset(1, 0)               'Range("A1:A" & LigoExportsLastRow).Cells


With LigoExportsWS
    .Range("A1").value = ws.Range("A10").value
    .Range("B1").value = ws.Range("B10").value
    .Range("C1").value = ws.Range("D10").value
    .Range("D1").value = ws.Range("E10").value
    .Range("E1").value = ws.Range("F10").value
End With

For i = 11 To WSLastRow Step 2

    Set LigoRanges = Application.Union(ws.Range("A" & i), ws.Range("B" & i), ws.Range("D" & i), ws.Range("E" & i), ws.Range("F" & i))
   
    For Each RangeValue In LigoRanges
        With LigoExportsDest
            .value = RangeValue.value
            .NumberFormat = "0.00"
        End With
        If RangeValue.Interior.Color = RGB(255, 255, 255) Then
            
            Else
            With LigoExportsDest
                .Interior.Color = RangeValue.Interior.Color
            End With
        End If
        Set LigoExportsDest = LigoExportsDest.Offset(0, 1)
    Next RangeValue
    Set LigoExportsDest = LigoExportsDest.Offset(1, -5)
Next i

With LigoExportsWS.Range("A:E")
    .HorizontalAlignment = xlHAlignCenter
    .Columns.AutoFit
End With


'add code to automatically filter by inconclusives then detected then not detected  <--------------------------------



Call OptimizeCode_End

 MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

'Notify user in seconds
  MsgBox "This ran fast as shit! " & MinutesElapsed & " minutes", vbInformation
End Sub

