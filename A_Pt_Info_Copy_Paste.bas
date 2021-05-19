Attribute VB_Name = "A_Pt_Info_Copy_Paste"
Sub OptimizeCode_Begin()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
End Sub
Sub OptimizeCode_End()

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub ColToMatrix(rngCol As Range, rngTL As Range)

    Dim arr, mtx(1 To 8, 1 To 12), i As Long, r As Long, c As Long, n As Long
    arr = rngCol.Value
    For i = 1 To UBound(arr, 1)
        n = i + 2 'account for starting 3 cells in
        r = 1 + ((n - 1) \ 12)
        c = ((n - 1) Mod 12) + 1
        mtx(r, c) = arr(i, 1)
    Next i
    rngTL.Resize(8, 12).Value = mtx
End Sub
Sub ImportPT(CalledFrom As String)

Call OptimizeCode_Begin

    Dim FileToOpen As Variant
    Dim xRet As Boolean
    Dim Name As String
    Dim ligoWB As Workbook, ligoWS As Worksheet, rrmWB As Worksheet
    Dim rackPos As Range, rrmDest As Range, rrmRackPos As Range, findSpace As Range
    Dim arrLigoMatch, arrRRMmatch, arrLoc
    Dim ligoLastRow As Long, rrmLastRow As Long
    Dim myPath As String
    myPath = "X:\Reditus\ligo_exports"
    ChDrive myPath
    ChDir myPath
   
    arrLoc = Array("    A-3", "    A-4", "    A-5", "    A-6", "    A-7", "    A-8", "    A-9", "    A-10", "    A-11", "    A-12", _
                   "    B-1", "    B-2", "    B-3", "    B-4", "    B-5", "    B-6", "    B-7", "    B-8", "    B-9", "    B-10", "    B-11", "    B-12", _
                   "    C-1", "    C-2", "    C-3", "    C-4", "    C-5", "    C-6", "    C-7", "    C-8", "    C-9", "    C-10", "    C-11", "    C-12", _
                   "    D-1", "    D-2", "    D-3", "    D-4", "    D-5", "    D-6", "    D-7", "    D-8", "    D-9", "    D-10", "    D-11", "    D-12", _
                   "    E-1", "    E-2", "    E-3", "    E-4", "    E-5", "    E-6", "    E-7", "    E-8", "    E-9", "    E-10", "    E-11", "    E-12", _
                   "    F-1", "    F-2", "    F-3", "    F-4", "    F-5", "    F-6", "    F-7", "    F-8", "    F-9", "    F-10", "    F-11", "    F-12", _
                   "    G-1", "    G-2", "    G-3", "    G-4", "    G-5", "    G-6", "    G-7", "    G-8", "    G-9", "    G-10", "    G-11", "    G-12", _
                   "    H-1", "    H-2", "    H-3", "    H-4", "    H-5", "    H-6", "    H-7", "    H-8", "    H-9", "    H-10", "    H-11", "    H-12")


    FileToOpen = Application.GetOpenFilename(Title:="", FileFilter:="Excel Files (*.xls*),*xls*")       'if file types change to csv or something else, this needs changed
        If FileToOpen <> False Then
   
            Name = CStr(FileToOpen)
            xRet = IsWorkBookOpenNow(Name)
           
            Set ligoWB = Application.Workbooks.Open(FileToOpen)
            Set rrmWB = ThisWorkbook.Worksheets("PATIENT INFO COPY-PASTE")
            Set ligoWS = ligoWB.Worksheets(1)
            
                With rrmWB
                    rrmLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                End With
                
            Set rrmRackPos = rrmWB.Range("A1:A" & rrmLastRow).Cells
       
            With rrmWB.Range("B4:E100")
                .WrapText = True
                .RowHeight = 15
                .ColumnWidth = 27.43
                .VerticalAlignment = xlVAlignTop
            End With
       
                If CalledFrom = "RACK 1" Then
                        With rrmWB.Range("B4:B97")
                            .ClearContents
                        End With
                    ElseIf CalledFrom = "RACK 2" Then
                        With rrmWB.Range("C4:C97")
                            .ClearContents
                        End With
                    ElseIf CalledFrom = "RACK 3" Then
                        With rrmWB.Range("D4:D97")
                            .ClearContents
                        End With
                    ElseIf CalledFrom = "RACK 4" Then
                        With rrmWB.Range("E4:E97")
                            .ClearContents
                        End With
                End If
                
                With ligoWS
                       ligoLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row                             'these files are generated from an older version of excel, must find the last row of this version
                       .Range("B:I").UnMerge
                End With
               
                Set findSpace = ligoWS.Range("A1:A50").Find(" ", , xlValues)                            'Find rack ID - search for where rack positioning begins then offset -1 row and put value on RRM
               
            If Not findSpace Is Nothing Then
                If CalledFrom = "RACK 1" Then
                    rrmWB.Range("B3").Value = findSpace.Offset(-1, 0).Value  '<----difference #1 - following buttons destination = C3, D3, E3.  The code to find the value is the same - rack ID location found, place rack id found in cell B3 of RRM
                ElseIf CalledFrom = "RACK 2" Then
                    rrmWB.Range("C3").Value = findSpace.Offset(-1, 0).Value
                ElseIf CalledFrom = "RACK 3" Then
                    rrmWB.Range("D3").Value = findSpace.Offset(-1, 0).Value
                ElseIf CalledFrom = "RACK 4" Then
                    rrmWB.Range("E3").Value = findSpace.Offset(-1, 0).Value
                End If
            End If
            
            If findSpace Is Nothing Then
                If (CalledFrom = "RACK 1" Or CalledFrom = "RACK 2" Or CalledFrom = "RACK 3" Or CalledFrom = "RACK 4") Then
                    MsgBox ("No Patients Found.")
                    Exit Sub
                End If
            End If

             
            If CalledFrom = "RACK 1" Then
                rrmWB.Range("B1").Value = ligoWS.Range("E1").End(xlDown).Value  '<-----difference #2 - following buttons destination = C1, D1, E1 - Find Date/Time rack was created - put value in B1 of RRM
            ElseIf CalledFrom = "RACK 2" Then
                rrmWB.Range("C1").Value = ligoWS.Range("E1").End(xlDown).Value
            ElseIf CalledFrom = "RACK 3" Then
                rrmWB.Range("D1").Value = ligoWS.Range("E1").End(xlDown).Value
            ElseIf CalledFrom = "RACK 4" Then
                rrmWB.Range("E1").Value = ligoWS.Range("E1").End(xlDown).Value
            End If
            
            For Each rackPos In ligoWS.Range("A1:A" & ligoLastRow).Cells                                'For each cell in ligo-exports file
                arrLigoMatch = Application.Match(rackPos.Value, arrLoc, 0)                              'Search for cell values that are stored in arrLoc
                   
                    If Not IsError(arrLigoMatch) Then
                        arrRRMmatch = Application.Match(rackPos.Value, rrmRackPos, 0)                   'if cell value is found, then find the same/matching cell value in Column A of RRM
                        If CalledFrom = "RACK 1" Then
                            Set rrmDest = rrmWB.Cells(arrRRMmatch, 2)               '<---difference #3 - following buttons destination = 3, 4, 5 - setting rrmDest as the RRM Destination - arrRRMmatch = row, 2 = column
                        ElseIf CalledFrom = "RACK 2" Then
                            Set rrmDest = rrmWB.Cells(arrRRMmatch, 3)
                        ElseIf CalledFrom = "RACK 3" Then
                            Set rrmDest = rrmWB.Cells(arrRRMmatch, 4)
                        ElseIf CalledFrom = "RACK 4" Then
                            Set rrmDest = rrmWB.Cells(arrRRMmatch, 5)
                        End If
                        rrmDest.Value = rackPos.Offset(0, 1).Value                                      'Once destination has been set, from Ligo file, offset 1 Column to get desired information and place value to rrmDest
                    End If
               
            Next rackPos
           
                If xRet <> True Then
                    ligoWB.Close False
                End If

        End If
        
Call OptimizeCode_End

End Sub
Function IsWorkBookOpenNow(FileName As String) As Boolean
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpenNow = False
    Case 70:   IsWorkBookOpenNow = True
    Case Else: Error ErrNo
    End Select
End Function
