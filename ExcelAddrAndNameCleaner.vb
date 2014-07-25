
' Clean Addresses and Person names - for Vodafone/TelstraClear Credit Review Team
' Author Hugh Van
' Email hvan021@gmail.com
' License: MIT

Private Sub Button1_Click()
    Dim usedRow As Integer
    Dim StreetNoColName As String, StreetNameColName As String, StreetTypeColName As String, StreetUnitColName As String, FirstNameColName As String, DobColName As String
    Dim TempStreetNoColName As String, TempStreetNameColName As String, TempStreetTypeColName As String
    Dim StreetUnitStat As Long, StreetNoStat As Long, StreetTypeStat As Long, FNameStat As Long
    Dim DuplicateHeaderRow As Boolean
    Dim streetno_val As String

    Dim Errcount As Long
    Dim ErrLines() As Long

    FirstNameColName = "D"
    StreetUnitColName = "F"
    StreetNoColName = "G"
    StreetNameColName = "H"
    StreetTypeColName = "I"
    DobColName = "E"

    TempStreetNoColName = "AA"
    TempStreetNameColName = "AB"
    TempStreetTypeColName = "AC"

    StreetUnitStat = 0
    StreetNoStat = 0
    StreetTypeStat = 0
    FNameStat = 0

    ErrCount = 0


    Dim i As Long
    Dim startCell As String, endCell As String
    With ActiveWorkbook.ActiveSheet

        ''' Remove duplicate header if necessary
        If InStr(LCase(Cells(2,"B").Value), "customer id") > 0 Then
            DuplicateHeaderRow = True
            .Rows(2).Delete
        End If

        For i = 2 To .UsedRange.Rows.Count Step 1
'            .Cells(i, copy_to_col).Value = .Cells(i, copy_from_col).Value
            startCell = StreetNoColName & i
            endCell = StreetTypeColName & i
            fromRange = StreetNoColName & i & ":" & StreetTypeColName & i
            toRange = TempStreetNoColName & i & ":" & TempStreetTypeColName & i
            .Range(fromRange).Copy
            .Range(toRange).Select
            .Paste

             ''' Clean first name
            If InStr(Cells(i, FirstNameColName), " ") > 0 Then
                FNameStat = FNameStat + 1
                Cells(i, FirstNameColName).Value = Left(Cells(i, FirstNameColName), InStr(Cells(i, FirstNameColName), " ") - 1)
            End If
            ''' End clean first names

            ''' Change dd.mm.yyy to dd/mm/yyy
            Cells(i, DobColName).Value = Replace(Cells(i, DobColName), ".", "/")



            FullStreetName = Cells(i, TempStreetNameColName).Value
            If NOT IsEmpty(FullStreetName) Then ''' If FullStreetName is empty then skip this row
                If InStr(LCase(FullStreetName), "floor") > 0 Then ''' If error
                    ReDim Preserve ErrLines(ErrCount)
                    ErrLines(Errcount) = i
                    ErrCount = ErrCount + 1
                Else ''' no potential error - go ahead and clean the address
                    ''' Clean Street No
                    If (IsEmpty(Cells(i, TempStreetNoColName))) Then
                        StreetNoStat = StreetNoStat + 1
                        ' .Range(StreetNoColName & i).NumberFormat = "@"

                        ' Clean "room" from cell
                        If InStr(LCase(FullStreetName), "room") > 0 Then
                            Cells(i, TempStreetNameColName).Value = Trim(Replace(FullStreetName, "Room", "", 1, -1, vbTextCompare))
                        End If

                        streetno_val = Trim(StreetNo(Cells(i, TempStreetNameColName).Value))
        '                .Cells(i, StreetNoColName).Value = StreetNo(Cells(i, TempStreetNameColName).Value)

                        If IsNumeric(streetno_val) Then
                            .Cells(i, StreetNoColName).Value = streetno_val
                        Else
                            StreetUnitStat = StreetUnitStat + 1
                            ' StreetNo contain "/" for eg. 17/19 then Unit:17, StreetNo: 19
                            If InStr(streetno_val, "/") > 0 Then
                                .Cells(i, StreetUnitColName).Value = Left(streetno_val, InStr(streetno_val, "/") - 1)
                                .Cells(i, StreetNoColName).Value = Right(streetno_val, Len(streetno_val) - InStr(streetno_val, "/"))
                            Else
                                ' StreetNo contains character eg. 103A then Unir: A, StreetNo: 103

                                For x = 1 To Len(streetno_val)

                                found_char = (Mid(streetno_val, x, 1) Like "[a-zA-Z]")
                                If found_char = True Then
        '                            ''' Attemp to alert if encounter Street No with Room number
        '                            If InStr(LCase(streetno_val), "room") > 0 Then
        '                                ''' ErrMsg = "StreetNo contains room number. Manual edit required at row " & i
        '                            Else
                                        newStreetNo = Left(streetno_val, x - 1)
                                        Unit = Right(streetno_val, Len(streetno_val) - (x - 1))

                                        .Cells(i, StreetUnitColName).Value = Unit
                                        .Cells(i, StreetNoColName).Value = newStreetNo
        '                            End If
                                End If
                                Next x
                            End If
                        End If
                    End If
                    ' End clean Street No

                    ''' Clean street type
                    If (IsEmpty(Cells(i, TempStreetTypeColName))) Then
                        StreetTypeStat = StreetTypeStat + 1
                        ' .Range(StreetTypeColName & i).NumberFormat = "@"

                        ''' Clean street type
                        .Cells(i, StreetTypeColName).Value = StreetType(Cells(i, TempStreetNameColName).Value)

                        ''' Now clean street name
                        .Cells(i, StreetNameColName).Value = StreetNameOnly(Cells(i, TempStreetNameColName).Value)
                    End If
                    ''' End clean street type

                    ''' Now clean street name
        '            If (IsEmpty(Cells(i, TempStreetTypeColName))) Then
        '                .Cells(i, StreetNameColName).Value = StreetNameOnly(Cells(i, TempStreetNameColName).Value)
        '            End If
                    ''' End clean street name
                End If


                ''' Cleer temp columns
                .Range(toRange).Select
                Selection.ClearContents
            End If ''' End if FullStreetName is NOT empty

            ''' Go back to street columns
            ' .Range(fromRange).Select
        Next i
        .Range("A1").Select
    End With

    Dim CSVFileName As String
    CSVFileName = ActiveWorkbook.Path & "\" & "Upload " & Format(DateAdd("d", 1, Date), "dd.mm.yyyy") & ".csv"
    'ExportAsCSV(CSVFileName)

    ActiveWorkbook.Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=CSVFileName, FileFormat:=xlCSV, CreateBackup:=False, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
   ' ActiveWorkbook.SaveAs Filename:=CSVFilename, FileFormat:=xlCSV, CreateBackup:=False

    If (StreetNoStat > 0 Or StreetTypeStat > 0 Or FNameStat > 0 ) Then
        ErrMsg = "No error found"
        If ErrCount > 0 Then
            ErrMsg = ErrCount & " error(s) found at line(s) "
            For j = LBound(ErrLines) To UBound(ErrLines)
                ErrMsg = ErrMsg & "### " & ErrLines(j) & " "
            Next j
            ErrMsg = ErrMsg & " ###. Manual edit required"
        End If
        ErrMsg = ErrMsg  & vbNewLine & vbNewLine
        ErrMsg = ErrMsg & "In total, we cleaned" & vbNewLine
        If DuplicateHeaderRow Then
            ErrMsg = ErrMsg & "- Duplicate header line" & vbNewLine
        End If
        MSgBox ( ErrMsg & _
                    "- StreetNo: " & StreetNoStat & vbNewLine & _
                    "- Street Unit: " & StreetUnitStat & vbNewLine & _
                    "- Street Type: " & StreetTypeStat & vbNewLine & _
                    "- Firstnames: " & FNameStat & vbNewLine & _
                    vbNewLine & VbNewLine & vbNewLine & _
                    "Your new CSV file is saved at: [" & CSVFileName & "]" & _
                    vbNewLine & vbNewLine & _
                    "When you close this dialog the program will reset itself to get ready for next use" & _
                    vbNewLine & vbNewLine & _
                    "As your CSV file has already been created, DO NOT save this file as csv again" & _
                    vbNewLine & _
                    "Select ***[Don't Save]*** when prompted." & _
                    vbNewLine & vbNewLine & _
                    "Have FUN :-)" _
                )
    End If

    Sheet1.Cells.Clear

End Sub

Private Function StreetNo(FullStreetName)
    StreetNo = Left(FullStreetName, InStr(FullStreetName, " "))
End Function


Private Function StreetType(FullStreetName)
    StreetType = Right(FullStreetName, Len(FullStreetName) - InStrRev(FullStreetName, " "))
End Function


Private Function StreetNameOnly(FullStreetName)
    If InStrRev(FullStreetName, " ") > 0 Then
        StreetNameOnly = Trim(Mid(FullStreetName, InStr(FullStreetName, " "), InStrRev(FullStreetName, " ") - InStr(FullStreetName, " ")))
    End If
End Function
