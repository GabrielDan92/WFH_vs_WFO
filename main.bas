Public rng              As Range, _
        rngSheet        As Range, _
        wsTranspose     As Worksheet, _
        wsTruePortal    As Worksheet, _
        wsOracle        As Worksheet, _
        wsFinal         As Worksheet, _
        wsTeams         As Worksheet, _
        dict            As Dictionary, _
        dict2           As Dictionary, _
        dayArr          As Variant, _
        arr             As Variant, _
        truePortalPath  As String, _
        oraclePath      As String, _
        lastRow         As Long, _
        month           As String, _
        inpYear         As String, _
        key             As Variant, _
        finalDate       As Date, _
        numberOfDays    As Integer
        
Sub Main(ByVal inputMonth As String, ByVal inputYear As String)

    With Application
        .ScreenUpdating = False
        .Calculation = xlManual
        .EnableEvents = False
        oldStatusBar = .DisplayStatusBar
        .DisplayStatusBar = True
        .StatusBar = "Please wait..."
    End With

    'month = inputMonth
    'inpYear = inputYear
    userDay = 1
    pass = "userDefinedPass"
    
    'create a date object to be passed to EoMonth function
        finalDate = DateSerial(inputYear, inputMonth, userDay)
    
    'get the last day of the month
        numberOfDays = CInt(Format(CDate(WorksheetFunction.EoMonth(finalDate, 0)), "dd"))
    
    Set wsTranspose = Sheet6
    Set wsTruePortal = Sheet2
    Set wsOracle = Sheet3
    Set wsFinal = Sheet1
    
    If inputYear = "19" Then
        Set wsTeams = Sheet7
    Else
        Set wsTeams = Sheet4
    End If
    
    '========================PREPARE THE DESTINATION FILE (I)==========================
    Call prepareFile1(pass)
    '============================RETRIEVE THE SOURCE DATA==============================
    Call retrieveData(inputMonth)
    '========================PREPARE THE DESTINATION FILE (II)=========================
    Call prepareFile2
    '============================TRANSPOSE THE SOURCE DATA=============================
    Call Transpose(inputMonth, inputYear)
    '=========================CALCULATE THE RETRIEVED RESULTS==========================
    Call calculateResults(pass)

    With Application
        .EnableEvents = True
        '.Calculation = xlAutomatic
        .ScreenUpdating = True
        .StatusBar = False
        .DisplayStatusBar = oldStatusBar
    End With

End Sub

Private Sub prepareFile1(ByVal pass As String)

    'clear any previous data from the workbook
        With wsTruePortal
            .Cells(1, 1).CurrentRegion.ClearContents
            .Cells(1, 1).CurrentRegion.ClearFormats
        End With
        With wsOracle
            .Cells(1, 1).CurrentRegion.ClearContents
            .Cells(1, 1).CurrentRegion.ClearFormats
        End With
        With wsTranspose
            .Unprotect (pass)         'unprotect the sheet
            If .AutoFilterMode Then wsTranspose.AutoFilterMode = False
            .Cells(1, 1).CurrentRegion.Offset(2).ClearContents
            .Cells(1, 1).CurrentRegion.Offset(2).ClearFormats
        End With
        With wsFinal
            .Unprotect (pass)         'unprotect the sheet
            If .AutoFilterMode Then wsFinal.AutoFilterMode = False
            anchorColumnForFinal = 1
            j = 1
            For i = 1 To 5
                .Cells(1, anchorColumnForFinal).CurrentRegion.Offset(2).Font.Bold = False
                .Cells(1, anchorColumnForFinal).CurrentRegion.Offset(2).ClearContents
                .Cells(1, anchorColumnForFinal).Value = wsTeams.Cells(1, j).Value
                anchorColumnForFinal = anchorColumnForFinal + 13
                j = j + 3
            Next i
        End With
    
End Sub

Private Sub retrieveData(ByVal inputMonth As String)
                                                                                                                                                                                        
'===========================HR PORTAL===========================

    'open the workbook
        Dim wb As Workbook
        
        With Application.FileDialog(msoFileDialogOpen)
            .Title = "Open the HR Portal file for month " & inputMonth
            .AllowMultiSelect = False
            .Show
            File = .SelectedItems(1)
        End With
        
        Set wb = Workbooks.Open(File, 2)
        
    'copy the data
        Set rngSheet = wb.Sheets(1).Cells(1, 1).CurrentRegion
        lastRow = rngSheet.Rows.Count
        arr = rngSheet
        wsTruePortal.Cells(1, 1).Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        Erase arr
        
    'close the workbook
        wb.Close (False)
        Set wb = Nothing
        
    'concat the first and last name elements in column 'A'
        With wsTruePortal
            .Columns("A:B").Delete Shift:=xlToLeft
            .Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            .Cells(1, 1).Value = "Nume Complet"
            .Range("A2:A" & lastRow).FormulaR1C1 = "=TRIM(CONCATENATE(R[0]C[2],"" "", R[0]C[3]))"
        End With
                                                                                                                                                                   
'=============================ORACLE==============================

    'open the workbook
        With Application.FileDialog(msoFileDialogOpen)
            .Title = "Open the Oracle file for month " & inputMonth
            .AllowMultiSelect = False
            .Show
            File = .SelectedItems(1)
        End With
        
        Set wb = Workbooks.Open(File, 2)
    
    'copy the data
        With wb.Sheets("Export Worksheet")
            If .AutoFilterMode Then .AutoFilterMode = False
            Set rngSheet = .Cells(1, 1).CurrentRegion
            arr = rngSheet
            wsOracle.Cells(1, 1).Resize(UBound(arr, 1), UBound(arr, 2)) = arr
            Erase arr
        End With
        
    'close the workbook
        wb.Close (False)
        Set wb = Nothing
        
        With wsOracle
            .Columns("A:A").Delete Shift:=xlToLeft
            .Columns("C:L").Delete Shift:=xlToLeft
            .Columns("D:X").Delete Shift:=xlToLeft
            .Columns("A:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End With

End Sub

Private Sub prepareFile2()
    
    'format the calendar day and remove the comma from the employee name
        Set rngSheet = wsOracle.Cells(1, 3).CurrentRegion
        lastRow = rngSheet.Rows.Count
        
        With wsOracle
            .Cells(1, 1).Value = "Full Name"
            .Cells(1, 2).Value = "Date"
            .Range("A2:A" & lastRow).FormulaR1C1 = "=SUBSTITUTE(TRIM(R[0]C[3]),"","","""")"
            .Columns("B:B").NumberFormat = "@"
            
            'convert the date to date serial
                Set rng = .Cells(1, 3).CurrentRegion
                lastRow = rng.Rows.Count
                
                For i = 2 To lastRow
                    .Cells(i, 2).Value = CLng(.Cells(i, 5).Value)
                Next i
        End With

    'push the presence type in the dict2 dictionary
        With ThisWorkbook.Sheets("Types of presence")
        
            Set dict2 = New Dictionary
            Set rngSheet = .Cells(1, 1).CurrentRegion
            lastRow = rngSheet.Rows.Count
            
            For i = 2 To lastRow
                dict2(.Cells(i, 1).Value) = .Cells(i, 2).Value
            Next i
            
        End With
    
    'import the Oracle employees to the dictionary
        With wsOracle
        
            Set dict = New Dictionary
            dict.CompareMode = vbTextCompare
            Set rngSheet = .Cells(1, 1).CurrentRegion
            lastRow = rngSheet.Rows.Count
            
            For i = 2 To lastRow
                dict(.Cells(i, 1).Value) = .Cells(i, 1).Value
            Next i
            If dict.Exists("") = True Then dict.Remove ("")
            
        End With

End Sub

Private Sub Transpose(ByVal inputMonth As String, ByVal inputYear As String)

    With wsTranspose
        .Columns("C:C").NumberFormat = "@"
        .Columns("D:D").NumberFormat = "@"
    End With
    
    Set rng = wsTruePortal.Range("A1").CurrentRegion
    lastRow = rng.Rows.Count
    
    'we start to retrieve values from True Portal to the working sheet starting with the 3rd row
    currentrow = 3
    
    'i' represents the  row index of the employee, going one row down means going to the next employee in the list
    For i = 2 To lastRow
    
        If dict.Exists(wsTruePortal.Cells(i, 1).Value) = False Then
            GoTo nextFor
        Else
            With wsTranspose
                .Range("A" & currentrow & ":A" & currentrow + numberOfDays) = wsTruePortal.Cells(i, 2).Value    'employee QR code
                .Range("B" & currentrow & ":B" & currentrow + numberOfDays) = wsTruePortal.Cells(i, 1).Value    'employee Full Name
                
                 For j = 1 To numberOfDays
                    .Cells(currentrow, 3).Value = CLng(DateValue(j & "-" & inputMonth & "-" & inputYear))       'calendar day
                    .Cells(currentrow, 4).Value = wsTruePortal.Cells(i, j + 4).Value                            'presence type
                    
                    If .Cells(currentrow, 4).Value <> "" Then
                        .Cells(currentrow, 5).Value = dict2(.Cells(currentrow, 4).Value)                        'in/out of office
                        If dict2(.Cells(currentrow, 4).Value) = "Office" Or _
                            dict2(.Cells(currentrow, 4).Value) = "Home office" Then                             'countifs formula
                                .Cells(currentrow, 6).FormulaR1C1 = "=COUNTIFS(ORACLE!C[-5],R[0]C[-4],ORACLE!C[-4],R[0]C[-3])"
                        End If
                    Else
                        .Cells(currentrow, 5).Value = "OFF"
                    End If
                    currentrow = currentrow + 1
                Next j
            End With
        End If
        
nextFor:
    Next i
    
    'clean up
        dict.RemoveAll
        dict2.RemoveAll

End Sub

Private Sub calculateResults(ByVal pass As String)

'push the employees from the 'wsTranspose' sheet to the dictionary
    With wsTranspose
        Set rngSheet = .Range("A1").CurrentRegion
        lastRow = rngSheet.Rows.Count
        
        For i = 3 To lastRow
            dict(.Cells(i, 2).Value) = .Cells(i, 1).Value
        Next i
    End With
    
    
'populate the 'Final Results' sheet
    anchorColumnForTeams = 1
    anchorColumnForFinal = 1
    i = 3

For k = 1 To 5
    
    'push the employees from the 'Teams' sheet to the dictionary2
        With wsTeams
            Set rngSheet = .Cells(1, anchorColumnForTeams).CurrentRegion
            lastRow = rngSheet.Rows.Count
            dict2.RemoveAll
            
            For j = 2 To lastRow
                dict2(Trim(.Cells(j, anchorColumnForTeams).Value)) = .Cells(j, anchorColumnForTeams + 1).Value
            Next j
        
        End With
    
            
    'populate the final table with the employees, QRs and formulas
        For Each key In dict.Keys
            If dict2.Exists(key) = True Then
                With wsFinal
                    .Cells(i, anchorColumnForFinal).Value = dict(key)       'QR
                    .Cells(i, anchorColumnForFinal + 1).Value = key         'Full Name
                    .Cells(i, anchorColumnForFinal + 10).FormulaR1C1 = "=COUNTIFS(Transpose!C2,RC[-9],Transpose!C5,""Office"")"
                    .Cells(i, anchorColumnForFinal + 9).FormulaR1C1 = "=COUNTIFS(Transpose!C2,RC[-8],Transpose!C5,""Home office"")"
                    .Cells(i, anchorColumnForFinal + 8).FormulaR1C1 = "=IFERROR(RC[1]/(RC[1]+RC[2]),0)"
                    .Cells(i, anchorColumnForFinal + 7).FormulaR1C1 = "=SUMIFS(Transpose!C6,Transpose!C2,RC[-6],Transpose!C5,""Office"")"
                    .Cells(i, anchorColumnForFinal + 6).FormulaR1C1 = "=SUMIFS(Transpose!C6,Transpose!C2,RC[-5],Transpose!C5,""Home office"")"
                    .Cells(i, anchorColumnForFinal + 5).FormulaR1C1 = "=IFERROR(RC[1]/(RC[1]+RC[2]),0)"
                    .Cells(i, anchorColumnForFinal + 4).FormulaR1C1 = "=RC[2]+RC[3]"
                    .Cells(i, anchorColumnForFinal + 3).FormulaR1C1 = "=IFERROR(RC[4]/RC[7],0)"
                    .Cells(i, anchorColumnForFinal + 2).FormulaR1C1 = "=IFERROR(RC[4]/RC[7],0)"
                    i = i + 1
                End With
            End If
        Next key
        
        
        With wsFinal
        
            Set rng = .Cells(1, anchorColumnForFinal).CurrentRegion
            lastRow = rng.Rows.Count
            
            .Cells(lastRow + 1, anchorColumnForFinal + 1).Value = "Total Average"
            .Cells(lastRow + 1, anchorColumnForFinal + 2).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 3).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 4).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 5).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 6).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 7).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 8).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 9).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            .Cells(lastRow + 1, anchorColumnForFinal + 10).FormulaR1C1 = "=AVERAGE(R[-" & (lastRow - 2) & "]C:R[-1]C)"
            
            .Range(.Cells(lastRow + 1, anchorColumnForFinal + 1), .Cells(lastRow + 1, anchorColumnForFinal + 10)).Font.Bold = True
            
        End With

        
    anchorColumnForTeams = anchorColumnForTeams + 3
    anchorColumnForFinal = anchorColumnForFinal + 13
    i = 3
     
Next k
     
Set dict = Nothing

'make the tables more readable
    wsTranspose.Columns("C:C").NumberFormat = "m/d/yyyy"
    With wsTranspose.Cells(2, 1).CurrentRegion
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    anchorColumnForFinal = 1
        For i = 1 To 5
            With wsFinal.Cells(1, anchorColumnForFinal).CurrentRegion
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
            End With
            anchorColumnForFinal = anchorColumnForFinal + 13
        Next i
     
'protect the sheets
    wsTranspose.Protect (pass)
    wsFinal.Protect (pass)
     
End Sub
