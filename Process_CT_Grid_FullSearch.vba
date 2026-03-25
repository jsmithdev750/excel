Option Explicit

Public Sub Process_CT_Grid_FullSearch()

    '-----------------------------------
    ' Declare variables
    '-----------------------------------
    Dim ws As Worksheet
    
    Dim lastRow As Long
    Dim lastCol As Long
    
    Dim contractCol As Long
    Dim headerRow As Long
    
    Dim i As Long
    Dim r As Long, c As Long, cc As Long
    
    Dim cellValue As String
    Dim rawInput As String
    Dim parts As Variant
    
    Dim regionName As String
    Dim contractName As String
    Dim valueNum As Variant
    
    Dim regionCol As Long
    Dim contractRow As Long
    
    Dim ctFound As Boolean
    Dim ctRow As Long
    Dim headerScanLimit As Long
    Dim contractLastRow As Long
    
    ' NEW: Store skipped rows
    Dim skippedList As Collection
    Set skippedList = New Collection
    
    '-----------------------------------
    ' Setup worksheet
    '-----------------------------------
    Set ws = ThisWorkbook.Sheets("CT GRID Last value")
    
    ' Refresh all queries/connections
    ThisWorkbook.RefreshAll
    
    ' Wait for refresh to complete
    DoEvents
    

    '-----------------------------------
    ' Detect boundaries
    '-----------------------------------
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    '-----------------------------------
    ' STEP 1: Find "Contract" column
    '-----------------------------------
    contractCol = 0
    
    For r = 1 To lastRow
        For c = 1 To lastCol
            If CleanText(ws.Cells(r, c).Value) = "contract" Then
                contractCol = c
                headerRow = r
                Exit For
            End If
        Next c
        If contractCol <> 0 Then Exit For
    Next r
    
    If contractCol = 0 Then
        MsgBox "Contract column not found.", vbCritical
        Exit Sub
    End If
    
    '-----------------------------------
    ' STEP 1.5: Find last row
    '-----------------------------------
    contractLastRow = ws.Cells(ws.Rows.Count, contractCol).End(xlUp).Row
    
    '-----------------------------------
    ' STEP 1.6: Clear CT columns
    '-----------------------------------
    headerScanLimit = Application.WorksheetFunction.Min(10, contractLastRow)

Dim colLastRow As Long

For c = contractCol + 1 To lastCol
    ' Check if this column header contains "ct"
    If InStr(1, CleanText(ws.Cells(headerRow, c).Value), "ct") > 0 Then
        
        ' Find the last used row in THIS column
        colLastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        
        ' Only clear below the header row
        If colLastRow > headerRow Then
            ws.Range(ws.Cells(headerRow + 1, c), ws.Cells(colLastRow, c)).ClearContents
        End If
    End If
Next c
    
    '-----------------------------------
    ' STEP 2: LOOP INPUT
    '-----------------------------------
    For i = 2 To lastRow
        
        rawInput = ws.Cells(i, 1).Value '<<< KEEP ORIGINAL
        cellValue = CleanText(rawInput)
        
        If cellValue = "" Then GoTo NextRow
        
        parts = Split(cellValue, " ")
        If UBound(parts) < 2 Then GoTo NextRow
        
        '-----------------------------------
        ' PARSING
        '-----------------------------------
        Dim partCount As Long
        Dim k As Long
        
        partCount = UBound(parts)
        regionName = parts(0)
        
        ' Value
        For k = partCount To 0 Step -1
            If Trim(parts(k)) <> "" Then
                valueNum = parts(k)
                Exit For
            End If
        Next k
        
        ' Contract
        contractName = ""
        For k = 1 To partCount
            If Trim(parts(k)) <> "" And parts(k) <> valueNum Then
                contractName = contractName & parts(k) & " "
            End If
        Next k
        
        contractName = Trim(contractName)
        
        ' Put contract into Column B
        'ws.Cells(i, 2).NumberFormat = "@"
        'ws.Cells(i, 2).Value = contractName
        
        contractName = CleanText(contractName)
        
        '-----------------------------------
        ' STEP 3: Find REGION
        '-----------------------------------
        regionCol = 0
        
        For cc = contractCol + 1 To lastCol
            If CleanText(ws.Cells(headerRow, cc).Value) = regionName Then
                regionCol = cc
                Exit For
            End If
        Next cc
        
        If regionCol = 0 Then
            skippedList.Add rawInput
            GoTo NextRow
        End If
        
        '-----------------------------------
        ' STEP 4: Find CONTRACT
        '-----------------------------------
        Dim sheetVal As Variant
        Dim inputDate As Date
        Dim sheetDate As Date
        Dim isInputDate As Boolean
        Dim isSheetDate As Boolean
        
        contractRow = 0
        
        isInputDate = TryParseMonthContract(contractName, inputDate)
        
        For r = headerRow + 1 To contractLastRow
            
            sheetVal = ws.Cells(r, contractCol).Value
            
            If IsDate(sheetVal) Then
                sheetDate = CDate(sheetVal)
                isSheetDate = True
            Else
                isSheetDate = TryParseMonthContract(CStr(sheetVal), sheetDate)
            End If
            
            If isInputDate And isSheetDate Then
                If Year(inputDate) = Year(sheetDate) And Month(inputDate) = Month(sheetDate) Then
                    contractRow = r
                    Exit For
                End If
            Else
                If replace(contractName, "-", "") = replace(CleanText(sheetVal), "-", "") Then
                    contractRow = r
                    Exit For
                End If
            End If
            
        Next r
        
        If contractRow = 0 Then
            skippedList.Add rawInput
            GoTo NextRow
        End If
        
        '-----------------------------------
        ' WRITE VALUE
        '-----------------------------------
        ws.Cells(contractRow, regionCol + 1).Value = valueNum
        
NextRow:
    Next i
    
    '-----------------------------------
    ' STEP 6: OUTPUT FILTERED CONTRACTS
    '-----------------------------------
    Dim outputCol As Long
    Dim outputStartRow As Long
    
    outputCol = FindHeaderColumn(ws, "contract that is been filter out")
    
    If outputCol = 0 Then
        MsgBox "Output column not found.", vbExclamation
        Exit Sub
    End If
    
    ' Clear old results
    ws.Range(ws.Cells(headerRow + 1, outputCol), _
             ws.Cells(ws.Rows.Count, outputCol)).ClearContents
    
    outputStartRow = headerRow
    
    ' Write skipped items
    For i = 1 To skippedList.Count
        ws.Cells(outputStartRow + i - 1, outputCol).Value = skippedList(i)
    Next i
    
    ' Recalculate formulas (either workbook or sheet)
    Application.CalculateFull   ' recalculates entire workbook
    ' or use ws.Calculate       ' recalculates only this sheet
    
    MsgBox "Processing completed.", vbInformation

End Sub

'-----------------------------------
' FIND HEADER COLUMN
'-----------------------------------
Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For r = 1 To lastRow
        For c = 1 To lastCol
            If CleanText(ws.Cells(r, c).Value) = CleanText(headerName) Then
                FindHeaderColumn = c
                Exit Function
            End If
        Next c
    Next r
    
    FindHeaderColumn = 0

End Function

'-----------------------------------
' CLEAN FUNCTION
'-----------------------------------
Private Function CleanText(ByVal txt As String) As String
    
    If txt = "" Then
        CleanText = ""
        Exit Function
    End If
    
    txt = replace(txt, Chr(160), " ")
    txt = Application.WorksheetFunction.Clean(txt)
    txt = Trim(txt)
    
    Do While InStr(txt, "  ") > 0
        txt = replace(txt, "  ", " ")
    Loop
    
    CleanText = LCase(txt)

End Function


'-----------------------------------
' PARSE MONTH CONTRACT
'-----------------------------------
Private Function TryParseMonthContract(ByVal txt As String, ByRef outDate As Date) As Boolean
    
    On Error GoTo Fail
    
    Dim temp As String
    
    temp = CleanText(txt)
    temp = replace(temp, "-", " ")
    temp = replace(temp, "sept", "sep")
    
    If temp Like "*[a-z][a-z][a-z]* *##" Then
        outDate = DateValue("1 " & Application.WorksheetFunction.Proper(temp))
        TryParseMonthContract = True
        Exit Function
    End If

Fail:
    TryParseMonthContract = False

End Function


