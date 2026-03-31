Option Explicit
'============================================================
' Main Import Macro (with dynamic destination offsets)
'============================================================
Public Sub Import_Old_Japan_Power_Curve()

    Dim wbOrigin As Workbook, wbDest As Workbook
    Dim wsOrigin As Worksheet, wsDest As Worksheet, wsWeekday As Worksheet
    Dim tokyoCell As Range, spreadsCell As Range
    Dim headerRow As Long
    Dim startCol As Long, endCol As Long
    Dim regionCols As Collection
    Dim regionCell As Range
    Dim regionStartCol As Long, regionEndCol As Long
    Dim wk1Row As Long, wk2Row As Long, wk3Row As Long
    Dim C As Long, r As Long, lastRow As Long
    
    Dim todayDate As String
    Dim originSheetName As String
    Dim destSheetName As String
    Dim originSearchHeader As String
    Dim destSearchHeader As String
    Dim originSheetNameInput As String
    Dim originSheetNameWeekDay As String
    
    Dim todayYYMMDD As String
    Dim destPattern As String
    Dim wb As Workbook
    
    Dim destTokyoCell As Range
    Dim destHeaderRow As Long, destStartCol As Long
    Dim rowOffset As Long, colOffset As Long
    
    Dim wsInput As Worksheet
    Dim colMap As Object
    Dim headers As Variant
    Dim found As Boolean
    Dim d As Range
    Dim h As Variant
    Dim colContract As Long
    Dim firstDataRow As Long, lastDataRow As Long
    Dim headerRowInput As Long
    Dim baseRow As Long
    
    Dim valToPaste As Variant
    Dim key As Variant
    Dim histKey As String
    Dim colKey As String
    
    Dim histSheet As Worksheet
    Dim histDateColumn As Long
    Dim contractColumn As Long
    Dim targetCell As Range
    Dim contract As String
    Dim firstHistRow As Long, lastHistRow As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


    ' Settings
    todayDate = Sheet1.Range("A3").value
    originSheetName = Sheet1.Range("A14").value
    destSheetName = Sheet1.Range("B14").value
    originSearchHeader = Sheet1.Range("A11").value
    destSearchHeader = Sheet1.Range("B11").value
    originSheetNameInput = Sheet1.Range("A15").value
    originSheetNameWeekDay = Sheet1.Range("A16").value
    
    
    '--------------------------------
    ' Destination file pattern
    '--------------------------------
    todayYYMMDD = Format(todayDate, "yy.mm.dd")
    destPattern = "*Vanir EEX Japan Power Curve_" & todayYYMMDD & "*"

    '--------------------------------
    ' Locate origin workbook
    '--------------------------------
    For Each wb In Workbooks
        If wb.Name Like "*NEW CURVE_OUTPUT*" Then
            Set wbOrigin = wb
            Exit For
        End If
    Next wb

    If wbOrigin Is Nothing Then
        MsgBox "Origin workbook not found", vbCritical
        GoTo ExitSafe
    End If

    '--------------------------------
    ' Locate destination workbook
    '--------------------------------
    For Each wb In Workbooks
        If wb.Name Like destPattern And Not wb.Name Like "*NEW FORMAT*" Then
            Set wbDest = wb
            Exit For
        End If
    Next wb

    If wbDest Is Nothing Then
        MsgBox "Destination workbook not open", vbCritical
        GoTo ExitSafe
    End If

    '--------------------------------
    ' Get origin/destination sheets (case-insensitive)
    '--------------------------------
    Set wsOrigin = GetSheetByNameInsensitive(wbOrigin, originSheetName)
    Set wsDest = GetSheetByNameInsensitive(wbDest, destSheetName)
    Set wsInput = GetSheetByNameInsensitive(wbOrigin, originSheetNameInput)
    Set wsWeekday = GetSheetByNameInsensitive(wbOrigin, originSheetNameWeekDay)
    
    If wsOrigin Is Nothing Then
        MsgBox "Sheet 'OUTPUT' not found in origin workbook", vbCritical
        GoTo ExitSafe
    End If
    
    If wsDest Is Nothing Then
        MsgBox "Sheet 'CURVE' not found in destination workbook", vbCritical
        GoTo ExitSafe
    End If

    If wsInput Is Nothing Then
        MsgBox "Sheet 'INPUT' not found in origin workbook", vbCritical
        GoTo ExitSafe
    End If

    If wsWeekday Is Nothing Then
        MsgBox "Sheet 'WEEKS_DAYS' not found in origin workbook", vbCritical
        GoTo ExitSafe
    End If
    
    '--------------------------------
    ' Clear old charts/pictures in destination
    '--------------------------------
    Dim shp As Shape
    
    wsDest.Activate
    DoEvents
    
    If wsDest.ProtectContents Then
        wsDest.Unprotect ' add password if needed
    End If
    
    For Each shp In wsDest.Shapes
        On Error Resume Next
        If shp.Type = msoChart Or shp.Type = msoPicture Then
            shp.Delete
        End If
        On Error GoTo 0
    Next shp
    
     '  Dim i As Long
    'For i = wsDest.ChartObjects.Count To 1 Step -1
     '   wsDest.ChartObjects(i).Delete
    'Next i

    'For i = wsDest.Shapes.Count To 1 Step -1
     '   If wsDest.Shapes(i).Type = msoPicture Then wsDest.Shapes(i).Delete
    'Next i

    '--------------------------------
    ' Find TOKYO AREA in origin and destination (dynamic)
    '--------------------------------
    Set tokyoCell = wsOrigin.Cells.Find(originSearchHeader, LookAt:=xlPart)
    If tokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found in origin sheet", vbCritical
        GoTo ExitSafe
    End If
    headerRow = tokyoCell.row
    startCol = tokyoCell.mergeArea.Column

    Set destTokyoCell = wsDest.Cells.Find(originSearchHeader, LookAt:=xlPart)
    If destTokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found in destination sheet", vbCritical
        GoTo ExitSafe
    End If
    destHeaderRow = destTokyoCell.row
    destStartCol = destTokyoCell.mergeArea.Column

    '--------------------------------
    ' Find SPREADS
    '--------------------------------
    Set spreadsCell = wsOrigin.Cells.Find(destSearchHeader, LookAt:=xlPart)
    If spreadsCell Is Nothing Then
        MsgBox "Spreads header not found", vbCritical
        GoTo ExitSafe
    End If
    endCol = spreadsCell.mergeArea.Columns(spreadsCell.mergeArea.Columns.Count).Column

    '--------------------------------
    ' Collect region headers
    '--------------------------------
    Set regionCols = New Collection
    C = startCol
    Do While C <= endCol
        If wsOrigin.Cells(headerRow, C).MergeCells Then
            regionCols.Add wsOrigin.Cells(headerRow, C)
            C = C + wsOrigin.Cells(headerRow, C).mergeArea.Columns.Count
        Else
            C = C + 1
        End If
    Loop

    '--------------------------------
    ' Process each region
    '--------------------------------
    Dim ch As ChartObject
    Dim chartLeft As Double, chartTop As Double
    Dim wk1Top As Double, wk2Top As Double
    
    For Each regionCell In regionCols

        regionStartCol = regionCell.mergeArea.Column
        regionEndCol = regionStartCol + regionCell.mergeArea.Columns.Count - 1

        wk1Row = headerRow + 2
        wk2Row = wk1Row + 7
        wk3Row = wk2Row + 7

        ' WEEK CONTRACTS
        rowOffset = wk1Row - headerRow
        colOffset = regionStartCol - startCol
        CopyRowFast wsOrigin, wsDest, wk1Row, regionStartCol, regionEndCol, destHeaderRow + rowOffset, destStartCol + colOffset
        rowOffset = wk2Row - headerRow
        CopyRowFast wsOrigin, wsDest, wk2Row, regionStartCol, regionEndCol, destHeaderRow + rowOffset, destStartCol + colOffset
        rowOffset = wk3Row - headerRow
        CopyRowFast wsOrigin, wsDest, wk3Row, regionStartCol, regionEndCol, destHeaderRow + rowOffset, destStartCol + colOffset

        ' DAY CONTRACTS (AREA logic) + red font check
        If InStr(1, regionCell.value, "AREA", vbTextCompare) > 0 Then

            Dim col1 As Long, col2 As Long, col3 As Long
            Dim contractDate As Date
            Dim destDate As Date
            Dim destRowStart As Long, destColStart As Long
            
            col1 = regionEndCol - 2
            col2 = regionEndCol - 1
            col3 = regionEndCol
            destDate = todayDate

            lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, col1).End(xlUp).row

            ' Dynamic paste destination
            destRowStart = destHeaderRow + (wk1Row - headerRow)
            destColStart = destStartCol + (col1 - startCol)

            wsDest.Range(wsDest.Cells(destRowStart, destColStart), _
                         wsDest.Cells(destRowStart + (lastRow - wk1Row), destColStart + (col3 - col1))).value = _
                wsOrigin.Range(wsOrigin.Cells(wk1Row, col1), wsOrigin.Cells(lastRow, col3)).value

            ' Red font only on last col if date condition met, else reset to black
            For r = 0 To wk3Row - wk1Row
                If IsDate(wsDest.Cells(destRowStart + r, destColStart + 1).value) Then
                    contractDate = wsDest.Cells(destRowStart + r, destColStart + 1).value
                    If contractDate <= destDate Or contractDate = destDate + 1 Then
                        wsDest.Cells(destRowStart + r, destColStart + 2).Font.Color = RGB(255, 0, 0)
                    Else
                        wsDest.Cells(destRowStart + r, destColStart + 2).Font.Color = RGB(0, 0, 0)
                    End If
                Else
                    wsDest.Cells(destRowStart + r, destColStart + 2).Font.Color = RGB(0, 0, 0)
                End If
            Next r

            ' Copy charts from origin to destination for AREA region
            wk1Top = wsDest.Rows(destHeaderRow + (wk1Row - headerRow)).Top
            wk2Top = wsDest.Rows(destHeaderRow + (wk2Row - headerRow)).Top

            For Each ch In wsOrigin.ChartObjects
                If ch.Left >= wsOrigin.Cells(headerRow, regionStartCol).Left And _
                   ch.Left + ch.Width <= wsOrigin.Cells(headerRow, regionEndCol).Left + wsOrigin.Cells(headerRow, regionEndCol).Width Then

                    ' Determine horizontal position
                    chartLeft = wsDest.Cells(destHeaderRow, destStartCol + (regionStartCol - startCol)).Left

                    ' Determine vertical position
                    If ch.Top < wsOrigin.Rows(wk2Row).Top Then
                        chartTop = wk1Top + wsDest.Rows(destHeaderRow + (wk1Row - headerRow)).Height
                    Else
                        chartTop = wk2Top + wsDest.Rows(destHeaderRow + (wk2Row - headerRow)).Height
                    End If

                    ' Copy chart as picture and paste
                    ch.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                    wsDest.Paste

                    ' Move pasted image
                    With wsDest.Shapes(wsDest.Shapes.Count)
                        .Left = chartLeft
                        .Top = chartTop
                        .LockAspectRatio = msoTrue
                    End With

                End If
            Next ch

        End If

        ' REMAINING CONTRACTS
        lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, regionStartCol).End(xlUp).row
        If lastRow > wk3Row Then
            rowOffset = wk3Row + 1 - headerRow
            colOffset = regionStartCol - startCol
            wsDest.Range(wsDest.Cells(destHeaderRow + rowOffset, destStartCol + colOffset), _
                         wsDest.Cells(destHeaderRow + rowOffset + (lastRow - wk3Row - 1), destStartCol + colOffset + (regionEndCol - regionStartCol))).value = _
                wsOrigin.Range(wsOrigin.Cells(wk3Row + 1, regionStartCol), wsOrigin.Cells(lastRow, regionEndCol)).value
        End If

    Next regionCell
    
'========================================================
    '-------------------------------
    ' Locate Column Headers
    '-------------------------------
    Set colMap = CreateObject("Scripting.Dictionary")
    headers = Array("TBL", "CBL", "KBL", "TPK", "CPK", "KPK", "TOPK", "COPK", "KOPK")
    
    For Each h In headers
        found = False
        For Each d In wsInput.UsedRange
            ' Only consider columns before the stop column
            If LCase(Trim(d.value)) = LCase(h) Then
                colMap(h) = d.Column
                ' Capture headerRow once (first header found)
                If headerRowInput = 0 Then headerRowInput = d.row
                    found = True
                Exit For
            End If
        Next d
        
        If Not found Then
            MsgBox "Column header '" & h & "' not found '", vbCritical
            Exit Sub
        End If
    Next h


'==================================
' WEEK DAY
    Dim colWeekDayMap As Object
    Dim headersWeekDay As Variant
    Dim foundWeekDay As Boolean
    Dim d1 As Range
    Dim h1 As Variant
    Dim colContractBL As Long, colContractPK As Long
    Dim firstDataRowBL As Long, lastDataRowBL As Long
    Dim firstDataRowPK As Long, lastDataRowPK As Long
    Dim headerRowInputWeekDay As Long
    Dim baseRow1 As Long
    
    Set colWeekDayMap = CreateObject("Scripting.Dictionary")
    headersWeekDay = Array("TBL", "CBL", "KBL", "TPK", "CPK", "KPK")
    
    For Each h1 In headersWeekDay
        foundWeekDay = False
        For Each d1 In wsWeekday.UsedRange
            ' Only consider columns before the stop column
            If LCase(Trim(d1.value)) = LCase(h1) Then
                colWeekDayMap(h1) = d1.Column
                ' Capture headerRow once (first header found)
                If headerRowInputWeekDay = 0 Then headerRowInputWeekDay = d1.row
                    foundWeekDay = True
                Exit For
            End If
        Next d1
        
        If Not foundWeekDay Then
            MsgBox "Column header '" & h1 & "' not found '", vbCritical
            Exit Sub
        End If
    Next h1
    
    '-------------------------------
    ' Contract column = column before TBL
    '-------------------------------
    colContractBL = colWeekDayMap("TBL") - 1
    If colContractBL < 1 Then
        MsgBox "Invalid contract column.", vbCritical
        Exit Sub
    End If
    
    colContractPK = colWeekDayMap("TPK") - 1
    If colContractPK < 1 Then
        MsgBox "Invalid contract column.", vbCritical
        Exit Sub
    End If
    
    '-------------------------------
    ' Data range -  BL
    '-------------------------------
    firstDataRowBL = headerRowInputWeekDay + 1
    lastDataRowBL = wsWeekday.Cells(wsWeekday.Rows.Count, colContractBL).End(xlUp).row
    If lastDataRowBL < firstDataRowBL Then
        MsgBox "No contract data found.", vbCritical
        Exit Sub
    End If
    
    '-------------------------------
    ' Data range -  PK
    '-------------------------------
    firstDataRowPK = headerRowInputWeekDay + 1
    lastDataRowPK = wsWeekday.Cells(wsWeekday.Rows.Count, colContractPK).End(xlUp).row
    If lastDataRowPK < firstDataRowPK Then
        MsgBox "No contract data found.", vbCritical
        Exit Sub
    End If
    
    '-------------------------------
    ' Display results in MsgBox
    '-------------------------------
    'Dim msg1 As String
    
   ' msg1 = "Header column mapping:" & vbCrLf
    'For Each key In colWeekDayMap.Keys
     '   msg1 = msg1 & key & " -> Column " & colWeekDayMap(key) & vbCrLf
    'Next key
    
    'msg1 = msg1 & "Contract column -> Column " & colContractBL & vbCrLf
    'msg1 = msg1 & "Data rows: " & firstDataRowBL & " to " & lastDataRowBL
    
    'MsgBox msg1, vbInformation, "bl Sheet Mapping"
    
    
    '-------------------------------
    ' Display results in MsgBox
    '-------------------------------
    'Dim msg2 As String
    
   ' msg2 = "Header column mapping:" & vbCrLf
    'For Each key In colWeekDayMap.Keys
     '   msg2 = msg2 & key & " -> Column " & colWeekDayMap(key) & vbCrLf
    'Next key
    
    'msg2 = msg2 & "Contract column -> Column " & colContractPK & vbCrLf
    'msg2 = msg2 & "Data rows: " & firstDataRowPK & " to " & lastDataRowPK
    
    'MsgBox msg2, vbInformation, "pk Sheet Mapping"
    
'===========================

    '-------------------------------
    ' Contract column = column before TBL
    '-------------------------------
    colContract = colMap("TBL") - 1
    If colContract < 1 Then
        MsgBox "Invalid contract column.", vbCritical
        Exit Sub
    End If

    '-------------------------------
    ' Data range
    '-------------------------------
    firstDataRow = headerRowInput + 1
    lastDataRow = wsInput.Cells(wsInput.Rows.Count, colContract).End(xlUp).row
    If lastDataRow < firstDataRow Then
        MsgBox "No contract data found.", vbCritical
        Exit Sub
    End If
    
    '-------------------------------
    ' Display results in MsgBox
    '-------------------------------
    'Dim msg As String
    
    'msg = "Header column mapping:" & vbCrLf
    'For Each key In colMap.Keys
        'msg = msg & key & " -> Column " & colMap(key) & vbCrLf
    'Next key
    
    'msg = msg & "Contract column -> Column " & colContract & vbCrLf
    'msg = msg & "Data rows: " & firstDataRow & " to " & lastDataRow
    
    'MsgBox msg, vbInformation, "INPUT Sheet Mapping"

'--------------------------------
' Loop through "Hist" sheets and process the data
'--------------------------------
For Each histSheet In wbDest.Worksheets
    
    If InStr(1, histSheet.Name, "Hist", vbTextCompare) = 1 Then

        histDateColumn = FindDateColumnFlexible(histSheet, Format(todayDate, "dd-mmm-yy"))

        If histDateColumn > 0 Then

            lastHistRow = histSheet.Cells(histSheet.Rows.Count, 1).End(xlUp).row

'--------------------------------
' Find "DAYS" section row
'--------------------------------
Dim iFind As Long
Dim daysStartRow As Long
daysStartRow = 0

daysStartRow = 0

Dim daysCell As Range

Set daysCell = histSheet.Columns(1).Find("DAYS", LookAt:=xlWhole, MatchCase:=False)

If Not daysCell Is Nothing Then
    daysStartRow = daysCell.row
End If

' Only enforce DAYS section for restricted sheets
If IsDayRestrictedSheet(histSheet.Name) Then
    If daysStartRow = 0 Then
        MsgBox "'DAYS' section not found in sheet: " & histSheet.Name, vbCritical
        Exit Sub
    End If
End If
            
            Dim contractDict As Object, contractDictBL As Object, contractDictPk As Object
            Set contractDict = CreateObject("Scripting.Dictionary")
            Set contractDictBL = CreateObject("Scripting.Dictionary")
            Set contractDictPk = CreateObject("Scripting.Dictionary")
            
            Dim i As Long
            For i = firstDataRow To lastDataRow
                contractDict(NormalizeContract(wsInput.Cells(i, colContract), daysStartRow, 0)) = i
            Next i
            
            For i = firstDataRowBL To lastDataRowBL
                contractDictBL(NormalizeContract(wsWeekday.Cells(i, colContractBL), daysStartRow, daysStartRow + 1)) = i
            Next i
            
            For i = firstDataRowPK To lastDataRowPK
                contractDictPk(NormalizeContract(wsWeekday.Cells(i, colContractPK), daysStartRow, daysStartRow + 1)) = i
            Next i
            
            colKey = ""
                        
            For Each key In colMap.Keys
                If InStr(1, histSheet.Name, key, vbTextCompare) > 0 Then
                    colKey = key
                    Exit For
                End If
            Next key
            
            If colKey = "" Then
                For Each key In colWeekDayMap.Keys
                    If InStr(1, histSheet.Name, key, vbTextCompare) > 0 Then
                        colKey = key
                        Exit For
                    End If
                Next key
            End If
            
            For r = 2 To lastHistRow
                
                contract = histSheet.Cells(r, 1).value
                
                If contract <> "" Then
                    
                    Dim normalizedContract As String
                    normalizedContract = NormalizeContract(contract, daysStartRow, r)
                    
                    Dim sourceType As String
                    Dim foundRow As Long
                    
                    foundRow = 0
                    sourceType = ""
                    
                    
                    If contractDict.exists(normalizedContract) Then
                        foundRow = contractDict(normalizedContract)
                        sourceType = "INPUT"
                    
                    ElseIf contractDictBL.exists(normalizedContract) Then
                        foundRow = contractDictBL(normalizedContract)
                        sourceType = "BL"
                    
                    ElseIf contractDictPk.exists(normalizedContract) Then
                        foundRow = contractDictPk(normalizedContract)
                        sourceType = "PK"
                    End If
                    
                    If foundRow > 0 And colKey <> "" Then
                        If sourceType = "INPUT" Then
                            If daysStartRow > 0 Then
                                If r >= daysStartRow Then GoTo SkipRow
                            End If
                        End If
                        
                        Select Case sourceType
                            Case "INPUT"
                                valToPaste = wsInput.Cells(foundRow, colMap(colKey)).value
                    
                            Case "BL", "PK"
                
                                valToPaste = wsWeekday.Cells(foundRow, colWeekDayMap(colKey)).value
                
                        End Select
                    
                        Set targetCell = histSheet.Cells(r, histDateColumn)
                        PasteIfSafe targetCell, valToPaste
                        ' Set red font if date id <= today +1 / today +1
                        If daysStartRow > 0 And r > daysStartRow Then
                            If IsDate(contract) Then
                                If Int(CDate(contract)) <= Int(destDate) + 1 Then
                                    targetCell.Font.Color = RGB(255, 0, 0)
                                Else
                                    targetCell.Font.Color = RGB(0, 0, 0)
                                End If
                                
                            End If
                        
                        End If
                        
                    
                    End If
                
                End If
SkipRow:
            Next r
        
        Else
            MsgBox "No matching date found in sheet: " & histSheet.Name, vbExclamation
        End If
    
    End If

Next histSheet

'================================

    wbDest.Save
    MsgBox "Import completed"

ExitSafe:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
'--------------------------------
' Fast row copy (no cell loops)
'--------------------------------
Private Sub CopyRowFast(wsSrc As Worksheet, wsDst As Worksheet, _
                        srcRow As Long, srcStartCol As Long, srcEndCol As Long, _
                        Optional dstRow As Long = 0, Optional dstCol As Long = 0)

    If dstRow = 0 Then dstRow = srcRow
    If dstCol = 0 Then dstCol = srcStartCol

    wsDst.Range(wsDst.Cells(dstRow, dstCol), wsDst.Cells(dstRow, dstCol + srcEndCol - srcStartCol)).value = _
        wsSrc.Range(wsSrc.Cells(srcRow, srcStartCol), wsSrc.Cells(srcRow, srcEndCol)).value

End Sub
' *** FIXED: Flexible date column search ***
Private Function FindDateColumnFlexible(ws As Worksheet, dateText As String) As Long
    Dim C As Range
    ' Try exact match first
    Set C = ws.Rows(1).Find(dateText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not C Is Nothing Then
        FindDateColumnFlexible = C.Column
        Exit Function
    End If
    ' Try partial match
    Set C = ws.Rows(1).Find(dateText, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not C Is Nothing Then FindDateColumnFlexible = C.Column
End Function

' Helper function to paste safely
Private Sub PasteIfSafe(targetCell As Range, value As Variant)
    If Not targetCell.HasFormula And _
       targetCell.Font.Color <> vbRed And _
       targetCell.Interior.Color <> RGB(255, 242, 204) And _
       Not targetCell.EntireRow.Hidden Then
        targetCell.value = value
    End If
End Sub
Private Function NormalizeContract(val As Variant, Optional daysRow As Long, Optional currentRow As Long) As String

    Dim txt As String
    Dim m As String, y As String
    
    If IsEmpty(val) Then Exit Function
    
    txt = Trim(CStr(val))
    If txt = "" Then Exit Function
    
    txt = LCase(txt)
    
    '-------------------------
    ' Remove spaces / dashes
    '-------------------------
    txt = Replace(txt, " ", "")
    
    '-------------------------
    ' Case 1: Quarter
    ' Handles Q226 / Q2-26 / Q2 26
    '-------------------------
    Dim qtxt As String
    qtxt = Replace(txt, "-", "")
    
    If qtxt Like "q#??" Then
        NormalizeContract = "Q" & Mid(qtxt, 2, 1) & "-" & Right(qtxt, 2)
        Exit Function
    End If
    
    '-------------------------
    ' Case 2: Month text
    ' Handles Dec26 / Dec-26 / Dec 26
    '-------------------------
    Dim clean As String
    clean = Replace(txt, "-", "")
    
    If clean Like "[a-z][a-z][a-z]##" Then
        m = Left(clean, 3)
        y = Right(clean, 2)
        NormalizeContract = UCase(m & "-" & y)
        Exit Function
    End If
    
    '-------------------------
    ' Case 3: Excel dates
    '-------------------------
    Dim d As Date
    If IsDate(val) Then

        d = CDate(val)

        If Day(d) = 1 Then
            ' Monthly contract
            If currentRow < daysRow Then
                NormalizeContract = UCase(Format(d, "mmm-yy"))
               
            Else
                NormalizeContract = UCase(Format(d, "dd-mmm-yy"))
            End If
                
        Else
            ' Daily contract
            NormalizeContract = UCase(Format(d, "dd-mmm-yy"))
        End If

        Exit Function
    End If
    
    '-------------------------
    ' Case 4: Leave others
    '-------------------------
    NormalizeContract = UCase(txt)

End Function
Private Function FindContractRow(ws As Worksheet, contractName As String) As Long

    Dim lastRow As Long, i As Long
    Dim cellValue As Variant
    Dim curveNorm As String, sheetNorm As String
    
    curveNorm = NormalizeContract(contractName)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, 1).value
        
        sheetNorm = NormalizeContract(cellValue)
        
        If curveNorm <> "" And sheetNorm <> "" Then
            If StrComp(curveNorm, sheetNorm, vbTextCompare) = 0 Then
                FindContractRow = i
                Exit Function
            End If
        End If
        
    Next i
    
    FindContractRow = 0

End Function
'============================================================
' Helper Function: Get Sheet by Name (Case-Insensitive)
'============================================================
Public Function GetSheetByNameInsensitive(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim cleanName As String
    cleanName = Replace(sheetName, Chr(160), " ") ' replace non-breaking spaces
    cleanName = Trim(cleanName)
    
    For Each ws In wb.Worksheets
        If StrComp(Trim(Replace(ws.Name, Chr(160), " ")), cleanName, vbTextCompare) = 0 Then
            Set GetSheetByNameInsensitive = ws
            Exit Function
        End If
    Next ws
End Function
'============================================================
' Helper Function: To check this few sheets have the word DAYS
'============================================================
Private Function IsDayRestrictedSheet(sheetName As String) As Boolean
    
    Dim arr As Variant
    Dim i As Long
    
    arr = Array("TBL", "CBL", "KBL", "TPK", "CPK", "KPK")
    
    For i = LBound(arr) To UBound(arr)
        If InStr(1, sheetName, arr(i), vbTextCompare) > 0 Then
            IsDayRestrictedSheet = True
            Exit Function
        End If
    Next i

End Function

