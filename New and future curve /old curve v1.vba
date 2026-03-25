Option Explicit
'============================================================
' Main Import Macro (with dynamic destination offsets)
'============================================================
Public Sub Import_Old_Japan_Power_Curve()

    Dim wbOrigin As Workbook, wbDest As Workbook
    Dim wsOrigin As Worksheet, wsDest As Worksheet
    Dim tokyoCell As Range, spreadsCell As Range
    Dim headerRow As Long
    Dim startCol As Long, endCol As Long
    Dim regionCols As Collection
    Dim regionCell As Range
    Dim regionStartCol As Long, regionEndCol As Long
    Dim wk1Row As Long, wk2Row As Long, wk3Row As Long
    Dim c As Long, r As Long, lastRow As Long
    
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

    '--------------------------------
    ' Destination file pattern
    '--------------------------------
    todayYYMMDD = Format(Sheet1.Range("A3").value, "yy.mm.dd")
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
    Set wsOrigin = GetSheetByNameInsensitive(wbOrigin, Sheet1.Range("A14").value)
    Set wsDest = GetSheetByNameInsensitive(wbDest, Sheet1.Range("B14").value)
    Set wsInput = GetSheetByNameInsensitive(wbOrigin, "INPUT")
    
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
    Set tokyoCell = wsOrigin.Cells.Find(Sheet1.Range("A11").value, LookAt:=xlPart)
    If tokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found in origin sheet", vbCritical
        GoTo ExitSafe
    End If
    headerRow = tokyoCell.Row
    startCol = tokyoCell.mergeArea.Column

    Set destTokyoCell = wsDest.Cells.Find(Sheet1.Range("A11").value, LookAt:=xlPart)
    If destTokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found in destination sheet", vbCritical
        GoTo ExitSafe
    End If
    destHeaderRow = destTokyoCell.Row
    destStartCol = destTokyoCell.mergeArea.Column

    '--------------------------------
    ' Find SPREADS
    '--------------------------------
    Set spreadsCell = wsOrigin.Cells.Find(Sheet1.Range("B11").value, LookAt:=xlPart)
    If spreadsCell Is Nothing Then
        MsgBox "Spreads header not found", vbCritical
        GoTo ExitSafe
    End If
    endCol = spreadsCell.mergeArea.Columns(spreadsCell.mergeArea.Columns.Count).Column

    '--------------------------------
    ' Collect region headers
    '--------------------------------
    Set regionCols = New Collection
    c = startCol
    Do While c <= endCol
        If wsOrigin.Cells(headerRow, c).MergeCells Then
            regionCols.Add wsOrigin.Cells(headerRow, c)
            c = c + wsOrigin.Cells(headerRow, c).mergeArea.Columns.Count
        Else
            c = c + 1
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
            destDate = Sheet1.Range("A3").value

            lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, col1).End(xlUp).Row

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
        lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, regionStartCol).End(xlUp).Row
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
                If headerRowInput = 0 Then headerRowInput = d.Row
                    found = True
                Exit For
            End If
        Next d
        
        If Not found Then
            MsgBox "Column header '" & h & "' not found '", vbCritical
            Exit Sub
        End If
    Next h

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
    lastDataRow = wsInput.Cells(wsInput.Rows.Count, colContract).End(xlUp).Row
    If lastDataRow < firstDataRow Then
        MsgBox "No contract data found.", vbCritical
        Exit Sub
    End If
    
    '-------------------------------
    ' Display results in MsgBox
    '-------------------------------
    Dim msg As String
    
    msg = "Header column mapping:" & vbCrLf
    For Each key In colMap.Keys
        msg = msg & key & " -> Column " & colMap(key) & vbCrLf
    Next key
    
    msg = msg & "Contract column -> Column " & colContract & vbCrLf
    msg = msg & "Data rows: " & firstDataRow & " to " & lastDataRow
    
    MsgBox msg, vbInformation, "INPUT Sheet Mapping"

'--------------------------------
' Loop through "Hist" sheets and process the data
'--------------------------------


For Each histSheet In wbDest.Worksheets
    
    If InStr(1, histSheet.Name, "Hist", vbTextCompare) = 1 Then

        histDateColumn = FindDateColumnFlexible(histSheet, Format(Sheet1.Range("A3").value, "dd-mmm-yy"))

        If histDateColumn > 0 Then

            lastHistRow = histSheet.Cells(histSheet.Rows.Count, 1).End(xlUp).Row

            For r = 2 To lastHistRow
                
                contract = histSheet.Cells(r, 1).value
                
                If contract <> "" Then
                    
                    Dim normalizedContract As String
                    normalizedContract = NormalizeContract(contract)
                    
                    ' ?? Find in INPUT sheet (NOT origin/output)
                    Dim i As Long
                    Dim foundRow As Long
                    foundRow = 0
                    
                    Dim contractDict As Object
                    Set contractDict = CreateObject("Scripting.Dictionary")
                    
                    For i = firstDataRow To lastDataRow
                        contractDict(NormalizeContract(wsInput.Cells(i, colContract).value)) = i
                    Next i
                    
                    If contractDict.exists(normalizedContract) Then
                        foundRow = contractDict(normalizedContract)
                    Else
                        foundRow = 0
                    End If
                    
                    If foundRow > 0 Then
                        
                        ' Example: pulling TBL (you can change to other columns)

                        
                        colKey = ""
                        
                        For Each key In colMap.Keys
                            If InStr(1, histSheet.Name, key, vbTextCompare) > 0 Then
                                colKey = key
                                Exit For
                            End If
                        Next key
                        
                    If colKey <> "" Then
                        valToPaste = wsInput.Cells(foundRow, colMap(colKey)).value
                        Set targetCell = histSheet.Cells(r, histDateColumn)
                        PasteIfSafe targetCell, valToPaste
                    End If
                                            
                    Else
                       ' MsgBox "Contract not found in INPUT sheet: " & contract, vbExclamation
                    End If
                
                End If
                
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
    Dim c As Range
    ' Try exact match first
    Set c = ws.Rows(1).Find(dateText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not c Is Nothing Then
        FindDateColumnFlexible = c.Column
        Exit Function
    End If
    ' Try partial match
    Set c = ws.Rows(1).Find(dateText, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not c Is Nothing Then FindDateColumnFlexible = c.Column
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
Private Function NormalizeContract(val As Variant) As String

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
    ' Case 3: Real Excel date
    '-------------------------
    On Error Resume Next
    Dim d As Date
    d = CDate(val)
    On Error GoTo 0
    
    If d <> 0 Then
        NormalizeContract = UCase(Format(d, "mmm-yy"))
        Exit Function
    End If
    
    '-------------------------
    ' Case 4: Leave others
    '-------------------------
    NormalizeContract = UCase(txt)

End Function
' *** NEW: Find contract row in Column A ***
Private Function FindContractRow(ws As Worksheet, contractName As String) As Long

    Dim lastRow As Long, i As Long
    Dim cellValue As Variant
    Dim curveNorm As String, sheetNorm As String
    
    curveNorm = NormalizeContract(contractName)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
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




