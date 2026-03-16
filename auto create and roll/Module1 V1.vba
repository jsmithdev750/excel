Option Explicit

Public Sub Old_Japan_Fizz_Curve()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wbOriginOld As Workbook, wbDestOld As Workbook
    Dim wsCurveOld As Worksheet, wsCurveDestOld As Worksheet
    Dim originPatternOld As String, destPatternOld As String
    Dim todayYYMMDDOld As String, todayDate As Date  ' *** FIXED: Proper Date type ***
    Dim wbOld As Workbook
    Dim fMTD As Range, fDestMTD As Range, fWest As Range
    Dim startRowOld As Long, lastRowOld As Long, startColOld As Long, endColOld As Long
    Dim rngCopyOld As Range, numRows As Long, numCols As Long
    
    '--------------------------------------------------------
    ' SAFE Date handling - *** FIXED ***
    '--------------------------------------------------------
    Dim activeWs As Worksheet
    Set activeWs = ActiveSheet  ' Use ActiveSheet instead of hardcoded Sheet1
    If activeWs.Range("A3").value = "" Then
        MsgBox "A3 date is empty!", vbCritical
        GoTo SafeExit
    End If
    todayDate = CDate(activeWs.Range("A3").value)
    todayYYMMDDOld = Format(todayDate, "yy.mm.dd")
    Debug.Print "Today date: " & todayDate & " | Format: " & todayYYMMDDOld
    
    '--------------------------------------------------------
    ' Workbook patterns
    '--------------------------------------------------------
    originPatternOld = "*FIZZ CURVE SHEET - MASTER v1*"
    destPatternOld = "*Vanir Japan Power Curve_PHYSICAL_" & todayYYMMDDOld & "*.xls*"
    
    ' Find workbooks...
    Set wbOriginOld = FindWorkbook(originPatternOld)
    If wbOriginOld Is Nothing Then GoTo SafeExit
    
    Set wbDestOld = FindWorkbook(destPatternOld)
    If wbDestOld Is Nothing Then GoTo SafeExit
    
    ' Find sheets...
    Set wsCurveOld = GetSheetByNameInsensitive(wbOriginOld, "Base_Peak_Combined")
    Set wsCurveDestOld = GetSheetByNameInsensitive(wbDestOld, "Curve")
    If wsCurveOld Is Nothing Or wsCurveDestOld Is Nothing Then GoTo SafeExit
    
    ' Find ranges...
    Set fMTD = wsCurveOld.Cells.Find("BASE/PEAK", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fMTD Is Nothing Then GoTo SafeExit
    
    startRowOld = fMTD.Row + 1
    startColOld = fMTD.Column
    lastRowOld = wsCurveOld.Cells(wsCurveOld.Rows.Count, startColOld).End(xlUp).Row
    If lastRowOld < startRowOld Then lastRowOld = startRowOld
    
    ' *** FIXED: Safe A10 access ***
    Dim westRegionText As String
    westRegionText = GetCellValueSafe(activeWs, "A10", "West Region")
    
    Set fWest = wsCurveOld.Cells.Find(westRegionText, LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fWest Is Nothing Then GoTo SafeExit
    
    ' Column calculation
    endColOld = IIf(fWest.MergeCells, fWest.mergeArea.Column + fWest.mergeArea.Columns.Count - 1, fWest.Column)
    
    ' Copy data
    numRows = lastRowOld - startRowOld + 1
    numCols = endColOld - startColOld + 1
    Set rngCopyOld = wsCurveOld.Range(wsCurveOld.Cells(startRowOld, startColOld), wsCurveOld.Cells(lastRowOld, endColOld))
    
    Set fDestMTD = wsCurveDestOld.Cells.Find("MtD", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fDestMTD Is Nothing Then GoTo SafeExit
    
    ' PASTE - *** FIXED paste range ***
    wsCurveDestOld.Range(fDestMTD, fDestMTD.Offset(numRows - 1, numCols - 1)).value = rngCopyOld.value
    
    ' Process region sheets
    Call ProcessCurveData(wbDestOld, wsCurveDestOld, fDestMTD, numRows, todayDate)
    
    wbDestOld.Save
    MsgBox "SUCCESS! Old Japan Fizz Curve pasted.", vbInformation
    GoTo SafeExit

ErrorHandler:
    MsgBox "ERROR " & Err.Number & ": " & Err.Description, vbCritical
    Debug.Print "ERROR: " & Err.Description & " at " & Now

SafeExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
'============================================================
' *** FIXED: ProcessCurveData with error handling ***
'============================================================
Private Sub ProcessCurveData(wbDest As Workbook, wsCurveDest As Worksheet, fDestMTD As Range, numRows As Long, todayDate As Date)
    On Error GoTo ProcError
    
    Dim regions As Variant: regions = Array("Tokyo", "Chubu", "Kansai", "Hokkaido", "Tohoku", "Hokuriku", "Chugoku", "Shikoku", "Kyushu")
    Dim regionColDict As Object: Set regionColDict = CreateObject("Scripting.Dictionary")
    
    Dim region As Variant, regionCell As Range
    Dim todayDateFormat As String: todayDateFormat = Format(todayDate, "dd-mmm-yy")
    
    ' *** STEP 1: Build region column dictionary ***
    For Each region In regions
        Set regionCell = FindRegionCellAnywhere(wsCurveDest, CStr(region))
        If Not regionCell Is Nothing Then
            regionColDict(CStr(region)) = regionCell.Column
            Debug.Print "Region " & region & " at column " & regionCell.Column
        End If
    Next region
    
    ' *** STEP 2: Process EACH region individually ***
    Dim wsBase As Worksheet, wsPeak As Worksheet
    Dim dateColBase As Long, dateColPeak As Long
    Dim contractName As String, curveBaseVal As Variant, curvePeakVal As Variant
    Dim baseRow As Long, peakRow As Long, r As Long
    
    For Each region In regions
        If Not regionColDict.Exists(CStr(region)) Then GoTo NextRegion
        
        ' Get sheets
        On Error Resume Next
        Set wsBase = wbDest.Worksheets(CStr(region) & " Base")
        Set wsPeak = wbDest.Worksheets(CStr(region) & " Peak")
        On Error GoTo ProcError
        
        If wsBase Is Nothing Or wsPeak Is Nothing Then
            Debug.Print "Skipping " & region & " - sheets missing"
            GoTo NextRegion
        End If
        
        ' Find date columns
        dateColBase = FindDateColumnFlexible(wsBase, todayDateFormat)
        dateColPeak = FindDateColumnFlexible(wsPeak, todayDateFormat)
        If dateColBase = 0 Or dateColPeak = 0 Then GoTo NextRegion
        
        Debug.Print region & " | Date cols: Base=" & dateColBase & ", Peak=" & dateColPeak
        
        ' *** STEP 3: For EACH Curve row, find matching contract IN BOTH sheets ***
        For r = fDestMTD.Row + 1 To fDestMTD.Row + numRows - 1
            contractName = Trim(CStr(wsCurveDest.Cells(r, fDestMTD.Column).value))
            If Len(contractName) = 0 Then GoTo NextContract
            
            ' Get values from Curve sheet
            Dim regionCol As Long: regionCol = regionColDict(CStr(region))
            curveBaseVal = wsCurveDest.Cells(r, regionCol).value
            curvePeakVal = wsCurveDest.Cells(r, regionCol + 1).value
            
            ' *** FIND CONTRACT IN BASE SHEET ***
            baseRow = FindContractRow(wsBase, contractName)
            If baseRow > 0 Then
                Debug.Print "  " & region & " Base: " & contractName & " @ row " & baseRow & " = " & curveBaseVal
                PasteIfSafe wsBase.Cells(baseRow, dateColBase), curveBaseVal
            End If
            
            ' *** FIND CONTRACT IN PEAK SHEET (SEPARATE SEARCH!) ***
            peakRow = FindContractRow(wsPeak, contractName)
            If peakRow > 0 Then
                Debug.Print "  " & region & " Peak: " & contractName & " @ row " & peakRow & " = " & curvePeakVal
                PasteIfSafe wsPeak.Cells(peakRow, dateColPeak), curvePeakVal
            End If
            
NextContract:       Next r
        
NextRegion:     Next region
        
    Debug.Print "? ALL regions processed!"
    Exit Sub
    
ProcError:
    Debug.Print "ERROR: " & Err.Description
End Sub
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
Private Function GetCellValueSafe(ws As Worksheet, cellAddr As String, defaultVal As String) As String
    On Error Resume Next
    GetCellValueSafe = CStr(ws.Range(cellAddr).value)
    If Err.Number <> 0 Then GetCellValueSafe = defaultVal
    On Error GoTo 0
End Function
Private Sub PasteIfSafe(targetCell As Range, value As Variant)
    If Not targetCell.HasFormula And _
       targetCell.Font.Color <> vbRed And _
       targetCell.Interior.Color <> RGB(255, 242, 204) And _
       Not targetCell.EntireRow.Hidden Then
        targetCell.value = value
    End If
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
' Keep existing helper functions...
Public Function GetSheetByNameInsensitive(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim(ws.Name), Trim(sheetName), vbTextCompare) = 0 Then
            Set GetSheetByNameInsensitive = ws
            Exit Function
        End If
    Next ws
End Function
Public Function FindRegionCellAnywhere(ws As Worksheet, regionName As String) As Range
    Dim cell As Range, checkCell As Range, cellValue As String
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            For Each checkCell In cell.mergeArea.Cells
                cellValue = Trim(CStr(checkCell.value))
                If StrComp(cellValue, regionName, vbTextCompare) = 0 Then
                    Set FindRegionCellAnywhere = cell.mergeArea.Cells(1, 1)
                    Exit Function
                End If
            Next
        Else
            cellValue = Trim(CStr(cell.value))
            If StrComp(cellValue, regionName, vbTextCompare) = 0 Then
                Set FindRegionCellAnywhere = cell
                Exit Function
            End If
        End If
    Next
    Set FindRegionCellAnywhere = Nothing
End Function
'============================================================
' *** NEW HELPER FUNCTIONS ***
'============================================================
Private Function FindWorkbook(pattern As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name Like pattern Then
            Set FindWorkbook = wb
            Debug.Print "Found: " & wb.Name
            Exit Function
        End If
    Next wb
    MsgBox "Workbook not found: " & pattern, vbCritical
End Function


