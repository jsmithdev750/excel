Option Explicit
'============================================================
' Main Import Macro
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

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    '--------------------------------
    ' Destination file pattern
    '--------------------------------
    todayYYMMDD = Format(Sheet1.Range("A3").Value, "yy.mm.dd")
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
    Set wsOrigin = GetSheetByNameInsensitive(wbOrigin, Sheet1.Range("A10").Value)
    Set wsDest = GetSheetByNameInsensitive(wbDest, Sheet1.Range("B10").Value)
    
    If wsOrigin Is Nothing Then
        MsgBox "Sheet 'OUTPUT' not found in origin workbook", vbCritical
        GoTo ExitSafe
    End If
    
    If wsDest Is Nothing Then
        MsgBox "Sheet 'CURVE' not found in destination workbook", vbCritical
        GoTo ExitSafe
    End If

    '--------------------------------
    ' Find TOKYO AREA
    '--------------------------------
    Set tokyoCell = wsOrigin.Cells.Find(Sheet1.Range("A7").Value, LookAt:=xlPart)

    If tokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found", vbCritical
        GoTo ExitSafe
    End If

    headerRow = tokyoCell.Row
    startCol = tokyoCell.mergeArea.Column

    '--------------------------------
    ' Find SPREADS
    '--------------------------------
    Set spreadsCell = wsOrigin.Cells.Find(Sheet1.Range("B7").Value, LookAt:=xlPart)

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
    For Each regionCell In regionCols

        regionStartCol = regionCell.mergeArea.Column
        regionEndCol = regionStartCol + regionCell.mergeArea.Columns.Count - 1

        wk1Row = headerRow + 2
        wk2Row = wk1Row + 7
        wk3Row = wk2Row + 7

        '--------------------------------
        ' WEEK CONTRACTS
        '--------------------------------
        CopyRowFast wsOrigin, wsDest, wk1Row, regionStartCol, regionEndCol
        CopyRowFast wsOrigin, wsDest, wk2Row, regionStartCol, regionEndCol
        CopyRowFast wsOrigin, wsDest, wk3Row, regionStartCol, regionEndCol

        '--------------------------------
        ' DAY CONTRACTS (AREA logic) + red font check
        '--------------------------------
        If InStr(1, regionCell.Value, "AREA", vbTextCompare) > 0 Then

            Dim col1 As Long, col2 As Long, col3 As Long
            Dim contractDate As Date
            Dim destDate As Date
            
            col1 = regionEndCol - 2
            col2 = regionEndCol - 1
            col3 = regionEndCol
            
            destDate = Sheet1.Range("A3").Value

            lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, col1).End(xlUp).Row

            ' Bulk copy
            wsDest.Range(wsDest.Cells(wk1Row, col1), wsDest.Cells(lastRow, col3)).Value = _
            wsOrigin.Range(wsOrigin.Cells(wk1Row, col1), wsOrigin.Cells(lastRow, col3)).Value
            
            ' Red font only on last col if date condition met (first wk contracts)
            For r = wk1Row To wk3Row
                If IsDate(wsDest.Cells(r, col2).Value) Then
                    contractDate = wsDest.Cells(r, col2).Value
                    If contractDate <= destDate Or contractDate = destDate + 1 Then
                        wsDest.Cells(r, col3).Font.Color = RGB(255, 0, 0)
                    End If
                End If
            Next r

        End If

        '--------------------------------
        ' REMAINING CONTRACTS
        '--------------------------------
        lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, regionStartCol).End(xlUp).Row

        If lastRow > wk3Row Then
            wsDest.Range(wsDest.Cells(wk3Row + 1, regionStartCol), _
                         wsDest.Cells(lastRow, regionEndCol)).Value = _
            wsOrigin.Range(wsOrigin.Cells(wk3Row + 1, regionStartCol), _
                           wsOrigin.Cells(lastRow, regionEndCol)).Value
        End If

    Next regionCell
    
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
                        rowNum As Long, startCol As Long, endCol As Long)
    wsDst.Range(wsDst.Cells(rowNum, startCol), wsDst.Cells(rowNum, endCol)).Value = _
    wsSrc.Range(wsSrc.Cells(rowNum, startCol), wsSrc.Cells(rowNum, endCol)).Value
End Sub

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

