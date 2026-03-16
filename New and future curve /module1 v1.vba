Option Explicit

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

    Set wsOrigin = wbOrigin.Sheets(1)
    Set wsDest = wbDest.Sheets(1)

    '--------------------------------
    ' Find TOKYO AREA
    '--------------------------------
    Set tokyoCell = wsOrigin.Cells.Find(Sheet1.Range("A7").Value, LookAt:=xlPart)

    If tokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found", vbCritical
        GoTo ExitSafe
    End If

    headerRow = tokyoCell.Row
    startCol = tokyoCell.MergeArea.Column

    '--------------------------------
    ' Find SPREADS
    '--------------------------------
    Set spreadsCell = wsOrigin.Cells.Find(Sheet1.Range("B7").Value, LookAt:=xlPart)

    If spreadsCell Is Nothing Then
        MsgBox "Spreads header not found", vbCritical
        GoTo ExitSafe
    End If

    endCol = spreadsCell.MergeArea.Columns(spreadsCell.MergeArea.Columns.Count).Column

    '--------------------------------
    ' Collect region headers
    '--------------------------------
    Set regionCols = New Collection

    c = startCol
    Do While c <= endCol

        If wsOrigin.Cells(headerRow, c).MergeCells Then
        
            regionCols.Add wsOrigin.Cells(headerRow, c)
            c = c + wsOrigin.Cells(headerRow, c).MergeArea.Columns.Count
            
        Else
        
            c = c + 1
            
        End If
        
    Loop

    '--------------------------------
    ' Process each region
    '--------------------------------
    For Each regionCell In regionCols

        regionStartCol = regionCell.MergeArea.Column
        regionEndCol = regionStartCol + regionCell.MergeArea.Columns.Count - 1

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
        ' DAY CONTRACTS (AREA logic)
        '--------------------------------
        If InStr(1, regionCell.Value, "AREA", vbTextCompare) > 0 Then

            Dim col1 As Long, col2 As Long, col3 As Long
            
            col1 = regionEndCol - 2
            col2 = regionEndCol - 1
            col3 = regionEndCol

            lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, col1).End(xlUp).Row

            wsDest.Range(wsDest.Cells(wk1Row, col1), wsDest.Cells(lastRow, col3)).Value = _
            wsOrigin.Range(wsOrigin.Cells(wk1Row, col1), wsOrigin.Cells(lastRow, col3)).Value

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
