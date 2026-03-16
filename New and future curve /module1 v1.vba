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
    ' Find TOKYO AREA in origin and destination (dynamic)
    '--------------------------------
    Set tokyoCell = wsOrigin.Cells.Find(Sheet1.Range("A7").Value, LookAt:=xlPart)
    If tokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found in origin sheet", vbCritical
        GoTo ExitSafe
    End If
    headerRow = tokyoCell.Row
    startCol = tokyoCell.mergeArea.Column

    Set destTokyoCell = wsDest.Cells.Find(Sheet1.Range("A7").Value, LookAt:=xlPart)
    If destTokyoCell Is Nothing Then
        MsgBox "Tokyo Area header not found in destination sheet", vbCritical
        GoTo ExitSafe
    End If
    destHeaderRow = destTokyoCell.Row
    destStartCol = destTokyoCell.mergeArea.Column

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
        rowOffset = wk1Row - headerRow
        colOffset = regionStartCol - startCol
        CopyRowFast wsOrigin, wsDest, wk1Row, regionStartCol, regionEndCol, destHeaderRow + rowOffset, destStartCol + colOffset
        rowOffset = wk2Row - headerRow
        CopyRowFast wsOrigin, wsDest, wk2Row, regionStartCol, regionEndCol, destHeaderRow + rowOffset, destStartCol + colOffset
        rowOffset = wk3Row - headerRow
        CopyRowFast wsOrigin, wsDest, wk3Row, regionStartCol, regionEndCol, destHeaderRow + rowOffset, destStartCol + colOffset

        '--------------------------------
        ' DAY CONTRACTS (AREA logic) + red font check
        '--------------------------------
        If InStr(1, regionCell.Value, "AREA", vbTextCompare) > 0 Then

            Dim col1 As Long, col2 As Long, col3 As Long
            Dim contractDate As Date
            Dim destDate As Date
            Dim destRowStart As Long, destColStart As Long
            
            col1 = regionEndCol - 2
            col2 = regionEndCol - 1
            col3 = regionEndCol
            destDate = Sheet1.Range("A3").Value

            lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, col1).End(xlUp).Row

            ' Dynamic paste destination
            destRowStart = destHeaderRow + (wk1Row - headerRow)
            destColStart = destStartCol + (col1 - startCol)

            wsDest.Range(wsDest.Cells(destRowStart, destColStart), _
                         wsDest.Cells(destRowStart + (lastRow - wk1Row), destColStart + (col3 - col1))).Value = _
                wsOrigin.Range(wsOrigin.Cells(wk1Row, col1), wsOrigin.Cells(lastRow, col3)).Value

            ' Red font only on last col if date condition met
            For r = 0 To wk3Row - wk1Row
                If IsDate(wsDest.Cells(destRowStart + r, destColStart + 1).Value) Then
                    contractDate = wsDest.Cells(destRowStart + r, destColStart + 1).Value
                    If contractDate <= destDate Or contractDate = destDate + 1 Then
                        wsDest.Cells(destRowStart + r, destColStart + 2).Font.Color = RGB(255, 0, 0)
                    End If
                End If
            Next r

        End If
        
    '--------------------------------
    ' Copy charts for AREA regions
    '--------------------------------
Dim ch As ChartObject
Dim chartLeft As Double, chartTop As Double
Dim wk1Top As Double, wk2Top As Double

' Get the top positions of the first and second week rows in destination
wk1Top = wsDest.Rows(destHeaderRow + (wk1Row - headerRow)).Top
wk2Top = wsDest.Rows(destHeaderRow + (wk2Row - headerRow)).Top

For Each ch In wsOrigin.ChartObjects
    ' Check if chart overlaps the AREA region columns
    If ch.Left >= wsOrigin.Cells(headerRow, regionStartCol).Left And _
       ch.Left + ch.Width <= wsOrigin.Cells(headerRow, regionEndCol).Left + wsOrigin.Cells(headerRow, regionEndCol).Width Then

        ' Determine horizontal position in destination
        chartLeft = wsDest.Cells(destHeaderRow, destStartCol + (regionStartCol - startCol)).Left

        ' Determine vertical position based on chart location in origin
        If ch.Top < wsOrigin.Rows(wk2Row).Top Then
            ' Chart below first week
            chartTop = wk1Top + wsDest.Rows(destHeaderRow + (wk1Row - headerRow)).Height
        Else
            ' Chart below second week
            chartTop = wk2Top + wsDest.Rows(destHeaderRow + (wk2Row - headerRow)).Height
        End If

        ' Copy chart as picture
        ch.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        ' Paste in destination
        wsDest.Paste

        ' Move pasted image
        With wsDest.Shapes(wsDest.Shapes.Count)
            .Left = chartLeft
            .Top = chartTop
            .LockAspectRatio = msoTrue
        End With
    End If
Next ch

        '--------------------------------
        ' REMAINING CONTRACTS
        '--------------------------------
        lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, regionStartCol).End(xlUp).Row
        If lastRow > wk3Row Then
            rowOffset = wk3Row + 1 - headerRow
            colOffset = regionStartCol - startCol
            wsDest.Range(wsDest.Cells(destHeaderRow + rowOffset, destStartCol + colOffset), _
                         wsDest.Cells(destHeaderRow + rowOffset + (lastRow - wk3Row - 1), destStartCol + colOffset + (regionEndCol - regionStartCol))).Value = _
                wsOrigin.Range(wsOrigin.Cells(wk3Row + 1, regionStartCol), wsOrigin.Cells(lastRow, regionEndCol)).Value
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
                        srcRow As Long, srcStartCol As Long, srcEndCol As Long, _
                        Optional dstRow As Long = 0, Optional dstCol As Long = 0)

    If dstRow = 0 Then dstRow = srcRow
    If dstCol = 0 Then dstCol = srcStartCol

    wsDst.Range(wsDst.Cells(dstRow, dstCol), wsDst.Cells(dstRow, dstCol + srcEndCol - srcStartCol)).Value = _
        wsSrc.Range(wsSrc.Cells(srcRow, srcStartCol), wsSrc.Cells(srcRow, srcEndCol)).Value

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

