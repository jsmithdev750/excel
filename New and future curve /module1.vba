Option Explicit
Public Sub Import_Old_Japan_Power_Curve()

    Dim wbOrigin As Workbook, wbDest As Workbook
    Dim wsOrigin As Worksheet, wsDest As Worksheet
    Dim tokyoCell As Range, spreadsCell As Range
    Dim headerRow As Long
    Dim startCol As Long, endCol As Long
    Dim regionCols As Collection
    Dim c As Long
    Dim r As Long
    Dim lastRow As Long
    
    Dim todayYYMMDD As String
    Dim destPattern As String
    
    '--------------------------------
    ' Date pattern for destination file
    '--------------------------------
    todayYYMMDD = Format(Sheet1.Range("A3").Value, "yy.mm.dd")
    destPattern = "*Vanir EEX Japan Power Curve_" & todayYYMMDD & "*"
    
    '--------------------------------
    ' Locate origin workbook
    '--------------------------------
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name Like "*NEW CURVE_OUTPUT*" Then
            Set wbOrigin = wb
            Exit For
        End If
    Next wb
    
    If wbOrigin Is Nothing Then
        MsgBox "Origin workbook not found", vbCritical
        Exit Sub
    End If
    
    '--------------------------------
    ' Locate destination workbook
    '--------------------------------
    Set wbDest = Nothing
    
    For Each wb In Workbooks
        If wb.Name Like destPattern _
        And Not wb.Name Like "*NEW FORMAT*" Then
            Set wbDest = wb
            Exit For
        End If
    Next wb
    
    If wbDest Is Nothing Then
        MsgBox "Destination workbook is not open!" & vbCrLf & _
               "Expected pattern: " & destPattern, vbCritical
        Exit Sub
    End If
    
    Set wsOrigin = wbOrigin.Sheets(1)
    Set wsDest = wbDest.Sheets(1)
    
    '--------------------------------
    ' Find TOKYO AREA
    '--------------------------------
    Set tokyoCell = wsOrigin.Cells.Find(Sheet1.Range("A7").Value, LookAt:=xlPart, MatchCase:=False)
    
    If tokyoCell Is Nothing Then
        MsgBox "TOKYO AREA not found", vbCritical
        Exit Sub
    End If
    
    headerRow = tokyoCell.Row
    startCol = tokyoCell.mergeArea.Column
    
    '--------------------------------
    ' Find SPREADS
    '--------------------------------
    Set spreadsCell = wsOrigin.Cells.Find(Sheet1.Range("B7").Value, LookAt:=xlPart, MatchCase:=False)
    
    If spreadsCell Is Nothing Then
        MsgBox "SPREADS not found", vbCritical
        Exit Sub
    End If
    
    endCol = spreadsCell.mergeArea.Columns(spreadsCell.mergeArea.Columns.Count).Column
    
    '--------------------------------
    ' Build region header list
    '--------------------------------
    
    Set regionCols = New Collection
    c = startCol
    
    Do While c <= endCol
    
        If wsOrigin.Cells(headerRow, c).MergeCells Then
        
            regionCols.Add wsOrigin.Cells(headerRow, c)
            c = wsOrigin.Cells(headerRow, c).mergeArea.Columns.Count + c
        
        Else
        
            c = c + 1
            
        End If
        
    Loop
    
    '--------------------------------
    ' Loop each region
    '--------------------------------
    
    Dim regionCell As Range
    Dim regionStartCol As Long
    Dim regionEndCol As Long
    
    For Each regionCell In regionCols
    
        regionStartCol = regionCell.mergeArea.Column
        regionEndCol = regionStartCol + regionCell.mergeArea.Columns.Count - 1
        
        Debug.Print "Processing region: " & regionCell.Value
        
        '--------------------------------
        ' WEEK CONTRACTS
        '--------------------------------
        
        Dim wk1Row As Long, wk2Row As Long, wk3Row As Long
        
        wk1Row = headerRow + 2
        wk2Row = wk1Row + 7
        wk3Row = wk2Row + 7
        
        CopyRowValues wsOrigin, wsDest, wk1Row, regionStartCol, regionEndCol
        CopyRowValues wsOrigin, wsDest, wk2Row, regionStartCol, regionEndCol
        CopyRowValues wsOrigin, wsDest, wk3Row, regionStartCol, regionEndCol
        
        '--------------------------------
        ' Day contract
        '--------------------------------
        
        If InStr(1, regionCell.Value, "AREA", vbTextCompare) > 0 Then
        
            Dim col1 As Long, col2 As Long, col3 As Long
            Dim pasteRow As Long
            
            col1 = regionEndCol - 2
            col2 = regionEndCol - 1
            col3 = regionEndCol
            
            pasteRow = wk1Row
            
            lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, col1).End(xlUp).Row
            
            For r = pasteRow To lastRow
            
                wsDest.Cells(r, col1).Value = wsOrigin.Cells(r, col1).Value
                wsDest.Cells(r, col2).Value = wsOrigin.Cells(r, col2).Value
                wsDest.Cells(r, col3).Value = wsOrigin.Cells(r, col3).Value
                
            Next r
            
        End If
        '--------------------------------
        ' REMAINING CONTRACTS
        '--------------------------------
        

        
        lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, regionStartCol).End(xlUp).Row
        
        For r = wk3Row + 1 To lastRow
        
            CopyRowValues wsOrigin, wsDest, r, regionStartCol, regionEndCol
            
        Next r
        
    Next regionCell
    
    MsgBox "Import completed"

End Sub


'--------------------------------
' Copy one row values only
'--------------------------------
Private Sub CopyRowValues(wsSrc As Worksheet, wsDst As Worksheet, _
                          rowNum As Long, startCol As Long, endCol As Long)

    Dim c As Long
    
    For c = startCol To endCol
    
        wsDst.Cells(rowNum, c).Value = wsSrc.Cells(rowNum, c).Value
        
    Next c

End Sub

