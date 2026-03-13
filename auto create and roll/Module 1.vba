Option Explicit

'============================================================
' Macro: Import Old Japan Fizz Curve (robust MTD + West Region search)
'============================================================
Public Sub Old_Japan_Fizz_Curve()

    Dim wbOriginOld As Workbook
    Dim wbDestOld As Workbook
    Dim wsCurveOld As Worksheet
    Dim wsCurveDestOld As Worksheet

    Dim originPatternOld As String
    Dim destPatternOld As String
    Dim todayYYMMDDOld As String
    Dim todayDDMMYY As String
    
    Dim wbOld As Workbook
    Dim fMTD As Range, fDestMTD As Range
    Dim fWest As Range
    Dim startRowOld As Long, lastRowOld As Long
    Dim startColOld As Long, endColOld As Long
    Dim rngCopyOld As Range
    Dim numRows As Long, numCols As Long
    Dim r As Long
    
    '--------------------------------------------------------
    ' Date
    '--------------------------------------------------------
    todayDDMMYY = Sheet1.Range("A3").Value
    todayYYMMDDOld = Format(todayDDMMYY, "yy.mm.dd")
    
    '--------------------------------------------------------
    ' Workbook patterns
    '--------------------------------------------------------
    originPatternOld = "*FIZZ CURVE SHEET - MASTER v1*"
    destPatternOld = "*Vanir Japan Power Curve_PHYSICAL_" & todayYYMMDDOld & "*.xls*"
    
    '--------------------------------------------------------
    ' Find origin workbook
    '--------------------------------------------------------
    For Each wbOld In Workbooks
        If wbOld.Name Like originPatternOld Then
            Set wbOriginOld = wbOld
            Exit For
        End If
    Next wbOld
    If wbOriginOld Is Nothing Then
        MsgBox "Origin workbook not open", vbCritical
        Exit Sub
    End If
    
    '--------------------------------------------------------
    ' Find destination workbook
    '--------------------------------------------------------
    For Each wbOld In Workbooks
        If wbOld.Name Like destPatternOld Then
            Set wbDestOld = wbOld
            Exit For
        End If
    Next wbOld
    If wbDestOld Is Nothing Then
        MsgBox "Destination workbook not open", vbCritical
        Exit Sub
    End If
    
    '--------------------------------------------------------
    ' Resolve sheets
    '--------------------------------------------------------
    Set wsCurveOld = GetSheetByNameInsensitive(wbOriginOld, "Base_Peak_Combined")
    Set wsCurveDestOld = GetSheetByNameInsensitive(wbDestOld, "Curve")
    If wsCurveOld Is Nothing Or wsCurveDestOld Is Nothing Then
        MsgBox "Required sheet missing", vbCritical
        Exit Sub
    End If
    
    '--------------------------------------------------------
    ' Find Base/Peak in origin sheet
    '--------------------------------------------------------
    Set fMTD = wsCurveOld.Cells.Find(What:="BASE/PEAK", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fMTD Is Nothing Then
        MsgBox "'BASE/PEAK' not found in origin sheet", vbCritical
        Exit Sub
    End If
    startRowOld = fMTD.Row + 1
    startColOld = fMTD.Column
    
    '--------------------------------------------------------
    ' Find last non-empty row below MTD
    '--------------------------------------------------------
    lastRowOld = startRowOld
    For r = startRowOld To wsCurveOld.Rows.Count
        If Application.WorksheetFunction.CountA(wsCurveOld.Rows(r)) = 0 Then Exit For
        lastRowOld = r
    Next r
    
    '--------------------------------------------------------
    ' Find West Region anywhere in origin sheet
    '--------------------------------------------------------
    Set fWest = wsCurveOld.Cells.Find(What:=Sheet1.Range("A10").Value, LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fWest Is Nothing Then
        MsgBox "'West Region' not found in origin sheet", vbCritical
        Exit Sub
    End If
    
    ' Get last column of West Region (merged aware)
    If fWest.MergeCells Then
        endColOld = fWest.mergeArea.Columns(fWest.mergeArea.Columns.Count).Column
    Else
        endColOld = fWest.Column
    End If
    
    '--------------------------------------------------------
    ' Compute rows and columns to copy
    '--------------------------------------------------------
    numRows = lastRowOld - startRowOld + 1
    numCols = endColOld - startColOld + 1
    
    If numRows <= 0 Or numCols <= 0 Then
        MsgBox "Invalid copy size. Check BASE/PEAK and West Region positions.", vbCritical
        Exit Sub
    End If
    
    '--------------------------------------------------------
    ' Set copy range in origin
    '--------------------------------------------------------
    Set rngCopyOld = wsCurveOld.Range(wsCurveOld.Cells(startRowOld, startColOld), _
                                      wsCurveOld.Cells(lastRowOld, endColOld))
    
    '--------------------------------------------------------
    ' Find MTD in destination sheet
    '--------------------------------------------------------
    Set fDestMTD = wsCurveDestOld.Cells.Find(What:="MtD", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fDestMTD Is Nothing Then
        MsgBox "'MTD' not found in destination sheet", vbCritical
        Exit Sub
    End If
    
    '--------------------------------------------------------
    ' Set destination range same size as copy range
    '--------------------------------------------------------
    Dim destRange As Range
    Set destRange = wsCurveDestOld.Range(fDestMTD, _
                                         wsCurveDestOld.Cells(fDestMTD.Row + numRows - 2, _
                                                              fDestMTD.Column + numCols - 1))
    
    '--------------------------------------------------------
    ' Paste values
    '--------------------------------------------------------
    destRange.Value = rngCopyOld.Value
    
    wbDestOld.Save
    MsgBox "Old Japan Fizz Curve pasted successfully", vbInformation

End Sub

'============================================================
' Helper: Get Sheet by Name (Case-Insensitive)
'============================================================
Public Function GetSheetByNameInsensitive(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim(ws.Name), Trim(sheetName), vbTextCompare) = 0 Then
            Set GetSheetByNameInsensitive = ws
            Exit Function
        End If
    Next ws
End Function

