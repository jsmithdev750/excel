Option Explicit

'============================================================
' Macro: Import New Japan Power Curve into MARKS sheet
'============================================================
Public Sub Import_New_Japan_Power_Curve()

    '-------------------------------
    ' Workbook & Worksheet variables
    '-------------------------------
    Dim wbOrigin As Workbook
    Dim wbDest As Workbook
    Dim wsCurve As Worksheet
    Dim wsMarks As Worksheet
    Dim wsWD As Worksheet
    Dim wb As Variant

    '-------------------------------
    ' File patterns and date
    '-------------------------------
    Dim originPattern As String
    Dim destPattern As String
    Dim todayYYMMDD As String
    Dim todayDDMMYY As Date

    '-------------------------------
    ' Header & column variables
    '-------------------------------
    Dim headerRow As Long
    Dim colMap As Object
    Dim headers As Variant
    Dim h As Variant, cell As Range
    Dim colContract As Long
    Dim stopWord As String
    Dim stopCol As Long
    Dim cStop As Long
    Dim pastePrevCol As Long, checkTestCol As Long
    Dim lastCol As Long
    Dim c As Range
    Dim F As Range
    Dim colIndex As Long

    '-------------------------------
    ' Data range & looping variables
    '-------------------------------
    Dim firstDataRow As Long, lastDataRow As Long
    Dim lastRow As Long
    Dim destRow As Long
    Dim pasteOrder As Variant
    Dim productMap As Object
    Dim key As Variant
    Dim markVal As Variant, changeVal As Variant
    Dim contractVal As Variant
    Dim prevVal As Variant, searchCol As Long
    Dim dayRow As Long, dayVal As Variant
    Dim weekRow As Long, weekVal As Variant
    Dim w As Long
    Dim dayCol As Long, weekCol As Long
    Dim r As Long

    '-------------------------------
    ' Day/Week specific variables
    '-------------------------------
    Dim hdrCell As Range, startRow As Long, rDay As Long
    Dim dayMark As Variant, dayContract As Variant
    Dim hdrCellP As Range, startRowP As Long, rDayP As Long
    Dim dayMarkP As Variant, dayContractP As Variant
    Dim wkRow As Long, wkContract As Variant, wkMark As Variant, wkChange As Variant
    Dim wkIndex As Long
    Dim wkParts() As String
    Dim wkNum As Long

    '-------------------------------
    ' File naming convention
    '-------------------------------
    todayDDMMYY = Sheet1.Range("A3").Value
    todayYYMMDD = Format(todayDDMMYY, "yy.mm.dd")

    '-------------------------------
    ' Workbooks patterns
    '-------------------------------
    originPattern = "*NEW CURVE_Simple Version Feb 2026*"
    destPattern = "*Vanir EEX Japan Power Curve_" & todayYYMMDD & " NEW FORMAT*"

    '-------------------------------
    ' Check if Origin Workbook is open
    '-------------------------------
    Set wbOrigin = Nothing
    For Each wb In Workbooks
        If wb.Name Like originPattern Then
            Set wbOrigin = wb
            Exit For
        End If
    Next wb
    If wbOrigin Is Nothing Then
        MsgBox "Origin workbook is not open!" & vbCrLf & "Expected pattern: " & originPattern, vbCritical
        Exit Sub
    End If

    '-------------------------------
    ' Check if Destination Workbook is open
    '-------------------------------
    Set wbDest = Nothing
    For Each wb In Workbooks
        If wb.Name Like destPattern Then
            Set wbDest = wb
            Exit For
        End If
    Next wb
    If wbDest Is Nothing Then
        MsgBox "Destination workbook is not open!" & vbCrLf & "Expected pattern: " & destPattern, vbCritical
        Exit Sub
    End If

    '-------------------------------
    ' Resolve Sheets
    '-------------------------------
    Set wsCurve = GetSheetByNameInsensitive(wbOrigin, "CURVE")
    Set wsMarks = GetSheetByNameInsensitive(wbDest, "MARKS")
    Set wsWD = GetSheetByNameInsensitive(wbOrigin, "WEEKS_DAYS")

    If wsCurve Is Nothing Or wsMarks Is Nothing Or wsWD Is Nothing Then
        MsgBox "One or more required sheets not found.", vbCritical
        Exit Sub
    End If
    
    '-------------------------------
    ' Determine Stop Column based on A7 (entire sheet)
    '-------------------------------
    stopWord = Trim(Sheet1.Range("A7").Value)
    stopCol = 0
    
    Dim stopCell As Range
    For Each stopCell In wsCurve.UsedRange
        If LCase(Trim(stopCell.Value)) = LCase(stopWord) Then
            stopCol = stopCell.Column
            Exit For
        End If
    Next stopCell
    
    If stopCol = 0 Then
        MsgBox "Stop word '" & stopWord & "' not found anywhere WB-Simple Version.", vbCritical
        Exit Sub
    End If

    '-------------------------------
    ' Locate Column Headers
    '-------------------------------
    Set colMap = CreateObject("Scripting.Dictionary")
    headers = Array("TBL", "CBL", "KBL", "TPK", "CPK", "KPK", "TOPK", "COPK", "KOPK")
    
    Dim found As Boolean
    For Each h In headers
        found = False
        For Each c In wsCurve.UsedRange
            ' Only consider columns before the stop column
            If c.Column < stopCol Then
                If LCase(Trim(c.Value)) = LCase(h) Then
                    colMap(h) = c.Column
                    ' Capture headerRow once (first header found)
                    If headerRow = 0 Then headerRow = c.Row
                    found = True
                    Exit For
                End If
            End If
        Next c
        
        If Not found Then
            MsgBox "Column header '" & h & "' not found before stop word '" & stopWord & "'", vbCritical
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
    firstDataRow = headerRow + 1
    lastDataRow = wsCurve.Cells(wsCurve.Rows.Count, colContract).End(xlUp).Row
    If lastDataRow < firstDataRow Then
        MsgBox "No contract data found.", vbCritical
        Exit Sub
    End If

    '-------------------------------
    ' Paste order & product mapping
    '-------------------------------
    pasteOrder = Array("TBL", "TPK", "TOPK", "CBL", "CPK", "COPK", "KBL", "KPK", "KOPK")
    Set productMap = CreateObject("Scripting.Dictionary")
    productMap("TBL") = "Tokyo Area Baseload"
    productMap("TPK") = "Tokyo Area Peakload"
    productMap("TOPK") = "Tokyo Area Off-Peak"
    productMap("CBL") = "Chubu Area Baseload"
    productMap("CPK") = "Chubu Area Peakload"
    productMap("COPK") = "Chubu Area Off-Peak"
    productMap("KBL") = "Kansai Area Baseload"
    productMap("KPK") = "Kansai Area Peakload"
    productMap("KOPK") = "Kansai Area Off-Peak"

    '-------------------------------
    ' Clear existing data in MARKS sheet
    '-------------------------------
    lastRow = wsMarks.Cells(wsMarks.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then wsMarks.Rows("2:" & lastRow).ClearContents
    destRow = 2

    '-------------------------------
    ' Set column formats
    '-------------------------------
    With wsMarks
        .Columns(1).NumberFormat = "dd mmm yyyy"
        .Columns(2).NumberFormat = "General"
        .Columns(3).NumberFormat = "mmm-yy"
        .Columns(4).NumberFormat = "0.00"
        .Columns(5).NumberFormat = "0.00"
    End With

    '-------------------------------
    ' Find Paste Previous Day & Check Test columns
    '-------------------------------
    pastePrevCol = 0
    checkTestCol = 0
    lastCol = wsCurve.Cells(headerRow, wsCurve.Columns.Count).End(xlToLeft).Column
    For cStop = 1 To lastCol
        Set cell = wsCurve.Cells(headerRow, cStop)
        If LCase(Trim(cell.Value)) = "paste previous day" Then pastePrevCol = cStop
        If LCase(Trim(cell.Value)) = "check test" Then checkTestCol = cStop
    Next cStop

    '-------------------------------
    ' Main Copy Loop: Day / Week / Numeric
    '-------------------------------
    For Each key In pasteOrder

        '------------- Day rows (Baseload) -------------
        If key = "TBL" Or key = "CBL" Or key = "KBL" Then

            Set hdrCell = wsWD.Columns("F").Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
            If hdrCell Is Nothing Then
                MsgBox key & " not found in WEEKS_DAYS column F", vbCritical
                Exit Sub
            End If
            startRow = hdrCell.Row + 1

            For rDay = startRow To startRow + 13
                dayMark = wsWD.Cells(rDay, "F").Value
                dayContract = wsWD.Cells(rDay, "E").Value

                wsMarks.Cells(destRow, 1).Value = todayDDMMYY
                wsMarks.Cells(destRow, 2).Value = productMap(key)
                If IsDate(dayContract) Then
                    wsMarks.Cells(destRow, 3).Value = "D" & Format(Day(dayContract), "00") & "-" & Format(dayContract, "Mmm-yy")
                Else
                    wsMarks.Cells(destRow, 3).Value = dayContract
                End If
                wsMarks.Cells(destRow, 4).Value = dayMark
                wsMarks.Cells(destRow, 5).Value = ""
                destRow = destRow + 1
            Next rDay
        End If

        '------------- Day rows (Peak) -------------
        If key = "TPK" Or key = "CPK" Or key = "KPK" Then

            Set hdrCellP = wsWD.Columns("H").Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
            If hdrCellP Is Nothing Then
                MsgBox key & " not found in WEEKS_DAYS column H", vbCritical
                Exit Sub
            End If
            startRowP = hdrCellP.Row + 1

            For rDayP = startRowP To startRowP + 13
                dayMarkP = wsWD.Cells(rDayP, "H").Value
                dayContractP = wsWD.Cells(rDayP, "E").Value

                wsMarks.Cells(destRow, 1).Value = todayDDMMYY
                wsMarks.Cells(destRow, 2).Value = productMap(key)
                If IsDate(dayContractP) Then
                    wsMarks.Cells(destRow, 3).Value = "D" & Format(Day(dayContractP), "00") & "-" & Format(dayContractP, "Mmm-yy")
                Else
                    wsMarks.Cells(destRow, 3).Value = dayContractP
                End If
                wsMarks.Cells(destRow, 4).Value = dayMarkP
                wsMarks.Cells(destRow, 5).Value = ""
                destRow = destRow + 1
            Next rDayP
        End If

        '------------- Week rows (Baseload) -------------
        If key = "TBL" Or key = "CBL" Or key = "KBL" Then
            For wkIndex = 0 To 2
                wkRow = startRow + (wkIndex * 7)
                wkContract = wsWD.Cells(wkRow, "B").Value
                wkMark = wsWD.Cells(wkRow, "C").Value
                wkChange = wsWD.Cells(wkRow + 1, "C").Value

                If Len(wkContract) > 0 Then
                    wkParts = Split(wkContract, "-")
                    If UBound(wkParts) = 1 Then
                        wkNum = CLng(Replace(wkParts(0), "Wk", ""))
                        wkContract = "Wk" & Format(wkNum, "00") & "-" & wkParts(1)
                    End If
                    wsMarks.Cells(destRow, 1).Value = todayDDMMYY
                    wsMarks.Cells(destRow, 2).Value = productMap(key)
                    wsMarks.Cells(destRow, 3).Value = wkContract
                    wsMarks.Cells(destRow, 4).Value = wkMark
                    wsMarks.Cells(destRow, 5).Value = wkChange
                    destRow = destRow + 1
                End If
            Next wkIndex
        End If

        '------------- Numeric rows -------------
        For r = firstDataRow To lastDataRow
            contractVal = wsCurve.Cells(r, colContract).Value
            markVal = wsCurve.Cells(r, colMap(key)).Value
            If key = "TOPK" Or key = "COPK" Or key = "KOPK" Then
                changeVal = ""
            Else
                changeVal = wsCurve.Cells(r, colMap(key) + 1).Value
            End If

            wsMarks.Cells(destRow, 1).Value = todayDDMMYY
            wsMarks.Cells(destRow, 2).Value = productMap(key)
            If IsDate(contractVal) Then
                wsMarks.Cells(destRow, 3).Value = DateSerial(Year(contractVal), Month(contractVal), 1)
            Else
                wsMarks.Cells(destRow, 3).Value = contractVal
            End If
            wsMarks.Cells(destRow, 4).Value = markVal
            wsMarks.Cells(destRow, 5).Value = changeVal
            destRow = destRow + 1
        Next r

    Next key

    '-------------------------------
    ' Format Header Row (A1:E1)
    '-------------------------------
    With wsMarks.Range("A1:E1")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    '-------------------------------
    ' Enable AutoFilter
    '-------------------------------
    If wsMarks.AutoFilterMode Then wsMarks.AutoFilterMode = False
    wsMarks.Range("A1:E1").AutoFilter

    If wsMarks.FilterMode Then wsMarks.ShowAllData

    '-------------------------------
    ' Save destination workbook
    '-------------------------------
    wbDest.Save

    MsgBox "Done", vbInformation

End Sub

'============================================================
' Helper Function: Get Sheet by Name (Case-Insensitive)
'============================================================
Public Function GetSheetByNameInsensitive(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim cleanName As String
    cleanName = Replace(sheetName, Chr(160), " ")
    cleanName = Trim(cleanName)
    
    For Each ws In wb.Worksheets
        If StrComp(Trim(Replace(ws.Name, Chr(160), " ")), cleanName, vbTextCompare) = 0 Then
            Set GetSheetByNameInsensitive = ws
            Exit Function
        End If
    Next ws
End Function
