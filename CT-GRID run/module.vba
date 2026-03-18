Option Explicit

'=============================
' User-defined type to hold term anchor info
'=============================
Type SAnchor
    kind As String    ' "week", "day", "month", "monthSpan", "quarter", "unknown"
    mon As Variant    ' 0-based month index (0=Jan)
    day As Variant    ' day number
    yy As Variant     ' year string (2-digit)
    hasMon As Boolean
    hasDay As Boolean
    hasYY As Boolean
End Type

'=============================
' Record type for each transformed row
'=============================
Type ContractRecord
    FullText As String   ' full transformed text (column C)
    RegionKey As String  ' parts(0) lowercased & trimmed
End Type

'=============================
' Main procedure
'=============================
Public Sub MapContractSeasons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long, j As Long
    Dim text As String, parts() As String, contract As String, legs() As String
    Dim outLegs() As String, transformedContract As String, newText As String
    Dim recs() As ContractRecord
    
    '=============================
    ' Step 1: Transform all contracts and store in array
    '=============================
    ReDim recs(1 To lastRow - 1)  ' assume header in row 1
    Dim recIndex As Long
    recIndex = 1
    
    For i = 2 To lastRow
        text = ws.Cells(i, 1).Value
        If Trim(text) = "" Then GoTo NextRow
        
        parts = Split(text, " ")
        If UBound(parts) < 1 Then GoTo NextRow
        
        ' Transform contract legs
        contract = parts(1)
        legs = Split(contract, "/")
        ReDim outLegs(LBound(legs) To UBound(legs))
        For j = LBound(legs) To UBound(legs)
            outLegs(j) = TransformTerm(Trim(legs(j)))
        Next j
        
        transformedContract = Join(outLegs, "/")
        ws.Cells(i, 2).Value = transformedContract  ' optional: store in col B
        
        ' Rebuild full text
        newText = parts(0) & " " & transformedContract
        If UBound(parts) > 1 Then
            For j = 2 To UBound(parts)
                newText = newText & " " & parts(j)
            Next j
        End If
        
        ' Store in array
        recs(recIndex).FullText = newText
        recs(recIndex).RegionKey = LCase(Trim(parts(0)))
        recIndex = recIndex + 1
        
NextRow:
    Next i
    
    ' Resize array to actual number of records
    If recIndex > 1 Then
        ReDim Preserve recs(1 To recIndex - 1)
    Else
        MsgBox "No valid data found!", vbExclamation
        Exit Sub
    End If
    
    '=============================
    ' Step 2: Read Region Order from workbook (case-insensitive, trim spaces)
    '=============================
    Dim orderCell As Range
    Dim found As Boolean
    found = False
    
    For Each orderCell In ws.UsedRange
        If LCase(replace(Trim(orderCell.Value), Chr(160), "")) = "region order by" Then
            found = True
            Exit For
        End If
    Next orderCell
    
    If Not found Then
        MsgBox "'Region Order BY' not found in Sheet1", vbCritical
        Exit Sub
    End If
    
    ' Region order list below the found cell
    Dim rngOrder As Range
    Set rngOrder = ws.Range(orderCell.Offset(1, 0), ws.Cells(ws.Rows.Count, orderCell.Column).End(xlUp))
    
    ' Build order dictionary
    Dim orderDict As Object
    Set orderDict = CreateObject("Scripting.Dictionary")
    Dim idx As Long, cell As Range
    idx = 1
    For Each cell In rngOrder
        If Trim(cell.Value) <> "" Then
            orderDict(LCase(Trim(cell.Value))) = idx
            idx = idx + 1
        End If
    Next cell
    
    '=============================
    ' Step 3: Sort array by RegionKey based on orderDict
    '=============================
    Call SortContractRecordsByRegion(recs, orderDict)
    
    '=============================
    ' Step 4: Write back sorted array to column C
    '=============================
    For i = 1 To UBound(recs)
        ws.Cells(i + 1, 3).Value = recs(i).FullText
    Next i
    
    MsgBox "Contracts transformed, rebuilt, and sorted by region!", vbInformation
End Sub

'=============================
' Sorting routine (bubble sort for simplicity)
' Can be replaced with more efficient sort if needed
'=============================
Sub SortContractRecordsByRegion(ByRef recs() As ContractRecord, ByRef orderDict As Object)
    Dim i As Long, j As Long
    Dim temp As ContractRecord
    Dim key1 As Long, key2 As Long
    
    For i = LBound(recs) To UBound(recs) - 1
        For j = i + 1 To UBound(recs)
            key1 = GetOrderIndex(recs(i).RegionKey, orderDict)
            key2 = GetOrderIndex(recs(j).RegionKey, orderDict)
            If key1 > key2 Then
                temp = recs(i)
                recs(i) = recs(j)
                recs(j) = temp
            End If
        Next j
    Next i
End Sub

'=============================
' Return index for sorting: if not in orderDict, assign large number (end of list)
'=============================
Function GetOrderIndex(ByVal key As String, ByRef orderDict As Object) As Long
    If orderDict.exists(key) Then
        GetOrderIndex = orderDict(key)
    Else
        GetOrderIndex = 9999
    End If
End Function
'=============================
' Transform a single term or fallback
'=============================
Function TransformTerm(term As String) As String
    Dim parts() As String
    Dim i As Long
    Dim fallbackYY As String
    Dim outParts() As String
    
    parts = Split(term, "/")
    
    ' Single part
    If UBound(parts) = 0 Then
        TransformTerm = TransformTermSingle(parts(0))
        Exit Function
    End If
    
    ' Determine fallback year from rightmost part
    fallbackYY = ""
    For i = UBound(parts) To 0 Step -1
        If Right(parts(i), 3) Like "-##" Then
            fallbackYY = Right(parts(i), 2)
            Exit For
        End If
    Next i
    
    ' Normalize each part with fallback
    ReDim outParts(LBound(parts) To UBound(parts))
    For i = LBound(parts) To UBound(parts)
        If i = LBound(parts) Then
            outParts(i) = MaterializeLeft(parts(i), fallbackYY)
        Else
            ' Append fallback year if missing
            If Right(parts(i), 3) Like "-##" Or fallbackYY = "" Then
                outParts(i) = parts(i)
            Else
                outParts(i) = parts(i) & "-" & fallbackYY
            End If
        End If
        outParts(i) = TransformTermSingle(outParts(i))
    Next i
    
    ' Remove year from first part
    If Right(outParts(LBound(outParts)), 3) Like "-##" Then
        outParts(LBound(outParts)) = Left(outParts(LBound(outParts)), Len(outParts(LBound(outParts))) - 3)
    End If
    
    TransformTerm = Join(outParts, "/")
End Function

'=============================
' Append year to left part if missing
'=============================
Function MaterializeLeft(part As String, fallbackYY As String) As String
    If Right(part, 3) Like "-##" Then
        MaterializeLeft = part
    ElseIf fallbackYY <> "" Then
        MaterializeLeft = part & "-" & fallbackYY
    Else
        MaterializeLeft = part
    End If
End Function

'=============================
' Transform a single leg (JS TransformLeg)
'=============================
Function TransformTermSingle(ByVal leg As String) As String
    Dim re As Object, matches As Object
    Dim d1 As Long, d2 As Long, m1 As Long, m2 As Long, y1 As Long, y2 As Long
    Dim lowercapLeg As String
    
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False
    
    leg = Trim(leg)
    
    ' Pattern 1: d1-d2-Mon-yy
    re.pattern = "^(\d{1,2})-(\d{1,2})-([A-Za-z]{3,9})-(\d{2})$"
    If re.Test(leg) Then
        Set matches = re.Execute(leg)(0)
        d1 = CLng(matches.SubMatches(0))
        d2 = CLng(matches.SubMatches(1))
        m2 = MonthIndex(matches.SubMatches(2))
        y2 = ToYYYY(matches.SubMatches(3))
        
        If d1 <= d2 Then
            m1 = m2
            y1 = y2
        Else
            m1 = (m2 + 11) Mod 12
            If m2 = 0 Then y1 = y2 - 1 Else y1 = y2
        End If
        
        TransformTermSingle = WeekCodesFromRangeInclusive(d1, m1, y1, d2, m2, y2)
        Exit Function
    End If
    
    ' Pattern 2: d1-Mon1-d2-Mon2-yy
    re.pattern = "^(\d{1,2})-([A-Za-z]{3,9})-(\d{1,2})-([A-Za-z]{3,9})-(\d{2})$"
    If re.Test(leg) Then
        Set matches = re.Execute(leg)(0)
        d1 = CLng(matches.SubMatches(0))
        m1 = MonthIndex(matches.SubMatches(1))
        d2 = CLng(matches.SubMatches(2))
        m2 = MonthIndex(matches.SubMatches(3))
        y1 = ToYYYY(matches.SubMatches(4))
        y2 = y1
        
        TransformTermSingle = WeekCodesFromRangeInclusive(d1, m1, y1, d2, m2, y2)
        Exit Function
    End If
    
    ' Pattern 3: d1-Mon1-yy1-d2-Mon2-yy2
    re.pattern = "^(\d{1,2})-([A-Za-z]{3,9})-(\d{2})-(\d{1,2})-([A-Za-z]{3,9})-(\d{2})$"
    If re.Test(leg) Then
        Set matches = re.Execute(leg)(0)
        d1 = CLng(matches.SubMatches(0))
        m1 = MonthIndex(matches.SubMatches(1))
        y1 = ToYYYY(matches.SubMatches(2))
        d2 = CLng(matches.SubMatches(3))
        m2 = MonthIndex(matches.SubMatches(4))
        y2 = ToYYYY(matches.SubMatches(5))
        
        TransformTermSingle = WeekCodesFromRangeInclusive(d1, m1, y1, d2, m2, y2)
        Exit Function
    End If
    
    ' Non-week transforms (quarters, summer, winter, FY)
    lowercapLeg = LCase(Trim(leg))
    
    Select Case True
        Case lowercapLeg Like "apr-jun-##": TransformTermSingle = "Q1-" & Right(lowercapLeg, 2): Exit Function
        Case lowercapLeg Like "apr-jun": TransformTermSingle = "Q1": Exit Function
        Case lowercapLeg Like "jul-sep-##": TransformTermSingle = "Q2-" & Right(lowercapLeg, 2): Exit Function
        Case lowercapLeg Like "jul-sep": TransformTermSingle = "Q2": Exit Function
        Case lowercapLeg Like "oct-dec-##": TransformTermSingle = "Q3-" & Right(lowercapLeg, 2): Exit Function
        Case lowercapLeg Like "oct-dec": TransformTermSingle = "Q3": Exit Function
        Case lowercapLeg Like "jan-mar-##": TransformTermSingle = "Q4-" & Format(CLng(Right(lowercapLeg, 2)) - 1, "00"): Exit Function
        Case lowercapLeg Like "jan-mar": TransformTermSingle = "Q4": Exit Function
        Case lowercapLeg Like "apr-sep-##": TransformTermSingle = "Sum-" & Right(lowercapLeg, 2): Exit Function
        Case lowercapLeg Like "apr-sep": TransformTermSingle = "Sum": Exit Function
        Case lowercapLeg Like "oct-mar-##": TransformTermSingle = "Win-" & Format(CLng(Right(lowercapLeg, 2)) - 1, "00"): Exit Function
        Case lowercapLeg Like "oct-mar": TransformTermSingle = "Win": Exit Function
        Case lowercapLeg Like "apr-mar-##": TransformTermSingle = "FY-" & Format(CLng(Right(lowercapLeg, 2)) - 1, "00"): Exit Function
        Case lowercapLeg Like "apr-mar": TransformTermSingle = "FY": Exit Function
        Case Else: TransformTermSingle = leg
    End Select
End Function

'=============================
' Convert date range to ISO weeks (same as JS weekCodesFromRangeInclusive)
'=============================
Function WeekCodesFromRangeInclusive(d1 As Long, m1 As Long, y1 As Long, _
                                     d2 As Long, m2 As Long, y2 As Long) As String
    Dim dtStart As Date, dtEnd As Date, dt As Date
    Dim dict As Object
    Dim wk As Long, yr As Long
    Dim arr As Variant
    Dim k As Variant
    Dim result As String
    Dim code As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    dtStart = DateSerial(y1, m1 + 1, d1)
    dtEnd = DateSerial(y2, m2 + 1, d2)
    
    For dt = dtStart To dtEnd
        wk = Application.WorksheetFunction.IsoWeekNum(dt)
        yr = Year(dt)
        Dim key As String
        key = wk & "-" & yr
        
        If Not dict.exists(key) Then
            dict.Add key, Array(False, 0, 0)
        End If
        
        arr = dict(key)
        ' Check if weekend
        If Weekday(dt, vbMonday) >= 6 Then
            arr(0) = True
            If arr(1) = 0 Or day(dt) < arr(1) Then arr(1) = day(dt)
            If arr(2) = 0 Or day(dt) > arr(2) Then arr(2) = day(dt)
        End If
        dict(key) = arr
    Next dt
    
    ' Build string
    result = ""
    For Each k In dict.Keys
        arr = dict(k)
        wk = Split(k, "-")(0)
        yr = Split(k, "-")(1)
        If arr(0) Then
            code = "WE " & Format(wk, "00") & " (D" & arr(1) & "-D" & arr(2) & ")"
        Else
            code = "Wk " & Format(wk, "00") & "-" & Right(yr, 2)
        End If
        If result <> "" Then result = result & ","
        result = result & code
    Next k
    
    WeekCodesFromRangeInclusive = result
End Function

'=============================
' Helpers
'=============================
Function MonthIndex(mon As String) As Long
    Select Case LCase(Left(mon, 3))
        Case "jan": MonthIndex = 0
        Case "feb": MonthIndex = 1
        Case "mar": MonthIndex = 2
        Case "apr": MonthIndex = 3
        Case "may": MonthIndex = 4
        Case "jun": MonthIndex = 5
        Case "jul": MonthIndex = 6
        Case "aug": MonthIndex = 7
        Case "sep": MonthIndex = 8
        Case "oct": MonthIndex = 9
        Case "nov": MonthIndex = 10
        Case "dec": MonthIndex = 11
        Case Else: MonthIndex = -1
    End Select
End Function

Function ToYYYY(yy As String) As Long
    Dim n As Long
    n = CLng(yy)
    If n < 50 Then
        ToYYYY = 2000 + n
    Else
        ToYYYY = 1900 + n
    End If
End Function



