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
    CategoryKey As String
    ContractKey As String
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
    Dim recIndex As Long
    
    '=============================
    ' Step 0: Build Region Order
    '=============================
    Dim orderDict As Object
    Set orderDict = BuildOrderDict(ws, "region order by")
    
    '=============================
    ' Step 0B: Build Contract Category Order
    '=============================
    Dim catOrderDict As Object
    Set catOrderDict = BuildOrderDict(ws, "contract category order by")
    
    '=============================
    ' Step 0C: Build Contract Order
    '=============================
    Dim contractOrderDict As Object
    Set contractOrderDict = BuildOrderDict(ws, "contract order by")
    
    '=============================
    ' Step 1: Transform + store
    '=============================
    ReDim recs(1 To lastRow - 1)
    recIndex = 1
    
    For i = 2 To lastRow
        text = ws.Cells(i, 1).Value
        If Trim(text) = "" Then GoTo NextRow
        
        parts = Split(text, " ")
        If UBound(parts) < 1 Then GoTo NextRow
        
        ' Transform contract
        contract = parts(1)
        legs = Split(contract, "/")
        ReDim outLegs(LBound(legs) To UBound(legs))
        
        For j = LBound(legs) To UBound(legs)
            outLegs(j) = TransformTerm(Trim(legs(j)))
        Next j
        
        transformedContract = Join(outLegs, "/")
        ws.Cells(i, 2).Value = transformedContract
        
        ' Rebuild full text
        newText = parts(0) & " " & transformedContract
        If UBound(parts) > 1 Then
            For j = 2 To UBound(parts)
                newText = newText & " " & parts(j)
            Next j
        End If
        
        ' Store
        recs(recIndex).FullText = newText
        recs(recIndex).RegionKey = GetPrimaryRegion(parts(0), orderDict)
        recs(recIndex).CategoryKey = GetContractCategory(transformedContract)
        recs(recIndex).ContractKey = GetContractOrderKey(transformedContract)
        recIndex = recIndex + 1
        
        
        'Refresh the wb link
        ThisWorkbook.RefreshAll
        
NextRow:
    Next i
    
    ' Resize array
    If recIndex > 1 Then
        ReDim Preserve recs(1 To recIndex - 1)
    Else
        MsgBox "No valid data found!", vbExclamation
        Exit Sub
    End If
    
    '=============================
    ' Step 2: Sort by Region -> Category -> Contract
    '=============================
    Call SortContractRecords(recs, orderDict, catOrderDict, contractOrderDict)
    
    '=============================
    ' Step 3: Output
    '=============================
    For i = 1 To UBound(recs)
        ws.Cells(i + 1, 3).Value = recs(i).FullText
    Next i
    
    MsgBox "Done! Sorted by Region, Category, and Contract Order.", vbInformation
End Sub

'=============================
' Sorting routine with 3 levels
'=============================
Sub SortContractRecords(ByRef recs() As ContractRecord, _
                        ByRef regionDict As Object, _
                        ByRef catDict As Object, _
                        ByRef contractDict As Object)
    Dim i As Long, j As Long
    Dim temp As ContractRecord
    Dim r1 As Long, r2 As Long
    Dim c1 As Long, c2 As Long
    Dim co1 As Long, co2 As Long
    
    For i = LBound(recs) To UBound(recs) - 1
        For j = i + 1 To UBound(recs)
            r1 = GetOrderIndex(recs(i).RegionKey, regionDict)
            r2 = GetOrderIndex(recs(j).RegionKey, regionDict)
            
            If r1 > r2 Then
                temp = recs(i): recs(i) = recs(j): recs(j) = temp
            ElseIf r1 = r2 Then
                c1 = GetOrderIndex(recs(i).CategoryKey, catDict)
                c2 = GetOrderIndex(recs(j).CategoryKey, catDict)
                
                If c1 > c2 Then
                    temp = recs(i): recs(i) = recs(j): recs(j) = temp
                ElseIf c1 = c2 Then
                    co1 = GetOrderIndex(LCase(recs(i).ContractKey), contractDict)
                    co2 = GetOrderIndex(LCase(recs(j).ContractKey), contractDict)
If co1 > co2 Then
    temp = recs(i): recs(i) = recs(j): recs(j) = temp

ElseIf co1 = co2 Then
    If CompareContractWithinType(recs(i).FullText, recs(j).FullText) > 0 Then
        temp = recs(i): recs(i) = recs(j): recs(j) = temp
    End If
End If

                End If
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
        m2 = monthIndex(matches.SubMatches(2))
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
        m1 = monthIndex(matches.SubMatches(1))
        d2 = CLng(matches.SubMatches(2))
        m2 = monthIndex(matches.SubMatches(3))
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
        m1 = monthIndex(matches.SubMatches(1))
        y1 = ToYYYY(matches.SubMatches(2))
        d2 = CLng(matches.SubMatches(3))
        m2 = monthIndex(matches.SubMatches(4))
        y2 = ToYYYY(matches.SubMatches(5))
        
        TransformTermSingle = WeekCodesFromRangeInclusive(d1, m1, y1, d2, m2, y2)
        Exit Function
    End If
    
    ' Pattern: single day like 20-Mar-26
    re.pattern = "^(\d{1,2})-([A-Za-z]{3,9})-(\d{2})$"
    If re.Test(leg) Then
        Set matches = re.Execute(leg)(0)
        d1 = CLng(matches.SubMatches(0))
        m1 = monthIndex(matches.SubMatches(1))
        y1 = ToYYYY(matches.SubMatches(2))
        
        ' single day ? use WeekCodesFromRangeInclusive with same start/end
        TransformTermSingle = WeekCodesFromRangeInclusive(d1, m1, y1, d1, m1, y1)
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
        Case lowercapLeg Like "apr-mar-##": TransformTermSingle = "FY" & Format(CLng(Right(lowercapLeg, 2)) - 1, "00"): Exit Function
        Case lowercapLeg Like "apr-mar": TransformTermSingle = "FY": Exit Function
        Case Else: TransformTermSingle = leg
    End Select
End Function

'=============================
' Convert date range to ISO weeks (same as JS weekCodesFromRangeInclusive)
'=============================
Function WeekCodesFromRangeInclusive(d1 As Long, m1 As Long, y1 As Long, _
                                     d2 As Long, m2 As Long, y2 As Long) As String
    Dim dtStart As Date, dtEnd As Date
    Dim dt As Date
    Dim result As String
    Dim wk As Long, yr As Long
    Dim weekDict As Object
    Dim totalDays As Long
    Dim startIsWeekend As Boolean, endIsWeekend As Boolean, isWE As Boolean
    Dim arrKeys() As Variant
    Dim firstYr As Long, lastYr As Long
    Dim i As Long
    
    dtStart = DateSerial(y1, m1 + 1, d1)
    dtEnd = DateSerial(y2, m2 + 1, d2)
    
    
    totalDays = dtEnd - dtStart + 1
    
    ' Determine weekend flags
    startIsWeekend = Weekday(dtStart, vbMonday) >= 6
    endIsWeekend = Weekday(dtEnd, vbMonday) >= 6
    
    ' Only mark as WE if BOTH start and end are weekends
    isWE = startIsWeekend And endIsWeekend
    
    If totalDays = 1 Then
        result = "D" & Format(d1, "00")
        WeekCodesFromRangeInclusive = result
        Exit Function
    End If
    
    If isWE Then
        ' Format WE week number as 2 digits, and day numbers as 2 digits
        result = "WE " & Format(Application.WorksheetFunction.IsoWeekNum(dtStart), "00") & _
                 "(D" & Format(d1, "00") & "-" & Format(d2, "00") & ")"
        WeekCodesFromRangeInclusive = result
        Exit Function
    End If
    
    ' If duration < 7 days, do NOT convert to WK - just return original range
    If totalDays < 7 Then
        result = d1 & "-" & d2 & "-" & Format(MonthName(m1 + 1, True), "mmm") & "-" & Right(CStr(y2), 2)
        WeekCodesFromRangeInclusive = result
        Exit Function
    End If
    
    ' Duration >= 7 days - split across ISO weeks
    Set weekDict = CreateObject("Scripting.Dictionary")
    
    For dt = dtStart To dtEnd
        ' Only include weekdays
        If Weekday(dt, vbMonday) < 6 Then
            wk = Application.WorksheetFunction.IsoWeekNum(dt)
            yr = Year(dt - Weekday(dt, vbMonday) + 4)
            Dim key As String
            key = wk & "-" & yr
            If Not weekDict.exists(key) Then
                weekDict.Add key, yr
            End If
        End If
    Next dt
    
    '===========================
    ' Build COMPRESSED result
    '===========================
    Dim weekList() As Long
    Dim yearList() As Long
    Dim k As Long
    Dim key2 As Variant
    
    ReDim weekList(0 To weekDict.Count - 1)
    ReDim yearList(0 To weekDict.Count - 1)
    
    ' Move dictionary to arrays
    k = 0
    For Each key2 In weekDict.Keys
        weekList(k) = CLng(Split(key2, "-")(0))
        yearList(k) = weekDict(key2)
        k = k + 1
    Next key2
    
    ' Sort weeks (important!)
Dim x As Long, y As Long
Dim tmpWk As Long, tmpYr As Long

For x = LBound(weekList) To UBound(weekList) - 1
    For y = x + 1 To UBound(weekList)
        
        If (yearList(x) > yearList(y)) Or _
           (yearList(x) = yearList(y) And weekList(x) > weekList(y)) Then
            
            ' Swap week
            tmpWk = weekList(x)
            weekList(x) = weekList(y)
            weekList(y) = tmpWk
            
            ' Swap year (IMPORTANT)
            tmpYr = yearList(x)
            yearList(x) = yearList(y)
            yearList(y) = tmpYr
            
        End If
        
    Next y
Next x
    
    ' Build compressed range
    Dim startWk As Long, endWk As Long
    startWk = weekList(LBound(weekList))
    endWk = weekList(UBound(weekList))
    
    Dim finalYear As Long
    finalYear = yearList(UBound(yearList))
    
    If startWk = endWk Then
        result = "Wk" & Format(startWk, "00") & "-" & Right(finalYear, 2)
    Else
        result = "Wk" & Format(startWk, "00") & "-Wk" & Format(endWk, "00") & "-" & Right(finalYear, 2)
    End If
    
    WeekCodesFromRangeInclusive = result
End Function

'=============================
' Helpers
'=============================
Function monthIndex(mon As String) As Long
    Select Case LCase(Left(mon, 3))
        Case "jan": monthIndex = 0
        Case "feb": monthIndex = 1
        Case "mar": monthIndex = 2
        Case "apr": monthIndex = 3
        Case "may": monthIndex = 4
        Case "jun": monthIndex = 5
        Case "jul": monthIndex = 6
        Case "aug": monthIndex = 7
        Case "sep": monthIndex = 8
        Case "oct": monthIndex = 9
        Case "nov": monthIndex = 10
        Case "dec": monthIndex = 11
        Case Else: monthIndex = -1
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
'=============================
' Updated GetPrimaryRegion
'=============================
Function GetPrimaryRegion(rawRegion As String, orderDict As Object) As String
    Dim r As String
    r = LCase(Trim(rawRegion))
    
    ' Full match first
    If orderDict.exists(r) Then
        GetPrimaryRegion = r
    Else
        ' fallback: first part only
        If InStr(r, "/") > 0 Then
            GetPrimaryRegion = Split(r, "/")(0)
        Else
            GetPrimaryRegion = r
        End If
    End If
End Function

'=============================
' Updated GetContractCategory
'=============================
Function GetContractCategory(contract As String) As String
    If InStr(contract, "/") > 0 Then
        GetContractCategory = "spread"
    ElseIf UBound(Split(contract, "-")) >= 2 Then
        GetContractCategory = "strips"
    Else
        GetContractCategory = "flat"
    End If
End Function
'=============================
' Build any order dictionary dynamically
'=============================
Function BuildOrderDict(ws As Worksheet, searchText As String) As Object
    Dim cell As Range, found As Boolean
    Dim orderDict As Object
    Dim idx As Long
    Set orderDict = CreateObject("Scripting.Dictionary")
    found = False
    
    ' Search for the header
    For Each cell In ws.UsedRange
        If LCase(Trim(cell.Value)) = LCase(searchText) Then
            found = True
            Exit For
        End If
    Next cell
    
    If Not found Then
        MsgBox "'" & searchText & "' not found!", vbCritical
        Exit Function
    End If
    
    ' Read values until blank
    idx = 1
    Dim r As Long
    r = cell.Row + 1
    Do While Trim(ws.Cells(r, cell.Column).Value) <> ""
        orderDict(LCase(Trim(ws.Cells(r, cell.Column).Value))) = idx
        idx = idx + 1
        r = r + 1
    Loop
    
    Set BuildOrderDict = orderDict
End Function
'=============================
' Get Contract Order key for sorting
'=============================
Function GetContractOrderKey(contract As String) As String
    Dim firstPart As String
    firstPart = contract
    
    ' DAY: D##
    If UCase(Left(firstPart, 1)) = "D" And IsNumeric(Mid(firstPart, 2, 2)) Then
        GetContractOrderKey = "day": Exit Function
    End If
    
    ' WE: starts with WE
    If UCase(Left(firstPart, 2)) = "WE" Then
        GetContractOrderKey = "we": Exit Function
    End If
    
    ' WK: starts with Wk
    If UCase(Left(firstPart, 2)) = "WK" Then
        GetContractOrderKey = "wk": Exit Function
    End If
    
    ' Month names (first 3 chars)
    Dim months As Variant
    months = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    Dim m As Variant
    For Each m In months
        If LCase(Left(firstPart, 3)) = m Then
            GetContractOrderKey = m: Exit Function
        End If
    Next m
    
    ' Q1-Q4
    If LCase(Left(firstPart, 2)) Like "q#" Then
        GetContractOrderKey = LCase(Left(firstPart, 2)): Exit Function
    End If
    
    ' Sum / Win (first 3 chars)
    If LCase(Left(firstPart, 3)) = "sum" Then
        GetContractOrderKey = "sum": Exit Function
    End If
    If LCase(Left(firstPart, 3)) = "win" Then
        GetContractOrderKey = "win": Exit Function
    End If
    
    ' FY (first 2 chars)
    If LCase(Left(firstPart, 2)) = "fy" Then
        GetContractOrderKey = "fy": Exit Function
    End If
    
    ' fallback
    GetContractOrderKey = firstPart
End Function


Function CompareContractWithinType(c1 As String, c2 As String) As Long
    
    Dim v1 As Double, v2 As Double
    
    v1 = GetIntraTypeValue(c1)
    v2 = GetIntraTypeValue(c2)
    
    If v1 > v2 Then
        CompareContractWithinType = 1
    ElseIf v1 < v2 Then
        CompareContractWithinType = -1
    Else
        CompareContractWithinType = 0
    End If

End Function
Function GetIntraTypeValue(contract As String) As Double
    
    Dim c As String
    c = LCase(contract)
    
    ' Use first leg only (important for spreads)
    If InStr(c, " ") > 0 Then c = Split(c, " ")(1)
    If InStr(c, "/") > 0 Then c = Split(c, "/")(0)
    
    '=====================
    ' DAY: D##
    '=====================
    If Left(c, 1) = "d" Then
        GetIntraTypeValue = val(Mid(c, 2))
        Exit Function
    End If
    
    '=====================
    ' WE: WE ##
    '=====================
If Left(c, 2) = "we" Then
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "we\s*(\d+)"
    re.IgnoreCase = True
    
    If re.Test(c) Then
        Set m = re.Execute(c)(0)
        GetIntraTypeValue = CLng(m.SubMatches(0))
    Else
        GetIntraTypeValue = 0
    End If
    Exit Function
End If
    
    '=====================
    ' WK: Wk##-YY
    '=====================
If Left(c, 2) = "wk" Then
    Dim wk As Long, yy As Long
    
    wk = val(Mid(c, 3, 2))
    yy = val(Right(c, 2))
    
    GetIntraTypeValue = yy * 100 + wk
    Exit Function
End If
    
    '=====================
    ' MONTH: Jan-YY
    '=====================
    Dim months As Variant
    months = Array("jan", "feb", "mar", "apr", "may", "jun", _
                   "jul", "aug", "sep", "oct", "nov", "dec")
    
    Dim i As Long
    For i = 0 To 11
        If Left(c, 3) = months(i) Then
            GetIntraTypeValue = val(Right(c, 2)) * 100 + i
            Exit Function
        End If
    Next i
    
    '=====================
    ' QUARTER: Q#-YY
    '=====================
    If Left(c, 1) = "q" Then
        Dim qNum As Long, qYear As Long
        
        qNum = val(Mid(c, 2, 1))
        qYear = val(Right(c, 2))
        
        ' Year priority, then quarter
        GetIntraTypeValue = qYear * 10 + qNum
        Exit Function
    End If
    
    '=====================
    ' SUM / WIN / FY
    '=====================
    If Left(c, 3) = "sum" Or Left(c, 3) = "win" Then
        GetIntraTypeValue = val(Right(c, 2))
        Exit Function
    End If
    
    If Left(c, 2) = "fy" Then
        GetIntraTypeValue = val(Right(c, 2))
        Exit Function
    End If
    
    ' fallback
    GetIntraTypeValue = 0

End Function



