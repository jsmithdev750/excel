Attribute VB_Name = "Module1"
Option Explicit

' FAST MONTHLY BILLING CHECK (Excel for Windows)
' Sheet1!A2 = invoice/delivery date (any date within that month)
' Sheet1!A5 = invoice workbook path (blank = this workbook)
' Sheet1!A9 = Trade Record V4 workbook path
'
' Outputs:
'   Client_Summary - one row per client with monthly invoice total
'   Billing_List   - one row per nonblank invoice amount (including numeric zero)
'   BG_CFD_Trades  - main V4 fields (File Ref through Client), plus match metadata
'   Match_Diagnostics - client/trade-month/BG-CFD/delivery-period counts
'
' Match: invoice client + invoice-header trade month = V4 Client + Contract Date month.
' BG/CFD detection checks BOTH Product text AND C/P (e.g., C/P=BG).
' Client_Name_Map (in THIS macro workbook) maps V4 Client names to Forward Bill Client names.
' Delivery Period must contain Sheet1!A2 invoice month. Brokerage amounts are not reconciled.
' FY26 = April 2026 - March 2027; Winter2026 = Oct 2026 - Mar 2027.
'
Private Const TRADE_READ_BATCH As Long = 4000
Private Const TRADE_WRITE_BATCH As Long = 250

Public Sub RunBillingFromSheet1()
    Dim cfg As Worksheet, invoiceBook As Workbook, tradeBook As Workbook
    Dim billOut As Worksheet, summaryOut As Worksheet, tradeOut As Worksheet, diagOut As Worksheet
    Dim pInvoice As String, pTrade As String, requestedMonth As String
    Dim openedInvoice As Boolean, openedTrade As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldCalc As XlCalculation, oldSecurity As Long, oldStatus As Variant
    Dim wanted As Object, wantedClients As Object, invoices As Object, totals As Object
    Dim allClientRows As Object, clientAliases As Object, periodCache As Object, periodInfo As Variant
    Dim v4Sh As Worksheet, templateSheet As Worksheet
    Dim hdr As Long, refCol As Long, dateCol As Long, prodCol As Long, clientCol As Long
    Dim maxTradeCols As Long, templateHdr As Long, templateRef As Long, nFound As Long
    Dim tradeBuf As Variant, buffered As Long, tradeNextRow As Long, candidateCount As Long
    Dim scanHeader As Long, scanRef As Long, scanDate As Long, scanProd As Long, scanClient As Long
    Dim lastTrade As Long, chunkStart As Long, chunkEnd As Long, n As Long, i As Long, c As Long
    Dim data As Variant, clientName As String, normClient As String, tradeMonth As String, key As String
    Dim ids As Collection, id As Variant, rec As Variant, tot As Variant
    Dim product As String, tag As String, ticket As String, refs As String
    Dim tradeCol As Long, cpCol As Long, periodCol As Long, invoiceCount As Long, message As String
    Dim periodText As String, periodStart As String, periodEnd As String
    Dim periodKnown As Boolean, periodFits As Boolean, qualifiedTag As Boolean
    Dim cpValue As String

    On Error GoTo Failure
    Set cfg = ThisWorkbook.Worksheets("Sheet1")
    requestedMonth = MonthKeyFast(cfg.Range("A2").Value2)
    pInvoice = NormalizePathFast(cfg.Range("A5").Value2)
    pTrade = NormalizePathFast(cfg.Range("A9").Value2)

    If Len(requestedMonth) = 0 Then
        MsgBox "Set Sheet1!A2 to a valid invoice month (for example 01-Oct-2026).", vbExclamation
        Exit Sub
    End If
    If Len(pTrade) = 0 Then
        MsgBox "Enter the full Trade Record V4 path in Sheet1!A9.", vbExclamation
        Exit Sub
    End If
    If Len(Dir$(pTrade, vbNormal)) = 0 Then
        MsgBox "Trade Record V4 was not found:" & vbCrLf & pTrade, vbExclamation
        Exit Sub
    End If
    If Len(pInvoice) > 0 Then
        If StrComp(pInvoice, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            If Len(Dir$(pInvoice, vbNormal)) = 0 Then
                MsgBox "Forward billing workbook was not found:" & vbCrLf & pInvoice, vbExclamation
                Exit Sub
            End If
        End If
    End If

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    oldSecurity = Application.AutomationSecurity
    oldStatus = Application.StatusBar
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3  'Disable macros in any workbook opened for reading.

    If Len(pInvoice) = 0 Or StrComp(pInvoice, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Set invoiceBook = ThisWorkbook
    Else
        Set invoiceBook = GetOpenWorkbookFast(pInvoice)
        If invoiceBook Is Nothing Then
            Application.StatusBar = "Opening invoice workbook..."
            Set invoiceBook = Workbooks.Open(Filename:=pInvoice, UpdateLinks:=0, ReadOnly:=True, _
                AddToMru:=False, IgnoreReadOnlyRecommended:=True)
            openedInvoice = True
        End If
    End If

    Set wanted = CreateObject("Scripting.Dictionary")
    wanted.CompareMode = vbTextCompare
    Set wantedClients = CreateObject("Scripting.Dictionary")
    wantedClients.CompareMode = vbTextCompare
    Set invoices = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare
    Set allClientRows = CreateObject("Scripting.Dictionary")
    allClientRows.CompareMode = vbTextCompare

    'Load aliases once, not once per trade row (important for large V4 workbooks).
    'Set clientAliases = LoadClientNameMapFast()
    Set periodCache = CreateObject("Scripting.Dictionary")
    periodCache.CompareMode = vbTextCompare

    Application.StatusBar = "Checking invoice cells for " & MonthLabelFast(requestedMonth) & "..."
    CollectInvoiceRowsFast invoiceBook, requestedMonth, wanted, wantedClients, invoices, totals
    invoiceCount = invoices.Count
    Set clientAliases = LoadClientNameMapFast(invoices)

    Set billOut = GetOutputSheetFast("Billing_List")
    Set summaryOut = GetOutputSheetFast("Client_Summary")
    Set tradeOut = GetOutputSheetFast("BG_CFD_Trades")
    Set diagOut = GetOutputSheetFast("Match_Diagnostics")
    If billOut.AutoFilterMode Then billOut.AutoFilterMode = False
    If summaryOut.AutoFilterMode Then summaryOut.AutoFilterMode = False
    If tradeOut.AutoFilterMode Then tradeOut.AutoFilterMode = False
    If diagOut.AutoFilterMode Then diagOut.AutoFilterMode = False
    billOut.UsedRange.Clear
    summaryOut.UsedRange.Clear
    tradeOut.UsedRange.Clear
    diagOut.UsedRange.Clear

    If invoiceCount > 0 Then
        Set tradeBook = GetOpenWorkbookFast(pTrade)
        If tradeBook Is Nothing Then
            Application.StatusBar = "Opening Trade Record V4..."
            Set tradeBook = Workbooks.Open(Filename:=pTrade, UpdateLinks:=0, ReadOnly:=True, _
                AddToMru:=False, IgnoreReadOnlyRecommended:=True)
            openedTrade = True
        End If

        'Use the first MAIN trade block (File Ref through Client), not later BUYER/SELLER blocks.
        For Each v4Sh In tradeBook.Worksheets
            If FindTradeHeadersFast(v4Sh, hdr, refCol, dateCol, prodCol, clientCol, tradeCol) Then
                nFound = nFound + 1
                If clientCol - refCol + 1 > maxTradeCols Then
                    maxTradeCols = clientCol - refCol + 1
                    templateHdr = hdr
                    templateRef = refCol
                    Set templateSheet = v4Sh
                End If
            End If
        Next v4Sh
        If nFound = 0 Then Err.Raise vbObjectError + 701, , _
            "No sheet with File Ref, Contract Date, Product and Client headers exists in Trade Record V4."

        WriteTradeHeadersFast tradeOut, templateSheet, templateHdr, templateRef, maxTradeCols
       ' ReDim tradeBuf(1 To TRADE_WRITE_BATCH, 1 To maxTradeCols + 9)
       ReDim tradeBuf(1 To TRADE_WRITE_BATCH, 1 To 20)
        tradeNextRow = 2

        For Each v4Sh In tradeBook.Worksheets
            If FindTradeHeadersFast(v4Sh, scanHeader, scanRef, scanDate, scanProd, scanClient, tradeCol) Then
                cpCol = FindCPColumnFast(v4Sh, scanHeader, scanClient)
                periodCol = FindPeriodColumnBilling(v4Sh, scanHeader, scanRef, scanClient)
                If periodCol = 0 Then Err.Raise vbObjectError + 732, , _
                    "Period header missing from MAIN trade block on V4 sheet: " & v4Sh.name
                lastTrade = Application.Max( _
                    v4Sh.Cells(v4Sh.Rows.Count, scanRef).End(xlUp).Row, _
                    v4Sh.Cells(v4Sh.Rows.Count, scanClient).End(xlUp).Row)
                If lastTrade > scanHeader Then
                    Application.StatusBar = "Searching " & v4Sh.name & " (" & _
                        Format$(lastTrade - scanHeader, "#,##0") & " records)..."
                    For chunkStart = scanHeader + 1 To lastTrade Step TRADE_READ_BATCH
                        chunkEnd = Application.Min(chunkStart + TRADE_READ_BATCH - 1, lastTrade)
                        data = v4Sh.Range(v4Sh.Cells(chunkStart, 1), _
                            v4Sh.Cells(chunkEnd, scanClient)).Value2
                        For i = 1 To UBound(data, 1)
                            clientName = SafeTextFast(data(i, scanClient))
                            If Len(clientName) > 0 Then
                                normClient = NormalizeClientFast(clientName)
                                'Translate the V4 name to its corresponding Forward Bill name.
                                'Names without an alias continue to match directly as before.
                                If clientAliases.Exists(normClient) Then
                                    normClient = CStr(clientAliases(normClient))
                                End If
                                If wantedClients.Exists(normClient) Then
                                    If allClientRows.Exists(normClient) Then
                                        allClientRows(normClient) = CLng(allClientRows(normClient)) + 1
                                    Else
                                        allClientRows.Add normClient, 1&
                                    End If
                                    tradeMonth = MonthKeyFast(data(i, scanDate))
                                    key = normClient & "|" & tradeMonth
                                    If wanted.Exists(key) Then
                                        Set ids = wanted(key)
                                        product = SafeTextFast(data(i, scanProd))
                                        cpValue = ""
                                        If cpCol > 0 Then cpValue = SafeTextFast(data(i, cpCol))
                                        tag = ProductTagFast(product, cpValue)
                                        qualifiedTag = (tag <> "Other")
                                        periodText = SafeTextFast(data(i, periodCol))
                                        periodKnown = False
                                        periodFits = False
                                        periodStart = ""
                                        periodEnd = ""
                                        If qualifiedTag Then
                                            If periodCache.Exists(periodText) Then
                                                periodInfo = periodCache(periodText)
                                                periodKnown = CBool(periodInfo(0))
                                                periodStart = CStr(periodInfo(1))
                                                periodEnd = CStr(periodInfo(2))
                                            Else
                                                periodKnown = TryPeriodBoundsBilling(periodText, periodStart, periodEnd)
                                                periodCache.Add periodText, Array(periodKnown, periodStart, periodEnd)
                                            End If
                                            If periodKnown Then
                                                periodFits = (CLng(requestedMonth) >= CLng(periodStart) And _
                                                    CLng(requestedMonth) <= CLng(periodEnd))
                                            End If
                                        End If
                                        ticket = SafeTextFast(data(i, scanRef))
                                        refs = ""
                                        For Each id In ids
                                            rec = invoices(CLng(id))
                                            rec(7) = CLng(rec(7)) + 1 'client + historical trade month
                                            If qualifiedTag Then
                                                rec(11) = CLng(rec(11)) + 1 'BG/CFD before delivery filter
                                                If Not periodKnown Then rec(12) = CLng(rec(12)) + 1
                                                If periodFits Then
                                                    rec(8) = CLng(rec(8)) + 1 'FINAL delivery-qualified count
                                                    rec(9) = AppendTicketFast(CStr(rec(9)), ticket)
                                                    If Len(refs) < 1200 Then
                                                        If Len(refs) > 0 Then refs = refs & ", "
                                                        refs = refs & CStr(rec(2)) & "!" & CStr(rec(3))
                                                    End If
                                                End If
                                            End If
                                            invoices(CLng(id)) = rec
                                        Next id

                                        'Only produce original V4 rows covering selected invoice month.
                                            If qualifiedTag And periodFits Then
                                            
                                                'Restore the client summary trade count.
                                                tot = totals(normClient)
                                                tot(4) = CLng(tot(4)) + 1
                                                totals(normClient) = tot
                                            
                                                buffered = buffered + 1
                                            
                                            
                                            'Copy exactly 20 fields from File Ref to Client.
                                            For c = 1 To 20
                                                tradeBuf(buffered, c) = data(i, scanRef + c - 1)
                                            Next c
                                            
                                            'Preserve ticket reference as text.
                                            tradeBuf(buffered, 1) = SafeTextFast(data(i, scanRef))
                                            
                                            candidateCount = candidateCount + 1
                                            
                                            If buffered = TRADE_WRITE_BATCH Then
                                                FlushTradesFast tradeOut, tradeBuf, buffered, _
                                                    tradeNextRow, 20
                                            End If

                                        End If
                                    End If
                                End If
                            End If
                        Next i
                        Erase data
                    Next chunkStart
                End If
            End If
        Next v4Sh
       If buffered > 0 Then
    FlushTradesFast tradeOut, tradeBuf, buffered, _
        tradeNextRow, 20
End If
        FormatTradeOutputFast tradeOut, maxTradeCols, templateSheet, templateHdr, templateRef, tradeNextRow - 1
    Else
        tradeOut.Range("A1").Value = "No populated invoice entries for " & MonthLabelFast(requestedMonth)
    End If

    WriteBillingResultsFast billOut, summaryOut, requestedMonth, invoices, totals
    WriteMatchDiagnosticsFast diagOut, requestedMonth, wanted, invoices, allClientRows
    summaryOut.Activate
    message = totals.Count & " client(s) to review, " & invoiceCount & " invoice cell(s), " & _
        candidateCount & " BG/CFD trade row(s)." & vbCrLf & vbCrLf & _
        "Results: Client_Summary, Billing_List, BG_CFD_Trades, Match_Diagnostics." & vbCrLf & _
        "BG/CFD records are candidates, not verified ticket-to-invoice matches."
    MsgBox message, vbInformation, "Billing check complete"
    GoTo CleanUp

Failure:
    message = Err.Description
    MsgBox "Billing check could not finish: " & message, vbExclamation
CleanUp:
    On Error Resume Next
    If openedTrade Then tradeBook.Close SaveChanges:=False
    If openedInvoice Then invoiceBook.Close SaveChanges:=False
    Application.StatusBar = oldStatus
    Application.AutomationSecurity = oldSecurity
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

Private Sub CollectInvoiceRowsFast(ByVal sourceBook As Workbook, ByVal monthWanted As String, _
    ByVal wanted As Object, ByVal wantedClients As Object, ByVal invoices As Object, _
    ByVal totals As Object)
    Dim sh As Worksheet, colB As Variant, headers As Variant, vals As Variant
    Dim lastR As Long, lastC As Long, hdrR As Long, nCols As Long
    Dim r As Long, j As Long, clientName As String, normClient As String
    Dim tradeKeys() As String, key As String, kMonth As String, bText As String
    Dim invVal As Variant, ids As Collection, rec(1 To 12) As Variant
    Dim tot(1 To 4) As Variant, currentTotal As Variant, idx As Long

    For Each sh In sourceBook.Worksheets
        If Not ShouldSkipInvoiceSheetFast(sh, sourceBook) Then
            lastR = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
            If lastR >= 2 Then
                colB = sh.Range("B1:B" & lastR).Value2
                clientName = ""
                hdrR = 0
                nCols = 0
                For r = 1 To lastR
                    bText = LCase$(SafeTextFast(colB(r, 1)))
                    If bText = "client:" Or bText = "client :" Then
                        clientName = SafeTextFast(sh.Cells(r, 3).Value2)
                        normClient = NormalizeClientFast(clientName)
                        hdrR = r + 1
                        nCols = 0
                        If hdrR <= lastR Then
                            If LCase$(SafeTextFast(colB(hdrR, 1))) = "month of delivery" Then
                                lastC = sh.Cells(hdrR, sh.Columns.Count).End(xlToLeft).Column
                                If lastC >= 3 Then
                                    nCols = lastC - 2
                                    ReDim tradeKeys(1 To nCols)
                                    If nCols = 1 Then
                                        tradeKeys(1) = MonthKeyFast(sh.Cells(hdrR, 3).Value2)
                                    Else
                                        headers = sh.Range(sh.Cells(hdrR, 3), sh.Cells(hdrR, lastC)).Value2
                                        For j = 1 To nCols
                                            tradeKeys(j) = MonthKeyFast(headers(1, j))
                                        Next j
                                    End If
                                End If
                            End If
                        End If
                    ElseIf bText = "total revenue" Or bText = "total" Then
                        clientName = ""
                        nCols = 0
                    ElseIf nCols > 0 And r > hdrR And Len(clientName) > 0 Then
                        If Len(bText) > 0 Then
                            If MonthKeyFast(colB(r, 1)) = monthWanted Then
                                If nCols = 1 Then
                                    ReDim vals(1 To 1, 1 To 1)
                                    vals(1, 1) = sh.Cells(r, 3).Value2
                                Else
                                    vals = sh.Range(sh.Cells(r, 3), sh.Cells(r, nCols + 2)).Value2
                                End If
                                For j = 1 To nCols
                                    invVal = vals(1, j)
                                    kMonth = tradeKeys(j)
                                    If InvoiceHasValueFast(invVal) And Len(kMonth) > 0 Then
                                        idx = invoices.Count + 1
                                        rec(1) = clientName
                                        rec(2) = sh.name
                                        rec(3) = sh.Cells(r, j + 2).Address(False, False)
                                        rec(4) = MonthLabelFast(monthWanted)
                                        rec(5) = MonthLabelFast(kMonth)
                                        rec(6) = invVal
                                        rec(7) = 0
                                        rec(8) = 0
                                        rec(9) = ""
                                        rec(10) = ""
                                        rec(11) = 0 'BG/CFD before period filtering
                                        rec(12) = 0 'unrecognised BG/CFD periods
                                        invoices.Add idx, rec
                                        key = normClient & "|" & kMonth
                                        If Not wanted.Exists(key) Then
                                            Set ids = New Collection
                                            wanted.Add key, ids
                                        End If
                                        Set ids = wanted(key)
                                        ids.Add idx
                                        If Not wantedClients.Exists(normClient) Then wantedClients.Add normClient, True
                                        If Not totals.Exists(normClient) Then
                                            tot(1) = clientName
                                            tot(2) = 0
                                            tot(3) = 0#
                                            tot(4) = 0
                                            totals.Add normClient, tot
                                        End If
                                        currentTotal = totals(normClient)
                                        currentTotal(2) = CLng(currentTotal(2)) + 1
                                        If IsNumeric(invVal) Then
                                            currentTotal(3) = CDbl(currentTotal(3)) + CDbl(invVal)
                                        End If
                                        totals(normClient) = currentTotal
                                    End If
                                Next j
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next sh
End Sub

Private Sub WriteBillingResultsFast(ByVal detail As Worksheet, ByVal clientsSheet As Worksheet, _
    ByVal requestedMonth As String, ByVal invoices As Object, ByVal totals As Object)
    Dim h As Variant, arr As Variant, sumArr As Variant
    Dim n As Long, i As Long, j As Long, rec As Variant, item As Variant, t As Variant
    h = Array("Client", "Invoice Sheet", "Invoice Cell", "Delivery Month", "Trade Month", _
        "Invoice Amount", "V4 Trade Rows", "BG/CFD Trade Rows", "BG/CFD Ticket Refs", "Status")
    detail.Range("A1:J1").Value = h
    n = invoices.Count
    If n > 0 Then
        ReDim arr(1 To n, 1 To 10)
        For i = 1 To n
            rec = invoices(i)
            If CLng(rec(7)) = 0 Then
                rec(10) = "NO CLIENT/TRADE-MONTH MATCH"
            ElseIf CLng(rec(11)) = 0 Then
                rec(10) = "NO BG/CFD PRODUCT"
            ElseIf CLng(rec(8)) = 0 Then
                If CLng(rec(12)) > 0 Then
                    rec(10) = "NO DELIVERY MATCH / CHECK PERIOD FORMAT"
                Else
                    rec(10) = "BG/CFD PERIOD DOES NOT COVER INVOICE MONTH"
                End If
            Else
                rec(10) = "BG/CFD DELIVERY-PERIOD CANDIDATES"
            End If
            For j = 1 To 10
                arr(i, j) = rec(j)
            Next j
        Next i
        detail.Range("A2").Resize(n, 10).Value2 = arr
    Else
        detail.Range("A2").Value2 = "No invoice entries for " & MonthLabelFast(requestedMonth)
    End If
    DetailFormatFast detail, n

    clientsSheet.Range("A1:F1").Value = Array("Client", "Invoice Month", "Invoice Cells", _
        "Monthly Invoice Total", "BG/CFD V4 Trades", "Status")
    n = totals.Count
    If n > 0 Then
        ReDim sumArr(1 To n, 1 To 6)
        i = 0
        For Each item In totals.Keys
            i = i + 1
            t = totals(item)
            sumArr(i, 1) = t(1)
            sumArr(i, 2) = MonthLabelFast(requestedMonth)
            sumArr(i, 3) = t(2)
            sumArr(i, 4) = t(3)
            sumArr(i, 5) = t(4)
            If CLng(t(4)) > 0 Then
                sumArr(i, 6) = "BG/CFD CANDIDATES FOUND"
            Else
                sumArr(i, 6) = "NO BG/CFD CANDIDATE"
            End If
        Next item
        clientsSheet.Range("A2").Resize(n, 6).Value2 = sumArr
    Else
        clientsSheet.Range("A2").Value2 = "No invoice entries for " & MonthLabelFast(requestedMonth)
    End If
    SummaryFormatFast clientsSheet, n
End Sub

Private Sub WriteTradeHeadersFast( _
    ByVal dest As Worksheet, _
    ByVal source As Worksheet, _
    ByVal sourceRow As Long, _
    ByVal sourceStartCol As Long, _
    ByVal colCount As Long)

    dest.Range("A1:T1").Value = Array( _
        "File Ref", "Contract Date", "Product", "Period", _
        "Client Role", "C/P", "Quantity", "MWh", _
        "Option_Style", "Option_Type", "Exercise_Type", _
        "Price/Premium", "JPY/kWh", "Strike", "JPY/kWh", _
        "Comm Rate", "JPY/kWh", "JPY", "Amount DUE", "Client")

    dest.Columns("A").NumberFormat = "@"

End Sub

Private Sub FlushTradesFast(ByVal dest As Worksheet, ByRef buffer As Variant, _
    ByRef n As Long, ByRef nextRow As Long, ByVal cols As Long)
    Dim temp As Variant, i As Long, j As Long
    If n < 1 Then Exit Sub
    If n = TRADE_WRITE_BATCH Then
        dest.Cells(nextRow, 1).Resize(n, cols).Value2 = buffer
    Else
        ReDim temp(1 To n, 1 To cols)
        For i = 1 To n
            For j = 1 To cols
                temp(i, j) = buffer(i, j)
            Next j
        Next i
        dest.Cells(nextRow, 1).Resize(n, cols).Value2 = temp
    End If
    nextRow = nextRow + n
    n = 0
End Sub

Private Function FindPeriodColumnBilling(ByVal sh As Worksheet, _
    ByVal headerRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long, h As String
    For c = firstCol To lastCol
        h = LCase$(SafeTextFast(sh.Cells(headerRow, c).Value2))
        Select Case h
            Case "period", "delivery period", "contract period"
                FindPeriodColumnBilling = c
                Exit Function
        End Select
    Next c
End Function

'Converts known V4 delivery products to inclusive YYYYMM bounds.
'Unknown periods return False (excluded from BG_CFD_Trades; counted in diagnostics).
Private Function TryPeriodBoundsBilling(ByVal rawPeriod As String, _
    ByRef startKey As String, ByRef endKey As String) As Boolean
    Dim clean As String, re As Object, matches As Object, m As Object
    Dim parts As String, aStart As String, aEnd As String
    Dim bStart As String, bEnd As String
    On Error GoTo NotRecognised
    startKey = "": endKey = ""
    clean = UCase$(Trim$(rawPeriod))
    clean = Replace(clean, ChrW$(&H2013), "-") 'en dash
    clean = Replace(clean, ChrW$(&H2014), "-") 'em dash
    clean = Replace(clean, ChrW$(&H2011), "-") 'nonbreaking hyphen
    clean = Replace(clean, ChrW$(&HFF0D), "-") 'fullwidth hyphen
    clean = Replace(clean, ChrW$(&HA0), " ")
    clean = Replace(clean, " ", "")
    clean = Replace(clean, vbTab, "")
    clean = Replace(clean, "-", "")
    clean = Replace(clean, "/", "")
    clean = Replace(clean, "_", "")
    If Len(clean) = 0 Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(FY[0-9]{2,4}|Q[1-4][0-9]{2,4}|WINTER[0-9]{2,4}|SUMMER[0-9]{2,4}|" & _
        "JANUARY[0-9]{2,4}|JAN[0-9]{2,4}|FEBRUARY[0-9]{2,4}|FEB[0-9]{2,4}|" & _
        "MARCH[0-9]{2,4}|MAR[0-9]{2,4}|APRIL[0-9]{2,4}|APR[0-9]{2,4}|" & _
        "MAY[0-9]{2,4}|JUNE[0-9]{2,4}|JUN[0-9]{2,4}|JULY[0-9]{2,4}|JUL[0-9]{2,4}|" & _
        "AUGUST[0-9]{2,4}|AUG[0-9]{2,4}|SEPTEMBER[0-9]{2,4}|SEPT[0-9]{2,4}|" & _
        "SEP[0-9]{2,4}|OCTOBER[0-9]{2,4}|OCT[0-9]{2,4}|" & _
        "NOVEMBER[0-9]{2,4}|NOV[0-9]{2,4}|DECEMBER[0-9]{2,4}|DEC[0-9]{2,4})"
    Set matches = re.Execute(clean)
    If matches.Count < 1 Or matches.Count > 2 Then Exit Function
    parts = ""
    For Each m In matches
        parts = parts & m.Value
    Next m
    'Reject partial/unknown parsing rather than returning wrong tickets.
    If StrComp(parts, clean, vbTextCompare) <> 0 Then Exit Function

    If Not TryPeriodTokenBilling(CStr(matches(0).Value), aStart, aEnd) Then Exit Function
    If matches.Count = 1 Then
        startKey = aStart
        endKey = aEnd
    Else
        If Not TryPeriodTokenBilling(CStr(matches(1).Value), bStart, bEnd) Then Exit Function
        startKey = aStart
        endKey = bEnd
    End If
    If CLng(endKey) < CLng(startKey) Then Exit Function
    TryPeriodBoundsBilling = True
    Exit Function
NotRecognised:
    startKey = "": endKey = ""
End Function

Private Function TryPeriodTokenBilling(ByVal token As String, _
    ByRef startKey As String, ByRef endKey As String) As Boolean
    Dim re As Object, matches As Object, yrText As String
    Dim prefix As String, yy As Long, mm As Long, q As Long
    On Error GoTo BadToken
    startKey = "": endKey = ""
    If Left$(UCase$(token), 1) = "Q" Then
        'Q426 means Q4 of 2026; do not accidentally parse year 426.
        prefix = UCase$(Left$(token, 2))
        yrText = Mid$(token, 3)
    Else
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "[0-9]{2,4}$"
        Set matches = re.Execute(token)
        If matches.Count <> 1 Then Exit Function
        yrText = CStr(matches(0).Value)
        prefix = UCase$(Left$(token, Len(token) - Len(yrText)))
    End If
    If Len(yrText) <> 2 And Len(yrText) <> 4 Then Exit Function
    yy = CLng(yrText)
    If Len(yrText) = 2 Then yy = 2000 + yy
    If yy < 2000 Or yy > 2099 Then Exit Function

    Select Case prefix
        Case "FY"
            startKey = Format$(DateSerial(yy, 4, 1), "yyyymm")
            endKey = Format$(DateSerial(yy + 1, 3, 1), "yyyymm")
        Case "WINTER"
            startKey = Format$(DateSerial(yy, 10, 1), "yyyymm")
            endKey = Format$(DateSerial(yy + 1, 3, 1), "yyyymm")
        Case "SUMMER"
            startKey = Format$(DateSerial(yy, 4, 1), "yyyymm")
            endKey = Format$(DateSerial(yy, 9, 1), "yyyymm")
        Case "Q1", "Q2", "Q3", "Q4"
            q = CLng(Right$(prefix, 1))
            startKey = Format$(DateSerial(yy, (q - 1) * 3 + 1, 1), "yyyymm")
            endKey = Format$(DateSerial(yy, q * 3, 1), "yyyymm")
        Case Else
            Select Case prefix
                Case "JAN", "JANUARY": mm = 1
                Case "FEB", "FEBRUARY": mm = 2
                Case "MAR", "MARCH": mm = 3
                Case "APR", "APRIL": mm = 4
                Case "MAY": mm = 5
                Case "JUN", "JUNE": mm = 6
                Case "JUL", "JULY": mm = 7
                Case "AUG", "AUGUST": mm = 8
                Case "SEP", "SEPT", "SEPTEMBER": mm = 9
                Case "OCT", "OCTOBER": mm = 10
                Case "NOV", "NOVEMBER": mm = 11
                Case "DEC", "DECEMBER": mm = 12
                Case Else: Exit Function
            End Select
            startKey = Format$(DateSerial(yy, mm, 1), "yyyymm")
            endKey = startKey
    End Select
    TryPeriodTokenBilling = True
    Exit Function
BadToken:
    startKey = "": endKey = ""
End Function

'Optional quick self-check. Run via Alt+F8: TestDeliveryPeriodsBilling
Public Sub TestDeliveryPeriodsBilling()
    Dim examples As Variant, item As Variant, sKey As String, eKey As String
    Dim passed As Long, Failed As Long, result As Boolean, covered As Boolean
    Dim expected As Boolean, details As String
    examples = Array( _
        Array("October-2026", True), Array("Q4-2026", True), _
        Array("Q1-2027", False), Array("FY26", True), _
        Array("FY-2026-FY-2027", True), Array("Winter2026", True), _
        Array("Summer-2026", False), Array("July-2026-March-2027", True), _
        Array("September-2026", False), Array("Q426", True), _
        Array("FY2027", False), Array("Jan-27", False), _
        Array("UNKNOWN PERIOD", False))
    For Each item In examples
        sKey = "": eKey = ""
        result = TryPeriodBoundsBilling(CStr(item(0)), sKey, eKey)
        covered = False
        If result Then covered = (CLng("202610") >= CLng(sKey) And CLng("202610") <= CLng(eKey))
        expected = CBool(item(1))
        If covered = expected Then
            passed = passed + 1
        Else
            Failed = Failed + 1
            details = details & CStr(item(0)) & " => " & sKey & " to " & eKey & vbCrLf
        End If
    Next item
    MsgBox passed & " passed; " & Failed & " failed (invoice month Oct-2026)." & _
        IIf(Failed > 0, vbCrLf & details, ""), vbInformation, "Period parser self-test"
End Sub

Private Function FindCPColumnFast(ByVal sh As Worksheet, ByVal headerRow As Long, _
    ByVal lastCol As Long) As Long
    Dim c As Long, h As String
    For c = 1 To lastCol
        h = LCase$(SafeTextFast(sh.Cells(headerRow, c).Value2))
        Select Case h
            Case "c/p", "c / p", "clearing platform"
                FindCPColumnFast = c
                Exit Function
        End Select
    Next c
End Function

Private Sub WriteMatchDiagnosticsFast(ByVal sh As Worksheet, ByVal invoiceMonth As String, _
    ByVal wanted As Object, ByVal invoices As Object, ByVal allClientRows As Object)
    Dim values As Variant, ids As Collection, k As Variant, keyText As String
    Dim p As Long, i As Long, rec As Variant, clientKey As String
    Dim totalRows As Long, matchedRows As Long, bgRows As Long
    Dim validRows As Long, unknownRows As Long, msg As String

    sh.Range("A1:I1").Value = Array("Invoice Client", "Invoice Month", "Trade Month", _
        "V4 Client Rows (Any Month)", "V4 Rows (Trade Month)", _
        "BG/CFD Rows (Trade Month)", "BG/CFD Rows (Delivery Month)", _
        "Unrecognised Period Rows", "Explanation")
    If wanted.Count = 0 Then
        sh.Range("A2").Value = "No populated invoice cells for this month."
        Exit Sub
    End If
    ReDim values(1 To wanted.Count, 1 To 9)
    i = 0
    For Each k In wanted.Keys
        i = i + 1
        keyText = CStr(k)
        p = InStrRev(keyText, "|")
        clientKey = Left$(keyText, p - 1)
        Set ids = wanted(keyText)
        rec = invoices(CLng(ids(1)))
        totalRows = 0
        If allClientRows.Exists(clientKey) Then totalRows = CLng(allClientRows(clientKey))
        matchedRows = CLng(rec(7))
        bgRows = CLng(rec(11))
        validRows = CLng(rec(8))
        unknownRows = CLng(rec(12))

        If totalRows = 0 Then
            msg = "Client not found in V4 (check Client_Name_Map)"
        ElseIf matchedRows = 0 Then
            msg = "No historical V4 contract-month match"
        ElseIf bgRows = 0 Then
            msg = "No BG/CFD product or C/P for the historical trade month"
        ElseIf validRows = 0 And unknownRows > 0 Then
            msg = "No delivery match. Some Period formats are UNRECOGNISED - review V4"
        ElseIf validRows = 0 Then
            msg = "BG/CFD found but delivery Period does NOT cover invoice month"
        ElseIf unknownRows > 0 Then
            msg = "Delivery candidates found; some other Period formats need review"
        Else
            msg = "Delivery period matches - inspect BG_CFD_Trades"
        End If

        values(i, 1) = rec(1)
        values(i, 2) = MonthLabelFast(invoiceMonth)
        values(i, 3) = rec(5)
        values(i, 4) = totalRows
        values(i, 5) = matchedRows
        values(i, 6) = bgRows
        values(i, 7) = validRows
        values(i, 8) = unknownRows
        values(i, 9) = msg
    Next k
    sh.Range("A2").Resize(wanted.Count, 9).Value2 = values
    HeaderFormatFast sh, 9, wanted.Count + 1
    sh.Columns("A").ColumnWidth = 38
    sh.Columns("B:C").ColumnWidth = 16
    sh.Columns("D:H").ColumnWidth = 25
    sh.Columns("I").ColumnWidth = 75
End Sub


Private Function FindTradeHeadersFast( _
    ByVal sh As Worksheet, _
    ByRef headerRow As Long, _
    ByRef fileCol As Long, _
    ByRef dateCol As Long, _
    ByRef productCol As Long, _
    ByRef clientCol As Long, _
    ByRef endCol As Long) As Boolean

    Dim r As Long
    Dim c As Long
    Dim txt As String
    Dim lastC As Long
    Dim sample As Variant

    For r = 1 To 20

        fileCol = 0
        dateCol = 0
        productCol = 0
        clientCol = 0

        lastC = Application.Min(100, _
            sh.Cells(r, sh.Columns.Count).End(xlToLeft).Column)

        If lastC >= 4 Then

            sample = sh.Range( _
                sh.Cells(r, 1), _
                sh.Cells(r, lastC)).Value2

            For c = 1 To lastC

                txt = LCase$(SafeTextFast(sample(1, c)))

                Select Case txt

                    Case "file ref", "file reference", "file_ref"
                        If fileCol = 0 Then fileCol = c

                    Case "contract date", "date of trade"
                        If dateCol = 0 Then dateCol = c

                    Case "product"
                        If productCol = 0 Then productCol = c

                    Case "client"
                        If clientCol = 0 Then clientCol = c

                End Select

            Next c

            If fileCol > 0 _
                And dateCol > fileCol _
                And productCol > dateCol _
                And clientCol > productCol Then

                headerRow = r
                endCol = lastC

                FindTradeHeadersFast = True
                Exit Function

            End If

        End If

    Next r

End Function



Private Sub FormatTradeOutputFast( _
    ByVal sh As Worksheet, _
    ByVal nSourceCols As Long, _
    ByVal template As Worksheet, _
    ByVal hdrRow As Long, _
    ByVal sourceStartCol As Long, _
    ByVal bottom As Long)

    HeaderFormatFast sh, 20, bottom

    sh.Range("A:T").ColumnWidth = 14

    sh.Columns("A").ColumnWidth = 18
    sh.Columns("A").NumberFormat = "@"

    sh.Columns("B").ColumnWidth = 17
    sh.Columns("B").NumberFormat = "dd-mmm-yyyy"

    sh.Columns("C").ColumnWidth = 48
    sh.Columns("D").ColumnWidth = 25
    sh.Columns("E").ColumnWidth = 15
    sh.Columns("G").ColumnWidth = 16
    sh.Columns("G").NumberFormat = "#,##0.00"

    sh.Columns("S").ColumnWidth = 19
    sh.Columns("S").NumberFormat = "#,##0.00"

    sh.Columns("T").ColumnWidth = 42

End Sub


Private Sub DetailFormatFast(ByVal sh As Worksheet, ByVal n As Long)
    HeaderFormatFast sh, 10, n + 1
    sh.Columns("A").ColumnWidth = 36
    sh.Columns("B").ColumnWidth = 23
    sh.Columns("C").ColumnWidth = 16
    sh.Columns("D:E").ColumnWidth = 16
    sh.Columns("F").ColumnWidth = 22
    sh.Columns("F").NumberFormat = "#,##0.00;[Red](#,##0.00);0.00"
    sh.Columns("G:H").ColumnWidth = 18
    sh.Columns("I").ColumnWidth = 44
    sh.Columns("I").NumberFormat = "@"
    sh.Columns("J").ColumnWidth = 37
End Sub

Private Sub SummaryFormatFast(ByVal sh As Worksheet, ByVal n As Long)
    HeaderFormatFast sh, 6, n + 1
    sh.Columns("A").ColumnWidth = 40
    sh.Columns("B").ColumnWidth = 17
    sh.Columns("C").ColumnWidth = 16
    sh.Columns("D").ColumnWidth = 24
    sh.Columns("D").NumberFormat = "#,##0.00;[Red](#,##0.00);0.00"
    sh.Columns("E").ColumnWidth = 21
    sh.Columns("F").ColumnWidth = 32
End Sub

Private Sub HeaderFormatFast(ByVal sh As Worksheet, ByVal cols As Long, ByVal bottom As Long)
    With sh.Range(sh.Cells(1, 1), sh.Cells(1, cols))
        .Interior.Color = RGB(24, 49, 83)
        .Font.Color = vbWhite
        .Font.Bold = True
        .RowHeight = 29
        .VerticalAlignment = xlCenter
    End With
    If bottom > 1 Then sh.Range(sh.Cells(1, 1), sh.Cells(bottom, cols)).AutoFilter
End Sub

Private Function GetOutputSheetFast(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOutputSheetFast = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOutputSheetFast Is Nothing Then
        Set GetOutputSheetFast = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOutputSheetFast.name = name
    End If
End Function


Private Function ShouldSkipInvoiceSheetFast( _
    ByVal sh As Worksheet, _
    ByVal wb As Workbook) As Boolean

    Select Case LCase$(sh.name)

        Case "billing_list", "client_summary", _
             "bg_cfd_trades", "match_diagnostics", _
             "client_name_map", "setting", _
             "bg_cfd_possible_missing"

            ShouldSkipInvoiceSheetFast = True

        Case "sheet1"

            If wb Is ThisWorkbook Then
                ShouldSkipInvoiceSheetFast = True
            End If

    End Select

End Function


' ===============================================================================
' EDITABLE CLIENT ALIASES
' A = client name EXACTLY as used in Forward Bill
' B = client name EXACTLY as used in Trade Record V4
' One row per alternative V4 name; duplicate Forward Bill names are allowed.
' The map is loaded into a Dictionary once per run for quick lookup.
' ===============================================================================


Public Sub SetupClientNameMap()

    Dim cfg As Worksheet
    Dim src As Workbook
    Dim sh As Worksheet
    Dim oldMap As Worksheet
    Dim filePath As String
    Dim saved As Object, defaults As Object
    Dim lastR As Long, mapLast As Long
    Dim r As Long, n As Long
    Dim oldLayout As Boolean
    Dim sheetName As String, v4Name As String
    Dim fbName As String, key As String
    Dim result() As Variant
    Dim openedHere As Boolean
    Dim hit As Range

    On Error GoTo Failed

    'Write the client mapping into Setting W:X
    Set cfg = ThisWorkbook.Worksheets("Setting")
    
    'Read Forward Bill file path from Sheet1
    filePath = NormalizePathFast( _
        ThisWorkbook.Worksheets("Sheet1").Range("A5").Value2)

    Set saved = CreateObject("Scripting.Dictionary")
    Set defaults = CreateObject("Scripting.Dictionary")
    saved.CompareMode = vbTextCompare
    defaults.CompareMode = vbTextCompare

    'Migrate old W:X:Y mapping into W:X.
    oldLayout = InStr(1, _
        SafeTextFast(cfg.Range("Y1").Value2), _
        "Trade Record V4 Client", vbTextCompare) > 0

    lastR = Application.Max( _
        cfg.Cells(cfg.Rows.Count, "W").End(xlUp).Row, _
        cfg.Cells(cfg.Rows.Count, "X").End(xlUp).Row, _
        cfg.Cells(cfg.Rows.Count, "Y").End(xlUp).Row)

    For r = 2 To lastR
        sheetName = SafeTextFast(cfg.Cells(r, "W").Value2)

        If Len(sheetName) > 0 Then
            If oldLayout Then
                v4Name = SafeTextFast(cfg.Cells(r, "Y").Value2)
            Else
                v4Name = SafeTextFast(cfg.Cells(r, "X").Value2)
            End If

            saved(NormalizeClientFast(sheetName)) = v4Name
        End If
    Next r

    'Use old mappings to prefill known client names.
    On Error Resume Next
    Set oldMap = ThisWorkbook.Worksheets("Client_Name_Map")
    On Error GoTo Failed

    If Not oldMap Is Nothing Then
        mapLast = oldMap.Cells(oldMap.Rows.Count, "A").End(xlUp).Row

        For r = 2 To mapLast
            fbName = NormalizeClientFast( _
                SafeTextFast(oldMap.Cells(r, "A").Value2))
            v4Name = SafeTextFast(oldMap.Cells(r, "B").Value2)

            If Len(fbName) > 0 And Len(v4Name) > 0 Then
                If Not defaults.Exists(fbName) Then
                    defaults.Add fbName, v4Name
                End If
            End If
        Next r
    End If

    'Open Forward Bill workbook.
    If Len(filePath) = 0 Or _
       StrComp(filePath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Set src = ThisWorkbook
    Else
        Set src = GetOpenWorkbookFast(filePath)

        If src Is Nothing Then
            Set src = Workbooks.Open( _
                Filename:=filePath, _
                UpdateLinks:=0, _
                ReadOnly:=True)
            openedHere = True
        End If
    End If

    ReDim result(1 To src.Worksheets.Count, 1 To 2)

    For Each sh In src.Worksheets

        If Not ShouldSkipInvoiceSheetFast(sh, src) Then

            n = n + 1
            result(n, 1) = sh.name
            key = NormalizeClientFast(sh.name)

            v4Name = ""

            If saved.Exists(key) Then
                v4Name = saved(key)
            Else
                Set hit = sh.Columns("B").Find( _
                    What:="Client:", _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    MatchCase:=False)

                If Not hit Is Nothing Then
                    fbName = NormalizeClientFast( _
                        SafeTextFast(sh.Cells(hit.Row, "C").Value2))

                    If defaults.Exists(fbName) Then
                        v4Name = defaults(fbName)
                    End If
                End If
            End If

            result(n, 2) = v4Name

        End If
    Next sh

    'Replace old mapping area.
    cfg.Range("W1:AA" & Application.Max(lastR, n + 1, 2)).ClearContents

    cfg.Range("W1:X1").Value = Array( _
        "Forward Bill Sheet", _
        "Trade Record V4 Client")

    For r = 1 To n
        cfg.Cells(r + 1, "W").Value2 = result(r, 1)
        cfg.Cells(r + 1, "X").Value2 = result(r, 2)
    Next r

    With cfg.Range("W1:X1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(24, 49, 83)
    End With

    cfg.Columns("W").ColumnWidth = 40
    cfg.Columns("X").ColumnWidth = 52

    MsgBox n & " Forward Bill sheets loaded.", vbInformation

CleanUp:
    On Error Resume Next
    If openedHere Then src.Close SaveChanges:=False
    Exit Sub

Failed:
    MsgBox Err.Description, vbExclamation
    Resume CleanUp

End Sub

Private Function LoadClientNameMapFast( _
    ByVal invoices As Object) As Object

    Dim aliases As Object
    Dim sheetClients As Object
    Dim cfg As Worksheet
    Dim oldMap As Worksheet

    Dim rec As Variant
    Dim lastR As Long
    Dim r As Long, i As Long
    Dim sheetKey As String
    Dim fbName As String, v4Name As String

    Set aliases = CreateObject("Scripting.Dictionary")
    Set sheetClients = CreateObject("Scripting.Dictionary")

    aliases.CompareMode = vbTextCompare
    sheetClients.CompareMode = vbTextCompare

    'Keep older client aliases as fallback.
    On Error Resume Next
    Set oldMap = ThisWorkbook.Worksheets("Client_Name_Map")
    On Error GoTo 0

    If Not oldMap Is Nothing Then
        lastR = oldMap.Cells(oldMap.Rows.Count, "A").End(xlUp).Row

        For r = 2 To lastR
            fbName = NormalizeClientFast( _
                SafeTextFast(oldMap.Cells(r, "A").Value2))

            v4Name = NormalizeClientFast( _
                SafeTextFast(oldMap.Cells(r, "B").Value2))

            If Len(fbName) > 0 And Len(v4Name) > 0 Then
                aliases(v4Name) = fbName
            End If
        Next r
    End If

    'Resolve worksheet names to invoice client names.
    For i = 1 To invoices.Count

        rec = invoices(i)

        sheetKey = NormalizeClientFast(CStr(rec(2)))
        fbName = NormalizeClientFast(CStr(rec(1)))

        If Not sheetClients.Exists(sheetKey) Then
            sheetClients.Add sheetKey, fbName
        ElseIf sheetClients(sheetKey) <> fbName Then
            sheetClients(sheetKey) = "#MULTIPLE_CLIENTS#"
        End If

    Next i

    Set cfg = ThisWorkbook.Worksheets("Setting")
    lastR = cfg.Cells(cfg.Rows.Count, "W").End(xlUp).Row

    For r = 2 To lastR

        sheetKey = NormalizeClientFast( _
            SafeTextFast(cfg.Cells(r, "W").Value2))

        v4Name = NormalizeClientFast( _
            SafeTextFast(cfg.Cells(r, "X").Value2))

        If Len(sheetKey) > 0 And Len(v4Name) > 0 Then

            If sheetClients.Exists(sheetKey) Then

                fbName = sheetClients(sheetKey)

                If fbName = "#MULTIPLE_CLIENTS#" Then
                    Err.Raise vbObjectError + 902, , _
                        "Multiple clients on Forward Bill sheet: " & _
                        cfg.Cells(r, "W").Value2
                End If

                'Settings mapping takes priority.
                aliases(v4Name) = fbName

            End If

        End If

    Next r

    Set LoadClientNameMapFast = aliases

End Function



Private Function GetOpenWorkbookFast(ByVal path As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            Set GetOpenWorkbookFast = wb
            Exit Function
        End If
    Next wb
End Function

Private Function InvoiceHasValueFast(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    InvoiceHasValueFast = Len(SafeTextFast(v)) > 0 And SafeTextFast(v) <> "-"
    'Numeric zero IS included.
End Function

Private Function ProductTagFast(ByVal productName As String, _
    Optional ByVal cpValue As String = "") As String
    Dim hasBg As Boolean, hasCfd As Boolean
    'Product can say "Balancing Group" without the literal "BG".
    'The trading platform/clearing column C/P can separately say "BG" or "CFD".
    hasBg = (StrComp(Trim$(cpValue), "BG", vbTextCompare) = 0) Or _
        (InStr(1, productName, "Balancing Group", vbTextCompare) > 0) Or _
        (InStr(1, productName, "BG", vbTextCompare) > 0)
    hasCfd = (StrComp(Trim$(cpValue), "CFD", vbTextCompare) = 0) Or _
        (InStr(1, productName, "CFD", vbTextCompare) > 0) Or _
        (InStr(1, productName, "Contract for Difference", vbTextCompare) > 0)
    If hasBg And hasCfd Then
        ProductTagFast = "BG + CFD"
    ElseIf hasBg Then
        ProductTagFast = "BG"
    ElseIf hasCfd Then
        ProductTagFast = "CFD"
    Else
        ProductTagFast = "Other"
    End If
End Function

Private Function AppendTicketFast(ByVal original As String, ByVal ticket As String) As String
    AppendTicketFast = original
    If Len(ticket) = 0 Or Len(original) > 3000 Then Exit Function
    If InStr(1, " | " & original & " | ", " | " & ticket & " | ", vbTextCompare) > 0 Then Exit Function
    If Len(original) = 0 Then
        AppendTicketFast = ticket
    Else
        AppendTicketFast = original & " | " & ticket
    End If
End Function

Private Function NormalizeClientFast(ByVal v As String) As String
    Dim s As String
    s = LCase$(Trim$(v))
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeClientFast = Trim$(s)
End Function

Private Function NormalizePathFast(ByVal v As Variant) As String
    Dim s As String
    s = SafeTextFast(v)
    If Len(s) >= 2 Then
        If Left$(s, 1) = """" And Right$(s, 1) = """" Then s = Mid$(s, 2, Len(s) - 2)
    End If
    NormalizePathFast = s
End Function

Private Function SafeTextFast(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeTextFast = Trim$(Replace(CStr(v), Chr$(160), " "))
End Function

Private Function MonthKeyFast(ByVal v As Variant) As String
    On Error GoTo BadDate
    Dim s As String, d As Date
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then
        If CDbl(v) < 1 Or CDbl(v) > 2958465 Then Exit Function
        d = DateAdd("d", CLng(Fix(CDbl(v))), DateSerial(1899, 12, 30))
        MonthKeyFast = Format$(d, "yyyymm")
        Exit Function
    End If
    s = SafeTextFast(v)
    If Len(s) = 0 Then Exit Function
    If IsDate("1-" & s) Then
        MonthKeyFast = Format$(CDate("1-" & s), "yyyymm")
    ElseIf IsDate(s) Then
        MonthKeyFast = Format$(CDate(s), "yyyymm")
    End If
BadDate:
End Function

Private Function MonthLabelFast(ByVal key As String) As String
    If Len(key) <> 6 Then Exit Function
    MonthLabelFast = Format$(DateSerial(CInt(Left$(key, 4)), CInt(Right$(key, 2)), 1), "mmm-yyyy")
End Function


Public Sub MoveBillingSettings()

    Dim oldSh As Worksheet
    Dim newSh As Worksheet
    Dim lastRow As Long

    Set oldSh = ThisWorkbook.Worksheets("Sheet1")
    Set newSh = ThisWorkbook.Worksheets("Setting")

    If Application.CountA(newSh.Range("A1:A10")) > 0 _
        Or Application.CountA(newSh.Range("W1:X50")) > 0 Then

        MsgBox "Setting already contains data. " & _
               "Review it before copying.", vbExclamation
        Exit Sub

    End If

    'Copy existing billing settings.
    oldSh.Range("A1:A10").Copy _
        Destination:=newSh.Range("A1")

    'Copy sheet-to-V4 client mappings.
    lastRow = Application.Max( _
        oldSh.Cells(oldSh.Rows.Count, "W").End(xlUp).Row, _
        oldSh.Cells(oldSh.Rows.Count, "X").End(xlUp).Row)

    oldSh.Range("W1:X" & lastRow).Copy _
        Destination:=newSh.Range("W1")

    newSh.Columns("W").ColumnWidth = 40
    newSh.Columns("X").ColumnWidth = 52

    newSh.Activate

    MsgBox "Settings moved successfully. " & _
           "Now update the VBA worksheet references.", vbInformation

End Sub


