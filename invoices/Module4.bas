Attribute VB_Name = "Module4"
Option Explicit

Public Sub GenerateNonGST()

    Dim wsSettings As Worksheet
    Dim gstWB As Workbook
    Dim invoiceDate As Date
    Dim dataMonth As String

    Dim ws As Worksheet
    Dim SelectedSheet As Worksheet
    Dim wb As Workbook
    Dim foundOpen As Boolean

    Dim gstSettingsWS As Worksheet
    Dim exchangeRate As Variant
    Dim gstInvoiceDate As Variant
    
    Dim targetDate As Date
    Dim monthLong As String
    Dim monthShort As String
    Dim yearLong As String
    Dim yearShort As String
    Dim wsName As String

    Set wsSettings = ThisWorkbook.Worksheets("Settings")

    '==========================
    'Read settings
    '==========================
    If Not IsDate(wsSettings.Range("B2").value) Then

        MsgBox "Invalid invoice date in Settings!B2.", vbCritical
        Exit Sub

    End If

    invoiceDate = wsSettings.Range("B2").value

    
    
    
    monthLong = Format(invoiceDate, "MMMM")
    monthShort = Format(invoiceDate, "MMM")
    yearLong = Format(invoiceDate, "yyyy")
    yearShort = Format(invoiceDate, "yy")

    '==========================
    'Check GST workbook is open
    '==========================
    foundOpen = False

    For Each wb In Application.Workbooks

        If InStr(1, LCase(wb.Name), "vgm invoice system", vbTextCompare) > 0 Then

            Set gstWB = wb
            foundOpen = True
            Exit For

        End If

    Next wb

    If Not foundOpen Then

        MsgBox _
            "Please open the GST workbook before continuing." & vbCrLf & _
            "(Workbook name should contain 'vgm invoice system').", _
            vbExclamation

        Exit Sub

    End If

    '==========================
    'Read GST settings sheet
    '==========================
    On Error Resume Next
    Set gstSettingsWS = gstWB.Worksheets(wsSettings.Range("B9").value)
    On Error GoTo 0

    If gstSettingsWS Is Nothing Then

        MsgBox _
            "Cannot find sheet '" & wsSettings.Range("B9").value & _
            "' in " & gstWB.Name, _
            vbCritical

        Exit Sub

    End If

    exchangeRate = gstSettingsWS.Range("B2").value
    gstInvoiceDate = gstSettingsWS.Range("A1").value
    
    '==========================
    'Sync invoice date
    '==========================
    If CDate(gstInvoiceDate) <> CDate(invoiceDate) Then
    
        gstSettingsWS.Range("F2").value = invoiceDate
    
        gstInvoiceDate = invoiceDate
    
    End If

    '==========================
    'User confirmation
    '==========================
    If MsgBox( _
        "Please verify:" & vbCrLf & vbCrLf & _
        "Exchange Rate : " & exchangeRate & vbCrLf & _
        "GST Workbook Invoice Date : " & _
            Format(gstInvoiceDate, "dd MMMM yyyy") & vbCrLf & _
        "Continue?", _
        vbYesNo + vbQuestion) = vbNo Then
    
        Exit Sub
    
    End If
    
    '==========================
    'Checking is folder existed
    '==========================
    Dim outputFolder As String

    outputFolder = GetInvoiceOutputFolder()
    
    If outputFolder = "" Then Exit Sub
    
    MsgBox "Saving location:" & vbCrLf & outputFolder

    '==========================
    'Find matching worksheets
    '==========================
    frmSelectSheet.cboSheets.Clear
    frmSelectSheet.SelectedSheet = ""
    
    Dim matchCount As Long
    
    matchCount = 0
    
    For Each ws In gstWB.Worksheets
    
        wsName = LCase(ws.Name)
    
        If _
            InStr(wsName, LCase(monthLong & " " & yearLong)) > 0 _
            Or InStr(wsName, LCase(monthShort & " " & yearLong)) > 0 _
            Or InStr(wsName, LCase(monthLong & " " & yearShort)) > 0 _
            Or InStr(wsName, LCase(monthShort & " " & yearShort)) > 0 Then
    
            matchCount = matchCount + 1
    
            Set SelectedSheet = ws
    
            frmSelectSheet.cboSheets.AddItem ws.Name
    
        End If
    
    Next ws
    
    If matchCount = 0 Then
    
    MsgBox _
        "No worksheet found for " & monthLong & " " & yearLong & ".", _
        vbCritical
    
        Exit Sub
    
    End If
    
    'Only one sheet found - use it automatically
    If matchCount = 1 Then
    
        MsgBox "Using worksheet:" & vbCrLf & SelectedSheet.Name, vbInformation
    
    Else
    
        'Multiple sheets found - let user choose
        frmSelectSheet.Show
    
        If frmSelectSheet.SelectedSheet = "" Then Exit Sub
    
        Set SelectedSheet = gstWB.Worksheets(frmSelectSheet.SelectedSheet)
    
        MsgBox "Using worksheet:" & vbCrLf & SelectedSheet.Name, vbInformation
    
    End If
    
    '===========================================================
    'NEXT STEP
    '
    'SelectedSheet is now the worksheet chosen by the user.
    '===========================================================
    Call ProcessNonGSTInvoices(gstWB, SelectedSheet)


    '==============================
    'Ensure calculations are complete
    '==============================
  '  gstWB.Calculate
    
   ' Do While Application.CalculationState <> xlDone
    '    DoEvents
    'Loop
   
    gstWB.Save
    MsgBox "GST generation completed."
    
    

End Sub
Private Sub ProcessNonGSTInvoices( _
    ByVal gstWB As Workbook, _
    ByVal wsData As Worksheet)

    Dim wsMap As Worksheet
    Dim wsClientSetting As Worksheet
    
    Dim lastRow As Long
    Dim r As Long

    Dim clientName As String
    Dim invoiceRef As String

    Dim startRow As Long
    Dim endRow As Long
    

    Set wsMap = ThisWorkbook.Worksheets("Sheet Name and C2")
    Set wsClientSetting = ThisWorkbook.Worksheets("Client Setting")

    lastRow = wsData.Cells(wsData.Rows.Count, "T").End(xlUp).Row

    r = 2

    Do While r <= lastRow

        If Trim(wsData.Cells(r, "AH").value) <> "" Then

            invoiceRef = wsData.Cells(r, "AH").value
            clientName = wsData.Cells(r, "T").value
            
            '==============================
            'Skip GST clients
            '==============================
            If IsGSTClient(wsClientSetting, clientName) Then
            
                Do While endRow < lastRow
            
                    If Trim(wsData.Cells(endRow + 1, "AH").value) <> "" Then Exit Do
            
                    If wsData.Cells(endRow + 1, "T").value <> clientName Then Exit Do
            
                    endRow = endRow + 1
            
                Loop
            
                r = endRow + 1
                GoTo ContinueLoop
            
            End If

            startRow = r

            endRow = r

            Do While endRow < lastRow

                If Trim(wsData.Cells(endRow + 1, "AH").value) <> "" Then Exit Do

                If wsData.Cells(endRow + 1, "T").value <> clientName Then Exit Do

                endRow = endRow + 1

            Loop

            '==============================
            'Generate Non template
            '==============================
            FillNonGSTTemplate _
                gstWB, _
                wsMap, _
                wsData, _
                clientName, _
                invoiceRef, _
                startRow, _
                wsClientSetting, _
                endRow
                
            '==============================
            'Export PDF
            '==============================
            ExportInvoicePDF _
                gstWB, _
                wsMap, _
                wsData, _
                clientName, _
                invoiceRef, _
                startRow

            '==============================
            'Generate Excel export if required
            '==============================
            GenerateExcelExport _
                wsClientSetting, _
                clientName, _
                invoiceRef, _
                wsData, _
                startRow, _
                endRow

            r = endRow + 1

        Else

            r = r + 1

        End If
ContinueLoop:
    Loop

End Sub
Private Sub FillNonGSTTemplate( _
    ByVal gstWB As Workbook, _
    ByVal wsMap As Worksheet, _
    ByVal wsData As Worksheet, _
    ByVal clientName As String, _
    ByVal invoiceRef As String, _
    ByVal startRow As Long, _
    ByVal clientSettingWS As Worksheet, _
    ByVal endRow As Long)

    Dim mapLastRow As Long
    Dim r As Long

    Dim templateSheetName As String
    Dim wsTemplate As Worksheet

    Dim totalCell As Range
    Dim totalRow As Long
    Dim tradeCount As Long
    
    Dim wsClientSetting As Worksheet
    Dim invoiceSplitType As String
    Dim minAmount As Double

    Dim foundClient As Boolean
    Dim lastClientRow As Long

    '==============================
    'Find template sheet
    '==============================
    Dim templateStartRow As Long
    
    foundClient = False
    templateSheetName = ""
    
    mapLastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    
    '----------------------------------------
    'Find where the Non-GST template section starts
    '----------------------------------------
    templateStartRow = 0
    
    For r = 1 To mapLastRow
    
        If InStr(1, _
                 wsMap.Cells(r, "A").value, _
                 "VGM Invoice System", _
                 vbTextCompare) > 0 Then
    
            templateStartRow = r + 1
            Exit For
    
        End If
    
    Next r
    
    If templateStartRow = 0 Then
    
        MsgBox "Cannot locate the Non-GST template section.", vbCritical
        Exit Sub

End If
    
    '----------------------------------------
    'Search ONLY the Non-GST template section
    '----------------------------------------
    For r = templateStartRow To mapLastRow
    
        If StrComp(Trim(wsMap.Cells(r, "B").value), _
                   Trim(clientName), _
                   vbTextCompare) = 0 Then
    
            foundClient = True
            templateSheetName = Trim(wsMap.Cells(r, "A").value)
            Exit For
    
        End If
    
    Next r
    
    '----------------------------------------
    'Use default template if no mapping exists
    '----------------------------------------
    If templateSheetName = "" Then
    
        templateSheetName = "Template - Energy"
    
    End If

    On Error Resume Next
    Set wsTemplate = gstWB.Worksheets(templateSheetName)
    On Error GoTo 0

    If wsTemplate Is Nothing Then Exit Sub

    '==============================
    'Invoice Number
    '==============================
    wsTemplate.Range("Q2").value = invoiceRef

    '==============================
    'Locate Total
    '==============================
    Set totalCell = wsTemplate.Cells.Find( _
        What:="Total", _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If totalCell Is Nothing Then Exit Sub

    totalRow = totalCell.Row

    '==============================
    'Delete existing trade rows
    '==============================
    If totalRow > 8 Then
        wsTemplate.Rows("8:" & totalRow - 1).Delete
    End If

    tradeCount = endRow - startRow + 1

    '==============================
    'Copy and INSERT copied cells
    '(Keeps formatting from source)
    '==============================
    wsData.Range("A" & startRow & ":S" & endRow).Copy

    wsTemplate.Range("A8").Insert Shift:=xlDown

    Application.CutCopyMode = False

    '==============================
    'Find Total again
    '==============================
    Set totalCell = wsTemplate.Cells.Find( _
        What:="Total", _
        LookIn:=xlValues, _
        LookAt:=xlWhole)
        
    
    Set wsClientSetting = ThisWorkbook.Worksheets("Client Setting")
    lastClientRow = clientSettingWS.Cells(wsClientSetting.Rows.Count, "A").End(xlUp).Row
    invoiceSplitType = GetSplitType(wsData.Cells(startRow, "C").value)
    
    For r = 2 To lastClientRow
    
        If StrComp(Trim(wsClientSetting.Cells(r, "A").value), _
                   Trim(clientName), vbTextCompare) = 0 Then
    
            If UCase(Trim(wsClientSetting.Cells(r, "E").value)) = invoiceSplitType Then
    
                If IsNumeric(wsClientSetting.Cells(r, "D").value) Then
                    minAmount = wsClientSetting.Cells(r, "D").value
                End If
    
                Exit For
    
            ElseIf Trim(wsClientSetting.Cells(r, "E").value) = "" Then
    
                If IsNumeric(wsClientSetting.Cells(r, "D").value) Then
                    minAmount = wsClientSetting.Cells(r, "D").value
                End If
    
                Exit For
    
            End If
    
        End If
    
    Next r
    
    ' TOTAL/SUM
    Dim totalAmount As Double
    
    totalAmount = Application.WorksheetFunction.Sum( _
        wsData.Range("S" & startRow & ":S" & endRow))
        
    If Not totalCell Is Nothing Then

        With totalCell.Offset(0, 4)
        
            If minAmount > 0 And totalAmount < minAmount Then
                .value = minAmount
            Else
                .Formula = "=SUM(S8:S" & tradeCount + 7 & ")"
            End If
        
        End With

    End If

End Sub
Private Function GetInvoiceOutputFolder() As String

    Dim wsSettings As Worksheet
    Dim basePath As String
    
    Dim invoiceDate As Date
    Dim yearFolder As String
    Dim monthFolder As String
    
    Dim yearPath As String
    Dim monthPath As String
    
    Dim fso As Object
    
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '==============================
    'Read settings
    '==============================
    basePath = "C:\Users\" & wsSettings.Range("A5").value & "\" & _
               Trim(wsSettings.Range("A16").value)
    
    If basePath = "" Then
        MsgBox "Output folder path missing in Settings!A16"
        Exit Function
    End If
    
    
    If Not IsDate(wsSettings.Range("B2").value) Then
        MsgBox "Invalid invoice date"
        Exit Function
    End If
    
    invoiceDate = wsSettings.Range("B2").value
    
    
    '==============================
    'Folder names
    '==============================
    yearFolder = Format(invoiceDate, "yyyy")
    
    monthFolder = Format(invoiceDate, "yyyymm") & " " & _
                  Format(invoiceDate, "mmm")
    
    
    '==============================
    'Year folder
    '==============================
    yearPath = fso.BuildPath(basePath, yearFolder)
    
    If Not fso.FolderExists(yearPath) Then
        
        fso.CreateFolder yearPath
        
    End If
    
    
    '==============================
    'Month folder
    '==============================
    monthPath = fso.BuildPath(yearPath, monthFolder)
    
    If Not fso.FolderExists(monthPath) Then
        
        fso.CreateFolder monthPath
        
    End If
    
    
    GetInvoiceOutputFolder = monthPath


End Function
Private Sub GenerateExcelExport( _
    ByVal clientSettingWS As Worksheet, _
    ByVal clientName As String, _
    ByVal invoiceRef As String, _
    ByVal wsData As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long)

    Dim r As Long
    Dim lastClientRow As Long

    Dim exportExcel As Boolean
    Dim shortName As String
    Dim splitType As String

    Dim settingsWS As Worksheet

    Dim templatePath As String
    Dim outputFolder As String
    Dim newFile As String

    Dim exportWB As Workbook
    Dim exportWS As Worksheet

    Dim sheetName As String
    Dim rowCount As Long
    

    Dim invoiceSplitType As String

    On Error GoTo ErrorHandler

    Set settingsWS = ThisWorkbook.Worksheets("Settings")

    '==============================
    'Find client setting
    '==============================
    lastClientRow = clientSettingWS.Cells(clientSettingWS.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastClientRow
    
        If StrComp(Trim(clientSettingWS.Cells(r, "A").value), _
                   Trim(clientName), _
                   vbTextCompare) = 0 Then
    
            'Split matches
            If UCase(Trim(clientSettingWS.Cells(r, "E").value)) = invoiceSplitType Then
    
                exportExcel = (clientSettingWS.Cells(r, "B").value = True)
                shortName = Trim(clientSettingWS.Cells(r, "C").value)
                Exit For
                
            ElseIf Trim(clientSettingWS.Cells(r, "E").value) = "" Then
                 'Normal client (no split configured)
                exportExcel = (clientSettingWS.Cells(r, "B").value = True)
                shortName = Trim(clientSettingWS.Cells(r, "C").value)
                Exit For

            End If
    
        End If
    
    Next r

    '==============================
    'No export required
    '==============================
    If exportExcel = False Then Exit Sub

    '==============================
    'Get save folder
    '==============================
    outputFolder = GetInvoiceOutputFolder()

    If outputFolder = "" Then Exit Sub

    '==============================
    'Template path
    '==============================
    templatePath = "C:\Users\" & _
                   Trim(settingsWS.Range("A5").value) & "\" & _
                   Trim(settingsWS.Range("A13").value)

    If Dir(templatePath) = "" Then
        Err.Raise vbObjectError + 1000, , _
            "Excel template not found:" & vbCrLf & templatePath
    End If

    '==============================
    'Build filename
    '==============================
    If shortName <> "" Then

        newFile = outputFolder & "\" & _
                  invoiceRef & "_" & shortName & ".xlsx"

    Else

        newFile = outputFolder & "\" & _
                  invoiceRef & ".xlsx"

    End If

    '==============================
    'Skip if already exists
    '==============================
    If Dir(newFile) <> "" Then Exit Sub

    '==============================
    'Copy template
    '==============================
    FileCopy templatePath, newFile

    '==============================
    'Open copied workbook
    '==============================
    Set exportWB = Workbooks.Open(newFile)

    Set exportWS = exportWB.Worksheets(1)

    '==============================
    'Rename sheet
    '==============================
    sheetName = Format(settingsWS.Range("B2").value, "mmmm yyyy")

    On Error Resume Next
    exportWS.Name = sheetName
    Err.Clear
    On Error GoTo ErrorHandler

    '==============================
    'Clear template
    '==============================
    exportWS.Rows("2:" & exportWS.Rows.Count).ClearContents

    '==============================
    'Paste values
    '==============================
    rowCount = endRow - startRow + 1

    exportWS.Range("A2:S" & rowCount + 1).value = _
        wsData.Range("A" & startRow & ":S" & endRow).value

    '==============================
    'Ensure calculations are complete
    '==============================
    'exportWB.Calculate
    
    'Do While Application.CalculationState <> xlDone
     '   DoEvents
    'Loop
    
    '==============================
    'Save
    '==============================
    exportWB.Save
    exportWB.Close SaveChanges:=False
    Set exportWB = Nothing

CleanExit:

    On Error Resume Next

    If Not exportWB Is Nothing Then
        exportWB.Close SaveChanges:=False
        Set exportWB = Nothing
    End If

    Exit Sub

ErrorHandler:

    On Error Resume Next

    If Not exportWB Is Nothing Then
        exportWB.Close SaveChanges:=False
        Set exportWB = Nothing
    End If

    MsgBox _
        "GenerateExcelExport failed." & vbCrLf & vbCrLf & _
        "Client : " & clientName & vbCrLf & _
        "Invoice : " & invoiceRef & vbCrLf & vbCrLf & _
        "Error " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume CleanExit

End Sub

Private Function IsGSTClient(ws As Worksheet, clientName As String) As Boolean

    Dim r As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow

        If StrComp(Trim(ws.Cells(r, "A").value), _
                   Trim(clientName), _
                   vbTextCompare) = 0 Then

            IsGSTClient = (ws.Cells(r, "D").value = True)
            Exit Function

        End If

    Next r

End Function
Private Sub ExportInvoicePDF( _
    ByVal gstWB As Workbook, _
    ByVal wsMap As Worksheet, _
    ByVal wsData As Worksheet, _
    ByVal clientName As String, _
    ByVal invoiceRef As String, _
    ByVal startRow As Long)

    Dim outputFolder As String
    Dim pdfFile As String

    Dim wsTemplate As Worksheet
    Dim templateSheetName As String

    Dim lastRow As Long
    Dim r As Long
    
    Dim wsClientSetting As Worksheet
    Dim shortName As String
    
    Dim invoiceSplitType As String

    outputFolder = GetInvoiceOutputFolder()

    If outputFolder = "" Then Exit Sub
    
    '==============================
    'Get client short name
    '==============================
    Set wsClientSetting = ThisWorkbook.Worksheets("Client Setting")
    invoiceSplitType = GetSplitType(wsData.Cells(startRow, "C").value)
    
    lastRow = wsClientSetting.Cells(wsClientSetting.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
    
        If StrComp(Trim(wsClientSetting.Cells(r, "A").value), _
                   Trim(clientName), _
                   vbTextCompare) = 0 Then
                   
            'Split matches
            If UCase(Trim(wsClientSetting.Cells(r, "E").value)) = invoiceSplitType Then
    
                shortName = Trim(wsClientSetting.Cells(r, "C").value)
                Exit For
                
            ElseIf Trim(wsClientSetting.Cells(r, "E").value) = "" Then
                 'Normal client (no split configured)
                shortName = Trim(wsClientSetting.Cells(r, "C").value)
                Exit For

            End If
    
            
            
            'Exit For
    
        End If
    
    Next r

    'Find template
    lastRow = wsMap.Cells(wsMap.Rows.Count, "B").End(xlUp).Row

    For r = 2 To lastRow

        If StrComp(Trim(wsMap.Cells(r, "B").value), _
                   Trim(clientName), _
                   vbTextCompare) = 0 Then

            templateSheetName = Trim(wsMap.Cells(r, "A").value)
            Exit For

        End If

    Next r

    If templateSheetName = "" Then
        templateSheetName = "Template - Energy"
    End If
    

    On Error Resume Next
    Set wsTemplate = gstWB.Worksheets(templateSheetName)
    On Error GoTo 0
    
    'Template doesn't exist in this workbook (likely a GST client)
    If wsTemplate Is Nothing Then Exit Sub

    'Only the default template needs the client name updated
    If templateSheetName = "Template - Energy" Then
        wsTemplate.Range("C2").value = clientName
    End If
    
    If shortName <> "" Then
    
        pdfFile = outputFolder & "\" & _
                  invoiceRef & "_" & shortName & ".pdf"
    
    Else
    
        pdfFile = outputFolder & "\" & _
                  invoiceRef & ".pdf"
    
    End If

    'Already exists
    If Dir(pdfFile) <> "" Then Exit Sub

    wsTemplate.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

End Sub

Private Function GetSplitType(ByVal product As String) As String

    product = UCase(product)

    If InStr(product, "BALANCING GROUP") > 0 _
        Or InStr(product, "(BG)") > 0 Then

        GetSplitType = "BG"

    Else
        GetSplitType = "EEX"

    End If

End Function

