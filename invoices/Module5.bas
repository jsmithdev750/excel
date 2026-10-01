Attribute VB_Name = "Module5"
Option Explicit

'====================================================================
' VGMJ INVOICE GENERATOR
'
' OUTPUT:
'
' 1. NORMAL VGMJ EXCEL
'    - ALWAYS created
'    - Template comes from Settings!A22
'    - Contains English + Japanese sheets
'
' 2. ADDITIONAL TRADE EXCEL
'    - ONLY created when Client Setting column B = TRUE
'    - Template comes from Settings!A13
'    - Filename ends with "_trade.xlsx"
'    - Trade sheet renamed to "MMM YYYY"
'
' EXAMPLE:
'
' Normal:
'   VGMJ202609001_Theme.xlsx
'
' Trade:
'   VGMJ202609001_Theme_trade.xlsx
'
' Trade sheet:
'   Sep 2026
'
'====================================================================


'====================================================================
' MAIN
'====================================================================

Public Sub GenerateVGMJ()

    Dim wsSettings As Worksheet
    Dim invoiceWB As Workbook
    Dim wb As Workbook
    Dim wsVGMJ As Worksheet

    Dim invoiceFolder As String
    Dim outputFolder As String

    Set wsSettings = ThisWorkbook.Worksheets("Settings")

    '====================================================
    ' CHECK INVOICE DATE
    '====================================================

    If Not IsDate(wsSettings.Range("A2").Value) Then

        MsgBox _
            "Invalid invoice date in Settings!A2.", _
            vbCritical

        Exit Sub

    End If

    '====================================================
    ' GET INVOICE DATA FOLDER FROM A16
    '====================================================

    invoiceFolder = _
        GetSettingsPath( _
            CStr(wsSettings.Range("A16").Value))

    If invoiceFolder = "" Then

        MsgBox _
            "Cannot determine the invoice data folder." & _
            vbCrLf & vbCrLf & _
            "Please check Settings!A16.", _
            vbCritical

        Exit Sub

    End If

    '====================================================
    ' FIND OPEN INVOICE DATA WORKBOOK
    '====================================================

    Set invoiceWB = Nothing

    For Each wb In Application.Workbooks

        If InStr( _
            1, _
            wb.Name, _
            "invoice data", _
            vbTextCompare) > 0 Then

            Set invoiceWB = wb

            Exit For

        End If

    Next wb

    '====================================================
    ' IF NOT OPEN, FIND AND OPEN FROM A16
    '====================================================

    If invoiceWB Is Nothing Then

        Set invoiceWB = _
            FindAndOpenInvoiceData(invoiceFolder)

    End If

    '====================================================
    ' CANNOT FIND INVOICE DATA
    '====================================================

    If invoiceWB Is Nothing Then

        MsgBox _
            "Cannot find the invoice data workbook. Check IT Project folder if it is not in the folder stated below." & _
            vbCrLf & vbCrLf & _
            "Folder checked:" & _
            vbCrLf & _
            invoiceFolder & _
            vbCrLf & vbCrLf & _
            "Please check Settings!A16.", _
            vbExclamation

        Exit Sub

    End If

    '====================================================
    ' FIND VGMJ SHEET
    '====================================================

    Set wsVGMJ = Nothing

    On Error Resume Next

    Set wsVGMJ = _
        invoiceWB.Worksheets("VGMJ")

    On Error GoTo 0

    If wsVGMJ Is Nothing Then

        MsgBox _
            "Cannot find sheet 'VGMJ' in:" & _
            vbCrLf & vbCrLf & _
            invoiceWB.Name, _
            vbCritical

        Exit Sub

    End If

    '====================================================
    ' GET OUTPUT FOLDER
    '====================================================

    outputFolder = _
        GetVGMJOutputFolder()

    If outputFolder = "" Then Exit Sub

    '====================================================
    ' PROCESS VGMJ
    '====================================================

    ProcessVGMJInvoices _
        wsVGMJ, _
        outputFolder

    '====================================================
    ' COMPLETED
    '====================================================

    MsgBox _
        "VGMJ generation completed." & _
        vbCrLf & vbCrLf & _
        "Output folder:" & _
        vbCrLf & _
        outputFolder, _
        vbInformation

End Sub


'====================================================================
' GET SETTINGS PATH
'====================================================================

Private Function GetSettingsPath( _
    ByVal storedPath As String) As String

    Dim wsSettings As Worksheet
    Dim username As String
    Dim cleanPath As String
    Dim resultPath As String

    Set wsSettings = _
        ThisWorkbook.Worksheets("Settings")

    username = _
        Trim(CStr( _
            wsSettings.Range("A5").Value))

    cleanPath = _
        Trim(storedPath)

    If username = "" Then

        MsgBox _
            "Settings!A5 is empty." & _
            vbCrLf & vbCrLf & _
            "Please enter the Windows username.", _
            vbCritical

        Exit Function

    End If

    If cleanPath = "" Then Exit Function

    '====================================================
    ' ALREADY FULL WINDOWS PATH
    '====================================================

    If Len(cleanPath) >= 2 Then

        If Mid(cleanPath, 2, 1) = ":" Then

            GetSettingsPath = _
                RemoveTrailingSlash(cleanPath)

            Exit Function

        End If

    End If

    '====================================================
    ' REMOVE LEADING SLASH
    '====================================================

    Do While Left(cleanPath, 1) = "\"

        cleanPath = _
            Mid(cleanPath, 2)

    Loop

    '====================================================
    ' BUILD FULL PATH
    '====================================================

    resultPath = _
        "C:\Users\" & _
        username & _
        "\" & _
        cleanPath

    GetSettingsPath = _
        RemoveTrailingSlash(resultPath)

End Function


'====================================================================
' REMOVE TRAILING SLASH
'====================================================================

Private Function RemoveTrailingSlash( _
    ByVal folderPath As String) As String

    Do While Len(folderPath) > 3 And _
             (Right(folderPath, 1) = "\" Or _
              Right(folderPath, 1) = "/")

        folderPath = _
            Left( _
                folderPath, _
                Len(folderPath) - 1)

    Loop

    RemoveTrailingSlash = _
        folderPath

End Function


'====================================================================
' FIND AND OPEN INVOICE DATA
'====================================================================

Private Function FindAndOpenInvoiceData( _
    ByVal invoiceFolder As String) As Workbook

    Dim fileName As String
    Dim fullPath As String
    Dim wb As Workbook

    Set FindAndOpenInvoiceData = Nothing

    '====================================================
    ' CHECK FOLDER
    '====================================================

    If Dir(invoiceFolder, vbDirectory) = "" Then

        Exit Function

    End If

    '====================================================
    ' XLSX
    '====================================================

    fileName = _
        Dir( _
            invoiceFolder & _
            "\*invoice data*.xlsx")

    '====================================================
    ' XLSM
    '====================================================

    If fileName = "" Then

        fileName = _
            Dir( _
                invoiceFolder & _
                "\*invoice data*.xlsm")

    End If

    '====================================================
    ' XLS
    '====================================================

    If fileName = "" Then

        fileName = _
            Dir( _
                invoiceFolder & _
                "\*invoice data*.xls")

    End If

    If fileName = "" Then

        Exit Function

    End If

    fullPath = _
        invoiceFolder & _
        "\" & _
        fileName

    '====================================================
    ' CHECK IF ALREADY OPEN
    '====================================================

    For Each wb In Application.Workbooks

        If StrComp( _
            wb.FullName, _
            fullPath, _
            vbTextCompare) = 0 Then

            Set FindAndOpenInvoiceData = wb

            Exit Function

        End If

    Next wb

    '====================================================
    ' OPEN WORKBOOK
    '====================================================

    On Error GoTo OpenError

    Set FindAndOpenInvoiceData = _
        Workbooks.Open(fullPath)

    Exit Function

OpenError:

    Set FindAndOpenInvoiceData = Nothing

End Function


'====================================================================
' PROCESS VGMJ INVOICES
'====================================================================

Private Sub ProcessVGMJInvoices( _
    ByVal wsData As Worksheet, _
    ByVal outputFolder As String)

    Dim wsClientSetting As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim clientName As String
    Dim invoiceRef As String

    Dim startRow As Long
    Dim endRow As Long

    Set wsClientSetting = _
        ThisWorkbook.Worksheets("Client Setting")

    '====================================================
    ' FIND LAST ROW
    ' COLUMN T = CLIENT
    '====================================================

    lastRow = _
        wsData.Cells( _
            wsData.Rows.Count, _
            "T").End(xlUp).Row

    If lastRow < 2 Then

        MsgBox _
            "No invoice data was found in the VGMJ sheet.", _
            vbExclamation

        Exit Sub

    End If

    '====================================================
    ' START ROW 2
    '====================================================

    r = 2

    Do While r <= lastRow

        '================================================
        ' CHECK INVOICE REF IN AH
        '================================================

        If Trim(CStr( _
            wsData.Cells(r, "AH").Value)) <> "" Then

            invoiceRef = _
                Trim(CStr( _
                    wsData.Cells(r, "AH").Value))

            clientName = _
                Trim(CStr( _
                    wsData.Cells(r, "T").Value))

            startRow = r
            endRow = r

            '================================================
            ' FIND ALL ROWS FOR THIS INVOICE
            '================================================

            Do While endRow < lastRow

                If Trim(CStr( _
                    wsData.Cells( _
                        endRow + 1, _
                        "AH").Value)) <> "" Then

                    Exit Do

                End If

                If StrComp( _
                    Trim(CStr( _
                        wsData.Cells( _
                            endRow + 1, _
                            "T").Value)), _
                    clientName, _
                    vbTextCompare) <> 0 Then

                    Exit Do

                End If

                endRow = _
                    endRow + 1

            Loop

            '================================================
            ' GENERATE NORMAL VGMJ EXCEL
            '================================================

            GenerateNormalVGMJExcel _
                clientName, _
                invoiceRef, _
                wsData, _
                startRow, _
                endRow, _
                outputFolder

            '================================================
            ' GENERATE TRADE EXCEL IF REQUIRED
            '================================================

            GenerateTradeExcelIfRequired _
                wsClientSetting, _
                clientName, _
                invoiceRef, _
                wsData, _
                startRow, _
                endRow, _
                outputFolder

            r = _
                endRow + 1

        Else

            r = _
                r + 1

        End If

    Loop

End Sub


'====================================================================
' GENERATE NORMAL VGMJ EXCEL
'====================================================================

Private Sub GenerateNormalVGMJExcel( _
    ByVal clientName As String, _
    ByVal invoiceRef As String, _
    ByVal wsData As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal outputFolder As String)

    Dim settingsWS As Worksheet

    Dim templatePath As String
    Dim newFile As String

    Dim exportWB As Workbook

    Dim wsEnglish As Worksheet
    Dim wsJapanese As Worksheet

    Dim shortName As String
    Dim invoiceDate As Date

    On Error GoTo ErrorHandler

    Set settingsWS = _
        ThisWorkbook.Worksheets("Settings")

    '====================================================
    ' GET INVOICE DATE FROM SETTINGS A2
    '====================================================

    If Not IsDate( _
        settingsWS.Range("A2").Value) Then

        Err.Raise _
            vbObjectError + 3000, , _
            "Invalid invoice date in Settings!A2."

    End If

    invoiceDate = _
        CDate( _
            settingsWS.Range("A2").Value)

    '====================================================
    ' GET SHORT NAME
    '====================================================

    shortName = _
        GetClientShortName( _
            clientName, _
            wsData, _
            startRow)

    '====================================================
    ' GET NORMAL TEMPLATE
    '====================================================

    templatePath = _
        GetSettingsPath( _
            CStr( _
                settingsWS.Range("A22").Value))

    If templatePath = "" Then

        Err.Raise _
            vbObjectError + 3001, , _
            "Settings!A22 is empty."

    End If

    '====================================================
    ' CHECK TEMPLATE
    '====================================================

    If Dir(templatePath) = "" Then

        Err.Raise _
            vbObjectError + 3002, , _
            "Normal VGMJ template not found:" & _
            vbCrLf & vbCrLf & _
            templatePath & _
            vbCrLf & vbCrLf & _
            "Please check Settings!A22."

    End If

    '====================================================
    ' BUILD FILE NAME
    '====================================================

    If shortName <> "" Then

        newFile = _
            outputFolder & "\" & _
            CleanFileName(invoiceRef) & "_" & _
            CleanFileName(shortName) & _
            ".xlsx"

    Else

        newFile = _
            outputFolder & "\" & _
            CleanFileName(invoiceRef) & _
            ".xlsx"

    End If

    '====================================================
    ' IF FILE ALREADY EXISTS, SKIP
    '====================================================

    If Dir(newFile) <> "" Then

        Exit Sub

    End If

    '====================================================
    ' COPY TEMPLATE
    '====================================================

    FileCopy _
        templatePath, _
        newFile

    If Dir(newFile) = "" Then

        Err.Raise _
            vbObjectError + 3003, , _
            "The normal VGMJ template could not be copied."

    End If

    '====================================================
    ' OPEN COPIED WORKBOOK
    '====================================================

    Set exportWB = _
        Workbooks.Open(newFile)

    '====================================================
    ' FIND ENGLISH SHEET
    '====================================================

    Set wsEnglish = Nothing

    On Error Resume Next

    Set wsEnglish = _
        exportWB.Worksheets("English")

    On Error GoTo ErrorHandler

    If wsEnglish Is Nothing Then

        Err.Raise _
            vbObjectError + 3004, , _
            "Cannot find sheet 'English' in:" & _
            vbCrLf & _
            newFile

    End If

    '====================================================
    ' FIND JAPANESE SHEET
    '====================================================

    Set wsJapanese = Nothing

    On Error Resume Next

    Set wsJapanese = _
        exportWB.Worksheets("Japanese")

    On Error GoTo ErrorHandler

    If wsJapanese Is Nothing Then

        Err.Raise _
            vbObjectError + 3005, , _
            "Cannot find sheet 'Japanese' in:" & _
            vbCrLf & _
            newFile

    End If

    '====================================================
    ' POPULATE ENGLISH
    '====================================================

    PopulateVGMJSheet _
        wsEnglish, _
        clientName, _
        invoiceDate, _
        wsData, _
        startRow, _
        endRow

    '====================================================
    ' POPULATE JAPANESE
    '====================================================

    PopulateVGMJSheet _
        wsJapanese, _
        clientName, _
        invoiceDate, _
        wsData, _
        startRow, _
        endRow

    '====================================================
    ' SAVE
    '====================================================

    exportWB.Save

    exportWB.Close _
        SaveChanges:=False

    Set exportWB = Nothing

    Exit Sub

ErrorHandler:

    On Error Resume Next

    If Not exportWB Is Nothing Then

        exportWB.Close _
            SaveChanges:=False

        Set exportWB = Nothing

    End If

    MsgBox _
        "Normal VGMJ Excel generation failed." & _
        vbCrLf & vbCrLf & _
        "Client: " & clientName & _
        vbCrLf & _
        "Invoice: " & invoiceRef & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & _
        vbCrLf & _
        Err.Description, _
        vbCritical

End Sub


'====================================================================
' GENERATE TRADE EXCEL IF REQUIRED
'====================================================================

Private Sub GenerateTradeExcelIfRequired( _
    ByVal clientSettingWS As Worksheet, _
    ByVal clientName As String, _
    ByVal invoiceRef As String, _
    ByVal wsData As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal outputFolder As String)

    Dim settingsWS As Worksheet

    Dim templatePath As String
    Dim newFile As String

    Dim exportWB As Workbook
    Dim exportWS As Worksheet

    Dim shortName As String
    Dim exportExcel As Boolean

    Dim invoiceDate As Date
    Dim tradeSheetName As String

    On Error GoTo ErrorHandler

    Set settingsWS = _
        ThisWorkbook.Worksheets("Settings")

    '====================================================
    ' GET INVOICE DATE
    '====================================================

    If Not IsDate( _
        settingsWS.Range("A2").Value) Then

        Err.Raise _
            vbObjectError + 3999, , _
            "Invalid invoice date in Settings!A2."

    End If

    invoiceDate = _
        CDate( _
            settingsWS.Range("A2").Value)

    '====================================================
    ' TRADE SHEET NAME
    '====================================================

    tradeSheetName = _
        Format( _
            invoiceDate, _
            "mmmm yyyy")

    '====================================================
    ' FIND CLIENT SETTING
    '====================================================

    If Not GetTradeClientSetting( _
        clientSettingWS, _
        clientName, _
        wsData, _
        startRow, _
        exportExcel, _
        shortName) Then

        Exit Sub

    End If

    If exportExcel = False Then

        Exit Sub

    End If

    '====================================================
    ' GET TRADE TEMPLATE FROM A13
    '====================================================

    templatePath = _
        GetSettingsPath( _
            CStr( _
                settingsWS.Range("A13").Value))

    If templatePath = "" Then

        Err.Raise _
            vbObjectError + 4000, , _
            "Settings!A13 is empty."

    End If

    If Dir(templatePath) = "" Then

        Err.Raise _
            vbObjectError + 4001, , _
            "Trade Excel template not found:" & _
            vbCrLf & vbCrLf & _
            templatePath & _
            vbCrLf & vbCrLf & _
            "Please check Settings!A13."

    End If

    '====================================================
    ' BUILD TRADE FILE NAME
    '====================================================

    If shortName <> "" Then

        newFile = _
            outputFolder & "\" & _
            CleanFileName(invoiceRef) & "_" & _
            CleanFileName(shortName) & _
            "_trade.xlsx"

    Else

        newFile = _
            outputFolder & "\" & _
            CleanFileName(invoiceRef) & _
            "_trade.xlsx"

    End If

    If Dir(newFile) <> "" Then

        Exit Sub

    End If

    '====================================================
    ' COPY TRADE TEMPLATE
    '====================================================

    FileCopy _
        templatePath, _
        newFile

    If Dir(newFile) = "" Then

        Err.Raise _
            vbObjectError + 4002, , _
            "The Trade Excel template could not be copied."

    End If

    Set exportWB = _
        Workbooks.Open(newFile)

    Set exportWS = _
        exportWB.Worksheets(1)

    exportWS.Name = _
        GetSafeSheetName( _
            tradeSheetName, _
            exportWB, _
            exportWS)

    PopulateTradeSheet _
        exportWS, _
        wsData, _
        startRow, _
        endRow

    exportWB.Save

    exportWB.Close _
        SaveChanges:=False

    Set exportWB = Nothing

    Exit Sub

ErrorHandler:

    On Error Resume Next

    If Not exportWB Is Nothing Then

        exportWB.Close _
            SaveChanges:=False

        Set exportWB = Nothing

    End If

    MsgBox _
        "Trade Excel generation failed." & _
        vbCrLf & vbCrLf & _
        "Client: " & clientName & _
        vbCrLf & _
        "Invoice: " & invoiceRef & _
        vbCrLf & vbCrLf & _
        "Trade template:" & _
        vbCrLf & _
        templatePath & _
        vbCrLf & vbCrLf & _
        "Output file:" & _
        vbCrLf & _
        newFile & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & _
        vbCrLf & _
        Err.Description, _
        vbCritical

End Sub


'====================================================================
' GET SAFE SHEET NAME
'====================================================================

Private Function GetSafeSheetName( _
    ByVal requestedName As String, _
    ByVal wb As Workbook, _
    ByVal currentSheet As Worksheet) As String

    Dim cleanName As String
    Dim testName As String
    Dim counter As Long

    cleanName = Trim(requestedName)

    cleanName = Replace(cleanName, ":", "")
    cleanName = Replace(cleanName, "\", "")
    cleanName = Replace(cleanName, "/", "")
    cleanName = Replace(cleanName, "?", "")
    cleanName = Replace(cleanName, "*", "")
    cleanName = Replace(cleanName, "[", "")
    cleanName = Replace(cleanName, "]", "")

    If Len(cleanName) > 31 Then

        cleanName = _
            Left(cleanName, 31)

    End If

    If cleanName = "" Then

        cleanName = "Trade"

    End If

    testName = cleanName
    counter = 1

    Do While SheetNameExists( _
        wb, _
        testName, _
        currentSheet)

        counter = counter + 1

        testName = _
            Left( _
                cleanName, _
                31 - Len(CStr(counter)) - 1) & _
            "_" & _
            CStr(counter)

    Loop

    GetSafeSheetName = testName

End Function


'====================================================================
' CHECK SHEET NAME EXISTS
'====================================================================

Private Function SheetNameExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String, _
    ByVal currentSheet As Worksheet) As Boolean

    Dim ws As Worksheet

    SheetNameExists = False

    For Each ws In wb.Worksheets

        If Not ws Is currentSheet Then

            If StrComp( _
                ws.Name, _
                sheetName, _
                vbTextCompare) = 0 Then

                SheetNameExists = True

                Exit Function

            End If

        End If

    Next ws

End Function


'====================================================================
' GET TRADE CLIENT SETTING
'====================================================================

Private Function GetTradeClientSetting( _
    ByVal ws As Worksheet, _
    ByVal clientName As String, _
    ByVal wsData As Worksheet, _
    ByVal startRow As Long, _
    ByRef exportExcel As Boolean, _
    ByRef shortName As String) As Boolean

    Dim r As Long
    Dim lastRow As Long

    Dim invoiceSplitType As String
    Dim settingSplitType As String

    exportExcel = False
    shortName = ""

    GetTradeClientSetting = False

    invoiceSplitType = _
        GetSplitType( _
            CStr( _
                wsData.Cells( _
                    startRow, _
                    "C").Value))

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            "A").End(xlUp).Row

    '====================================================
    ' FIRST: CLIENT + MATCHING SPLIT
    '====================================================

    For r = 2 To lastRow

        If StrComp( _
            Trim(CStr( _
                ws.Cells(r, "A").Value)), _
            Trim(clientName), _
            vbTextCompare) = 0 Then

            settingSplitType = _
                UCase(Trim(CStr( _
                    ws.Cells(r, "E").Value)))

            If settingSplitType = _
                UCase(invoiceSplitType) Then

                exportExcel = _
                    IsSettingTrue( _
                        ws.Cells(r, "B").Value)

                shortName = _
                    Trim(CStr( _
                        ws.Cells(r, "C").Value))

                GetTradeClientSetting = True

                Exit Function

            End If

        End If

    Next r

    '====================================================
    ' SECOND: CLIENT + BLANK SPLIT
    '====================================================

    For r = 2 To lastRow

        If StrComp( _
            Trim(CStr( _
                ws.Cells(r, "A").Value)), _
            Trim(clientName), _
            vbTextCompare) = 0 Then

            settingSplitType = _
                Trim(CStr( _
                    ws.Cells(r, "E").Value))

            If settingSplitType = "" Then

                exportExcel = _
                    IsSettingTrue( _
                        ws.Cells(r, "B").Value)

                shortName = _
                    Trim(CStr( _
                        ws.Cells(r, "C").Value))

                GetTradeClientSetting = True

                Exit Function

            End If

        End If

    Next r

End Function


'====================================================================
' POPULATE TRADE SHEET
'====================================================================

Private Sub PopulateTradeSheet( _
    ByVal ws As Worksheet, _
    ByVal wsData As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long)

    Dim rowCount As Long

    rowCount = _
        endRow - startRow + 1

    If ws.Rows.Count >= 2 Then

        ws.Rows("2:" & ws.Rows.Count).ClearContents

    End If

    ws.Range( _
        "A2:S" & _
        rowCount + 1).Value = _
            wsData.Range( _
                "A" & startRow & _
                ":S" & endRow).Value

End Sub


'====================================================================
' GET CLIENT SHORT NAME
'====================================================================

Private Function GetClientShortName( _
    ByVal clientName As String, _
    ByVal wsData As Worksheet, _
    ByVal startRow As Long) As String

    Dim ws As Worksheet

    Dim r As Long
    Dim lastRow As Long

    Dim invoiceSplitType As String
    Dim settingSplitType As String

    Set ws = _
        ThisWorkbook.Worksheets("Client Setting")

    GetClientShortName = ""

    invoiceSplitType = _
        GetSplitType( _
            CStr( _
                wsData.Cells( _
                    startRow, _
                    "C").Value))

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            "A").End(xlUp).Row

    '====================================================
    ' FIRST: CLIENT + SPLIT TYPE
    '====================================================

    For r = 2 To lastRow

        If StrComp( _
            Trim(CStr( _
                ws.Cells(r, "A").Value)), _
            Trim(clientName), _
            vbTextCompare) = 0 Then

            settingSplitType = _
                UCase(Trim(CStr( _
                    ws.Cells(r, "E").Value)))

            If settingSplitType = _
                UCase(invoiceSplitType) Then

                GetClientShortName = _
                    Trim(CStr( _
                        ws.Cells(r, "C").Value))

                Exit Function

            End If

        End If

    Next r

    '====================================================
    ' SECOND: NORMAL CLIENT
    '====================================================

    For r = 2 To lastRow

        If StrComp( _
            Trim(CStr( _
                ws.Cells(r, "A").Value)), _
            Trim(clientName), _
            vbTextCompare) = 0 Then

            settingSplitType = _
                Trim(CStr( _
                    ws.Cells(r, "E").Value))

            If settingSplitType = "" Then

                GetClientShortName = _
                    Trim(CStr( _
                        ws.Cells(r, "C").Value))

                Exit Function

            End If

        End If

    Next r

End Function


'====================================================================
' CHECK WHETHER CLIENT SETTING COLUMN B IS TRUE
'====================================================================

Private Function IsSettingTrue( _
    ByVal Value As Variant) As Boolean

    If IsError(Value) Then

        IsSettingTrue = False

        Exit Function

    End If

    If VarType(Value) = vbBoolean Then

        IsSettingTrue = CBool(Value)

        Exit Function

    End If

    Select Case UCase(Trim(CStr(Value)))

        Case "TRUE", "YES", "Y", "1"

            IsSettingTrue = True

        Case Else

            IsSettingTrue = False

    End Select

End Function


'====================================================================
' GET VGMJ OUTPUT FOLDER
'====================================================================

Private Function GetVGMJOutputFolder() As String

    Dim wsSettings As Worksheet

    Dim basePath As String
    Dim vgmjPath As String
    Dim monthPath As String

    Dim invoiceDate As Date
    Dim monthFolder As String

    Dim fso As Object

    Set wsSettings = _
        ThisWorkbook.Worksheets("Settings")

    Set fso = _
        CreateObject( _
            "Scripting.FileSystemObject")

    basePath = _
        GetSettingsPath( _
            CStr( _
                wsSettings.Range("A16").Value))

    If basePath = "" Then

        MsgBox _
            "Settings!A16 is empty or invalid.", _
            vbCritical

        Exit Function

    End If

    '====================================================
    ' INVOICE DATE FROM A2
    '====================================================

    If Not IsDate( _
        wsSettings.Range("A2").Value) Then

        MsgBox _
            "Invalid invoice date in Settings!A2.", _
            vbCritical

        Exit Function

    End If

    invoiceDate = _
        CDate( _
            wsSettings.Range("A2").Value)

    If Not fso.FolderExists(basePath) Then

        CreateFolderRecursive _
            fso, _
            basePath

    End If

    vgmjPath = _
        fso.BuildPath( _
            basePath, _
            "VGMJ")

    If Not fso.FolderExists(vgmjPath) Then

        fso.CreateFolder _
            vgmjPath

    End If

    monthFolder = _
        Format( _
            invoiceDate, _
            "yyyymm") & _
        " " & _
        Format( _
            invoiceDate, _
            "mmm")

    monthPath = _
        fso.BuildPath( _
            vgmjPath, _
            monthFolder)

    If Not fso.FolderExists(monthPath) Then

        fso.CreateFolder _
            monthPath

    End If

    GetVGMJOutputFolder = _
        monthPath

End Function


'====================================================================
' CREATE FOLDER RECURSIVELY
'====================================================================

Private Sub CreateFolderRecursive( _
    ByVal fso As Object, _
    ByVal folderPath As String)

    Dim parentPath As String

    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = _
        fso.GetParentFolderName( _
            folderPath)

    If parentPath <> "" Then

        If Not fso.FolderExists(parentPath) Then

            CreateFolderRecursive _
                fso, _
                parentPath

        End If

    End If

    If Not fso.FolderExists(folderPath) Then

        fso.CreateFolder _
            folderPath

    End If

End Sub


'====================================================================
' CLEAN FILE NAME
'====================================================================

Private Function CleanFileName( _
    ByVal text As String) As String

    Dim badCharacters As Variant
    Dim i As Long

    badCharacters = _
        Array( _
            "\", _
            "/", _
            ":", _
            "*", _
            "?", _
            """", _
            "<", _
            ">", _
            "|")

    CleanFileName = _
        Trim(text)

    For i = _
        LBound(badCharacters) To _
        UBound(badCharacters)

        CleanFileName = _
            Replace( _
                CleanFileName, _
                badCharacters(i), _
                "_")

    Next i

End Function


'====================================================================
' POPULATE ENGLISH / JAPANESE SHEET
'====================================================================

Private Sub PopulateVGMJSheet( _
    ByVal ws As Worksheet, _
    ByVal clientName As String, _
    ByVal invoiceDate As Date, _
    ByVal wsData As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long)

    Dim headerCell As Range
    Dim totalCell As Range

    Dim headerRow As Long
    Dim firstDataRow As Long

    Dim tradeCount As Long

    Dim r As Long
    Dim destRow As Long

    Dim isJapanese As Boolean
    Dim fileRefHeader As String

    '====================================================
    ' JAPANESE PAYMENT DUE DATE VARIABLES
    '====================================================

    Dim paymentDueLabel As Range
    Dim paymentDueCell As Range
    Dim paymentDueDate As Date
    Dim paymentDueText As String

    '====================================================
    ' DETERMINE ENGLISH / JAPANESE
    '====================================================

    isJapanese = _
        (StrComp( _
            ws.Name, _
            "Japanese", _
            vbTextCompare) = 0)

    '====================================================
    ' ENTITY
    '====================================================

    If isJapanese Then

        ws.Range("B4").Value = _
            clientName

    Else

        ws.Range("B2").Value = _
            clientName

    End If

    '====================================================
    ' INVOICE DATE
    '
    ' Uses Settings!A2 via invoiceDate.
    '====================================================

    If isJapanese Then

        ws.Range("K11").Value = _
            ChrW(&H8ACB) & _
            ChrW(&H6C42) & _
            ChrW(&H66F8) & _
            ChrW(&H767A) & _
            ChrW(&H884C) & _
            ChrW(&H65E5) & _
            ChrW(&HFF1A) & _
            Year(invoiceDate) & _
            ChrW(&H5E74) & _
            Month(invoiceDate) & _
            ChrW(&H6708) & _
            Day(invoiceDate) & _
            ChrW(&H65E5)

    Else

        ws.Range("K9").Value = _
            "Invoice Date:" & _
            Format( _
                invoiceDate, _
                "dd-mmm-yy")

    End If

    '====================================================
    ' INVOICE NUMBER
    '====================================================

    If isJapanese Then

        ws.Range("K10").Value = _
            ChrW(&H8ACB) & _
            ChrW(&H6C42) & _
            ChrW(&H66F8) & _
            ChrW(&H756A) & _
            ChrW(&H53F7) & _
            ChrW(&HFF1A) & _
            "VGMJ " & _
            Format( _
                invoiceDate, _
                "yyyymm") & _
            "001"

    Else

        ws.Range("K8").Value = _
            "Invoice Number:VGMJ " & _
            Format( _
                invoiceDate, _
                "yyyymm") & _
            "001"

    End If

    '====================================================
    ' PAYMENT DUE DATE
    '
    ' IMPORTANT:
    ' Payment Due Date uses the LAST DAY of the
    ' month containing Settings!A2.
    '
    ' Examples:
    ' Settings!A2 = 01-Sep-2026 -> 30-Sep-26
    ' Settings!A2 = 15-Oct-2026 -> 31-Oct-26
    '
    ' English:
    '   format = dd-mmm-yy
    '
    ' Japanese:
    '   same underlying month-end date,
    '   while preserving the Japanese cell format.
    '
    ' IMPORTANT:
    ' Japanese C16 is NOT touched because it contains
    ' the invoice amount formula such as:
    '
    ' =CONCATENATE("¥",TEXT(L40,"#,##0"),"-")
    '====================================================

    paymentDueDate = _
        DateSerial( _
            Year(invoiceDate), _
            Month(invoiceDate) + 1, _
            0)

    If Not isJapanese Then

        '================================================
        ' ENGLISH PAYMENT DUE DATE
        '================================================

        ws.Range("C16").Value = _
            paymentDueDate

        ws.Range("C16").NumberFormat = _
            "dd-mmm-yy"

    Else

        '================================================
        ' JAPANESE TEXT:
        '
        ' ?????
        '================================================

        paymentDueText = _
            ChrW(&H304A) & _
            ChrW(&H652F) & _
            ChrW(&H6255) & _
            ChrW(&H671F) & _
            ChrW(&H9650)

        Set paymentDueLabel = Nothing
        Set paymentDueCell = Nothing

        Set paymentDueLabel = _
            ws.Cells.Find( _
                What:=paymentDueText, _
                After:=ws.Range("A1"), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False)

        If Not paymentDueLabel Is Nothing Then

            '================================================
            ' FIND CELL IMMEDIATELY TO THE RIGHT
            '================================================

            If paymentDueLabel.MergeCells Then

                Set paymentDueCell = _
                    paymentDueLabel.MergeArea.Cells( _
                        1, _
                        paymentDueLabel.MergeArea.Columns.Count).Offset( _
                            0, _
                            1)

            Else

                Set paymentDueCell = _
                    paymentDueLabel.Offset(0, 1)

            End If

            '================================================
            ' DESTINATION MAY ALSO BE MERGED
            '================================================

            If paymentDueCell.MergeCells Then

                Set paymentDueCell = _
                    paymentDueCell.MergeArea.Cells(1, 1)

            End If

            '================================================
            ' WRITE SAME DUE DATE AS ENGLISH
            '================================================

            paymentDueCell.Value = _
                paymentDueDate

        Else

            MsgBox _
                "Cannot find the Japanese payment due date field " & _
                "in sheet '" & ws.Name & "'.", _
                vbExclamation

        End If

    End If

    '====================================================
    ' UPDATE INVOICE DESCRIPTION MONTH
    '
    ' Both English and Japanese use Settings!A2
    ' via invoiceDate.
    '====================================================

    UpdateInvoiceNarrative _
        ws, _
        invoiceDate, _
        isJapanese

    '====================================================
    ' FILE REF HEADER
    '====================================================

    If isJapanese Then

        fileRefHeader = _
            ChrW(&H6210) & _
            ChrW(&H7D04) & _
            ChrW(&H756A) & _
            ChrW(&H53F7)

    Else

        fileRefHeader = _
            "File Ref"

    End If

    '====================================================
    ' FIND FILE REF HEADER
    '====================================================

    Set headerCell = _
        ws.Cells.Find( _
            What:=fileRefHeader, _
            After:=ws.Range("A1"), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)

    If headerCell Is Nothing Then

        MsgBox _
            "Cannot find '" & _
            fileRefHeader & _
            "' header in sheet '" & _
            ws.Name & _
            "'.", _
            vbCritical

        Exit Sub

    End If

    headerRow = _
        headerCell.Row

    firstDataRow = _
        headerRow + 1

    tradeCount = _
        endRow - startRow + 1

    '====================================================
    ' FIND TOTAL ROW
    '====================================================

    Set totalCell = _
        ws.Cells.Find( _
            What:="Total", _
            After:=ws.Range("A1"), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)

    '====================================================
    ' DELETE OLD DETAIL ROWS
    '====================================================

    If Not totalCell Is Nothing Then

        If totalCell.Row > firstDataRow Then

            ws.Rows( _
                firstDataRow & ":" & _
                totalCell.Row - 1).Delete

        End If

    End If

    '====================================================
    ' INSERT REQUIRED ROWS
    '====================================================

    If tradeCount > 1 Then

        ws.Rows( _
            firstDataRow + 1 & ":" & _
            firstDataRow + tradeCount - 1).Insert _
                Shift:=xlDown, _
                CopyOrigin:=xlFormatFromLeftOrAbove

    End If

    '====================================================
    ' COPY DATA
    '====================================================

    For r = startRow To endRow

        destRow = _
            firstDataRow + _
            (r - startRow)

        '================================================
        ' A - FILE REF
        '================================================

        ws.Cells( _
            destRow, _
            "A").Value = _
                wsData.Cells( _
                    r, _
                    "A").Value

        '================================================
        ' B - CONTRACT DATE
        '================================================

        ws.Cells( _
            destRow, _
            "B").Value = _
                wsData.Cells( _
                    r, _
                    "B").Value

        ws.Cells( _
            destRow, _
            "B").NumberFormat = _
                "dd-mmm-yy"

        '================================================
        ' C - PRODUCT
        '================================================

        ws.Cells( _
            destRow, _
            "C").Value = _
                ConvertVGMJProduct( _
                    CStr( _
                        wsData.Cells( _
                            r, _
                            "C").Value))

        '================================================
        ' D - CONTRACT
        '================================================

        ws.Cells( _
            destRow, _
            "D").Value = _
                wsData.Cells( _
                    r, _
                    "D").Value

        '================================================
        ' E - BUY / SELL
        '================================================

        ws.Cells( _
            destRow, _
            "E").Value = _
                wsData.Cells( _
                    r, _
                    "E").Value

        '================================================
        ' F - C/P
        '================================================

        ws.Cells( _
            destRow, _
            "F").Value = _
                wsData.Cells( _
                    r, _
                    "F").Value

        '================================================
        ' G - QUANTITY
        '================================================

        ws.Cells( _
            destRow, _
            "G").Value = _
                wsData.Cells( _
                    r, _
                    "G").Value

        '================================================
        ' H - UNIT
        '================================================

        ws.Cells( _
            destRow, _
            "H").Value = _
                "JPY/kWh"

        '================================================
        ' I - UNIT
        '================================================

        ws.Cells( _
            destRow, _
            "I").Value = _
                "JPY/kWh"

        '================================================
        ' J - TAX
        '================================================

        ws.Cells( _
            destRow, _
            "J").Value = 0

        ws.Cells( _
            destRow, _
            "J").NumberFormat = _
                "0%"

        '================================================
        ' K - BLANK
        '================================================

        ws.Cells( _
            destRow, _
            "K").ClearContents

        '================================================
        ' L - AMOUNT
        '================================================

        ws.Cells( _
            destRow, _
            "L").Value = _
                wsData.Cells( _
                    r, _
                    "S").Value

        '================================================
        ' M - CURRENCY
        '================================================

        If isJapanese Then

            ws.Cells( _
                destRow, _
                "M").Value = _
                    ChrW(&H5186)      ' ?

        Else

            ws.Cells( _
                destRow, _
                "M").Value = _
                    "JPY"

        End If

    Next r

End Sub


'====================================================================
' UPDATE INVOICE NARRATIVE
'
' Uses invoiceDate, which comes from Settings!A2.
'
' ENGLISH EXAMPLE:
'
' We hereby invoice you for the fees related to your contract for
' September 2026 as detailed below. Kindly review and process the
' payment accordingly.
'
' JAPANESE EXAMPLE:
'
' 2026?9??????????????????????
' ???????????????????????
'====================================================================

Private Sub UpdateInvoiceNarrative( _
    ByVal ws As Worksheet, _
    ByVal invoiceDate As Date, _
    ByVal isJapanese As Boolean)

    Dim narrativeCell As Range
    Dim searchText As String
    Dim japaneseNarrative As String
    Dim contractMonth As Date

    Set narrativeCell = Nothing

    '====================================================
    ' CONTRACT MONTH
    '
    ' The invoice narrative describes the PREVIOUS month.
    '
    ' Example:
    ' Settings!A2 = 01-Oct-2026
    ' Contract month = September 2026
    '====================================================

    contractMonth = _
        DateAdd( _
            "m", _
            -1, _
            invoiceDate)

    If Not isJapanese Then

        '================================================
        ' ENGLISH
        '================================================

        searchText = _
            "We hereby invoice you for the fees related to your contract for"

        Set narrativeCell = _
            ws.Cells.Find( _
                What:=searchText, _
                After:=ws.Range("A1"), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False)

        If Not narrativeCell Is Nothing Then

            If narrativeCell.MergeCells Then

                Set narrativeCell = _
                    narrativeCell.MergeArea.Cells(1, 1)

            End If

            narrativeCell.Value = _
                "We hereby invoice you for the fees related to your contract for " & _
                Format(contractMonth, "mmmm yyyy") & _
                " as detailed below. Kindly review and process the payment accordingly."

        End If

    Else

        '================================================
        ' JAPANESE
        '
        ' Search for:
        ' ???
        '
        ' Japanese is built with ChrW so it does not turn
        ' into ????? in the VBA editor.
        '================================================

        searchText = _
            ChrW(&H3054) & _
            ChrW(&H6210) & _
            ChrW(&H7D04)

        Set narrativeCell = _
            ws.Cells.Find( _
                What:=searchText, _
                After:=ws.Range("A1"), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False)

        If Not narrativeCell Is Nothing Then

            If narrativeCell.MergeCells Then

                Set narrativeCell = _
                    narrativeCell.MergeArea.Cells(1, 1)

            End If

            japaneseNarrative = _
                CStr(Year(contractMonth)) & _
                ChrW(&H5E74) & _
                CStr(Month(contractMonth)) & _
                ChrW(&H6708) & _
                ChrW(&H5206) & _
                ChrW(&H306E)

            japaneseNarrative = _
                japaneseNarrative & _
                ChrW(&H8CB4) & _
                ChrW(&H793E) & _
                ChrW(&H304C) & _
                ChrW(&H3054) & _
                ChrW(&H6210) & _
                ChrW(&H7D04)

            japaneseNarrative = _
                japaneseNarrative & _
                ChrW(&H3055) & _
                ChrW(&H308C) & _
                ChrW(&H305F) & _
                ChrW(&H8AF8) & _
                ChrW(&H304A) & _
                ChrW(&H624B)

            japaneseNarrative = _
                japaneseNarrative & _
                ChrW(&H6570) & _
                ChrW(&H6599) & _
                ChrW(&H306B) & _
                ChrW(&H3064) & _
                ChrW(&H3044) & _
                ChrW(&H3066)

            japaneseNarrative = _
                japaneseNarrative & _
                ChrW(&H3001) & _
                ChrW(&H4E0B) & _
                ChrW(&H8A18) & _
                ChrW(&H306E) & _
                ChrW(&H901A) & _
                ChrW(&H308A)

            japaneseNarrative = _
                japaneseNarrative & _
                ChrW(&H3054) & _
                ChrW(&H8ACB) & _
                ChrW(&H6C42) & _
                ChrW(&H7533) & _
                ChrW(&H3057) & _
                ChrW(&H4E0A)

            japaneseNarrative = _
                japaneseNarrative & _
                ChrW(&H3052) & _
                ChrW(&H307E) & _
                ChrW(&H3059) & _
                ChrW(&H306E) & _
                ChrW(&H3067) & _
                ChrW(&H3054)

            japaneseNarrative = _
                japaneseNarrative & _
                ChrW(&H67FB) & _
                ChrW(&H53CE) & _
                ChrW(&H4E0B) & _
                ChrW(&H3055) & _
                ChrW(&H3044) & _
                ChrW(&H3002)

            narrativeCell.Value = _
                japaneseNarrative

        End If

    End If

End Sub


'====================================================================
' GET SPLIT TYPE
'====================================================================

Private Function GetSplitType( _
    ByVal product As String) As String

    product = _
        UCase(Trim(product))

    If InStr( _
        product, _
        "BALANCING GROUP") > 0 _
        Or InStr( _
            product, _
            "(BG)") > 0 Then

        GetSplitType = _
            "BG"

    Else

        GetSplitType = _
            "EEX"

    End If

End Function


'====================================================================
' CONVERT VGMJ PRODUCT NAME
'====================================================================

Private Function ConvertVGMJProduct( _
    ByVal product As String) As String

    Dim p As String

    p = _
        UCase(Trim(product))

    '====================================================
    ' TOKYO
    '====================================================

    If InStr( _
        1, _
        p, _
        "TOKYO AREA BASE", _
        vbTextCompare) > 0 Then

        ConvertVGMJProduct = _
            "TBL"

        Exit Function

    End If

    If InStr( _
        1, _
        p, _
        "TOKYO AREA PEAK", _
        vbTextCompare) > 0 Then

        ConvertVGMJProduct = _
            "TPK"

        Exit Function

    End If

    '====================================================
    ' KANSAI
    '====================================================

    If InStr( _
        1, _
        p, _
        "KANSAI AREA BASE", _
        vbTextCompare) > 0 Then

        ConvertVGMJProduct = _
            "KBL"

        Exit Function

    End If

    If InStr( _
        1, _
        p, _
        "KANSAI AREA PEAK", _
        vbTextCompare) > 0 Then

        ConvertVGMJProduct = _
            "KPK"

        Exit Function

    End If

    '====================================================
    ' CHUBU
    '====================================================

    If InStr( _
        1, _
        p, _
        "CHUBU AREA BASE", _
        vbTextCompare) > 0 Then

        ConvertVGMJProduct = _
            "CBL"

        Exit Function

    End If

    If InStr( _
        1, _
        p, _
        "CHUBU AREA PEAK", _
        vbTextCompare) > 0 Then

        ConvertVGMJProduct = _
            "CPK"

        Exit Function

    End If

    '====================================================
    ' NO MATCH
    '====================================================

    ConvertVGMJProduct = _
        product

End Function






