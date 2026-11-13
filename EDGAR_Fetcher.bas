Attribute VB_Name = "EDGAR_Fetcher"
Sub FetchAndPopulateEDGAR()
    Dim ticker As String
    Dim email As String
    Dim shellCmd As String
    Dim scriptPath As String
    Dim csvPath As String
    Dim folderPath As String
    Dim userProfile As String
    Dim wsh As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim ws As Worksheet
    
    ' Use ActiveSheet to ensure we update the sheet the user is currently viewing/editing
    Set ws = ActiveSheet
    
    ' Read ticker directly from cell C7
    ticker = Trim(ws.Range("C7").Value)
    If ticker = "" Then
        ticker = InputBox("No ticker found in cell C7. Enter Stock Ticker (e.g., UBER):", "SEC EDGAR DCF Fetcher", "UBER")
        If ticker = "" Then Exit Sub
        ws.Range("C7").Value = UCase(ticker)
    End If
    
    ' Ask for SEC email (read from cell C8)
    email = Trim(ws.Range("C8").Value)
    If email = "" Or InStr(email, "@") = 0 Then
        email = InputBox("Enter SEC User-Agent Email (required by SEC guidelines to pull data):", "SEC EDGAR DCF Fetcher", "julie@example.com")
        If email = "" Then Exit Sub
        ws.Range("C8").Value = email
    End If
    
    ' Construct absolute paths using the USERPROFILE environment variable.
    ' This bypasses any OneDrive/SharePoint URL path issues
    userProfile = Environ("USERPROFILE")
    folderPath = userProfile & "\.gemini\antigravity\scratch\edgar_dcf_excel"
    scriptPath = folderPath & "\fetch_edgar.py"
    csvPath = folderPath & "\edgar_data.csv"
    
    ' Remove old CSV to prevent stale data loading if Python fails
    On Error Resume Next
    Kill csvPath
    On Error GoTo 0
    
    ' Call python fetch_edgar.py using uv in the user profile directory
    shellCmd = "cmd.exe /c cd /d """ & folderPath & """ && """ & userProfile & "\.local\bin\uv.exe"" run python fetch_edgar.py --ticker " & ticker & " --email """ & email & """"
    
    MsgBox "Fetching SEC EDGAR data for " & ticker & ". Excel will temporarily pause while downloading...", vbInformation, "Fetching..."
    
    Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.Run shellCmd, windowStyle, waitOnReturn
    On Error GoTo 0
    
    ' Check if CSV was created
    If Dir(csvPath) = "" Then
        ' Prompt with details on how to debug if it fails
        MsgBox "Failed to retrieve data from SEC EDGAR." & vbCrLf & vbCrLf & _
               "Please verify:" & vbCrLf & _
               "1. Internet connection is active." & vbCrLf & _
               "2. Python and uv are installed." & vbCrLf & _
               "3. Ticker '" & ticker & "' is correct." & vbCrLf & vbCrLf & _
               "Debug Path checked: " & csvPath, vbCritical, "Error"
        Exit Sub
    End If
    
    ' Load the CSV data and write to active sheet
    Call LoadEdgarCSVIntoSheet(csvPath, ticker)
    
    MsgBox "DCF Model successfully populated for " & ticker & "!", vbInformation, "Success"
End Sub

Sub LoadEdgarCSVIntoSheet(csvPath As String, ticker As String)
    Dim fileNum As Integer
    Dim lineStr As String
    Dim parts() As String
    Dim key As String
    Dim rowNum As Long
    Dim colIdx As Long
    Dim ws As Worksheet
    Dim idx As Integer
    Dim cfRows As Variant
    Dim r As Variant
    Dim capexVals(1 To 5) As Double
    Dim daVals(1 To 5) As Double
    
    ' Use ActiveSheet to ensure we update the sheet the user is currently viewing/editing
    Set ws = ActiveSheet
    
    Application.EnableEvents = False
    On Error GoTo CleanExit
    
    ' Unhide all rows first to reset any previous ticker's hidden rows
    ws.Cells.EntireRow.Hidden = False
    
    ' Update Ticker in C7
    ws.Range("C7").Value = UCase(ticker)
    
    fileNum = FreeFile
    Open csvPath For Input As #fileNum
    
    ' Read Header (Years)
    Line Input #fileNum, lineStr
    parts = Split(lineStr, ",")
    
    For idx = 1 To UBound(parts)
        colIdx = 4 + idx ' Columns E, F, G, H, I
        ws.Cells(10, colIdx).Value = parts(idx) & "A"
        ws.Cells(59, colIdx).Value = parts(idx) & "A"
        ws.Cells(102, colIdx).Value = parts(idx) & "A"
        ws.Cells(121, colIdx).Value = parts(idx) & "A"
    Next idx
    
    ' Read Row values
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineStr
        parts = Split(lineStr, ",")
        If UBound(parts) > 0 Then
            key = Trim(parts(0))
            If key = "CompanyName" Then
                ws.Range("C6").Value = parts(UBound(parts))
            ElseIf key = "LatestQuarterlyCash" Then
                ws.Range("I6").Value = Val(parts(UBound(parts)))
            ElseIf key = "LatestQuarterlyDebt" Then
                ws.Range("I7").Value = Val(parts(UBound(parts)))
            ElseIf key = "LatestQuarterlyShares" Then
                ws.Range("F7").Value = Val(parts(UBound(parts)))
            ElseIf key = "CF_Capex" Then
                For idx = 1 To UBound(parts)
                    If idx <= 5 Then capexVals(idx) = Abs(Val(parts(idx)))
                Next idx
            Else
                If key = "DepreciationAmortizationIS" Then
                    For idx = 1 To UBound(parts)
                        If idx <= 5 Then daVals(idx) = Abs(Val(parts(idx)))
                    Next idx
                End If
                rowNum = GetRowForKey(key)
                If rowNum > 0 Then
                    For idx = 1 To UBound(parts)
                        colIdx = 4 + idx ' Columns E, F, G, H, I
                        ws.Cells(rowNum, colIdx).Value = Val(parts(idx))
                    Next idx
                End If
            End If
        End If
    Loop
    
    Close #fileNum
    
    ' Keep historical Cash Flow Statement input cells (Columns E to I) empty as requested
    cfRows = Array(106, 107, 108, 109, 110, 112, 114, 115, 116, 117)
    For Each r In cfRows
        For colIdx = 5 To 9
            ws.Cells(r, colIdx).ClearContents
        Next colIdx
    Next r
    
    ' Populate PP&E Schedule (Rows 123 to 133) for historical years 2022A to 2025A (Columns F to I)
    ws.Range("E123:E126").ClearContents
    ws.Range("E133").ClearContents
    
    ' Row 123: Beginning of period (BOP = prior year EOP PP&E)
    ws.Range("F123").Formula = "=E73"
    ws.Range("G123").Formula = "=F126"
    ws.Range("H123").Formula = "=G126"
    ws.Range("I123").Formula = "=H126"
    
    ' Row 124: Plus: CapEX (positive historical CapEX for 2022A to 2025A)
    For idx = 2 To 5
        colIdx = 4 + idx ' Columns F, G, H, I
        ws.Cells(124, colIdx).Value = capexVals(idx)
        ws.Cells(125, colIdx).Value = -daVals(idx)
    Next idx
    
    ' Row 126: End of Period (links to Balance Sheet PP&E Row 73 for 2022A to 2025A)
    ws.Range("F126").Formula = "=F73"
    ws.Range("G126").Formula = "=G73"
    ws.Range("H126").Formula = "=H73"
    ws.Range("I126").Formula = "=I73"
    
    ' Row 133: Total D&A (links to Income Statement D&A Row 23 for 2022A to 2025A)
    ws.Range("F133").Formula = "=F23"
    ws.Range("G133").Formula = "=G23"
    ws.Range("H133").Formula = "=H23"
    ws.Range("I133").Formula = "=I23"
    
    ' Ensure Row 132 (% of revenue) in projected columns J to N links to prior column (=I132..=M132)
    ' to prevent circular reference with Row 131 (=J132*J12)
    ws.Range("J132").Formula = "=I132"
    ws.Range("K132").Formula = "=J132"
    ws.Range("L132").Formula = "=K132"
    ws.Range("M132").Formula = "=L132"
    ws.Range("N132").Formula = "=M132"
    
    ' Dynamically scan and hide optional/empty rows (excluding Row 23 D&A so it is never hidden)
    Call AutoHideEmptyRows(ws)

CleanExit:
    Application.EnableEvents = True
End Sub

Sub AutoHideEmptyRows(ws As Worksheet)
    ' List of row numbers that are optional and can be hidden if they contain no data
    Dim rowsToCheck As Variant
    Dim r As Variant
    Dim c As Long
    Dim isEmptyRow As Boolean
    Dim valToCheck As Double
    
    ' Row 23 (D&A) is intentionally excluded so it is never hidden
    rowsToCheck = Array(19, 20, 21, 32, 34, 65, 69, 70, 72, 82, 86, 91, 98)
    
    For Each r In rowsToCheck
        isEmptyRow = True
        
        ' Check columns E, F, G, H, I (columns 5 to 9)
        For c = 5 To 9
            valToCheck = Abs(Val(ws.Cells(r, c).Value))
            If valToCheck > 0.0001 Then
                isEmptyRow = False
                Exit For
            End If
        Next c
        
        ' Apply hide/unhide logic
        If r = 34 Then
            ' If NCI is zero, hide both "Net income incl. NCI" (33) and "Less: Net income attributable to NCI" (34)
            ws.Rows(33).EntireRow.Hidden = isEmptyRow
            ws.Rows(34).EntireRow.Hidden = isEmptyRow
        Else
            ws.Rows(r).EntireRow.Hidden = isEmptyRow
        End If
    Next r
End Sub

Function GetRowForKey(key As String) As Long
    Select Case key
        ' Income Statement
        Case "Revenue": GetRowForKey = 12
        Case "COGS": GetRowForKey = 15
        Case "OperationSupport": GetRowForKey = 19
        Case "SalesMarketing": GetRowForKey = 20
        Case "ResearchDevelopment": GetRowForKey = 21
        Case "GeneralAdministrative": GetRowForKey = 22
        Case "DepreciationAmortizationIS": GetRowForKey = 23
        Case "InterestExpense": GetRowForKey = 27
        Case "OtherIncome": GetRowForKey = 28
        Case "TaxProvision": GetRowForKey = 31
        Case "EquityInvestmentIncome": GetRowForKey = 32
        Case "NCI_Income": GetRowForKey = 34
        Case "StockBasedCompensation": GetRowForKey = 38
        Case "DilutedShares": GetRowForKey = 41
        
        ' Balance Sheet
        Case "CashEquivalents": GetRowForKey = 63
        Case "ShortTermInvestments": GetRowForKey = 64
        Case "RestrictedCashCurrent": GetRowForKey = 65
        Case "AccountsReceivable": GetRowForKey = 66
        Case "PrepaidExpenses": GetRowForKey = 67
        Case "RestrictedCashNoncurrent": GetRowForKey = 69
        Case "RestrictedInvestmentsNoncurrent": GetRowForKey = 70
        Case "InvestmentsNoncurrent": GetRowForKey = 71
        Case "EquityMethodInvestments": GetRowForKey = 72
        Case "PP_E": GetRowForKey = 73
        Case "LeaseAssets": GetRowForKey = 74
        Case "Intangibles": GetRowForKey = 75
        Case "Goodwill": GetRowForKey = 76
        Case "OtherAssets": GetRowForKey = 77
        Case "AccountsPayable": GetRowForKey = 81
        Case "ShortTermDebt": GetRowForKey = 82
        Case "LeaseLiabilitiesCurrent": GetRowForKey = 83
        Case "AccruedLiabilitiesCurrent": GetRowForKey = 84
        Case "LongTermInsuranceReserves": GetRowForKey = 86
        Case "LongTermDebt": GetRowForKey = 87
        Case "LeaseLiabilitiesNoncurrent": GetRowForKey = 88
        Case "OtherLiabilitiesNoncurrent": GetRowForKey = 89
        Case "RedeemableNCI": GetRowForKey = 91
        Case "APIC": GetRowForKey = 94
        Case "AccumOCI": GetRowForKey = 95
        Case "AccumDeficit": GetRowForKey = 96
        Case "NoncontrollingInterest": GetRowForKey = 98
        
        ' Historical Cash Flow inputs are intentionally not populated (kept empty)
        Case Else: GetRowForKey = 0
    End Select
End Function
