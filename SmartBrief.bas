' SmartBrief VBA Module for LCMC Task Manager
' Fetches email summaries from the Email Summarizer API and writes to column L
'
' Installation:
' 1. Open Excel, press Alt+F11 to open VBA Editor
' 2. Right-click on your workbook in Project Explorer
' 3. Insert > Module
' 4. Paste this code
' 5. Also paste the ThisWorkbook code (see bottom of file) into ThisWorkbook module
' 6. Save as .xlsm (macro-enabled workbook)
'
' Auto-run behavior:
' - SmartBriefAll runs automatically when workbook opens
' - SmartBriefRow runs when company name changes in column A
'
' Manual usage:
' - SmartBrief: Fetch summary for current row
' - SmartBriefAll: Fetch summaries for all rows
' - SmartBriefAllSilent: Same as SmartBriefAll but no prompts (used by auto-run)

Option Explicit

' API Configuration - Update this to your server
Private Const API_URL As String = "http://192.168.1.141:5002/api/task-manager/summaries/"
Private Const MATCH_THRESHOLD As Double = 0.5

' Main SmartBrief function - fetches summary for current row
Public Sub SmartBrief()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim currentRow As Long
    Dim companyName As String
    Dim initialRequest As String
    Dim introEmail As String
    Dim remindersStart As String
    Dim summary As String

    Set ws = ActiveSheet
    currentRow = ActiveCell.Row

    ' Skip header row
    If currentRow < 2 Then
        MsgBox "Please select a data row (row 2 or below)", vbExclamation, "SmartBrief"
        Exit Sub
    End If

    ' Get company name from column A
    companyName = Trim(CStr(ws.Cells(currentRow, "A").Value))
    If companyName = "" Then
        MsgBox "No company name found in column A for row " & currentRow, vbExclamation, "SmartBrief"
        Exit Sub
    End If

    ' Get dates from columns E, F, G
    initialRequest = FormatDateForAPI(ws.Cells(currentRow, "E").Value)
    introEmail = FormatDateForAPI(ws.Cells(currentRow, "F").Value)
    remindersStart = FormatDateForAPI(ws.Cells(currentRow, "G").Value)

    ' Show progress
    Application.StatusBar = "SmartBrief: Fetching summary for " & companyName & "..."

    ' Fetch summary from API
    summary = FetchSummaryFromAPI(companyName, initialRequest, introEmail, remindersStart)

    ' Write to column L
    If summary <> "" Then
        ws.Cells(currentRow, "L").Value = summary
        Application.StatusBar = "SmartBrief: Summary loaded for " & companyName
        MsgBox "Summary loaded successfully!", vbInformation, "SmartBrief"
    Else
        Application.StatusBar = "SmartBrief: No summary found for " & companyName
        MsgBox "No matching summaries found for " & companyName, vbInformation, "SmartBrief"
    End If

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description, vbCritical, "SmartBrief Error"
End Sub

' Fetch summary for ALL rows with company names (batch processing)
Public Sub SmartBriefAll()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim companyName As String
    Dim summary As String
    Dim processedCount As Long
    Dim foundCount As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found in the worksheet", vbExclamation, "SmartBrief"
        Exit Sub
    End If

    If MsgBox("This will fetch summaries for all " & (lastRow - 1) & " rows. Continue?", _
              vbYesNo + vbQuestion, "SmartBrief All") = vbNo Then
        Exit Sub
    End If

    processedCount = 0
    foundCount = 0

    For i = 2 To lastRow
        companyName = Trim(CStr(ws.Cells(i, "A").Value))

        If companyName <> "" Then
            Application.StatusBar = "SmartBrief: Processing row " & i & " of " & lastRow & " (" & companyName & ")"
            DoEvents

            summary = FetchSummaryFromAPI(companyName, _
                FormatDateForAPI(ws.Cells(i, "E").Value), _
                FormatDateForAPI(ws.Cells(i, "F").Value), _
                FormatDateForAPI(ws.Cells(i, "G").Value))

            If summary <> "" Then
                ws.Cells(i, "L").Value = summary
                foundCount = foundCount + 1
            End If

            processedCount = processedCount + 1
        End If
    Next i

    Application.StatusBar = False
    MsgBox "Processed " & processedCount & " rows." & vbCrLf & _
           "Found summaries for " & foundCount & " companies.", vbInformation, "SmartBrief Complete"

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error at row " & i & ": " & Err.Description, vbCritical, "SmartBrief Error"
End Sub

' Core function to call the API
Private Function FetchSummaryFromAPI(companyName As String, initialRequest As String, _
                                      introEmail As String, remindersStart As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim jsonBody As String
    Dim response As String
    Dim summaryText As String

    ' Create HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Build URL with encoded company name
    url = API_URL & URLEncode(companyName)

    ' Build JSON body
    jsonBody = "{" & _
        """initial_request"": """ & initialRequest & """," & _
        """intro_email"": """ & introEmail & """," & _
        """reminders_start"": """ & remindersStart & """," & _
        """threshold"": " & Replace(CStr(MATCH_THRESHOLD), ",", ".") & _
    "}"

    ' Make POST request
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    ' Check response
    If http.Status = 200 Then
        response = http.responseText
        summaryText = ParseBestSummary(response)
    Else
        Debug.Print "API Error: " & http.Status & " - " & http.responseText
        summaryText = ""
    End If

    FetchSummaryFromAPI = summaryText
    Exit Function

ErrorHandler:
    Debug.Print "FetchSummaryFromAPI Error: " & Err.Description
    FetchSummaryFromAPI = ""
End Function

' Parse JSON response and extract the best summary
' Returns formatted text with subject, match score, and summary
Private Function ParseBestSummary(jsonResponse As String) As String
    On Error GoTo ErrorHandler

    Dim result As String
    Dim summariesStart As Long
    Dim summaryBlock As String
    Dim subject As String
    Dim summaryText As String
    Dim matchScore As String
    Dim fromEmail As String
    Dim emailDate As String

    ' Simple JSON parsing (no external library needed)
    ' Look for the first summary in the array

    summariesStart = InStr(jsonResponse, """summaries""")
    If summariesStart = 0 Then
        ParseBestSummary = ""
        Exit Function
    End If

    ' Check if summaries array is empty
    If InStr(jsonResponse, """summaries"": []") > 0 Or InStr(jsonResponse, """summaries"":[]") > 0 Then
        ParseBestSummary = ""
        Exit Function
    End If

    ' Extract first summary block (best match - API returns sorted by score)
    subject = ExtractJSONValue(jsonResponse, "subject")
    summaryText = ExtractJSONValue(jsonResponse, "summary")
    matchScore = ExtractJSONValue(jsonResponse, "match_score")
    fromEmail = ExtractJSONValue(jsonResponse, "from")
    emailDate = ExtractJSONValue(jsonResponse, "date")

    If summaryText = "" Then
        ParseBestSummary = ""
        Exit Function
    End If

    ' Format the output for column J
    ' Short format: just the summary with match info on first line
    result = "[" & Format(Val(matchScore) * 100, "0") & "% match] " & subject & vbLf & vbLf & summaryText

    ParseBestSummary = result
    Exit Function

ErrorHandler:
    Debug.Print "ParseBestSummary Error: " & Err.Description
    ParseBestSummary = ""
End Function

' Extract a value from JSON by key (simple parser)
Private Function ExtractJSONValue(json As String, key As String) As String
    On Error GoTo ErrorHandler

    Dim keyPos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim value As String
    Dim searchKey As String

    searchKey = """" & key & """:"
    keyPos = InStr(json, searchKey)

    If keyPos = 0 Then
        ExtractJSONValue = ""
        Exit Function
    End If

    valueStart = keyPos + Len(searchKey)

    ' Skip whitespace
    Do While Mid(json, valueStart, 1) = " "
        valueStart = valueStart + 1
    Loop

    ' Check if value is a string (starts with quote) or number
    If Mid(json, valueStart, 1) = """" Then
        ' String value
        valueStart = valueStart + 1
        valueEnd = valueStart

        ' Find closing quote (handle escaped quotes)
        Do While valueEnd <= Len(json)
            If Mid(json, valueEnd, 1) = """" Then
                If Mid(json, valueEnd - 1, 1) <> "\" Then
                    Exit Do
                End If
            End If
            valueEnd = valueEnd + 1
        Loop

        value = Mid(json, valueStart, valueEnd - valueStart)
        ' Unescape common sequences
        value = Replace(value, "\""", """")
        value = Replace(value, "\\", "\")
        value = Replace(value, "\n", vbLf)
        value = Replace(value, "\r", "")
        value = Replace(value, "\t", vbTab)
    Else
        ' Number or other value
        valueEnd = valueStart
        Do While valueEnd <= Len(json)
            Dim c As String
            c = Mid(json, valueEnd, 1)
            If c = "," Or c = "}" Or c = "]" Then
                Exit Do
            End If
            valueEnd = valueEnd + 1
        Loop
        value = Trim(Mid(json, valueStart, valueEnd - valueStart))
    End If

    ExtractJSONValue = value
    Exit Function

ErrorHandler:
    ExtractJSONValue = ""
End Function

' URL encode a string
Private Function URLEncode(str As String) As String
    Dim i As Long
    Dim c As String
    Dim result As String

    result = ""
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        Select Case Asc(c)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126  ' 0-9, A-Z, a-z, -, ., _, ~
                result = result & c
            Case 32  ' Space
                result = result & "%20"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i

    URLEncode = result
End Function

' Format date for API (YYYY-MM-DD format)
Private Function FormatDateForAPI(dateValue As Variant) As String
    On Error GoTo ErrorHandler

    If IsEmpty(dateValue) Or dateValue = "" Then
        FormatDateForAPI = ""
        Exit Function
    End If

    If IsDate(dateValue) Then
        FormatDateForAPI = Format(CDate(dateValue), "yyyy-mm-dd")
    Else
        ' Try to parse as string
        FormatDateForAPI = CStr(dateValue)
    End If

    Exit Function

ErrorHandler:
    FormatDateForAPI = ""
End Function

' Quick test function to verify API connectivity
Public Sub TestAPIConnection()
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim testUrl As String

    Set http = CreateObject("MSXML2.XMLHTTP")
    testUrl = Replace(API_URL, "/api/task-manager/summaries/", "/health")

    Application.StatusBar = "Testing connection to " & testUrl & "..."

    http.Open "GET", testUrl, False
    http.send

    Application.StatusBar = False

    If http.Status = 200 Then
        MsgBox "API connection successful!" & vbCrLf & vbCrLf & _
               "Server: " & Replace(API_URL, "/api/task-manager/summaries/", "") & vbCrLf & _
               "Status: Online", vbInformation, "SmartBrief API Test"
    Else
        MsgBox "API returned status: " & http.Status & vbCrLf & _
               "Response: " & Left(http.responseText, 200), vbExclamation, "SmartBrief API Test"
    End If

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Cannot connect to API server." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure the Email Summarizer service is running on:" & vbCrLf & _
           Replace(API_URL, "/api/task-manager/summaries/", ""), vbCritical, "SmartBrief API Test"
End Sub

' Silent version of SmartBriefAll - no prompts, used for auto-run
Public Sub SmartBriefAllSilent()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim companyName As String
    Dim summary As String
    Dim foundCount As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    foundCount = 0

    For i = 2 To lastRow
        companyName = Trim(CStr(ws.Cells(i, "A").Value))

        ' Only fetch if company exists and column L is empty
        If companyName <> "" And Trim(CStr(ws.Cells(i, "L").Value)) = "" Then
            Application.StatusBar = "SmartBrief: Fetching summary for " & companyName & "..."
            DoEvents

            summary = FetchSummaryFromAPI(companyName, _
                FormatDateForAPI(ws.Cells(i, "E").Value), _
                FormatDateForAPI(ws.Cells(i, "F").Value), _
                FormatDateForAPI(ws.Cells(i, "G").Value))

            If summary <> "" Then
                ws.Cells(i, "L").Value = summary
                foundCount = foundCount + 1
            End If
        End If
    Next i

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "SmartBriefAllSilent Error: " & Err.Description
End Sub

' Fetch summary for a specific row (used by Worksheet_Change event)
Public Sub SmartBriefRow(rowNum As Long)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim companyName As String
    Dim summary As String

    Set ws = ActiveSheet

    If rowNum < 2 Then Exit Sub

    companyName = Trim(CStr(ws.Cells(rowNum, "A").Value))
    If companyName = "" Then Exit Sub

    Application.StatusBar = "SmartBrief: Fetching summary for " & companyName & "..."

    summary = FetchSummaryFromAPI(companyName, _
        FormatDateForAPI(ws.Cells(rowNum, "E").Value), _
        FormatDateForAPI(ws.Cells(rowNum, "F").Value), _
        FormatDateForAPI(ws.Cells(rowNum, "G").Value))

    If summary <> "" Then
        ws.Cells(rowNum, "L").Value = summary
    End If

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "SmartBriefRow Error: " & Err.Description
End Sub

' ============================================================================
' PASTE THE CODE BELOW INTO "ThisWorkbook" MODULE (not in a regular module)
' To access: In VBA Editor, double-click "ThisWorkbook" in Project Explorer
' ============================================================================
'
' Private Sub Workbook_Open()
'     ' Auto-fetch summaries for rows that don't have one yet
'     Application.OnTime Now + TimeValue("00:00:02"), "SmartBriefAllSilent"
' End Sub
'
' ============================================================================
' PASTE THE CODE BELOW INTO THE WORKSHEET MODULE (e.g., "Sheet1")
' To access: In VBA Editor, double-click the sheet name in Project Explorer
' ============================================================================
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     ' Auto-fetch summary when company name changes in column A
'     If Target.Column = 1 And Target.Row > 1 Then
'         If Target.Value <> "" Then
'             Application.OnTime Now + TimeValue("00:00:01"), "'SmartBriefRow " & Target.Row & "'"
'         End If
'     End If
' End Sub
