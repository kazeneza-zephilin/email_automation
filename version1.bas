' Email Processing Automation - Manual Processing Version
' Processes emails in inbox manually when user runs the script

' Removed automatic processing - no longer needed
' Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
' End Sub

Sub ProcessAllEmails()
    ' Enhanced Email Processing Automation with User Input
    ' User workflow: Alt+F8 -> Enter event name -> Enter start date -> Process emails
    
    Dim inbox As Folder
    Dim mail As MailItem
    Dim processedCount As Long
    Dim eventCount As Long
    Dim responseCount As Long
    Dim docPath As String
    Dim eventFolder As String
    Dim filePath As String
    Dim fso As Object
    Dim responseType As String
    Dim eventName As String
    Dim startDate As String
    Dim trackingStartDate As Date
    Dim userResponse As Integer
    
    ' Add delay to ensure Outlook is ready using Sleep instead of Application.Wait
    Dim startTime As Date
    startTime = Now
    Do While DateDiff("s", startTime, Now) < 2
        DoEvents ' Allow other processes to run
    Loop
    
    ' Step 3: Get event name from user
    eventName = InputBox("Enter the name of the event to track responses for:", "Event Tracking Setup", "Annual Diplomatic Reception")
    If eventName = "" Then
        MsgBox "Event tracking cancelled. No event name provided.", vbExclamation, "Cancelled"
        Exit Sub
    End If
    
    ' Step 4: Get start date from user
    startDate = InputBox("Enter the start date for tracking responses (MM/DD/YYYY):" & vbCrLf & "Only emails received on or after this date will be processed.", "Start Date", Format(Date - 7, "mm/dd/yyyy"))
    If startDate = "" Then
        MsgBox "Event tracking cancelled. No start date provided.", vbExclamation, "Cancelled"
        Exit Sub
    End If
    
    ' Validate and convert start date
    On Error GoTo InvalidDate
    trackingStartDate = CDate(startDate)
    On Error GoTo ErrorHandler
    
    ' Additional date validation
    If trackingStartDate > Date Then
        MsgBox "Start date cannot be in the future. Please enter a valid date.", vbExclamation, "Invalid Date"
        Exit Sub
    End If
    
    ' Step 5: Confirmation popup
    userResponse = MsgBox("Ready to process email responses for:" & vbCrLf & vbCrLf & _
                         "Event: " & eventName & vbCrLf & _
                         "Start Date: " & Format(trackingStartDate, "mmmm dd, yyyy") & vbCrLf & vbCrLf & _
                         "This will scan your inbox for responses received on or after the start date." & vbCrLf & vbCrLf & _
                         "Do you want to proceed?", vbYesNo + vbQuestion, "Confirm Processing")
    
    If userResponse = vbNo Then
        MsgBox "Email processing cancelled by user.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' Initialize counters
    processedCount = 0
    eventCount = 0
    responseCount = 0
    
    ' Setup folder and file paths
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    docPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    eventFolder = docPath & "\events"
    
    ' Create events folder if it doesn't exist
    If Not fso.FolderExists(eventFolder) Then
        fso.CreateFolder(eventFolder)
    End If
    
    ' Create filename based on event name
    Dim cleanEventName As String
    cleanEventName = Replace(Replace(Replace(eventName, " ", "_"), ":", ""), "/", "")
    filePath = eventFolder & "\" & cleanEventName & "_Responses.xlsm"
    
    ' Create file if it doesn't exist
    If Not fso.FileExists(filePath) Then
        Call CreateEventFile(filePath, eventName)
    End If
    
    ' Step 6: Process emails
    On Error GoTo ErrorHandler
    
    ' Use the current Outlook application instance directly
    ' Since we're running from within Outlook, use the Application object
    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(6) ' 6 = olFolderInbox
    
    ' Check if inbox is accessible
    If inbox Is Nothing Then
        MsgBox "❌ ERROR: Cannot access Outlook inbox." & vbCrLf & _
               "Please ensure Outlook is running and try again.", vbCritical, "Inbox Access Error"
        Exit Sub
    End If
    
    ' Process emails received on or after the start date
    Dim emailsProcessed As Long
    emailsProcessed = 0
    
    ' Sort emails by received time to process newest first
    inbox.Items.Sort "[ReceivedTime]", True
    
    ' Process emails with enhanced error handling
    Dim mailItem As Object
    For Each mailItem In inbox.Items
        On Error Resume Next ' Continue processing even if one email fails
        
        ' Skip if we can't access this email item
        If Not mailItem Is Nothing Then
            ' Skip if not a MailItem (could be meeting requests, etc.)
            If TypeName(mailItem) = "MailItem" Then
                processedCount = processedCount + 1
                
                ' Check if email was received on or after start date
                If mailItem.ReceivedTime >= trackingStartDate Then
                    ' Check if email is a reply to the specific event being tracked
                    If IsReplyToEvent(mailItem, eventName) Then
                        eventCount = eventCount + 1
                        
                        ' Parse the email response
                        responseType = ParseEmailResponse(mailItem)
                        
                        ' Log all responses (including Unknown for manual review)
                        Call LogResponse(filePath, mailItem.SenderName, mailItem.SenderEmailAddress, responseType, mailItem.ReceivedTime, mailItem.Subject, eventName)
                        responseCount = responseCount + 1
                        emailsProcessed = emailsProcessed + 1
                    End If
                Else
                    ' Since emails are sorted by date, we can break early if we've gone past our start date
                    Exit For
                End If
            End If
        End If
        
        On Error GoTo ErrorHandler ' Reset error handling
    Next mailItem
    
    ' Clean up objects
    Set mailItem = Nothing
    Set inbox = Nothing
    
    ' Step 8: Success confirmation
    MsgBox "Email processing completed successfully!" & vbCrLf & vbCrLf & _
           "Event: " & eventName & vbCrLf & _
           "Tracking Period: " & Format(trackingStartDate, "mmmm dd, yyyy") & " to " & Format(Date, "mmmm dd, yyyy") & vbCrLf & vbCrLf & _
           "📊 RESULTS:" & vbCrLf & _
           "Total emails scanned: " & processedCount & vbCrLf & _
           "Event-related emails found: " & eventCount & vbCrLf & _
           "Responses recorded: " & responseCount & vbCrLf & vbCrLf & _
           "✅ Results saved to: " & filePath & vbCrLf & vbCrLf & _
           "Opening Excel file now...", vbInformation, "Processing Complete"
    
    ' Open the Excel file
    Shell "explorer.exe """ & filePath & """", vbNormalFocus
    Exit Sub

InvalidDate:
    MsgBox "Invalid date format. Please use MM/DD/YYYY format (e.g., 07/15/2025).", vbCritical, "Invalid Date"
    Exit Sub

ErrorHandler:
    ' Enhanced error handling with specific error codes
    Dim errorMsg As String
    
    Select Case Err.Number
        Case -2147418111 ' Call was rejected by callee
            errorMsg = "❌ OUTLOOK BUSY ERROR:" & vbCrLf & vbCrLf & _
                      "Outlook is currently busy or locked by another process." & vbCrLf & vbCrLf & _
                      "SOLUTIONS:" & vbCrLf & _
                      "1. Close any open email windows in Outlook" & vbCrLf & _
                      "2. Wait 10 seconds and try again" & vbCrLf & _
                      "3. Restart Outlook completely" & vbCrLf & _
                      "4. Make sure no other email programs are running" & vbCrLf & _
                      "5. Check if Outlook is syncing emails (wait for it to finish)"
        Case -2147221236, -2147221238, -2147221240 ' Common Outlook automation errors
            errorMsg = "❌ OUTLOOK CONNECTION ERROR:" & vbCrLf & vbCrLf & _
                      "Outlook automation was rejected. This can happen when:" & vbCrLf & _
                      "• Outlook is not fully loaded" & vbCrLf & _
                      "• Another process is using Outlook" & vbCrLf & _
                      "• Outlook security settings block automation" & vbCrLf & vbCrLf & _
                      "SOLUTIONS:" & vbCrLf & _
                      "1. Close and restart Outlook completely" & vbCrLf & _
                      "2. Wait 30 seconds after opening Outlook" & vbCrLf & _
                      "3. Run this script as Administrator" & vbCrLf & _
                      "4. Disable antivirus email scanning temporarily"
        Case 70 ' Permission denied
            errorMsg = "❌ PERMISSION ERROR:" & vbCrLf & vbCrLf & _
                      "Access denied. Please:" & vbCrLf & _
                      "1. Run Outlook as Administrator" & vbCrLf & _
                      "2. Check file permissions in Documents folder" & vbCrLf & _
                      "3. Ensure Excel is not already open with the file"
        Case Else
            errorMsg = "❌ PROCESSING ERROR:" & vbCrLf & vbCrLf & _
                      "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
                      "Please try again. If problem persists:" & vbCrLf & _
                      "1. Restart Outlook and Excel" & vbCrLf & _
                      "2. Run as Administrator" & vbCrLf & _
                      "3. Contact support with error number"
    End Select
    
    MsgBox errorMsg, vbCritical, "Processing Error"
End Sub

Function IsReplyToEvent(mail As MailItem, eventName As String) As Boolean
    Dim subject As String
    Dim body As String
    
    subject = LCase(mail.Subject)
    body = LCase(mail.Body)
    eventName = LCase(eventName)
    
    ' Check if subject starts with "Re:" and contains the event name
    If InStr(subject, "re:") > 0 Then
        ' Look for various patterns that indicate this is a reply to the event invitation
        If InStr(subject, "event: " & eventName) > 0 Or _
           InStr(subject, "invitation: " & eventName) > 0 Or _
           InStr(subject, eventName & " - rsvp") > 0 Or _
           InStr(subject, eventName & " rsvp") > 0 Or _
           InStr(subject, eventName) > 0 Then
            IsReplyToEvent = True
            Exit Function
        End If
    End If
    
    ' Also check if the body contains the event name and RSVP-related keywords
    ' This catches forwarded responses or replies without proper "Re:" format
    If InStr(body, eventName) > 0 And _
       (InStr(body, "rsvp") > 0 Or _
        InStr(body, "attend") > 0 Or _
        InStr(body, "invitation") > 0 Or _
        InStr(body, "reception") > 0 Or _
        InStr(body, "confirm") > 0) Then
        IsReplyToEvent = True
        Exit Function
    End If
    
    ' Default to false if no match found
    IsReplyToEvent = False
End Function

Function ParseEmailResponse(mail As MailItem) As String
    Dim emailText As String
    Dim responseType As String
    
    ' Combine subject and body for analysis
    emailText = LCase(mail.Subject & " " & mail.Body)
    
    ' Remove common email artifacts
    emailText = Replace(emailText, vbCrLf, " ")
    emailText = Replace(emailText, vbLf, " ")
    emailText = Replace(emailText, vbTab, " ")
    
    ' Strong YES indicators (prioritized)
    If InStr(emailText, "yes, i will attend") > 0 Or _
       InStr(emailText, "yes i will attend") > 0 Or _
       InStr(emailText, "yes, i'll attend") > 0 Or _
       InStr(emailText, "yes i'll attend") > 0 Or _
       InStr(emailText, "yes, i will be there") > 0 Or _
       InStr(emailText, "yes i will be there") > 0 Or _
       InStr(emailText, "yes, count me in") > 0 Or _
       InStr(emailText, "yes count me in") > 0 Or _
       InStr(emailText, "i confirm my attendance") > 0 Or _
       InStr(emailText, "i will definitely attend") > 0 Or _
       InStr(emailText, "absolutely yes") > 0 Or _
       InStr(emailText, "definitely yes") > 0 Then
        responseType = "Yes"
    
    ' Strong NO indicators (prioritized)
    ElseIf InStr(emailText, "no, i cannot attend") > 0 Or _
           InStr(emailText, "no i cannot attend") > 0 Or _
           InStr(emailText, "no, i can't attend") > 0 Or _
           InStr(emailText, "no i can't attend") > 0 Or _
           InStr(emailText, "no, i will not attend") > 0 Or _
           InStr(emailText, "no i will not attend") > 0 Or _
           InStr(emailText, "no, i won't attend") > 0 Or _
           InStr(emailText, "no i won't attend") > 0 Or _
           InStr(emailText, "unfortunately, i cannot") > 0 Or _
           InStr(emailText, "unfortunately i cannot") > 0 Or _
           InStr(emailText, "unable to attend") > 0 Or _
           InStr(emailText, "will not be able to attend") > 0 Or _
           InStr(emailText, "won't be able to attend") > 0 Then
        responseType = "No"
    
    ' MAYBE/TENTATIVE indicators
    ElseIf InStr(emailText, "maybe") > 0 Or _
           InStr(emailText, "might attend") > 0 Or _
           InStr(emailText, "tentative") > 0 Or _
           InStr(emailText, "not sure") > 0 Or _
           InStr(emailText, "let you know") > 0 Or _
           InStr(emailText, "will confirm later") > 0 Or _
           InStr(emailText, "possibly") > 0 Or _
           InStr(emailText, "depends on") > 0 Or _
           InStr(emailText, "will try to") > 0 Then
        responseType = "Maybe"
    
    ' Simple YES patterns (less specific)
    ElseIf (InStr(emailText, " yes ") > 0 And InStr(emailText, "attend") > 0) Or _
           (InStr(emailText, "yes,") > 0 And InStr(emailText, "attend") > 0) Or _
           (InStr(emailText, "yes.") > 0 And InStr(emailText, "attend") > 0) Or _
           InStr(emailText, "i will attend") > 0 Or _
           InStr(emailText, "i'll attend") > 0 Or _
           InStr(emailText, "i will be there") > 0 Or _
           InStr(emailText, "i'll be there") > 0 Or _
           InStr(emailText, "count me in") > 0 Or _
           InStr(emailText, "see you there") > 0 Then
        responseType = "Yes"
    
    ' Simple NO patterns (less specific)
    ElseIf (InStr(emailText, " no ") > 0 And InStr(emailText, "attend") > 0) Or _
           (InStr(emailText, "no,") > 0 And InStr(emailText, "attend") > 0) Or _
           (InStr(emailText, "no.") > 0 And InStr(emailText, "attend") > 0) Or _
           InStr(emailText, "cannot attend") > 0 Or _
           InStr(emailText, "can't attend") > 0 Or _
           InStr(emailText, "will not attend") > 0 Or _
           InStr(emailText, "won't attend") > 0 Or _
           InStr(emailText, "sorry, no") > 0 Or _
           InStr(emailText, "sorry no") > 0 Then
        responseType = "No"
    
    Else
        responseType = "Unknown"
    End If
    
    ParseEmailResponse = responseType
End Function

Sub CreateEventFile(filePath As String, eventName As String)
    Dim ExcelApp As Object
    Dim wb As Object
    Dim ws As Object
    
    ' Enhanced error handling for file creation
    On Error GoTo CreateError
    
    ' Silent file creation
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    ExcelApp.ScreenUpdating = False
    
    ' Force Excel to only save to the specified location
    ExcelApp.DefaultFilePath = ""
    
    Set wb = ExcelApp.Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "EventResponses"
    
    ' Set headers with event information
    ws.Cells(1, 1).Value = "Event: " & eventName
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Interior.Color = RGB(220, 230, 255)
    
    ws.Cells(2, 1).Value = "Generated: " & Format(Now, "mmmm dd, yyyy hh:mm AM/PM")
    ws.Cells(2, 1).Font.Italic = True
    
    ' Set column headers
    ws.Cells(4, 1).Value = "Name"
    ws.Cells(4, 2).Value = "Email"
    ws.Cells(4, 3).Value = "Response"
    ws.Cells(4, 4).Value = "Date Received"
    ws.Cells(4, 5).Value = "Subject"
    ws.Cells(4, 6).Value = "Event Name"
    ws.Cells(4, 7).Value = "Processing Notes"
    
    ' Format headers
    ws.Range("A4:G4").Font.Bold = True
    ws.Range("A4:G4").Interior.Color = RGB(200, 220, 255)
    ws.Range("A1:G1").Merge
    ws.Columns("A:G").AutoFit
    
    ' Save to specific path only and ensure no read-only
    On Error Resume Next
    If Dir(filePath) <> "" Then Kill filePath ' Delete existing file if any
    On Error GoTo CreateError
    
    wb.SaveAs Filename:=filePath, FileFormat:=52, ReadOnlyRecommended:=False
    wb.Close SaveChanges:=True
    
    ' Clean up
    Set ws = Nothing
    Set wb = Nothing
    ExcelApp.Quit
    Set ExcelApp = Nothing
    Exit Sub

CreateError:
    ' Error handling for file creation
    If Not ExcelApp Is Nothing Then
        ExcelApp.DisplayAlerts = True
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        ExcelApp.Quit
        Set ExcelApp = Nothing
    End If
    MsgBox "❌ ERROR creating Excel file:" & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please ensure you have Excel installed and try again.", vbCritical, "File Creation Error"
End Sub

Sub LogResponse(filePath As String, senderName As String, senderEmail As String, responseType As String, receivedTime As Date, emailSubject As String, eventName As String)
    Dim ExcelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim lastRow As Long
    
    ' Enhanced error handling for Excel operations
    On Error GoTo LogError
    
    ' Silent processing - suppress all alerts and popups
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    ExcelApp.ScreenUpdating = False
    
    ' Check if file exists before trying to open
    If Dir(filePath) = "" Then
        ' File doesn't exist, create it first
        Call CreateEventFile(filePath, eventName)
    End If
    
    ' Open the existing file
    Set wb = ExcelApp.Workbooks.Open(filePath, ReadOnly:=False)
    Set ws = wb.Sheets("EventResponses")
    
    ' Find the last row with data (starting from row 4 since headers are in row 4)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
    If lastRow < 4 Then
        ' If no data exists yet, start at row 5 (first data row)
        lastRow = 5
    Else
        ' Add to next available row
        lastRow = lastRow + 1
    End If
    
    ' Add new row for every response (no duplicate checking)
    ws.Cells(lastRow, 1).Value = senderName
    ws.Cells(lastRow, 2).Value = senderEmail
    ws.Cells(lastRow, 3).Value = responseType
    ws.Cells(lastRow, 4).Value = Format(receivedTime, "mm/dd/yyyy hh:mm AM/PM")
    ws.Cells(lastRow, 5).Value = emailSubject
    ws.Cells(lastRow, 6).Value = eventName
    ws.Cells(lastRow, 7).Value = "Processed " & Format(Now, "mm/dd/yyyy hh:mm AM/PM")
    
    ' Color-code responses for easy viewing
    Select Case UCase(responseType)
        Case "YES"
            ws.Cells(lastRow, 3).Interior.Color = RGB(200, 255, 200) ' Light green
        Case "NO"
            ws.Cells(lastRow, 3).Interior.Color = RGB(255, 200, 200) ' Light red
        Case "MAYBE"
            ws.Cells(lastRow, 3).Interior.Color = RGB(255, 255, 200) ' Light yellow
        Case "UNKNOWN"
            ws.Cells(lastRow, 3).Interior.Color = RGB(220, 220, 220) ' Light gray
    End Select
    
    ' Auto-fit columns and save
    ws.Columns("A:G").AutoFit
    wb.Save
    wb.Close SaveChanges:=True
    
    ' Clean up objects
    Set ws = Nothing
    Set wb = Nothing
    ExcelApp.Quit
    Set ExcelApp = Nothing
    Exit Sub

LogError:
    ' Error handling for Excel operations
    If Not ExcelApp Is Nothing Then
        ExcelApp.DisplayAlerts = True
        ExcelApp.Quit
        Set ExcelApp = Nothing
    End If
    ' Don't show error message here as it will interrupt the main process
    ' The main function will handle overall error reporting
End Sub
