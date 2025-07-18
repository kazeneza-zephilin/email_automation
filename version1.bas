' Email Processing Automation - Manual Processing Version
' Processes emails in inbox manually when user runs the script

' Removed automatic processing - no longer needed
' Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
' End Sub

Sub ProcessAllEmails()
    ' Enhanced Email Processing Automation with User Input
    ' User workflow: Alt+F8 -> Enter event name -> Enter start date -> Process emails
    
    Dim inbox As folder
    Dim mail As mailItem
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
    
    ' Reset static counter for fresh run
    Call ResetResponseCounter
    
    ' Setup folder and file paths
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    docPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    eventFolder = docPath & "\events"
    
    ' Create events folder if it doesn't exist
    If Not fso.FolderExists(eventFolder) Then
        fso.CreateFolder (eventFolder)
    End If
    
    ' Create filename based on event name
    Dim cleanEventName As String
    cleanEventName = Replace(Replace(Replace(eventName, " ", "_"), ":", ""), "/", "")
    filePath = eventFolder & "\" & cleanEventName & "_Responses.xlsm"
    
    ' Create file if it doesn't exist OR clear existing data for same event
    If Not fso.FileExists(filePath) Then
        Call CreateEventFile(filePath, eventName)
    Else
        ' File exists - we need to clear existing data for fresh run
        MsgBox "Existing event file found. Data will be refreshed for new run.", vbInformation, "Refreshing Data"
    End If
    
    ' Step 6: Process emails
    On Error GoTo ErrorHandler
    
    ' Use the current Outlook application instance directly
    ' Since we're running from within Outlook, use the Application object
    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(6) ' 6 = olFolderInbox
    
    ' Check if inbox is accessible
    If inbox Is Nothing Then
        MsgBox "ERROR: Cannot access Outlook inbox." & vbCrLf & _
               "Please ensure Outlook is running and try again.", vbCritical, "Inbox Access Error"
        Exit Sub
    End If
    
    ' Show loading dialog
    Dim loadingForm As Object
    Set loadingForm = CreateObject("WScript.Shell")
    
    ' Show initial progress message
    MsgBox "Processing emails in progress..." & vbCrLf & vbCrLf & _
           "This may take a few moments depending on inbox size." & vbCrLf & _
           "Please wait...", vbInformation, "Processing Emails"
    
    ' Process emails received on or after the start date
    Dim emailsProcessed As Long
    emailsProcessed = 0
    
    ' Don't sort emails - process them in natural order to avoid date filtering issues
    ' inbox.Items.Sort "[ReceivedTime]", True
    
    ' Process emails with enhanced error handling
    Dim mailItem As Object
    Dim currentProgress As Long
    Dim totalEmails As Long
    Dim debugCounter As Long
    
    ' Get total count for progress tracking
    totalEmails = inbox.Items.Count
    currentProgress = 0
    debugCounter = 0
    
    ' Process emails silently - debug focus on Excel automation only
    
    For Each mailItem In inbox.Items
        On Error Resume Next ' Continue processing even if one email fails
        
        currentProgress = currentProgress + 1
        
        ' Show progress every 100 emails
        If currentProgress Mod 100 = 0 Or currentProgress = totalEmails Then
            Application.StatusBar = "Processing emails: " & currentProgress & " of " & totalEmails & " (" & Format(currentProgress / totalEmails, "0%") & ")"
            DoEvents ' Allow UI to update
        End If
        
        ' Skip if we can't access this email item or if it's invalid
        If Not mailItem Is Nothing Then
            ' Try to access TypeName to verify the object is valid
            testAccess = TypeName(mailItem)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextEmail
            End If
            
            ' Skip if not a MailItem (could be meeting requests, etc.)
            If testAccess = "MailItem" Then
                processedCount = processedCount + 1
                
                ' Get receivedTime with error handling
                Dim emailDate As Date
                emailDate = mailItem.receivedTime
                If Err.Number <> 0 Then
                    ' If we can't get received time, use current date as fallback
                    emailDate = Date
                    Err.Clear
                End If
                
                ' Check if email was received on or after start date
                If emailDate >= trackingStartDate Then
                    ' Check if email is a reply to the specific event being tracked
                    If IsReplyToEvent(mailItem, eventName) Then
                        eventCount = eventCount + 1
                        
                        ' Parse the email response
                        responseType = ParseEmailResponse(mailItem)
                        
                        ' Get email properties with simple error handling
                        Dim senderName As String
                        Dim senderEmail As String
                        Dim emailSubject As String
                        
                        senderName = "Unknown Sender"
                        senderEmail = "unknown@email.com"
                        emailSubject = "No Subject"
                        
                        ' Try to get sender name
                        On Error Resume Next
                        If Not IsEmpty(mailItem.senderName) Then
                            senderName = CStr(mailItem.senderName)
                        End If
                        
                        ' Try to get sender email
                        If Not IsEmpty(mailItem.SenderEmailAddress) Then
                            senderEmail = CStr(mailItem.SenderEmailAddress)
                        End If
                        
                        ' Try to get subject
                        If Not IsEmpty(mailItem.subject) Then
                            emailSubject = CStr(mailItem.subject)
                        End If
                        On Error GoTo NextEmail
                        
                        ' Always log the response (even with default values)
                        Call LogResponse(filePath, senderName, senderEmail, responseType, emailDate, emailSubject, eventName)
                        responseCount = responseCount + 1
                        emailsProcessed = emailsProcessed + 1
                    End If
                End If
                ' REMOVED: Early exit that was causing the problem
                ' We now continue processing all emails regardless of date order
            End If
        End If
        
NextEmail:
        ' Clear any errors and continue
        If Err.Number <> 0 Then Err.Clear
    Next mailItem
    
    ' Clear status bar
    Application.StatusBar = False
    
    ' Clean up objects
    Set mailItem = Nothing
    Set inbox = Nothing
    
    ' Step 8: Success confirmation
    Dim debugInfo As String
    debugInfo = ""
    
    ' Add debug information if no responses were found
    If responseCount = 0 And eventCount = 0 Then
        debugInfo = vbCrLf & vbCrLf & "DEBUG INFO:" & vbCrLf & _
                   "No matching emails found. Please check:" & vbCrLf & _
                   "• Event name entered: '" & eventName & "'" & vbCrLf & _
                   "• Email subjects should contain this event name" & vbCrLf & _
                   "• Emails should be received after: " & Format(trackingStartDate, "mm/dd/yyyy") & vbCrLf & _
                   "• Check if emails are actually in your inbox"
    ElseIf eventCount > 0 And responseCount = 0 Then
        debugInfo = vbCrLf & vbCrLf & "FOUND EMAILS BUT NO RESPONSES:" & vbCrLf & _
                   "Found " & eventCount & " event emails but couldn't parse responses." & vbCrLf & _
                   "This might be due to response text format."
    End If
    
    MsgBox "Email processing completed successfully!" & vbCrLf & vbCrLf & _
           "Event: " & eventName & vbCrLf & _
           "Tracking Period: " & Format(trackingStartDate, "mmmm dd, yyyy") & " to " & Format(Date, "mmmm dd, yyyy") & vbCrLf & vbCrLf & _
           "RESULTS:" & vbCrLf & _
           "Total emails scanned: " & processedCount & vbCrLf & _
           "Event-related emails found: " & eventCount & vbCrLf & _
           "Responses recorded: " & responseCount & debugInfo & vbCrLf & vbCrLf & _
           "Results saved to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
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
            errorMsg = "OUTLOOK BUSY ERROR:" & vbCrLf & vbCrLf & _
                      "Outlook is currently busy or locked by another process." & vbCrLf & vbCrLf & _
                      "SOLUTIONS:" & vbCrLf & _
                      "1. Close any open email windows in Outlook" & vbCrLf & _
                      "2. Wait 10 seconds and try again" & vbCrLf & _
                      "3. Restart Outlook completely" & vbCrLf & _
                      "4. Make sure no other email programs are running" & vbCrLf & _
                      "5. Check if Outlook is syncing emails (wait for it to finish)"
        Case -2147221236, -2147221238, -2147221240 ' Common Outlook automation errors
            errorMsg = "OUTLOOK CONNECTION ERROR:" & vbCrLf & vbCrLf & _
                      "Outlook automation was rejected. This can happen when:" & vbCrLf & _
                      "• Outlook is not fully loaded" & vbCrLf & _
                      "• Another process is using Outlook" & vbCrLf & _
                      "• Outlook security settings block automation" & vbCrLf & vbCrLf & _
                      "SOLUTIONS:" & vbCrLf & _
                      "1. Close and restart Outlook completely" & vbCrLf & _
                      "2. Wait 30 seconds after opening Outlook" & vbCrLf & _
                      "3. Run this script as Administrator" & vbCrLf & _
                      "4. Disable antivirus email scanning temporarily"
        Case 438 ' Object doesn't support this property or method
            errorMsg = "EXCEL AUTOMATION ERROR:" & vbCrLf & vbCrLf & _
                      "Excel automation failed. This can happen when:" & vbCrLf & _
                      "• Excel is not installed or not properly registered" & vbCrLf & _
                      "• Excel is busy or locked by another process" & vbCrLf & _
                      "• File permissions issue" & vbCrLf & vbCrLf & _
                      "SOLUTIONS:" & vbCrLf & _
                      "1. Close all Excel windows and try again" & vbCrLf & _
                      "2. Restart Outlook and Excel" & vbCrLf & _
                      "3. Run as Administrator" & vbCrLf & _
                      "4. Check if Excel is properly installed" & vbCrLf & _
                      "5. Try running Excel manually first"
        Case 70 ' Permission denied
            errorMsg = "PERMISSION ERROR:" & vbCrLf & vbCrLf & _
                      "Access denied. Please:" & vbCrLf & _
                      "1. Run Outlook as Administrator" & vbCrLf & _
                      "2. Check file permissions in Documents folder" & vbCrLf & _
                      "3. Ensure Excel is not already open with the file"
        Case Else
            errorMsg = "PROCESSING ERROR:" & vbCrLf & vbCrLf & _
                      "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
                      "Please try again. If problem persists:" & vbCrLf & _
                      "1. Restart Outlook and Excel" & vbCrLf & _
                      "2. Run as Administrator" & vbCrLf & _
                      "3. Contact support with error number"
    End Select
    
    MsgBox errorMsg, vbCritical, "Processing Error"
End Sub

Sub ResetResponseCounter()
    ' Reset the static counter in LogResponse for new automation run
    Static dummy As Long
    Call LogResponse("RESET", "", "", "", Date, "", "")
End Sub

Function IsReplyToEvent(mail As mailItem, eventName As String) As Boolean
    On Error Resume Next
    
    ' Validate input parameters
    If mail Is Nothing Then
        IsReplyToEvent = False
        Exit Function
    End If
    
    If eventName = "" Then
        IsReplyToEvent = False
        Exit Function
    End If
    
    Dim subject As String
    Dim body As String
    Dim cleanEventName As String
    
    ' Safely get subject and body
    subject = ""
    body = ""
    
    If Not IsEmpty(mail.subject) And Not IsNull(mail.subject) Then
        subject = LCase(Trim(CStr(mail.subject)))
    End If
    
    If Not IsEmpty(mail.body) And Not IsNull(mail.body) Then
        body = LCase(Trim(CStr(mail.body)))
    End If
    
    eventName = LCase(eventName)
    
    ' Extract core event name without "Event:" prefix for more flexible matching
    cleanEventName = eventName
    If Left(cleanEventName, 6) = "event:" Then
        cleanEventName = Trim(Mid(cleanEventName, 7))
    End If
    
    ' Test 1: Check if subject contains the event name or clean event name
    If InStr(subject, eventName) > 0 Or InStr(subject, cleanEventName) > 0 Then
        IsReplyToEvent = True
        Exit Function
    End If
    
    ' Test 2: Check for "Re:" prefix with event-related patterns
    If InStr(subject, "re:") > 0 Then
        If InStr(subject, "event:") > 0 Or _
           InStr(subject, "invitation:") > 0 Or _
           InStr(subject, "rsvp") > 0 Or _
           InStr(subject, "reception") > 0 Or _
           InStr(subject, "diplomatic") > 0 Then
            IsReplyToEvent = True
            Exit Function
        End If
    End If
    
    ' Test 3: Check body content
    If (InStr(body, eventName) > 0 Or InStr(body, cleanEventName) > 0) And _
       (InStr(body, "rsvp") > 0 Or _
        InStr(body, "attend") > 0 Or _
        InStr(body, "invitation") > 0 Or _
        InStr(body, "reception") > 0 Or _
        InStr(body, "confirm") > 0) Then
        IsReplyToEvent = True
        Exit Function
    End If
    
    ' Test 4: Catch-all for reception/diplomatic
    If InStr(subject, "re:") > 0 And _
       (InStr(subject, "reception") > 0 Or InStr(subject, "diplomatic") > 0) Then
        IsReplyToEvent = True
        Exit Function
    End If
    
    ' Default to false if no match found
    IsReplyToEvent = False
End Function

Function ParseEmailResponse(mail As mailItem) As String
    On Error Resume Next
    
    ' Validate input
    If mail Is Nothing Then
        ParseEmailResponse = "Unknown"
        Exit Function
    End If
    
    Dim emailText As String
    Dim responseType As String
    Dim subject As String
    Dim body As String
    
    ' Safely get subject and body
    subject = ""
    body = ""
    
    If Not IsEmpty(mail.subject) And Not IsNull(mail.subject) Then
        subject = CStr(mail.subject)
    End If
    
    If Not IsEmpty(mail.body) And Not IsNull(mail.body) Then
        body = CStr(mail.body)
    End If
    
    ' Combine subject and body for analysis
    emailText = LCase(subject & " " & body)
    
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
    
    On Error GoTo CreateError
    
    ' Create Excel application
    Set ExcelApp = CreateObject("Excel.Application")
    If ExcelApp Is Nothing Then
        MsgBox "ERROR: Cannot create Excel application. Please ensure Excel is installed.", vbCritical, "Excel Error"
        Exit Sub
    End If
    
    ' Silent file creation
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    ExcelApp.ScreenUpdating = False
    
    ' Create new workbook
    Set wb = ExcelApp.Workbooks.Add
    Set ws = wb.Sheets(1)
    
    ' Rename sheet
    On Error Resume Next
    ws.Name = "EventResponses"
    On Error GoTo CreateError
    
    ' Set headers with event information
    ws.Cells(1, 1).Value = "Event: " & eventName
    ws.Cells(2, 1).Value = "Generated: " & Format(Now, "mmmm dd, yyyy hh:mm AM/PM")
    
    ' Apply basic formatting
    On Error Resume Next
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Interior.Color = RGB(220, 230, 255)
    ws.Cells(2, 1).Font.Italic = True
    On Error GoTo CreateError
    
    ' Set column headers
    ws.Cells(4, 1).Value = "Name"
    ws.Cells(4, 2).Value = "Email"
    ws.Cells(4, 3).Value = "Response"
    ws.Cells(4, 4).Value = "Date Received"
    ws.Cells(4, 5).Value = "Subject"
    ws.Cells(4, 6).Value = "Event Name"
    ws.Cells(4, 7).Value = "Processing Notes"
    
    ' Format headers
    On Error Resume Next
    ws.Range("A4:G4").Font.Bold = True
    ws.Range("A4:G4").Interior.Color = RGB(200, 220, 255)
    ws.Range("A1:G1").Merge
    ws.Columns("A:G").AutoFit
    On Error GoTo CreateError
    
    ' Delete existing file if present
    On Error Resume Next
    If Dir(filePath) <> "" Then Kill filePath
    On Error GoTo CreateError
    
    ' Save file
    wb.SaveAs FileName:=filePath, FileFormat:=52, ReadOnlyRecommended:=False
    wb.Close SaveChanges:=True
    
    ' Clean up
    Set ws = Nothing
    Set wb = Nothing
    ExcelApp.Quit
    Set ExcelApp = Nothing
    Exit Sub

CreateError:
    ' Error handling for file creation
    On Error Resume Next
    If Not ExcelApp Is Nothing Then
        ExcelApp.DisplayAlerts = True
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        ExcelApp.Quit
        Set ExcelApp = Nothing
    End If
    MsgBox "ERROR creating Excel file. Please ensure you have Excel installed and try again.", vbCritical, "File Creation Error"
End Sub

Sub LogResponse(filePath As String, senderName As String, senderEmail As String, responseType As String, receivedTime As Date, emailSubject As String, eventName As String)
    Dim ExcelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim lastRow As Long
    Static responseCounter As Long
    
    On Error GoTo LogError
    
    ' Handle reset call
    If filePath = "RESET" Then
        responseCounter = 0
        Exit Sub
    End If
    
    ' Validate input parameters
    If filePath = "" Or eventName = "" Then
        Exit Sub ' Skip if essential parameters are missing
    End If
    
    ' Ensure we have minimum required data
    If senderName = "" Then senderName = "Unknown Sender"
    If senderEmail = "" Then senderEmail = "unknown@email.com"
    If emailSubject = "" Then emailSubject = "No Subject"
    If responseType = "" Then responseType = "Unknown"
    
    ' Create Excel application
    Set ExcelApp = CreateObject("Excel.Application")
    If ExcelApp Is Nothing Then
        Exit Sub ' Skip silently if Excel unavailable
    End If
    
    ' Silent processing
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    ExcelApp.ScreenUpdating = False
    
    ' Open or create file
    If Dir(filePath) = "" Then
        Call CreateEventFile(filePath, eventName)
        responseCounter = 0  ' Reset counter for new file
    End If
    
    ' Open workbook
    Set wb = ExcelApp.Workbooks.Open(filePath, ReadOnly:=False, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    ' For first response of this run, always clear existing data
    If responseCounter = 0 Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        If lastRow > 4 Then
            ws.Range("A5:G" & lastRow).Clear
        End If
        lastRow = 4
    Else
        ' Find the actual last row for subsequent responses
        lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        If lastRow < 4 Then lastRow = 4
    End If
    
    responseCounter = responseCounter + 1
    lastRow = lastRow + 1
    
    ' Write data
    ws.Cells(lastRow, 1).Value = senderName
    ws.Cells(lastRow, 2).Value = senderEmail
    ws.Cells(lastRow, 3).Value = responseType
    ws.Cells(lastRow, 4).Value = Format(receivedTime, "mm/dd/yyyy hh:mm AM/PM")
    ws.Cells(lastRow, 5).Value = emailSubject
    ws.Cells(lastRow, 6).Value = eventName
    ws.Cells(lastRow, 7).Value = "Processed " & Format(Now, "mm/dd/yyyy hh:mm AM/PM")
    
    ' Apply formatting
    On Error Resume Next
    Select Case UCase(responseType)
        Case "YES": ws.Cells(lastRow, 3).Interior.Color = RGB(200, 255, 200)
        Case "NO": ws.Cells(lastRow, 3).Interior.Color = RGB(255, 200, 200)
        Case "MAYBE": ws.Cells(lastRow, 3).Interior.Color = RGB(255, 255, 200)
        Case "UNKNOWN": ws.Cells(lastRow, 3).Interior.Color = RGB(220, 220, 220)
    End Select
    On Error GoTo LogError
    
    ' Save and close
    wb.Save
    wb.Close SaveChanges:=False
    
    ' Clean up
    Set ws = Nothing
    Set wb = Nothing
    ExcelApp.Quit
    Set ExcelApp = Nothing
    
    Exit Sub

LogError:
    ' Simple error handling with cleanup
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    If Not ws Is Nothing Then Set ws = Nothing
    If Not ExcelApp Is Nothing Then
        ExcelApp.Quit
        Set ExcelApp = Nothing
    End If
    ' Continue processing next email silently
    On Error GoTo 0
End Sub


