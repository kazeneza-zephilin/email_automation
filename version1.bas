
Sub CreateAttendanceFile()
    Dim ExcelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim docPath As String
    Dim filePath As String

    ' Get path to Documents folder
    docPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    filePath = docPath & "\EventAttendance.xlsm"

    ' Start Excel and create workbook
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = True
    Set wb = ExcelApp.Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "AttendanceLog"

    ' Set headers
    ws.Cells(1, 1).Value = "Name"
    ws.Cells(1, 2).Value = "Email"
    ws.Cells(1, 3).Value = "Response"
    ws.Cells(1, 4).Value = "Timestamp"

    ' Save the file as macro-enabled
    wb.SaveAs filePath, 52 ' 52 = xlOpenXMLWorkbookMacroEnabled
    MsgBox "File created successfully at: " & filePath
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim arr() As String
    Dim i As Integer
    arr = Split(EntryIDCollection, ",")
    
    For i = 0 To UBound(arr)
        Dim mail As MailItem
        On Error Resume Next
        Set mail = Application.Session.GetItemFromID(arr(i))
        If Not mail Is Nothing Then
            Call ProcessMail(mail)
        End If
        On Error GoTo 0
    Next i
End Sub

Sub ProcessMail(mail As MailItem)
    On Error Resume Next
    Dim response As String
    Dim ExcelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim docPath As String
    Dim filePath As String

    ' Only process relevant emails
    If mail.Subject Like "*Event*" Then
        response = LCase(mail.Body)

        If InStr(response, "yes") > 0 Or InStr(response, "i will attend") > 0 Then
            docPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
            filePath = docPath & "\EventAttendance.xlsm"

            Set ExcelApp = GetObject(, "Excel.Application")
            Set wb = ExcelApp.Workbooks("EventAttendance.xlsm")
            Set ws = wb.Sheets("AttendanceLog")

            lastRow = ws.Cells(ws.Rows.Count, "A").End(-4162).Row + 1 ' -4162 = xlUp

            ws.Cells(lastRow, 1).Value = mail.SenderName
            ws.Cells(lastRow, 2).Value = mail.SenderEmailAddress
            ws.Cells(lastRow, 3).Value = "Yes"
            ws.Cells(lastRow, 4).Value = Now

            MsgBox "Response logged for: " & mail.SenderName
        End If
    End If
End Sub
