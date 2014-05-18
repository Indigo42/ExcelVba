Attribute VB_Name = "Utilities"
' ****************************************************************************
' ****************************************************************************
'
'   Description:            Utility procedures.
'
' ****************************************************************************
' ****************************************************************************
Option Explicit
Option Private Module

' ****************************************************************************
'
'   Procedure name:         AddRecipientToMailItem()
'
'   Description:            This procedure simplifies the process of creating
'                           a new recipient and setting its type.
'
'   Example call:           Here is how you would add "Joe Smith" as a "CC"
'                           recipient to a MailItem object named "eml":
'
'                               AddRecipientToMailItem eml, olCC, "Joe Smith"
'
'
' ****************************************************************************
Private Sub AddRecipientToMailItem(ByRef mailMsg As Outlook.MailItem, _
                                   ByVal recipientType _
                                       As Outlook.OlMailRecipientType, _
                                   ByVal recipientName As String)
    
    On Error GoTo ErrorHandler
    
    Dim errDesc As String                   ' Error description.
    Dim errNum As Long                      ' Error number.
    Dim errSrc As String                    ' Error source.
    Dim recip As Outlook.Recipient
    
    Set recip = mailMsg.Recipients.Add(recipientName)
    recip.Type = recipientType
    
    Set recip = Nothing
    
    Exit Sub

ErrorHandler:
    ' Clean up recip.
    If Not recip Is Nothing Then
        Set recip = Nothing
    End If
    
    ' Save information about the error.
    errSrc = Err.Source
    errNum = Err.Number
    errDesc = Err.Description
    
    ' Regenerate the error.
    Err.Clear
    Err.Raise errNum, errSrc, errDesc
End Sub

' ****************************************************************************
'
'   Procedure name:         CreateTestEmail()
'
'   Description:            Creates and displays a test email.
'
' ****************************************************************************
Public Sub CreateTestEmail(Optional ByVal attachmentPath As String = "")
    On Error GoTo ErrorHandler

    Const fileNotFound As Long = 53         ' Error code for "File not found."
    Const outlookActionBlocked _
        As Long = 287                       ' Error code when an Outlook
                                            ' action is blocked by certain
                                            ' security features.
    
    Dim errDesc As String                   ' Error description.
    Dim errNum As Long                      ' Error number.
    Dim errSrc As String                    ' Error source.
    Dim fso As Scripting.FileSystemObject
    Dim eml As Outlook.MailItem
    Dim ol As Outlook.Application
    
    ' Connect to Outlook and create the email message.
    Set ol = New Outlook.Application
    Set eml = ol.CreateItem(olMailItem)
    
    ' Set the email's "To" recipients.
    AddRecipientToMailItem eml, olTo, "jim@server.fake"
    AddRecipientToMailItem eml, olTo, "pam@server.fake"
    AddRecipientToMailItem eml, olTo, "roy@server.fake"
    
    ' Set the email's "CC" recipients.
    AddRecipientToMailItem eml, olCC, "angela@server.fake"
    AddRecipientToMailItem eml, olCC, "dwight@server.fake"
    AddRecipientToMailItem eml, olCC, "michael@server.fake"
    
    ' Set the email's subject and body.
    eml.Subject = "Test"
    eml.Body = _
        "Hello," & vbNewLine & _
        vbNewLine & _
        "This is a test." & vbNewLine & _
        vbNewLine
    
    ' Add the attachment to the email, if applicable.
    If attachmentPath <> "" Then
        Set fso = New Scripting.FileSystemObject
        
        If fso.FileExists(attachmentPath) Then
            eml.Attachments.Add attachmentPath
        Else
            Err.Raise fileNotFound, _
                      , _
                      "Could not find the file that you wanted to " & _
                          "attach: """ & attachmentPath & """."
        End If
        
        Set fso = Nothing
    End If
            
    ' Display the email.
    eml.Display
    
    ' Clean up eml.
    Set eml = Nothing
    
    ' Clean up ol.
    Set ol = Nothing
    
    Exit Sub

ErrorHandler:
    ' Clean up fso.
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    
    ' Clean up eml.
    If Not eml Is Nothing Then
        Set eml = Nothing
    End If
    
    ' Clean up ol.
    If Not ol Is Nothing Then
        Set ol = Nothing
    End If
    
    ' Save information about the error.
    errSrc = Err.Source
    errNum = Err.Number
    Select Case errNum
        Case outlookActionBlocked
            errDesc = _
                "The Outlook operation was blocked by a security feature."
        Case Else
            errDesc = Err.Description
    End Select
    
    ' Regenerate the error.
    Err.Clear
    Err.Raise errNum, errSrc, errDesc
End Sub
