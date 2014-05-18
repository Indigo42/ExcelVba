VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdCreateTestEmail_Click()
    On Error GoTo ErrorHandler
    
    Dim attachmentPath As String
    Dim attachmentPrompt As String
    Dim confirmationPrompt As String
    Dim confirmationResponse As VbMsgBoxResult
    
    ' Confirm that the user wants to continue.
    confirmationPrompt = _
        "Are you sure that you want to create and display a test email?"
    confirmationResponse = MsgBox(confirmationPrompt, vbYesNo)
    If confirmationResponse = vbNo Then Exit Sub
    
    ' Get the path of the attachment, if applicable.
    attachmentPrompt = _
        "Please enter the path of the file that you want to attach, or " & _
        "click ""Cancel"" if you do not want to attach a file."
    attachmentPath = InputBox(attachmentPrompt)
    
    ' Create and display the test email.
    CreateTestEmail attachmentPath

    Exit Sub
    
ErrorHandler:
    ' Display an error message.
    MsgBox Err.Description, vbCritical
End Sub
