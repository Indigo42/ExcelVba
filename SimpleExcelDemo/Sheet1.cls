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

Private Sub cmdRunTheDemo_Click()
    Dim inputIsValid As Boolean
    Dim userInput As String

    inputIsValid = False
    Do Until inputIsValid
        userInput = InputBox("Please enter a number.")
        
        ' Exit if the user clicked Cancel or typed nothing in the text box.
        If userInput = "" Then Exit Sub
        
        If IsNumeric(userInput) Then
            inputIsValid = True
        Else
            MsgBox "Invalid input.  Please try again.", vbCritical
        End If
    Loop
    
    ' Write to an arbitrary cell.
    Me.Range("A6").Value = CDbl(userInput)
End Sub
