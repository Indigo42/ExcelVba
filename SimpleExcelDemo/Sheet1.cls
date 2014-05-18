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
    Dim targetRange As Range
    Dim userInput As String
    
    Set targetRange = Me.Range("A6")    ' Arbitrary cell.

    inputIsValid = False
    Do Until inputIsValid
        userInput = InputBox("Please enter a number.")
        
        ' If the user clicked Cancel (or typed nothing in the text box), exit
        ' the loop.
        If userInput = "" Then Exit Do
        
        If IsNumeric(userInput) Then
            inputIsValid = True
            targetRange.Value = CDbl(userInput)
        Else
            MsgBox "Invalid input.  Please try again.", vbCritical
        End If
    Loop
    
    Set targetRange = Nothing
End Sub
