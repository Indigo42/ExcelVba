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
'   Procedure name:         FixComments()
'
'   Description:
'
' ****************************************************************************
Public Sub FixComments(ByRef wb As Workbook)
    Dim c As Comment
    Dim ws As Worksheet
         
    For Each ws In wb.Worksheets
        For Each c In ws.Comments
            With c.Shape
                .Left = c.Parent.Left
                .Top = c.Parent.Top

                .Width = 200
                .Height = 200
            End With
        Next c
    Next ws
End Sub
