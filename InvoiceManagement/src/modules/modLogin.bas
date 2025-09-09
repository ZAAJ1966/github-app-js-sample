Attribute VB_Name = "modLogin"
Option Explicit

Public Function ValidateLogin(uName As String, pWord As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Worksheets("Users")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow 'assuming headers in row1
        If ws.Cells(r, 1).Value = uName And ws.Cells(r, 2).Value = pWord Then
            gCurrentUser = uName
            gCurrentRole = ws.Cells(r, 3).Value
            ValidateLogin = True
            Exit Function
        End If
    Next r
    ValidateLogin = False
End Function