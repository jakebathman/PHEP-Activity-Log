Attribute VB_Name = "m_Add_Activity_Row"
'v3

Option Explicit

Public Sub mAddActivityRow(strAct$, strDate$, dblTime#, intDateCol%)
Dim i%, j%, c%
Dim intHeaderRow%, intFirstEmptyActRow%, intTotalsRow%
Dim intExistingActivityRow%
Dim strActiveSheetName$
Dim boolActExists As Boolean

strActiveSheetName = ActiveSheet.Name
intDateCol = intDateCol + 2

Sheets("Refs").Range("P2").Value = strActiveSheetName

Call fCalcLocations(intHeaderRow, intFirstEmptyActRow, intTotalsRow)

If intTotalsRow < intFirstEmptyActRow Or intFirstEmptyActRow = 0 Then
    Rows(intTotalsRow).Insert
    intFirstEmptyActRow = intTotalsRow
    intTotalsRow = intTotalsRow + 1
End If



'Check for existing row for activity

For i = intHeaderRow + 1 To intTotalsRow - 1
    If Cells(i, 1).Value = strAct Then
        boolActExists = True
        intExistingActivityRow = i
        Exit For
    End If
Next i

If intExistingActivityRow = 0 Then intExistingActivityRow = 500
If boolActExists And Cells(intExistingActivityRow, intDateCol).Value > 0 Then
    Select Case MsgBox("Whoops!" & vbNewLine & vbNewLine & "There's already a value for that activity & date. " _
        & "Add the two together?", vbYesNoCancel, "Activity exists on that date!")
        Case vbYes
            Call fAddNewLineAndActivity(intExistingActivityRow, intDateCol, Cells(intExistingActivityRow, intDateCol).Value + dblTime, strAct)
            Rows(intTotalsRow - 1).Delete
            intTotalsRow = intTotalsRow - 1
        Case vbNo
            Call fAddNewLineAndActivity(intFirstEmptyActRow, intDateCol, dblTime, strAct)
        Case Else
            End
    End Select
ElseIf boolActExists Then
    Rows(intFirstEmptyActRow).Delete
    Call fAddNewLineAndActivity(intExistingActivityRow, intDateCol, dblTime, strAct)
Else
    Call fAddNewLineAndActivity(intFirstEmptyActRow, intDateCol, dblTime, strAct)
End If

End Sub


Public Function fAddNewLineAndActivity(intRow%, intCol%, dblT#, strA$)
    'Add row & info
    Dim i%
    Cells(intRow, 1).Value = strA
    Cells(intRow, intCol).Value = dblT
    For i = 1 To 16
        With Cells(intRow, i)
            Select Case i
                Case 1
                    .Style = "Normal"
                    .HorizontalAlignment = xlRight
                    .Font.Bold = True
                Case 16
                    .Style = "Normal"
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                Case Else
                    .Style = "Normal"
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = False
            End Select
        End With
    Next i
End Function

