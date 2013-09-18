Attribute VB_Name = "m_Add_Activity_Row"
'v4.2.1

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
                    .Style = "ActivityName"
                Case 16
                    .Style = "Normal"
                    .Font.Bold = True
                Case Else
                    .Style = "Normal"
                    .Font.Bold = False
            End Select
        End With
    Next i
End Function


Public Sub mUpdateHourLabels()
    Dim i%, intDateCol%
    Dim strHrs$, strHrsToEight$
    Dim dblHrs#, dblHrsToEight#
    With frmAddActivity
        If .cmbDate.Value <> vbNullString Then
            For i = 1 To 100
                If ActiveSheet.Cells(i, 1).Value Like "Total:" Then
                    intDateCol = .cmbDate.ListIndex + 2
                    dblHrs = ActiveSheet.Cells(i, intDateCol).Value
                    strHrs = Format(dblHrs, "#0.00")
                    If dblHrs = 1 Then strHrs = strHrs & " hour" Else strHrs = strHrs & " hours"
                    .lblDayTotalHours = "This day's current total is " & strHrs

                    dblHrsToEight = 8 - dblHrs
                    strHrsToEight = Format(dblHrsToEight, "#0.00")
                    If dblHrsToEight = 1 Then strHrsToEight = strHrsToEight & " hour" Else strHrsToEight = strHrsToEight & " hours"
                    Select Case dblHrsToEight
                        Case Is > 0
                            .lblHrsToEight = "You need " & strHrsToEight & " to reach 8"
                            .lblHrsToEight.ForeColor = &HC0&        'red
                        Case 0
                            .lblHrsToEight = "You right at 8 hours for this day!"
                            .lblHrsToEight.ForeColor = &H8000&      'green
                        Case Is < 0
                            .lblHrsToEight = "You're over 8 hours by " & Replace(strHrsToEight, "-", "", , , vbTextCompare)
                            .lblHrsToEight.ForeColor = &HC000C0     'pink
                    End Select
                    Exit Sub
                End If
            Next i
        Else
            .lblDayTotalHours = vbNullString
            .lblHrsToEight = vbNullString
        End If
    End With
End Sub
