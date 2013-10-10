Attribute VB_Name = "m_Setup_Worksheet_Exist"
Option Explicit

Public Sub mSetupWorksheetExists()

    Dim i%, j%
    Dim shtRefs As Worksheet
    Dim intWorksheetNameCol%, intWorksheetExistsCol%

    Set shtRefs = ThisWorkbook.Sheets("Refs")

    ' Look for columns WorksheetName and WorksheetExists
    ' Should be at column X and Y, respectively
    Debug.Print shtRefs.Range("X2").FormulaR1C1

    With shtRefs
        For j = 1 To 100
            If StrComp(.Cells(1, j).Value, "WorksheetName", vbTextCompare) = 0 Then
                intWorksheetNameCol = j
                If StrComp(.Cells(1, j + 1).Value, "WorksheetExists", vbTextCompare) = 0 Then
                    intWorksheetExistsCol = j + 1
                    Exit For
                End If
            End If
        Next j
    End With

    If intWorksheetExistsCol = 0 Or intWorksheetNameCol = 0 Then
        shtRefs.Range("X:X").Clear
        shtRefs.Range("Y:Y").Clear
        intWorksheetNameCol = 24        ' x
        intWorksheetExistsCol = 25      ' y
        shtRefs.Range("X1").Value = "WorksheetName"
        shtRefs.Range("Y1").Value = "WorksheetExists"
        shtRefs.Range("X2:X124").FormulaR1C1 = "=""FY"" & RIGHT(RC[-17],2) & ""-"" & IF(RC[-21]<10,""0"" & RC[-21],RC[-21])"
    End If

    Call updateWorksheetExists

End Sub

