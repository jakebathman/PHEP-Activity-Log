Attribute VB_Name = "f_Calc_Locations"
Option Explicit

Public Function fCalcLocations(ByRef intHeaderRow, ByRef intFirstEmptyActRow, ByRef intTotalsRow)
Dim i%
For i = 1 To 100
    Select Case Cells(i, 1).Value
        Case "Activity"
            intHeaderRow = i
        Case vbNullString
            If intHeaderRow <> 0 Then intFirstEmptyActRow = i
        Case "Total:"
            intTotalsRow = i: Exit For
    End Select
Next i

End Function
