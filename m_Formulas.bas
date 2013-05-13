Attribute VB_Name = "m_Formulas"
'v2

Option Explicit

Public Sub mFormulas(ByVal intTotRow%, ByVal intHeadRow%)
Dim i%, j%
Dim strBotRange$, strRightRange$

strBotRange = Cells(intTotRow, 2).Address(False, False) & ":" _
            & Cells(intTotRow, 16).Address(False, False)

strRightRange = Cells(intHeadRow + 1, 16).Address(False, False) & ":" _
            & Cells(intTotRow - 1, 16).Address(False, False)


For j = 2 To 16
    Cells(intTotRow, j).Formula = "=sum(" & Range(Cells(intHeadRow + 1, j), Cells(intTotRow - 1, j)).Address(False, False) & ")"
Next j

For i = intHeadRow + 1 To intTotRow - 1
    Cells(i, 16).Formula = "=sum(" & Range(Cells(i, 2), Cells(i, 15)).Address(False, False) & ")"
Next i

'   Set Conditional Formatting

Dim rngBottom As Range
Dim rngRight As Range
Dim rngSuperTot As Range

Set rngBottom = Range(Cells(intTotRow, 2), Cells(intTotRow, 15))
Set rngRight = Range(Cells(intHeadRow + 1, 16), Cells(intTotRow - 1, 16))
Set rngSuperTot = Range(Cells(intTotRow, 16).Address)

Call fSetFormatConditions(rngBottom, "=8")
Call fSetFormatConditions(rngRight, "=80")
Call fSetFormatConditions(rngSuperTot, "=80")



End Sub



Public Function fSetFormatConditions(rng As Range, fm As String)
rng.FormatConditions.Delete

    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=fm
    rng.FormatConditions(1).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16777024
        .TintAndShade = 0
    End With
    With rng.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
    End With
    rng.FormatConditions(1).StopIfTrue = False
    
End Function
