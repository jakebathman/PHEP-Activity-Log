Attribute VB_Name = "m_Formulas"
'v4.2.1

Option Explicit

Public Sub mFormulas(ByVal intTotRow%, ByVal intHeadRow%)
    Dim i%, j%
    Dim strBotRange$, strRightRange$
    Dim rngBottom As Range
    Dim rngRight As Range
    Dim rngSuperTot As Range
    Dim rngWeekOne As Range, rngWeekTwo As Range
    Dim rngBelowTable As Range
    Dim rngWeekOneTotal As Range, rngWeekTwoTotal As Range
    Dim rngWeekendDays As Range


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

    Set rngWeekOne = Range(Cells(intTotRow, 2), Cells(intTotRow, 8))
    Set rngWeekTwo = Range(Cells(intTotRow, 9), Cells(intTotRow, 15))
    Set rngBelowTable = Range(Cells(intTotRow + 1, 1), Cells(intTotRow + 200, 20))

    'rngBelowTable.Select
    rngBelowTable.Clear

    Set rngWeekOneTotal = Cells(intTotRow + 1, 8)
    Set rngWeekTwoTotal = Cells(intTotRow + 1, 15)

    With rngWeekOneTotal
        .Formula = "=sum(" & rngWeekOne.Address(False, False) & ")"
        .Offset(0, -1).Value = "Week 1 Total:"
        .Style = "Calculation"
        .Offset(0, -1).Style = "ActivityName"
    End With
    With rngWeekTwoTotal
        .Formula = "=sum(" & rngWeekTwo.Address(False, False) & ")"
        .Offset(0, -1).Value = "Week 2 Total:"
        .Style = "Calculation"
        .Offset(0, -1).Style = "ActivityName"
    End With

    '   Set Conditional Formatting

    Set rngBottom = Range(Cells(intTotRow, 2), Cells(intTotRow, 15))
    Set rngRight = Range(Cells(intHeadRow + 1, 16), Cells(intTotRow - 1, 16))
    Set rngSuperTot = Range(Cells(intTotRow, 16).Address)
    Set rngWeekendDays = Range(Range(Cells(intTotRow, 7), Cells(intTotRow, 8)).Address(False, False) & "," & Range(Cells(intTotRow, 14), Cells(intTotRow, 15)).Address(False, False))

    Call fSetFormatConditions(rngBottom, "=8", True)
    Call fSetFormatConditions(rngRight, "=80", False)
    Call fSetFormatConditions(rngSuperTot, "=80", True)
    Call fSetFormatConditions(rngWeekOneTotal, "=40", True)
    Call fSetFormatConditions(rngWeekTwoTotal, "=40", True)
    Call fSetFormatConditions(rngWeekendDays, "=8", False)



End Sub



Public Function fSetFormatConditions(rng As Range, fm As String, Optional boolHasLowerLimit As Boolean)
    rng.FormatConditions.Delete

    'Greater Than
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=fm
    rng.FormatConditions(1).SetFirstPriority

    With rng.FormatConditions(1).Font
        '        .Bold = True
        '        .Italic = False
        '        .ColorIndex = xlAutomatic
        '        .TintAndShade = 0
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With rng.FormatConditions(1).Interior
        '        .PatternColorIndex = 0
        '        .Color = 16101350
        '        .TintAndShade = 0
        '        .PatternTintAndShade = 0
        .PatternColorIndex = xlAutomatic
        .Color = 11513845
        .TintAndShade = 0
    End With
    rng.FormatConditions(1).StopIfTrue = False

    If boolHasLowerLimit Then
        'Equal To
        rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:=fm
        'rng.FormatConditions(2).SetFirstPriority

        With rng.FormatConditions(2).Font
            .Bold = True
            .Italic = False
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
        With rng.FormatConditions(2).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 12645807
            .TintAndShade = 0
        End With
        rng.FormatConditions(2).StopIfTrue = False


        '        'Less Than
        '        rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:=fm
        '        'rng.FormatConditions(3).SetFirstPriority
        '
        '        With rng.FormatConditions(3).Font
        '            .Bold = True
        '            .Italic = False
        '            .ColorIndex = xlAutomatic
        '            .TintAndShade = 0
        '        End With
        '        With rng.FormatConditions(3).Interior
        '            .PatternColorIndex = xlAutomatic
        '            .Color = 11513845
        '            .TintAndShade = 0
        '        End With
        '        rng.FormatConditions(3).StopIfTrue = False
    End If

End Function

'GREATER THAN
'    With Selection.FormatConditions(1).Font
'        .Bold = True
'        .Italic = False
'        .ColorIndex = xlAutomatic
'        .TintAndShade = 0
'    End With
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = 0
'        .Color = 16101350
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
'    Range("P7").Select

'EQUAL TO
'    Range("B8:C8").Select
'    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=8"
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    With Selection.FormatConditions(1).Font
'        .Bold = True
'        .Italic = False
'        .ColorIndex = xlAutomatic
'        .TintAndShade = 0
'    End With
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = xlAutomatic
'        .Color = 12645807
'        .TintAndShade = 0
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False

'LESS THAN
'    Range("B8:C8").Select
'    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=8"
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    With Selection.FormatConditions(1).Font
'        .Bold = True
'        .Italic = False
'        .ColorIndex = xlAutomatic
'        .TintAndShade = 0
'    End With
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = xlAutomatic
'        .Color = 11513845
'        .TintAndShade = 0
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
'




