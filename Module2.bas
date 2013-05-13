Attribute VB_Name = "Module2"
'v4

Option Explicit


Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'

    Range("B8:O8").Select
    Cells.FormatConditions.Delete
    Range("B8:O8").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = 0
        .Color = 16101350
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("P7").Select
    Range("B8").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=80"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16777024
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = 0
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799951170384838
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("P8").Select
    Range("B8").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=80"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16777024
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = 0
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799951170384838
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("B8:O8").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 12645807
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("B8:O8").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 11513845
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWorkbook.Save
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("B8").Select
    Selection.FormatConditions.Delete
    
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Range("B8:C8").Select
    Cells.FormatConditions.Delete
    Range("B8:C8").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = 0
        .Color = 16101350
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("P7").Select
    
    Range("B8:C8").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 12645807
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("B8:C8").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 11513845
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
