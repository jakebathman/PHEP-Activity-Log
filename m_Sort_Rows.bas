Attribute VB_Name = "m_Sort_Rows"
'v4.6

Option Explicit

Public Sub mSortRows(intHeaderRow, intTotalsRow)
    Dim sh As Worksheet
    Dim rn As Range
    Dim ky As Range

    Set sh = ActiveWorkbook.ActiveSheet
    Set rn = sh.Range(Cells(intHeaderRow, 1), Cells(intTotalsRow - 1, 16))
    Set ky = sh.Range(Cells(intHeaderRow + 1, 1), Cells(intTotalsRow - 1, 1))

    sh.Sort.SortFields.Clear
    sh.Sort.SortFields.Add Key:=ky, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sh.Sort
        .SetRange rn
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub

