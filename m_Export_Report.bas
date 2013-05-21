Attribute VB_Name = "m_Export_Report"
'v4.1

Option Explicit

Public Sub mExportMonthlyReport()

    Dim v
    ' This code fixes the report month formula in "Refs" and can be removed in later versions (but won't hurt anything to stay in)
    With ThisWorkbook.Sheets("Refs")
        If .Range("I21").Value Like "September" Or .Range("I21").Value Like "October" Then
            v = ActiveSheet.Name
            .Visible = xlSheetVisible
            .Range("I2:I124").ClearContents
            .Range("I2").FormulaR1C1 = "=CHOOSE(IF(MONTH(RC[-3])+1=13,1,MONTH(RC[-3])+1),""December"",""January"",""February"",""March"",""April"",""May"",""June"",""July"",""August"",""September"",""October"",""November"")"
            .Activate
            .Range("I2").AutoFill Destination:=Range("I2:I124"), Type:=xlFillDefault
            Sheets(v).Activate
            .Visible = xlSheetVeryHidden
        End If
    End With

    frmExportSheetSelection.Show


End Sub




Public Function fMonthToInteger(strMonth$)
    Select Case strMonth
        Case "January": fMonthToInteger = 1
        Case "February": fMonthToInteger = 2
        Case "March": fMonthToInteger = 3
        Case "April": fMonthToInteger = 4
        Case "May": fMonthToInteger = 5
        Case "June": fMonthToInteger = 6
        Case "July": fMonthToInteger = 7
        Case "August": fMonthToInteger = 8
        Case "September": fMonthToInteger = 9
        Case "October": fMonthToInteger = 10
        Case "November": fMonthToInteger = 11
        Case "December": fMonthToInteger = 12
    End Select
End Function

Public Function fIntegerToMonth(intInt%)
    Select Case intInt
        Case 1: fIntegerToMonth = "January"
        Case 2: fIntegerToMonth = "February"
        Case 3: fIntegerToMonth = "March"
        Case 4: fIntegerToMonth = "April"
        Case 5: fIntegerToMonth = "May"
        Case 6: fIntegerToMonth = "June"
        Case 7: fIntegerToMonth = "July"
        Case 8: fIntegerToMonth = "August"
        Case 9: fIntegerToMonth = "September"
        Case 10: fIntegerToMonth = "October"
        Case 11: fIntegerToMonth = "November"
        Case 12: fIntegerToMonth = "December"
    End Select
End Function




Public Sub mExportSheets(ByVal strSheet1$, Optional ByVal strSheet2$, Optional ByVal strSheet3$)

    Unload frmExportSheetSelection
    ThisWorkbook.Save

    Dim strFileName$, strEmpName$
    Dim strOutputPrefix$
    Dim dtExportDate As Date
    Dim strCurDir As String
    Dim strAWName As String
    Dim strNewWBName As String
    Dim strTempNewWBName As String
    Dim vbOpenFolder
    Dim strPath As String
    Dim intLastRow%, i%, intNumSheetsToExport%, intNextRow%
    Dim rngData As Range
    Dim cSheets As New Collection
    Dim S, v, arr()

    If strSheet2 = vbNullString Then
        intNumSheetsToExport = 1
        cSheets.Add strSheet1
    ElseIf strSheet3 = vbNullString Then
        intNumSheetsToExport = 2
        cSheets.Add strSheet1
        cSheets.Add strSheet2
    Else
        intNumSheetsToExport = 3
        cSheets.Add strSheet1
        cSheets.Add strSheet2
        cSheets.Add strSheet3
    End If

    i = 1
    'Sort the sheets
    If cSheets.Count > 1 Then
        ReDim arr(1 To cSheets.Count)
        For Each S In cSheets
            arr(i) = S
            i = i + 1
        Next
        For i = 1 To cSheets.Count
            cSheets.Remove 1
        Next
        Call QSortInPlace(arr, -1, -1, False, vbTextCompare, False)
        For i = 1 To UBound(arr)
            cSheets.Add arr(i)
        Next i
    End If


    If ThisWorkbook.Sheets("Refs").Range("N2").Value = vbNullString Then
        strEmpName = Trim(StrConv(InputBox("Sorry, I don't know who you are!" & vbNewLine & vbNewLine & "Please enter your full name.", "Who are you?!"), vbProperCase))
        ThisWorkbook.Sheets("Refs").Range("N2").Value = strEmpName
    Else
        strEmpName = ThisWorkbook.Sheets("Refs").Range("N2").Value
    End If

    dtExportDate = DateSerial(Year(ThisWorkbook.Sheets(strSheet1).Range("B5").Value), fMonthToInteger(ThisWorkbook.Sheets(strSheet1).Range("B3").Value), 1)

    strOutputPrefix = Format(dtExportDate, "yyyy.mm") & " " & Trim(Mid(strEmpName, InStr(1, strEmpName, " "))) & " monthly activity log"
    strNewWBName = strOutputPrefix & ".xlsx"

    strAWName = ThisWorkbook.Name
    strCurDir = ThisWorkbook.Path
    Workbooks.Add

    strTempNewWBName = ActiveWorkbook.Name
    intNextRow = 1

    'Workbooks(strCurWBName).Sheets(strCurWSName).Activate
    For Each S In cSheets
        For i = 1 To 50
            If Workbooks(strAWName).Sheets(S).Cells(i, 1) = "Total:" Then intLastRow = i: Exit For
        Next i

        With Workbooks(strAWName).Sheets(S)
            Set rngData = .Range(.Cells(5, 1), .Cells(intLastRow, 16))
        End With

        rngData.Copy
        Application.DisplayAlerts = False
        Workbooks(strTempNewWBName).Sheets(1).Cells(intNextRow, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Workbooks(strTempNewWBName).Sheets(1).Cells(intNextRow, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Workbooks(strTempNewWBName).Sheets(1).Name = Trim(Mid(strEmpName, InStr(1, strEmpName, " "))) & " " & Format(dtExportDate, "yyyy.mm")
        Application.CutCopyMode = False

        intNextRow = intNextRow + intLastRow - 2

    Next

    With Workbooks(strTempNewWBName).Sheets(1)
        .Rows(1).Insert (1)
        .Range("A1").Value = strEmpName
        .Range("A2").Select
        Selection.Copy
        .Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Range("A1").HorizontalAlignment = xlCenter
        Application.CutCopyMode = False
    End With

    With Workbooks(strTempNewWBName)
        .Sheets("Sheet2").Delete
        .Sheets("Sheet3").Delete
        .Sheets(1).Columns("A:Z").EntireColumn.AutoFit
    End With

    Call fSetPrintSettings(Workbooks(strTempNewWBName).Sheets(1))


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''


    strPath = strCurDir & "\Monthly Activity Reports\"
    strFileName = strPath & strNewWBName
    If Len(Dir(strPath, vbDirectory)) = 0 Then
        MkDir strPath
    End If
    ActiveWorkbook.SaveAs FileName:=strFileName, FileFormat:=XlFileFormat.xlOpenXMLWorkbook, CreateBackup:=False


    Application.DisplayAlerts = True

    vbOpenFolder = MsgBox("The file was exported successfully. You may find it in the same local directory as this workbook, in a new folder called Monthly Activity Reports" & vbCrLf & vbCrLf _
                        & "Would you like to put a copy on the PHEP drive as well?", vbYesNo)

    If vbOpenFolder = vbYes Then
        'Shell "explorer.exe " & strPath, vbNormalFocus
        strPath = "\\ccdata01\homeland_security\PHEP Documentation\Monthly Reports\" & Left(strSheet1, 4) & "\"
        strFileName = strPath & strNewWBName
        ActiveWorkbook.SaveAs FileName:=strFileName, FileFormat:=XlFileFormat.xlOpenXMLWorkbook, CreateBackup:=False
    End If


    Workbooks(strNewWBName).Close


End Sub




Public Function fSetPrintSettings(ByRef S As Worksheet)

    With S.PageSetup
        .Orientation = xlLandscape
        .Zoom = 100
    End With
End Function

