Attribute VB_Name = "m_New_Sheet"
'v4.1

Option Explicit

Public Sub mNewSheet()

    Dim i%, j%
    Dim dtToday As Date
    Dim dtIndex As Date
    Dim intYear%, intPeriodsBetween%, intCurPayPeriod%, intNextPayPeriod%, intOldPayPeriod%
    Dim strNewSheetName$, strPayPeriod$, strReallyNewSheetName$, strNextPayPeriod$, strOldPayPeriod$, strOldSheetName$
    Dim intNextYear%, intOldYear%
    Dim wsNewSheet As Worksheet
    Dim boolSheetDoesntExist As Boolean

    ActiveWorkbook.Sheets("Refs").Visible = xlVeryHidden
    ActiveWorkbook.Sheets("templatesheet").Visible = xlVeryHidden


    dtToday = Now
    dtIndex = DateSerial(2011, 12, 26)    'index date, starts FY12-01 pay period
    intYear = Year(Now)
    intNextYear = intYear + 1
    intOldYear = intYear - 1
    intPeriodsBetween = DateDiff("ww", dtIndex, dtToday, vbMonday) / 2
    intCurPayPeriod = (intPeriodsBetween + 1) Mod 26
    intNextPayPeriod = (intCurPayPeriod + 1) Mod 26
    intOldPayPeriod = intCurPayPeriod - 1
    If intOldPayPeriod = 0 Then intOldPayPeriod = 26

    strPayPeriod = fMakeTwoDigitPayPeriod(intCurPayPeriod)
    strNextPayPeriod = fMakeTwoDigitPayPeriod(intNextPayPeriod)
    strOldPayPeriod = fMakeTwoDigitPayPeriod(intOldPayPeriod)


    If dtToday > DateSerial(intYear, 8, 31) And dtToday < DateSerial(intYear, 12, 31) Then
        intYear = intNextYear
    End If
    If (dtToday + 14) > DateSerial(intYear, 8, 31) And dtToday < DateSerial(intYear, 12, 31) Then
        intNextYear = intYear + 1
    Else
        intNextYear = intYear
    End If
    If (dtToday - 14) < DateSerial(intYear, 8, 31) And dtToday > DateSerial(intYear, 1, 1) Then
        intOldYear = intYear - 1
    Else
        intOldYear = intYear
    End If


    strNewSheetName = "FY" & Right(CStr(intYear), 2) & "-" & strPayPeriod
    strReallyNewSheetName = "FY" & Right(CStr(intNextYear), 2) & "-" & strNextPayPeriod
    strOldSheetName = "FY" & Right(CStr(intOldYear), 2) & "-" & strOldPayPeriod

    boolSheetDoesntExist = True

    For i = 1 To Sheets.Count
        If StrComp(strNewSheetName, Sheets(i).Name, vbTextCompare) = 0 Then
            boolSheetDoesntExist = False
            If MsgBox("Looks like you already have a tracking sheet for the current pay period." & vbNewLine & vbNewLine _
                    & "If you're getting antsy, you can create one for the next pay period. Do that?", vbYesNo, "Sheet exists!!") = vbYes Then
                Call fAddSheet(strReallyNewSheetName, strNewSheetName, intNextYear, intNextPayPeriod)
            End If
            Exit For
        End If
    Next i

    If boolSheetDoesntExist Then
        'Call fAddSheet(strNewSheetName, strOldSheetName, intYear, intCurPayPeriod)
        Call fAddSheet(strNewSheetName, "MAIN", intYear, intCurPayPeriod)
    End If

    Call MaintenanceForAddActivityButton

End Sub



Public Function fMakeTwoDigitPayPeriod(intPP)
    If intPP < 10 Then
        fMakeTwoDigitPayPeriod = "0" & CStr(intPP)
    Else
        fMakeTwoDigitPayPeriod = CStr(intPP)
    End If
End Function



Public Function fAddSheet(strShtName, strOld, yr, pp)

    Dim sht As Worksheet
    Call mShowALLTHETHINGS
    On Error GoTo errFixMissingSheet
    With ActiveWorkbook
        .Sheets("templatesheet").Copy After:=.Sheets(strOld)
        Set sht = .Sheets(.Sheets(strOld).Index + 1)
    End With
    sht.Name = strShtName
    sht.Range("B1").Value = yr
    sht.Range("B2").Value = pp
    sht.Visible = xlSheetVisible
    sht.Activate

    Sheets("Refs").Range("P2").Value = sht.Name
    Call mHideSOMEOFTHETHINGS(True, True)

    Exit Function


errFixMissingSheet:
    strOld = "MAIN"
    Resume

End Function
