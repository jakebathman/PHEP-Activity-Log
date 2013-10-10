Attribute VB_Name = "m_New_Sheet"
'v4.3

Public boolSheetAdded As Boolean
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

    strNewSheetName = "FY" & Right(CStr(intYear), 2) & "-" & strPayPeriod
    strReallyNewSheetName = "FY" & Right(CStr(intNextYear), 2) & "-" & strNextPayPeriod
    strOldSheetName = "FY" & Right(CStr(intOldYear), 2) & "-" & strOldPayPeriod

    Call updateWorksheetExists

    For i = 2 To 124
        With Sheets("Refs")
            If .Cells(i, 4).Value <= dtToday And .Cells(i, 5).Value >= dtToday Then
                If .Range("Y" & i).Value = False Then
                    Call fAddSheet(.Range("X" & i).Value, "MAIN", .Range("G" & i).Value, .Range("C" & i).Value)
                Else
                    frmAddSheet.Show
                End If
                Exit For
            End If
        End With
    Next i


    Unload frmAddSheet
    Call updateWorksheetExists


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
        .Sheets("templatesheet").Copy After:=.Sheets("MAIN")
        Set sht = .Sheets(.Sheets("MAIN").Index + 1)
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
