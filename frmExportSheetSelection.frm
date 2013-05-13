VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportSheetSelection 
   Caption         =   "Which sheets do you want to export?"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   OleObjectBlob   =   "frmExportSheetSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportSheetSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v3

Option Explicit





Private Sub btnCancel_Click()
    Unload frmExportSheetSelection
End Sub

Private Sub btnExport_Click()
    Dim v, i
    Dim s1, s2, s3
        
    With frmExportSheetSelection
        If .chk1.Value + .chk2.Value + .chk3.Value = False Then Unload frmExportSheetSelection: Exit Sub
        If .chk1.Value = True Then s1 = Trim(Mid(.chk1.Caption, 1, InStr(1, .chk1.Caption, " ", vbTextCompare)))
        If .chk2.Value = True Then s2 = Trim(Mid(.chk2.Caption, 1, InStr(1, .chk2.Caption, " ", vbTextCompare)))
        If .chk3.Value = True Then s3 = Trim(Mid(.chk3.Caption, 1, InStr(1, .chk3.Caption, " ", vbTextCompare)))
    End With

    Call mExportSheets(s1, s2, s3)
        
End Sub

Private Sub cmbMonths_Change()

Dim arrSheets(), arrPeriodsAndMonth()
Dim i%, m%, j%
Dim v, x, y, z
Dim boolHasExportableStuff As Boolean

'If Not frmExportSheetSelection.Visible Then Exit Sub

frmExportSheetSelection.chk1.Visible = False
frmExportSheetSelection.chk2.Visible = False
frmExportSheetSelection.chk3.Visible = False
frmExportSheetSelection.lblNothingToExport.Visible = False

    i = 1
    For Each v In ThisWorkbook.Sheets
        If Left(v.Name, 3) Like "FY1" Then
            ReDim Preserve arrSheets(1 To i)
            arrSheets(i) = v.Name
            i = i + 1
        End If
    Next
    
    ReDim arrPeriodsAndMonth(1 To 12, 1 To 3, 1 To 4)
        
    For i = 1 To UBound(arrSheets)
        m = fMonthToInteger(ThisWorkbook.Sheets(arrSheets(i)).Range("B3").Value)
        For j = 1 To 3
            If arrPeriodsAndMonth(m, j, 1) = vbNullString Then Exit For
        Next j
        arrPeriodsAndMonth(m, j, 1) = arrSheets(i)
        arrPeriodsAndMonth(m, j, 2) = ThisWorkbook.Sheets(arrSheets(i)).Range("B3").Value
        arrPeriodsAndMonth(m, j, 3) = ThisWorkbook.Sheets(arrSheets(i)).Range("B5").Value
        arrPeriodsAndMonth(m, j, 4) = ThisWorkbook.Sheets(arrSheets(i)).Range("O5").Value
    Next i

    m = fMonthToInteger(frmExportSheetSelection.cmbMonths.Value)
    boolHasExportableStuff = False
    For j = 1 To 3
        v = arrPeriodsAndMonth(m, j, 2) 'Month Name
        x = arrPeriodsAndMonth(m, j, 1) 'Period (sheet) name
        y = arrPeriodsAndMonth(m, j, 3) 'Period start date
        z = arrPeriodsAndMonth(m, j, 4) 'Period end date
        If v <> vbNullString Then
            If v Like frmExportSheetSelection.cmbMonths.Value Then
                boolHasExportableStuff = True
                y = MonthName(Month(y), True) & " " & Day(y)
                z = MonthName(Month(z), True) & " " & Day(z)
                frmExportSheetSelection.Controls("chk" & j).Caption = x & " (" & y & " to " & z & ")"
                frmExportSheetSelection.Controls("chk" & j).Visible = True
            End If
        End If
    Next j
    If boolHasExportableStuff = False Then frmExportSheetSelection.lblNothingToExport.Visible = True







End Sub

Private Sub UserForm_Activate()
    With frmExportSheetSelection
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_Initialize()
Dim strMonths$
Dim v, c

strMonths = "January;February;March;April;May;June;July;August;September;October;November;December"

    v = Split(strMonths, ";", -1, vbTextCompare)
    For c = 0 To 11 'starts at 0 because the variant array v() does so
        frmExportSheetSelection.cmbMonths.AddItem v(c)
    Next c
    
    MonthName (Month(DateSerial(Year(Now), Month(Now) - 1, Day(Now))))
    frmExportSheetSelection.cmbMonths.Value = MonthName(Month(DateSerial(Year(Now), Month(Now) - 1, Day(Now))))

    

End Sub
