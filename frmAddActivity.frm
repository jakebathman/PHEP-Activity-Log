VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddActivity 
   Caption         =   "Add activity to the log"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   OleObjectBlob   =   "frmAddActivity.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v4

Option Explicit

Public rngWorkingRange As Range ' bounded by date row, activity column, totals row, and total column
Public boolEnableEvents As Boolean
Public boolAddAnother As Boolean





Private Sub btnCancel_Click()
    Unload frmAddActivity
End Sub

Public Sub btnSave_Click()
boolAddAnother = False
Call btnSaveAndAdd_Click

'    With frmAddActivity
'        .lblBlankMessage.Visible = False
'        If .cmbActivity.Value <> vbNullString And .cmbDate <> vbNullString And .txtHours.Value <> vbNullString Then
'            Call mMain(frmAddActivity.cmbActivity.Value, frmAddActivity.cmbDate.Value, frmAddActivity.txtHours.Value, frmAddActivity.cmbDate.ListIndex, False)
'        Else
'            .lblBlankMessage.Visible = True
'        End If
'    End With
End Sub

Public Sub btnSaveAndAdd_Click()
    Dim v

    With frmAddActivity
        .lblBlankMessage.Visible = False
        .lblZeroHrs.Visible = False
        On Error Resume Next
        v = Round(.txtHours.Value / 0.25, 0) * 0.25
        .txtHours.Value = v
        
        On Error GoTo errFixBlankTime
        If .cmbActivity.Value <> vbNullString And .cmbDate <> vbNullString And .txtHours.Value <> vbNullString And .txtHours.Value >= 0.25 Then
            Call mMain(frmAddActivity.cmbActivity.Value, frmAddActivity.cmbDate.Value, frmAddActivity.txtHours.Value, frmAddActivity.cmbDate.ListIndex, boolAddAnother)
        ElseIf cmbActivity.Value = vbNullString Or .cmbDate = vbNullString Or .txtHours.Value = vbNullString Then
            .lblBlankMessage.Visible = True
        ElseIf CDbl(.txtHours.Value) = 0 Then
            .lblZeroHrs.Visible = True
        End If
        On Error GoTo 0
    End With

    Exit Sub

errFixBlankTime:
    If frmAddActivity.txtHours.Value = vbNullString Then frmAddActivity.txtHours.Value = 0
    Resume

End Sub

Private Sub cmbActivity_Change()
    If frmAddActivity.cmbDate.Value = vbNullString Then
        frmAddActivity.cmbDate.SetFocus
    Else
        frmAddActivity.cmbDate.SetFocus
        frmAddActivity.txtHours.SetFocus
    End If
End Sub

Private Sub cmbDate_Change()
    Call mUpdateHourLabels
    frmAddActivity.txtHours.SetFocus
End Sub

Private Sub cmbDate_Enter()
    frmAddActivity.cmbDate.DropDown
End Sub





Private Sub txtHours_Change()
Dim v
If Not boolEnableEvents Then Exit Sub
boolEnableEvents = False

v = frmAddActivity.txtHours.Value
    If Len(v) = 1 And v = "." Then frmAddActivity.txtHours.Value = "0."
    If Not IsNumeric(frmAddActivity.txtHours.Value) Then
        If frmAddActivity.txtHours.Value <> vbNullString Then
            Call MsgBox("Please enter decimal numbers only!"): frmAddActivity.txtHours.Value = vbNullString
        End If
    End If

boolEnableEvents = True
End Sub

Private Sub UserForm_Activate()
    Dim strNow$, i%, intDateCol%

    With frmAddActivity
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.6 * Application.Height) - (0.5 * .Height)
    End With
    frmAddActivity.lblBlankMessage.Visible = False
    frmAddActivity.lblZeroHrs.Visible = False

    On Error Resume Next
    strNow = WeekdayName(Weekday(Now, vbMonday), True, vbMonday) & " " & Format(DateSerial(Year(Now), Month(Now), Day(Now)), "Short Date", vbMonday)
    If Sheets("Refs").Range("O2").Value <> vbNullString Then frmAddActivity.cmbDate.Value = strNow
    On Error GoTo 0

    Call mUpdateHourLabels

    frmAddActivity.cmbActivity.DropDown
End Sub

Private Sub UserForm_Initialize()
    Dim strActiveSheetName$
    Dim i%, j%
    Dim intHeaderRow%, intFirstBlankRow%, intNumActivities%, intTotalsRow%
    strActiveSheetName = ActiveSheet.Name
    intNumActivities = 0
    boolAddAnother = True



    For i = 1 To 50
        If Cells(i, 1).Value = "Activity" Then intHeaderRow = i: Exit For
    Next i

    For i = intHeaderRow To 75
        If Cells(i, 1).Value = vbNullString Then intFirstBlankRow = i: Exit For
    Next i

    Call mUpdateCategories

    For i = 2 To 50
        If Sheets("Refs").Cells(i, 2).Value = vbNullString Then Exit For
        frmAddActivity.cmbActivity.AddItem Sheets("Refs").Cells(i, 2).Value
        intNumActivities = intNumActivities + 1
    Next i
    frmAddActivity.cmbActivity.ListRows = intNumActivities

    For j = 2 To 15
        If ActiveSheet.Cells(intHeaderRow, j).Value = "Total" Then Exit For
        frmAddActivity.cmbDate.AddItem ActiveSheet.Cells(intHeaderRow, j).Value & " " & ActiveSheet.Cells(intHeaderRow - 1, j).Value
    Next j

    For i = 1 To 100
        If ActiveSheet.Cells(i, 1).Value Like "Total:" Then
            intTotalsRow = i: Exit For
        End If
    Next i

    'Set rngWorkingRange = Range(Cells(intHeaderRow - 1, 1), Cells(intTotalsRow, 16))

    boolEnableEvents = True

End Sub
