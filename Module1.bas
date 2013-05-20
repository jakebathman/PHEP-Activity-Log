Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'    ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False _
'        , DisplayAsIcon:=False, Left:=221.875, Top:=330, Width:=286.875, _
'        Height:=103.75).Select
        
Dim Obj As Object
Dim Code As String

Sheets("MAIN").Select

'create button
    Set Obj = ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=220, Top:=220, Width:=275, Height:=100)
    'Set Obj = ActiveSheet.OLEObjects("frmBugButton").Object
    Obj.Name = "frmBugBtn"
'button text
    With ActiveSheet.OLEObjects("frmBugBtn").Object
        .Caption = ":( " & vbCrLf & vbCrLf & "Something 's Broken" & vbCrLf & "(report a bug)"
        .BackColor = &HC0&
        .ForeColor = &HFFFFFF
        .Font.Size = 14
        .Font.Bold = True
    End With
        
End Sub
