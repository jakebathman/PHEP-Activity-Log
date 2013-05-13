VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'v4

Option Explicit

Private Sub btnExportSheet_Click()
    'Call MsgBox("Coming soon...", vbOKOnly + vbExclamation, "Hold your horses!")
    Call mExportMonthlyReport
End Sub

Private Sub btnHistory_Click()
    On Error GoTo errBadSheetName
    Sheets(Sheets("Refs").Range("P2").Value).Activate

Exit Sub

errBadSheetName:
    frmPickWorkingSheet.lblMainLabel.Caption = "Which sheet are you currently working on?" & vbNewLine & vbNewLine & "I'll try to remember it for next time."
    frmPickWorkingSheet.Show
    Sheets(Sheets("Refs").Range("P2").Value).Activate
    Unload frmPickWorkingSheet
End Sub

Private Sub btnNewSheet_Click()
Call mNewSheet
End Sub

Private Sub btnPrefs_Click()
    frmPrefs.Show
End Sub

Private Sub btnUpdateCode_Click()
    If MsgBox("Update this tool to the latest version?" & vbNewLine & vbNewLine & "This will look for a newer version on the PHEP drive.", vbOKCancel) = vbOK Then
        Call uUpdateCode
    End If
End Sub
