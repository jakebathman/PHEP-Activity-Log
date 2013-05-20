VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'v4.1

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

Private Sub CommandButton1_Click()

End Sub

Private Sub frmBugButton_Click()
    frmBug.Show
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    With Sheets("Refs")
        If .Range("Q2").Value = "FALSE" Or .Range("Q2").Value = False Or .Range("Q2").Value = vbNullString Then
            If MsgBox("Some back-end code needs to be udpated. You should do that soon. How about now?", vbYesNo + vbExclamation) = vbYes Then
                .Range("Q1").Value = "UpdateCodeInSync"
                .Range("Q2").Value = "FALSE"
                Call uUpdateTheUpdateCode
                Call uOneTimeCode
            End If
        End If
    End With

    

End Sub
