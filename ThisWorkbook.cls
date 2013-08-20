VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'v4.2

Option Explicit

Private Sub Workbook_Open()
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
