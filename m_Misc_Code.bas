Attribute VB_Name = "m_Misc_Code"
Option Explicit

Public Sub mShowALLTHETHINGS()
Dim sh As Worksheet
    For Each sh In Application.Worksheets
        Debug.Print sh.Name
        Debug.Print sh.Visible
        If sh.Visible = xlSheetHidden Or sh.Visible = xlSheetVeryHidden Or sh.Visible = False Then
            sh.Visible = xlSheetVisible
        End If
    Next sh
    'Sheet2.Visible = xlSheetVeryHidden
    'Sheet4.Visible = xlSheetVeryHidden
End Sub

Public Sub mHideSOMEOFTHETHINGS(HideRefs As Boolean, HideTemplate As Boolean)
    If HideRefs Then Sheet2.Visible = xlSheetVeryHidden 'Refs
    If HideTemplate Then Sheet4.Visible = xlSheetVeryHidden 'templatesheet
End Sub

