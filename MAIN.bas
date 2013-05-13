Attribute VB_Name = "MAIN"
'v3

Option Explicit

Public Sub mMain(ByVal a$, ByVal d$, ByVal t#, ByVal c%, Optional boolAddActivityAndAnother As Boolean)

'this sub is called when the Add Activity dialog box is filled out

    Dim i%, j%, intTotalsRow%, intHeaderRow%, intFirstEmptyActRow%
    
       
    Application.ScreenUpdating = False
    
    With Sheets("Refs")
        If .Range("Q2").Value = "False" Or .Range("Q2").Value = vbNullString Then
            .Range("Q1").Value = "UpdateCodeInSync"
            .Range("Q2").Value = "FALSE"
            'Call uUpdateTheUpdateCode
        End If
    End With
    
    Call mHideSOMEOFTHETHINGS(True, True)

    Call mAddActivityRow(a, d, t, c)
    
    Unload frmAddActivity
    
    Call fCalcLocations(intHeaderRow, intFirstEmptyActRow, intTotalsRow)
    
    Call mSortRows(intHeaderRow, intTotalsRow)
    
    Call mFormulas(intTotalsRow, intHeaderRow)
    
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    
    If boolAddActivityAndAnother Then frmAddActivity.Show
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = False

End Sub



