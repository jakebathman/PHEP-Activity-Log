Attribute VB_Name = "MAIN"

Option Explicit

Public Sub mMain(ByVal a$, ByVal d$, ByVal t#, ByVal c%, Optional boolAddActivityAndAnother As Boolean)

'this sub is called when the Add Activity dialog box is filled out

    Dim i%, j%, intTotalsRow%, intHeaderRow%, intFirstEmptyActRow%
    
    
    
    Call mHideSOMEOFTHETHINGS(True, True)



    Call mAddActivityRow(a, d, t, c)
    
    Unload frmAddActivity
    
    If boolAddActivityAndAnother Then frmAddActivity.Show
    
    Call fCalcLocations(intHeaderRow, intFirstEmptyActRow, intTotalsRow)
    
    Call mSortRows(intHeaderRow, intTotalsRow)
    
    Call mFormulas(intTotalsRow, intHeaderRow)

End Sub
