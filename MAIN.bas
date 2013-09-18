Attribute VB_Name = "MAIN"
'v4.2.1

Option Explicit

Public Sub mMain(ByVal a$, ByVal d$, ByVal t#, ByVal c%, Optional boolAddActivityAndAnother As Boolean)

    'this sub is called when the Add Activity dialog box is filled out

    Dim i%, j%, intTotalsRow%, intHeaderRow%, intFirstEmptyActRow%


    Application.ScreenUpdating = False

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



