Attribute VB_Name = "u_One_Time_Code"
Option Explicit

Public Sub uOneTimeCode()

    Dim strLogFilePathAndName

    strLogFilePathAndName = ActiveWorkbook.Path & "\" & "ErrorLog.txt"

    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Starting One Time Code": Close #1

    ' Check for PatchesInstalled column, and add header if not there
    Dim i%, j%, intLastCol%, intPatchCol%
    With ThisWorkbook.Sheets("Refs")
        ' Find the correct column
        For j = 1 To 200
            If StrComp("PatchesInstalled", .Cells(1, j).Value, vbTextCompare) = 0 Then intPatchCol = j: Exit For
            If .Cells(1, j).Value = vbNullString And .Cells(1, j + 1).Value = vbNullString Then intLastCol = j - 1: Exit For
        Next j
        ' LOG
        Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Done looking for PatchesInstalled column in Refs... intLastCol = " & intLastCol & " and intPatchCol = " & intPatchCol: Close #1

        If intLastCol > 0 Then
            .Cells(1, intLastCol + 1).Value = "PatchesInstalled"
            intPatchCol = intLastCol
            ' LOG
            Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Done looking for PatchesInstalled column in Refs... intLastCol = " & intLastCol & " and intPatchCol = " & intPatchCol: Close #1

        End If
    End With

    ' Force patches to re-install
    ThisWorkbook.Sheets("Refs").Range(Cells(2, intPatchCol), Cells(5, intPatchCol)).Value = vbNullString






    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "About to call patch4_2_1 module": Close #1





    Call patch4_2_1(intPatchCol)
    Call patch4_2_2(intPatchCol)

    Kill strLogFilePathAndName

End Sub


Private Sub patch4_2_1(intPatchesCol As Integer)
    Dim strLogFilePathAndName

    strLogFilePathAndName = ActiveWorkbook.Path & "\" & "ErrorLog.txt"

    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Starting patch v4.2.1": Close #1

    ' Check if the patch is installed first
    Dim i%, j%, intLastColOfRefs%
    With ThisWorkbook.Sheets("Refs")
        For i = 2 To 100
            If .Cells(i, intPatchesCol).Value = "v4.2.1" Then
                Exit Sub
            End If
        Next i
    End With
    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Continuing with patch v4.2.1 (intPatchesCol = " & intPatchesCol & ")": Close #1


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Patch v4.2.1                                        '
    '   Adds Bug Report button on MAIN sheet                '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Obj As Object
    Dim Code As String
    Dim shMainSheet As Worksheet
    Dim boolButtonExists As Boolean
    Dim btnBug


    boolButtonExists = False

    Set shMainSheet = ThisWorkbook.Sheets("MAIN")

    On Error Resume Next
    ' Try to get the name of the bug report button, which will error if it doesn't exist
    For Each btnBug In shMainSheet.OLEObjects
        If btnBug.Name = "frmBugButton" Then boolButtonExists = True: Exit For
        Err.Clear
    Next btnBug
    On Error GoTo 0

    If Not boolButtonExists Then
        ' LOG
        Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "button doesn't exist, adding to MAIN sheet": Close #1

        'create button
        Set Obj = shMainSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=220, Top:=220, Width:=275, Height:=100)
        Obj.Name = "frmBugButton"
        'button text
        With ActiveSheet.OLEObjects("frmBugButton").Object
            .Caption = ":( " & vbCrLf & vbCrLf & "Something 's Broken" & vbCrLf & "(report a bug)"
            .BackColor = &HC0&
            .ForeColor = &HFFFFFF
            .Font.Size = 14
            .Font.Bold = True
        End With
    End If

    'add execution code to sheet code module

    ' Modified from: http://www.cpearson.com/excel/vbe.aspx

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim S As String
    Dim LineNum As Long
    Dim v

    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Starting to add code to MAIN module": Close #1

    Set VBProj = ActiveWorkbook.VBProject
    For Each v In VBProj.VBComponents
        If StrComp(v.Properties("Name"), "MAIN", vbTextCompare) = 0 Then
            Set VBComp = v
            Set CodeMod = VBComp.CodeModule
            Exit For
        End If
    Next

    ' tests that the procedure doesn't yet exist
    If Not CodeMod.Find("frmBugButton_Click", 1, 1, CodeMod.CountOfLines, 255, True, True, False) Then
        ' LOG
        Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Code doesn't exist, writing it now": Close #1

        LineNum = CodeMod.CountOfLines + 1
        S = "Private Sub frmBugButton_Click()" & vbCrLf & _
          "    frmBug.Show" & vbCrLf & _
            "End Sub"
        CodeMod.InsertLines LineNum, S
    End If
    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Noting the patch was installed (intPatchesCol = " & intPatchesCol & ")": Close #1

    ' note that the patch is installed, so this doesn't keep running
    With ThisWorkbook.Sheets("Refs")
        For i = 2 To 100
            If .Cells(i, intPatchesCol).Value = vbNullString Then
                .Cells(i, intPatchesCol).Value = "v4.2.1"
                ' LOG
                Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Noting the patch was installed at row " & i & " and col " & intPatchesCol: Close #1

                Exit For
            End If
        Next i
    End With
    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Patch v4.2.1 complete": Close #1



End Sub

Private Sub patch4_2_2(intPatchesCol As Integer)
    Dim strLogFilePathAndName

    strLogFilePathAndName = ActiveWorkbook.Path & "\" & "ErrorLog.txt"
    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, " ": Close #1


    ' Check if the patch is installed first
    Dim i%, j%, intLastColOfRefs%
    Dim shtRefs As Worksheet

    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Starting patch v4.2.2": Close #1

    Set shtRefs = ThisWorkbook.Sheets("Refs")

    With shtRefs
        For i = 2 To 100
            If .Cells(i, intPatchesCol).Value = "v4.2.2" Then
                Exit Sub
            End If
        Next i
    End With

    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Done looking for patch in Refs... intLastColOfRefs = " & intLastColOfRefs & " and intPatchesCol = " & intPatchesCol: Close #1


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Patch v4.2.2                                        '
    '   Fixes calculation of FY13/FY14 bridge pay period    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim intFYColumn%
    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Starting patch v4.2.2 because it looks like it's not installed": Close #1

    ' In "Refs" sheet, find column for FY formula (should be column G)

    For j = 1 To 200
        If StrComp("FY", shtRefs.Cells(1, j).Value, vbTextCompare) = 0 Then intFYColumn = j: Exit For
    Next j

    ' Loop down the column, changing all formulas to the new one
    For i = 2 To 200
        If shtRefs.Cells(i, intFYColumn).Value = vbNullString Then Exit For
        shtRefs.Cells(i, intFYColumn).FormulaR1C1 = "=IF(RC[-3]<DATE(2012,9,1),2012,IF(RC[-3]<DATE(2013,9,1),2013,IF(RC[-3]<DATE(2014,9,1),2014,IF(RC[-3]<DATE(2015,9,1),2015,2016))))"
    Next i

    ' LOG
    Open strLogFilePathAndName For Append As #1: Print #1, Now & " " & "Noting patch v4.2.2 in Refs at column " & intPatchesCol: Close #1

    ' note that the patch is installed, so this doesn't keep running
    With shtRefs
        For i = 2 To 100
            If .Cells(i, intPatchesCol).Value = vbNullString Then
                .Cells(i, intPatchesCol).Value = "v4.2.2"
                Exit For
            End If
        Next i
    End With

End Sub









