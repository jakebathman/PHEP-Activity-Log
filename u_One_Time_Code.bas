Attribute VB_Name = "u_One_Time_Code"
Option Explicit

Public Sub uOneTimeCode()

Call patch4_2_1
Call patch4_2_2


End Sub


Public Sub patch4_2_1()


    ' Check if the patch is installed first
    Dim i%, j%, intLastColOfRefs%, intPatchesCol%
    With ThisWorkbook.Sheets("Refs")
        ' Find the correct column
        For j = 1 To 200
            If StrComp("PatchesInstalled", .Cells(1, j).Value, vbTextCompare) = 0 Then intPatchesCol = j: Exit For
            If .Cells(1, j).Value = vbNullString And .Cells(1, j + 1).Value = vbNullString Then intLastColOfRefs = j - 1: Exit For
        Next j
        If intLastColOfRefs > 0 Then
            .Cells(1, intLastColOfRefs + 1).Value = "PatchesInstalled"
            intPatchesCol = intLastColOfRefs
        End If
        For i = 2 To 100
            If .Cells(i, intPatchesCol).Value = "v4.2.1" Then
                Exit Sub
            End If
        Next i
    End With


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Patch v4.2.1                                        '
'   Adds Bug Report button on MAIN sheet                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Obj As Object
    Dim Code As String
    Dim shMainSheet As Worksheet

    Set shMainSheet = ThisWorkbook.Sheets("MAIN")

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


    'add execution code to sheet code module

    ' Modified from: http://www.cpearson.com/excel/vbe.aspx

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim S As String
    Dim LineNum As Long
    Dim v

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

        LineNum = CodeMod.CountOfLines + 1
        S = "Private Sub frmBugButton_Click()" & vbCrLf & _
            "    frmBug.Show" & vbCrLf & _
            "End Sub"
        CodeMod.InsertLines LineNum, S
    End If

    ' note that the patch is installed, so this doesn't keep running
    With ThisWorkbook.Sheets("Refs")
        For i = 2 To 100
            If .Cells(i, intPatchesCol).Value = vbNullString Then
                .Cells(i, intPatchesCol).Value = "v4.2.1"
                Exit For
            End If
        Next i
    End With

End Sub

Public Sub patch4_2_2()


    ' Check if the patch is installed first
    Dim i%, j%, intLastColOfRefs%, intPatchesCol%
    Dim shtRefs As Worksheet

    Set shtRefs = ThisWorkbook.Sheets("Refs")

    With shtRefs
        ' Find the correct column
        For j = 1 To 200
            If StrComp("PatchesInstalled", .Cells(1, j).Value, vbTextCompare) = 0 Then intPatchesCol = j: Exit For
            If .Cells(1, j).Value = vbNullString And .Cells(1, j + 1).Value = vbNullString Then intLastColOfRefs = j - 1: Exit For
        Next j
        If intLastColOfRefs > 0 Then
            .Cells(1, intLastColOfRefs + 1).Value = "PatchesInstalled"
            intPatchesCol = intLastColOfRefs
        End If
        For i = 2 To 100
            If .Cells(i, intPatchesCol).Value = "v4.2.2" Then
                Exit Sub
            End If
        Next i
    End With


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Patch v4.2.2                                        '
    '   Fixes calculation of FY13/FY14 bridge pay period    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim intFYColumn%

    ' In "Refs" sheet, find column for FY formula (should be column G)

    For j = 1 To 200
        If StrComp("FY", shtRefs.Cells(1, j).Value, vbTextCompare) = 0 Then intFYColumn = j: Exit For
    Next j

    ' Loop down the column, changing all formulas to the new one
    For i = 2 To 200
        If shtRefs.Cells(i, intFYColumn).Value = vbNullString Then Exit For
        shtRefs.Cells(i, intFYColumn).FormulaR1C1 = "=IF(RC[-2]<=DATE(2012,9,1),2012,IF(RC[-2]<=DATE(2013,9,1),2013,IF(RC[-2]<=DATE(2014,9,1),2014,IF(RC[-2]<=DATE(2015,9,1),2015,2016))))"
    Next i


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



