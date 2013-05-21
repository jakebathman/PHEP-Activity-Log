Attribute VB_Name = "u_One_Time_Code"
Option Explicit

Public Sub uOneTimeCode()


    ' Patch to add Bug Report button on MAIN sheet
    ' Patch v4.2.1


    ' Check if the patch is installed first
    Dim i%, j%
    With ThisWorkbook.Sheets("Refs")
    ' Find the correct column
    For j = 1 To 200
        If StrComp("PatchesInstalled", .Cells(1, j).Value, vbTextCompare) = 0 Then Exit For
    Next j
        For i = 2 To 100
            If .Cells(i, j).Value = "v4.2.1" Then
                Exit Sub
            End If
        Next i
    End With


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
            If .Cells(i, j).Value = vbNullString Then
                .Cells(i, j).Value = "v4.2.1"
                Exit For
            End If
        Next i
    End With



End Sub



