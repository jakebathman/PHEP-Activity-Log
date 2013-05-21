Attribute VB_Name = "m_Misc_Code"
'v4.1

Option Explicit
Public Const pthUpdatedWorkbookPath = "\\ccdata01\homeland_security\PHEP Documentation\Monthly Reports\Activity Tracking\"
'Public Const pthUpdatedWorkbookPath = "C:\Users\e008922\Dropbox\_Work\Monthly reports\AO reports\PHEP drive\"

Public Const strActivityCategories = "Administrative Work;Budget or Documentation;Conference;Conference Call or Webinar;Exercise (hosted or attended);Incident Response;Inventory Management;IT Management or Maintenance;Meeting (in office);Meeting (out of office);Personnel Management;Planning or Resource Updates;Public Event or Outreach;Research or Analysis;Time Off;Training (attended);Training (conducted);Traveling;Volunteer Management"
Public arrActivityCategories(1 To 19) As String

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

Public Sub mHideSOMEOFTHETHINGS(HideRefs As Boolean, HideTemplates As Boolean)
    If HideRefs Then Sheet2.Visible = xlSheetVeryHidden    'Refs
    If HideTemplates Then
        Sheet4.Visible = xlSheetVeryHidden    'templatesheet
        'Sheets("reporttemplatesheet").Visible = xlSheetVeryHidden 'reporttemplatesheet
    End If
End Sub


Public Sub mUpdateCategories()
    Dim v, vCell
    Dim c%
    Dim wb As Workbook
    Dim shtRefs As Worksheet
    Dim tmpRange As Range
    Dim rngNamedActCategoryRange As Range

    v = Split(strActivityCategories, ";", -1, vbTextCompare)
    For c = 0 To 18    'starts at 0 because the variant array v() does so
        arrActivityCategories(c + 1) = v(c)
    Next c


    Set shtRefs = ActiveWorkbook.Sheets("Refs")
    Set rngNamedActCategoryRange = shtRefs.Range(Cells(2, 2).Address, Cells(UBound(arrActivityCategories) + 1, 2).Address)

    'If rngNamedActCategoryRange.Count = UBound(arrActivityCategories) Then Exit Sub


    'Set wb = ActiveWorkbook

    c = 1
    For Each vCell In rngNamedActCategoryRange
        vCell.Value = arrActivityCategories(c)
        c = c + 1
    Next

End Sub































Public Sub uUpdateTheUpdateCode()
    Dim arrListOfModules(), arrListOfNewModules(), arrListOfModulesAfterUpdate()
    Dim intNumModules%, intNumNewModules%, i%, j%
    Dim fVBProj As VBIDE.VBProject
    Dim tVBProj As VBIDE.VBProject
    Dim tFilePathFull$, strVers$, strVersNew$
    Dim t
    Dim c%
    Dim v
    Dim vC As VBComponent
    Dim boolTriedTwice As Boolean, boolNoMissingModules As Boolean
    Dim actApp As Application
    Dim actWB As Workbook
    Dim actWS As Worksheet
    Dim strActWBFileName$, strActWBFilePath$, strActWBFullPath$, strActWBBackupPath$, strActWBFileTitle$
    Dim strNewWBFileName$, strNewWBFilePath$, strNewWBFullPath$, strNewWbFileTitle$
    Dim strFileNameConventionInside$
    Dim intCurVersion%, dblCurVersion#
    Dim intNewVersion%, dblNewVersion#
    Dim strFileNameExtras$

    Dim boolServerFileIsDifferent As Boolean


    ' Check for macro security settings

    boolServerFileIsDifferent = True


    Set actApp = Application
    Set actWB = actApp.ActiveWorkbook
    Set actWS = actWB.ActiveSheet
    actWB.Save
    strActWBFileName = actWB.Name
    strActWBFilePath = actWB.Path
    strActWBFullPath = strActWBFilePath & "\" & strActWBFileName
    strActWBBackupPath = strActWBFilePath & "\OLD_" & strActWBFileName
    strActWBFileTitle = Replace(strActWBFileName, ".xlsm", "", Compare:=vbTextCompare)
    strFileNameConventionInside = "PHEP activity log v"

    boolTriedTwice = False

    strVers = Sheets("Refs").Range("L2").Value
    
    On Error GoTo errOtherUpdateErr

    Set tVBProj = ActiveWorkbook.VBProject

    Dim f
    'Get updated file using path
    On Error GoTo errCouldntListDir
    f = Dir(pthUpdatedWorkbookPath)
    tFilePathFull = pthUpdatedWorkbookPath & f
    On Error GoTo errOtherUpdateErr

    If f <> vbNullString And boolServerFileIsDifferent Then
        Application.StatusBar = "Server file is different! Starting update of update code..."

        Dim app As New Excel.Application
        Dim book As Excel.Workbook
        Set book = app.Workbooks.Open(tFilePathFull)
        
        t = Timer
        While Timer < t + 0.2
            DoEvents
        Wend
        
        Set fVBProj = book.VBProject

        Call uListModules(arrListOfNewModules, intNumModules, fVBProj)
        Call uListModules(arrListOfModules, intNumNewModules, tVBProj)
        c = 1
        Err.Clear


'''''''''''''''     NEW UNTESTED CODE COPIED FROM CODE UPDATE MODULE    ''''''''''''''''''''''''''

        Dim strTmpFldr$, strTmpFldrPath$
        Dim Fs As Object


        strTmpFldr = "tmpcodemodules"
        If StrComp(Right(strActWBFilePath, 1), "\", vbTextCompare) <> 0 Then strTmpFldr = "\" & strTmpFldr
        v = strActWBFilePath & strTmpFldr
        strTmpFldrPath = v
        Debug.Print v
        If Len(Dir(strActWBFilePath & strTmpFldr, vbDirectory)) <> 0 Then
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Fs.DeleteFolder v, True
        End If
        MkDir (strActWBFilePath & strTmpFldr)

        For Each vC In fVBProj.VBComponents
            Application.StatusBar = "Looking for Update Code module..."

            On Error GoTo errOtherUpdateErr
            If vC.Name = "u_Update_Code" Or vC.Name = "frmWorking" Then
                v = ExportVBComponent(vC, strActWBFilePath & strTmpFldr, , True)
                If v <> True Then Call MsgBox("Problem with " & vC.Name & " export :(")
            t = Timer
            While Timer < t + 0.1
                DoEvents
            Wend
            End If
            
            c = c + 1
            On Error Resume Next
            Debug.Print Len(tVBProj.VBComponents("frmWorking").Name)
            If Err.Number = 0 Then Call UpdateProgressBar(" ", (c / fVBProj.VBComponents.Count) * 100): Err.Clear
            On Error GoTo errOtherUpdateErr
        Next


        Dim m As VBComponent
        'delete existing u_Update_Code and frmWorking
        For Each m In actWB.VBProject.VBComponents
            Debug.Print actWB.VBProject.VBComponents.Count
            If m.Name = "frmworking" Then  'or m.Name <> "u_Update_Code" Then
                Application.StatusBar = "Deleting " & m.Name
                actWB.VBProject.VBComponents.Remove m
                On Error Resume Next
                    actWB.VBProject.VBComponents.Item("u_Update_Code1").Activate
                    Debug.Print actWB.VBProject.VBComponents.Item("u_Update_Code1").Name
                    If Err.Number <> 0 Then actWB.VBProject.VBComponents.Remove m
                On Error GoTo errOtherUpdateErr
            End If
        Next

        c = 0

        'now do the importing
        'modified version of code found on: http://stackoverflow.com/questions/10380312/loop-through-files-in-a-folder-using-vba

        Dim MyObj As Object, MySource As Object, file As Variant
        Dim sText As String
        Dim vArray
        Dim lCnt As Long
        Dim tmpComp As VBComponent
        Dim tmpCodeMod As CodeModule
        Dim tmpModName$
        Dim S
        Dim TempVBComp As VBComponent
        Dim filepth$
        Dim newfilepth$
        Dim ModName$

        file = Dir(strActWBFilePath & strTmpFldr & "\")
        While (file <> "")
            filepth = strActWBFilePath & strTmpFldr & "\" & file
            Debug.Print c & ": " & file
            On Error Resume Next
            If InStr(1, file, "u_Update_Code", vbTextCompare) > 0 Then
                Name filepth As strActWBFilePath & strTmpFldr & "\tmp" & file
                newfilepth = strActWBFilePath & strTmpFldr & "\tmp" & file
                tmpModName = "tmp" & Left(file, Len(file) - 4)
                ModName = Left(file, Len(file) - 4)
                Set TempVBComp = actWB.VBProject.VBComponents.Import(newfilepth)
                TempVBComp.Name = tmpModName
                Set tmpComp = actWB.VBProject.VBComponents(ModName)
                Set tmpCodeMod = tmpComp.CodeModule


                With tmpCodeMod
                    .DeleteLines 1, .CountOfLines
                    S = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                    .InsertLines 1, S
                End With
                actWB.VBProject.VBComponents.Remove TempVBComp

            ElseIf InStr(1, file, ".frx", vbTextCompare) = 0 Then
                Application.StatusBar = "Importing module " & file
                actWB.VBProject.VBComponents.Import (strActWBFilePath & strTmpFldr & "\" & file)
            End If
            If InStr(1, file, ".frx", vbTextCompare) > 0 Then intNumNewModules = intNumNewModules - 1
            On Error GoTo errOtherUpdateErr
            'Call UpdateProgressBar(" ", (c / intNumNewModules) * 100):
            t = Timer
            While Timer < t + 0.05
                DoEvents
            Wend
            c = c + 1
            file = Dir
        Wend

        v = book.Sheets("Refs").Range("L2").Value

        
        book.Close SaveChanges:=False
        Set book = Nothing
        app.Quit
        Set app = Nothing


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '   Delete the tmpcodemodules folder, now that we're done with it
        Debug.Print "Starting attempt to delete temp folder"
        If Len(Dir(strActWBFilePath & strTmpFldr, vbDirectory)) <> 0 Then
            Debug.Print "Looks like it exists!"
            Debug.Print "Setting Fs to an object"
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Debug.Print "Deleting Fs, which is at " & strTmpFldrPath
            On Error Resume Next
            Fs.DeleteFolder strTmpFldrPath, True
            On Error GoTo errOtherUpdateErr
            Debug.Print "Finished deleting Fs"
        End If


        'Pesky class modules shouldn't hang around
        For Each vC In tVBProj.VBComponents
            If vC.Type = vbext_ct_ClassModule Then tVBProj.VBComponents.Remove vC
        Next






        'v = TotalCodeLinesInVBComponent(tVBProj.VBComponents("v_Version_Num")) - 3
        'Debug.Print v
        strVersNew = CStr(v)
        Sheets("Refs").Range("R1").Value = "UpdateCodeVersion"
        Sheets("Refs").Range("R2").Value = strVersNew
        Sheets("Refs").Range("Q2").Value = "TRUE"

        'Call MsgBox("Update to the updating code complete!!" & vbNewLine & vbNewLine _
         & "This is Version " & strVersNew & " of this tool." & vbNewLine & vbNewLine _
         & " ")
        Application.StatusBar = "Update code had to be completed (seriously). It's done!"

    Else
        'Call MsgBox("Looks like you've got the latest version!" & vbNewLine & vbNewLine _
                  & "This is Version " & strVers & vbNewLine & vbNewLine _
                  & "It's possible you're not able to access the PHEP drive, which may result in this message.")
    End If

    Exit Sub

errCouldntListDir:
    If Not boolTriedTwice Then
        If MsgBox("Looks like something went wront trying to access the updated code. You may not be able to connect to the PHEP drive." & vbNewLine & vbNewLine _
                & "Try Again?", vbYesNo, "I can't connect! :(") = vbYes Then
            Resume
        Else
            Exit Sub
        End If
    Else
        Call MsgBox("I've tried again and failed. You probably can't connect to the PHEP drive." & vbNewLine & vbNewLine _
                  & "Go get Jake, he'll know what to do...", vbOK, ":(")
    End If


errOtherUpdateErr:
    On Error Resume Next
    book.Close SaveChanges:=False
    Set book = Nothing
    app.Quit
    Set app = Nothing
    On Error GoTo 0

    MsgBox ("Sorry! Something went wrong :(" & vbNewLine & vbNewLine & "The code was NOT updated." _
          & vbNewLine & vbNewLine & "Error #: " & Err.Number & vbNewLine & "Error text: " & Err.Description)


End Sub


Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, _
                                  FolderName As String, _
                                  Optional FileName As String, _
                                  Optional OverwriteExisting As Boolean = True) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function exports the code module of a VBComponent to a text
' file. If FileName is missing, the code will be exported to
' a file with the same name as the VBComponent followed by the
' appropriate extension.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim FName As String
    'Extension = ".txt"
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(FileName) = vbNullString Then
        FName = VBComp.Name & Extension
    Else
        FName = FileName
        If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
            FName = FName & Extension
        End If
    End If

    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        FName = FolderName & FName
    Else
        FName = FolderName & "\" & FName
    End If

    If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill FName
        Else
            ExportVBComponent = False
            Exit Function
        End If
    End If

    VBComp.Export FileName:=FName
    ExportVBComponent = True

End Function

Public Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select

End Function


























' Code below copied from http://www.cpearson.com/excel/vbe.aspx
Function CopyModule(ByVal iItemNum, _
                    FromVBProject As VBIDE.VBProject, _
                    ToVBProject As VBIDE.VBProject, _
                    OverwriteExisting As Boolean, strPathToWB As String) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyModule
' This function copies a module from one VBProject to
' another. It returns True if successful or False
' if an error occurs.
'
' Parameters:
' --------------------------------
' FromVBProject         The VBProject that contains the module
'                       to be copied.
'
' ToVBProject           The VBProject into which the module is
'                       to be copied.
'
' CodeModuleName            The name of the module to copy.
'
' OverwriteExisting     If True, the VBComponent named CodeModuleName
'                       in ToVBProject will be removed before
'                       importing the module. If False and
'                       a VBComponent named CodeModuleName exists
'                       in ToVBProject, the code will return
'                       False.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim VBComp As VBIDE.VBComponent
    Dim FName As String
    Dim CompName As String
    Dim S As String
    Dim SlashPos As Long
    Dim ExtPos As Long
    Dim TempVBComp As VBIDE.VBComponent

    '''''''''''''''''''''''''''''''''''''''''''''
    ' Do some housekeeping validation.
    '''''''''''''''''''''''''''''''''''''''''''''
    If FromVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If

    If Trim(FromVBProject.VBComponents.Item(iItemNum).Name) = vbNullString Then
        CopyModule = False
        Exit Function
    End If

    If ToVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If

    If FromVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If

    If ToVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If

    On Error Resume Next
    Set VBComp = FromVBProject.VBComponents.Item(iItemNum)
    If Err.Number <> 0 Then
        CopyModule = False
        Exit Function
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FName is the name of the temporary file to be
    ' used in the Export/Import code.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Debug.Print FName
    FName = strPathToWB & "\" & FromVBProject.VBComponents.Item(iItemNum).Name & ".bas"
    'FName = Environ("Temp") & "\vbComps\" & FromVBProject.VBComponents.Item(iItemNum).Name & ".bas"
    If OverwriteExisting = True Then
        ''''''''''''''''''''''''''''''''''''''
        ' If OverwriteExisting is True, Kill
        ' the existing temp file and remove
        ' the existing VBComponent from the
        ' ToVBProject.
        ''''''''''''''''''''''''''''''''''''''
        If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            Err.Clear
            Kill FName
            If Err.Number <> 0 Then
                CopyModule = False
                Exit Function
            End If
        End If
        ToVBProject.VBComponents.Remove ToVBProject.VBComponents(FromVBProject.VBComponents.Item(iItemNum).Name)
    Else
        '''''''''''''''''''''''''''''''''''''''''
        ' OverwriteExisting is False. If there is
        ' already a VBComponent named CodeModuleName,
        ' exit with a return code of False.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Set VBComp = ToVBProject.VBComponents(FromVBProject.VBComponents.Item(iItemNum).Name)
        If Err.Number <> 0 Then
            If Err.Number = 9 Then
                ' module doesn't exist. ignore error.
            Else
                ' other error. get out with return value of False
                CopyModule = False
                Exit Function
            End If
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Do the Export and Import operation using FName
    ' and then Kill FName.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FromVBProject.VBComponents.Item(iItemNum).Export FileName:=FName

    '''''''''''''''''''''''''''''''''''''
    ' Extract the module name from the
    ' export file name.
    '''''''''''''''''''''''''''''''''''''
    SlashPos = InStrRev(FName, "\")
    ExtPos = InStrRev(FName, ".")
    CompName = Mid(FName, SlashPos + 1, ExtPos - SlashPos - 1)

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Document modules (SheetX and ThisWorkbook)
    ' cannot be removed. So, if we are working with
    ' a document object, delete all code in that
    ' component and add the lines of FName
    ' back in to the module.
    ''''''''''''''''''''''''''''''''''''''''''''''
    Set VBComp = Nothing
    Set VBComp = ToVBProject.VBComponents(CompName)

    If VBComp Is Nothing Then
        ToVBProject.VBComponents.Import FileName:=FName
    Else
        If VBComp.Type = vbext_ct_Document Then
            ' VBComp is destination module
            Set TempVBComp = ToVBProject.VBComponents.Import(FName)
            ' TempVBComp is source module
            With VBComp.CodeModule
                .DeleteLines 1, .CountOfLines
                S = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                .InsertLines 1, S
            End With
            On Error GoTo 0
            ToVBProject.VBComponents.Remove TempVBComp
        End If
    End If
    Kill FName
    CopyModule = True
End Function










Public Function TotalCodeLinesInVBComponent(VBComp As VBIDE.VBComponent) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the total number of code lines (excluding blank lines and
' comment lines) in the VBComponent referenced by VBComp. Returns -1
' if the VBProject is locked.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim S As String
    Dim LineCount As Long

    If VBComp.Collection.Parent.Protection = vbext_pp_locked Then
        TotalCodeLinesInVBComponent = -1
        Exit Function
    End If

    With VBComp.CodeModule
        For N = 1 To .CountOfLines
            S = .Lines(N, 1)
            If Trim(S) = vbNullString Then
                ' blank line, skip it
            ElseIf Left(Trim(S), 1) = "'" Then
                ' comment line, skip it
            Else
                LineCount = LineCount + 1
            End If
        Next N
    End With
    TotalCodeLinesInVBComponent = LineCount
End Function



'****************************
'*                          *
'*   Count lines of code    *
'*      Just for fun        *
'*                          *
'****************************

Function CountCodeLines()
    Dim VBCodeModule As Object
    Dim NumLines As Long, N As Long
    With ActiveWorkbook
        For N = 1 To .VBProject.VBComponents.Count
            Set VBCodeModule = .VBProject.VBComponents(N).CodeModule
            NumLines = NumLines + VBCodeModule.CountOfLines
        Next
    End With
    NumLines = NumLines - 13    ' exclude this module from the count
    'MsgBox "Total number of lines of code in the project = " & NumLines, , "Code Lines"
    Set VBCodeModule = Nothing
    CountCodeLines = NumLines
    Debug.Print "All project modules contain " & CountCodeLines & " lines of code"
End Function


'****************************************************
'*                                                  *
'*              List all open workbooks             *
'*                                                  *
'****************************************************



Sub ListOpenBooks()
'lists each book that's OPEN
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        Debug.Print "Open workbook: " & wb.Name
    Next wb
End Sub






'---------------------------------------------------------------------------------------
' Procedure : ReviseVersionNumberComment
' Author    : E008922
' Date      : 10/16/2012
' Purpose   : Revise the version number comment at the top of every code module
'---------------------------------------------------------------------------------------
'
Public Sub ReviseVersionNumberComment() 'Optional sOld, Optional rNew)

    On Error GoTo ReviseVersionNumberComment_Error
    Dim v, x, y
    Dim S$, r$

    'If sOld = vbNullString Then s = "'v3"
    'If rNew = vbNullString Then r = "'v4.1"
    S = "'v4"
    r = "'v4.1"

    For Each v In ThisWorkbook.VBProject.VBComponents
        x = v.CodeModule.Find(S, 1, 1, 2, 5, False, True, False)
        If x Then
            Call v.CodeModule.ReplaceLine(1, r)
        End If
    Next

    On Error GoTo 0
    Exit Sub

ReviseVersionNumberComment_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReviseVersionNumberComment of Module m_Misc_Code"
End Sub






