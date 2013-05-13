Attribute VB_Name = "u_Update_Code"
'v4

Option Explicit
Public Const pthUpdatedWorkbookPath = "\\ccdata01\homeland_security\PHEP Documentation\Monthly Reports\Activity Tracking\"
'Public Const pthUpdatedWorkbookPath = "C:\Users\e008922\Dropbox\_Work\Monthly reports\AO reports\PHEP drive\"


Public Sub uUpdateCode()
    Dim arrListOfModules(), arrListOfNewModules(), arrListOfModulesAfterUpdate()
    Dim intNumModules%, intNumNewModules%, i%, j%
    Dim fVBProj As VBIDE.VBProject
    Dim tVBProj As VBIDE.VBProject
    Dim tFilePathFull$, strVers$, strVersNew$
    Dim t
    Dim c%
    Dim v, vC
    Dim boolTriedTwice As Boolean, boolNoMissingModules As Boolean
    Dim actApp As Application
    Dim actWB As Workbook
    Dim actWS As Worksheet
    Dim strActWBFileName$, strActWBFilePath$, strActWBFullPath$, strActWBBackupPath$, strActWBFileTitle$
    Dim strNewWBFileName$, strNewWBFilePath$, strNewWBFullPath$, strNewWbFileTitle$
    Dim strFileNameConventionInside$
    Dim intCurVersion%
    Dim intNewVersion%
    Dim strFileNameExtras$
    
    Dim boolServerFileIsDifferent As Boolean
    
    
' Check for macro security settings
If Not AddRefsIfAccessAllowed Then Exit Sub
    
    boolServerFileIsDifferent = False


    Set actApp = Application
    Set actWB = actApp.ActiveWorkbook
    Set actWS = actWB.ActiveSheet
    strActWBFileName = actWB.Name
    strActWBFilePath = actWB.Path
        strActWBFullPath = strActWBFilePath & "\" & strActWBFileName
        strActWBBackupPath = strActWBFilePath & "\OLD_" & strActWBFileName
        strActWBFileTitle = Replace(strActWBFileName, ".xlsm", "", Compare:=vbTextCompare)
    actWB.Save
    strFileNameConventionInside = "PHEP activity log v"
    
    boolTriedTwice = False
    
    strVers = Sheets("Refs").Range("L2").Value
    
    On Error Resume Next
        Err.Clear
        Debug.Print Len(ThisWorkbook.VBProject.VBComponents("frmWorking").Name)
        Debug.Print Err.Number
        Debug.Print Err.Description
        If Err.Number = 0 Then Call InitializeProgressBar
    On Error GoTo errOtherUpdateErr
    
    
    Set tVBProj = ActiveWorkbook.VBProject
    
        Dim f
        'Get updated file using path
        On Error GoTo errCouldntListDir
        f = Dir(pthUpdatedWorkbookPath)
        tFilePathFull = pthUpdatedWorkbookPath & f

'        Do While f <> ""
'            Debug.Print f
'            Debug.Print FileLen(pthUpdatedWorkbookPath & f)
'            Debug.Print FileDateTime(pthUpdatedWorkbookPath & f)
'            Get next File
'            f = Dir()
'        Loop
    On Error GoTo errOtherUpdateErr
    '   This IF statement used to check for a file name dissimilar in any way, very uninteligently
    '   If f <> vbNullString And StrComp(f, ActiveWorkbook.Name, vbTextCompare) <> 0 Then
    
    Debug.Print ActiveWorkbook.Name
    Debug.Print f
    Debug.Print strActWBFileName
        strFileNameExtras = Mid(strActWBFileName, 1, InStr(1, strActWBFileName, strFileNameConventionInside, vbTextCompare) - 1)
    If Len(strFileNameExtras & f) <> Len(strActWBFileName) And InStr(1, strActWBFileName, strFileNameConventionInside, vbTextCompare) >= 1 Then
        boolServerFileIsDifferent = True
    Else
        If StrComp(strFileNameExtras & f, strActWBFileName, vbTextCompare) <> 0 Then boolServerFileIsDifferent = True
    End If
    
    
    
    Debug.Print strFileNameExtras
    
    If f <> vbNullString And boolServerFileIsDifferent Then
        Application.DisplayAlerts = False
        actWB.SaveAs FileName:=strActWBBackupPath, FileFormat:=XlFileFormat.xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
        Application.StatusBar = "Saved a backup..."
        
        Dim app As New Excel.Application
        Dim book As Excel.Workbook
        Set book = app.Workbooks.Open(tFilePathFull)
        strNewWBFileName = strFileNameExtras & f
        strNewWBFilePath = book.Path
            strNewWBFullPath = strActWBFilePath & "\" & strNewWBFileName
        Debug.Print "New full path: " & strNewWBFullPath
        strNewWbFileTitle = Replace(strNewWBFileName, ".xlsm", "", Compare:=vbTextCompare)
        Debug.Print Mid(strActWBFileTitle, InStr(1, strFileNameConventionInside, strActWBFileTitle, vbTextCompare) + Len(strFileNameConventionInside))
        intCurVersion = CInt(Mid(strActWBFileTitle, Len(strActWBFileTitle), 1))
        intNewVersion = CInt(Mid(strNewWbFileTitle, Len(strNewWbFileTitle), 1))

        Set fVBProj = book.VBProject
    
    Application.DisplayAlerts = False
        actWB.SaveAs FileName:=strNewWBFullPath, FileFormat:=XlFileFormat.xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    Application.StatusBar = "Saved with updated file name..."
        
        Call uListModules(arrListOfNewModules, intNumModules, fVBProj)
        Call uListModules(arrListOfModules, intNumNewModules, tVBProj)
        c = 1
        Err.Clear
        For Each vC In fVBProj.VBComponents
            Application.StatusBar = "Updating module " & c & " of " & fVBProj.VBComponents.Count
            On Error Resume Next
                Err.Clear
                Debug.Print Len(tVBProj.VBComponents("frmWorking").Name)
                If Err.Number = 0 Then Call UpdateProgressBar("Module " & c & " of " & fVBProj.VBComponents.Count, (c / fVBProj.VBComponents.Count) * 100): Err.Clear
            On Error GoTo errOtherUpdateErr
            If vC.Name <> "u_Update_Code" And vC.Name <> "u_List_Modules" Then
            v = CopyModule(c, fVBProj, tVBProj, True, actWB.Path)

            End If
            t = Timer
            While Timer < t + 0.1
                actWB.Activate
                DoEvents
            Wend
            c = c + 1
            On Error Resume Next
                Debug.Print Len(tVBProj.VBComponents("frmWorking").Name)
                If Err.Number = 0 Then Call UpdateProgressBar(" ", (c / fVBProj.VBComponents.Count) * 100): Err.Clear
            On Error GoTo errOtherUpdateErr
        Next
        
        
        book.Close SaveChanges:=False
        Set book = Nothing
        app.Quit
        Set app = Nothing
        
        Application.StatusBar = "Done updating!"
        
        v = TotalCodeLinesInVBComponent(tVBProj.VBComponents("v_Version_Num")) - 3
        'Debug.Print v
        strVersNew = CStr(v)
        Sheets("Refs").Range("L2").Value = strVersNew
        Sheets("Refs").Range("Q2").Value = "FALSE"


' Check to make sure none of the modules are missing, which indicates something went wrong
        Application.StatusBar = "Checking that the update didn't break anything..."
Call uListModules(arrListOfModulesAfterUpdate, tVBProj.VBComponents.Count, tVBProj)
If fCompareArrays(arrListOfModules, arrListOfModulesAfterUpdate) = False Then boolNoMissingModules = False Else boolNoMissingModules = True

        If boolNoMissingModules Then
        Application.StatusBar = "Killing old file..."
        Kill strActWBFullPath
            ThisWorkbook.Save
            Application.StatusBar = "Update complete!!"
            Call MsgBox("Update complete!!" & vbNewLine & vbNewLine _
                    & "This is Version " & strVersNew & " of this tool." & vbNewLine & vbNewLine _
                    & "(Your old workbook was backed up in the same folder, just in case)")
        Else
            GoTo errOtherUpdateErr
        End If
        
    Else
        Call MsgBox("Looks like you've got the latest version!" & vbNewLine & vbNewLine _
                & "This is Version " & strVers & vbNewLine & vbNewLine _
                & "It's possible you're not able to access the PHEP drive, which may result in this message.")
        Call UnloadAllForms
    End If


Call MaintenanceForAddActivityButton
Call UnloadAllForms
Application.StatusBar = False
Exit Sub

errCouldntListDir:
    If Not boolTriedTwice Then
        If MsgBox("Looks like something went wront trying to access the updated code. You may not be able to connect to the PHEP drive." & vbNewLine & vbNewLine _
                & "Try Again?", vbYesNo, "I can't connect! :(") = vbYes Then
            Resume
        Else
            Call UnloadAllForms
            Application.StatusBar = False
            Exit Sub
        End If
    Else
        Call MsgBox("I've tried again and failed. You probably can't connect to the PHEP drive." & vbNewLine & vbNewLine _
                & "Go get Jake, he'll know what to do...", vbOK, ":(")
        Call UnloadAllForms
        Application.StatusBar = False
    End If

    
errOtherUpdateErr:
    On Error Resume Next
    book.Close SaveChanges:=False
    Set book = Nothing
    app.Quit
    Set app = Nothing
    On Error GoTo 0

    Call MsgBox("Sorry! Something went wrong :(" & vbNewLine & vbNewLine & "The code was NOT updated." _
            & vbNewLine & vbNewLine & "Error #: " & Err.Number & vbNewLine & "Error text: " & Err.Description)
            
    Call MsgBox("Hey, this is important." & vbNewLine & vbNewLine & "THIS WORKBOOK IS NOW BROKEN!" _
            & vbNewLine & vbNewLine & "The workbook was saved, so you won't lose any data." _
            & vbNewLine & vbNewLine & "***************************************************" _
            & vbNewLine & "THIS WORKBOOK WILL NOW QUIT. PLEASE RE-OPEN IT, AND TRY TO UPDATE AGAIN." _
            & vbNewLine & "***************************************************", vbOKOnly + vbCritical, "WORKBOOK WILL NOW CLOSE ... NO DATA WILL BE LOST!")
            
            Call UnloadAllForms
            Application.StatusBar = False
            ThisWorkbook.Close SaveChanges:=False
            


End Sub



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
    Dim s As String
    Dim SlashPos As Long
    Dim ExtPos As Long
    Dim TempVBComp As VBIDE.VBComponent
    Dim strExt$
    Dim t
    
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
    
Call CheckAndUpdateProgressBar

    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FName is the name of the temporary file to be
    ' used in the Export/Import code.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case FromVBProject.VBComponents.Item(iItemNum).Type
        Case vbext_ct_Document
            strExt = ".cls"
        Case vbext_ct_MSForm
            strExt = ".frm"
        Case vbext_ct_StdModule
            strExt = ".bas"
        Case Else
            strExt = ".bas"
    End Select
    FName = strPathToWB & "\" & FromVBProject.VBComponents.Item(iItemNum).Name & strExt
    Debug.Print FName
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
Call CheckAndUpdateProgressBar
t = Timer
While Timer < t + 0.1
    DoEvents
Wend
Call CheckAndUpdateProgressBar
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
Call CheckAndUpdateProgressBar
t = Timer
While Timer < t + 0.1
    DoEvents
Wend
Call CheckAndUpdateProgressBar
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
Call CheckAndUpdateProgressBar
t = Timer
While Timer < t + 0.1
    DoEvents
Wend
Call CheckAndUpdateProgressBar
    Else
        If VBComp.Type = vbext_ct_Document Then
            ' VBComp is destination module
            Set TempVBComp = ToVBProject.VBComponents.Import(FName)
            ' TempVBComp is source module
            With VBComp.CodeModule
                .DeleteLines 1, .CountOfLines
                s = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                .InsertLines 1, s
            End With
            On Error GoTo 0
            ToVBProject.VBComponents.Remove TempVBComp
        End If
    End If
    Kill FName
    If FromVBProject.VBComponents.Item(iItemNum).Type = vbext_ct_MSForm Then FName = Replace(FName, ".frm", ".frx"): Kill FName
    CopyModule = True
End Function



Public Sub CountTheLines()
    Dim N As Long
    Dim s As String
    Dim LineCount As Long
    Dim v
    
    If ThisWorkbook.VBProject.Protection = vbext_pp_locked Then
        LineCount = -1
        Exit Sub
    End If
    
    For Each v In ThisWorkbook.VBProject.VBComponents
        With v.CodeModule
            For N = 1 To .CountOfLines
                s = .Lines(N, 1)
                If Trim(s) = vbNullString Then
                    ' blank line, skip it
                ElseIf Left(Trim(s), 1) = "'" Then
                    ' comment line, skip it
                Else
                    LineCount = LineCount + 1
                End If
            Next N
        End With
    Next
    MsgBox ("There are " & LineCount & " lines in this project.")
End Sub






Public Function TotalCodeLinesInVBComponent(VBComp As VBIDE.VBComponent) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns the total number of code lines (excluding blank lines and
    ' comment lines) in the VBComponent referenced by VBComp. Returns -1
    ' if the VBProject is locked.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim N As Long
        Dim s As String
        Dim LineCount As Long
        
        If VBComp.Collection.Parent.Protection = vbext_pp_locked Then
            TotalCodeLinesInVBComponent = -1
            Exit Function
        End If
        
        With VBComp.CodeModule
            For N = 1 To .CountOfLines
                s = .Lines(N, 1)
                If Trim(s) = vbNullString Then
                    ' blank line, skip it
                ElseIf Left(Trim(s), 1) = "'" Then
                    ' comment line, skip it
                Else
                    LineCount = LineCount + 1
                End If
            Next N
        End With
        TotalCodeLinesInVBComponent = LineCount
    End Function


















Public Function fCompareArrays(ByRef arrOld, ByRef arrNew) As Boolean
    ' Takes two arrays as input, spits out boolean TRUE if all elements of arrOld are found in arrNew
    ' Returning TRUE indicates no data was lost during an external operation on the items in the array
    
Dim i%, c%, intDiff%
Dim boolTmp As Boolean
Dim v, x
Dim aO(), aN()

ReDim aO(1 To UBound(arrOld))
ReDim aN(1 To UBound(arrNew))


For i = 1 To UBound(arrOld)
    aO(i) = arrOld(i, 1)
Next i
For i = 1 To UBound(arrNew)
    aN(i) = arrNew(i, 1)
Next i

For Each v In aO
    boolTmp = False
    For Each x In aN
        Debug.Print "Checking " & v & " against " & x
        If StrComp(v, x, vbTextCompare) = 0 Then boolTmp = True: Exit For
    Next
    If Not boolTmp Then intDiff = intDiff + 1: Debug.Print "intDiff Changed...now it's " & intDiff
Next

If intDiff > 0 Then fCompareArrays = False Else fCompareArrays = True

Debug.Print "Function done...returned " & fCompareArrays
    
End Function








Public Function CheckAndUpdateProgressBar()
    On Error Resume Next
        Debug.Print Len(ThisWorkbook.VBProject.VBComponents("frmWorking").Name)
        If Err.Number = 0 Then Call UpdateProgressBar
    On Error GoTo 0
End Function














Public Sub UpdateProgressBar(Optional strCaption As String, Optional dblPctTitle As Double)
    Dim arrRotatingChar(1 To 4)
    Dim iBarLeft#, iBarWidth#, iBarRight#
    Dim iBGLeft#, iBGWidth#, iBGRight#
    Dim iBarTwoWidth#, iBarTwoRight#
    Dim NewBarRight#
    Dim steps#
    
    arrRotatingChar(1) = "|"
    arrRotatingChar(2) = " | "
    arrRotatingChar(3) = "/"
    arrRotatingChar(4) = " / "
    
    Select Case frmWorking.lblProgressText.Caption
        Case "|"
            frmWorking.lblProgressText.Caption = arrRotatingChar(2)
        Case "/"
            frmWorking.lblProgressText.Caption = arrRotatingChar(3)
        Case "--"
            frmWorking.lblProgressText.Caption = arrRotatingChar(4)
        Case "\"
            frmWorking.lblProgressText.Caption = arrRotatingChar(1)
    End Select
    
    iBarLeft = frmWorking.lblMovingBar.Left
    iBarWidth = frmWorking.lblMovingBar.Width
    iBarRight = iBarLeft + iBarWidth
    
    iBGWidth = frmWorking.Label3.Width
    iBGLeft = frmWorking.Label3.Left
    iBGRight = iBGWidth + iBGLeft
    
    iBarTwoWidth = frmWorking.lblMoving2.Width
    iBarTwoRight = iBarTwoWidth + 10
    
    steps = Round((iBGWidth / 47), 0)
    
    If Round(iBarRight + steps + 1, 0) > iBGRight Then
        If Round(iBarLeft + steps + 1, 0) > iBGRight Then 'reset bar to the left
            If iBarTwoWidth > 0 Then
                frmWorking.lblMoving2.Width = 0
                frmWorking.lblMovingBar.Left = steps + 10
                frmWorking.lblMovingBar.Width = 85
            Else
                frmWorking.lblMovingBar.Left = 10
                frmWorking.lblMovingBar.Width = 85
                frmWorking.lblMoving2.Width = 0
            End If
        Else
            frmWorking.lblMovingBar.Left = iBarLeft + steps
            frmWorking.lblMovingBar.Width = iBGRight - (iBarLeft + steps) - 2
            NewBarRight = frmWorking.lblMovingBar.Left + 85 'measures new width of green bar if spills over
            frmWorking.lblMoving2.Width = (NewBarRight - iBGRight)
        End If
    Else
        frmWorking.lblMovingBar.Left = iBarLeft + steps
    End If
    
    If dblPctTitle > 0 Then frmWorking.Caption = Round(dblPctTitle, 1) & "% Complete"
    If dblPctTitle = 200 Then
        frmWorking.Caption = "100% Done!"
        frmWorking.lblMovingBar.Left = 10
        frmWorking.lblMovingBar.Width = frmWorking.Label3.Width - 4
    End If
    
    If strCaption <> vbNullString Then
        frmWorking.lblModuleUpdateText.Caption = strCaption
    End If
    
    frmWorking.Repaint

End Sub


Public Sub InitializeProgressBar()

    With frmWorking
        .Show False
        .Height = 60
        .Width = 443
        .Top = Application.Top + (Application.Height / 2) - (.Height / 2) - 75
        .Left = Application.Left + (Application.Width / 2) - (.Width / 2)
        '.Label2.Caption = vbNullString
        .lblModuleUpdateText.Caption = vbNullString
    End With

End Sub





Public Sub UnloadAllForms()

    Dim objLoop As Object
    On Error Resume Next
    For Each objLoop In VBA.UserForms
        If TypeOf objLoop Is UserForm Then Unload objLoop
    Next objLoop

End Sub







Public Function AddRefsIfAccessAllowed()

Dim Response As VbMsgBoxResult, v

    'Test to ensure access is allowed
     If Application.Version > 9 Then
           Dim VisualBasicProject As Object
           On Error Resume Next
           Set VisualBasicProject = ActiveWorkbook.VBProject
           If Not Err.Number = 0 Then
                'For Each v In Application.CommandBars
                '    Debug.Print v.Name
                'Next
                Response = MsgBox(vbNewLine & "Your current security settings do not allow the code in this workbook" & vbNewLine & _
                        " to work as designed and you will get some error messages." & vbNewLine & vbNewLine & _
                        "To allow the code to function correctly and without errors you need" & vbNewLine & _
                        " to change your security setting as follows:" & vbNewLine & vbNewLine & _
                        "    1. Select File - Options - Trust Center - Trust Center Settings... show the security dialog" & vbNewLine & _
                        "    2. Select Macro Settings on the left" & vbNewLine & _
                        "    2. Click the 'Trusted Sources' tab" & vbNewLine & _
                        "    3. Place a checkmark next to 'Trust Access to VBA Project Object Model'" & vbNewLine & _
                        "    4. Click OK." & vbNewLine & vbNewLine & _
                        "Do you want the security dialog shown now?", vbOKOnly + vbCritical)
                        If Response = vbOK Then Application.CommandBars("Macro").Controls("Security...").Execute
                 AddRefsIfAccessAllowed = False
                 Exit Function
           Else
                AddRefsIfAccessAllowed = True
           End If
     End If

     'Call AddReference

End Function


Public Sub AddReference()

     Dim Reference As Object

     With ThisWorkbook.VBProject
           For Each Reference In .References
                 If Reference.Description Like "Microsoft Visual Basic for Applications Extensibility*" Then Exit Sub
           Next
           .References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
     End With

End Sub











Public Sub MaintenanceForAddActivityButton()
    Dim i%, c%
    Dim v
    Dim arrSheetList(), arrTemplateCodeLines()
    Dim strTemplateCode$
    
    ReDim arrSheetList(1 To 100)
    ReDim arrTemplateCodeLines(1 To 50)
    i = 1
    
    With ThisWorkbook.VBProject
        For Each v In .VBComponents
            If strTemplateCode <> vbNullString Then Exit For
            If v.Type = vbext_ct_Document Then
                Debug.Print v.Properties.Item("Name")
                If v.Properties.Item("Name") Like "templatesheet" Then
                    For i = 1 To v.CodeModule.CountOfLines
                        strTemplateCode = v.CodeModule.Lines(1, v.CodeModule.CountOfLines)
                        Debug.Print strTemplateCode
                    Next i
                End If
            End If
        Next
    End With

    With ThisWorkbook.VBProject
        For Each v In .VBComponents
            If v.Type = vbext_ct_Document Then
                Debug.Print v.Properties.Item("Name")
                If Left(v.Properties.Item("Name"), 3) Like "FY1" Then
                    v.CodeModule.DeleteLines 1, v.CodeModule.CountOfLines
                    v.CodeModule.AddFromString (strTemplateCode)
                End If
            End If
        Next
    End With



End Sub















