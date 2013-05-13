Attribute VB_Name = "u_Update_Code"
Option Explicit
Public Const pthUpdatedWorkbookPath = "\\ccdata01\homeland_security\PHEP Documentation\Monthly Reports\Activity Tracking\"
'Public Const pthUpdatedWorkbookPath = "C:\Users\e008922\Dropbox\_Work\Monthly reports\AO reports\PHEP drive\"


Public Sub uUpdateCode()
    Dim arrListOfModules(), arrListOfNewModules()
    Dim intNumModules%, intNumNewModules%, i%, j%
    Dim fVBProj As VBIDE.VBProject
    Dim tVBProj As VBIDE.VBProject
    Dim tFilePathFull$, strVers$, strVersNew$
    Dim t
    Dim v
    Dim boolTriedTwice As Boolean
    
    boolTriedTwice = False
    
    strVers = Sheets("Refs").Range("L2").Value
    
    
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
    If f <> vbNullString And StrComp(f, ActiveWorkbook.Name, vbTextCompare) <> 0 Then
        Dim app As New Excel.Application
        Dim book As Excel.Workbook
        Set book = app.Workbooks.Add(tFilePathFull)

        Set fVBProj = book.VBProject
        
        
        Call uListModules(arrListOfModules, intNumModules, fVBProj)
        Call uListModules(arrListOfNewModules, intNumNewModules, tVBProj)
        
        For i = 1 To intNumNewModules
            'Cells(15 + i, 2).Value = arrListOfModules(i, 1)
            If arrListOfNewModules(i, 1) <> "u_Update_Code" And arrListOfNewModules(i, 1) <> "u_List_Modules" Then
                If arrListOfNewModules(i, 2) = "Code Module" Or arrListOfNewModules(i, 2) = "UserForm" Or arrListOfNewModules(i, 2) = "Document Module" Then
                        v = CopyModule(arrListOfNewModules(i, 1), fVBProj, tVBProj, True)
                        'Cells(15 + i, 1).Value = v
                        'Cells(15 + i, 2).Value = arrListOfNewModules(i, 1)
                End If
            End If
'            t = Timer
'            While Timer < t + 0.25
'                DoEvents
'            Wend
        Next i
        
        
        book.Close SaveChanges:=False
        Set book = Nothing
        app.Quit
        Set app = Nothing
        
        v = TotalCodeLinesInVBComponent(tVBProj.VBComponents("v_Version_Num")) - 3
        'Debug.Print v
        strVersNew = CStr(v)
        Sheets("Refs").Range("L2").Value = strVersNew
        
        
        Call MsgBox("Update complete!!" & vbNewLine & vbNewLine _
                & "This is Version " & strVersNew & " of this tool." & vbNewLine & vbNewLine _
                & " ")
        
        
    Else
        Call MsgBox("Looks like you've got the latest version!" & vbNewLine & vbNewLine _
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
    MsgBox ("Sorry! Something went wrong :(" & vbNewLine & vbNewLine & "The code was NOT updated.")


End Sub



' Code below copied from http://www.cpearson.com/excel/vbe.aspx
Function CopyModule(ByVal ModuleName As String, _
    FromVBProject As VBIDE.VBProject, _
    ToVBProject As VBIDE.VBProject, _
    OverwriteExisting As Boolean) As Boolean
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
    ' ModuleName            The name of the module to copy.
    '
    ' OverwriteExisting     If True, the VBComponent named ModuleName
    '                       in ToVBProject will be removed before
    '                       importing the module. If False and
    '                       a VBComponent named ModuleName exists
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
    
    If Trim(ModuleName) = vbNullString Then
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
    Set VBComp = FromVBProject.VBComponents(ModuleName)
    If Err.Number <> 0 Then
        CopyModule = False
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FName is the name of the temporary file to be
    ' used in the Export/Import code.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FName = Environ("Temp") & "\" & ModuleName & ".bas"
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
        With ToVBProject.VBComponents
            .Remove .Item(ModuleName)
        End With
    Else
        '''''''''''''''''''''''''''''''''''''''''
        ' OverwriteExisting is False. If there is
        ' already a VBComponent named ModuleName,
        ' exit with a return code of False.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Set VBComp = ToVBProject.VBComponents(ModuleName)
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
    FromVBProject.VBComponents(ModuleName).Export FileName:=FName
    
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
