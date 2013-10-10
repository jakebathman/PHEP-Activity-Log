Attribute VB_Name = "u_List_Modules"
'v4.3

Option Explicit

Sub uListModules(ByRef arrList(), ByRef intNumMods, ByRef vbActiveProj)
    'Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    'Dim WS As Worksheet
    'Dim Rng As Range
    Dim c%

    c = 1

    'Set VBProj = ActiveWorkbook.VBProject
    'Set WS = ActiveWorkbook.Worksheets("Sheet1")
    'Set Rng = WS.Range("A1")
    intNumMods = vbActiveProj.VBComponents.Count
    ReDim arrList(1 To intNumMods, 1 To 2)

    For Each VBComp In vbActiveProj.VBComponents
        arrList(c, 1) = VBComp.Name
        arrList(c, 2) = ComponentTypeToString(VBComp.Type)
        c = c + 1
        'Rng(1, 1).Value = VBComp.Name
        'Rng(1, 2).Value = ComponentTypeToString(VBComp.Type)
        'Set Rng = Rng(2, 1)
    Next VBComp
End Sub


Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function
