VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickWorkingSheet 
   Caption         =   "Which sheet is your currently active one?"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6945
   OleObjectBlob   =   "frmPickWorkingSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPickWorkingSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v4.1

Option Explicit



Private Sub btnSubmit_Click()
    With frmPickWorkingSheet
        If .cmbSheets.Value <> vbNullString Then
            Sheets("Refs").Range("P2").Value = .cmbSheets.Value
            .Hide
        End If
    End With
End Sub

Private Sub UserForm_Activate()
    With frmPickWorkingSheet
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim i%
    Dim v
    Dim arrSheets()
    
    i = 1
    For Each v In ThisWorkbook.Sheets
        If Left(v.Name, 3) Like "FY1" Then
            ReDim Preserve arrSheets(1 To i)
            arrSheets(i) = v.Name
            i = i + 1
        End If
    Next
    
    If QSortInPlace(arrSheets, -1, -1, True, vbTextCompare, False) = True Then
        For i = 1 To UBound(arrSheets)
            frmPickWorkingSheet.cmbSheets.AddItem (arrSheets(i))
        Next i
    End If
End Sub

Private Sub UserForm_Terminate()
    Sheets("Refs").Range("P2").Value = ThisWorkbook.Sheets(2).Name
End Sub
