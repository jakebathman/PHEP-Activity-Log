VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrefs 
   Caption         =   "User Preferences"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7020
   OleObjectBlob   =   "frmPrefs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v4.1

Option Explicit


Private Sub btnCancel_Click()
    Unload frmPrefs
End Sub

Private Sub btnSave_Click()
    Sheets("Refs").Range("O2").Value = frmPrefs.chkPreSelectToday.Value
    Sheets("Refs").Range("N2").Value = frmPrefs.cmbNames.Value
    Me.Hide
    Unload frmPrefs
End Sub

Private Sub UserForm_Activate()
With frmPrefs
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
End With
End Sub

Private Sub UserForm_Initialize()
Dim i%, intNumNames%

For i = 2 To 15
    If Sheets("Refs").Cells(i, 1).Value = vbNullString Then Exit For
    frmPrefs.cmbNames.AddItem Sheets("Refs").Cells(i, 1).Value
    intNumNames = intNumNames + 1
Next i
frmPrefs.cmbNames.ListRows = intNumNames

If Sheets("Refs").Range("N2").Value <> vbNullString Then frmPrefs.cmbNames.Value = Sheets("Refs").Range("N2").Value

End Sub



