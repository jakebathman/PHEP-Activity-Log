VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDayDetailPicker 
   Caption         =   "Which day would you like to manage in detail?"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10995
   OleObjectBlob   =   "frmDayDetailPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDayDetailPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v4.4

Option Explicit



Private Sub UserForm_Initialize()
    Dim i%
    For i = 1 To 7
        With frmDayDetailPicker
            .Controls("CommandButton" & i).BackColor = &HDBDCF2
        End With
    Next i
    For i = 8 To 14
        With frmDayDetailPicker
            .Controls("CommandButton" & i).BackColor = &HF1D9C5
        End With
    Next i
End Sub
