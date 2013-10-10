VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddSheet 
   Caption         =   "Add a sheet"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   OleObjectBlob   =   "frmAddSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub btnAddSheet_Click()
    Call fAddSheet(frmAddSheet.cmbSheetsToAdd.Value, "MAIN", "20" & Mid(frmAddSheet.cmbSheetsToAdd.Value, 3, 2), Right(frmAddSheet.cmbSheetsToAdd.Value, 2))

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    With frmAddSheet
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.6 * Application.Height) - (0.5 * .Height)
    End With
    
    If frmAddSheet.lblNextPP.Caption = vbNullString Or frmAddSheet.lblNextPP.Caption = "FY00-00" Then
        Call MsgBox("You're too efficient!" & vbNewLine & vbNewLine & "All sheets for the 10 nearest pay periods have been added already. Try again in a few weeks!")
        Me.Hide
    End If
    
    End Sub

Private Sub UserForm_Initialize()
    Dim i%, x%, f%, p%
    Dim dtToday As Date
    Dim strNextPP$

    dtToday = Now

    For i = 2 To 124
        With Sheets("Refs")
            If .Cells(i, 4).Value <= dtToday And .Cells(i, 5).Value >= dtToday Then
                ' get 2 sheets in the immediate future (out of the next 5)
                For x = 1 To 5
                    If f = 2 Then Exit For
                    If .Range("Y" & i + x).Value = False Then
                        frmAddSheet.cmbSheetsToAdd.AddItem (.Range("X" & i + x).Value)
                        If f = 0 Then strNextPP = .Range("X" & i + x).Value
                        f = f + 1
                    End If
                Next x

                ' get 2 sheets in the immediate future (out of the next 5)
                For x = -1 To -5 Step -1
                    If p = 2 Then Exit For
                    If .Range("Y" & i + x).Value = False Then
                        frmAddSheet.cmbSheetsToAdd.AddItem (.Range("X" & i + x).Value)
                        p = p + 1
                    End If
                Next x
                Exit For
            End If
        End With
    Next i

    frmAddSheet.lblNextPP.Caption = strNextPP
    

    

End Sub
