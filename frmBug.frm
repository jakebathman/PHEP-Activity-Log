VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBug 
   Caption         =   "Report a bug"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   OleObjectBlob   =   "frmBug.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnEmailJake_Click()
    ThisWorkbook.FollowHyperlink "mailto:jbathman@co.collin.tx.us?subject=Activity%20log%20bug%20report&body=Note%3A%20if%20this%20is%20a%20bug%20report%2C%20make%20sure%20to%20describe%3A%20%0A%0A1)%20What%20steps%20will%20reproduce%20the%20problem%3F%20%0A2)%20What%20do%20you%20see%2C%20such%20as%20error%20messages%20or%20dialog%20boxes%3F%0A%0AYou%20can%20also%20attach%20screenshots%2C%20if%20it%20will%20help."
    Unload Me
End Sub

Private Sub btnSubmitBug_Click()
    ThisWorkbook.FollowHyperlink "https://code.google.com/p/phep-activity-log/issues/entry?template=Defect%20report%20from%20user"
    Unload Me
End Sub



Private Sub UserForm_Activate()
    Dim dblWidthBySix#
    With frmBug
        .imgHeader.AutoSize = False
        .imgHeader.AutoSize = True
        .Width = .imgHeader.Width
        .imgHeader.Left = 0
        .imgHeader.Top = 0
        .lblHeaderMagic.AutoSize = True
        .lblHeaderMagic.Left = (.imgHeader.Width / 3) - 15
        .lblHeaderMagic.Top = 8

        dblWidthBySix = .Width / 6

        .lblInstructions.Left = (.Width / 2) - (.lblInstructions.Width / 2)
        .lblNoteGmail.Left = (.Width / 2) - (.lblNoteGmail.Width / 2)
        .lblOtherOption.Left = dblWidthBySix
        .btnEmailJake.Left = .lblOtherOption.Left + .lblOtherOption.Width + (dblWidthBySix / 2)
        .btnSubmitBug.Left = (.Width / 2) - (.btnSubmitBug.Width / 2)
        .btnClose.Left = (.Width / 2) - (.btnClose.Width / 2)

        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_Click()
    MsgBox ("Width in points = " & ActiveWindow.Width & Chr(13) & _
            "Depth in points = " & ActiveWindow.Height & Chr(13) & _
            "Width in Pixels = " & ActiveWindow.PointsToScreenPixelsX(ActiveWindow.Width) & Chr(13) & _
            "Depth in Pixels = " & ActiveWindow.PointsToScreenPixelsY(ActiveWindow.Height))

    Me.Label4.Caption = "One point is equal to " & Round(ActiveWindow.PointsToScreenPixelsY(ActiveWindow.Height) / ActiveWindow.Height, 4) & " pixels"
End Sub
