Attribute VB_Name = "Module3"
Option Explicit

'Variable to hold name of ToggleButton that was clicked.
Public Clicked As String

' Code is from: http://support.microsoft.com/kb/213714

Public Sub ExclusiveToggleButtons()

   Dim Toggle As Control

   'Loop through all of the ToggleButtons on Frame1
   For Each Toggle In frmAddActivity.Controls
    
    If InStr(1, Toggle.Name, "tog", vbTextCompare) > 0 Then
      'If Name of ToggleButton matches name of ToggleButton
      'that was clicked...
      If Toggle.Name = Clicked Then
         '...select the button
          Toggle.Value = True
       Else
         '...deselect the button
         Toggle.Value = False
       End If
       End If
    Next
End Sub
