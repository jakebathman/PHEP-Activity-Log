Attribute VB_Name = "z_Change_Log"
'v4.2

Option Explicit

Public Function fChangeLog(ByRef s1 As String, ByRef s2 As String, ByRef s3 As String, ByRef s4 As String)
    'Takes in three empty string variables to return (BYREF!) the current changelog in sections
    Dim S$

    '   Version header
    s1 = "Version 4.1"

    '   New Features
    s2 = s2 & vbTab & "-- Big red button to quickly submit bug reports to Google Code project" & vbNewLine
    s2 = s2 & vbTab & "-- " & vbNewLine
    s2 = s2 & vbTab & "-- " & vbNewLine

    '   Bug fixes
    s3 = s3 & vbTab & "-- " & vbNewLine
    s3 = s3 & vbTab & "-- " & vbNewLine
    s3 = s3 & vbTab & "-- " & vbNewLine

    '   Known issues
    s4 = s4 & vbTab & "-- Update code is still broken in some instances" & vbNewLine
    s4 = s4 & vbTab & "-- " & vbNewLine
    s4 = s4 & vbTab & "-- " & vbNewLine




End Function
