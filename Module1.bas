Attribute VB_Name = "Module1"
'v4

Option Explicit

Public Function fChangeLog(ByRef s1 As String, ByRef s2 As String, ByRef s3 As String, ByRef s4 As String)
    'Takes in three string variables to return the current changelog in sections
Dim s$

'   Version header
s1 = "Version 4"

'   New Features
s2 = s2 & vbTab & "-- " & vbNewLine
s2 = s2 & vbTab & "-- " & vbNewLine
s2 = s2 & vbTab & "-- " & vbNewLine

'   Bug fixes
s3 = s3 & vbTab & "-- " & vbNewLine
s3 = s3 & vbTab & "-- " & vbNewLine
s3 = s3 & vbTab & "-- " & vbNewLine

'   Known issues
s4 = s4 & vbTab & "-- " & vbNewLine
s4 = s4 & vbTab & "-- " & vbNewLine
s4 = s4 & vbTab & "-- " & vbNewLine




End Function
