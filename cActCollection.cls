VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cActCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Coll As Collection
Private CollKeys As Collection

    Private Sub Class_Initialize()
        Set Coll = New Collection
        Set CollKeys = New Collection
    End Sub
    
    Private Sub Class_Terminate()
        Set Coll = Nothing
        Set CollKeys = Nothing
    End Sub
