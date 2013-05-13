VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPeriodActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'v2

'Under development
'(NOT currently used or called)

Option Explicit

Public pDate As Date
Public pDay As String
Public pTime As Double
Public pAct As String

''
Public Property Get vDate() As Date
    vDate = pDate
End Property
Public Property Let vDate(ByVal vNewDate As Date)
    pDate = vNewDate
End Property

''
Public Property Get vDay() As String
    vDay = pDay
End Property
Public Property Let vDay(ByVal vNewDay As String)
    pDay = vNewDay
End Property

''
Public Property Get Vtime() As Double
    Vtime = pTime
End Property
Public Property Let Vtime(ByVal vNewTime As doube)
    pTime = vNewTime
End Property

''
Public Property Get vAct() As String
    vAct = pAct
End Property
Public Property Let vAct(ByVal vNewAct As String)
    pAct = vNewAct
End Property

