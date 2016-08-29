VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Name As String
Private StartTime As Date

Public Function SetName(Value As String) As Stopwatch
    Name = Value
End Function

Private Sub Class_Initialize()
    StartTime = Now
End Sub

Private Sub Class_Terminate()
    Debug.Print Format(Now, "hh:mm:ss") & ".." & Int((Now - StartTime) * 86400) & "[s]", Name
End Sub
