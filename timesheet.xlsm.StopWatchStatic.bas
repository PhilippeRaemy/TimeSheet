Attribute VB_Name = "StopWatchStatic"
Option Explicit
Function NewStopWatch(Name As String) As StopWatch
    Set NewStopWatch = New StopWatch
    NewStopWatch.SetName Name
End Function
