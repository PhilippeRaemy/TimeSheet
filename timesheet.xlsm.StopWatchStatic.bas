Attribute VB_Name = "StopWatchStatic"
Option Explicit
Function NewStopWatch(Name As String) As Stopwatch
    Set NewStopWatch = New Stopwatch
    NewStopWatch.SetName Name
End Function
