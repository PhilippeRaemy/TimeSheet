VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaskTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public StartTime As Date, EndTime As Date

Private pTime As Single
Private dTimes(1 To 7) As Single
Private mTaskName As String
Private SubTasks As Scripting.Dictionary
Private LastSubTask As TaskTime
Private IsByWeekDay As Boolean
Private mMaxTaskNameLength As Integer

Public Property Get TaskName() As String
    TaskName = mTaskName
End Property

Public Property Let TaskName(Value As String)
    mTaskName = Value
    mMaxTaskNameLength = Len(Value)
End Property

Public Property Get MaxTaskNameLength() As Integer
    MaxTaskNameLength = mMaxTaskNameLength
End Property

Public Property Get Time() As Single
    Time = pTime
End Property
Public Sub Increment(Optional i As Single = 0.25, Optional RecurseLevels As Integer = 2, Optional SubTaskName As String = "", Optional SubSubTaskName As String = "", Optional WithDateRangeBounds As Boolean = False, Optional TTime As Date = 0, Optional ByWeekDay As Boolean)
On Error GoTo Err_Proc
GoTo Proc
Err_Proc:
    If vbYes = MsgBox(Err.Description & vbCrLf & "Debug?", vbYesNo Or vbCritical, "Error") Then
        Stop
        Resume
    End If
    Exit Sub
Proc:
    ' Debug.Print TaskName, RecurseLevels, i, SubTaskName, SubSubTaskName, WithDateRangeBounds, Format(TTime, "yyyymmdd"), ByWeekDay
    pTime = pTime + i
    If ByWeekDay Then
        IsByWeekDay = ByWeekDay
        dTimes(DatePart("w", TTime, vbMonday)) = dTimes(DatePart("w", TTime, vbMonday)) + i
    End If
    If RecurseLevels > 0 Then
        If SubTasks Is Nothing Then
            Set SubTasks = New Scripting.Dictionary
            SubTasks.CompareMode = TextCompare
        End If
        If SubTaskName <> "" Or LastSubTask Is Nothing Then
            Dim doCreate As Boolean: doCreate = True
            If SubTasks.Exists(SubTaskName) Then
              If IsObject(SubTasks(SubTaskName)) Then
                Set LastSubTask = SubTasks(SubTaskName)
                doCreate = False
              End If
            End If
            If doCreate Then
              Set LastSubTask = New TaskTime
              LastSubTask.TaskName = SubTaskName
              Set SubTasks(SubTaskName) = LastSubTask
            End If
        End If
        LastSubTask.Increment i, RecurseLevels - 1, SubSubTaskName, TTime:=TTime, WithDateRangeBounds:=WithDateRangeBounds, ByWeekDay:=ByWeekDay
        If MaxTaskNameLength < LastSubTask.MaxTaskNameLength Then mMaxTaskNameLength = LastSubTask.MaxTaskNameLength
    End If
    If WithDateRangeBounds And TTime <> 0 Then
        If StartTime = 0 Or EndTime = 0 Then
            StartTime = TTime
            EndTime = TTime
        Else
            If StartTime > TTime Then StartTime = TTime
            If EndTime < TTime Then EndTime = TTime
        End If
    End If
End Sub

Public Function ToString( _
    Optional WithBounds As Boolean, _
    Optional AsArray As Boolean, _
    Optional OrdedBy As String, _
    Optional TopLevel As Boolean = True _
) As String
    Dim SubTask As Variant
    Dim SubTasksArray() As Variant, i As Integer, d As Integer
    Dim OrderByValue As Variant
    If pTime = 0 Then Exit Function
    ToString = IIf(TaskName = "", "Unnamed task", TaskName) & Chr(255)
    If Not AsArray Then
        If IsByWeekDay Then
            For d = LBound(dTimes) To UBound(dTimes)
                ToString = ToString & " ." & FormatTime(dTimes(d))
            Next d
              ToString = ToString & " |"
        End If
        ToString = ToString & " " & FormatTime(pTime)
    End If
    If WithBounds Then ToString = ToString & "  {" & Format(StartTime, "yyyy-mm-dd") & " => " & Format(EndTime, "yyyy-mm-dd") & "}"
    If Not SubTasks Is Nothing Then
        ReDim SubTasksArray(SubTasks.Count - 1)
        For Each SubTask In SubTasks.Keys
            OrderByValue = IIf(OrdedBy = "Time", -SubTasks(SubTask).Time, _
                           IIf(OrdedBy = "Date", SubTasks(SubTask).StartTime, _
                           SubTasks(SubTask).TaskName _
                           ))
            If AsArray Then
                SubTasksArray(i) = Array(OrderByValue, SubTasks(SubTask).ToString(WithBounds, AsArray, OrdedBy, TopLevel:=False) & "|")
            Else
                SubTasksArray(i) = Array(OrderByValue, vbCrLf & "*   " & Replace(Replace(Replace(SubTasks(SubTask).ToString(WithBounds, AsArray, OrdedBy, TopLevel:=False), vbCrLf, vbCrLf & "    "), "**  ", "*** "), "*   ", "**  "))
            End If
            i = i + 1
        Next SubTask
        QuickSort SubTasksArray
        For i = LBound(SubTasksArray) To UBound(SubTasksArray)
            ToString = ToString & SubTasksArray(i)(1)
        Next i
    End If
    If TopLevel Then
        ToString = ExplodePadder(ToString)
    End If
End Function
Private Function FormatTime(tm As Single) As String
    If tm = 0 Then
        FormatTime = "      "
    Else
        FormatTime = Right("  " & Format(Int(tm), "#0") & ":" & Right(Format(tm / 24, "hh:mm"), 2), 6)
    End If
End Function

Private Function ExplodePadder(s As String) As String
Dim a As Variant, l As Integer, i As Integer, pad As Integer, addspace As String
Const trailing = " . "

    a = Split(s, vbCrLf)
    Dim line As Variant
    For Each line In a
        If Len(line) > l Then l = Len(line)
    Next line
    For i = LBound(a) To UBound(a)
        pad = l - Len(a(i)) + 1
        addspace = String(pad Mod Len(trailing), " ")
        pad = pad - pad Mod Len(trailing)
        a(i) = Replace(a(i), Chr(255), addspace & Replace(String(pad, " "), String(Len(trailing), " "), trailing))
    Next i
    ExplodePadder = Join(a, vbCrLf)
End Function
