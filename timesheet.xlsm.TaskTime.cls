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
Private dTimes() As Single
Private mPeriodStartDate As Date
Private mTaskName As String
Private SubTasks As Scripting.Dictionary
Private LastSubTask As TaskTime
Private mGrouping As TaskTimeGrouping
Private mMaxTaskNameLength As Integer
Private pTimePerDay As Single
Private Chr_Horiz As String
Private Chr_Cross As String
Private Chr_Vert As String

Private Const DayNames = "MonTueWedThuFriSatSun"
Private Const MonthNames = "JanFebMarAprMayJunJulAugSepOctNovDec"


Public Property Get TaskName() As String
    TaskName = mTaskName
End Property

Public Property Let TaskName(value As String)
    mTaskName = value
    mMaxTaskNameLength = Len(value)
End Property

Public Property Get MaxTaskNameLength() As Integer
    MaxTaskNameLength = mMaxTaskNameLength
End Property

Public Property Get Time() As Single
    Time = pTime
End Property

Public Property Let TimePerDay(value As Single)
    pTimePerDay = value
End Property

Public Sub Increment(Optional i As Single = 0.25, Optional RecurseLevels As Integer = 2, Optional SubTaskName As String = "", Optional SubSubTaskName As String = "", Optional WithDateRangeBounds As Boolean = False, Optional TTime As Date = 0, Optional Grouping As TaskTimeGrouping = TaskTimeGrouping.None, Optional PeriodStartDate As Date)
Dim week As Integer

On Error GoTo Err_Proc
GoTo Proc
Err_Proc:
    If vbYes = MsgBox(Err.Description & vbCrLf & "Debug?", vbYesNo Or vbCritical, "Error") Then
        Stop
        Resume
    End If
    Exit Sub
Proc:
    If mPeriodStartDate = 0 And Not PeriodStartDate = 0 Then
        mPeriodStartDate = PeriodStartDate
        Select Case Grouping
            Case TaskTimeGrouping.ByDay:
                ReDim Preserve dTimes(1 To 7)
            Case TaskTimeGrouping.ByWeek:
                ReDim Preserve dTimes(1 To 5)
            Case TaskTimeGrouping.ByMonth:
                ReDim Preserve dTimes(1 To 12)
        End Select
    End If
        

    ' Debug.Print TaskName, RecurseLevels, i, SubTaskName, SubSubTaskName, WithDateRangeBounds, Format(TTime, "yyyymmdd"), Grouping
    pTime = pTime + i
    mGrouping = Grouping
    Select Case Grouping
        Case TaskTimeGrouping.ByDay:
            dTimes(DatePart("w", TTime, vbMonday)) = dTimes(DatePart("w", TTime, vbMonday)) + i
        Case TaskTimeGrouping.ByWeek:
            week = DatePart("ww", TTime, vbMonday, vbFirstJan1) - DatePart("ww", mPeriodStartDate, vbMonday, vbFirstJan1) + 1
            dTimes(week) = dTimes(week) + i
        Case TaskTimeGrouping.ByMonth:
            dTimes(DatePart("m", TTime)) = dTimes(DatePart("m", TTime)) + i
    End Select
    
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
              LastSubTask.TimePerDay = pTimePerDay
              Set SubTasks(SubTaskName) = LastSubTask
            End If
        End If
        LastSubTask.Increment i, RecurseLevels - 1, SubSubTaskName, TTime:=TTime, WithDateRangeBounds:=WithDateRangeBounds, Grouping:=Grouping, PeriodStartDate:=PeriodStartDate
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
    Dim HorizSeparator As String
    Dim Header As String
    Dim AddHeader As Boolean
    If pTime = 0 Then Exit Function
    Dim CellWidth As Integer
    
    AddHeader = TopLevel And Not AsArray And mGrouping <> TaskTimeGrouping.None
    
    Select Case mGrouping
        Case TaskTimeGrouping.ByDay: CellWidth = 5
        Case TaskTimeGrouping.ByWeek: CellWidth = 5
        Case TaskTimeGrouping.ByMonth: CellWidth = 6
        Case Else: CellWidth = 5
    End Select


    
    If AddHeader Or Not SubTasks Is Nothing Then
        HorizSeparator = String(Len(TaskName), Chr_Horiz) & Chr(254)
        If AddHeader Then Header = Chr(253)
        If Not AsArray Then
            If mGrouping <> TaskTimeGrouping.None Then
                For d = LBound(dTimes) To UBound(dTimes)
                    HorizSeparator = HorizSeparator & Chr_Horiz + Chr_Cross & FormatTime(-1, CellWidth)
                    If AddHeader Then
                        Select Case mGrouping
                            Case TaskTimeGrouping.ByDay:
                                Header = Header & " " & Chr_Vert & FormatTime(0, CellWidth, Mid(DayNames, 3 * d - 2, 3))
                            Case TaskTimeGrouping.ByWeek:
                                Header = Header & " " & Chr_Vert & FormatTime(0, CellWidth, Mid(MonthNames, 3 * d - 2, 3))
                            Case TaskTimeGrouping.ByMonth:
                                Header = Header & " " & Chr_Vert & FormatTime(0, CellWidth, Mid(MonthNames, 3 * d - 2, 3))
                        End Select
                    End If
                Next d
                HorizSeparator = HorizSeparator & Chr_Horiz + Chr_Cross
                If AddHeader Then Header = Header & " " & Chr_Vert
            End If
            HorizSeparator = HorizSeparator & Chr_Horiz & FormatTime(-1, CellWidth)
            If AddHeader Then Header = Header & " " & FormatTime(0, CellWidth)
            If WithBounds Then
                Header = Header & String(28, " ")
                HorizSeparator = HorizSeparator & String(28, Chr_Horiz)
            End If
        End If
        HorizSeparator = HorizSeparator & vbCrLf
        If AddHeader Then Header = Header & vbCrLf
    End If
    If AddHeader Then
        ToString = Header
    End If
    ToString = Header & HorizSeparator & IIf(TaskName = "", "Unnamed task", TaskName) & Chr(255) _
        
    If Not AsArray Then
        If mGrouping <> TaskTimeGrouping.None Then
            For d = LBound(dTimes) To UBound(dTimes)
                ToString = ToString & " " & Chr_Vert & FormatTime(dTimes(d), CellWidth)
            Next d
            ToString = ToString & " " & Chr_Vert
        End If
        ToString = ToString & " " & FormatTime(pTime, CellWidth)
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
                SubTasksArray(i) = Array(OrderByValue, SubTasks(SubTask).ToString(WithBounds, AsArray, OrdedBy, TopLevel:=False) & Chr_Vert)
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

Private Function NextColumn(ByRef Column As Integer) As Integer
    Column = Column + 1
    NextColumn = Column
End Function

Private Function NextRow(ByRef Target As Range) As Range
    Set Target = Target.Worksheet.Range( _
        Target.Worksheet.Cells(Target.Row + 1, Target.Column), _
        Target.Worksheet.Cells(Target.Row + Target.Rows.Count, Target.Column + Target.Columns.Count - 1) _
    )
    Set NextRow = Target
End Function

Public Sub ApplyToRange( _
    Target As Range, _
    Optional AsArray As Boolean, _
    Optional OrderBy As String, _
    Optional WithBounds As Boolean, _
    Optional Level As Integer = 0 _
)
    Dim SubTask As Variant
    Dim SubTasksArray() As Variant, i As Integer, d As Integer
    Dim OrderByValue As Variant
    Dim Column As Integer
    If pTime = 0 Then Exit Sub
    
    Dim AddHeader As Boolean
    AddHeader = Level = 0 And Not AsArray And mGrouping <> TaskTimeGrouping.None
    
    If AddHeader Or Not SubTasks Is Nothing Then
        If Not AsArray Then
            If mGrouping <> TaskTimeGrouping.None Then
                Column = 1 'skip 1st column
                For d = LBound(dTimes) To UBound(dTimes)
                    If AddHeader Then
                        Select Case mGrouping
                            Case TaskTimeGrouping.ByDay:
                                Target.Cells(1, NextColumn(Column)).value = Mid(DayNames, 3 * d - 2, 3)
                            Case TaskTimeGrouping.ByWeek:
                                Target.Cells(1, NextColumn(Column)).value = Mid(MonthNames, 3 * d - 2, 3)
                            Case TaskTimeGrouping.ByMonth:
                                Target.Cells(1, NextColumn(Column)).value = Mid(MonthNames, 3 * d - 2, 3)
                        End Select
                    End If
                Next d
            End If
        End If
    End If
    Set Target = Target.Worksheet.Range( _
        Target.Worksheet.Cells(Target.Row + 1, Target.Column), _
        Target.Worksheet.Cells(Target.Row + Target.Rows.Count, Target.Column + Target.Columns.Count - 1) _
    )
    Column = 0
    Target.Cells(1, NextColumn(Column)).value = IIf(TaskName = "", "Unnamed task", TaskName)
        
    If Not AsArray Then
        If mGrouping <> TaskTimeGrouping.None Then
            For d = LBound(dTimes) To UBound(dTimes)
                Target.Cells(1, NextColumn(Column)).value = dTimes(d)
            Next d
        End If
        Target.Cells(1, NextColumn(Column)).value = pTime
    End If
    If WithBounds Then Target.Cells(1, NextColumn(Column)).value = "{" & Format(StartTime, "yyyy-mm-dd") & " => " & Format(EndTime, "yyyy-mm-dd") & "}"
    If Not SubTasks Is Nothing Then
        ReDim SubTasksArray(SubTasks.Count - 1)
        For Each SubTask In SubTasks.Keys
            OrderByValue = IIf(OrderBy = "Time", -SubTasks(SubTask).Time, _
                           IIf(OrderBy = "Date", SubTasks(SubTask).StartTime, _
                           SubTasks(SubTask).TaskName _
                           ))
            SubTasksArray(i) = Array(OrderByValue, SubTasks(SubTask))
            i = i + 1
        Next SubTask
        QuickSort SubTasksArray
        For i = LBound(SubTasksArray) To UBound(SubTasksArray)
            ApplyToRange NextRow(Target), AsArray, OrderBy, WithBounds, Level + 1
        Next i
    End If
End Sub


Private Function FormatTime(tm As Single, CellWidth As Integer, Optional txt As String) As String
    If txt <> "" Then
        txt = Trim(Mid(txt, 1, CellWidth))
        FormatTime = " " & Mid(Mid("      ", 1, (CellWidth - Len(txt)) / 2) & txt & "     ", 1, CellWidth) & " "
    ElseIf tm = 0 Then
        FormatTime = String(CellWidth + 2, " ")
    ElseIf tm = -1 Then
        FormatTime = String(CellWidth + 2, Chr_Horiz) ' "------"
    Else
        FormatTime = Right("    " & Format(Int(tm), "#0") & ":" & Right(Format(tm / 24, "hh:mm"), 2), CellWidth + 2)
    End If
    Debug.Assert Len(FormatTime) = CellWidth + 2
    FormatTime = FormatTime & FormatTimeInDay(tm, CellWidth)
End Function

Private Function FormatTimeInDay(tm As Single, CellWidth As Integer) As String
    If pTimePerDay < 0 Then Exit Function
    If tm = 0 Then
        FormatTimeInDay = String(CellWidth, " ")
    ElseIf tm = -1 Then
        FormatTimeInDay = String(CellWidth, Chr_Horiz) ' "------"
    Else
        Dim fullDays As Integer: fullDays = Int(tm / pTimePerDay)
        If fullDays > 0 Then
            FormatTimeInDay = Right("   " & Format(fullDays, "#0"), CellWidth - 2)
        Else
            FormatTimeInDay = String(CellWidth - 2, "   ")
        End If
        Select Case tm / pTimePerDay - Int(tm / pTimePerDay)
            Case Is < 1 / 8:
                If fullDays > 0 Then
                    FormatTimeInDay = FormatTimeInDay & " d"
                Else
                    Select Case 8 * tm
                        Case 1 To 3: FormatTimeInDay = Right("   " & ChrW(&HBC) & "h", CellWidth)
                        Case 3 To 5: FormatTimeInDay = Right("   " & ChrW(&HBD) & "h", CellWidth)
                        Case 5 To 7: FormatTimeInDay = Right("   " & ChrW(&HBE) & "h", CellWidth)
                    End Select
                End If
            Case Is < 3 / 16: FormatTimeInDay = FormatTimeInDay & ChrW(&H215B) & "d" ' 1/8
            Case Is < 5 / 16: FormatTimeInDay = FormatTimeInDay & ChrW(&HBC) & "d" ' 1/4
            Case Is < 7 / 16: FormatTimeInDay = FormatTimeInDay & ChrW(&H215C) & "d" ' 3/8
            Case Is < 9 / 16: FormatTimeInDay = FormatTimeInDay & ChrW(&HBD) & "d" ' 1/2
            Case Is < 11 / 16: FormatTimeInDay = FormatTimeInDay & ChrW(&H215C) & "d" ' 5/8
            Case Is < 13 / 16: FormatTimeInDay = FormatTimeInDay & ChrW(&HBE) & "d" ' 3/4
            Case Is < 15 / 16: FormatTimeInDay = FormatTimeInDay & ChrW(&H215E) & "d" ' 7/8
            Case Else: FormatTimeInDay = Right("     " & Format(1 + Int(tm / pTimePerDay), "#0"), CellWidth - 1) & "d"
        End Select
    End If
    Debug.Assert Len(FormatTimeInDay) = CellWidth
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
        a(i) = Replace(a(i), Chr(254), String(pad + Len(addspace), Chr_Horiz))
        a(i) = Replace(a(i), Chr(253), String(pad + Len(addspace), " "))
    Next i
    ExplodePadder = Join(a, vbCrLf)
End Function

Private Sub Class_Initialize()
    Chr_Horiz = ChrW(&H2500)
    Chr_Cross = ChrW(&H253C)
    Chr_Vert = ChrW(&H2502)
End Sub
