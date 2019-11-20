Attribute VB_Name = "TaskTimeStatic"
Option Explicit
Public Enum TaskTimeGrouping
    None = 0
    ByDay = 1
    ByWeek = 2
    ByMonth = 3
End Enum

Public Function TextReport( _
    DataRange As Range, _
    DateRange As Range, _
    PeriodRange As Range, _
    PeriodValue As Variant, _
    CategoryReferences As Range, _
    Title As String, _
    TimePerDay As Single, _
    Optional WithDateRangeBounds As Boolean = False, _
    Optional AsArray As Boolean = False, _
    Optional OrderBy As String = "", _
    Optional Grouping As TaskTimeGrouping = TaskTimeGrouping.None, _
    Optional RecurseLevels As Integer = 2 _
) As String
Dim r As Integer, c As Integer, CatCell As Range, DataCell As Range
Dim color As Long
Dim ColorDic As Scripting.Dictionary
Dim Category As Variant
Dim txt As String
Dim GlobalTaskTime As TaskTime
Dim CurrentTaskTime As TaskTime
Dim PeriodStartDate As Date: PeriodStartDate = 0

On Error GoTo Err_Proc
GoTo Proc
Err_Proc:
    If vbYes = MsgBox(Err.Description & vbCrLf & "Debug?", vbYesNo Or vbCritical, "Error") Then
        Stop
        Resume
    End If
    TextReport = Err.Description
    Exit Function
Proc:
    Set ColorDic = New Scripting.Dictionary
    Set GlobalTaskTime = New TaskTime
    GlobalTaskTime.TaskName = Title
    GlobalTaskTime.TimePerDay = TimePerDay
    For Each CatCell In CategoryReferences
      ColorDic.Add CatCell.Interior.color, CatCell.value
    Next CatCell
    For r = 1 To PeriodRange.Rows.Count
        If PeriodRange.Cells(r, 1).value = PeriodValue Then
            If PeriodStartDate = 0 Then PeriodStartDate = DateRange.Cells(r, 1).value
            For c = 1 To DataRange.Columns.Count
                Set DataCell = DataRange.Cells(r, c)
                color = DataCell.Interior.color
                If ColorDic.Exists(color) Then
                    GlobalTaskTime.Increment SubTaskName:=ColorDic(color), _
                        SubSubTaskName:=DataCell.value, _
                        RecurseLevels:=RecurseLevels, _
                        Grouping:=Grouping, _
                        TTime:=DateRange.Cells(r, 1).value, _
                        WithDateRangeBounds:=WithDateRangeBounds, _
                        PeriodStartDate:=PeriodStartDate
                End If
            Next c
        End If
      ' Debug.Print r
    Next r
    TextReport = GlobalTaskTime.ToString(WithDateRangeBounds, AsArray, OrderBy)

    If TextReport = "" Then TextReport = "No data"
End Function


