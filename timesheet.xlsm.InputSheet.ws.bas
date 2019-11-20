VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim report As String
Dim Title As String
Dim InputRng As Range: Set InputRng = Range("InputRange")
Dim YearsRng As Range: Set YearsRng = Range("YearList")
Dim MonthRng As Range: Set MonthRng = Range("MonthsList")
Dim WeeksRng As Range: Set WeeksRng = Range("WeeksList")
Dim DatesRng As Range: Set DatesRng = Range("Dates")
Dim DayDueTime As Single: DayDueTime = Range("DayDueTime").value * 24

Dim Year As Integer: Year = Range("SummaryYear").value
Dim Mth_ As Integer: Mth_ = Month(Range("SummaryMonth").value)
Dim week As Integer: week = DatePart("ww", Range("SummaryWeek").value, vbMonday)
Dim Day_ As Date:    Day_ = Range("SummaryDay").value
    
    Title = "Activity Summary"
    Select Case Target.Address
        Case Range("YearTag").Address
            report = TextReport(InputRng, DatesRng, YearsRng, Year, Range("Tasksref"), Target.value, DayDueTime, WithDateRangeBounds:=True, OrderBy:="Time") & vbCrLf
        Case Range("SummaryYear").Address
            report = TextReport(InputRng, DatesRng, YearsRng, Year, Range("Tasksref"), Target.value, DayDueTime, WithDateRangeBounds:=True, OrderBy:="Date", Grouping:=ByMonth) & vbCrLf
        Case Range("SummaryMonth").Address
            report = TextReport(InputRng, DatesRng, MonthRng, Mth_, Range("Tasksref"), Format(Target, "mmm-yyyy"), DayDueTime, OrderBy:="Time") & vbCrLf
        Case Range("SummaryWeek").Address
            report = TextReport(InputRng, DatesRng, WeeksRng, week, Range("Tasksref"), "Week " & Target.value, DayDueTime, OrderBy:="Time") & vbCrLf
        Case Range("WeekTag").Address
            report = TextReport(InputRng, DatesRng, WeeksRng, week, Range("Tasksref"), "Week " & Range("SummaryWeek").value, -1, OrderBy:="Name", Grouping:=TaskTimeGrouping.ByDay, RecurseLevels:=2) & vbCrLf
        Case Range("SummaryDay").Address
            report = TextReport(InputRng, DatesRng, DatesRng, Day_, Range("Tasksref"), Format(Target.value, "dd-mmm"), DayDueTime) & vbCrLf
        Case Range("DayTag").Address
            ThisWorkbook.GoNow
        Case Range("WorkWeek").Address
            report = WorkReport("WeeklyAggregates", "yyyy-ww (mmm-dd)", "Hol", "ADM", "*")
        Case Range("WorkMonth").Address
            report = WorkReport("MonthlyAggregates", "mmm-yy", "Hol", "ADM", "*")
        Case Else
             If RangeRelation(Target, Range("TasksRefFullRange")) = "Included" Then
                report = TextReport(InputRng, DatesRng, YearsRng, Year, Target, Target.value, DayDueTime, WithDateRangeBounds:=True, OrderBy:="Date")
                Title = "Yearly summary"
            ElseIf RangeRelation(Target, Range("InputRange")) = "Included" Then
                Cancel = TryFillWithPreviousTask(Target)
            End If
    End Select
    If report <> "" Then
        ClipBoard.SetText report
        FrmTaskReport.ShowMessage report, Title:=Title
        Cancel = True
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Static selectedRow As Long
    If selectedRow = Target.Row Then Exit Sub
    selectedRow = Target.Row
    TimeSheet.SetSummaries Target
End Sub
