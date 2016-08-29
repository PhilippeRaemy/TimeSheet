Attribute VB_Name = "TimeSheet"
Option Explicit
Sub CmdTask(TaskName As Range)
    Dim sel As Range, cell As Range, r As Integer
    Dim ClearedColor  As Long
    ClearedColor = TaskName.Worksheet.Range("ClearTaskref").Interior.Color
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Set sel = Application.Selection
    If sel.Cells(1, 1).Interior.Color = ClearedColor And Not TaskName.Interior.Color = ClearedColor Then ' only fill-in gaps in the selection
        For Each cell In sel.Cells
            If cell.Interior.Color = ClearedColor Then
                cell.Interior.Color = TaskName.Interior.Color
                cell.Interior.Pattern = TaskName.Interior.Pattern
                cell.Font.Color = TaskName.Font.Color
            End If
        Next cell
    Else ' override whatever is selected
        sel.Interior.Color = TaskName.Interior.Color
        sel.Interior.Pattern = TaskName.Interior.Pattern
        sel.Font.Color = TaskName.Font.Color
    End If
    For r = 1 To sel.Rows.Count
        sel.Cells(r, 1).Value = sel.Cells(r, 1).Value
    Next r
    Application.Calculation = xlCalculationAutomatic
    PivotSheet.PivotTables("WeeklyAggregates").PivotCache.Refresh
    Application.ScreenUpdating = True
End Sub

Sub SetSummaries(Target As Range)
Dim DayDate As Date
    If Not RangeRelation(InputSheet.Range("InputRange"), Target) = "Including" Then Exit Sub
    
    DayDate = InputSheet.Range("Dates").Cells(Target.row - InputSheet.Range("Dates").row + 1, 1).Value
    
    Dim SummaryWeek As Date:    SummaryWeek = DateAdd("d", 1 - DatePart("w", DayDate, vbMonday), DayDate)
    Dim SummaryMonth As Date:   SummaryMonth = DateAdd("d", 1 - Day(DayDate), DayDate)
    Dim SummaryYear As Integer: SummaryYear = Year(DayDate)
    ' check values before to apply to save triggering unnecessary recalculation
    SetSummary DayDate, "SummaryDay", ""
    SetSummary SummaryWeek, "SummaryWeek", "PieChartWeekly"
    SetSummary SummaryMonth, "SummaryMonth", "PieChartMonthly"
End Sub
Sub SetSummary(Value As Variant, RangeName As String, PieChartName As String)
    If InputSheet.Range(RangeName).Value <> Value Then
        ' Debug.Print "SetSummary:" & RangeName, PieChartName, value
        InputSheet.Range(RangeName).Value = Value
        If PieChartName <> "" Then FormatPieChartByName PieChartName
        If PieChartName = "SummaryMonth" Then
            FormatPieChartByName "SummaryYear"
        End If
    End If
End Sub

Sub FormatPieCharts()
    Dim pieChart As ChartObject
    For Each pieChart In InputSheet.ChartObjects
        If pieChart.Name Like "PieChart*" Then
            FormatPieChart pieChart.Chart
        End If
    Next pieChart
End Sub

Sub FormatPieChartByName(PieChartName As String)
    FormatPieChart InputSheet.ChartObjects(PieChartName).Chart
End Sub

Function GetWorkbookFolderName() As String
    Static WorkbookFolderName As String
    If WorkbookFolderName = "" Then
        Dim a As Variant
        a = Split(ThisWorkbook.FullName, "\")
        ReDim Preserve a(UBound(a) - 1)
        WorkbookFolderName = Join(a, "\")
    End If
    GetWorkbookFolderName = WorkbookFolderName
End Function
Sub FormatPieChart(pieChart As Chart)
    Dim ser As Series, i As Integer
    Dim total As Double
    Dim labels As DataLabels
    Dim label As DataLabel
    ' pieChart.ApplyChartTemplate GetWorkbookFolderName & "\PieTimesheet.crtx"
    pieChart.ApplyDataLabels
    Set labels = pieChart.FullSeriesCollection(1).DataLabels
    labels.ShowValue = False
    labels.ShowPercentage = True
    labels.ShowLegendKey = True
    With labels.Format.ThreeD
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With
    With labels.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Transparency = 0.75
        .Solid
    End With


    For Each ser In pieChart.SeriesCollection
'         ser.ApplyDataLabels
        total = 0
        For i = LBound(ser.Values) To UBound(ser.Values)
            total = total + ser.Values(i)
        Next i
        If total = 0 Then total = 1
        For i = LBound(ser.Values) To UBound(ser.Values)
            Set label = ser.DataLabels(i)
            label.ShowBubbleSize = False
            label.ShowCategoryName = False
            label.ShowSeriesName = False
            label.ShowValue = False
            If ser.XValues(i) = "" Then
                label.ShowPercentage = False
            Else
                Select Case ser.Values(i) / total
                    Case Is > 0.03:
                        label.ShowPercentage = True
                    Case Else:
                        label.ShowPercentage = False
                End Select
            End If
            label.ShowLegendKey = label.ShowPercentage
        Next i
    Next ser
End Sub

Public Function CountColored(r As Range, ref As Range)
Dim cell As Range
    For Each cell In r
        If cell.Interior.Color = ref.Interior.Color Then CountColored = CountColored + 1
    Next cell
End Function
Public Function TimeRanges(r As Range, NoWorkRef As Range) As String
Dim NoWorkColor As Long
Dim IsWork As Boolean
Dim c As Integer
    NoWorkColor = NoWorkRef.Interior.Color
    For c = 1 To r.Columns.Count
        Select Case r.Cells(1, c).Interior.Color
            Case NoWorkColor
                If IsWork Then
                    TimeRanges = TimeRanges & "-" & Format((c - 1) / r.Columns.Count, "hh:mm")
                    IsWork = False
                End If
            Case Else
                If Not IsWork Then
                    TimeRanges = IIf(TimeRanges = "", "", TimeRanges & ", ") & Format((c - 1) / r.Columns.Count, "hh:mm")
                    IsWork = True
                End If
        End Select
    Next c
    If IsWork Then
      TimeRanges = TimeRanges & "-24:00"
      IsWork = False
    End If
End Function

Public Function TaskList(DataRange As Range, PeriodRange As Range, PeriodValue As Variant, CategoryReferences As Range, Title As String) As Variant
    Dim a As Variant, b As Variant, i As Integer
    a = Split(TextReport(DataRange, PeriodRange, PeriodValue, CategoryReferences, Title, DatesRange:=Nothing, AsArray:=True), "|")
    b = Array()
    ReDim b(UBound(a), 0)
    For i = 0 To UBound(a)
        b(i, 0) = a(i)
    Next i
    TaskList = b
End Function

Public Function TextReport( _
    DataRange As Range, _
    PeriodRange As Range, _
    PeriodValue As Variant, _
    CategoryReferences As Range, _
    Title As String, _
    Optional WithDateRangeBounds As Boolean = False, _
    Optional DatesRange As Range = Nothing, _
    Optional AsArray As Boolean = False, _
    Optional OrderBy As String = "", _
    Optional ByWeekDay As Boolean = False, _
    Optional RecurseLevels As Integer = 2 _
) As String
Dim r As Integer, c As Integer, CatCell As Range, DataCell As Range
Dim Color As Long
Dim ColorDic As Scripting.Dictionary
Dim Category As Variant
Dim txt As String
Dim GlobalTaskTime As TaskTime
Dim CurrentTaskTime As TaskTime
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
    For Each CatCell In CategoryReferences
      ColorDic.Add CatCell.Interior.Color, CatCell.Value
    Next CatCell
    For r = 1 To PeriodRange.Rows.Count
        If PeriodRange.Cells(r, 1).Value = PeriodValue Then
            For c = 1 To DataRange.Columns.Count
                Set DataCell = DataRange.Cells(r, c)
                Color = DataCell.Interior.Color
                If ColorDic.Exists(Color) Then
                    If DatesRange Is Nothing Then
                        GlobalTaskTime.Increment SubTaskName:=ColorDic(Color), SubSubTaskName:=DataCell.Value, RecurseLevels:=RecurseLevels, ByWeekDay:=ByWeekDay
                    Else
                        GlobalTaskTime.Increment SubTaskName:=ColorDic(Color), SubSubTaskName:=DataCell.Value, RecurseLevels:=RecurseLevels, ByWeekDay:=ByWeekDay, TTime:=DatesRange.Cells(r, 1).Value, WithDateRangeBounds:=WithDateRangeBounds
                    End If
                End If
            Next c
        End If
      ' Debug.Print r
    Next r
    TextReport = GlobalTaskTime.ToString(WithDateRangeBounds, AsArray, OrderBy)

    If TextReport = "" Then TextReport = "No data"
End Function

Function TryFillWithPreviousTask(rng As Range) As Boolean
    Static refColors As Variant
    If IsEmpty(refColors) Then
        Dim r As Integer, TasksRefFullRange As Range
        Set TasksRefFullRange = InputSheet.Range("TasksRefFullRange")
        refColors = Array()
        ReDim refColors(1 To TasksRefFullRange.Rows.Count)
        For r = 1 To TasksRefFullRange.Rows.Count
            refColors(r) = TasksRefFullRange.Cells(r, 1).Interior.Color
        Next r
    End If
    Dim c As Integer, cellColor As Long, taskColor As Long, i As Integer
    cellColor = rng.Interior.Color
    If ArrayContains(refColors, cellColor) Then Exit Function
    c = -1
    While rng.Offset(0, c).Interior.Color = cellColor
        c = c - 1
    Wend
    taskColor = rng.Offset(0, c).Interior.Color
    If Not ArrayContains(refColors, taskColor) Then Exit Function
    c = 0
    While rng.Offset(0, c).Interior.Color = cellColor
        rng.Offset(0, c).Interior.Color = taskColor
        c = c - 1
    Wend
    TryFillWithPreviousTask = True
End Function

Function RangeRelation(r1 As Range, r2 As Range) As String
Dim hRelation As String
Dim vRelation As String
    hRelation = IntervalRelation(r1.Column, r1.Column + r1.Columns.Count, r2.Column, r2.Column + r2.Columns.Count)
    vRelation = IntervalRelation(r1.row, r1.row + r1.Rows.Count, r2.row, r2.row + r2.Rows.Count)
    If hRelation = vRelation Then
        RangeRelation = vRelation
    Else
        RangeRelation = "Disjointed"
    End If
End Function
Function IntervalRelation(x1 As Long, x2 As Long, y1 As Long, y2 As Long) As String
    If x2 < y1 Then
        IntervalRelation = "Disjointed"
    ElseIf y2 < x1 Then
        IntervalRelation = "Disjointed"
    ElseIf x1 <= y1 And x2 >= y2 Then
        IntervalRelation = "Including"
    ElseIf y1 <= x1 And y2 >= x2 Then
        IntervalRelation = "Included"
    Else
        IntervalRelation = "Overlapping"
    End If
End Function