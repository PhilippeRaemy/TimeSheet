Attribute VB_Name = "TimeSheet"
Option Explicit

Sub CmdTask(TaskName As Range)
    Dim sel As Range, cell As Range, r As Integer
    Dim SelectedColor As Long
    Set sel = Application.Selection
    SelectedColor = sel.Cells(1, 1).Interior.color
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Select Case sel.Cells(1, 1).Interior.color
        Case TaskName.Interior.color ' override whatever is selected
            sel.Interior.color = TaskName.Interior.color
            sel.Interior.Pattern = TaskName.Interior.Pattern
            sel.Font.color = TaskName.Font.color
        Case SelectedColor ' only fill-in gaps in the selection
            For Each cell In sel.Cells
                If cell.Interior.color = SelectedColor Then
                    cell.Interior.color = TaskName.Interior.color
                    cell.Interior.Pattern = TaskName.Interior.Pattern
                    cell.Font.color = TaskName.Font.color
                End If
            Next cell
    End Select
    For r = 1 To sel.Rows.Count
        sel.Cells(r, 1).value = sel.Cells(r, 1).value
    Next r
    Application.Calculation = xlCalculationAutomatic
    PivotSheet.PivotTables("WeeklyAggregates").PivotCache.Refresh
    PivotSheet.PivotTables("MonthlyAggregates").PivotCache.Refresh
    Application.ScreenUpdating = True
End Sub

Sub SetSummaries(Target As Range)
Dim DayDate As Date
    If Not RangeRelation(InputSheet.Range("InputRange"), Target) = "Including" Then Exit Sub
    
    DayDate = InputSheet.Range("Dates").Cells(Target.Row - InputSheet.Range("Dates").Row + 1, 1).value
    
    Dim SummaryWeek As Date:    SummaryWeek = DateAdd("d", 1 - DatePart("w", DayDate, vbMonday), DayDate)
    Dim SummaryMonth As Date:   SummaryMonth = DateAdd("d", 1 - day(DayDate), DayDate)
    Dim SummaryYear As Integer: SummaryYear = Year(DayDate)
    ' check values before to apply to save triggering unnecessary recalculation
    SetSummary DayDate, "SummaryDay", ""
    SetSummary SummaryWeek, "SummaryWeek", "PieChartWeekly"
    SetSummary SummaryMonth, "SummaryMonth", "PieChartMonthly"
    SetSummary SummaryMonth, "SummaryMonth", "PieChartYearly"
End Sub
Sub SetSummary(value As Variant, RangeName As String, PieChartName As String)
    If InputSheet.Range(RangeName).value <> value Then
        ' Debug.Print "SetSummary:" & RangeName, PieChartName, value
        InputSheet.Range(RangeName).value = value
        If PieChartName <> "" Then FormatPieChartByName PieChartName
        If PieChartName = "PieChartMonthly" Then
            FormatPieChartByName "PieChartYearly"
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

    labels.Format.TextFrame2.TextRange.Font.Size = 7

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
        If cell.Interior.color = ref.Interior.color Then CountColored = CountColored + 1
    Next cell
End Function
Public Function TimeRanges(rng As Range, WorkColorRef As Range) As String
Static WorkColors() As Long
Static inited As String
Dim IsWork As Boolean
Dim r As Integer, c As Integer, i As Integer
Dim cell As Range
Dim color As Long: color = -1
Dim cellColor As Long
Dim isWorkColor As Boolean


    If Not inited = WorkColorRef.Address Then
        ReDim WorkColors(WorkColorRef.Rows.Count * WorkColorRef.Columns.Count - 1)
        For Each cell In WorkColorRef
            WorkColors(i) = cell.Interior.color
            i = i + 1
        Next cell
        inited = WorkColorRef.Address
    End If
    
    For c = 1 To rng.Columns.Count
        cellColor = rng.Cells(1, c).Interior.color
        If Not color = cellColor Then
            color = cellColor
            isWorkColor = False
            For i = LBound(WorkColors) To UBound(WorkColors)
                If color = WorkColors(i) Then
                    isWorkColor = True
                    ' Debug.Print c, color, isWorkColor, "HIT on color", i, WorkColors(i)
                    Exit For
                End If
            Next i
            ' Debug.Print c, color, isWorkColor
            If IsWork And Not isWorkColor Then
                TimeRanges = TimeRanges & "-" & Format((c - 1) / rng.Columns.Count, "hh:mm")
                IsWork = False
            ElseIf Not IsWork And isWorkColor Then
                TimeRanges = IIf(TimeRanges = "", "", TimeRanges & ", ") & Format((c - 1) / rng.Columns.Count, "hh:mm")
                IsWork = True
            End If
        End If
    Next c
    If IsWork Then
      TimeRanges = TimeRanges & "-24:00"
      IsWork = False
    End If
End Function

'Public Function TaskList(DataRange As Range, PeriodRange As Range, PeriodValue As Variant, CategoryReferences As Range, Title As String) As Variant
'    Dim a As Variant, b As Variant, i As Integer
'    a = Split(TextReport(DataRange, PeriodRange, PeriodValue, CategoryReferences, Title, DatesRange:=Nothing, AsArray:=True), "|")
'    b = Array()
'    ReDim b(UBound(a), 0)
'    For i = 0 To UBound(a)
'        b(i, 0) = a(i)
'    Next i
'    TaskList = b
'End Function

Function TryFillWithPreviousTask(rng As Range) As Boolean
    Static refColors As Variant
    If IsEmpty(refColors) Then
        Dim r As Integer, TasksRefFullRange As Range
        Set TasksRefFullRange = InputSheet.Range("TasksRefFullRange")
        refColors = Array()
        ReDim refColors(1 To TasksRefFullRange.Rows.Count)
        For r = 1 To TasksRefFullRange.Rows.Count
            refColors(r) = TasksRefFullRange.Cells(r, 1).Interior.color
        Next r
    End If
    Dim c As Integer, cellColor As Long, taskColor As Long, i As Integer
    cellColor = rng.Interior.color
    If ArrayContains(refColors, cellColor) Then Exit Function
    c = -1
    While rng.Offset(0, c).Interior.color = cellColor
        c = c - 1
    Wend
    taskColor = rng.Offset(0, c).Interior.color
    If Not ArrayContains(refColors, taskColor) Then Exit Function
    c = 0
    While rng.Offset(0, c).Interior.color = cellColor
        rng.Offset(0, c).Interior.color = taskColor
        c = c - 1
    Wend
    TryFillWithPreviousTask = True
End Function

Function RangeRelation(r1 As Range, r2 As Range) As String
Dim hRelation As String
Dim vRelation As String
    hRelation = IntervalRelation(r1.Column, r1.Column + r1.Columns.Count, r2.Column, r2.Column + r2.Columns.Count)
    vRelation = IntervalRelation(r1.Row, r1.Row + r1.Rows.Count, r2.Row, r2.Row + r2.Rows.Count)
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

Sub ButtonImportAppointments_Click()
    ImportAppointments InputSheet.Range("SummaryDay").value
End Sub

Public Sub ImportAppointments(day As Date)
    Dim appts() As Appointment
    Dim appt As Appointment
    Dim vappt As Variant 'Appointment
    Dim InputRange As Range
    Set InputRange = InputSheet.Range("InputRange")
    Dim StartCell As Range
    Dim EndCell As Range
    Dim cell As Range
    Dim TaskCell As Range
    Dim tasks As String
    
    If Not RangeRelation(InputRange, Selection) = "Including" Then Exit Sub
    
    appts = OutlookAccess.FindAppts(day, day + 1)
    On Error Resume Next
    If UBound(appts) < 0 Then Exit Sub
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    For Each cell In InputSheet.Range("TasksRefFullRange").Columns(2).Cells
        tasks = IIf(tasks = "", "Tasks: ", tasks & ", ") & cell.value
    Next cell
    
    For Each vappt In appts
        Set appt = vappt
        If Not Left(appt.Subject, 8) = "Canceled" Then
            Set StartCell = InputRange.Cells(1 + Selection.Row - InputRange.Row, appt.StartTick(1 / 4))
            Set EndCell = InputRange.Cells(1 + Selection.Row - InputRange.Row, appt.EndTick(1 / 4))
            InputSheet.Range(StartCell, EndCell).Select
            Set TaskCell = Nothing
            For Each cell In InputSheet.Range("TasksRefFullRange")
                If InStr(1, appt.Subject, cell.value & ":", vbTextCompare) <> 0 _
                Or InStr(1, cell.value & ":", appt.Subject, vbTextCompare) <> 0 Then
                    Set TaskCell = cell
                    Exit For
                End If
            Next cell
            If TaskCell Is Nothing Then
                Dim task As String: task = InputBox(appt.Subject & vbCrLf & "-------" & vbCrLf & tasks, "Please pick a task", "ADM")
                If task <> "" Then
                    For Each cell In InputSheet.Range("TasksRefFullRange").Cells
                        Debug.Print task, cell.Address, cell.value, InStr(1, task, cell.value, vbTextCompare), InStr(1, cell.value, task, vbTextCompare)
                        If InStr(1, task, cell.value, vbTextCompare) <> 0 _
                        Or InStr(1, cell.value, task, vbTextCompare) <> 0 Then
                            Set TaskCell = cell
                            Exit For
                        End If
                    Next cell
                    If TaskCell Is Nothing Then
                        Set TaskCell = InputSheet.Range("DefaultAdminPattern")
                    End If
                End If
            End If
            If Not TaskCell Is Nothing Then
                If IsEmpty(StartCell.value) Then StartCell.value = appt.Subject
                CmdTask TaskCell
                Selection.BorderAround XlLineStyle.xlDot, xlThick, color:=RGB(255, 0, 0)
                Debug.Print appt.ToString(), appt.StartTick(1 / 4), appt.EndTick(1 / 4)
            End If
        End If
    Next vappt
End Sub
Public Sub WorkReportStarter()
    Debug.Print WorkReport("MonthlyAggregates", "mmm-yy", "ADM", "Hol", "*")
End Sub
Public Function WorkReport(PivotTableName As String, dateformat As String, ParamArray Categories() As Variant) As String
    Dim p As PivotTable, Pivot As PivotTable
    Dim Header As Range
    Dim DataRows As Variant, DataRow As Variant, TotalDataRow As Variant
    Dim r As Integer, c As Integer, cc As Integer
    Dim widths() As Integer
     
    For Each p In PivotSheet.PivotTables
        If p.Name = PivotTableName Then
            Set Pivot = p
            Exit For
        End If
    Next p
    If Pivot Is Nothing Then Exit Function
    Set Header = Pivot.ColumnRange.Rows(Pivot.ColumnRange.Rows.Count)
    DataRows = Array()
    DataRow = Array()
    TotalDataRow = Array()
    
    Dim TotalCol As Integer: TotalCol = 2 + UBound(Categories)
    Dim TotalRow As Integer: TotalRow = 1 + Pivot.DataBodyRange.Rows.Count
    
    ReDim DataRows(0 To TotalRow)
    ReDim DataRow(0 To TotalCol)
    ReDim TotalDataRow(0 To TotalCol)
    DataRow(TotalCol) = "Total"
    DataRow(0) = CStr(Pivot.RowRange.Cells(1, 1).value)
    For cc = 0 To UBound(Categories)
        DataRow(cc + 1) = IIf(CStr(Categories(cc)) = "*", "Other", CStr(Categories(cc)))
    Next cc
    TotalDataRow(0) = "Total"
    For cc = 1 To TotalCol
        TotalDataRow(cc) = 0
    Next cc
    
    DataRows(0) = DataRow
    
    For r = 0 To Pivot.DataBodyRange.Rows.Count - 1
        DataRow = Array()
        ReDim DataRow(0 To TotalCol)
        DataRow(0) = Format(Pivot.RowRange.Cells(r + 2, 1), dateformat)
        For c = 1 To Pivot.ColumnRange.Columns.Count
            For cc = 1 To 1 + UBound(Categories)
                If LCase(Trim(Header.Cells(1, c).value)) Like LCase(Trim(Categories(cc - 1))) Then
                    If IsEmpty(DataRow(cc)) Then DataRow(cc) = 0
                    DataRow(cc) = DataRow(cc) + Pivot.DataBodyRange.Cells(r + 1, c).value
                    Exit For
                End If
            Next cc
        Next c
        DataRow(TotalCol) = 0
        For cc = 1 To TotalCol - 1
            DataRow(TotalCol) = DataRow(TotalCol) + DataRow(cc)
            TotalDataRow(cc) = TotalDataRow(cc) + DataRow(cc)
            DataRow(cc) = Application.text(DataRow(cc), "[h]:mm")
        Next cc
        TotalDataRow(TotalCol) = TotalDataRow(TotalCol) + DataRow(TotalCol)
        DataRow(UBound(DataRow)) = Application.text(DataRow(UBound(DataRow)), "[h]:mm")
        DataRows(r + 1) = DataRow
    Next r
    For cc = 0 To TotalCol
        TotalDataRow(cc) = Application.text(TotalDataRow(cc), "[h]:mm")
    Next cc
    DataRows(TotalRow) = TotalDataRow

    ReDim widths(TotalCol)
    For cc = 0 To TotalCol
        For r = LBound(DataRows) To UBound(DataRows)
            If Len(DataRows(r)(cc)) > widths(cc) Then widths(cc) = Len(DataRows(r)(cc))
        Next r
    Next cc
    WorkReport = WorkReport & vbCrLf
    Dim separator As String
    For c = 0 To UBound(widths)
        separator = separator & String(widths(c) + 1, "-") & "-|"
    Next c
    separator = separator & vbCrLf
    For r = 0 To TotalRow
        For c = 0 To UBound(DataRows(r))
            WorkReport = WorkReport & String(widths(c) + 1 - Len(DataRows(r)(c)), " ") & DataRows(r)(c) & " |"
        Next c
        WorkReport = WorkReport & vbCrLf
        If r = 0 Or r = TotalRow - 1 Then
            WorkReport = WorkReport & separator
        End If
    Next r
    WorkReport = WorkReport & separator
    
End Function



