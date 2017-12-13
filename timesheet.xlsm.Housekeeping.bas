Attribute VB_Name = "Housekeeping"
Option Explicit
Sub Housekeeping()
Dim sh As Shape
Dim Target As Range, targetAddress As String
Dim positionTarget As Range
Dim i As Integer
Const TaskButtonName = "TaskButton_"
    For Each sh In InputSheet.Shapes
        Debug.Print sh.Name, TypeName(sh), sh.OnAction
        If sh.Name Like "Button*" And Not sh.OnAction Like "'*'" Then sh.Delete
        If (sh.Name Like "TaskButton*" Or sh.Name Like "Button*") And sh.OnAction Like "'*'" Then
            targetAddress = Split(sh.OnAction, """")(1)
            Set Target = InputSheet.Range(targetAddress)
            Set positionTarget = InputSheet.Cells(Target.Row, Target.Column + 1)
            If Not sh.Name = TaskButtonName & targetAddress Then
              sh.Name = TaskButtonName & targetAddress
            End If
            sh.Left = positionTarget.Left + positionTarget.Width - positionTarget.Height * 1.3
            sh.Top = positionTarget.Top
            sh.Height = positionTarget.Height
            sh.Width = positionTarget.Height * 1.2
          sh.Fill.ForeColor.RGB = Target.Interior.Color
        End If
    Next sh
    For i = 1 To InputSheet.Range("TasksRef").Rows.Count + 1
        Set Target = InputSheet.Range("TasksRef").Cells(i, 1)
        Set sh = InputSheet.Shapes(TaskButtonName & Replace(Target.Address, "$", ""))
        sh.ZOrder msoBringToFront
        '  sh.Fill.BackColor.RGB = target.Interior.Color ' does not work for buttons :(
    Next i
  
End Sub
Sub ResetColors()
Dim Cell As Range, r As Integer
Dim trg As Range, trgName As Variant
Dim ch As ChartObject, sc As Series, sp As Point
Dim ChartHeaderRangeAddress As String
Dim ChartHeaderRange As Range
    Application.Calculation = xlCalculationManual
    For r = 1 To InputSheet.Range("TasksRef").Rows.Count
        Set Cell = InputSheet.Range("TasksRef").Cells(r, 1)
        For Each trgName In Array("DailyTotals", "DailyTotalsForPivot")
            Set trg = ActiveWorkbook.Names(trgName).RefersToRange
            ' trg.Cells(1, r).Value = cell.Value
            trg.Columns(r).Interior.Color = Cell.Interior.Color
        Next trgName
    Next r
    For Each ch In InputSheet.ChartObjects
        If ch.Chart.SeriesCollection.Count = 1 Then
            Debug.Print ch.Name, ch.Left, ch.Top
            ch.Top = Range("PieChartRange").Top
            ch.Height = Range("PieChartRange").Height
            ch.Left = Range("PieChartRange").Left + IIf(ch.Name = "PieChartYearly", 0, IIf(ch.Name = "PieChartMonthly", 1, IIf(ch.Name = "PieChartWeekly", 2, 3))) * Range("PieChartRange").Width / 3
            ch.Width = Range("PieChartRange").Width / 3
            ch.Chart.ChartArea.Left = ch.Left
            ch.Chart.ChartArea.Top = ch.Top
            ch.Chart.ChartArea.Width = ch.Width - 5
            ch.Chart.ChartArea.Height = ch.Height - 5
            
            
            Set sc = ch.Chart.SeriesCollection(1)
            ChartHeaderRangeAddress = Split(sc.Formula, ",")(1)
            Set ChartHeaderRange = ActiveWorkbook.Worksheets(Split(ChartHeaderRangeAddress, "!")(0)).Range(Split(ChartHeaderRangeAddress, "!")(1))
            r = 0
            For Each Cell In ChartHeaderRange.Rows(1).Cells
                r = r + 1
                Set sp = sc.Points(r)
                If Cell.Interior.Color = 16777215 Then
                    sp.Format.Fill.Visible = msoFalse
                Else
                    sp.Format.Fill.Visible = msoTrue
                    sp.Format.Fill.ForeColor.RGB = Cell.Interior.Color
                    sp.Format.line.ForeColor.RGB = Cell.Interior.Color
                End If
            Next Cell
        Else
            For Each sc In ch.Chart.SeriesCollection
                On Error Resume Next
                Dim refCell As Range
                Dim TasksRefFullRange As Range
                Set TasksRefFullRange = InputSheet.Range("TasksRefFullRange")
                Set refCell = TasksRefFullRange.Cells(WorksheetFunction.Match(Trim(sc.Name), TasksRefFullRange.Columns(1), 0), 1)
                If refCell Is Nothing Then
                    Set refCell = TasksRefFullRange.Cells(WorksheetFunction.Match(Trim(sc.Name), TasksRefFullRange.Columns(2), 0), 2)
                End If
                On Error GoTo 0
                Debug.Print ch.Name, sc.Name, sc.Formula
                If Not refCell Is Nothing Then
                    sc.Format.Fill.ForeColor.RGB = refCell.Interior.Color
                    sc.Format.line.Visible = False
                    sc.Format.ThreeD.BevelBottomDepth = 6
                    sc.Format.ThreeD.BevelBottomInset = 6
                    sc.Format.ThreeD.BevelBottomType = msoBevelCircle
                    sc.Format.ThreeD.BevelTopDepth = 6
                    sc.Format.ThreeD.BevelTopInset = 6
                    sc.Format.ThreeD.BevelTopType = msoBevelCircle
                End If
             
            Next sc
        End If
    Next ch
    resetWeeklyChartColors
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub ResetPivots()
    ResetPivot "WeeklyAggregates"
    ResetPivot "MonthlyAggregates"
End Sub

Sub ResetPivot(pivotName As String)
Dim Pivot As PivotTable, pField As PivotField, pFields As PivotFields
Set Pivot = PivotSheet.PivotTables(pivotName)
Set pFields = Pivot.PivotFields
Dim space As String
    
    Debug.Print Pivot.PivotFields.Count
    Application.Calculation = xlCalculationManual
    For Each pField In pFields
        Debug.Print pField.Name, pField.Caption, pField.Orientation
        Select Case pField.Name
            Case "Year", "Month", "Week", "Values"
            Case Else
              
              If Not pField.Orientation = xlHidden Then pField.Orientation = xlHidden
        End Select
    Next pField
    For Each pField In pFields
        Debug.Print pField.Name, pField.Caption, pField.Orientation
        Select Case pField.Name
            Case "Year", "Month", "Week", "Values", "End of week"
            Case Else
              space = ""
              pField.Orientation = xlDataField
              pField.Function = xlSum
              On Error Resume Next
              Do
                  pField.Caption = pField.SourceName & space
                  If Err.Number = 0 Then
                      On Error GoTo 0
                      Exit Do
                  End If
                  Err.Clear
                  space = space & " "
              Loop
        End Select
    Next pField
    'For Each pField In pFields
    '  Debug.Print pField.Name, pField.Caption, pField.Orientation
    '  pField.Caption = pField.SourceName
    'Next pField
    resetWeeklyChartColors
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub resetWeeklyChartColors()
    Const bevel = 6
    Dim Chart  As ChartObject
    Dim TaskRefRange As Range, TaskRefCell As Range
    Set TaskRefRange = Range("TasksRefFullRange")
    Dim ser As Series
    For Each Chart In InputSheet.ChartObjects
        If Chart.Chart.SeriesCollection.Count <> 1 Then
            For Each ser In Chart.Chart.SeriesCollection
                Debug.Print Chart.Name, ser.Name,
                For Each TaskRefCell In TaskRefRange.Columns(2).Cells
                    If Trim(TaskRefCell.value) = Trim(ser.Name) Then
                        Debug.Print TaskRefCell.Interior.Color,
                        ser.Format.Fill.ForeColor.RGB = TaskRefCell.Interior.Color
                    End If
                Next TaskRefCell
                ser.Format.ThreeD.BevelTopInset = bevel
                ser.Format.ThreeD.BevelTopDepth = bevel
                ser.Format.ThreeD.BevelBottomInset = bevel
                ser.Format.ThreeD.BevelBottomDepth = bevel
                
                Debug.Print
            Next ser
        End If
    Next Chart
  
End Sub
Sub SetComments()
Dim Cell As Range
    For Each Cell In Selection
        On Error Resume Next
        Cell.AddComment
        Cell.Comment.Visible = False
        Cell.Comment.Shape.TextFrame.Characters.Font.Bold = False
        Cell.Comment.Shape.TextFrame.Characters.Font.Size = 8
        Cell.Comment.Shape.TextFrame.Characters.Font.Name = "Monofur"
        Cell.Comment.Shape.Width = 100
        Cell.Comment.Shape.Height = 12
        Cell.Comment.text text:="Dblclick for summary"
    Next Cell
End Sub

Sub SetCommentsFormat()
    Dim shp As Shape
    On Error Resume Next
    For Each shp In InputSheet.Shapes
        If shp.AutoShapeType = msoShapeRectangle And Not shp.TextFrame Is Nothing Then
            Debug.Print shp.AutoShapeType,
            Debug.Print shp.TextFrame.Characters.text,
            Debug.Print shp.TopLeftCell.Address
            Debug.Print shp.TopLeftCell.Select
            Debug.Print
            shp.TextFrame.Characters.Font.Size = 8
            shp.TextFrame.Characters.Font.Bold = False
            shp.TextFrame.Characters.Font.Name = "Monofur"
        End If
    Next shp
End Sub
