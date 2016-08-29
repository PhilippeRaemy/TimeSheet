Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.ChartObjects("PieChartWeekly").Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).Points(16).Select
    ActiveChart.ChartGroups(1).FirstSliceAngle = 355
    ActiveChart.ChartGroups(1).FirstSliceAngle = 355
    Selection.Format.Fill.Visible = msoFalse
End Sub
