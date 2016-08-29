Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveChart.SeriesCollection(1).ApplyDataLabels
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.ChartObjects("PieChartWeekly").Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.ApplyChartTemplate ( _
        "I:\My Documents\templates\Charts\PieTimesheet.crtx")
End Sub
