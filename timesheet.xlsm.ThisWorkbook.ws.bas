VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Activate()
    GoNow
End Sub

Public Sub GoNow()
Dim Row As Integer, col As Integer
    Row = Int(Now - InputSheet.Range("Dates").Cells(1, 1).value)
    If Row < 0 Then Exit Sub
    col = 3 + Int((Now - Int(Now)) * 24 * 4)
    On Error Resume Next
    InputSheet.Select
    InputSheet.Cells(InputSheet.Range("Dates").Row + Row, col).Select
End Sub
