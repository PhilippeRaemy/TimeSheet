VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTaskReport 
   Caption         =   "Task Report"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17115
   OleObjectBlob   =   "timesheet.xlsm.FrmTaskReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTaskReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()
    UserForm_Resize
End Sub

Private Sub UserForm_Resize()
    ' Me.CmdOkCancel.Move Me.InsideWidth - Me.CmdOkCancel.Width - Me.TextBox.Left, Me.InsideHeight - Me.CmdOkCancel.Height - Me.TextBox.Top
    Me.TextBox.Width = Me.InsideWidth - 2 * Me.TextBox.Left
    Me.TextBox.Height = Me.InsideHeight - 2 * Me.TextBox.Top
End Sub

Public Sub ShowMessage(text As String, Title As String)
    Me.TextBox.text = text
    Me.Caption = Title
    Me.Show ' True
End Sub
