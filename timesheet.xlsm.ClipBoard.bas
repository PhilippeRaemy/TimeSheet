Attribute VB_Name = "ClipBoard"
Option Explicit
Public Sub SetText(txt As String)
    Dim Clip As New MSForms.DataObject
    If txt = "" Then
        Clip.Clear
    Else
        Clip.SetText txt
    End If
    Clip.PutInClipboard
End Sub

Public Function GetText() As String
    Dim Clip As New MSForms.DataObject
    Clip.GetFromClipboard
    GetText = Clip.GetText
End Function

Public Sub CumulateTextAndPrint(txt As String, ParamArray args() As Variant)
    txt = strings.Format(txt, args)
    Debug.Print txt
    SetText GetText & vbCrLf & txt
End Sub
