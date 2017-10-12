Attribute VB_Name = "ArrayHelpers"
Function ArrayDim(ByVal v As Variant) As Integer
Dim d As Integer
    If Not IsArray(v) Then Exit Function
    On Error GoTo ExitFct:
    
    While True
        d = UBound(v, ArrayDim + 1)
        ArrayDim = ArrayDim + 1
    Wend
    
ExitFct:

End Function

Function Concat(ByVal V1 As Variant, ByVal V2 As Variant) As Variant
Dim Ad1 As Integer, Ad2 As Integer
Dim a As Variant, i As Integer, K As Integer, Ix As Integer
Ad1 = ArrayDim(V1)
Ad2 = ArrayDim(V2)
If Ad1 <> Ad2 Then
    Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays Must Have Same Number Dimensions"
End If
Select Case Ad1
    Case 0
        Concat = Array(V1, V2)
    Case 1
        a = Array()
        ReDim a(UBound(V1) - LBound(V1) + UBound(V2) - LBound(V2) + 1)
        For i = LBound(V1) To UBound(V1): a(Ix) = V1(i): Ix = Ix + 1: Next i
        For i = LBound(V2) To UBound(V2): a(Ix) = V2(i): Ix = Ix + 1: Next i
    Case 2
        If Not (LBound(V1, 2) = LBound(V2, 2) And UBound(V1, 2) = UBound(V2, 2)) Then
            Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays' Second Dimension Does Not Match"
        End If
        a = Array()
        ReDim a(UBound(V1) - LBound(V1) + UBound(V2) - LBound(V2) + 1, LBound(V1, 2) To UBound(V1, 2))
        For i = LBound(V1, 1) To UBound(V1, 1): For j = LBound(V1, 2) To UBound(V1, 2): a(Ix, j) = V1(i, j): Ix = Ix + 1: Next j: Next i
        For i = LBound(V2, 1) To UBound(V2, 1): For j = LBound(V2, 2) To UBound(V2, 2): a(Ix, j) = V2(i, j): Ix = Ix + 1: Next j: Next i
    Case 3
        If Not ( _
                    LBound(V1, 2) = LBound(V2, 2) And UBound(V1, 2) = UBound(V2, 2) _
            And LBound(V1, 3) = LBound(V2, 3) And UBound(V1, 3) = UBound(V2, 3) _
        ) Then
            Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays' Second Or Third Dimension Do Not Match"
        End If
        a = Array()
        ReDim a(UBound(V1) - LBound(V1) + UBound(V2) - LBound(V2) + 1, LBound(V1, 2) To UBound(V1, 2), LBound(V1, 3) To UBound(V1, 3))
        For i = LBound(V1, 1) To UBound(V1, 1): For j = LBound(V1, 2) To UBound(V1, 2): For K = LBound(V1, 3) To UBound(V1, 3): a(Ix, j, K) = V1(i, j, K): Ix = Ix + 1: Next K: Next j: Next i
        For i = LBound(V2, 1) To UBound(V2, 1): For j = LBound(V2, 2) To UBound(V2, 2): For K = LBound(V2, 3) To UBound(V2, 3): a(Ix, j, K) = V2(i, j, K): Ix = Ix + 1: Next K: Next j: Next i
    Case Else
        Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays Of More Than 3 Dimensions Are Not Supported"
End Select

End Function

Public Function FlattenArray(ParamArray Parms() As Variant) As Variant
Dim a As Variant, i As Integer, P As Integer, Pp As Integer
Dim b As Variant, Pa As Variant
    Pa = Parms
    While LBound(Pa) = UBound(Pa)
        Pa = Pa(UBound(Pa))
    Wend
    a = Array()
    For P = LBound(Pa) To UBound(Pa)
        If IsArray(Pa(P)) Then
            b = FlattenArray(Pa(P))
            ReDim Preserve a(i + UBound(b) - UBound(b) + 1)
            For Pp = LBound(b) To UBound(b)
                a(i) = b(Pp)
                i = i + 1
            Next Pp
        Else
            ReDim Preserve a(i)
            a(i) = Pa(P)
            i = i + 1
        End If
    Next P
    FlattenArray = a
End Function

' Make Variant(Height, Width) From Variant(Height)(Width)
Public Function Make2DArray(ByVal i As Variant, Optional ByVal Width As Integer = -1) As Variant
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
    Logger.Error "Error {0} In {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Make2DArray"
    Err.Raise Logger.ErrNumber, Logger.ErrSource, Logger.ErrDescription, Logger.ErrHelpFile, Logger.ErrHelpContext
    Resume
    Exit Function
Proc:
Dim r As Integer, c As Integer, a() As Variant
    If Not IsArray(i) Then Exit Function
    If UBound(i) = -1 Then Exit Function
    If Not IsArray(i(LBound(i))) Then
        Width = 0
    ElseIf Width = -1 Then
        Width = UBound(i(LBound(i)))
    End If
    ReDim a(LBound(i) To UBound(i), Width)
    For r = LBound(i) To UBound(i)
        For c = 0 To Width
            If IsArray(i(r)) Then
                If c <= UBound(i(r)) Then
                    a(r, c) = i(r)(c)
                End If
            Else
                a(r, c) = i(r)
            End If
        Next c
    Next r
    Make2DArray = a
End Function

Sub TestArrayDim()
    Dim s As Integer
    Dim V1(4) As Integer
    Dim V2(4, 4) As Integer
    Dim V3(4, 4, 4) As Integer
    Dim V4(4, 4, 4, 4) As Integer
    
    Debug.Print ArrayDim(s)
    Debug.Print ArrayDim(V1)
    Debug.Print ArrayDim(V2)
    Debug.Print ArrayDim(V3)
    Debug.Print ArrayDim(V4)
End Sub

Sub TestForEach()
    Dim a, E, i
    a = Array(1, 2, 3)
    Debug.Print "A=[";: For Each E In a: Debug.Print E; ",";: Next E: Debug.Print "]"
    For Each E In a: E = E * 2: Next E
    Debug.Print "A=[";: For Each E In a: Debug.Print E; ",";: Next E: Debug.Print "]"
    For i = LBound(a) To UBound(a): a(i) = a(i) * 2: Next i
    Debug.Print "A=[";: For Each E In a: Debug.Print E; ",";: Next E: Debug.Print "]"
End Sub

Function ArrayToString(a As Variant) As String
Dim i As Integer, j As Integer
Dim results() As String, r As Integer
    Select Case ArrayDim(a)
        Case 0: ArrayToString = "Array()"
        Case 1:
            For i = LBound(a) To UBound(a)
                ArrayToString = IIf(ArrayToString = "", "Array(", ArrayToString & ", ")
                If ArrayDim(a(i)) = 0 Then
                    ArrayToString = ArrayToString & CStr(a(i))
                Else
                    ArrayToString = ArrayToString & ArrayToString(a(i))
                End If
            Next i
            ArrayToString = ArrayToString & ")"
        Case 2:
            For i = LBound(a, 1) To UBound(a, 1)
                ArrayToString = IIf(ArrayToString = "", "Array(", ArrayToString & ", ")
                For j = LBound(a, 2) To UBound(a, 2)
                    ArrayToString = ArrayToString & IIf(j = LBound(a, 2), "( ", ", ")
                    If ArrayDim(a(i, i)) = 0 Then
                        ArrayToString = ArrayToString & CStr(a(i, j))
                    Else
                        ArrayToString = ArrayToString(a(i, i))
                    End If
                Next j
                ArrayToString = ArrayToString & ")"
            Next i
            ArrayToString = ArrayToString & ")"
        Case Else
            Err.Raise vbObjectError, "ArrayHelpers.ArrayToString", "3-dim arrays or higher not supported"
    End Select
End Function


Sub testArrayToString()
    Dim a(2, 3), i, j
    For i = LBound(a, 1) To UBound(a, 1)
        For j = LBound(a, 2) To UBound(a, 2)
            a(i, j) = i & "-" & j
        Next j
    Next i
    Debug.Print ArrayToString(a)
End Sub

Public Sub QuickSort(arr, Optional Lo As Long = -1, Optional Hi As Long = -1)
On Error GoTo Err_Proc
GoTo Proc
Err_Proc:
  If vbYes = MsgBox(Err.Description & vbCrLf & "Debug?", vbYesNo Or vbCritical, "Error") Then
    Stop
    Resume
  End If
  Exit Sub
Proc:
' sorts a 2-dimensional array on 1st dimension
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  Lo = IIf(Lo >= 0, Lo, LBound(arr))
  Hi = IIf(Hi >= 0, Hi, UBound(arr))
  tmpLow = Lo
  tmpHi = Hi
  varPivot = arr((tmpLow + tmpHi) \ 2)(0)
  Do While tmpLow <= tmpHi
    Do While arr(tmpLow)(0) < varPivot And tmpLow < Hi
      tmpLow = tmpLow + 1
    Loop
    Do While varPivot < arr(tmpHi)(0) And tmpHi > Lo
      tmpHi = tmpHi - 1
    Loop
    If tmpLow <= tmpHi Then
      varTmp = arr(tmpLow)
      arr(tmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
  If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
End Sub

Public Function ArrayContains(a As Variant, value As Variant) As Boolean
    If Not IsArray(a) Then Exit Function
    Dim v As Variant
    For Each v In a
        If v = value Then
            ArrayContains = True
            Exit Function
        End If
    Next v
End Function


