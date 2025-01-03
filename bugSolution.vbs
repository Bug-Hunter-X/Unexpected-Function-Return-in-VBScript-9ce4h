Function f(x)
  Dim result
  If x > 10 Then
    result = x + 1
  Else
    result = x - 1
  End If
  f = result
End Function
MsgBox f(2)