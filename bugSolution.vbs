Function IsEmptyOrNull(a)
  If IsEmpty(a) Or a Is Nothing Then
    IsEmptyOrNull = True
  Else
    IsEmptyOrNull = False
  End If
End Function

Function f(a, b)
  If IsEmptyOrNull(a) Then
    Exit Function
  End If
  'Rest of the function code
End Function