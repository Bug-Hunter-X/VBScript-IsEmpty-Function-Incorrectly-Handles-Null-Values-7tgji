Function f(a, b)
  If IsEmpty(a) Then
    Exit Function 'This line causes an error if a is Null'
  End If
  'Rest of the function code
End Function