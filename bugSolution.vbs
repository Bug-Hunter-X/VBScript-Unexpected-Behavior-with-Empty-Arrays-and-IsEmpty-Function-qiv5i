Function f(a)
  If IsArray(a) Then
    If UBound(a) < 0 Then Exit Function 'Handles empty arrays correctly
  ElseIf IsEmpty(a) Then 
    Exit Function
  End If
  If InStr(1, a, ";") > 0 Then
    Err.Raise 1001, , "Semicolons are not allowed!" 
  End If
end function