Function f(a)
  If IsEmpty(a) Then Exit Function
  If InStr(1, a, ";") > 0 Then
    Err.Raise 1001, , "Semicolons are not allowed!" 
  End If
end function