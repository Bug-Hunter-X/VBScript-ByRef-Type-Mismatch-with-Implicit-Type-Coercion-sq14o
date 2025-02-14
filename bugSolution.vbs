Function f(ByRef a)
  If IsNumeric(a) Then
    a = a + 1
  ElseIf VarType(a) = vbString Then
    a = a & "1"
  Else
    Err.Raise 13, , "Type mismatch: Parameter must be numeric or string."
  End If
end function

x = 10
f x
MsgBox x 'Outputs 11

x = "hello"
f x
MsgBox x 'Outputs "hello1"

x = true
' this will raise the error
f x