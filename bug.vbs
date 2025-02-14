Function f(ByRef a)
  a = a + 1
end function

x = 10
f x
MsgBox x 'Outputs 11, as expected

x = "hello"
f x
MsgBox x 'Throws error: Type mismatch