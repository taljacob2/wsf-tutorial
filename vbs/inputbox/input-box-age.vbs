Option Explicit

Dim age
age = InputBox("What is your age?", "Title")

if Not IsNumeric(age) Or age = "" then
    MsgBox "Age is not valid. Please try again."
else
    MsgBox "Age is a valid number. Your age is : " &age
end if
