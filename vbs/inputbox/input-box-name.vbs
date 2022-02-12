Option Explicit

Dim name
name = InputBox("What is your name?", "Title", "Type your name here")

MsgBox name, vbOKOnly, "This is your name"

' if name = "tal" then
'     MsgBox "Hi Developer!", vbOKOnly
' elseif name <> "tal" then ' NOTE: the "<>" stands for "not equal".
'     MsgBox "You are not tal!", vbOKOnly
' end if


' if name = "tal" then
'     MsgBox "Hi Developer!", vbOKOnly
' elseif Not name = "tal" then ' NOTE: the "Not" stands for "not equal".
'     MsgBox "You are not tal!", vbOKOnly
' end if


' if name = "" then
'     MsgBox "Please input something, and try again. (and do not press the 'cancel' button)", vbOKOnly
' elseif name = "tal" then
'     MsgBox "Hi Developer!", vbOKOnly
' elseif Not name = "tal" then ' NOTE: the "Not" stands for "not equal".
'     MsgBox "You are not tal!", vbOKOnly
' end if


if name = "" then
    MsgBox "Please input something, and try again."_
            &" (and do not press the 'cancel' button)", vbOKOnly
elseif name = "tal" then
    MsgBox "Hi Developer!", vbOKOnly
elseif Not name = "tal" then ' NOTE: the "Not" stands for "not equal".
    MsgBox "You are not tal!", vbOKOnly
end if

