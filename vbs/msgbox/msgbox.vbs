Option Explicit

Dim result
result = msgbox("Guess a Button", vbAbortRetryIgnore, "Title")

if result = vbRetry Or result = vbIgnore then
    msgbox "retried or ignored"
else
    msgbox "aborting"
end if


' if result = vbRetry then
'     msgbox "retried"
' elseif result = vbIgnore then
'     msgbox "ignored"
' else
'     msgbox "aborting"
' end if

