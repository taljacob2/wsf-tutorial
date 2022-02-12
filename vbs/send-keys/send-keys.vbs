Option Explicit

Dim wshell
Set wshell = CreateObject("wscript.shell")

wshell.Run("notepad.exe")
wscript.Sleep 2000 ' Sleep main thread for 2 seconds.
wshell.SendKeys "Hello World"
wshell.SendKeys "{ENTER}"
wshell.SendKeys "How are you?"
