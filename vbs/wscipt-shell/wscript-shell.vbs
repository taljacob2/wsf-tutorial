Option Explicit

Dim wshell
Set wshell = CreateObject("wscript.shell") 

Dim result
result = MsgBox("Do you want to open paint?", vbYesNo)
if result = vbYes then
    wshell.Run("mspaint.exe")
else
    MsgBox ("We open GoogleChrome then...")
    wshell.Run("""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk""")
end if
