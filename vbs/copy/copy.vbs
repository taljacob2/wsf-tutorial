Option Explicit

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile "C:\Users\Tal\Desktop\*.url", "C:\Users\Tal\TEST\", True
WScript.Echo "Copied."
