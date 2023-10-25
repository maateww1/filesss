

Dim x
x=0
Do While x<100
Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys(chr(&hAF))
x=x+1
Loop




