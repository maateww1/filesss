Set wshShell = CreateObject("WScript.Shell")
x = 0
do 
WScript.Sleep(60000)
strName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set MyFile = fso.CreateTextFile("C:\Users\" + strName +"\Desktop\NoBooty"& x &".b", True)
   MyFile.WriteLine("Shit Mayn Check Out This Dope Shit")
   MyFile.Close
x = x + 1
loop
