
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("test.txt",1)
x=msgbox(objFileToRead.ReadAll() ,0, "Your Title Here")

objFileToRead.Close
Set objFileToRead = Nothing
