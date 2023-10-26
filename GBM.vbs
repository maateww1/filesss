Set wshShell = WScript.CreateObject("WScript.Shell")
sUserName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
sWinDir = oFSO.GetSpecialFolder(0)
sWallPaper = "%appdata%\gucci.png"
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True
