' https://soypocket.com/it/vbs-powershell-silent-install/

Dim objFso, objWsh, vbsPath, vbsFolder, ps1Path
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objWsh = WScript.CreateObject("WScript.Shell")
vbsPath = WScript.ScriptFullName
vbsFolder = objFso.GetFile(vbsPath).ParentFolder
ps1Path = vbsFolder & "\test.ps1"

Const OPT = "powershell -executionpolicy unrestricted -noexit"

objWsh.Run OPT & ps1Path,0,false

Set objFso = Nothing
Set objWsh = Nothing

Wscript.Quit