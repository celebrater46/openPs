' https://soypocket.com/it/vbs-powershell-silent-install/

' test.ps1 is working. 
' but it is killed by something when opening by VBS.
' perhaps the cause is Security Software

Dim objFso, objWsh, vbsPath, vbsFolder, ps1Path
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objWsh = WScript.CreateObject("WScript.Shell")
vbsPath = WScript.ScriptFullName
vbsFolder = objFso.GetFile(vbsPath).ParentFolder
ps1Path = vbsFolder & "\test.ps1"
' WScript.echo ps1Path
Const OPT = "powershell -executionpolicy unrestricted -noexit"

objWsh.Run OPT & ps1Path,1,true

' wscript.sleep(3000)

Set objFso = Nothing
Set objWsh = Nothing

Wscript.Quit