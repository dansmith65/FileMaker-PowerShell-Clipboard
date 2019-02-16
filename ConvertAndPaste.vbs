'Run PowerShell script, then send Ctrl+v if it didn't return an error
'http://www.leporelo.eu/blog.aspx?id=run-scheduled-tasks-with-winform-gui-in-powershell

Set objShell=CreateObject("WScript.Shell")
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strCMD="powershell.exe -sta -noProfile -nologo -file """ & scriptdir & "\Convert-FMClip.ps1" & """ " 
exitCode = objShell.Run(strCMD, 1, true)

If exitCode = 0 Then
	objShell.SendKeys "^v"
Else
	'MsgBox "Exit Code: " & exitCode
End If
