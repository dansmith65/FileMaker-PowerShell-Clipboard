'prevents a window from opening when the script is run
'http://www.leporelo.eu/blog.aspx?id=run-scheduled-tasks-with-winform-gui-in-powershell

Dim objShell
Set objShell=CreateObject("WScript.Shell")

filePath = Wscript.Arguments(0)
param1 = Wscript.Arguments(1)

strCMD="powershell -sta -noProfile -NonInteractive -nologo -file """ & filePath & """ " & param1

objShell.Run strCMD,0