' Silent launcher: starts watchdog.pyw, which in turn starts watcher.py and
' relaunches it if it dies silently. Launching watcher.py directly here would
' mean nothing ever restarts it until the next logon.
' Keep this file in <project>/scripts/.
' ASCII only on purpose: WScript reads .vbs as ANSI, so a Chinese literal
' would be corrupted. The path is derived from this file's own location.
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
py = "C:\Users\Administrator\.workbuddy\binaries\python\versions\3.14.3\pythonw.exe"
script = scriptDir & "\watchdog.pyw"
If Not fso.FileExists(py) Then
    WScript.Echo "pythonw.exe not found: " & py
    WScript.Quit 2
End If
If Not fso.FileExists(script) Then
    WScript.Echo "watchdog.pyw not found: " & script
    WScript.Quit 3
End If
If Not fso.FileExists(scriptDir & "\watcher.py") Then
    WScript.Echo "watcher.py not found: " & scriptDir
    WScript.Quit 4
End If
sh.Run """" & py & """ """ & script & """", 0, False
