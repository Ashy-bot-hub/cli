' Silent launcher for watcher.py (keep this file in <project>/scripts/).
' ASCII only on purpose: WScript reads .vbs as ANSI, so a Chinese literal
' would be corrupted. The path is derived from this file's own location.
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
py = "C:\Users\Administrator\.workbuddy\binaries\python\versions\3.14.3\pythonw.exe"
script = scriptDir & "\watcher.py"
If Not fso.FileExists(py) Then
    WScript.Echo "pythonw.exe not found: " & py
    WScript.Quit 2
End If
If Not fso.FileExists(script) Then
    WScript.Echo "watcher.py not found: " & script
    WScript.Quit 3
End If
sh.Run """" & py & """ """ & script & """", 0, False
