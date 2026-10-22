' Launcher for the Startup folder:
'   %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
'
' Why not derive the path like start_watcher.vbs does?
'   Because this file lives in the Startup folder, so its own location is not
'   the project folder.
'
' So "target" must be hard-coded. Two cases:
'   a) The path is pure ASCII  -> write it as a normal string (as shown below).
'   b) The path contains non-ASCII characters (e.g. Chinese) -> DO NOT write it
'      literally; .vbs is read as ANSI and it would be garbled.
'      Instead run:
'          python tools\gen_chrw_path.py "D:\your project\scripts\watcher.py"
'      and paste the produced  "..." & ChrW(...)  expression as the value of
'      target. That keeps this file ASCII-only.
'
' On failure a startup_error.txt is written next to this file.
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
errfile = base & "\startup_error.txt"

py = "C:\Users\Administrator\.workbuddy\binaries\python\versions\3.14.3\pythonw.exe"
target = "C:\path\to\project\scripts\watcher.py"

If Not fso.FileExists(py) Then
    Set ef = fso.CreateTextFile(errfile, True)
    ef.Write "E2 pythonw.exe not found: " & py
    ef.Close
    WScript.Quit 2
End If
If Not fso.FileExists(target) Then
    Set ef = fso.CreateTextFile(errfile, True)
    ef.Write "E3 watcher.py not found: " & target
    ef.Close
    WScript.Quit 3
End If
On Error Resume Next
sh.Run """" & py & """ """ & target & """", 0, False
If Err.Number <> 0 Then
    Set ef = fso.CreateTextFile(errfile, True)
    ef.Write "E9 run failed: " & Err.Description
    ef.Close
    WScript.Quit 9
End If
