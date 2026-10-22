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
'      NOTE: this applies to comments too. Any non-ASCII byte anywhere in this
'      file breaks the launcher silently.
'
' Hardened behaviour:
'   - launches watchdog.pyw (NOT watcher.py): the watchdog starts watcher.py and
'     relaunches it within minutes if it dies. Launching watcher.py directly
'     here would mean a silent crash stays unnoticed until the next logon.
'   - falls back to python.exe when pythonw.exe is missing
'   - retries 12 x 10s for the target file (cold boot can be slow)
'   - appends a timestamped line to <project>\_logs\startup.log on success,
'     which is the only reliable proof that the boot start really fired
'     (change TRACE_DIR below if you use a different log_dir)
'   - deletes a stale error file after a good start
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
errfile = base & "\startup_error.txt"

pyw = "C:\Users\Administrator\.workbuddy\binaries\python\versions\3.14.3\pythonw.exe"
pyn = "C:\Users\Administrator\.workbuddy\binaries\python\versions\3.14.3\python.exe"
target = "C:\path\to\project\scripts\watchdog.pyw"

If fso.FileExists(pyw) Then
    py = pyw
ElseIf fso.FileExists(pyn) Then
    py = pyn
Else
    WriteErr "E2 no python interpreter: " & pyw
End If

ready = False
For i = 1 To 12
    If fso.FileExists(target) Then
        ready = True
        Exit For
    End If
    WScript.Sleep 10000
Next
If Not ready Then WriteErr "E3 watchdog.pyw not found: " & target

runerr = ""
On Error Resume Next
sh.Run """" & py & """ """ & target & """", 0, False
If Err.Number <> 0 Then runerr = "E9 run failed: " & Err.Description
Err.Clear
On Error GoTo 0
If runerr <> "" Then WriteErr runerr

' ---- success: drop stale error, then leave a trace
On Error Resume Next
If fso.FileExists(errfile) Then fso.DeleteFile errfile, True
projroot = fso.GetParentFolderName(fso.GetParentFolderName(target))
tracedir = projroot & "\_logs"
If Not fso.FolderExists(tracedir) Then fso.CreateFolder tracedir
ts = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
     Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
Set tf = fso.OpenTextFile(tracedir & "\startup.log", 8, True, -1)
tf.WriteLine ts & "  startup fired | python=" & py & " | target=" & target
tf.Close

Sub WriteErr(msg)
    Dim f
    On Error Resume Next
    Set f = fso.CreateTextFile(errfile, True, -1)
    f.Write msg
    f.Close
    WScript.Quit 2
End Sub
