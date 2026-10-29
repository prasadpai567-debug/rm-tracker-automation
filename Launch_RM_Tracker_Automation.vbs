Option Explicit

Dim shell, fso, baseDir, pythonw, scriptPath
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = fso.BuildPath(baseDir, ".venv\Scripts\pythonw.exe")
scriptPath = fso.BuildPath(baseDir, "rm_tracker_app.py")

If Not fso.FileExists(pythonw) Then
    pythonw = "pythonw.exe"
End If

shell.CurrentDirectory = baseDir
shell.Run """" & pythonw & """ """ & scriptPath & """", 0, False
