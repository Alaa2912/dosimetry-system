Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.Run "cmd /c python -c ""import xlrd"" 2>nul || pip install xlrd", 0, True
WshShell.Run "python -m uvicorn app.main:app --host 0.0.0.0 --port 8000", 0, False
WScript.Sleep 1500
WshShell.Run "http://localhost:8000"
