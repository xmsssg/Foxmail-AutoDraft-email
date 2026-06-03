Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = "D:\Study\AutoEmail\Code"
shell.Run """C:\Python314\pythonw.exe"" -m app.main --gui", 0, False
