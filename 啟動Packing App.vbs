Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d ""D:\user\lydia_chen\Documents\packing-app"" && rmdir /s /q __pycache__ 2>nul & python app.py", 0, False
