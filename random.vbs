Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "Notepad.exe"
For i = 1 to 10
	wscript.sleep 300      
	objShell.SendKeys " "
Next

