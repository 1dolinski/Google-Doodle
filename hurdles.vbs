set wshshell = wscript.CreateObject("wscript.shell")
wscript.sleep 4000
For i = 1 to 400
wshshell.SendKeys "{left}"
wscript.sleep 50
wshshell.SendKeys "{right}"
Next 
