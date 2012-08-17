set wshshell = wscript.CreateObject("wscript.shell")
wscript.sleep 4000
For i = 1 to 5
wshshell.SendKeys " "
wscript.sleep 212
wshshell.SendKeys " "  
wscript.sleep 950
Next 
  
For i = 1 to 4
wshshell.SendKeys " "
wscript.sleep  400
wshshell.SendKeys " "  
wscript.sleep 950
Next   

wscript.sleep 150
For i = 1 to 4
wshshell.SendKeys " "
wscript.sleep 600
wshshell.SendKeys " "  
wscript.sleep 950
Next  

wscript.sleep 150
For i = 1 to 4
wshshell.SendKeys " "
wscript.sleep 750  
wshshell.SendKeys " "  
wscript.sleep 950
Next       