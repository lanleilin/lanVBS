On Error Resume Next 
Dim wsh,ye 
set wsh=createobject("wscript.shell") 
for i=1 to 10 
wscript.sleep 700 
wsh.AppActivate("银翼") 
wsh.sendKeys "^v" 
wsh.sendKeys i 
wsh.sendKeys "%s" 
next 
wscript.quit