Add-Type -AssemblyName System.Windows.Forms


start chrome "the http link to where you want to login"
#chrome so that it opens the chrome brower, otherwise pick another

Start-Sleep -Seconds 5 

$wshell = New-Object -ComObject WScript.Shell
$wshell.AppActivate("Google Chrome") #change if using different browser
Start-Sleep -Milliseconds 200

[System.Windows.Forms.SendKeys]::SendWait("username")
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}") #or tab depending if it is the same page or not
Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait("password")
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")