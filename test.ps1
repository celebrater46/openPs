Add-Type -AssemblyName System.Windows.Forms
Start-Sleep -m 500

# change window
Start-Process -FilePath "C:\Windows\system32\notepad.exe"
# add-type -assembly microsoft.visualbasic
# [microsoft.visualbasic.interaction]::AppActivate("notepad")

Start-Sleep -m 1000
[System.Windows.Forms.SendKeys]::SendWait("Hello World.")
