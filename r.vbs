
Set objShell = CreateObject("WScript.Shell")
Dim command, password

' Define the SSH command and password
command = "ssh -p 443 -R0:127.0.0.1:22 -o BatchMode=yes -o StrictHostKeyChecking=no -o ServerAliveInterval=30 mrjsT6z7Iwf+tcp@a.pinggy.io"
password = "123"

' Start the SSH session
objShell.Run command, 1, False
WScript.Sleep 2000 ' Wait for the SSH prompt

' Send the password
objShell.SendKeys password
objShell.SendKeys "{ENTER}"

WScript.Sleep 5000 ' Wait for commands to execute
