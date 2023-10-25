Dim arr

arr = Array("calc.exe", "cmd.exe")

Set shell=CreateObject("wscript.shell")
for each item in arr:
    Shell.Run(item)
next





