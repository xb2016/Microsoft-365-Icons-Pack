' ReflashShortcut.vbs
' cscript ReflashShortcut.vbs "lnkPath"

Option Explicit
Dim shell, lnk

Set shell = CreateObject("WScript.Shell")
Set lnk = shell.CreateShortcut(WScript.Arguments(0))
lnk.Save

WScript.Quit
