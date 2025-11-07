' CreateShortcut.vbs
' cscript CreateShortcut.vbs "lnkPath" "vbsPath"

Option Explicit
Dim args, lnkPath, vbsPath, shell, lnk, origArgs, fullArgs

Set args = WScript.Arguments
lnkPath = args(0)
vbsPath = args(1)
Set shell = CreateObject("WScript.Shell")
Set lnk = shell.CreateShortcut(lnkPath)

origArgs = lnk.Arguments
fullArgs = """" & lnk.TargetPath & """"
If Len(origArgs) > 0 Then
    fullArgs = fullArgs & " " & origArgs
End If

If InStr(LCase(fullArgs), LCase(vbsPath)) > 0 Then
    WScript.Quit
End If

lnk.TargetPath = vbsPath
lnk.Arguments = fullArgs
lnk.Save

WScript.Quit
