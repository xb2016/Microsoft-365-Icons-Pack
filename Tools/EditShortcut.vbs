' CreateShortcut.vbs
' cscript CreateShortcut.vbs "lnkPath" "vbsPath"
' cscript CreateShortcut.vbs "lnkPath" "vbsPath" restore

Option Explicit
Dim args, lnkPath, vbsPath, restoreMode
Dim shell, lnk, origArgs, fullArgs, cleanedArgs, firstArg, pos1, pos2

Set args = WScript.Arguments
If args.Count < 2 Then WScript.Quit
lnkPath = args(0)
vbsPath = args(1)

restoreMode = False
If args.Count >= 3 Then
    If LCase(args(2)) = "restore" Then restoreMode = True
End If

Set shell = CreateObject("WScript.Shell")
Set lnk = shell.CreateShortcut(lnkPath)

If restoreMode Then
    cleanedArgs = Trim(lnk.Arguments)
    If LCase(lnk.TargetPath) = LCase(vbsPath) Then
        If Left(cleanedArgs, 1) = """" Then
            pos2 = InStr(2, cleanedArgs, """")
            If pos2 > 0 Then
                firstArg = Mid(cleanedArgs, 2, pos2 - 2)
                origArgs = Trim(Mid(cleanedArgs, pos2 + 1))
                lnk.TargetPath = firstArg
                lnk.Arguments = origArgs
                lnk.Save
            End If
        End If
    End If
    WScript.Quit
End If

origArgs = lnk.Arguments
fullArgs = """" & lnk.TargetPath & """"
If Len(origArgs) > 0 Then fullArgs = fullArgs & " " & origArgs

If InStr(LCase(fullArgs), LCase(vbsPath)) > 0 Then
    WScript.Quit
End If

lnk.TargetPath = vbsPath
lnk.Arguments = fullArgs
lnk.Save

WScript.Quit
