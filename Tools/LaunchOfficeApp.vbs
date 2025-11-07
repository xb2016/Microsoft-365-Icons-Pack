' LaunchOfficeApp.vbs
' cscript LaunchOfficeApp.vbs "appPath" [opts...]

Option Explicit
Dim args, strCmd, i, objShell

Set args = WScript.Arguments

If args.Count = 0 Then
    MsgBox "Please provide the path of the program to launch, e.g.: LaunchOfficeApp.vbs ""C:\...\App.exe""", vbExclamation, "Error"
    WScript.Quit 1
End If

strCmd = Chr(34) & args(0) & Chr(34)

For i = 1 To args.Count - 1
    strCmd = strCmd & " " & args(i)
Next

Set objShell = CreateObject("WScript.Shell")
objShell.Run strCmd, 1, False

WScript.Quit
