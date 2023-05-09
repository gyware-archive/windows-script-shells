If WScript.Arguments.Named("host") = "scriptcontrol" Then
    Do While True
        On Error Resume Next
        Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
        If Err.Number <> 0 Then
            WScript.StdErr.WriteLine "Can't create the ""MSScriptControl.ScriptControl"" object: This script must be run in 32-bit CScript to use ScriptControl (usually located at ""%windir%\SysWOW64\cscript.exe"")"
            WScript.Quit
        End If
        On Error GoTo 0
        With CreateObject("MSScriptControl.ScriptControl")
            .Language = "VBScript"
            .AllowUI = True
            On Error Resume Next
            WScript.StdOut.Write ">>> "
            WScript.StdOut.WriteLine .AddCode(WScript.StdIn.ReadLine)
            If Err.Number <> 0 Then
                WScript.StdErr.WriteLine Err.Number & " " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End With
    Loop
Else
    Do While True
        On Error Resume Next
        WScript.StdOut.Write ">>> "
        WScript.StdOut.WriteLine Eval(WScript.StdIn.ReadLine)
        If Err.Number <> 0 Then
            WScript.StdErr.WriteLine Err.Number & " " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Loop
End If
