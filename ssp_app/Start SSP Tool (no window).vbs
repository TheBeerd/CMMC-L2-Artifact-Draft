' CMMC L2 SSP Drafting Tool — silent launcher
' Starts the tool with no console window and opens your browser.
'
' To STOP the tool afterwards, double-click "Stop SSP Tool.vbs"
' (there is no window to close with this launcher).

Dim shell, fso, here, pyw
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

here = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = here

' pythonw.exe runs Python without a console window
pyw = "pythonw.exe"

On Error Resume Next
shell.Run """" & pyw & """ app.py", 0, False
If Err.Number <> 0 Then
    MsgBox "Could not start Python." & vbCrLf & vbCrLf & _
           "Run ""Setup (run once).bat"" first, or install Python " & _
           "from https://python.org (tick ""Add Python to PATH"").", _
           vbExclamation, "CMMC L2 SSP Drafting Tool"
    WScript.Quit 1
End If
On Error Goto 0

' Give the server a moment to bind the port, then open the browser
WScript.Sleep 2500
shell.Run "http://127.0.0.1:5000", 1, False
