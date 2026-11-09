' CMMC L2 SSP Drafting Tool — stop the silent launcher
' Only needed if you started the tool with "Start SSP Tool (no window).vbs".
' If you used the .bat launcher, just close its window instead.

Dim shell, answer
Set shell = CreateObject("WScript.Shell")

answer = MsgBox("Stop the CMMC L2 SSP Drafting Tool?" & vbCrLf & vbCrLf & _
                "Make sure you have saved your work in the browser first.", _
                vbQuestion + vbYesNo, "CMMC L2 SSP Drafting Tool")

If answer = vbYes Then
    ' /F force, /IM by image name. Hidden window, wait for completion.
    shell.Run "taskkill /F /IM pythonw.exe", 0, True
    MsgBox "The tool has been stopped.", vbInformation, "CMMC L2 SSP Drafting Tool"
End If
