Option Explicit
Dim shell, response1, response2, regPath

Set shell = CreateObject("WScript.Shell")
regPath = "register-MEX.reg" ' Se till att filen ligger i samma mapp

' Första varningen
response1 = MsgBox("Are you sure you want to open this file? It will format your .MEX extension if you already have one. If you have a .MEX extension, this will reformat it and you could lose important .MEX files. (If you've never used Regedit or opened a .reg file, you likely don't have a .MEX extension as it's not included with Windows by default - but you can always double-check in the Registry Editor!)", _
                   vbYesNo + vbExclamation, "Warning - .MEX Extension Format")

If response1 = vbNo Then
    MsgBox "Action canceled. Nothing has been changed.", vbInformation, "Cancelled"
    WScript.Quit
End If

' Andra varningen
response2 = MsgBox("Last warning! Are you sure you want to format or create your .MEX extension?", _
                   vbYesNo + vbExclamation, "Final Confirmation")

If response2 = vbNo Then
    MsgBox "Action canceled. Nothing has been changed.", vbInformation, "Cancelled"
    WScript.Quit
End If

' Öppna .reg-filen om användaren bekräftar två gånger
On Error Resume Next
shell.Run "regedit.exe /s """ & regPath & """", 1, True

If Err.Number <> 0 Then
    MsgBox "Could not run register-MEX.reg. Please make sure the file exists.", vbCritical, "Error"
End If
