' Created by YOGESHWARAN TL
' Runs run_nightly.bat with no window: no flash on the screen every 10 minutes
' (the Adv Pro check reads the screen), and nothing for anyone to close.
Set fso = CreateObject("Scripting.FileSystemObject")
CreateObject("WScript.Shell").Run """" & fso.GetParentFolderName(WScript.ScriptFullName) & "\run_nightly.bat""", 0, False
