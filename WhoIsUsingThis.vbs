If WScript.Arguments.Count = 0 Then WScript.Quit

Set WshShell = CreateObject("WScript.Shell")
Set WshEnv = WshShell.Environment("PROCESS")

TargetPath = WScript.Arguments(0)
Quote = Chr(34)
ScriptPath = "D:\Users\joty79\scripts\WhoIsUsingThis\WhoIsUsingThis.ps1"

' Set flag so the script knows it was launched hidden
WshEnv("WHOISUSING_HIDDEN") = "1"

' Silent launch: powershell detects WT/SafeMode and re-launches appropriately
Command = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Quote & ScriptPath & Quote & " " & Quote & TargetPath & Quote

WshShell.Run Command, 0, False
