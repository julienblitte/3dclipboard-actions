' Author: Roy
' Date: Sun May 27, 2012

Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys "%{TAB}" 'switch back to application
WshShell.SendKeys Replace(Clipboard.Value, "-", "{TAB}")
