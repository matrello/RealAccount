Set WshShell=WScript.CreateObject("WScript.Shell")
WScript.Sleep(2000)

activated = False: dt = 100: Time2wait = 4000

Do While (Not activated And Time2wait > 0)
	activated = WshShell.AppActivate("Microsoft Office Outlook")
	WScript.Sleep (dt): Time2wait = Time2wait - dt
Loop

If Time2Wait >= dt Then
	WScript.Sleep(dt): WshShell.SendKeys("{TAB 2}{ENTER}")
Else
	MsgBox "Safety Dialog not found!!!"
End If
