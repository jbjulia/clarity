' This program sends unobtrusive keystrokes to keep your screen awake
' during inactivity, on positions that may contain privileged restrictions.

Sub initClarity()
	Dim i: i = 0
	Dim WshShell: Set WshShell = WScript.CreateObject("WScript.Shell")
	
	' Prompt user before initializing script
	If MsgBox ("Would you like to replenish your mana?", vbQuestion + vbYesNo, "Clarity") = vbNo Then
		Set WshShell = Nothing
		WScript.Quit
	End If
	
	' Will loop for 12 hours on a 60 second timer
	Do While (i < 720)
		If Err.Number <> 0 Then
			handleErrors
		Else:
			WScript.SendKeys ("{SCROLLLOCK 2}") ' Simulate Scroll Lock keystroke
			WScript.Sleep (60 * 1000)  ' Sleep for 60 seconds
			i = i + 1
		End If
	Loop
	
	initClarity
	
End Sub

Sub handleErrors()
	MsgBox Err.Description, vbExclamation + vbOKCancel, "Error: " & CStr(Err.Number)
	
	Set WshShell = Nothing
	WScript.Quit
	
End Sub

initClarity()