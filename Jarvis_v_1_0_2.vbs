Set wshshell = wscript.CreateObject("wscript.shell")
Sub Jarvis_User_Interaction()
	Set Sapi = Wscript.CreateObject("SAPI.SpVoice")

	Dim Input

	wshshell.run "%windir%\Speech\Common\sapisvr.exe -SpeechUX"
	Sapi.speak "Please speak, or type, what you want to open"
	Input=InputBox ("Please speak, or type, what you want to open.")
	Command = UCase(Input)

	If InStr(Command,"HELLO") > 0 Then
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
	strFile = "C:\Users\" + strName + "\Desktop\Jarvis\username.txt"
	Set objFile = objFSO.OpenTextFile(strFile)
	Do Until objFile.AtEndOfStream
	    strLine= objFile.ReadLine
	Loop
	Sapi.speak "Hello" + strLine
	strNamed = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
	wshshell.run("""C:\Users\" + strNamed + "\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
	Wscript.quit

	If InStr(Command, "CHANGE") > 0Then
		If InStr(Command, "NAME") > 0 Then
			username = InputBox("What do you want Jarvis to call you?", "Jarvis v.1.0.2 Setup Wizard")
			confirmation = MsgBox("Are you sure you want to be called " + username + "?", vbYesNo, "Jarvis v.1.0.2 Setup Wizard")
			Select Case confirmation
			Case vbYes
				Dim directory, objFile
				outFile = "C:\Users\" + strName + "\Desktop\Jarvis\username.txt"
				Set objFile =  objFSO.createTextFile(outFile,True)
				objFile.Write username & vbCrLf
				objFile.Close
				speechobject.speak "Welcome to Jarvis Version one point oh two."
				wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
				Wscript.quit
			Case vbNo
				MsgBox("Unable to save username")
				Wscript.quit
			End Select
		End If
	End If

	If Command = "YOUTUBE" Then
	Sapi.speak "Opening youtube"
	wshshell.run "www.youtube.com"
	wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")

	Else
	If Command = "INSTRUCTABLES" Then
	Sapi.speak "Opening instructables"
	wshshell.run "www.instructables.com"
	wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
	
	Else
	If Command = "GOOGLE" Then
	Sapi.speak "Opening google"
	wshshell.run "www.google.com"
	wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
	
	Else
	If Command = "COMMAND PROMPT" Then
	Sapi.speak "Opening command prompt"
	wshshell.run "cmd"
	wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
	
	Else
	If Command = "CALCULATOR" Then
	Sapi.speak "Opening calculator"
	wshshell.run "calc"
	wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
	
	Else
	If Command = "NOTEPAD" Then
	Sapi.speak "Opening notepad"
	wshshell.run "notepad"
	wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
	
	Else
	If Command = "" Then
	Wscript.quit
	Else
	
	
	Sapi.speak "I don't recognize your input, Please try something else"
	wshshell.run("""C:\Users\ryank\Desktop\Programs\Coding Resources\Jarvis_v_1_0_2.vbs""")
	End If
	End If
	End If
	End If
	End If
	End If
	End If
	End If
End Sub
Set speechobject = Wscript.CreateObject("SAPI.SpVoice")
Set objFSO=CreateObject("Scripting.FileSystemObject")
strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strFile = "C:\Users\" + strName + "\Desktop\Jarvis\username.txt"
Set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
Loop
If strLine = "" Then
objFile.Close
username = InputBox("What do you want Jarvis to call you?", "Jarvis v.1.0.2 Setup Wizard")
confirmation = MsgBox("Are you sure you want to be called " + username + "?", vbYesNo, "Jarvis v.1.0.2 Setup Wizard")
Select Case confirmation
Case vbYes
	Dim directory, objFile
	outFile = "C:\Users\" + strName + "\Desktop\Jarvis\username.txt"
	Set objFile =  objFSO.createTextFile(outFile,True)
	objFile.Write username & vbCrLf
	objFile.Close
	speechobject.speak "Welcome to Jarvis Version one point oh two."
	Call Jarvis_User_Interaction()
Case vbNo
	MsgBox("Unable to save username")
End Select
Else
speechobject.speak "Hello again " + strLine
Call Jarvis_User_Interaction()
End If
