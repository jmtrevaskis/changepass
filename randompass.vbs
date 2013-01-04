'local administrator password reset tool
'By James Trevaskis
'http://blog.metasplo.it
'Last edited: 04/01/2013


'input checks
If WScript.Arguments.Count = 1 Then
  inputFileName = WScript.Arguments.Item(0)
Else
  Wscript.Echo "Usage: randompass.vbs inputFileName"
  Wscript.Quit
End If

'get username/password
username=InputBox("Enter Username")
password=InputBox("Enter Password")

'initial sets
Set objFSO=CreateObject("Scripting.FileSystemObject")

'read input file
Set objFile = objFSO.OpenTextFile(inputFileName, ForReading)
Const ForReading = 1

'prepare file for writing
outFile="passwords.csv"
If objFSO.FileExists (outFile) then
  Const ForAppending = 8
  Set objFile2 = objFSO.OpenTextFile(outFile, ForAppending)
else
  Set objFile2 = objFSO.CreateTextFile(outFile, True)
  objFile2.Write "server,password,date" & vbCrLf
End If


'main loop
Do Until objFile.AtEndOfStream
	'generate random password
	newPassword = GetRandom(15)
	
	'get computer name
	targetComputerName = objFile.ReadLine

	'pspasswd string
	stringPspasswd = "pspasswd.exe \\" & targetComputerName & " -u " & username & " -p " & password & " administrator " & newPassword

	'display string
	stringDisplay = targetComputerName & "," & newPassword & "," & Date()

	'echo to console
	WScript.Echo stringPspasswd
    objFile2.Write stringDisplay & vbCrLf
	
	'perform password change and display output
	Set objShell = CreateObject("WScript.Shell")
	Set objExec = objShell.Exec("pspasswd.exe \\" & targetComputerName & " -u " & username & " -p " & password & " administrator " & newPassword)
	'objShell.run("pspasswd.exe \\" & targetComputerName & " -u " & username & " -p " & password & " administrator " & newPassword)
	Do
		line = objExec.StdOut.ReadLine()
		WScript.Echo stringStdOut = line
	Loop While Not objExec.Stdout.atEndOfStream
	
Loop

'cleanup files
objFile2.Close
objFile.Close




'EOF

'randomize password func
Function GetRandom(Count)
    Randomize

    For i = 1 To Count
        If (Int((1 - 0 + 1) * Rnd + 0)) Then
            GetRandom = GetRandom & Chr(Int((90 - 65 + 1) * Rnd + 65))
        Else
            GetRandom = GetRandom & Chr(Int((57 - 48 + 1) * Rnd + 48))
        End If
    Next
End Function
