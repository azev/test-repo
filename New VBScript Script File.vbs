''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description	: 
' Parameters	: 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

CONST TITLE = "My Script Title"

'	If (WScript.Arguments.Count = 0) Then
'		MsgBox "Missing parameters.", 48, ""
'		WScript.Quit
'	End if

Dim shell : Set shell = CreateObject("WScript.Shell")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim rx : Set rx = New RegExp
Dim oWMI : Set oWMI = GetObject("winmgmts:!\\.\root\cimv2")

shell.CurrentDirectory = fso.GetParentFolderName(wscript.ScriptFullName)
Dim temp : temp = shell.ExpandEnvironmentStrings("%Temp%")


With rx	
	.IgnoreCase = True
	.Global = True
	.Multiline = True
End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' array functions, pop(), push(), ...
' https://gallery.technet.microsoft.com/scriptcenter/c05af93f-1213-4238-9c96-6218141bf66d

'shell.Run "X:\BIN\playwav\playwav.exe X:\SCRIPTS\beep.wav", 0

Function readUTF8(fileName)
	'https://stackoverflow.com/questions/13851473/read-utf-8-text-file-in-vbscript
	Dim oStream, strData

	Set oStream = CreateObject("ADODB.Stream")
	oStream.CharSet = "utf-8"
	oStream.Open
	oStream.LoadFromFile(fileName)

	readUTF8 = oStream.ReadText()

	oStream.Close
	Set oStream = Nothing
End Function

Sub runConsoleMode()
	Dim sExecutable : sExecutable = LCase(Mid(Wscript.FullName, InstrRev(Wscript.FullName,"\")+1)) ' cscript.exe or wscript.exe
	If sExecutable <> "cscript.exe" Then
		Dim oWMI, oShell, sPath, colProcesses, objProcess, cmd
		Set oShell = CreateObject("wscript.shell")
		Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
		sPath = replace(Wscript.ScriptFullName, "\", "\\")
		Set colProcesses = oWMI.ExecQuery ("Select * from Win32_Process WHERE CommandLine LIKE '%wscript.exe%" & sPath & "%'")
		For Each objProcess in colProcesses
			cmd = replace(objProcess.CommandLine, "wscript.exe", "cscript.exe", 1, -1, 1)
			oShell.Run cmd
			Wscript.Quit
		Next
	End If
End Sub

Sub SetPriority(priority) ' 64:idle, 16384:below normal, 32:normal, 32768:above normal, 128:high priority, 256:Realtime
	Dim oWMI : Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
	Dim sPath : sPath = replace(Wscript.ScriptFullName, "\", "\\")
	Dim colProcesses, objProcess
	Set colProcesses = oWMI.ExecQuery ("Select * from Win32_Process WHERE CommandLine LIKE '%wscript.exe%" & sPath & "%'")
	For Each objProcess in colProcesses
		objProcess.SetPriority(priority)			
	Next
End Sub

function bind(cmdStr, arrParams)
	Dim i
	For i = 0 to uBound(arrParams)
		cmdStr = replace(cmdStr, "$"&i+1, arrParams(i))
	Next
	cmdStr = replace(cmdStr, "'", chr(34))
	bind = cmdStr
End Function

Sub mkDirTree(dir)
	If Not fso.FolderExists(dir) Then
		mkDirTree fso.GetParentFolderName(dir)
		fso.createFolder dir
	End If
End Sub

Sub GUImoveFile(srcPath, tgtPath)
	' srcPath may include wildcard. Eg: c:\*.txt
	' https://technet.microsoft.com/en-us/library/ee176633.aspx
	Const FOF_CREATEPROGRESSDLG = &H0&
	Set objFolder = appShell.NameSpace(tgtPath) 
	objFolder.MoveHere srcPath, FOF_CREATEPROGRESSDLG
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set shell = Nothing
Set fso = Nothing
Set rx = Nothing	
