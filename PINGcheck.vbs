' DESCRIPTION:
' Scans specific IP addresses on the local network using PING or SSH,
' then emails approximately how long the node has been offline.

' Requires PortQry.exe in the same directory, for SSH queries

' Create a file in the same directory named ScanMe.txt, and add content like below:
'    //Use "//" to ignore a line
'    //To scan for SSH instead of PING, prepend "22_" to the line. Ex: 22_192.168.1.1
'    
'    //192.168.1.167
'    //22_192.168.1.164
'    22_192.168.1.100
'    10.12.25.5
'    192.168.1.5
'    192.168.1.250

Dim strComputer, strMailTo, strSubj, strMsg, objFSO, strType
Dim dtDate, dtTime, strEvent

'strComputer = WScript.Arguments.Item(0)
strMailTo = "jcool@peanuts.com"
strAddrList = "ScanMe.txt"
nInternal = 30000 'Time, in milliseconds, to wait between scans

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Check for LOGS folder and create if not there
If Not objFSO.FolderExists("Logs") Then
	objFSO.CreateFolder("Logs")
End If

While TRUE
	If objFSO.FileExists(strAddrList) Then
		Set objFile = objFSO.OpenTextFile(strAddrList, 1)
		Do While Not objFile.AtEndOfStream
			strComputer = objFile.ReadLine
			strComputer = Trim(strComputer)
			If strComputer <> "" And (Left(strComputer,2) <> "//")Then
				WScript.Echo "Next item in list: " & strComputer
				If Left(strComputer,3) = "22_" Then
					strComputer = Replace(strComputer,"22_","")
					ScanSSH (strComputer)
				Else	
					PingAddress (strComputer)
				End If
			End If
		Loop
		objFile.Close
	Else
		strSubj = "Error Encountered in PING Monitor"
		strMsg = time() & " - " & "Could not find input file: " & strAddrList
		WScript.Echo strMsg
		MailNotify strSubj, strMsg
		Wscript.Quit
	End If
	'Wait for 30 seconds
	WScript.Sleep(nInternal)
Wend

Sub ScanSSH (strComputer)
	strType = "SSH"
	Set objShell = CreateObject("WScript.Shell")
	strCmd = "%COMSPEC% /c PortQry.exe -n " & strComputer & " -nr -e 22"
	nRet = objShell.Run (strCmd,1,True)
	'Wscript.Echo "  Return of PortQry: " & nRet
	If nRet = 0 Then 'SSH is listening
		If objFSO.FileExists("UIDs\" & strType & strComputer & ".uid") Then
			'Delete the UID file, it's back online
			strSubj = strComputer & " is back online"
			strMsg = time() & " - " & strComputer & " is back online"
			objFSO.DeleteFile "UIDs\" & strType & strComputer & ".uid" 
			WScript.Echo "  " & strMsg
			MailNotify strSubj, strMsg
			LogIt strComputer, date(), time(), "Node back online"
		Else
			WScript.Echo "  Found " & strComputer & " via SSH"
		End If
	Else
		strSubj = "Unable to find " & strComputer & " via SSH"
		strMsg = time() & " - " & "Unable to find " & strComputer & " via SSH"
		WScript.Echo "  " & strSubj
		NoFinder strComputer, strSubj, strMsg, strType
	End If
End Sub

Sub PingAddress(strComputer)
	strType = "PING"
	Set objWMILocalSvc = GetObject("winmgmts:\\.\root\cimv2")
	Set objPing = objWMILocalSvc.ExecQuery ("Select * from Win32_PingStatus " & _
		"Where Address = '" & strComputer & "'")
	For Each objStatus in objPing
		If objStatus.StatusCode <> 0 Then 
			strSubj = "Unable to PING " & strComputer
			strMsg = time() & " - " & "Unable to PING " & strComputer
			WScript.Echo "  " & strSubj
			NoFinder strComputer, strSubj, strMsg, strType
		Else
			If objFSO.FileExists("UIDs\" & strType & strComputer & ".uid") Then
				'Delete the UID file, it's back online
				strSubj = strComputer & " is back online"
				strMsg = time() & " - " & strComputer & " is back online"
				objFSO.DeleteFile "UIDs\" & strType & strComputer & ".uid" 
				WScript.Echo "  " & strMsg
				MailNotify strSubj, strMsg
				LogIt strComputer, date(), time(), "Node is back online"
			Else
				WScript.Echo "  Found " & strComputer
				'Check a list to see if this is being scanned, then run a scan if it is
			End If
		End If
	Next
End Sub

'Placeholder for tossing a found address to a monitoring tool for analysis
Sub ServerMonitor
	'Examples of stuff to check:
	' Page File Usage
	' Processor Usage
	' Free Disk Space
	' Memory Usage
	' Service Status
End Sub

Sub NoFinder (strComputer, strSubj, strMsg, strType)
		'	strSubj = "Unable to PING " & strComputer
		'	strMsg = time() & " - " & "Unable to PING " & strComputer
			If Not objFSO.FileExists("UIDs\" & strType & strComputer & ".uid") Then
				'Create the UID file
				objFSO.CreateTextFile "UIDs\" & strType & strComputer & ".uid", True
				'Only notify if this is the first time we're not seeing it
				WScript.Echo "  " & strMsg
				MailNotify strSubj, strMsg
				LogIt strComputer, date(), time(), "Node offline"
			ElseIf objFSO.FileExists("UIDs\" & strType & strComputer & ".uid") Then
				Set objFileName = objFSO.GetFile("UIDs\" & strType & strComputer & ".uid")
				Wscript.Echo "  UID File Created: " & objFileName.DateCreated
				dtCreated = objFileName.DateCreated
				nDiff = DateDiff("n",dtCreated,Now)
				Wscript.Echo "  Time difference: " & nDiff & " minutes"
				If objFileName.Size > 0 Then
					'Read the text in the file to see if and when we last notified
					Set fileRead = objFSO.OpenTextFile("UIDs\" & strType & strComputer & ".uid", 1) 
					strReadLine = fileRead.ReadLine
					fileRead.Close
				Else
					strReadLine = ""
				End If
				If nDiff > 60 Then
					If  strReadLine <> "SIXTY MIN" Then
						Set filetxt = objFSO.OpenTextFile("UIDs\" & strType & strComputer & ".uid", 2, True) 
						filetxt.WriteLine "SIXTY MIN"
						filetxt.Close
						strSubj = strComputer & " offline for more than one hour"
						strMsg = time() & " - " & strComputer & " has been offline for more than one hour"
						Wscript.Echo "  " & strMsg
						MailNotify strSubj, strMsg
						LogIt strComputer, date(), time(), "Node offline for more than one hour"
					End If
				ElseIf nDiff > 30 Then
					If  strReadLine <> "THIRTY MIN" Then
						Set filetxt = objFSO.OpenTextFile("UIDs\" & strType & strComputer & ".uid", 2, True) 
						filetxt.WriteLine "THIRTY MIN"
						filetxt.Close  
						strSubj = strComputer & " offline for more than thirty minutes"
						strMsg = time() & " - " & strComputer & " has been offline for more than thirty minutes, per " &_
							strType & " scan."
						Wscript.Echo "  " & strMsg
						MailNotify strSubj, strMsg
						LogIt strComputer, date(), time(), "Node offline for more than thirty minutes"
					End If
				ElseIf nDiff > 5 Then
					If strReadLine <> "FIVE MIN" Then 
						Set filetxt = objFSO.OpenTextFile("UIDs\" & strType & strComputer & ".uid", 2, True) 
						filetxt.WriteLine "FIVE MIN"
						filetxt.Close  
						strSubj = strComputer & " offline for more than five minutes"
						strMsg = time() & " - " & strComputer & " has been offline for more than five minutes, per " &_
							strType & " scan."
						Wscript.Echo "  " & strMsg
						MailNotify strSubj, strMsg
						LogIt strComputer, date(), time(), "Node offline for more than five minutes"
					End If
				ElseIf nDiff > 2 Then
					If  strReadLine <> "TWO MIN" Then
						Set filetxt = objFSO.OpenTextFile("UIDs\" & strType & strComputer & ".uid", 2, True) 
						filetxt.WriteLine "TWO MIN"
						filetxt.Close  
						strSubj = strComputer & " offline for more than two minutes"
						strMsg = time() & " - " & strComputer & " has been offline for more than two minutes, per " &_
							strType & " scan."
						Wscript.Echo "  " & strMsg
						MailNotify strSubj, strMsg
						LogIt strComputer, date(), time(), "Node offline for more than two minutes"
					End If
				End If
			End If
End Sub

Sub MailNotify (strSubj, strMsg)
	Set objEmail = CreateObject("CDO.Message")
	objEmail.From = "PingMonitor@peanuts.com"
	objEmail.To = strMailTo
	objEmail.Subject = strSubj
	objEmail.Textbody = strMsg
	objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.yourmailserver.com"
	objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	objEmail.Send
End Sub


'____________________ UPDATE THE LOG FILE ________________________
Sub LogIt (strComputer, dtDate, dtTime, strEvent)
	nDay = Day(Now)
	nMon = Month(Now)
	nYer = Year(Now)
	strToday = nYer & nMon & nDay
	Set filetxt = objFSO.OpenTextFile("Logs\" & strToday & ".log", 8, True) 
	filetxt.WriteLine(strComputer & "," & dtDate & "," & dtTime & "," & strEvent)
	filetxt.Close
End Sub
