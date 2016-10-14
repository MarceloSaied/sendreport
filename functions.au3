

Func  CreateLogAttachment()  ; init log file
	FileClose($logFile)
	Sleep(2000)
	FileCopy($logFileName,StringReplace($logFileName,".log",".txt"),1)
	$logFile=FileOpen($logFileName,1)
EndFunc


;~ 			SendMail(" STARTUP FAILED ","ERR11216. An error occurred when listing directory." & $ProfileDir & @crlf & " when finding  SAP ABAP message server" & @CRLF & @crlf)



Func  CheckINIfile()  ;checks if ini file exist
	If FileExists($fileini) Then
		Return 1
	Else
		IniWrite($fileini, "Activity", "SendReport", "Yes")
		IniWrite($fileini, "Mail", "Enable", "Yes")
		IniWrite($fileini, "Mail", "SmtpServer", "AMR.SMTP.BBVA.COM")
		IniWrite($fileini, "Mail", "FromAddress", "SendReport@" & @ComputerName)
		IniWrite($fileini, "Mail", "ToAddress", "GroupOrUser@bbva.com")
		IniWrite($fileini, "Mail", "CcAddress", "GroupOrUser@bbva.com")
		IniWrite($fileini, "Mail", "RetryFrecuencyMin", "10")
		IniWrite($fileini, "Mail", "DropTimeMin", "120")
		FileWriteLine($logFile,'INI file was created with defaults. please configure it as needed' & @crlf)
	endif
EndFunc

