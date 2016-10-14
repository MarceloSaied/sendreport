Func  StartLog()  ; init log file
	$logFile=FileOpen($logFileName,2)
	FileWriteLine($logFile,'Start of activities' & @TAB & @TAB  & _NowCalc() & @crlf)
EndFunc


Func  CreateLogAttachment()  ; init log file
	FileClose($logFile)
	Sleep(2000)
	FileCopy($logFileName,StringReplace($logFileName,".log",".txt"),1)
	$logFile=FileOpen($logFileName,1)
EndFunc


;~ 			SendMail(" STARTUP FAILED ","ERR11216. An error occurred when listing directory." & $ProfileDir & @crlf & " when finding  SAP ABAP message server" & @CRLF & @crlf)
