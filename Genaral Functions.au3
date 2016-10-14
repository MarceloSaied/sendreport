; Com Error Handler
Func MyErrFunc()
    $HexNumber = Hex($oMyError.number, 8)
    $oMyRet[0] = $HexNumber
    $oMyRet[1] = StringStripWS($oMyError.description, 3)
;~     ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
    ConsoleWrite("Send mail Description:" & $oMyRet[1] & @LF)
    SetError(1); something to check for when this function returns
    Return
EndFunc

Func PortCheck($SmtpServer,$px)
	Opt('TCPTimeout', 1000)
	TCPStartup() ;Start TCP services
	$py=TCPNameToIP($SmtpServer)
	ConsoleWrite("Checking connectio to SMTP server " & $SmtpServer & "  IP:" & $py  & @CRLF )
	FileWriteLine($logFile,"Checking connectio to SMTP server " & $SmtpServer & "  IP:" & $py  & @CRLF )
	$testport = TCPConnect($py,$px)
	If @error Then
		ConsoleWrite("SMTP Mail connection TCPConnect failed with WSA error: " & @error & @CRLF & $testport & @CRLF )
		FileWriteLine($logFile,"SMTP Mail connection TCPConnect failed with WSA error: " & @error & @CRLF & $testport & @CRLF )
		Return False
	else
		TCPCloseSocket($px)
		TCPShutdown ( )
		Return True
	endif
EndFunc
;
