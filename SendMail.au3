
Func SendMail($mailalert,$MailMessage)
	Local $var = IniRead($fileini, "Mail", "Enable", "No")
	If StringLower(StringStripWS($var,3))<>"no" then
		$FromName = IniRead($fileini, "Mail", "FromAddress", "SendReport@" & @ComputerName)
		$FromAddress = @ComputerName & "@bbva.com"
		$Importance = "High" ; "Normal"
		$replayTo= "noreply@bbva.com"
		CreateLogAttachment()
		$BccAddress = "marcelo.saied@bbva.com"
		$IPPort = 25 ;$IPPort=465                          ; GMAIL port used for sending the mail
		$ssl = 0      ;$ssl=1                               ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS
		Local $SmtpServer = IniRead($fileini, "Mail", "SmtpServer", "No.SMTP")
		Local $ToAddress = IniRead($fileini, "Mail", "ToAddress", "Noreply@bbva.com")
		Local $CcAddress = IniRead($fileini, "Mail", "CcAddress", "Noreply@bbva.com")
		$Subject = $mailalert
		$aInfo= _Date_Time_GetTimeZoneInformation()
		$Body = "Mail report" & @CRLF & @CRLF & _
		"Server Time = " & @MON & "-" & @MDAY & "-" & @YEAR & " " & @HOUR & ":" & @MIN & @CRLF


		ConsoleWrite("Testing SMTP Mail port." & @CRLF )
		FileWriteLine($logFile,"Testing SMTP Mail port."& @CRLF)
		If PortCheck($SmtpServer,25) Then
			ConsoleWrite("Sending Mail." & @CRLF & "Subject=" & $Subject & @CRLF  & "Body=" & $Body & @CRLF )
			FileWriteLine($logFile,"Sending Mail." & @CRLF & "Subject=" & $Subject & @CRLF  & "Body=" & $Body & @CRLF)
			$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $IPPort, $ssl,$replayTo)
			If @error Then
				FileWriteLine($logFile,'re intent in 1 seconds ' & @crlf)
				ConsoleWrite(@CRLF & 're intent in 1 seconds '  & @crlf)
				Sleep(1000)
				$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $IPPort, $ssl,$replayTo)
				$err=@error
				ConsoleWrite("First intent." & "  Error sending message. Error code:" & $err & "  Description:" & $rc & @CRLF )
				FileWriteLine($logFile,"First intent." & "  Error sending message. Error code:" & $err & "  Description:" & $rc & @CRLF)
				If $err Then
					FileWriteLine($logFile,'re intent in 5 seconds ' & @crlf)
					ConsoleWrite(@CRLF &'re intent in 5 seconds '  & @crlf)
					Sleep(5000)
					$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $IPPort, $ssl,$replayTo)
					$err=@error
					ConsoleWrite("Second intent." & "  Error sending message. Error code:" & $err & "  Description:" & $rc & @CRLF )
					FileWriteLine($logFile,"Second intent." & "  Error sending message. Error code:" & $err & "  Description:" & $rc & @CRLF)
					If $err Then
						FileWriteLine($logFile,'re intent in 60 seconds ' & @crlf)
						ConsoleWrite(@CRLF & 're intent in 60 seconds '  & @crlf)
						Sleep(60000)
						$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $IPPort, $ssl,$replayTo)
						If @error Then
							ConsoleWrite("Third intent." & "  Error sending message. Error code:" & @error & "  Description:" & $rc & @CRLF )
							FileWriteLine($logFile,"Third intent." & "  Error sending message. Error code:" & @error & "  Description:" & $rc & @CRLF)
							Return "Error sending message. Error code:" & @error & "  Description:" & $rc
						EndIf
					endif
				EndIf
			EndIf
		Else
		ConsoleWrite("Mail not sent" & @CRLF )
		FileWriteLine($logFile,"Mail not sent" & @CRLF )
		endif
		Return "Ok"
	endif
EndFunc
;

; The UDF
Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance="Normal", $IPPort = 25, $ssl = 0,$replayTo = "",$s_Username = "" ,$s_Password = "")
    Local $objEmail = ObjCreate("CDO.Message")
    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    $objEmail.To = $s_ToAddress
	$objEmail.ReplyTo = $replayTo
    Local $i_Error = 0
    Local $i_Error_desciption = ""
    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
    $objEmail.Subject = $s_Subject
    If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
        $objEmail.HTMLBody = $as_Body
    Else
        $objEmail.Textbody = $as_Body & @CRLF
    EndIf
    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
        For $x = 1 To $S_Files2Attach[0]
            $S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
;~          ConsoleWrite('@@ Debug : $S_Files2Attach[$x] = ' & $S_Files2Attach[$x] & @LF & '>Error code: ' & @error & @LF) ;### Debug Console
            If FileExists($S_Files2Attach[$x]) Then
                ConsoleWrite('+> File attachment added: ' & $S_Files2Attach[$x] & @LF)
                $objEmail.AddAttachment($S_Files2Attach[$x])
            Else
                ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
                SetError(1)
                Return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    If Number($IPPort) = 0 then $IPPort = 25
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
    ;Update settings
    $objEmail.Configuration.Fields.Update
    ; Set Email Importance
    Switch $s_Importance
        Case "High"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "High"
        Case "Normal"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Normal"
        Case "Low"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Low"
    EndSwitch
    $objEmail.Fields.Update
    ; Sent the Message
    $objEmail.Send
    If @error Then
        SetError(2)
        Return $oMyRet[1]
    EndIf
    $objEmail=""
EndFunc   ;==>_INetSmtpMailCom
;
