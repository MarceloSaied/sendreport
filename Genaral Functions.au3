

#region app greneral
	Func _Licence()
		ConsoleWrite('++_Licence() = '& @crlf)
		If @Compiled Then
			Local $iDateCalc = _DateDiff('s', "2016/12/8 00:00:00", _NowCalc())
			If $iDateCalc > 0 Then Exit
		EndIf
	EndFunc



#endregion
#region Log
	Func  StartLog()  ; init log file
		$logFile=FileOpen($logFileName,2)
		FileWriteLine($logFile,'Start of activities' & @TAB & @TAB  & _NowCalc() & @crlf)
	EndFunc
#endregion
#region Error Handeler
	Global Const $oErrorHandler = ObjEvent("AutoIt.Error", "ObjErrorHandler")
	Func ObjErrorHandler()
		ConsoleWrite(   "A COM Error has occured!" & @CRLF  & @CRLF & _
                                "err.description is: "    & @TAB & $oErrorHandler.description    & @CRLF & _
                                "err.windescription:"     & @TAB & $oErrorHandler & @CRLF & _
                                "err.number is: "         & @TAB & Hex($oErrorHandler.number, 8)  & @CRLF & _
                                "err.lastdllerror is: "   & @TAB & $oErrorHandler.lastdllerror   & @CRLF & _
                                "err.scriptline is: "     & @TAB & $oErrorHandler.scriptline     & @CRLF & _
                                "err.source is: "         & @TAB & $oErrorHandler.source         & @CRLF & _
                                "err.helpfile is: "       & @TAB & $oErrorHandler.helpfile       & @CRLF & _
                                "err.helpcontext is: "    & @TAB & $oErrorHandler.helpcontext & @CRLF _
                            )
	EndFunc
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
#endregion
#region sciTe
	Func _ClearSciteConsole()
		if $debugflag=1 then ConsoleWrite('++_ClearSciteConsole() = '& @crlf)
		ControlSend("[CLASS:SciTEWindow]", "", "Scintilla2", "+{F5}")
	EndFunc
#endregion
#region time
	Func  Sec2Time($nr_sec)
		$sec2time_hour = Int($nr_sec / 3600)
		$sec2time_min = Int(($nr_sec - $sec2time_hour * 3600) / 60)
		$sec2time_sec = $nr_sec - $sec2time_hour * 3600 - $sec2time_min * 60
		Return StringFormat('%02d:%02d:%02d', $sec2time_hour, $sec2time_min, $sec2time_sec)
	EndFunc   ;==>Sec2Time
	Func _GetUnixTime($sDate = 0);Date Format: 2013/01/01 00:00:00 ~ Year/Mo/Da Hr:Mi:Se
		Local $aSysTimeInfo = _Date_Time_GetTimeZoneInformation()
		Local $utcTime = ""
		If Not $sDate Then $sDate = _NowCalc()
		If Int(StringLeft($sDate, 4)) < 1970 Then Return ""
		If $aSysTimeInfo[0] = 2 Then
			$utcTime = _DateAdd('n', $aSysTimeInfo[1] + $aSysTimeInfo[7], $sDate)
		Else
			$utcTime = _DateAdd('n', $aSysTimeInfo[1], $sDate)
		EndIf
		Return _DateDiff('s', "1970/01/01 00:00:00", $utcTime)
	EndFunc   ;==>_GetUnixTime
#endregion
#region file work
	Func _ByteSuffix($iBytes)
		$iIndex = 0
		Dim $aArray[9] = [' bytes', ' KB', ' MB', ' GB', ' TB', ' PB', ' EB', ' ZB', ' YB']
		While $iBytes > 1023
			$iIndex += 1
			$iBytes /= 1024
		WEnd
		Return Round($iBytes) & $aArray[$iIndex]
	EndFunc   ;==>ByteSuffix
	Func _IsFolder($sFolder)
		Local $sAttribute = FileGetAttrib($sFolder)
		If @error Then
			If $SkipobtainAtributesFlag=0 Then
;~ 				MsgBox(4096, "Error 1053", "Could not obtain the file attributes."&@crlf&"$sFolder="&$sFolder&@crlf&"$sAttribute="&$sAttribute,0)
				Return 0
			else
				Return 0
			endif
		endif
		Return StringInStr($sAttribute,"D")
	EndFunc
	Func _FileInUse($filename)
		$handle = FileOpen($filename, 1)
		$result = False
		if $handle = -1 then $result = True
		FileClose($handle)
		ConsoleWrite('_FileInUse $filename= ' & $filename& " $result= "&$result & @crlf )
		return $result
	EndFunc
	Func _GetTempDir()
	   ConsoleWrite('++_GetTempDir() = '& @crlf)
		If @OSVersion="WIN_2003" Or @OSVersion="WIN_2000"  Or @OSVersion="WIN_XP" Then
			$tempDir=@TempDir
		else
			$tempDir=@HomeDrive&@HomePath&"\AppData\Local\Temp"
		endif
		Return $tempdir
	EndFunc
#endregion
#region Firewall
	Func _SetFirewall()
		$f_sav = @ScriptDir & '\Pocket-SF.sav'
		$_GetIP = @IPAddress1
		Local $s_reg = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run', $f_fopen, $f_Socket, $f_tray
	;~ 	Local $p_MD5 = PluginOpen('MD5.dll')

		$s_locip = @IPAddress1

		If Not FileExists($f_sav) Then
			ShellExecute("cmd.exe", 'netsh firewall add allowedprogram ' & @ScriptFullPath _
				& " " & 'ACNalertFT ENABLE ALL', "", "", @SW_HIDE) ;Add firewall exception

			For $i = 1 To 3
				FileWrite($f_sav, @CRLF)
			Next
		EndIf
	EndFunc
#endregion
#region controls
	Func _IsCheckedType($iControlID)
		$sts=$GUI_UNCHECKED
		If _IsChecked($iControlID) Then $sts=$GUI_CHECKED
		Return $sts
	EndFunc
	Func _IsChecked($iControlID)
		Return BitAND(GUICtrlRead($iControlID), 1) = 1
	EndFunc
#endregion
#region Network
	Func _ReverseDNS($IPAddress)
		ConsoleWrite('++_ReverseDNS() = '& $IPAddress & @crlf)
		GUICtrlSetData($LBL_Settings_DNSregistration,"DNS Registration Lookup")
		$IPAddress = StringStripWS($IPAddress,3)
		$sCommand="nslookup "& $IPAddress
		$ResponseText = _GetDOSOutput($sCommand)
		If StringInStr($ResponseText,"*** UnKnown")>0 Then
			GUICtrlSetData($LBL_Settings_DNSregistration,"DNS Registration Lookup UnKnown")
			Return "Unknown"
		endif
		GUICtrlSetState($LBL_Settings_DNSregistration,@SW_HIDE)
		$x1 = StringInStr($ResponseText, "Name:")
		$x2 = StringInStr($ResponseText, "Address",0,-1)
		If $x1 = 0 or $x2 = 0 Then Return "Unknown"
		Dim $arr[3]
		$arr[1]=StringStripWS(StringMid($ResponseText, $x1 + 6, $x2 - $x1 - 6),3)
		$arr[2]=StringStripWS(StringMid($ResponseText, $x2 + 8, 17),3)
		If $x1 > 0 and $x2 > 0 Then Return $arr
		GUICtrlSetData($LBL_Settings_DNSregistration,"DNS Registration Lookup Error")
		Return "UnknownError"
	EndFunc
	Func _ReverseDNSConnector($IPAddress)
		if $debugflag=1 then ConsoleWrite('++_ReverseDNSConnector() = '& $IPAddress & @crlf)
		$IPAddress = StringStripWS($IPAddress,3)
		$sCommand="nslookup "& $IPAddress
		$ResponseText = _GetDOSOutput($sCommand)
		if $debugflag=1 then ConsoleWrite('ResponseText = ' & $ResponseText & @crlf )
		If StringInStr($ResponseText,"*** UnKnown")>0 Then
			Return "Unknown"
		endif
		$x1 = StringInStr($ResponseText, "Name:")
		$x2 = StringInStr($ResponseText, "Address",0,-1)
		If $x1 = 0 or $x2 = 0 Then Return "Unknown"
		Dim $arr[3]
		$arr[1]=StringStripWS(StringMid($ResponseText, $x1 + 6, $x2 - $x1 - 6),3)
		$arr[2]=StringStripWS(StringMid($ResponseText, $x2 + 8, 17),3)
		If $x1 > 0 and $x2 > 0 Then Return $arr[2]
		Return "UnknownError"
	EndFunc
	Func _GetFQDN()
		if $debugflag=1 then ConsoleWrite('++_GetFQDN() = '& @crlf)
		$objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & @ComputerName & "\root\cimv2")
		If @error Then Return SetError(2, 0, "")
		$colItems = $objWMIService.ExecQuery("SELECT Name,Domain FROM Win32_ComputerSystem ", "WQL", 0x30)
		If IsObj($colItems) Then
			For $objItem In $colItems
				If $objItem.Domain = "" Then
;~ 					Return  $objItem.Name
					Return SetError(3, 0, "")
				Else
					Return  $objItem.Name & "." & $objItem.Domain
				endif
			Next
		Else
			Return SetError(3, 0, "")
		EndIf
	EndFunc
	Func _GetServerDNSip()
		if $debugflag=1 then ConsoleWrite('++_GetServerDNSip() = '& @crlf)
		$fqdn=_GetFQDN()
		_consolewrite("FQDN = "&$fqdn)
		If @Compiled Then
			$ip=_ReverseDNSConnector($fqdn)
		Else
			$ip=@IPAddress1
		endif
		_consolewrite("ServerDNSip = "&$ip)
		Return $ip
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
#endregion
#region Encription
	Func _Hashing($Password,$Hashflag=0)
		ConsoleWrite('++_Hashing() = '& @crlf )
			if $Hashflag=0 then
				Local $bEncrypted = _StringEncrypt(1,$Password,$HashingPassword,1)
			Else
				Local $bEncrypted = _StringEncrypt(0,$Password,$HashingPassword,1)
			EndIf
		return $bEncrypted
	EndFunc
#endregion
#region Design
	Func _color($dato)
;~ 		ConsoleWrite('++_color() = '& @crlf )
		If $dato="blue" Then Return "0x0000FF"
		If $dato="lightgreen" Then Return "0x00FF00"
		If $dato="Darkgreen" Then Return "0x088A08"
;~ 		If $dato="green" Then Return "0x3A5F0B"  41A317
		If $dato="green" Then Return "0x99CC00"
		If $dato="red" Then Return "0xFF0000"
		If $dato="orange" Then Return "0xFF6600"
		If $dato="yellow" Then Return "0xFFFF00"
		If $dato="Purple" Then Return "0xFF00FF"
		If $dato="lightblue" Then Return "0x00FFFF"
		If $dato="black" Then Return "0x000000"
		If $dato="white" Then Return "0xffffff"
		If $dato="blue" Then Return "0x0000ff"
	EndFunc
#endregion


