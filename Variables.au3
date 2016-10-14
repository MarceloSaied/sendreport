#Region Base
	Global $oMyRet[2]
	Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
#EndRegion

#Region 0) StartLog()
;~ 	$ScripClass=""
;~ 	$logFile=""
;~ 	$horario=@HOUR & "_" & @MIN
;~ 	$logFileName=@ScriptDir & "\" & StringMid(@ScriptName,1,StringLen(@ScriptName)-4)  & "_" & @MON & "-" & @MDAY & "-" & @YEAR & "_" & $horario& ".log"
;~ 	$AttachFiles = @ScriptDir & "\" & StringMid(@ScriptName,1,StringLen(@ScriptName)-4)   & "_" & @MON & "-" & @MDAY & "-" & @YEAR & "_" & $horario & ".txt"
#EndRegion



#Region Alerting
;~ 	$RetryFrecuencyMin= IniRead("SAP101.ini", "Mail", "RetryFrecuencyMin", "1")
;~ 	$RetryFrecuencyMin=$RetryFrecuencyMin*60
;~ 	$DropTimeMin= IniRead("SAP101.ini", "Mail", "DropTimeMin", "1")
;~ 	$DropTimeMin=$DropTimeMin*60
#EndRegion

