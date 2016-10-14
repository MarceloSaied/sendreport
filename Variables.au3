#Region Base
	Global $oMyRet[2]
	Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
	$debugflag=0
#EndRegion

#Region 0) StartLog()
	Global $fileini="SendReport.ini"
	$ScripClass=""
	$logFile=""
	$horario=@HOUR & "_" & @MIN
	$logFileName=@ScriptDir & "\" & StringMid(@ScriptName,1,StringLen(@ScriptName)-4)  & "_" & @MON & "-" & @MDAY & "-" & @YEAR & "_" & $horario& ".log"
	$AttachFiles = @ScriptDir & "\" & StringMid(@ScriptName,1,StringLen(@ScriptName)-4)   & "_" & @MON & "-" & @MDAY & "-" & @YEAR & "_" & $horario & ".txt"
#EndRegion



#Region
	Global $FileToAtach=""
	Global $EmailRecipient=""

#EndRegion

