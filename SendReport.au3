#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=sendreport.ico
#AutoIt3Wrapper_Outfile=SendReport.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Comment="By Marcelo N. Saied "
#AutoIt3Wrapper_Res_Description="Send report attached by mail"
#AutoIt3Wrapper_Res_Fileversion=0.1.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=0.1
#AutoIt3Wrapper_Res_LegalCopyright=Marcelo N. Saied
#AutoIt3Wrapper_Res_Field=ProductVersion|Version 1.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Field=Productname|SendReport
#AutoIt3Wrapper_Res_Field=Manufacturer |Marcelo N. Saied. All rights reserved
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/mergeonly
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;===================================================================
;===================================================================
;===                                                             ===
;===   Name   = SendReport                                       ===
;===                                                             ===
;===   Description: Sewnd mail report with atachment             ===
;===                                                             ===
;===   Designed By:                                              ===
;===                                                             ===
;===                                                             ===
;===   Author: Marcelo N. Saied                                  ===
;===           marcelosaied@gmail.com                            ===
;===                                                             ===
;===   Automation and Scripting Language: AUTOIT  v3.3.10.1      ===
;===           http://www.autoitscript.com/site/                 ===
;===                                                             ===
;===   Created on: October , 2016                                ===
;===                                                             ===
;===================================================================
;===================================================================
; -----------  Variables
   #include <Variables.au3>
   #include <secret\secrets.au3>
; ------------ Modules
   #include <UDF_System.au3>
   #include <UDF_local.au3>
; --------- general functions
	#include <Genaral Functions.au3>
#region 0)
	StartLog()   ; init log file
#EndRegion
	#region version number
		$ScriptVersion="V1"
		_FileWriteLog($logFile,'Send Report Version =' &   $ScriptVersion & @crlf)
		ConsoleWrite('Send Report Version =' &  $ScriptVersion & @crlf)
		_FileWriteLog($logFile,'####################################################################################' & @crlf)
	#EndRegion
	#region Check ini file
		checkINIfile()
	#EndRegion


;~ 		SendMail(" STARTUP FAILED "," Database is Down" )

_FileWriteLog($logFile,'End of activities' & @TAB & @TAB  & _NowCalc() & @crlf)
_FileWriteLog($logFile,'~~~~~~~~~~~~~~~~~~~~~~      END         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~' & @crlf)
FileClose($logFile)