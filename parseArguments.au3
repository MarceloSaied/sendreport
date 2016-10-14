#Region  -----------------------------  parse arguments ------------------------------
   if $CmdLine[0]=0 then
	  _TeeOutput($helpText,1)
	   exit 1
   endif
   if StringStripWS($CmdLine[1],4)="?"   then  ; output help
	  _TeeOutput($helpText,1)
	    exit 3
   endif
    if $CmdLine[0]<>4 then  ; if no 3 argumentes then exit
	  _TeeOutput("No enough arguments. <Usage>:  sendreport.exe -f <filename>  -r <mail@Address> ",1)
	  exit 5
   endif

   for $x=1 to $CmdLine[0]
	  Select
		 case $CmdLine[$x] = "-f"
			$x = $x + 1
			if StringLen($CmdLine[$x]) > 3 then   $FileToAtach=$CmdLine[$x]
		 case $CmdLine[$x] = "-r"
			$x = $x + 1
			if StringLen($CmdLine[$x]) > 3 then $EmailRecipient=$CmdLine[$x]
		 Case Else
			_TeeOutput("No enough arguments. <Usage>: sendreport.exe -f <filename>  -r <mail@Address>)  ",0)
			_TeeOutput("                              sendreport.exe ?  for help  ",0)
			_TeeOutput(@CRLF,0)
			exit 6
	  EndSelect
   Next

#EndRegion
#Region  ----------------------------- Main code ------------------------------------
		_TeeOutput('$EmailRecipient = ' & $EmailRecipient ,1)
		_TeeOutput('$FileToAtach = ' & $FileToAtach ,1)

#endregion



