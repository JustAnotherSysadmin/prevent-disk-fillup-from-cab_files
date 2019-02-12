#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         John Lucas

 Script Purpose:
   Prevent and clean temporary files that fill up the hard disk located
   in c:\windows\temp named "cab_xxxxx." (where xxxxx is a number) as well as
   removing the log files that are too large (>2GB in size) which cause
   the cab file creation process crash. Logs that casue the whole mess are
   located in c:\windows\Logs\CBS\CbsPersist_YYYYMMDDHHMMSS.log
   To save space, we remove files that are greater than 1GB in size.

   Here is a good background description of the problem from  https://www.bigfix.me/fixlet/details/11930
   (Text copied from the above link)
	  We've had repeated instances where a Windows 7 x64 client runs out of hard drive space, and
	  found that C:\Windows\TEMP is being consumed with hundreds of files with names following
	  the pattern "cab_XXXX_X", generally 100 MB each, and these files are constantly generated
	  until the system runs out of space.  Upon removing the files & rebooting, the files start
	  being generated again.

	  This appears to be caused by large Component-Based Servicing logs.  These are
	  stored at C:\Windows\Logs\CBS.  The current log file is named "cbs.log".  When "cbs.log" reaches
	  a certain size, a cleanup process renames the log to "CbsPersist_YYYYMMDDHHMMSS.log" and then
	  attempts to compress it into a .cab file.

	  However, when the cbs.log reaches a size of 2 GB before that cleanup process compresses it, the
	  file is to large to be handled by the makecab.exe utility.  The log file is renamed
	  to CbsPersist_date_time.log, but when the makecab process attempts to compress it the process
	  fails (but only after consuming some 100 MB under \Windows\Temp).  After this, the cleanup process
	  runs repeatedly (approx every 20 minutes in my experience).  The process fails every time, and
	  also consumes a new ~ 100 MB in \Windows\Temp before dying.  This is repeated until the system
	  runs out of drive space.

	  This Action will remove the log files that are too large to be handled by makecab.exe.  Any
	  \Windows\Logs\CBS\CbsPersist_XXX.log file larger than 1,000 MB is removed, as well as
	  any cab_ files from C:\Windows\Temp.

Script Function:
   1. Check to see if any %SystemRoot%\temp\cab_*. files exist.  If so, log and delete them.
   2. Check to see if any %SystemRoot%\Logs\CBS\CbsPersist_*.log files exist that are
	   bigger than 1GB.  If so, log and delete them.
   3. If either of above checks are true, then also:
     a) Delete %SystemRoot%\Logs\CBS\CbsPersist_*.log   older than X days
	 b) Delete %SystemRoot%\Logs\CBS\CbsPersist_*.cab   older than X days
   4. Write out log file if cab_* files are detected or if large log files exist

#ce ----------------------------------------------------------------------------


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;      ___   _   _    ____   _       _   _   ____    _____   ____
;     |_ _| | \ | |  / ___| | |     | | | | |  _ \  | ____| / ___|
;      | |  |  \| | | |     | |     | | | | | | | | |  _|   \___ \
;      | |  | |\  | | |___  | |___  | |_| | | |_| | | |___   ___) |
;     |___| |_| \_|  \____| |_____|  \___/  |____/  |_____| |____/
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#include <Constants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <Date.au3>
#include <Array.au3>
;#include <WinAPIFiles.au3>
;#include <Inet.au3>




;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;      _____   _   _   _   _    ____   _____   ___    ___    _   _   ____
;     |  ___| | | | | | \ | |  / ___| |_   _| |_ _|  / _ \  | \ | | / ___|
;     | |_    | | | | |  \| | | |       | |    | |  | | | | |  \| | \___ \
;     |  _|   | |_| | | |\  | | |___    | |    | |  | |_| | | |\  |  ___) |
;     |_|      \___/  |_| \_|  \____|   |_|   |___|  \___/  |_| \_| |____/
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Func _GetDOSOutput($sCommand)
    Local $iPID, $sOutput = ""

    $iPID = Run('"' & @ComSpec & '" /c ' & $sCommand, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
    While 1
        $sOutput &= StdoutRead($iPID, False, False)
        If @error Then
            ExitLoop
        EndIf
        Sleep(10)
    WEnd
    Return $sOutput
 EndFunc   ;==>_GetDOSOutput

#comments-start
; list all files returned ---- not used
Func _IHPfilesExistList($sPathAndFile)
    ; Assign a Local variable the search handle of all files in the current directory.
    Local $hSearch = FileFindFirstFile($sPathAndFile)

	MsgBox($MB_SYSTEMMODAL, "sPathAndFile", $sPathAndFile)

    ; Check if the search was successful, if not display a message and return False.
    If $hSearch = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
        Return False
    EndIf

    ; Assign a Local variable the empty string which will contain the files names found.
    Local $sFileName = "", $iResult = 0

    While 1
        $sFileName = FileFindNextFile($hSearch)
        ; If there is no more file matching the search.
        If @error Then ExitLoop

        ; Display the file name.
        $iResult = MsgBox(BitOR($MB_SYSTEMMODAL, $MB_OKCANCEL), "", "File: " & $sFileName)
        If $iResult <> $IDOK Then ExitLoop ; If the user clicks on the cancel/close button.
    WEnd

    ; Close the search handle.
    FileClose($hSearch)
EndFunc   ;==> _IHPfilesExistList
#comments-end

Func _IHPfilesExist($sPathAndFile)
    ; Assign a Local variable the search handle of all files in the current directory.
    Local $hSearch = FileFindFirstFile($sPathAndFile)

	;MsgBox($MB_SYSTEMMODAL, "sPathAndFile", $sPathAndFile)

    ; Check if the search was successful, if not display a message and return False.
    If $hSearch = -1 Then
        ;MsgBox($MB_SYSTEMMODAL, $sPathAndFile, "Error: No files/directories matched the search pattern.")
        Return False
    EndIf

    If Not($hSearch = -1) Then
        ;MsgBox($MB_SYSTEMMODAL, $sPathAndFile, "Files/directories DID match the search pattern.")
        Return True
    EndIf

    ; Close the search handle.
    FileClose($hSearch)
EndFunc   ;==> _IHPfilesExist

Func _IHPLargeCbsPersistLogExist()
   $sRawOutput = (_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.log /C "cmd /c if @fsize gtr 1048576000 echo Found"'))
   ;$sRawOutput = (_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.log /C "cmd /c if @fsize gtr 1040 echo Found"'))
   ;$sLargeFound = StringRegExp("Found", $sRawOutput, 1)
   $sLargeFound = StringRegExp($sRawOutput, "(?s)Found", 0)
   ;MsgBox($MB_SYSTEMMODAL, "In Func _IHPLargeCbsPersistLogExist()", "$sRawOutput: " & $sRawOutput)
   ;MsgBox($MB_SYSTEMMODAL, "In Func _IHPLargeCbsPersistLogExist()", "$sLargeFound: " & $sLargeFound)
   ;If ($sLargeFound) Then
   ;   Return True
   ;Else
   ;   Return False
   ;EndIf
   Return $sLargeFound
EndFunc


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;      __  __      _      ___   _   _
;     |  \/  |    / \    |_ _| | \ | |
;     | |\/| |   / _ \    | |  |  \| |
;     | |  | |  / ___ \   | |  | |\  |
;     |_|  |_| /_/   \_\ |___| |_| \_|
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


; -----------------------------------------------------------------------------------------------
; Declairation of Globals and misc stuff to get things going
; -----------------------------------------------------------------------------------------------

; This next variable is where we put the output of everything
Global $gListOfStuff = ""

; Create a variable for dos enviornment variable
Global $gSystemRoot = EnvGet("SystemRoot")

; At the very end of this script, will we write out a log file?
; This is programitically set to True if we actually delete anything.
Global $gWriteLogFile = False

; Default Log file path and name
$gIHPLogFile = "c:\scripts\prevent-disk-fillup-from-cab_files.txt"

; Used if the user specifies a command line argument [program.exe] /log C:\path-to\logfile.Log
$gIHPLogFileCmdLine = ""
;$gIHPLogFileCmdLineDetected = False

; Was this script run from scheduledtasks?  If so, log it when command line parameters are checked.
$gIHPscriptRunFromScheduledTasksMsg = ""

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;   Command line parameters

; Deal with the command line parameters
If Not( $cmdLine[0] = 0 ) Then
   ;MsgBox($MB_SYSTEMMODAL, "$cmdLine parameters number of elements", $cmdLine[0])
   For $i = 1 To UBound($cmdLine) - 1 Step 1
      ;MsgBox($MB_SYSTEMMODAL, "Array", $cmdLine[$i])
	  $sMsg = ""
	  $sTMP = $cmdLine[$i]
	  Switch ($sTMP)
	     Case 1 To 8
		    $sMsg &= "You passed a number that is between 1 and 8"
		 Case "/Silent"
		    $sMsg = ($sTMP & " is golden")
		 Case "/log"
		    ;$sMsg = ($sTMP & " command line switch.  Log file is: " & $cmdLine[$i + 1] & @CRLF)
			$gIHPLogFileCmdLine = $cmdLine[$i + 1]
			;$gIHPLogFileCmdLineDetected = True
		 Case "/scheduledtasks"
			$gIHPscriptRunFromScheduledTasksMsg &= "( run with command line Switch /scheduledtasks )"
         Case "/help","help","verbose","-v","/v","/?","-help"
			$sMsg  = "/help -> display this box" & @CRLF
			$sMsg &= "/log  [log file]  -> change the location of the log output" & @CRLF
			$sMsg &= "/scheduledtasks   -> add to log entry that the script was run automaticaly" & @CRLF
			$sMsg &= "  example:  [program.exe]  /log c:\scripts\prevent-disk-fillup-from-cab_files.log" & @CRLF
		 Case Else
			;if Not($gIHPLogFileCmdLineDetected) Then
			;   $sMsg = "you have reached the default switch case Else"
			;EndIf
	  EndSwitch
	  If Not( $sMsg = "" ) Then
		 MsgBox($MB_SYSTEMMODAL, "Command line parameter detected", $sMsg)
	  EndIf
   Next
   ;MsgBox($MB_SYSTEMMODAL, "endfor", "out of for loop")
EndIf



;;;;;;;;;;;;;;;;;;;;;
; Purpose of this next If statement is to copy the log filepath passed as command line input argument via /Log

; debug point
;MsgBox($MB_SYSTEMMODAL, "$gIHPLogFile before If", "$gIHPLogFile before If  " & $gIHPLogFile)
;MsgBox($MB_SYSTEMMODAL, "$gIHPLogFileCmdLine before If", "$gIHPLogFileCmdLine before If  " & $gIHPLogFileCmdLine)

If Not( $gIHPLogFileCmdLine = "" ) Then
  $gIHPLogFile = $gIHPLogFileCmdLine
  ;MsgBox($MB_SYSTEMMODAL, "$gIHPLogFile in If", "$gIHPLogFile in If  " & $gIHPLogFile)
  ;MsgBox($MB_SYSTEMMODAL, "$gIHPLogFileCmdLine in If", "$gIHPLogFileCmdLine in If  " & $gIHPLogFileCmdLine)
EndIf


; header info to top of log file (will only be written to disk if the $gWriteLogFile is made true
$gListOfStuff &= ("--------------------------------------------------------------------------------------------------------------------" & @CRLF)
$gListOfStuff &= _NowCalcDate() & "  " & _NowTime(5) & " - " & EnvGet("ComputerName") & @CRLF
$gListOfStuff &= ("-------------------------------------------------" & @CRLF)

If Not($gIHPscriptRunFromScheduledTasksMsg = "" ) Then
   $gListOfStuff &= $gIHPscriptRunFromScheduledTasksMsg & @CRLF
EndIf


; debug point
;MsgBox($MB_SYSTEMMODAL, "dates", $gListOfStuff)

; -----------------------------------------------------------------------------------------------
;  Script Function: 1. Check to see if any %SystemRoot%\temp\cab_*. files exist.  If so, log and delete them.
; -----------------------------------------------------------------------------------------------
If (_IHPfilesExist($gSystemRoot & "\temp\cab_*.")) Then
    ; debug point
    ;MsgBox($MB_SYSTEMMODAL, "cab_*.", "Files found")

	;Allow log to be written to a file
	$gWriteLogFile = True

    ; Before doing anything, lets clear up some space for breathing room
	;$gListOfStuff &= ("Removing " & $gSystemRoot & "\temp\cab_*." & @CRLF )
	$gListOfStuff &= ("Removing " & $gSystemRoot & "\temp\cab_*." )
    $gListOfStuff &= (_GetDOSOutput("del /f /s %SystemRoot%\temp\cab_*.") )
EndIf

; debug point
;If Not(_IHPfilesExist($gSystemRoot & "\temp\cab_*.")) Then
;    MsgBox($MB_SYSTEMMODAL, "cab_*.", "Files not found")
;EndIf

; -----------------------------------------------------------------------------------------------
; Script Function: 2. Check to see if any %SystemRoot%\Logs\CBS\CbsPersist_*.log files exist
;                     that are bigger than 1GB.  If so, log and delete them.
; -----------------------------------------------------------------------------------------------

;If (_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.log /C "cmd /c if @fsize gtr 1048576000 echo @path"') ) Then

If (_IHPLargeCbsPersistLogExist()) Then
	; debug point
	;MsgBox($MB_SYSTEMMODAL, "Logs greater than 1GB", "Files found--pingpinghere")

	;Allow log to be written to a file
	$gWriteLogFile = True

	;$gListOfStuff &= ("Removing " & $gSystemRoot & "\Logs\CBS\CbsPersist_*.log that are larger than 1GB"	& @CRLF )
	$gListOfStuff &= ("Removing " & $gSystemRoot & "\Logs\CBS\CbsPersist_*.log that are larger than 1GB")
    $gListOfStuff &= (_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.log /C "cmd /c if @fsize gtr 1048576000 del /s /f @path"') )

	;origial code
	;$gListOfStuff &= ("Del ...\Logs\CBS\CbsPersist_*.log larger than 1GB" & @CRLF )
    ;$gListOfStuff &= (_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.log /C "cmd /c if @fsize gtr 1048576000 echo /s /f @path"') )
EndIf

; debug point
;If Not(_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.log /C "cmd /c if @fsize gtr 1048576000 echo /s /f @path"') ) Then
;    MsgBox($MB_SYSTEMMODAL, "No logs greater than 1GB", "No Files found--pingpinghere2")
;EndIf

If ( $gWriteLogFile = True ) Then
   ;MsgBox($MB_SYSTEMMODAL, "YouAreHere", "$gWriteLogFile: " & $gWriteLogFile)
   ; -----------------------------------------------------------------------------------------------
   ;$gListOfStuff &= ("Removing " & $gSystemRoot & "\Logs\CBS\CbsPersist_*.log older than 15 days" & @CRLF )
   $gListOfStuff &= ("Removing " & $gSystemRoot & "\Logs\CBS\CbsPersist_*.log older than 15 days" )
   $gListOfStuff &= (_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.log /D -"15" /C "cmd /c del /s /f @path"') )

   ; -----------------------------------------------------------------------------------------------
   ;$gListOfStuff &= ("Removing " & $gSystemRoot & "\Logs\CBS\CbsPersist_*.cab older than 15 days" & @CRLF )
   $gListOfStuff &= (@CRLF & "Removing " & $gSystemRoot & "\Logs\CBS\CbsPersist_*.cab older than 15 days" )
   $gListOfStuff &= (_GetDOSOutput('forfiles /p %SystemRoot%\Logs\CBS\ /M CbsPersist_*.cab /D -"15" /C "cmd /c del /s /f @path"') )
EndIf

$gListOfStuff &= (@CRLF & "--- End of run --- " & _NowCalcDate() & "  " & _NowTime(5) & @CRLF & @CRLF)

;FileDelete ( "c:\scripts\prevent-disk-fillup-from-cab_files2.txt" )

; If we deleted anything, enter into the if statement and append to the log file
If ($gWriteLogFile = True) Then
   FileWrite  ( $gIHPLogFile, $gListOfStuff )
EndIf

