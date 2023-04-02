#include-once
#include <AutoItConstants.au3>
#include <FileConstants.au3>
#include <WinAPILocale.au3>
#include <Timers.au3>
; #INDEX# =======================================================================================================================
; Title .........: AutoIt Library, version 1.22
; Description....: Miscellaneous Functions
; Version history: 1.22		added: _WindowsScrollInactiveWindow() to read Windows setting to scroll an inactive window
;                  1.19		added: _ProcessRunsAlready() returns if a process is already running when the same process is started again (singleton), process name may include path, when no processes given @ScriptName is used
;                  			added: _ProcessInstances() gets number of instances of a process or processes, process name may include path, when no processes given @ScriptName is used
;                  			improved: _ProcessGetProcessId(), process name may include path
;                  			added: _CHMShow() to show CHM file (having manual, help info, etc.) located in @ScriptDir, file may have or not have the .chm extension
;                  			added: _CHMShowOnF1() to show CHM file (having manual, help info, etc.) located in @ScriptDir after pressing F1 on a GUI, file may have or not have the .chm extension
;                  1.18		added: _SoundGetWaveVolume() to get app volume of script, Windows Vista, 7, 8, 10 only
;                  1.17		added: _WindowsBuildNumber() and _WindowsUBRNumber() to read build and update build release numbers. Can be used to determine if Windows has been updated
;                  1.16		added: _FileGetSizeTimed() tries to get file size in a loop when FileGetSize() fails (when Windows isn't done yet) for instance just after downloading the file
;                  1.14		added: _ProcessGetProcessId() to get process id by name
;                  1.00		initial version
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;	_Console
;	_LogSwitch
;	_LogIsOn
;	_LogFile
;	_LogWrite
;	_LogDelete
;	_LogRead
;	_CHMShow
;	_CHMShowOnF1
;	_FileGetSizeTimed
;	_ProcessGetProcessId
;	_ProcessInstances
;	_ProcessRunsAlready
;	_RegistryRead
;	_WindowsVersion
;	_WindowsBuildNumber
;	_WindowsUBRNumber
;	_WindowsScrollInactiveWindow
;	_LocaleDecimal
;	_LocaleThousand
;	_SoundGetWaveVolume
; ===============================================================================================================================

; List of 21 functions
;	_Console						Debugging				Writes text or array of text to console with @CRLF
;	_LogSwitch						Debugging				Switches logging on or off
;	_LogIsOn						Debugging				Returns if logging is switched on
;	_LogFile						Debugging				Sets log file
;	_LogWrite						Debugging				Writes event to log file if logging is on
;	_LogDelete						Debugging				Deletes a log file
;	_LogRead						Debugging				Reads the content of a log file
;	_CHMShow						Help / Manual			Shows CHM file (having manual, help info, etc.) located in @ScriptDir, file may have or not have the .chm extension
;	_CHMShowOnF1					Help / Manual			Shows CHM file (having manual, help info, etc.) located in @ScriptDir after pressing F1 on a GUI, file may have or not have the .chm extension
;	_FileGetSizeTimed				File					Tries to get file size in a loop when FileGetSize() fails (when Windows isn't done yet) for instance just after downloading the file
;	_ProcessGetProcessId			Process information		Gets process id of given process name
;	_ProcessInstances				Process information		Gets number of instances of a process or processes, process name may include path
;	_ProcessRunsAlready				Process information		Returns if a process is already running when the same process is started again (singleton), process name may include path
;	_RegistryRead					System information		Reads key of registry taking in account a 64 bit Windows
;															Remark: For reading some keys from the registry administrator rights could be needed
;	_WindowsVersion					System information		Reads Windows version from registry
;	_WindowsBuildNumber				System information		Reads Windows build number from registry
;	_WindowsUBRNumber				System information		Reads Windows update build release number from registry
;	_WindowsScrollInactiveWindow	System information		Reads Windows setting to scroll an inactive window
;	_LocaleDecimal					System information		Returns decimal sign on local computer
;	_LocaleThousand					System information		Returns thousands sign on local computer
;	_SoundGetWaveVolume				System information		Returns app volume of script, Windows Vista, 7, 8, 10 only

; #FUNCTION# ====================================================================================================================
; Name...........: _Console
; Description....: Writes text or array of text to console with @CRLF
; Syntax.........: _Console($vText1 [,$vText2 = Default [,$vText3 = Default [,$vText4 = Default [,$vText5 = Default [,$vText6 = Default [,$vText7 = Default [,$vText8 = Default [,$vText9 = Default [,$vText10 = Default]]]]]]]]])
; Parameters.....: $vText1					- Text or array to write to console
;                  $vText2					- Text to write or boolean: True = Text on new line, False = Not
;                  $vText3 to 10			- Text 3 to 10 to write to console
; Return values..: None
; Modified.......:
; ===============================================================================================================================
Func _Console($vText1,$vText2 = Default,$vText3 = Default,$vText4 = Default,$vText5 = Default,$vText6 = Default,$vText7 = Default,$vText8 = Default,$vText9 = Default,$vText10 = Default)
	If IsArray($vText1) Then
		If UBound($vText1) > 0 Then
			For $nRow = 0 To UBound($vText1)-1
				If UBound($vText1,$UBOUND_COLUMNS) = 0 Then
					ConsoleWrite($vText1[$nRow])
				Else
					For $nColumn = 0 To UBound($vText1,$UBOUND_COLUMNS)-1
						ConsoleWrite($vText1[$nRow][$nColumn] & " ")
					Next
				EndIf
				If $vText2 = Default Or (IsBool($vText2) And $vText2) Then ConsoleWrite(@CRLF)
			Next
		EndIf
	Else
		If $vText2 = Default Then
			ConsoleWrite($vText1 & @CRLF)
		ElseIf $vText3 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF)
		ElseIf $vText4 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF)
		ElseIf $vText5 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF & $vText4 & @CRLF)
		ElseIf $vText6 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF & $vText4 & @CRLF & $vText5 & @CRLF)
		ElseIf $vText7 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF & $vText4 & @CRLF & $vText5 & @CRLF & $vText6 & @CRLF)
		ElseIf $vText8 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF & $vText4 & @CRLF & $vText5 & @CRLF & $vText6 & @CRLF & $vText7 & @CRLF)
		ElseIf $vText9 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF & $vText4 & @CRLF & $vText5 & @CRLF & $vText6 & @CRLF & $vText7 & @CRLF & $vText8 & @CRLF)
		ElseIf $vText10 = Default Then
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF & $vText4 & @CRLF & $vText5 & @CRLF & $vText6 & @CRLF & $vText7 & @CRLF & $vText8 & @CRLF & $vText9 & @CRLF)
		Else
			ConsoleWrite($vText1 & @CRLF & $vText2 & @CRLF & $vText3 & @CRLF & $vText4 & @CRLF & $vText5 & @CRLF & $vText6 & @CRLF & $vText7 & @CRLF & $vText8 & @CRLF & $vText9 & @CRLF & $vText10 & @CRLF)
		EndIf
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LogSwitch
; Description....: Switches logging on or off
; Syntax.........: _LogSwitch([$bLog = True])
; Parameters.....: $bLog = True				- True = Switch logging on, False = Switch off
; Return values..: None
; Modified.......:
; ===============================================================================================================================
Func _LogSwitch($bLog = True)
	Assign("_bLogSwitch",$bLog,$ASSIGN_FORCEGLOBAL)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LogIsOn
; Description....: Returns if logging is switched on
; Syntax.........: _LogIsOn()
; Parameters.....: None
; Return values..: True = Logging is on, False = Off
; Modified.......:
; ===============================================================================================================================
Func _LogIsOn()
	If IsDeclared("_bLogSwitch") <> $DECLARED_GLOBAL Or Not Eval("_bLogSwitch") Then Return False
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LogFile
; Description....: Sets log file
; Syntax.........: _LogFile($sLogFile)
; Parameters.....: $sLogFile				- File for logging info
; Return values..: None
; Modified.......:
; ===============================================================================================================================
Func _LogFile($sLogFile = @MyDocumentsDir & "\Log.txt")
	If StringLen($sLogFile) = 0 Then $sLogFile = @MyDocumentsDir & "\Log.txt"
	Assign("_sLogFile",$sLogFile,$ASSIGN_FORCEGLOBAL)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LogWrite
; Description....: Writes event to log file if logging is on
; Syntax.........: _LogWrite($sText [,$sLogFile = ""])
; Parameters.....: $sText					- Event to log in format: current date time $sText
;                  $sLogFile				- Log file to write to, if empty then file set by _LogFile is used
; Return values..: True = Event has been written, False = Logging is off or couldn't open log file
; Modified.......:
; ===============================================================================================================================
Func _LogWrite($sText,$sLogFile = "")
	If IsDeclared("_bLogSwitch") <> $DECLARED_GLOBAL Or Not Eval("_bLogSwitch") Then Return False	; logging is off, no writing
	If IsDeclared("_sLogFile") = $DECLARED_GLOBAL And StringLen($sLogFile) = 0 Then $sLogFile = Eval("_sLogFile")
	If StringLen($sLogFile) = 0 Then $sLogFile = @MyDocumentsDir & "\Log.txt"
	Local $hFileHandle = FileOpen($sLogFile,$FO_APPEND)
	If $hFileHandle = -1 Then Return False
	FileWrite($hFileHandle,@YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & " " & $sText & @CRLF)
	FileClose($hFileHandle)
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LogDelete
; Description....: Deletes a log file
; Syntax.........: _LogDelete([$sLogFile = ""])
; Parameters.....: $sLogFile				- Log file to delete, if empty then file set by _LogFile is used
; Return values..: True = File is deleted, False = Logging wasn't switched on during program run or file doesn't exist
; Modified.......:
; ===============================================================================================================================
Func _LogDelete($sLogFile = "")
	If IsDeclared("_sLogFile") = $DECLARED_GLOBAL And StringLen($sLogFile) = 0 Then $sLogFile = Eval("_sLogFile")
	If StringLen($sLogFile) = 0 Then $sLogFile = @MyDocumentsDir & "\Log.txt"
	If FileExists($sLogFile) Then
		FileDelete($sLogFile)
		Return True
	EndIf
	Return False
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LogRead
; Description....: Reads the content of a log file
; Syntax.........: _LogRead([$sLogFile = ""])
; Parameters.....: $sLogFile				- Log file to read, if empty then file set by _LogFile is used
; Return values..: Content of log file, if empty log file is empty or not readable
; Modified.......:
; ===============================================================================================================================
Func _LogRead($sLogFile = "")
	If IsDeclared("_sLogFile") = $DECLARED_GLOBAL And StringLen($sLogFile) = 0 Then $sLogFile = Eval("_sLogFile")
	If StringLen($sLogFile) = 0 Then $sLogFile = @MyDocumentsDir & "\Log.txt"
	If FileExists($sLogFile) Then
		Return FileRead($sLogFile)
	EndIf
	Return ""
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _CHMShow
; Description....: Shows CHM file (having manual, help info, etc.) located in @ScriptDir, file may have or not have the .chm extension
; Syntax.........: _CHMShow( $cCHMFile [, $cTopic = ""] )
; Parameters.....: $cCHMFile				- CHM (help manual) file
;                  $cTopic					- Topic (section) to show
; Return values..: None
; Modified.......:
; ===============================================================================================================================
Func _CHMShow($cCHMFile,$cTopic = "")
	ShellExecute(@WindowsDir & "\hh.exe","its:" & @ScriptDir & "\" & $cCHMFile & (StringLower(StringRight($cCHMFile,4)) = ".chm" ? "" : ".chm") & (StringLen($cTopic) = 0 ? "" : "::/" & $cTopic))
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _CHMShowOnF1
; Description....: Shows CHM file (having manual, help info, etc.) located in @ScriptDir after pressing F1 on a GUI, file may have or not have the .chm extension
; Syntax.........: _CHMShowOnF1( $cCHMFile [, $cTopic = ""] )
; Parameters.....: $cCHMFile				- CHM (help manual) file
;                  $cTopic					- Topic (section) to show
; Return values..: None
; Modified.......:
; ===============================================================================================================================
Func _CHMShowOnF1($cCHMFile,$cTopic = "")
	GUISetHelp(@WindowsDir & "\hh.exe its:" & @ScriptDir & "\" & $cCHMFile & (StringLower(StringRight($cCHMFile,4)) = ".chm" ? "" : ".chm") & (StringLen($cTopic) = 0 ? "" : "::/" & $cTopic))
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FileGetSizeTimed
; Description....: Tries to get file size in a loop when FileGetSize() fails (when Windows isn't done yet) for instance just after downloading the file
; Syntax.........: _FileGetSizeTimed($sFile [, $nTimeout = 5000 [, $nWaitTime = 100]] )
; Parameters.....: $sFile					- File to get size of
;                  $nTimeout				- Timeout in milliseconds, default 5000 ms
;                  $nWaitTime				- Time pause for checking again, default 100 ms
; Return values..: File size or 0 when failed or actually size is 0 (check @error)
; Error values...: @error = 1				- Failed to get size, you might need to increase $Timeout value
; Modified.......:
; ===============================================================================================================================
Func _FileGetSizeTimed($sFile,$nTimeout = 5000,$nWaitTime = 100)
	Local $nFileSize,$hFileSizeTimer = _Timer_Init(),$nError = 0

	While _Timer_Diff($hFileSizeTimer) < $nTimeout
		$nFileSize = FileGetSize($sFile)
		$nError = @error
		If $nError <> 0 Or $nFileSize = 0 Then	; check on error or 0 size which indicates that the FileGetSize() function is yet able to return correct size
			Sleep($nWaitTime)
			ContinueLoop
		EndIf
		Return $nFileSize
	WEnd
	If $nError = 0 Then
		Return 0
	Else
		SetError(1,0,0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _ProcessGetProcessId
; Description....: Gets process id of given process name, process name may include path
; Syntax.........: _ProcessGetProcessId( $cProcessName )
; Parameters.....: $cProcessName			- Name of process (executable, may include path)
; Return values..: Process Id or 0 if not found or -1 if process list couldn't be built
; Modified.......:
; ===============================================================================================================================
Func _ProcessGetProcessId($cProcessName)
	Local $aProcessDirs = StringSplit($cProcessName,":\")
	If $aProcessDirs[0] > 0 Then $cProcessName = $aProcessDirs[$aProcessDirs[0]]
	Local $aProcessesList = ProcessList($cProcessName)
	If @error = 1 Then Return -1
	If $aProcessesList[0][0] = 0 Then Return 0
	Return $aProcessesList[1][1]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _ProcessRunsAlready
; Description....: Returns if a process is already running when the same process is started again (singleton), process name may include path, when no processes given @ScriptName is used
; Syntax.........: _ProcessRunsAlready( [$cProcessName1 [,$cProcessName2] [,$cProcessName3]]] )
; Parameters.....: $cProcessName1			- Name of process (executable, may include path)
;                  $cProcessName2			- Equivalent name of process (executable, may include path), for instance 32 or 64 bit version
;                  $cProcessName3			- Equivalent name of process (executable, may include path), for instance 32 or 64 bit version
; Return values..: True if process is already running thus 2 or more times as the current instances is also running for checking
; Modified.......:
; ===============================================================================================================================
Func _ProcessRunsAlready($cProcessName1 = "",$cProcessName2 = "",$cProcessName3 = "")
	If StringLen($cProcessName1 & $cProcessName2 & $cProcessName3) = 0 Then	Return _ProcessInstances(@ScriptName) > 1		; check for more than 1 @ScriptName process
	; strip paths from process names
	Local $aProcessDirs = StringSplit($cProcessName1,":\")
	If $aProcessDirs[0] > 0 Then $cProcessName1 = $aProcessDirs[$aProcessDirs[0]]
	$aProcessDirs = StringSplit($cProcessName2,":\")
	If $aProcessDirs[0] > 0 Then $cProcessName2 = $aProcessDirs[$aProcessDirs[0]]
	$aProcessDirs = StringSplit($cProcessName3,":\")
	If $aProcessDirs[0] > 0 Then $cProcessName3 = $aProcessDirs[$aProcessDirs[0]]
	; filter out same process names
	If StringLower($cProcessName3) = StringLower($cProcessName2) Then $cProcessName3 = ""
	If StringLower($cProcessName2) = StringLower($cProcessName1) Then $cProcessName2 = ""
	; check if runs twice or more
	If StringLen($cProcessName3) = 0 And StringLen($cProcessName2) = 0 Then Return _ProcessInstances($cProcessName1) > 1	; check for more than 1 $cProcessName1 process
	If StringLen($cProcessName3) = 0 Then Return _ProcessInstances($cProcessName1)+_ProcessInstances($cProcessName2) > 1	; check for more than 1 $cProcessName1 + $cProcessName2 process
	Return _ProcessInstances($cProcessName1)+_ProcessInstances($cProcessName2)+_ProcessInstances($cProcessName3) > 1		; check for more than 1 $cProcessName1 + $cProcessName2 + $cProcessName3 + process
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _ProcessInstances
; Description....: Gets number of instances of a process or processes, process name may include path, when no processes given @ScriptName is used
; Syntax.........: _ProcessInstances( [$cProcessName1 [,$cProcessName2] [,$cProcessName3] [,$cProcessName4] [,$cProcessName5]]]]] )
; Parameters.....: $cProcessName1			- Name of a process (executable, may include path)
;                  $cProcessName2			- Name of a process (executable, may include path)
;                  $cProcessName3			- Name of a process (executable, may include path)
;                  $cProcessName4			- Name of a process (executable, may include path)
;                  $cProcessName5			- Name of a process (executable, may include path)
; Return value...: Number of instances of a process or processes
; Modified.......:
; ===============================================================================================================================
Func _ProcessInstances($cProcessName1 = "",$cProcessName2 = "",$cProcessName3 = "",$cProcessName4 = "",$cProcessName5 = "")
	Local $aProcesses = [$cProcessName1,$cProcessName2,$cProcessName3,$cProcessName4,$cProcessName5],$aProcessDirs,$aProcess,$iCountInstances = 0,$cProcessName,$bAlreadyChecked

	For $iProcess = 0 To UBound($aProcesses)-1
		If $iProcess = 0 And StringLen($aProcesses[$iProcess]) = 0 Then $aProcesses[$iProcess] = @ScriptName
		If StringLen($aProcesses[$iProcess]) = 0 Then ContinueLoop
		; strip path from process name
		$aProcessDirs = StringSplit($aProcesses[$iProcess],":\")
		If $aProcessDirs[0] = 0 Then ContinueLoop
		$cProcessName = $aProcessDirs[$aProcessDirs[0]]
		; check an identical given process name only once
		$bAlreadyChecked = False
		If $iProcess > 0 Then
			For $nCheckProcess = 0 To $iProcess-1
				If $aProcesses[$nCheckProcess] = $aProcessDirs[$aProcessDirs[0]] Then
					$bAlreadyChecked = True
					ExitLoop
				EndIf
			Next
		EndIf
		If $bAlreadyChecked Then ContinueLoop	; process has been checked already, don't count
		; count processes
		$aProcess = ProcessList($aProcessDirs[$aProcessDirs[0]])
		If @error Then ContinueLoop
		If $aProcess[0][0] > 0 Then $iCountInstances += $aProcess[0][0]
	Next
	Return $iCountInstances
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _RegistryRead
; Description....: Reads key of registry taking in account a 64 bit Windows
; Syntax.........: _RegistryRead($sKey, $sValueName)
; Parameters.....: $sKey					- Registry key
;                  $sValueName				- Value name to read
; Remarks........: For reading some keys from the registry administrator rights could be needed
; Return values..: Read key value
; Modified.......:
; ===============================================================================================================================
Func _RegistryRead($sKey,$sValueName)
	If @OSArch <> "X86" Then $sKey = StringLeft($sKey,StringInStr($sKey,"\")-1) & "64\" & StringTrimLeft($sKey,StringInStr($sKey,"\"))
	Return RegRead($sKey,$sValueName)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WindowsVersion
; Description....: Reads Windows version from registry
; Syntax.........: _WindowsVersion()
; Parameters.....: None
; Return values..: Windows version number
; Modified.......:
; ===============================================================================================================================
Func _WindowsVersion()
	Return Number(_RegistryRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\","CurrentVersion"))
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WindowsBuildNumber
; Description....: Reads Windows build number from registry
; Syntax.........: _WindowsBuildNumber()
; Parameters.....: None
; Return values..: Windows build number
; Modified.......:
; ===============================================================================================================================
Func _WindowsBuildNumber()
	Return Number(_RegistryRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\","CurrentBuildNumber"))
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WindowsUBRNumber
; Description....: Reads Windows update build release number from registry
; Syntax.........: _WindowsUBRNumber()
; Parameters.....: None
; Return values..: Windows update build release number
; Modified.......:
; ===============================================================================================================================
Func _WindowsUBRNumber()
	Return Number(_RegistryRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\","UBR"))
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WindowsScrollInactiveWindow
; Description....: Reads Windows setting to scroll an inactive window
; Syntax.........: _WindowsScrollInactiveWindow()
; Parameters.....: None
; Return values..: True if scrolling an inactive window is allowed
; Modified.......:
; ===============================================================================================================================
Func _WindowsScrollInactiveWindow()
	Return BitAND(Number(_RegistryRead("HKEY_CURRENT_USER\Control Panel\Desktop","MouseWheelRouting")),2) = 2
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LocaleDecimal
; Description....: Returns decimal sign on local computer
; Syntax.........: _LocaleDecimal()
; Parameters.....: None
; Return values..: Decimal sign
; Modified.......:
; ===============================================================================================================================
Func _LocaleDecimal()
	Local $LocaleID = _WinAPI_GetUserDefaultLCID()
	Return _WinAPI_GetLocaleInfo($LocaleID,$LOCALE_SDECIMAL)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _LocaleThousand
; Description....: Returns thousands sign on local computer
; Syntax.........: _LocaleThousand($i)
; Parameters.....: None
; Return values..: Thousands sign
; Modified.......:
; ===============================================================================================================================
Func _LocaleThousand()
	Local $LocaleID = _WinAPI_GetUserDefaultLCID()
	Return _WinAPI_GetLocaleInfo($LocaleID,$LOCALE_STHOUSAND)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _SoundGetWaveVolume
; Description....: Returns app volume of script, Windows Vista, 7, 8, 10 only
; Syntax.........: _SoundGetWaveVolume([$iValueOnError = -1])
; Parameters.....: $iValueOnError			- Value to return when an error occurs
; Return values..: App volume of script or $iValueOnError at an error
; Error values...: @error = 1				- Unable to create Struct
;                  @error = 2				- Dll file not found
;                  @error = 3				- Wrong call so not on Windows Vista, 7, 8 or 10
;                  @error = 4				- Internal error, array not returned
;                  @error = 5				- Volume wasn't received
;                  @error = 6				- Volume couldn't read
; Modified.......:
; ===============================================================================================================================
Func _SoundGetWaveVolume($iValueOnError = -1)
	Local $LPDWORD,$aMMRESULT,$iVolume

	$LPDWORD = DllStructCreate("dword")
	If @error <> 0 Then
		SetError(1)												; 1 = unable to create Struct
		Return $iValueOnError
	EndIf
	; get app volume of this script
	$aMMRESULT = DllCall("winmm.dll","uint","waveOutGetVolume","ptr",0,"long_ptr",DllStructGetPtr($LPDWORD))
	Switch @error
		Case 1
			SetError(2)											; 2 = dll file not found
			Return $iValueOnError
		Case 2,3,4,5
			SetError(3)											; 3 = wrong call so not on Windows Vista, 7, 8 or 10
			Return $iValueOnError
	EndSwitch
	If not IsArray($aMMRESULT) Then
		SetError(4)												; 4 = internal error, array not returned
		Return $iValueOnError
	EndIf
	If $aMMRESULT[0] <> 0 Then
		SetError(5)												; 5 = volume wasn't received
		Return $iValueOnError
	EndIf
	$iVolume = DllStructGetData($LPDWORD,1)
	If @error <> 0 Then
		SetError(6)												; 6 = volume couldn't read
		Return $iValueOnError
	EndIf
	Return Round(100*$iVolume/4294967295)						; return in range 0 to 100 as SoundSetWaveVolume()
EndFunc