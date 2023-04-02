#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_HiDpi=Y
#AutoIt3Wrapper_Icon=Images\logo.ico
#AutoIt3Wrapper_Outfile=Macronize32.exe
#AutoIt3Wrapper_Outfile_x64=Macronize.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=Macronize is a macro builder for automating your programs by creating macro's on mouse clicks, key presses and other automation actions.
#AutoIt3Wrapper_Res_Description=Macronize, automating your computer
#AutoIt3Wrapper_Res_Fileversion=1.1.1.0
#AutoIt3Wrapper_Res_ProductName=Macronize
#AutoIt3Wrapper_Res_ProductVersion=1.1.1.0
#AutoIt3Wrapper_Res_LegalCopyright=P.E. Verbeek
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_File_Add=Readme.txt
#AutoIt3Wrapper_Res_File_Add=Help\Macronize.chm
#AutoIt3Wrapper_Res_File_Add=Images\important.ico
#AutoIt3Wrapper_Res_File_Add=Images\info.ico
#AutoIt3Wrapper_Res_File_Add=Images\settings.png
#AutoIt3Wrapper_Res_File_Add=Images\questionmark.png
#AutoIt3Wrapper_Res_File_Add=Images\new.png
#AutoIt3Wrapper_Res_File_Add=Images\open.png
#AutoIt3Wrapper_Res_File_Add=Images\save.png
#AutoIt3Wrapper_Res_File_Add=Images\saveas.png
#AutoIt3Wrapper_Res_File_Add=Images\close.png
#AutoIt3Wrapper_Res_File_Add=Images\arrowup.png
#AutoIt3Wrapper_Res_File_Add=Images\arrowdown.png
#AutoIt3Wrapper_Res_File_Add=Images\minus.png
#AutoIt3Wrapper_Res_File_Add=Images\plus.png
#AutoIt3Wrapper_Res_File_Add=Images\edit.png
#AutoIt3Wrapper_Res_File_Add=Images\duplicate.png
#AutoIt3Wrapper_Res_File_Add=Images\detect.png
#AutoIt3Wrapper_Res_File_Add=Images\pointer.png
#AutoIt3Wrapper_Res_File_Add=Images\enable.png
#AutoIt3Wrapper_Res_File_Add=Images\disable.png
#AutoIt3Wrapper_Res_File_Add=Images\play.png
#AutoIt3Wrapper_Res_File_Add=Images\record.png
#AutoIt3Wrapper_Res_File_Add=Images\visible.png
#AutoIt3Wrapper_Res_File_Add=Images\invisible.png
#AutoIt3Wrapper_Res_File_Add=Images\viewlist.png
#AutoIt3Wrapper_Res_File_Add=Images\viewreport.png
#AutoIt3Wrapper_Res_File_Add=Images\viewgrid.png
#AutoIt3Wrapper_Res_File_Add=Images\invertcolor.png
#AutoIt3Wrapper_Res_File_Add=Images\numbers.png
#AutoIt3Wrapper_Res_File_Add=Images\disabled.bmp
#AutoIt3Wrapper_Res_File_Add=Images\mark.bmp
#AutoIt3Wrapper_Res_File_Add=Images\disabledmark.bmp
#AutoIt3Wrapper_Res_File_Add=Images\error.bmp
#AutoIt3Wrapper_Res_File_Add=Images\disablederror.bmp
#AutoIt3Wrapper_Res_File_Add=Sounds\ping.wav
#AutoIt3Wrapper_Res_File_Add=Sounds\metal.wav
#AutoIt3Wrapper_Res_File_Add=Sounds\bounce.wav
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "..\AutoIt Library\Window, Screen, Mouse and Control.au3"
#include "..\AutoIt Library\Lists.au3"
#include "..\AutoIt Library\Dialogue.au3"
#include "..\AutoIt Library\String and File String.au3"
#include "..\AutoIt Library\Logics and Math.au3"
#include "..\AutoIt Library\Miscellaneous.au3"
#include "_inputmask.au3"
#include <WinAPI.au3>
#include <GDIPlus.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstants.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include <GuiListView.au3>
#include <GuiImageList.au3>
#include <GuiMenu.au3>
#include <GuiDateTimePicker.au3>
#include <Array.au3>
#include <Misc.au3>
#include <String.au3>
#include <Date.au3>
#include <Sound.au3>
#include <Timers.au3>
#include <ScreenCapture.au3>
#include <Process.au3>

; set dpi awareness when not compiled
If Not (@Compiled ) Then
	If @OSVersion = "WIN_10" Then
		DllCall("User32.dll","bool","SetProcessDpiAwarenessContext","HWND","DPI_AWARENESS_CONTEXT"-2)
	ElseIf @OSVersion = "WIN_81" Then
		DllCall("User32.dll","bool","SetProcessDPIAware")
	EndIf
EndIf

; GetWinState Constants
;~ Global Const $WIN_STATE_EXISTS = 1		; Window exists
;~ Global Const $WIN_STATE_VISIBLE = 2		; Window is visible
;~ Global Const $WIN_STATE_ENABLED = 4		; Window is enabled
;~ Global Const $WIN_STATE_ACTIVE = 8		; Window is active
;~ Global Const $WIN_STATE_MINIMIZED = 16	; Window is minimized
;~ Global Const $WIN_STATE_MAXIMIZED = 32	; Window is maximized
Global Const $ProgramName = "Macronize",$FileExtension = ".mcz",$SettingsFile = @ScriptDir & "\" & $ProgramName & ".ini", _
			 $Version = "1.1.1.0", $ReleaseDate = " 29th of December 2022",$MaxRecentFiles = 9,$FileFailSave = "_FailSave_", _
			 $UndoTempFile = @TempDir & "\_Undo" & $ProgramName & @AutoItPID & "-",$FileAttributeDelimiter = Chr(14),$HelpFile = $ProgramName & ".chm"
Global $Themes = [[0x333333,0xE0E0E0,0xE0E0E0,0x000000],[0x222233,0xD0D0E0,0xD0D0E0,0x000000]]	; background, text, box and text input color
Global $ThemesTitles = ["Windows colors","Dark colors","Dark blue colors"]
Global $MainWindowActions = ["Show window","Minimize window","Hide window"]
Global $MacroStepTypes = [["Mouse click","Click a mouse button on a position or other mouse action",0x000080,0xC0C0FF,"typemouse.htm"],["Key press","Press a key or keys",0x800000,0xFFA0A0,"typekey.htm"], _
["Delaying","Delay a macro for a while",0x008000,0xA0FFA0,"typedelay.htm"],["Window control","Make a change to the program window",0x107040,0xB0FFE0,"typewindow.htm"], _
["Randomize","For the following steps randomize mouse clicks and movement",0x703000,0xFFC0A0,"typerandom.htm"],["Label","Add a jump label",0x506030,0xF0FFD0,"typelabel.htm"], _
["Jump to","Jump to label or other jump action",0x307000,0xD0FFA0,"typejump.htm"],["End block","End block for a loop of if",0xFF0000,0xFFA0A0,"typeend.htm"], _
["Repeat loop","Repeat a block for a number of times or check the counter (if step)",0x900080,0xFFA0C0,"typerepeat.htm"],["Time loop","Execute a block for a time period",0x804040,0xFFE0E0,"typetime.htm"], _
["Input","Input a mouse button click or a key press",0x407040,0xE0FFE0,"typeinput.htm"],["Pixel(s) change","Check if a pixel or area has been changed. Keep mouse pointer out of the way of area!",0x002090,0xA0C0FF,"typepixel.htm"], _
["Time alarm","Execute a block on based of a time/date",0x900030,0xFFA0D0,"typealarm.htm"],["Sound","Play and control playing sound",0x904000,0xFFE0A0,"typesound.htm"], _
["Counter","Set, add to or check a counter",0x409070,0xE0FFFF,"typecounter.htm"],["File or Folder control","Check file or folder on existence or changes",0x400070,0xE0A0FF,"typefile.htm"], _
["Pause or stop","Pause or Stop the macro",0x601040,0xFFB0E0,"typecontrol.htm"],["Program/Process control","Start a process/program and other process control",0x804030,0xFFE0D0,"typeprocess.htm"], _
["Screen or window capture","Capture the screen, window or an area",0x506000,0xF0FFA0,"typecapture.htm"],["Computer shutdown","Shutdown or restart your computer",0x106080,0xB0FFFF,"typeshutdown.htm"], _
["Comment line","Add a comment line for clarification",0x508040,0xF0FFE0,"typecomment.htm"]]
Global Enum $StepTypeMouse,$StepTypeKey,$StepTypeDelay,$StepTypeWindow,$StepTypeRandom,$StepTypeLabel,$StepTypeJump,$StepTypeEnd,$StepTypeRepeat,$StepTypeTime,$StepTypeInput, _
$StepTypePixel,$StepTypeAlarm,$StepTypeSound,$StepTypeCounter,$StepTypeFile,$StepTypeControl,$StepTypeProcess,$StepTypeCapture,$StepTypeShutdown,$StepTypeComment
Global $MacroMousePositions = ["Relative to program window position","Absolute screen coordinates","Relative to mouse position","Relative to screen center","Relative to window center"]
Global $StartStopKeys = [["Alt+Space","!{SPACE}"],["Ctrl+Space","^{SPACE}"],["Ctrl+Alt+Space","^!{SPACE}"],["Ctrl+Alt+S","^!s"],["Escape","{ESC}"]]
Global $PauseKeys = [["Alt+Backspace","!{BACKSPACE}"],["Ctrl+Backspace","^{BACKSPACE}"],["Ctrl+Alt+Backspace","^!{BACKSPACE}"],["Ctrl+Alt+P","^!p"]]
Global $MacroMouseButtons = [ _
["Left button click","01",$MOUSE_CLICK_PRIMARY],["Right button click","02",$MOUSE_CLICK_SECONDARY],["Middle button click","04",$MOUSE_CLICK_MIDDLE],["Move mouse","01",-1], _
["Left button down","01",$MOUSE_CLICK_PRIMARY],["Right button down","02",$MOUSE_CLICK_SECONDARY],["Middle button down","04",$MOUSE_CLICK_MIDDLE], _
["Left button up","01",$MOUSE_CLICK_PRIMARY],["Right button up","02",$MOUSE_CLICK_SECONDARY],["Middle button up","04",$MOUSE_CLICK_MIDDLE], _
["Push position","01",-1],["Pop position","01",-1]]
Global $MacroKeys = [ _
["characters","","",False],["Space","{SPACE}","20",False],["Enter","{ENTER}","0D",False],["Backspace","{BS}","08",False],["Delete","{DEL}","2E",False],["Up","{UP}","26",False],["Down","{DOWN}","28",False], _
["Left","{LEFT}","25",False],["Right","{RIGHT}","27",False],["Home","{HOME}","24",False],["End","{END}","23",False],["Escape","{ESC}","1B",False],["Insert","{INS}","2D",False],["Page Up","{PgUp}","21",False], _
["Page Down","{PgDn}","22",False],["Tab","{Tab}","09",False],["F1","{F1}","70",False],["F2","{F2}","71",False],["F3","{F3}","72",False],["F4","{F4}","73",False],["F5","{F5}","74",False],["F6","{F6}","75",False], _
["F7","{F7}","76",False],["F8","{F8}","77",False],["F9","{F9}","78",False],["F10","{F10}","79",False],["F11","{F11}","7A",False],["F12","{F12}","7B",False],["Pause","{PAUSE}","13",False], _
["Numeric Keypad 0","{NUMPAD0}","60",False],["Numeric Keypad 1","{NUMPAD1}","61",False],["Numeric Keypad 2","{NUMPAD2}","62",False],["Numeric Keypad 3","{NUMPAD3}","63",False],["Numeric Keypad 4","{NUMPAD4}","64",False], _
["Numeric Keypad 5","{NUMPAD5}","65",False],["Numeric Keypad 6","{NUMPAD6}","66",False],["Numeric Keypad 7","{NUMPAD7}","67",False],["Numeric Keypad 8","{NUMPAD8}","68",False],["Numeric Keypad 9","{NUMPAD9}","69",False], _
["Numeric Keypad *","{NUMPADMULT}","6A",False],["Numeric Keypad /","{NUMPADDIV}","6F",False],["Numeric Keypad +","{NUMPADADD}","6B",False],["Numeric Keypad -","{NUMPADMIN}","6D",False],["Numeric Keypad .","{NUMPADDOT}","6E",False], _
["Break","{BREAK}","03",False],["Caps Lock","{CAPSLOCK}","14",False],["Numeric Lock","{NUMLOCK}","90",False],["Scroll Lock","{SCROLLLOCK}","91",False],["Print Screen","{PRINTSCREEN}","2C",False],["Sleep","{SLEEP}","5F",False], _
["Volume Up","{VOLUME_UP}","AF",False],["Volume Down","{VOLUME_DOWN}","AE",False],["Volume Mute","{VOLUME_MUTE}","AD",False],["Play / Pause","{MEDIA_PLAY_PAUSE}","B3",False],["Stop","{MEDIA_STOP}","B2",False], _
["Next","{MEDIA_NEXT}","B0",False],["Previous","{MEDIA_PREV}","B1",False], _
["Browser Back","{BROWSER_BACK}","A6",False],["Browser Forward","{BROWSER_FORWARD}","A7",False],["Browser Refresh","{BROWSER_REFRESH}","A8",False],["Browser Stop","{BROWSER_STOP}","A9",False], _
["Browser Search","{BROWSER_SEARCH}","AA",False],["Browser Favourites","{BROWSER_FAVORITES}","AB",False],["Browser Home","{BROWSER_HOME}","AC",False], _
["Launch Email","{LAUNCH_MAIL}","B4",False],["Launch Media","{LAUNCH_MEDIA}","B5",False],["Launch Application 1","{LAUNCH_APP1}","B6",False],["Launch Application 2","{LAUNCH_APP2}","B7",False], _
["Left Shift","{LSHIFT}","A0",False],["Right Shift","{RSHIFT}","A1",False],["Left Control","{LCTRL}","A2",False],["Right Control","{RCTRL}","A3",False],["Alt","{ALT}","12",False],["Left Alt","{LALT}","",False], _
["Right Alt","{RALT}","",False],["Left Windows Key","{LWIN}","5B",False],["Right Windows Key","{RWIN}","5C",False],["Shift Down","{SHIFTDOWN}","10",False],["Shift Up","{SHIFTUP}","",False], _
["Control Down","{CTRLDOWN}","11",False],["Control Up","{CTRLUP}","",False],["Alt Down","{ALTDOWN}","12",False],["Alt Up","{ALTUP}","",False],["Left Windows Key Down","{LWINDOWN}","5B",False], _
["Left Windows Key Up","{LWINUP}","",False],["Right Windows Key Down","{RWINDOWN}","5C",False],["Right Windows Key Up","{RWINUP}","",False],["Windows Apps Key","{APPSKEY}","",False]]
Global $MacroKeyModes = [["None",""],["Shift","+"],["Control","^"],["Alt","!"],["Windows Key","#"],["Control Shift","^+"],["Control Alt","^!"],["Alt Shift","!+"],["Control Alt Shift","^!+"]]
Global $MacroTextMode = ["Enter characters","Random a to z, A to Z and 0 to 9","Random a to z and A to Z","Random a to z","Random a to z and 0 to 9","Random 0 to 9"]
Global $MacroDelayTypes = ["Delay of this step","Delay of every step"]
Global $MacroWindowActions = ["Maximize window","Restore window to former size","Minimize window","Move window to position","Resize window","Close window","If window maximized","If window not maximized","If window minimized", _
"If window not minimized","Show window","Hide window","Force closing of window","Switch to window in this process","If window exists in this process","If window not exists in this process","Switch to first found window"]
Global $MacroWindowPositions = ["Absolute screen coordinates","Relative to program window position"]
Global $MacroWindowSizes = ["Absolute size","Relative to previous size"]
Global $MacroJumpTypes = ["Jump to label","Restart macro","Exit macro","Exit block","Exit loop"]
Global $MacroRepeatTypes = ["Repeat loop","If counter is equal to","If counter is not equal to","If counter is greater or equal than","If counter is smaller or equal than","If counter is between","If counter is not between","If counter is odd","If counter is even"]
Global $MacroInputTypes = ["Click or press","Pop up message","Pop up confirmation"]
Global $MacroContinueKeys = [ _
["Left mouse button","01"],["Right mouse button","02"],["Middle mouse button","04"],["Space","20"],["Enter","0D"],["Backspace","08"],["Delete","2E"],["Up","26"],["Down","28"],["Left","25"],["Right","27"], _
["Home","24"],["End","23"],["Escape","18"],["Insert","2D"],["Page Up","21"],["Page Down","22"],["Tab","09"],["F1","70"],["F2","71"],["F3","72"],["F4","73"],["F5","74"],["F6","75"],["F7","76"],["F8","77"],["F9","78"],["F10","79"],["F11","7A"],["F12","7B"], _
["Shift","10"],["Control","11"],["Alt","12"],["Pause","13"],["Caps Lock","14"],["Break","03"],["Print Screen","2C"]]
Global $MacroPixelTypes = ["Record pixel area","If pixel area changed","Record pixel color","If pixel color changed"]
Global $MacroAlarmTypes = ["Specific time and/or date","On the minute","Hourly","Daily"]
Global $MacroAlarmSpecifics = ["After set time, run block once","After set time, run block always","Before set time, run block once","Before set time, run block always"]
Global $MacroSoundTypes = ["Beep","Play sound","Stop sounds","Set volume","Pause sounds","Resume sounds"]
Global $MacroCounterTypes = ["Set counter","Set counter during playing","Add to counter","If counter is equal to","If counter is not equal to","If counter is greater or equal than","If counter is smaller equal than","If counter is between","If counter is not between","If counter is odd","If counter is even"]
Global $MacroFileTypes = ["File exists","File not exist","Folder exists","Folder not exist","File to check change of","If file changed","If file not changed"]
Global $MacroControlTypes = ["Pause macro, continue yourself","Stop macro","Stop macro and close " & $ProgramName]
Global $MacroProcessTypes = ["Switch to selected process","Start program","If process exists","If process not exists","Process checking off","Process checking on","Set checking to new proces"]
Global $MacroProcessWindowTypes = ["Use current selected window","Set window by keyword","Don't remember window, only process"]
Global $MacroProcessWindowStates = ["Normal size","Maximized","Minimized","Hidden"]
Global $MacroCaptureActions = ["Capture screen on","Capture screen area on","Capture window on","Capture window area on","Capture off","Capture attributes"]
Global $MacroShutDownActions = ["Log off current user","Go into standby mode", "Hibernate, writing memory to disk","Restart computer","Shut down computer","Shut down computer with force","Power down, kill processes immediately"]
Global $MacroCommentTypes = ["Text","Line"]
Global $User32DLL,$HideStartupMessage,$HideDetectionMessage,$HideShowMessage,$MacroFile,$MacroFileVersion,$StepView,$ProgramTheme,$ShowInvisible,$ViewInvertColors,$ViewNumbers,$HidePlayMessage,$HideEarlierVersionWarning,$HideLoopWarning,$HideProcessCheckWarning, _
	   $DefaultStepDelay,$LeftIndent,$HidePlayToolTip,$ActivateTimeOut,$ActivateDelay,$HideRecordToolTip,$HideRecordMessage,$HideShowRecordWarning,$HideMoveTimeWarning,$HideStepDelayWarning
Global $ProcessList[1][5]
; $ProcessList[process][0]		Process name
; $ProcessList[process][1]		Window name
; $ProcessList[process][2]		Window handle
; $ProcessList[process][3]		Process id
; $ProcessList[process][4]		Invisible
Global $MacroStepsCopy[1],$Undo,$Undos = _StackCreate(),$MenuItemUndo,$MenuItemRedo,$RecentFiles = _ShiftCreate(),$RecentMenuItems[1],$RecentFilesMenu,$MenuItemRecent,$CurrentProcess,$CurrentWindow,$GUI,$Changed = False,$CloseByMacro = False
Global $ColorHighLight = _ColorBGRtoRGB(_WinAPI_GetSysColor($COLOR_HIGHLIGHT)),$GUIPosition[4]
Global $Close,$New,$Open,$Save,$SaveAs,$Settings,$Help,$Add,$Change,$Duplicate,$Delete,$Up,$Down,$Show,$Enable,$Disable
Global $WindowsVisible,$WindowsInvisible,$StepViewList,$StepViewReport,$StepViewGrid,$StepViewInverseColor,$StepViewNumbers
Global $StartStopKey,$PauseKey,$Stopped,$Paused,$StepDelayLabel,$StepDelay,$StepDelayAlways,$StepDelayAlwaysLabel,$Record,$ShowRecording,$ShowRecordingLabel,$MoveTimeLabel,$MoveTime, _
       $Play,$LoopCountLabel,$LoopCount,$ShowMain,$ShowMainLabel,$ShowBar,$ShowBarLabel,$TestMacro,$TestMacroLabel,$ProcessesLabel,$Processes,$Macros,$MacroStepsLabel
Global Const $MacroStepEnableDisableColumn = 10,$MacroStepMarkColumn = $MacroStepEnableDisableColumn+1,$MacroStepColorColumn = $MacroStepEnableDisableColumn+2, _
			$MacroStepWindowColumn = $MacroStepEnableDisableColumn+3,$MacroStepColumns = $MacroStepEnableDisableColumn+5
Global $MacroSteps[1][$MacroStepColumns]
; $MacroSteps[step][0]		Type $MacroStepType[]
; $MacroSteps[step][1]		Description, if empty program creates description
; $MacroSteps[step][2]		Subtype such as $MacroMouseButtons[]
; $MacroSteps[step][3]		Value type such width, x coordinate and other types
; $MacroSteps[step][4]		Value type such height, y coordinate and others types
; $MacroSteps[step][5]		If used, position type $MacroMousePositions[] and $MacroWindowPositions[] else other value types
; $MacroSteps[step][6]		If used, randomize x value else other value types
; $MacroSteps[step][7]		If used, randomize y value else other value types
; $MacroSteps[step][8]		Value types
; $MacroSteps[step][9]		Value types or process id
; $MacroSteps[step][10]		0 = enabled, 1 = disabled
; $MacroSteps[step][11]		0 = not bookmarked, 1 = bookmarked
; $MacroSteps[step][12]		temporary color storage
; $MacroSteps[step][13]		window handle belonging to step
; $MacroSteps[step][14]		process id belonging to step
Global $MacroStepsBlock[1],$MacroPreviousStep,$MacroStep,$MacroStepCount = 0,$MacroStepType,$MacroDescription
Global $MacroMouseButton,$MacroPositionX,$MacroPositionY,$MacroMousePosition,$MacroMouseMoveInstantaneous,$MacroMouseMoveInstantaneousLabel,$MacroPositionXRandom,$MacroPositionYRandom,$MacroDetectPosition,$MacroDetectClick,$MacroShowPosition
Global $MacroKey,$MacroKeyMode,$MacroText,$MacroWindowAction,$MacroWindowMoveInstantaneous,$MacroWindowMoveInstantaneousLabel,$MacroWindowResizeInstantaneous,$MacroWindowResizeInstantaneousLabel,$MacroWindowPosition,$MacroWindowSize,$MacroWidth,$MacroHeight,$MacroRandomMouse,$MacroRandomDelay
Global $MacroJumpType,$MacroJumpLabel,$MacroRepeatType,$MacroRepeatCount,$MacroRepeatValue1,$MacroRepeatValue2,$MacroTimeLoop,$MacroContinueKey,$MacroAlarmType,$MacroAlarmTime,$MacroAlarmDate,$MacroAlarmSpecific
Global $MacroMouseButtonLabel,$MacroMouseClicks,$MacroMouseClicksLabel,$MacroLabelX,$MacroLabelY,$MacroPositionXRandomLabel,$MacroPositionYRandomLabel,$MacroKeyLabel,$MacroKeyModeLabel,$MacroTextLabel,$MacroTextSendLabel,$MacroTextSend
Global $MacroDelayTypeLabel,$MacroDelayType,$MacroDelayLabel,$MacroDelayRandomLabel,$MacroDelay,$MacroDelayRandom,$MacroDelayStepLabel,$MacroDelayStep
Global $MacroWindowActionLabel,$MacroWidthLabel,$MacroHeightLabel,$MacroRandomMouseLabel,$MacroRandomDelayLabel,$MacroJumpTypeLabel,$MacroJumpLabelLabel,$MacroWindowWindowLabel,$MacroWindowWindow
Global $MacroRepeatTypeLabel,$MacroRepeatCountLabel,$MacroRepeatValue1aLabel,$MacroRepeatValue1bLabel,$MacroRepeatValue1cLabel,$MacroRepeatValue2Label,$MacroTimeLoopLabel
Global $MacroContinueKeyLabel,$MacroAlarmTypeLabel,$MacroAlarmTimeLabel,$MacroAlarmDateLabel,$MacroAlarmSpecificLabel,$MacroSoundTypeLabel,$MacroSoundFileLabel,$MacroSoundVolumeLabel
Global $MacroInputTypeLabel,$MacroInputType,$MacroPixelTypeLabel,$MacroPixelType,$MacroPixelWidthLabel,$MacroPixelWidth,$MacroPixelHeightLabel,$MacroPixelHeight,$MacroPixelId,$MacroPixelIdLabel1,$MacroPixelIdLabel2
Global $MacroSoundType,$MacroSoundSelect,$MacroSoundFile,$MacroSoundVolume,$MacroSoundVolumeUpDown,$MacroSoundFrequencyLabel,$MacroSoundFrequency,$MacroSoundDurationLabel,$MacroSoundDuration
Global $MacroCounterTypeLabel,$MacroCounterType,$MacroCounterNameLabel,$MacroCounterName,$MacroCounterNameList,$MacroCounterSet,$MacroCounterSetLabel,$MacroCounterRandom,$MacroCounterRandomLabel
Global $MacroCounterValue1,$MacroCounterValue1aLabel,$MacroCounterValue1bLabel,$MacroCounterValue1cLabel,$MacroCounterValue2Label,$MacroCounterValue2
Global $MacroFileTypeLabel,$MacroFileType,$MacroFileFileLabel,$MacroFileFile,$MacroFileFileSelect,$MacroFilePathLabel,$MacroFilePath,$MacroFilePathSelect
Global $MacroFileStampLabel,$MacroFileStamp,$MacroFileStampSelect,$MacroFileChangeLabel,$MacroFileChange,$MacroFileChangeSelect
Global $MacroControlTypeLabel,$MacroControlType
Global $MacroProcessTypeLabel,$MacroProcessType,$MacroProcessShowProcessCurrentLabel,$MacroProcessShowProcessCurrent,$MacroProcessShowProcessLabel,$MacroProcessShowProcess,$MacroProcessWindowTypeLabel,$MacroProcessWindowType
Global $MacroProcessProgramLabel,$MacroProcessProgram,$MacroProcessProgramSelect,$MacroProcessWindowLabel,$MacroProcessWindow,$MacroProcessWorkDirLabel,$MacroProcessWorkDir,$MacroProcessWorkDirSelect,$MacroProcessWait,$MacroProcessWaitLabel
Global $MacroProcessTest,$MacroProcessWindowStateLabel,$MacroProcessWindowState,$MacroProcessTimeoutLabel,$MacroProcessTimeout,$MacroProcessProcessLabel,$MacroProcessProcess,$MacroProcessProcessSelect
Global $MacroCaptureActionLabel,$MacroCaptureAction,$MacroCaptureXLabel,$MacroCaptureYLabel,$MacroCaptureWidthLabel,$MacroCaptureHeightLabel,$MacroCaptureScreenX,$MacroCaptureScreenY,$MacroCaptureScreenWidth,$MacroCaptureScreenHeight
Global $MacroCaptureWindowX,$MacroCaptureWindowY,$MacroCaptureWindowWidth,$MacroCaptureWindowHeight,$MacroCaptureShow,$MacroCaptureDetect
Global $MacroCapturePathLabel,$MacroCapturePath,$MacroCapturePathSelect,$MacroCapturePointer,$MacroCapturePointerLabel
Global $MacroShutDownActionLabel,$MacroShutDownAction,$MacroShutDownAsk,$MacroShutDownAskLabel,$MacroCommentTypeLabel,$MacroCommentType

FileChangeDir(@ScriptDir)						; make working directory the script directory
_DialogueTitle($ProgramName)					; all dialogues gets program name title bar
ReadSettings()									; read settings
ExtractFiles()									; extract all files from executable
StartupMessage()								; show warning on playing macro's
GUI()											; show program interface
CleanUp()										; clean up such as removing extracted files

Func ExtractFiles()								; extract all files from executable
	FileInstall("Readme.txt",@ScriptDir & "\Readme.txt",$FC_OVERWRITE)
	FileInstall("Help\Macronize.chm",@ScriptDir & "\Macronize.chm",$FC_OVERWRITE)
	FileInstall("Images\important.ico",@ScriptDir & "\important.ico",$FC_OVERWRITE)
	FileInstall("Images\info.ico",@ScriptDir & "\info.ico",$FC_OVERWRITE)
	FileInstall("Images\settings.png",@ScriptDir & "\settings.png",$FC_OVERWRITE)
	FileInstall("Images\questionmark.png",@ScriptDir & "\questionmark.png",$FC_OVERWRITE)
	FileInstall("Images\new.png",@ScriptDir & "\new.png",$FC_OVERWRITE)
	FileInstall("Images\open.png",@ScriptDir & "\open.png",$FC_OVERWRITE)
	FileInstall("Images\save.png",@ScriptDir & "\save.png",$FC_OVERWRITE)
	FileInstall("Images\saveas.png",@ScriptDir & "\saveas.png",$FC_OVERWRITE)
	FileInstall("Images\close.png",@ScriptDir & "\close.png",$FC_OVERWRITE)
	FileInstall("Images\arrowup.png",@ScriptDir & "\arrowup.png",$FC_OVERWRITE)
	FileInstall("Images\arrowdown.png",@ScriptDir & "\arrowdown.png",$FC_OVERWRITE)
	FileInstall("Images\minus.png",@ScriptDir & "\minus.png",$FC_OVERWRITE)
	FileInstall("Images\plus.png",@ScriptDir & "\plus.png",$FC_OVERWRITE)
	FileInstall("Images\cut.png",@ScriptDir & "\cut.png",$FC_OVERWRITE)
	FileInstall("Images\copy.png",@ScriptDir & "\copy.png",$FC_OVERWRITE)
	FileInstall("Images\edit.png",@ScriptDir & "\edit.png",$FC_OVERWRITE)
	FileInstall("Images\duplicate.png",@ScriptDir & "\duplicate.png",$FC_OVERWRITE)
	FileInstall("Images\detect.png",@ScriptDir & "\detect.png",$FC_OVERWRITE)
	FileInstall("Images\pointer.png",@ScriptDir & "\pointer.png",$FC_OVERWRITE)
	FileInstall("Images\enable.png",@ScriptDir & "\enable.png",$FC_OVERWRITE)
	FileInstall("Images\disable.png",@ScriptDir & "\disable.png",$FC_OVERWRITE)
	FileInstall("Images\play.png",@ScriptDir & "\play.png",$FC_OVERWRITE)
	FileInstall("Images\record.png",@ScriptDir & "\record.png",$FC_OVERWRITE)
	FileInstall("Images\visible.png",@ScriptDir & "\visible.png",$FC_OVERWRITE)
	FileInstall("Images\invisible.png",@ScriptDir & "\invisible.png",$FC_OVERWRITE)
	FileInstall("Images\viewlist.png",@ScriptDir & "\viewlist.png",$FC_OVERWRITE)
	FileInstall("Images\viewreport.png",@ScriptDir & "\viewreport.png",$FC_OVERWRITE)
	FileInstall("Images\viewgrid.png",@ScriptDir & "\viewgrid.png",$FC_OVERWRITE)
	FileInstall("Images\invertcolor.png",@ScriptDir & "\invertcolor.png",$FC_OVERWRITE)
	FileInstall("Images\numbers.png",@ScriptDir & "\numbers.png",$FC_OVERWRITE)
	FileInstall("Images\disabled.bmp",@ScriptDir & "\disabled.bmp",$FC_OVERWRITE)
	FileInstall("Images\mark.bmp",@ScriptDir & "\mark.bmp",$FC_OVERWRITE)
	FileInstall("Images\disabledmark.bmp",@ScriptDir & "\disabledmark.bmp",$FC_OVERWRITE)
	FileInstall("Images\error.bmp",@ScriptDir & "\error.bmp",$FC_OVERWRITE)
	FileInstall("Images\disablederror.bmp",@ScriptDir & "\disablederror.bmp",$FC_OVERWRITE)
	FileInstall("Sounds\ping.wav",@ScriptDir & "\ping.wav",$FC_OVERWRITE)
	FileInstall("Sounds\metal.wav",@ScriptDir & "\metal.wav",$FC_OVERWRITE)
	FileInstall("Sounds\bounce.wav",@ScriptDir & "\bounce.wav",$FC_OVERWRITE)
EndFunc

Func CleanUp()									; clean up after exit
	FileDelete(@ScriptDir & "\important.ico")
	FileDelete(@ScriptDir & "\info.ico")
	FileDelete(@ScriptDir & "\settings.png")
	FileDelete(@ScriptDir & "\questionmark.png")
	FileDelete(@ScriptDir & "\new.png")
	FileDelete(@ScriptDir & "\open.png")
	FileDelete(@ScriptDir & "\save.png")
	FileDelete(@ScriptDir & "\saveas.png")
	FileDelete(@ScriptDir & "\close.png")
	FileDelete(@ScriptDir & "\arrowup.png")
	FileDelete(@ScriptDir & "\arrowdown.png")
	FileDelete(@ScriptDir & "\minus.png")
	FileDelete(@ScriptDir & "\plus.png")
	FileDelete(@ScriptDir & "\cut.png")
	FileDelete(@ScriptDir & "\copy.png")
	FileDelete(@ScriptDir & "\edit.png")
	FileDelete(@ScriptDir & "\duplicate.png")
	FileDelete(@ScriptDir & "\detect.png")
	FileDelete(@ScriptDir & "\pointer.png")
	FileDelete(@ScriptDir & "\enable.png")
	FileDelete(@ScriptDir & "\disable.png")
	FileDelete(@ScriptDir & "\play.png")
	FileDelete(@ScriptDir & "\record.png")
	FileDelete(@ScriptDir & "\visible.png")
	FileDelete(@ScriptDir & "\invisible.png")
	FileDelete(@ScriptDir & "\viewlist.png")
	FileDelete(@ScriptDir & "\viewreport.png")
	FileDelete(@ScriptDir & "\viewgrid.png")
	FileDelete(@ScriptDir & "\invertcolor.png")
	FileDelete(@ScriptDir & "\numbers.png")
	FileDelete(@ScriptDir & "\disabled.bmp")
	FileDelete(@ScriptDir & "\mark.bmp")
	FileDelete(@ScriptDir & "\disabledmark.bmp")
	FileDelete(@ScriptDir & "\error.bmp")
	FileDelete(@ScriptDir & "\disablederror.bmp")
	FileDelete(@ScriptDir & "\ping.wav")
	FileDelete(@ScriptDir & "\metal.wav")
	FileDelete(@ScriptDir & "\bounce.wav")
	If _StackSize($Undos) > 0 Then FileDelete($UndoTempFile & "*")
	FileDelete(FileFailSave())
EndFunc

Func StartUpMessage()							; show warning on playing macro's
	If Not $HideStartupMessage Then
		_Message("Create your macro with sense. A macro can really mesh up your program/system." & @CRLF & @CRLF & "And your computer may do things that interfere with the running of your macro.",$HideStartupMessage,$ProgramName & " - Notice","Don't show this message at startup","I acknowledge",@ScriptDir & "\important.ico",0,0,0,$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
		WriteSettings()
	EndIf
EndFunc

Func ReadSettings()								; read main program settings
	$GUIPosition[0] = Number(IniRead($SettingsFile,"General","Main Window X",-1))
	$GUIPosition[1] = Number(IniRead($SettingsFile,"General","Main Window Y",-1))
	$GUIPosition[2] = Number(IniRead($SettingsFile,"General","Main Window Width",0))
	$GUIPosition[3] = Number(IniRead($SettingsFile,"General","Main Window Height",0))
	$ProgramTheme = Number(IniRead($SettingsFile,"General","Theme",-1))
	$ShowInvisible = Number(IniRead($SettingsFile,"General","Show Invisible Windows","0")) <> 0
	$ViewNumbers = Number(IniRead($SettingsFile,"General","Step List View Line Numbers","0")) <> 0
	$ViewInvertColors = Number(IniRead($SettingsFile,"General","Step List View Inverse Colors","0")) <> 0
	$StepView = Number(IniRead($SettingsFile,"General","Step List View","1"))
	$StartStopKey = Number(IniRead($SettingsFile,"General","Start Stop Key","0"))
	$PauseKey = Number(IniRead($SettingsFile,"General","Pause Key","0"))
	$LeftIndent = Number(IniRead($SettingsFile,"General","Left Indent","4"))
	$HideStartupMessage = Number(IniRead($SettingsFile,"General","Hide Startup Message","0")) <> 0
	$HideDetectionMessage = Number(IniRead($SettingsFile,"Macro Edit","Hide Detection Message","0")) <> 0
	$HideShowMessage = Number(IniRead($SettingsFile,"Macro Edit","Hide Show Message","0")) <> 0
	$HideProcessCheckWarning = Number(IniRead($SettingsFile,"Macro Edit","Hide Process Check Warning","0")) <> 0
	$HidePlayMessage = Number(IniRead($SettingsFile,"Macro Play","Hide Play Message","0")) <> 0
	$HideEarlierVersionWarning = Number(IniRead($SettingsFile,"Macro Play","Hide Version Warning","0")) <> 0
	$HidePlayToolTip = Number(IniRead($SettingsFile,"Macro Play","Hide Play Tooltip","0")) <> 0
	$HideLoopWarning = Number(IniRead($SettingsFile,"Macro Play","Hide Loop Warning","0")) <> 0
	$HideStepDelayWarning = Number(IniRead($SettingsFile,"Macro Play","Hide Step Delay Warning","0")) <> 0
	$DefaultStepDelay = Number(IniRead($SettingsFile,"Macro Play","Default Delay","100"))
	$ActivateTimeOut = Number(IniRead($SettingsFile,"Macro Play","Window Activate Timeout","5000"))
	$ActivateDelay = Number(IniRead($SettingsFile,"Macro Play","Window Activate Delay","500"))
	$HideRecordMessage = Number(IniRead($SettingsFile,"Macro Record","Hide Record Message","0")) <> 0
	$HideRecordToolTip = Number(IniRead($SettingsFile,"Macro Record","Hide Record Tooltip","0")) <> 0
	$HideShowRecordWarning = Number(IniRead($SettingsFile,"Macro Record","Hide Show Recording Warning","0")) <> 0
	$HideMoveTimeWarning = Number(IniRead($SettingsFile,"Macro Record","Hide Move Time Warning","0")) <> 0
EndFunc

Func WriteSettings()							; write main program settings
	IniWrite($SettingsFile,"General","Theme",$ProgramTheme)
	IniWrite($SettingsFile,"General","Show Invisible Windows",$ShowInvisible ? "1" : "0")
	IniWrite($SettingsFile,"General","Show Invisible Windows",$ShowInvisible ? "1" : "0")
	IniWrite($SettingsFile,"General","Step List View",$StepView)
	IniWrite($SettingsFile,"General","Step List View Line Numbers",$ViewNumbers ? "1" : "0")
	IniWrite($SettingsFile,"General","Step List View Inverse Colors",$ViewInvertColors ? "1" : "0")
	IniWrite($SettingsFile,"General","Hide Startup Message", $HideStartupMessage ? "1" : "0")
	IniWrite($SettingsFile,"General","Start Stop Key",$StartStopKey)
	IniWrite($SettingsFile,"General","Pause Key",$PauseKey)
	IniWrite($SettingsFile,"General","Left Indent",$LeftIndent)
	IniWrite($SettingsFile,"Macro Edit","Hide Detection Message",$HideDetectionMessage ? "1" : "0")
	IniWrite($SettingsFile,"Macro Edit","Hide Show Message",$HideShowMessage ? "1" : "0")
	IniWrite($SettingsFile,"Macro Edit","Hide Process Check Warning",$HideProcessCheckWarning ? "1" : "0")
	IniWrite($SettingsFile,"Macro Play","Hide Play Message",$HidePlayMessage ? "1" : "0")
	IniWrite($SettingsFile,"Macro Play","Hide Version Warning",$HideEarlierVersionWarning ? "1" : "0")
	IniWrite($SettingsFile,"Macro Play","Hide Play Tooltip",$HidePlayToolTip ? "1" : "0")
	IniWrite($SettingsFile,"Macro Play","Hide Loop Warning",$HideLoopWarning ? "1" : "0")
	IniWrite($SettingsFile,"Macro Play","Hide Step Delay Warning",$HideStepDelayWarning ? "1" : "0")
	IniWrite($SettingsFile,"Macro Play","Default Delay",$DefaultStepDelay)
	IniWrite($SettingsFile,"Macro Play","Window Activate Timeout",$ActivateTimeOut)
	IniWrite($SettingsFile,"Macro Play","Window Activate Delay",$ActivateDelay)
	IniWrite($SettingsFile,"Macro Record","Hide Record Message",$HideRecordMessage ? "1" : "0")
	IniWrite($SettingsFile,"Macro Record","Hide Record Tooltip",$HideRecordToolTip ? "1" : "0")
	IniWrite($SettingsFile,"Macro Record","Hide Show Recording Warning",$HideShowRecordWarning ? "1" : "0")
	IniWrite($SettingsFile,"Macro Record","Hide Move Time Warning",$HideMoveTimeWarning ? "1" : "0")
EndFunc

Func ReadFileName()								; read current macro file
	$MacroFile = IniRead($SettingsFile,"General","Macro File","")
EndFunc

Func WriteFileName()							; write current macro file
	IniWrite($SettingsFile,"General","Macro File",$MacroFile)
EndFunc

Func WriteMainWindowPosition()					; read main window size and position
	IniWrite($SettingsFile,"General","Main Window X",$GUIPosition[0])
	IniWrite($SettingsFile,"General","Main Window Y",$GUIPosition[1])
	IniWrite($SettingsFile,"General","Main Window Width",$GUIPosition[2])
	IniWrite($SettingsFile,"General","Main Window Height",$GUIPosition[3])
EndFunc

Func ReadRecentFiles()							; read recent macro files
	Local $RecentFile
	_ShiftClear($RecentFiles)
	For $File = $MaxRecentFiles To 1 Step -1
		$RecentFile = IniRead($SettingsFile,"Recent Files","Recent File " & $File,"")
		If StringLen($RecentFile) > 0 Then _ShiftPush($RecentFiles,$RecentFile)
	Next
	If _ShiftIsEmpty($RecentFiles) Then
		$RecentFilesMenu = False
	Else
		If $RecentFilesMenu Then
			For $File = 0 To UBound($RecentMenuItems)-1
				GUICtrlDelete($RecentMenuItems[$File])
			Next
		EndIf
		$RecentFilesMenu = True
		ReDim $RecentMenuItems[_ShiftSize($RecentFiles)]
		For $File = 0 To _ShiftSize($RecentFiles)-1
			$RecentMenuItems[$File] = GUICtrlCreateMenuItem(_ShiftGet($RecentFiles,$File),$MenuItemRecent)
		Next
	EndIf
EndFunc

Func WriteRecentFiles()							; write recent macro files
	IniDelete($SettingsFile,"Recent Files")
	If Not _ShiftIsEmpty($RecentFiles) Then
		For $File = 1 To _ShiftSize($RecentFiles)
			IniWrite($SettingsFile,"Recent Files","Recent File " & $File,_ShiftGet($RecentFiles,$File-1))
		Next
	EndIf
EndFunc

Func AddRecentFile($File)						; add file to recent macro files list
	If Not $RecentFilesMenu Or $File <> _ShiftGet($RecentFiles,0) Then
		_ShiftPush($RecentFiles,$File)
		If _ShiftSize($RecentFiles) > $MaxRecentFiles Then _ShiftPop($RecentFiles)
		WriteRecentFiles()
		ReadRecentFiles()
	EndIf
EndFunc

Func Label($Label,$X,$Y,$Width = -1,$Height = -1,$ToolTip = "",$Dock = -1)
	Local $Control = GUICtrlCreateLabel($Label,$X,$Y,$Width,$Height)
	If $ProgramTheme > -1 Then GUICtrlSetColor(-1,$Themes[$ProgramTheme][1])
	If $ToolTip <> "" Then GUICtrlSetTip(-1,$ToolTip)
	Switch $Dock
		Case 0
			GUICtrlSetResizing(-1,$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKSIZE)
		Case 1
			GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKSIZE)
	EndSwitch
	Return $Control
EndFunc

Func Input($X,$Y,$Width,$Height,$ToolTip = "",$Places = -2,$Dock = -1)
	Local $Control = GUICtrlCreateInput("",$X,$Y,$Width,$Height)
	If $ProgramTheme > -1 Then
		GUICtrlSetBkColor(-1,$Themes[$ProgramTheme][2])
		GUICtrlSetColor(-1,$Themes[$ProgramTheme][3])
	EndIf
	If $ToolTip <> "" Then GUICtrlSetTip(-1,$ToolTip)
	If $Dock = 0 Then GUICtrlSetResizing(-1,$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKSIZE)
	If $Places = -1 Then
		_Inputmask_add($Control,$iIM_INTEGER)
	ElseIf $Places = 0 Then
		_Inputmask_add($Control,$iIM_POSITIVE_INTEGER)
	elseIf $Places > 0 Then
		_Inputmask_add($Control,$iIM_POSITIVE_INTEGER,0,"",$Places)
	EndIf
	If $Places > -2 Then GUICtrlSetData(-1,0)
	Return $Control
EndFunc

Func Checkbox($Text,$X,$Y,$Width,$Height,$ToolTip, ByRef $Checkbox, ByRef $Label,$Dock = -1)
	$Checkbox = GUICtrlCreateCheckbox("",$X,$Y,16,16)
	If $ToolTip <> "" Then GUICtrlSetTip(-1,$ToolTip)
	If $Dock = 0 Then GUICtrlSetResizing(-1,$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKSIZE)
	If $Text <> "" Then
		$Label = Label($Text,$X+17,$Y,-1,-1,$ToolTip,$Dock)
		If $ToolTip <> "" Then GUICtrlSetTip(-1,$ToolTip)
	EndIf
EndFunc

Func Button($Image,$Text,$ToolTip,$X,$Y,$Width,$Height,$Dock = -1)
	Local $Control = _GraphicButton($Image,$Text,$ToolTip,$X,$Y,$Width,$Height)
	Switch $Dock
		Case 0
			GUICtrlSetResizing(-1,$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKSIZE)
		Case 1
			GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKSIZE)
		Case 2
			GUICtrlSetResizing(-1,$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKSIZE)
	EndSwitch
	Return $Control
EndFunc

Func ShowButtonsToolTip()
	Local $ProcessName

	If $MacroStep = -1 Then
		$ProcessName = $ProcessList[SelectedProcessIndex()][0]
	Else
		$ProcessName = ProcessNameById($MacroSteps[$MacroStep][$MacroStepWindowColumn+1])
	EndIf
	GUICtrlSetTip($Show,"Show position in selected process " & $ProcessName)
	GUICtrlSetTip($MacroShowPosition,"Show position in selected process " & $ProcessName)
	GUICtrlSetTip($MacroCaptureShow,"Show top left corner of area of selected process " & $ProcessName)
	GUICtrlSetTip($MacroDetectPosition,"Detect a mouse click at a position relative to program window of selected process " & $ProcessName)
EndFunc

Func GUI()										; build program interfaces and show main interface
	Const $Margin = 10,$ProcessLabelWidth = 120,$ControlHeight = 25,$LabelWidth = 150,$ButtonWidth = 105,$ButtonHeight = 23,$SmallButtonWidth = 35,$ComboWidth = 250
	Const $ControlsWidth = 3*$Margin+$ButtonWidth,$ControlsHeight = _WindowMenuHeight()+3*$Margin+Floor(1.5*$ControlHeight)+$ButtonHeight
	Const $GUIWidth = 540+$ControlsWidth,$GUIHeight = 520+$ControlsHeight
	Const $ViewButtonWidth = 24,$ViewButtonHeight = 16,$ViewButtonMargin = 5,$GUIStepWidth = 600,$GUIStepHeight = 270,$InputWidth = 60,$InputHeight = 20
	Const $GUICopyWidth = 300,$GUICopyHeight = 100,$GUIFindWidth = 2*$Margin+$LabelWidth+$ComboWidth,$GUIFindHeight = 100,$GUIGoToWidth = 300,$GUIGoToHeight = 80
	Const $GUIEnableDisableWidth = 380,$GUIEnableDisableHeight = 100
	Const $GUISettingsWidth = 400,$GUISettingsHeight = 570,$SettingsControlHeight = 25,$SettingsLabelWidth = 160
	Local $ListWidth,$ListHeight,$StepHelp
	Local $MenuFile,$MenuEdit,$MenuSearch,$MenuOptions,$MenuHelp,$MenuItemNew,$MenuItemOpen,$MenuItemSave,$MenuItemSaveAs,$MenuItemExit
	Local $MenuItemCut,$MenuItemCutSteps,$MenuItemCopy,$MenuItemCopySteps,$MenuItemPaste,$MenuItemPasteAfter,$MenuItemEnableDisableSteps,$MenuItemFind,$MenuItemFindNext,$MenuItemFindPrevious,$MenuItemFindGoTo
	Local $MenuItemMark,$MenuItemMarkNext,$MenuItemMarkPrevious,$MenuItemDisabledNext,$MenuItemDisabledPrevious
	Local $MenuItemSettings,$MenuItemHelp,$MenuItemSite,$MenuItemAbout
	Local $ProcessRefresh = False,$aMsg,$PositionY,$ViewX,$ListImage,$Tooltip,$MouseX,$MouseY
	Local $FormerPosition[4],$MacroStepSelected,$AddingStep,$MacroSave,$MacroStepHelp,$MacroCancel,$GUIStep,$GUIShowArea,$GUIShowAreaShow = 0,$GUIShowAreaTimer
	Local $GUICopy,$CopyFromStep,$CopyToStep,$CopyStep,$CopyCut,$CopyCopy,$CopyClose,$GUIFind,$FindType,$FindKeywordLabel,$FindKeyword,$FindStepTypeLabel,$FindStepType,$FindFind,$FindClose,$FindStep
	Local $FindTypes = ["Keyword in description","Step type"],$GUIGoTo,$FindGoToLabel,$FindGoTo,$FindGoToJump,$FindGoToClose
	Local $GUIEnableDisable,$EnableDisableFromStep,$EnableDisableToStep,$EnableDisableEnable,$EnableDisableDisable,$EnableDisableClose
	Local $GUISettings,$SettingsSave,$SettingsReset,$SettingsCancel,$SettingsStartStopKey,$SettingsPauseKey,$SettingsShowStartupMessage,$SettingsShowStartupMessageLabel, _
		  $SettingsDefaultStepDelay,$SettingsShowLoopWarning,$SettingsShowLoopWarningLabel,$SettingsLeftIndent, _
		  $SettingsShowDetectionMessage,$SettingsShowDetectionMessageLabel,$SettingsShowShowMessage,$SettingsShowShowMessageLabel, _
		  $SettingsShowPlayMessage,$SettingsShowPlayMessageLabel,$SettingsShowVersionWarning,$SettingsShowVersionWarningLabel, _
	      $SettingsShowPlayToolTip,$SettingsShowPlayToolTipLabel,$SettingsActivateTimeOut,$SettingsActivateDelay,$SettingsShowProcessCheckWarning,$SettingsShowProcessCheckWarningLabel, _
		  $SettingsShowRecordMessage,$SettingsShowRecordMessageLabel,$SettingsShowShowRecordWarning,$SettingsShowShowRecordWarningLabel, _
		  $SettingsShowMoveTimeWarning,$SettingsShowMoveTimeWarningLabel,$SettingsShowStepDelayWarning,$SettingsShowStepDelayWarningLabel, _
		  $SettingsShowRecordToolTip,$SettingsShowRecordToolTipLabel,$SettingsThemes

	$User32DLL = DllOpen("user32.dll")

	_Inputmask_init(50)				; Maximum number of InputMasks for input controls, also includes controls of settings window
	_Inputmask_MsgRegister()		; Register for _InputMask

	_GDIPlus_StartUp()				; initialize graphics system
	_GraphicButtons()				; initialize use of graphical buttons

	; build settings window
	$GUISettings = GUICreate("Settings of " & $ProgramName,$GUISettingsWidth,$GUISettingsHeight,-1,-1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU)
	If $ProgramTheme > -1 Then GUISetBkColor($Themes[$ProgramTheme][0])
	$PositionY = $Margin
	Label("Stop play/record macro key",$Margin,$PositionY+3,$SettingsLabelWidth,$InputHeight)
	$SettingsStartStopKey = GUICtrlCreateCombo("",$Margin+$SettingsLabelWidth+$Margin,$PositionY,$GUISettingsWidth-3*$Margin-$SettingsLabelWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Key to stop the playing or recording of the macro")
	_GUICtrlComboBox_AddArray($SettingsStartStopKey,$StartStopKeys)
	$PositionY += $SettingsControlHeight
	Label("Pause macro key",$Margin,$PositionY+3,$SettingsLabelWidth,$InputHeight)
	$SettingsPauseKey = GUICtrlCreateCombo("",$Margin+$SettingsLabelWidth+$Margin,$PositionY,$GUISettingsWidth-3*$Margin-$SettingsLabelWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Key to pause the macro during playing")
	_GUICtrlComboBox_AddArray($SettingsPauseKey,$PauseKeys)
	$PositionY += $SettingsControlHeight
	Checkbox("Show startup message",$Margin,$PositionY,-1,-1,"Show introduction message at startup",$SettingsShowStartupMessage,$SettingsShowStartupMessageLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show detection information",$Margin,$PositionY,-1,-1,"Show select process information at mouse click detection in a popup box when pressing the Detect button",$SettingsShowDetectionMessage,$SettingsShowDetectionMessageLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show show information",$Margin,$PositionY,-1,-1,"Show select process information at coordinates showing in a popup box when pressing the Show button",$SettingsShowShowMessage,$SettingsShowShowMessageLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show Play macro information",$Margin,$PositionY,-1,-1,"Show Play macro information in a popup box when pressing the Play button",$SettingsShowPlayMessage,$SettingsShowPlayMessageLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show earlier version warning",$Margin,$PositionY,-1,-1,"Show earlier version warning when pressing the Play button",$SettingsShowVersionWarning,$SettingsShowVersionWarningLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show Stop macro key tooltip",$Margin,$PositionY,-1,-1,"Show Stop macro key tooltip near mouse pointer during the playing of a macro",$SettingsShowPlayToolTip,$SettingsShowPlayToolTipLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show loop warning",$Margin,$PositionY,-1,-1,"Show warning that looping a macro is dangerous",$SettingsShowLoopWarning,$SettingsShowLoopWarningLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show small step delay warning",$Margin,$PositionY,-1,-1,"Show warning of a small step delay being dangerous",$SettingsShowStepDelayWarning,$SettingsShowStepDelayWarningLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show process checking warning",$Margin,$PositionY,-1,-1,"Show warning at setting the process checking off",$SettingsShowProcessCheckWarning,$SettingsShowProcessCheckWarningLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show Record macro information",$Margin,$PositionY,-1,-1,"Show record message on key repetition",$SettingsShowRecordMessage,$SettingsShowRecordMessageLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show Stop recording key tooltip",$Margin,$PositionY,-1,-1,"Show Stop recording key tooltip near mouse pointer during the recording of macro steps",$SettingsShowRecordToolTip,$SettingsShowRecordToolTipLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show show recording warning",$Margin,$PositionY,-1,-1,"Show the message on showing the recording by updating the macro step list",$SettingsShowShowRecordWarning,$SettingsShowShowRecordWarningLabel)
	$PositionY += $SettingsControlHeight
	Checkbox("Show mouse movement recording warning",$Margin,$PositionY,-1,-1,"Show the warning on a small interval time of recording the mouse movement",$SettingsShowMoveTimeWarning,$SettingsShowMoveTimeWarningLabel)
	$PositionY += $SettingsControlHeight
	Label("Block indent spaces",$Margin,$PositionY+3,$SettingsLabelWidth,$InputHeight)
	$SettingsLeftIndent = Input($Margin+$SettingsLabelWidth+$Margin,$PositionY,$InputWidth,$InputHeight,"Number of spaces to left indent blocks in macro step list",2)
	$PositionY += $SettingsControlHeight
	Label("Window activation timeout",$Margin,$PositionY+3,$SettingsLabelWidth,$InputHeight)
	$SettingsActivateTimeOut = Input($Margin+$SettingsLabelWidth+$Margin,$PositionY,$InputWidth,$InputHeight,"Timeout in milliseconds when trying to activate a window during macro play",6)
	$PositionY += $SettingsControlHeight
	Label("Window activation delay",$Margin,$PositionY+3,$SettingsLabelWidth,$InputHeight)
	$SettingsActivateDelay = Input($Margin+$SettingsLabelWidth+$Margin,$PositionY,$InputWidth,$InputHeight,"Delay in milliseconds after a window is activated in order to actually being able to process it",6)
	$PositionY += $SettingsControlHeight
	Label("Default step delay ms",$Margin,$PositionY+3,$SettingsLabelWidth,$InputHeight)
	$SettingsDefaultStepDelay = Input($Margin+$SettingsLabelWidth+$Margin,$PositionY,$InputWidth,$InputHeight,"Default delay in milliseconds after executing a time critical macro step",6)
	$PositionY += $SettingsControlHeight+10
	Label("Color scheme",$Margin,$PositionY+3,$SettingsLabelWidth,$InputHeight)
	$SettingsThemes = GUICtrlCreateCombo("",$Margin+$SettingsLabelWidth+$Margin,$PositionY,$GUISettingsWidth-3*$Margin-$SettingsLabelWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Color scheme. After selecting restart " & $ProgramName & " to fully apply the color scheme")
	_GUICtrlComboBox_AddArray($SettingsThemes,$ThemesTitles)

	$SettingsSave = GUICtrlCreateButton("Save",$Margin,$GUISettingsHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	GUICtrlSetTip(-1,"Save settings")
	$SettingsReset = GUICtrlCreateButton("Reset",($GUISettingsWidth-$ButtonWidth)/2,$GUISettingsHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	GUICtrlSetTip(-1,"Reset settings")
	$SettingsCancel = GUICtrlCreateButton("Cancel",$GUISettingsWidth-$Margin-$ButtonWidth,$GUISettingsHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	GUICtrlSetTip(-1,"Cancel changed settings")

	; build macro add/edit step window (all step types)
	$GUIStep = GUICreate("",$GUIStepWidth,$GUIStepHeight,-1,-1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU)
	If $ProgramTheme > -1 Then GUISetBkColor($Themes[$ProgramTheme][0])
	$PositionY = $Margin
	Label("Step type",$Margin,$PositionY)
	$MacroStepType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select step type")
	_GUICtrlComboBox_AddArray($MacroStepType,$MacroStepTypes)
	$StepHelp = _GraphicButton(@ScriptDir & "\questionmark.png","","Get help on the selected step type",$Margin+$LabelWidth+$ComboWidth+$Margin,$PositionY-5,$SmallButtonWidth,$ButtonHeight)
	$PositionY += $ControlHeight
	Label("Description",$Margin,$PositionY)
	$MacroDescription = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$Margin-$LabelWidth-$Margin,$InputHeight,"Leave empty for default description. Use $1, $2, etc. to insert values in step description")
	; mouse
	$PositionY = $Margin+2*$ControlHeight
	$MacroMouseButtonLabel = Label("Mouse action",$Margin,$PositionY)
	$MacroMouseButton = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a mouse action")
	_GUICtrlComboBox_AddArray($MacroMouseButton,$MacroMouseButtons)
	$MacroDetectPosition = _GraphicButton(@ScriptDir & "\detect.png","Detect","",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-6,$ButtonWidth,$ButtonHeight)
	$PositionY += $ControlHeight
	$MacroLabelX = Label("X position",$Margin,$PositionY)
	$MacroPositionX = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"X position in pixels",-1)
	$PositionY += $ControlHeight
	$MacroLabelY = Label("Y position",$Margin,$PositionY)
	$MacroPositionY = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Y position in pixels",-1)
	$MacroMousePosition = GUICtrlCreateCombo("",$Margin+$LabelWidth+$Margin+$InputWidth,$PositionY-4-$ControlHeight/2,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a position type")
	_GUICtrlComboBox_AddArray($MacroMousePosition,$MacroMousePositions)
	$MacroShowPosition = _GraphicButton(@ScriptDir & "\pointer.png","Show","",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-6-$ControlHeight/2,$ButtonWidth,$ButtonHeight)
	$PositionY += $ControlHeight
	$MacroPositionXRandomLabel = Label("Randomize X position",$Margin,$PositionY)
	$MacroPositionXRandom = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Randomize X position in pixels for instance to mimic human behaviour",0)
	Checkbox("Move instantaneously",$Margin+$LabelWidth+$Margin+$InputWidth,$PositionY-3,-1,-1,"Move mouse to coordinates without the macro delay",$MacroMouseMoveInstantaneous,$MacroMouseMoveInstantaneousLabel)
	$PositionY += $ControlHeight
	$MacroPositionYRandomLabel = Label("Randomize Y position",$Margin,$PositionY)
	$MacroPositionYRandom = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Randomize Y position in pixels for instance to mimic human behaviour",0)
	$PositionY += $ControlHeight
	$MacroMouseClicksLabel = Label("Number of clicks",$Margin,$PositionY)
	$MacroMouseClicks = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Number of mouse clicks to perform",0)
	; key
	$PositionY = $Margin+2*$ControlHeight
	$MacroKeyModeLabel = Label("Key mode",$Margin,$PositionY)
	$MacroKeyMode = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a key mode")
	_GUICtrlComboBox_AddArray($MacroKeyMode,$MacroKeyModes)
	$PositionY += $ControlHeight
	$MacroKeyLabel = Label("Special key",$Margin,$PositionY)
	$MacroKey = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a special key")
	_GUICtrlComboBox_AddArray($MacroKey,$MacroKeys)
	$PositionY += $ControlHeight
	$MacroTextSendLabel = Label("Sending mode",$Margin,$PositionY)
	$MacroTextSend = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	_GUICtrlComboBox_AddArray($MacroTextSend,$MacroTextMode)
	$PositionY += $ControlHeight
	$MacroTextLabel = Label("Character or text",$Margin,$PositionY)
	$MacroText = Input($Margin+$LabelWidth,$PositionY-3,$ComboWidth,"Character or text")
	;delay
	$PositionY = $Margin+2*$ControlHeight
	$MacroDelayTypeLabel = Label("Delay action",$Margin,$PositionY)
	$MacroDelayType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a delay action, this step or every step")
	_GUICtrlComboBox_AddArray($MacroDelayType,$MacroDelayTypes)
	$PositionY += $ControlHeight
	$MacroDelayLabel = Label("Delay macro for",$Margin,$PositionY)
	$MacroDelay = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Delay of this step in milliseconds",0)
	$MacroDelayStepLabel = Label("Delay all steps",$Margin,$PositionY)
	$MacroDelayStep = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Delay for every step in milliseconds",0)
	$PositionY += $ControlHeight
	$MacroDelayRandomLabel = Label("Randomize delay",$Margin,$PositionY)
	$MacroDelayRandom = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Randomize delay in milliseconds to mimic human behaviour",0)
	; window
	$PositionY = $Margin+2*$ControlHeight
	$MacroWindowActionLabel = Label("Window action",$Margin,$PositionY)
	$MacroWindowAction = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select an action to perform on window")
	_GUICtrlComboBox_AddArray($MacroWindowAction,$MacroWindowActions)
	$PositionY += $ControlHeight
	$MacroWidthLabel = Label("Width",$Margin,$PositionY)
	$MacroWidth = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Width in pixels",0)
	$MacroWindowWindowLabel = Label("Window by keyword",$Margin,$PositionY)
	$MacroWindowWindow = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter a keyword of the title of the window you want to be activated")
	$PositionY += $ControlHeight
	$MacroHeightLabel = Label("Height",$Margin,$PositionY)
	$MacroHeight = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Height in pixels",0)
	$MacroWindowPosition = GUICtrlCreateCombo("",$Margin+$LabelWidth+$Margin+$InputWidth,$PositionY-4-$ControlHeight/2,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a window position")
	_GUICtrlComboBox_AddArray($MacroWindowPosition,$MacroWindowPositions)
	$MacroWindowSize = GUICtrlCreateCombo("",$Margin+$LabelWidth+$Margin+$InputWidth,$PositionY-4-$ControlHeight/2,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a window size type")
	_GUICtrlComboBox_AddArray($MacroWindowSize,$MacroWindowSizes)
	$PositionY += $ControlHeight
	Checkbox("Move instantaneously",$Margin+$LabelWidth+$Margin+$InputWidth,$PositionY-3,-1,-1,"Move window to new coordinates without the macro delay",$MacroWindowMoveInstantaneous,$MacroWindowMoveInstantaneousLabel)
	Checkbox("Resize instantaneously",$Margin+$LabelWidth+$Margin+$InputWidth,$PositionY-3,-1,-1,"Resize window to new size without the macro delay",$MacroWindowResizeInstantaneous,$MacroWindowResizeInstantaneousLabel)
	; randomize
	$PositionY = $Margin+2*$ControlHeight
	$MacroRandomMouseLabel = Label("Mouse movement",$Margin,$PositionY)
	$MacroRandomMouse = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Randomize mouse movement in pixels to mimic human behaviour, 0 = set mouse randomize off",0)
	$PositionY += $ControlHeight
	$MacroRandomDelayLabel = Label("Step delay",$Margin,$PositionY)
	$MacroRandomDelay = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Randomize step delay in milliseconds to mimic human behaviour, 0 = set delay randomize off",0)
	; jump
	$PositionY = $Margin+2*$ControlHeight
	$MacroJumpTypeLabel = Label("Jump type",$Margin,$PositionY)
	$MacroJumpType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a jump action to perform")
	_GUICtrlComboBox_AddArray($MacroJumpType,$MacroJumpTypes)
	$PositionY += $ControlHeight
	$MacroJumpLabelLabel = Label("Label to jump to",$Margin,$PositionY)
	$MacroJumpLabel = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a label to jump to")
	; repeat
	$PositionY = $Margin+2*$ControlHeight
	$MacroRepeatTypeLabel = Label("Repeat action",$Margin,$PositionY)
	$MacroRepeatType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a repeat action to perform, repeat loop or if condition")
	_GUICtrlComboBox_AddArray($MacroRepeatType,$MacroRepeatTypes)
	$PositionY += $ControlHeight
	$MacroRepeatCountLabel = Label("Number of times",$Margin,$PositionY)
	$MacroRepeatCount = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Repeat block of macro steps a number of times, 0 = infinite loop",0)
	$MacroRepeatValue1aLabel = Label("Repeat counter =",$Margin,$PositionY)
	$MacroRepeatValue1bLabel = Label("Repeat counter not =",$Margin,$PositionY)
	$MacroRepeatValue1cLabel = Label("Repeat counter > or =",$Margin,$PositionY)
	$MacroRepeatValue1 = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Equal to value",0)
	$PositionY += $ControlHeight
	$MacroRepeatValue2Label = Label("Repeat counter < or =",$Margin,$PositionY)
	$MacroRepeatValue2 = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Smaller or equal than value",0)
	; time loop
	$PositionY = $Margin+2*$ControlHeight
	$MacroTimeLoopLabel = Label("Time in ms",$Margin,$PositionY)
	$MacroTimeLoop = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Total time to loop block in milliseconds",0)
	; input
	$PositionY = $Margin+2*$ControlHeight
	$MacroInputTypeLabel = Label("Input type",$Margin,$PositionY)
	$MacroInputType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select an input type")
	_GUICtrlComboBox_AddArray($MacroInputType,$MacroInputTypes)
	$PositionY += $ControlHeight
	$MacroContinueKeyLabel = Label("Button/key to continue",$Margin,$PositionY)
	$MacroContinueKey = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a mouse button or key to press to continue macro")
	_GUICtrlComboBox_AddArray($MacroContinueKey,$MacroContinueKeys)
	; pixels
	$PositionY = $Margin+2*$ControlHeight
	$MacroPixelTypeLabel = Label("Pixel set type",$Margin,$PositionY)
	$MacroPixelType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a pixel set type")
	_GUICtrlComboBox_AddArray($MacroPixelType,$MacroPixelTypes)
	$PositionY = $Margin+5*$ControlHeight
	$MacroPixelWidthLabel = Label("Area width",$Margin,$PositionY)
	$MacroPixelWidth = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Width of pixel area to check for changes",0)
	$MacroPixelHeightLabel = Label("Area height",$Margin+$Margin+$ComboWidth,$PositionY)
	$MacroPixelHeight = Input($Margin+$LabelWidth+$Margin+$ComboWidth,$PositionY-3,$InputWidth,$InputHeight,"Height of pixel area to check for changes",0)
	$PositionY = $Margin+6*$ControlHeight
	$MacroPixelIdLabel1 = Label("Area id number",$Margin,$PositionY)
	$MacroPixelIdLabel2 = Label("Pixel id number",$Margin,$PositionY)
	$MacroPixelId = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Enter a number for identification of area or pixel",0)
	; alarm
	$PositionY = $Margin+2*$ControlHeight
	$MacroAlarmTypeLabel = Label("Alarm type",$Margin,$PositionY)
	$MacroAlarmType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a alarm type")
	_GUICtrlComboBox_AddArray($MacroAlarmType,$MacroAlarmTypes)
	$PositionY += $ControlHeight
	$MacroAlarmTimeLabel = Label("Time of alarm",$Margin,$PositionY)
	$MacroAlarmTime = GUICtrlCreateDate($GUIStep,$Margin+$LabelWidth,$PositionY-3,90,$InputHeight,$DTS_TIMEFORMAT)
	$PositionY += $ControlHeight
	$MacroAlarmSpecificLabel = Label("When and how",$Margin,$PositionY)
	$MacroAlarmSpecific = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select when block is runned and how: once or always")
	_GUICtrlComboBox_AddArray($MacroAlarmSpecific,$MacroAlarmSpecifics)
	$PositionY += $ControlHeight
	$MacroAlarmDateLabel = Label("Date of alarm",$Margin,$PositionY)
	$MacroAlarmDate = GUICtrlCreateDate($GUIStep,$Margin+$LabelWidth,$PositionY-3,200,$InputHeight,$DTS_LONGDATEFORMAT)
	; sound
	$PositionY = $Margin+2*$ControlHeight
	$MacroSoundTypeLabel = Label("Sound type",$Margin,$PositionY)
	$MacroSoundType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a sound type")
	_GUICtrlComboBox_AddArray($MacroSoundType,$MacroSoundTypes)
	$PositionY += $ControlHeight
	$MacroSoundFileLabel = Label("Sound file",$Margin,$PositionY)
	$MacroSoundFile = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter a sound file")
	$MacroSoundSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select  a sound file",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	$MacroSoundFrequencyLabel = Label("Beep frequency",$Margin,$PositionY)
	$MacroSoundFrequency = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Set beep frequency in Herz",0)
	$PositionY += $ControlHeight
	$MacroSoundVolumeLabel = Label("Volume 0 to 100%",$Margin,$PositionY)
	$MacroSoundVolume = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Sound volume percentage from 0 to 100%")
	$MacroSoundVolumeUpDown = GUICtrlCreateUpdown($MacroSoundVolume)
	$MacroSoundDurationLabel = Label("Beep duration",$Margin,$PositionY)
	$MacroSoundDuration = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Set beep duration in milliseconds",0)
	; counter
	$PositionY = $Margin+2*$ControlHeight
	$MacroCounterTypeLabel = Label("Counter action",$Margin,$PositionY)
	$MacroCounterType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a counter action to perform, set counter or if condition")
	_GUICtrlComboBox_AddArray($MacroCounterType,$MacroCounterTypes)
	$PositionY += $ControlHeight
	$MacroCounterNameLabel = Label("Counter name",$Margin,$PositionY)
	$MacroCounterName = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Name of the counter")
	$MacroCounterNameList = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a counter")
	$PositionY += $ControlHeight
	$MacroCounterSetLabel = Label("Set counter value",$Margin,$PositionY)
	$MacroCounterSet = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,-1)
	Checkbox("Randomize value",$Margin+$LabelWidth+$InputWidth+10,$PositionY-3,-1,-1,"Value will be implemented as random number",$MacroCounterRandom,$MacroCounterRandomLabel)
	$MacroCounterValue1aLabel = Label("Counter =",$Margin,$PositionY)
	$MacroCounterValue1bLabel = Label("Counter not =",$Margin,$PositionY)
	$MacroCounterValue1cLabel = Label("Counter > or =",$Margin,$PositionY)
	$MacroCounterValue1 = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,-1)
	$PositionY += $ControlHeight
	$MacroCounterValue2Label = Label("Counter < or =",$Margin,$PositionY)
	$MacroCounterValue2 = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Smaller or equal than value",-1)
	SetCounterLabelsAndTips()
	; file/path
	$PositionY = $Margin+2*$ControlHeight
	$MacroFileTypeLabel = Label("File or path action",$Margin,$PositionY)
	$MacroFileType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a file or path action")
	_GUICtrlComboBox_AddArray($MacroFileType,$MacroFileTypes)
	$PositionY += $ControlHeight
	$MacroFileFileLabel = Label("File to check",$Margin,$PositionY)
	$MacroFileFile = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter a file folder and name to check")
	$MacroFileFileSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the file to check",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	$MacroFilePathLabel = Label("Folder to check",$Margin,$PositionY)
	$MacroFilePath = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter a folder to check")
	$MacroFilePathSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the folder to check",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	$MacroFileStampLabel = Label("File to record change",$Margin,$PositionY)
	$MacroFileStamp = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter a file to record a stamp for checking a change later on")
	$MacroFileStampSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the file to record change",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	$MacroFileChangeLabel = Label("File to check for change",$Margin,$PositionY)
	$MacroFileChange = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter a file to check if it has been changed or not")
	$MacroFileChangeSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the file to check if changed or not",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	; control
	$PositionY = $Margin+2*$ControlHeight
	$MacroControlTypeLabel = Label("Control action",$Margin,$PositionY)
	$MacroControlType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a macro control action to perform")
	_GUICtrlComboBox_AddArray($MacroControlType,$MacroControlTypes)
	; process
	$PositionY = $Margin+2*$ControlHeight
	$MacroProcessTypeLabel = Label("Select process action",$Margin,$PositionY)
	$MacroProcessType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a process action")
	_GUICtrlComboBox_AddArray($MacroProcessType,$MacroProcessTypes)
	$PositionY += $ControlHeight
	$MacroProcessWindowTypeLabel = Label("Window storage",$Margin,$PositionY)
	$MacroProcessWindowType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a window storage method")
	_GUICtrlComboBox_AddArray($MacroProcessWindowType,$MacroProcessWindowTypes)
	$MacroProcessProgramLabel = Label("Program to start",$Margin,$PositionY)
	$MacroProcessProgram = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter the program with path to start (A program which needs administrator rights won't start)")
	$MacroProcessProgramSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the program to start (A program which needs administrator rights won't start)",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	$MacroProcessProcessLabel = Label("Process to check",$Margin,$PositionY)
	$MacroProcessProcess = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter the process executable to check existence of")
	$MacroProcessProcessSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the process executable to check existence of",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	$PositionY += $ControlHeight
	$MacroProcessWindowLabel = Label("Window by keyword",$Margin,$PositionY)
	$MacroProcessWindow = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter a keyword of the title of the window you want to be activated")
	$PositionY += $ControlHeight
	$MacroProcessShowProcessCurrentLabel = Label("Switch to process",$Margin,$PositionY)
	$MacroProcessShowProcessCurrent = Label("",$Margin+$LabelWidth,$PositionY,$GUIStepWidth-2*$Margin-$LabelWidth,3*$InputHeight)
	$MacroProcessWorkDirLabel = Label("Working folder",$Margin,$PositionY)
	$MacroProcessWorkDir = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter the working folder")
	$MacroProcessWorkDirSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the working folder of the program",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	GUICtrlSetTip(-1,"Select the working folder")
	$PositionY += $ControlHeight
	$MacroProcessShowProcessLabel = Label("Step currently set to",$Margin,$PositionY)
	$MacroProcessShowProcess = Label("",$Margin+$LabelWidth,$PositionY,$GUIStepWidth-2*$Margin-$LabelWidth,3*$InputHeight)
	$MacroProcessWindowStateLabel = Label("Window action",$Margin,$PositionY)
	$MacroProcessWindowState =  GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a window action to perform")
	_GUICtrlComboBox_AddArray($MacroProcessWindowState,$MacroProcessWindowStates)
	$PositionY += $ControlHeight
	Checkbox("Pause macro until you close the program",$Margin+$LabelWidth,$PositionY-3,-1,-1,"The macro is paused, you close the started program and it will continue",$MacroProcessWait,$MacroProcessWaitLabel)
	$MacroProcessTest = _GraphicButton(@ScriptDir & "\detect.png","Test","Start program to test this step is working",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-15,$ButtonWidth,$ButtonHeight)
	$PositionY += $ControlHeight
	$MacroProcessTimeoutLabel = Label("Timeout for activating",$Margin,$PositionY)
	$MacroProcessTimeout = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Timeout in milliseconds for trying to activate the window of the started program",0)
	; capture
	$PositionY = $Margin+2*$ControlHeight
	$MacroCaptureActionLabel = Label("Capture action",$Margin,$PositionY)
	$MacroCaptureAction = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a capture action to perform")
	_GUICtrlComboBox_AddArray($MacroCaptureAction,$MacroCaptureActions)
	$MacroCaptureDetect = _GraphicButton(@ScriptDir & "\detect.png","Detect","Detect top left corner of area by a mouse click at a position",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-6,$ButtonWidth,$ButtonHeight)
	$PositionY += $ControlHeight
	$MacroCapturePathLabel = Label("Folder to store images",$Margin,$PositionY)
	$MacroCapturePath = Input($Margin+$LabelWidth,$PositionY-3,$GUIStepWidth-$LabelWidth-3*$Margin-$ButtonWidth,$InputHeight,"Enter the folder to store the captured images" & @CRLF & "If left empty, folder MyDocuments\" & $ProgramName & " is used")
	$MacroCapturePathSelect = _GraphicButton(@ScriptDir & "\open.png","Select","Select the folder to store the captured images",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-5,$ButtonWidth,$ButtonHeight)
	$MacroCaptureXLabel = Label("Area x position",$Margin,$PositionY)
	$MacroCaptureScreenX = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"X position of screen area",0)
	$MacroCaptureWindowX = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"X position of area in window",0)
	$PositionY += $ControlHeight
	$MacroCaptureYLabel = Label("Area y position",$Margin,$PositionY)
	$MacroCaptureScreenY = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Y position of screen area",0)
	$MacroCaptureWindowY = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Y position of area in window",0)
	$MacroCaptureShow = _GraphicButton(@ScriptDir & "\pointer.png","Show","",$GUIStepWidth-$Margin-$ButtonWidth,$PositionY-6-$ControlHeight/2,$ButtonWidth,$ButtonHeight)
	Checkbox("Capture the mouse pointer in images",$Margin+$LabelWidth,$PositionY-3,-1,-1,"",$MacroCapturePointer,$MacroCapturePointerLabel)
	$PositionY += $ControlHeight
	$MacroCaptureWidthLabel = Label("Area width",$Margin,$PositionY)
	$MacroCaptureScreenWidth = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Width of screen area, -1 = whole area",-1)
	$MacroCaptureWindowWidth = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Width of window area, -1 = whole area",-1)
	$PositionY += $ControlHeight
	$MacroCaptureHeightLabel = Label("Area height",$Margin,$PositionY)
	$MacroCaptureScreenHeight = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Height of screen area, -1 = whole area",-1)
	$MacroCaptureWindowHeight = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Height of window area, -1 = whole area",-1)
	; shutdown
	$PositionY = $Margin+2*$ControlHeight
	$MacroShutDownActionLabel = Label("Shutdown action",$Margin,$PositionY)
	$MacroShutDownAction = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a computer shut down action to perform")
	_GUICtrlComboBox_AddArray($MacroShutDownAction,$MacroShutDownActions)
	$PositionY += $ControlHeight
	Checkbox("Confirm action by asking",$Margin+$LabelWidth,$PositionY-3,-1,-1,"Before doing the shutdown action, you have to confirm it",$MacroShutDownAsk,$MacroShutDownAskLabel)
	; comment
	$PositionY = $Margin+2*$ControlHeight
	$MacroCommentTypeLabel = Label("Comment type",$Margin,$PositionY)
	$MacroCommentType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a comment type")
	_GUICtrlComboBox_AddArray($MacroCommentType,$MacroCommentTypes)

	$MacroSave = _GraphicButton(@ScriptDir & "\save.png","Save","Save this step, Ctrl+Enter",$Margin,$GUIStepHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	$MacroStepHelp = _GraphicButton(@ScriptDir & "\questionmark.png","","Get help how to use this window",$GUIStepWidth/2-$SmallButtonWidth/2,$GUIStepHeight-$Margin-$ButtonHeight,$SmallButtonWidth,$ButtonHeight)
	$MacroCancel = _GraphicButton(@ScriptDir & "\close.png","Cancel","Cancel this step",$GUIStepWidth-$Margin-$ButtonWidth,$GUIStepHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	Local $GUIStepKeys[1][2] = [["{ENTER}",$MacroSave]]
    GUISetAccelerators($GUIStepKeys,$GUIStep)

	; build copy steps interface
	$GUICopy = GUICreate("",$GUICopyWidth,$GUICopyHeight,-1,-1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU,$WS_EX_TOPMOST)
	If $ProgramTheme > -1 Then GUISetBkColor($Themes[$ProgramTheme][0])
	$PositionY = $Margin
	Label("Start step of range",$Margin,$PositionY)
	$CopyFromStep = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Enter start step of range",0)
	GUICtrlSetData(-1,1)
	$PositionY += $ControlHeight
	Label("End step of range",$Margin,$PositionY)
	$CopyToStep = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Enter end step of range",0)
	GUICtrlSetData(-1,1)
	$CopyCopy = _GraphicButton(@ScriptDir & "\copy.png","Copy","Copy range",$Margin,$GUICopyHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	$CopyCut = _GraphicButton(@ScriptDir & "\cut.png","Cut","Cut range",$Margin,$GUICopyHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	$CopyClose = _GraphicButton(@ScriptDir & "\close.png","Close","Close this window",$GUICopyWidth-$Margin-$ButtonWidth,$GUICopyHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)

	; build enable/disable steps interface
	$GUIEnableDisable = GUICreate("",$GUIEnableDisableWidth,$GUIEnableDisableHeight,-1,-1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU,$WS_EX_TOPMOST)
	If $ProgramTheme > -1 Then GUISetBkColor($Themes[$ProgramTheme][0])
	$PositionY = $Margin
	Label("Start step of range",$Margin,$PositionY)
	$EnableDisableFromStep = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Enter start step of range",0)
	GUICtrlSetData(-1,1)
	$PositionY += $ControlHeight
	Label("End step of range",$Margin,$PositionY)
	$EnableDisableToStep = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Enter end step of range",0)
	GUICtrlSetData(-1,1)
	$EnableDisableEnable = _GraphicButton(@ScriptDir & "\enable.png","Enable","Enable range",$Margin,$GUICopyHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	$EnableDisableDisable = _GraphicButton(@ScriptDir & "\disable.png","Disable","Disable range",$Margin+$ButtonWidth+$Margin,$GUICopyHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	$EnableDisableClose = _GraphicButton(@ScriptDir & "\close.png","Close","Close this window",$GUIEnableDisableWidth-$Margin-$ButtonWidth,$GUICopyHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)

	; build find interface
	$GUIFind = GUICreate("",$GUIFindWidth,$GUIFindHeight,-1,-1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU,$WS_EX_TOPMOST)
	If $ProgramTheme > -1 Then GUISetBkColor($Themes[$ProgramTheme][0])
	$PositionY = $Margin
	Label("Find step by",$Margin,$PositionY)
	$FindType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select how to find a step")
	_GUICtrlComboBox_AddArray($FindType,$FindTypes)
	$PositionY += $ControlHeight
	$FindKeywordLabel = Label("Find keyword",$Margin,$PositionY)
	$FindKeyword = Input($Margin+$LabelWidth,$PositionY-3,$ComboWidth,$InputHeight,"Enter a keyword of the step description to find a step")
	$FindStepTypeLabel = Label("Find step type",$Margin,$PositionY)
	GUICtrlSetState($FindStepTypeLabel,$GUI_HIDE)
	$FindStepType = GUICtrlCreateCombo("",$Margin+$LabelWidth,$PositionY-4,$ComboWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetTip(-1,"Select a step type to find")
	_GUICtrlComboBox_AddArray($FindStepType,$MacroStepTypes)
	GUICtrlSetState($FindStepType,$GUI_HIDE)
	$FindFind = _GraphicButton(@ScriptDir & "\detect.png","Find","Find next step with search condition",$Margin,$GUIFindHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	$FindClose = _GraphicButton(@ScriptDir & "\close.png","Close","Close this window",$GUIFindWidth-$Margin-$ButtonWidth,$GUIFindHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	Local $GUIFindKeys[1][2] = [["{ENTER}",$FindFind]]
    GUISetAccelerators($GUIFindKeys,$GUIFind)

	; build go to step interface
	$GUIGoTo = GUICreate("",$GUIGoToWidth,$GUIGoToHeight,-1,-1,$WS_CAPTION+$WS_POPUP+$WS_SYSMENU,$WS_EX_TOPMOST)
	If $ProgramTheme > -1 Then GUISetBkColor($Themes[$ProgramTheme][0])
	$PositionY = $Margin
	Label("Go to step number",$Margin,$PositionY)
	$FindGoTo = Input($Margin+$LabelWidth,$PositionY-3,$InputWidth,$InputHeight,"Enter step to go to",0)
	GUICtrlSetData(-1,1)
	$FindGoToJump = _GraphicButton(@ScriptDir & "\detect.png","Go","Go to entered step",$Margin,$GUIGoToHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	$FindGoToClose = _GraphicButton(@ScriptDir & "\close.png","Close","Close this window",$GUIGoToWidth-$Margin-$ButtonWidth,$GUIGoToHeight-$Margin-$ButtonHeight,$ButtonWidth,$ButtonHeight)
	Local $GUIGoToKeys[1][2] = [["{ENTER}",$FindGoToJump]]
    GUISetAccelerators($GUIGoToKeys,$GUIGoTo)

	; build main interface
	$GUIPosition[2] = _Maximum($GUIPosition[2],$GUIWidth)
	$GUIPosition[3] = _Maximum($GUIPosition[3],$GUIHeight)
	$ListWidth = $GUIPosition[2]-$ControlsWidth
	$ListHeight = $GUIPosition[3]-$ControlsHeight
	$GUI = GUICreate("",$GUIPosition[2],$GUIPosition[3],$GUIPosition[0],$GUIPosition[1],$WS_MINIMIZEBOX+$WS_CAPTION+$WS_POPUP+$WS_SYSMENU+$WS_SIZEBOX)
	If $ProgramTheme > -1 Then GUISetBkColor($Themes[$ProgramTheme][0])
	SetWindowTitle()							; show program name/version and macro file in main window title bar
	SetF1($GUI,$HelpFile)						; pop up manual by pressing F1
	; build program menu
	$MenuFile = GUICtrlCreateMenu("&File")
	$MenuEdit = GUICtrlCreateMenu("&Edit")
	$MenuSearch = GUICtrlCreateMenu("&Search")
	$MenuOptions = GUICtrlCreateMenu("&Options")
	$MenuHelp = GUICtrlCreateMenu("&Help")
	$MenuItemNew = GUICtrlCreateMenuItem("&New" & @TAB & "Ctrl+N",$MenuFile)
	$MenuItemOpen = GUICtrlCreateMenuItem("&Open" & @TAB & "Ctrl+O",$MenuFile)
	$MenuItemRecent = GUICtrlCreateMenu("Recent macro's",$MenuFile)
	ReadRecentFiles()
	GUICtrlCreateMenuItem("",$MenuFile)
	$MenuItemSave = GUICtrlCreateMenuItem("&Save" & @TAB & "Ctrl+S",$MenuFile)
	$MenuItemSaveAs = GUICtrlCreateMenuItem("Save &As",$MenuFile)
	GUICtrlCreateMenuItem("",$MenuFile)
	$MenuItemExit = GUICtrlCreateMenuItem("&Exit",$MenuFile)
	$MenuItemUndo = GUICtrlCreateMenuItem("&Undo" & @TAB & "Ctrl+Z",$MenuEdit)
	GUICtrlSetState(-1,$GUI_DISABLE)
	$MenuItemRedo = GUICtrlCreateMenuItem("&Redo" & @TAB & "Ctrl+Y",$MenuEdit)
	GUICtrlSetState(-1,$GUI_DISABLE)
	GUICtrlCreateMenuItem("",$MenuEdit)
	$MenuItemCut = GUICtrlCreateMenuItem("Cu&t" & @TAB & "Ctrl+X",$MenuEdit)
	$MenuItemCopy = GUICtrlCreateMenuItem("&Copy" & @TAB & "Ctrl+C",$MenuEdit)
	GUICtrlCreateMenuItem("",$MenuEdit)
	$MenuItemCutSteps = GUICtrlCreateMenuItem("Cut range of steps",$MenuEdit)
	$MenuItemCopySteps = GUICtrlCreateMenuItem("Copy range of steps",$MenuEdit)
	GUICtrlCreateMenuItem("",$MenuEdit)
	$MenuItemPaste = GUICtrlCreateMenuItem("&Paste" & @TAB & "Ctrl+V",$MenuEdit)
	GUICtrlSetState(-1,$GUI_DISABLE)
	$MenuItemPasteAfter = GUICtrlCreateMenuItem("Paste &after" & @TAB & "Shift+Ctrl+V",$MenuEdit)
	GUICtrlSetState(-1,$GUI_DISABLE)
	GUICtrlCreateMenuItem("",$MenuEdit)
	$MenuItemEnableDisableSteps = GUICtrlCreateMenuItem("Enable or Disable steps",$MenuEdit)
	$MenuItemFind = GUICtrlCreateMenuItem("&Find" & @TAB & "Ctrl+F",$MenuSearch)
	$MenuItemFindNext = GUICtrlCreateMenuItem("Find &next" & @TAB & "F3",$MenuSearch)
	GUICtrlSetState(-1,$GUI_DISABLE)
	$MenuItemFindPrevious = GUICtrlCreateMenuItem("Find &previous" & @TAB & "Shift+F3",$MenuSearch)
	GUICtrlSetState(-1,$GUI_DISABLE)
	GUICtrlCreateMenuItem("",$MenuSearch)
	$MenuItemFindGoTo = GUICtrlCreateMenuItem("&Go to step number" & @TAB & "Ctrl+G",$MenuSearch)
	GUICtrlCreateMenuItem("",$MenuSearch)
	$MenuItemMark = GUICtrlCreateMenuItem("&Bookmark step" & @TAB & "Ctrl+F2",$MenuSearch)
	$MenuItemMarkNext = GUICtrlCreateMenuItem("Next bookmark" & @TAB & "F2",$MenuSearch)
	$MenuItemMarkPrevious = GUICtrlCreateMenuItem("Previous bookmark" & @TAB & "Shift+F2",$MenuSearch)
	GUICtrlCreateMenuItem("",$MenuSearch)
	$MenuItemDisabledNext = GUICtrlCreateMenuItem("Next disabled step" & @TAB & "F4",$MenuSearch)
	$MenuItemDisabledPrevious = GUICtrlCreateMenuItem("Previous disabled step" & @TAB & "Shift+F4",$MenuSearch)
	$MenuItemSettings = GUICtrlCreateMenuItem("&Settings",$MenuOptions)
	$MenuItemHelp = GUICtrlCreateMenuItem("&Help" & @TAB & "F1",$MenuHelp)
	$MenuItemSite = GUICtrlCreateMenuItem($ProgramName & " website",$MenuHelp)
	$MenuItemAbout = GUICtrlCreateMenuItem("&About",$MenuHelp)
	; process controls
	$ProcessesLabel = Label("Select a process",$Margin,$Margin,-1,-1,"",1)
	$WindowsVisible = Button(@ScriptDir & "\visible.png","","Click to hide invisible windows",$GUIPosition[2]-$Margin-$ViewButtonWidth,$Margin-5,$ViewButtonWidth,$ButtonHeight-1,0)
	$WindowsInvisible = Button(@ScriptDir & "\invisible.png","","Click to show invisible windows too",$GUIPosition[2]-$Margin-$ViewButtonWidth,$Margin-5,$ViewButtonWidth,$ButtonHeight-1,0)
	If $ShowInvisible Then
		GUICtrlSetState($WindowsInvisible,$GUI_HIDE)
		GUICtrlSetState($WindowsVisible,$GUI_SHOW)
	Else
		GUICtrlSetState($WindowsInvisible,$GUI_SHOW)
		GUICtrlSetState($WindowsVisible,$GUI_HIDE)
	EndIf
	$Processes = GUICtrlCreateCombo("",$Margin+$ProcessLabelWidth,$Margin-4,$GUIPosition[2]-2*$Margin-$ProcessLabelWidth-$ViewButtonMargin-$ViewButtonWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKTOP)
	GUICtrlSetTip(-1,"Select a process to perform macro on")
	_GUICtrlComboBox_SetDroppedWidth($Processes,$GUIPosition[2]-2*$Margin-$ProcessLabelWidth)
	; macro steps controls
	$ViewX = $Margin+$ListWidth-$ViewButtonWidth
	$StepViewList = Button(@ScriptDir & "\viewlist.png","","View with small rows and vertically columns",$ViewX,$Margin+$ControlHeight,$ViewButtonWidth,$ViewButtonHeight,0)
	$ViewX -= $ViewButtonWidth+$ViewButtonMargin
	$StepViewReport = Button(@ScriptDir & "\viewreport.png","","View with rows in one column",$ViewX,$Margin+$ControlHeight,$ViewButtonWidth,$ViewButtonHeight,0)
	$ViewX -= $ViewButtonWidth+$ViewButtonMargin
	$StepViewGrid = Button(@ScriptDir & "\viewgrid.png","","View with rows between lines in one column",$ViewX,$Margin+$ControlHeight,$ViewButtonWidth,$ViewButtonHeight,0)
	$ViewX -= $ViewButtonWidth+2*$ViewButtonMargin
	$StepViewInverseColor = Button(@ScriptDir & "\invertcolor.png","","",$ViewX,$Margin+$ControlHeight,$ViewButtonWidth,$ViewButtonHeight,0)
	If $ViewInvertColors Then
		GUICtrlSetTip(-1,"Set colors to normal")
	Else
		GUICtrlSetTip(-1,"Invert colors")
	EndIf
	$ViewX -= $ViewButtonWidth+2*$ViewButtonMargin
	$StepViewNumbers = Button(@ScriptDir & "\numbers.png","","",$ViewX,$Margin+$ControlHeight,$ViewButtonWidth,$ViewButtonHeight,0)
	If $ViewNumbers Then
		GUICtrlSetTip(-1,"Hide the step line numbers")
	Else
		GUICtrlSetTip(-1,"Show the step line numbers")
	EndIf
	$MacroStepsLabel = Label("Macro steps",$Margin,$Margin+$ControlHeight,$ViewX-$Margin-2,15,"",1)
	; macro steps list
	$PositionY = $Margin+1.7*$ControlHeight
	$Macros = GUICtrlCreateListView("",$Margin,$PositionY,$ListWidth,$ListHeight,$LVS_LIST+$LVS_NOCOLUMNHEADER+$LVS_SINGLESEL,$LVS_EX_FULLROWSELECT+$WS_EX_CLIENTEDGE)
	If $ProgramTheme > -1 Then GUICtrlSetBkColor(-1,$Themes[$ProgramTheme][2])
	GUICtrlSetTip(-1,"Sequencial steps of the macro")
	GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKTOP)
	_GUICtrlListView_AddColumn($Macros,"",$ListWidth-_WinAPI_GetSystemMetrics($SM_CXVSCROLL)-2*_WinAPI_GetSystemMetrics($SM_CXEDGE))
	SetStepView($StepView)
	$ListImage = _GUIImageList_Create(6,20,5,3,5,5)
	_GUIImageList_AddBitmap($ListImage,@ScriptDir & "\disabled.bmp")
	_GUIImageList_AddBitmap($ListImage,@ScriptDir & "\mark.bmp")
	_GUIImageList_AddBitmap($ListImage,@ScriptDir & "\disabledmark.bmp")
	_GUIImageList_AddBitmap($ListImage,@ScriptDir & "\error.bmp")
	_GUIImageList_AddBitmap($ListImage,@ScriptDir & "\disablederror.bmp")
	_GUICtrlListView_SetImageList($Macros,$ListImage,1)
	; macro steps edit controls
	$Add = Button(@ScriptDir & "\plus.png","Add","Add a macro step, Ctrl++",$Margin+$ListWidth+$Margin,$PositionY,$ButtonWidth,$ButtonHeight,0)
	$Change = Button(@ScriptDir & "\edit.png","Change","Change selected macro step, Ctrl+Enter",$Margin+$ListWidth+$Margin,$PositionY+$ButtonHeight+$Margin,$ButtonWidth,$ButtonHeight,0)
	$Duplicate = Button(@ScriptDir & "\duplicate.png","Duplicate","Duplicate selected macro step, Ctrl+D",$Margin+$ListWidth+$Margin,$PositionY+2*($ButtonHeight+$Margin),$ButtonWidth,$ButtonHeight,0)
	$Delete = Button(@ScriptDir & "\minus.png","Delete","Delete selected macro step, Ctrl+Delete",$Margin+$ListWidth+$Margin,$PositionY+3*($ButtonHeight+$Margin),$ButtonWidth,$ButtonHeight,0)
	$Up = Button(@ScriptDir & "\arrowup.png","Move up","Move selected macro step up, Ctrl+Up",$Margin+$ListWidth+$Margin,$PositionY+4*($ButtonHeight+$Margin),$ButtonWidth,$ButtonHeight,0)
	$Down = Button(@ScriptDir & "\arrowdown.png","Move down","Move selected macro step down, Ctrl+Down",$Margin+$ListWidth+$Margin,$PositionY+5*($ButtonHeight+$Margin),$ButtonWidth,$ButtonHeight,0)
	$Show = Button(@ScriptDir & "\pointer.png","Show","",$Margin+$ListWidth+$Margin,$PositionY+6*($ButtonHeight+$Margin),$ButtonWidth,$ButtonHeight,0)
	$Enable = Button(@ScriptDir & "\enable.png","Enable","Activate disabled step again",$Margin+$ListWidth+$Margin,$PositionY+7*($ButtonHeight+$Margin),$ButtonWidth,$ButtonHeight,0)
	$Disable = Button(@ScriptDir & "\disable.png","Disable","Disable step, it will be skipped during macro play",$Margin+$ListWidth+$Margin,$PositionY+7*($ButtonHeight+$Margin),$ButtonWidth,$ButtonHeight,0)
	; play and record controls
	$PositionY += $ListHeight
	Checkbox("Show recording",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-9*$ControlHeight-$Margin,-1,-1,"If checked, " & $ProgramName & " updates the macro step list on every mouse click and key press",$ShowRecording,$ShowRecordingLabel,0)
	$MoveTimeLabel = Label("Move",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-7.9*$ControlHeight-$Margin,-1,-1,"",0)
	$MoveTime = Input($Margin+$ListWidth+$Margin+40,$PositionY-$ButtonHeight-7.9*$ControlHeight-$Margin-3,$ButtonWidth-50,$InputHeight,"Record mouse movement at an interval in milliseconds, 0 = Don't record mouse movement",0,0)
	$Record = Button(@ScriptDir & "\record.png","Record","Record mouse clicks and key presses, Ctrl+R",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-7*$ControlHeight-$Margin,$ButtonWidth,$ButtonHeight,0)
	Checkbox("Test, no clicks",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-6*$ControlHeight-$Margin,-1,-1,"If checked, mouse is only moved to the specified positions but no mouse buttons are pressed and no keys are pressed",$TestMacro,$TestMacroLabel,0)
	$ShowMain =  GUICtrlCreateCombo("",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-5*$ControlHeight-$Margin,$ButtonWidth,$ButtonHeight,$CBS_DROPDOWNLIST+$WS_VSCROLL)
	GUICtrlSetResizing(-1,$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKSIZE)
	GUICtrlSetTip(-1,"Whilst playing show, minized or hide " & $ProgramName)
	_GUICtrlComboBox_AddArray($ShowMain,$MainWindowActions)
	Checkbox("Show progress",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-4*$ControlHeight-$Margin,-1,-1,"If checked, show a macro progress bar during playing",$ShowBar,$ShowBarLabel,0)
	$StepDelayLabel = Label("Delay",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-2.8*$ControlHeight-$Margin,-1,-1,"",0)
	$StepDelay = Input($Margin+$ListWidth+$Margin+40,$PositionY-$ButtonHeight-2.8*$ControlHeight-$Margin-3,$ButtonWidth-50,$InputHeight,"Delay for time critical steps in milliseconds" & @CRLF & "Note: A small delay may cause macro playing problems",0,0)
	Checkbox("On every step",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-2.1*$ControlHeight-$Margin,-1,-1,"Use delay on every single macro step",$StepDelayAlways,$StepDelayAlwaysLabel,0)
	$LoopCountLabel = Label("Loop",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-$ControlHeight-$Margin,-1,-1,"",0)
	$LoopCount = Input($Margin+$ListWidth+$Margin+40,$PositionY-$ButtonHeight-$ControlHeight-$Margin-3,$ButtonWidth-50,$InputHeight,"Number of loops when looping macro, 0 = Run indefinitely until stopped",0,0)
	$Play = Button(@ScriptDir & "\play.png","Play","",$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-$Margin,$ButtonWidth,$ButtonHeight,0)
	SetPlayTip()
	$PositionY = $GUIPosition[3]-_WindowMenuHeight()-$Margin-$ButtonHeight
	$New = Button(@ScriptDir & "\new.png","New","Create a new macro, Ctrl+N",$Margin,$PositionY,$ButtonWidth,$ButtonHeight,1)
	$Open = Button(@ScriptDir & "\open.png","Open","Open a saved macro, Ctrl+O",$Margin+$ButtonWidth+$Margin,$PositionY,$ButtonWidth,$ButtonHeight,1)
	$Save = Button(@ScriptDir & "\save.png","Save","Save this macro, Ctrl+S",$Margin+2*($ButtonWidth+$Margin),$PositionY,$ButtonWidth,$ButtonHeight,1)
	GUICtrlSetState(-1,$GUI_DISABLE)
	$SaveAs = Button(@ScriptDir & "\saveas.png","Save As","Save this macro to a different file",$Margin+3*($ButtonWidth+$Margin),$PositionY,$ButtonWidth,$ButtonHeight,1)
	GUICtrlSetState(-1,$GUI_DISABLE)
	$Close = Button(@ScriptDir & "\close.png","Exit","Close " & $ProgramName,$GUIPosition[2]-$Margin-$ButtonWidth,$PositionY,$ButtonWidth,$ButtonHeight,2)
	$Settings = Button(@ScriptDir & "\settings.png","","Go to settings",$Margin+$ListWidth-$SmallButtonWidth,$PositionY,$SmallButtonWidth,$ButtonHeight,2)
	$Help = Button(@ScriptDir & "\questionmark.png","","Get help on " & $ProgramName,$Margin+$ListWidth-$SmallButtonWidth-$Margin-$SmallButtonWidth,$PositionY,$SmallButtonWidth,$ButtonHeight,2)
	; program hotkeys
	Local $GUIKeys[31][2] = [["^n",$MenuItemNew],["^o",$MenuItemOpen],["^s",$MenuItemSave],["^z",$MenuItemUndo],["^y",$MenuItemRedo],["^+z",$MenuItemRedo],["^x",$MenuItemCut],["^c",$MenuItemCopy], _
	["^v",$MenuItemPaste],["^t",$MenuItemPasteAfter],["^f",$MenuItemFind],["{F3}",$MenuItemFindNext],["+{F3}",$MenuItemFindPrevious],["^g",$MenuItemFindGoTo],["^{F2}", _
	$MenuItemMark],["{F2}",$MenuItemMarkNext],["+{F2}",$MenuItemMarkPrevious],["{F4}",$MenuItemDisabledNext],["+{F4}",$MenuItemDisabledPrevious], _
	["^=",$Add],["^+=",$Add],["^{NUMPADADD}",$Add],["^{ENTER}",$Change],["^{DEL}",$Delete],["^-",$Delete],["^{NUMPADSUB}",$Delete],["^{UP}",$Up],["^{DOWN}",$Down],["^d",$Duplicate], _
	["^p",$Play],["^r",$Record]]
    GUISetAccelerators($GUIKeys,$GUI)

	GUISetState(@SW_SHOW,$GUI)
	If Not FillProcesses() Then					; does program have enough rights to do its task?
		_Ok("A list of processes (programs) couldn't be made. Perhaps " & $ProgramName & " hasn't enough rights to do so.")
	Else
		ReadFileName()
		If FileExists(FileFailSave()) Then		; check if a program crash fail safe macro file exists
			$MacroFile = FileFailSave()
			OpenMacro(False)
			$MacroFile = StringReplace($MacroFile,"_FailSave_","")
			WriteFileName()
			Changed()
			_Ok("Fail save file for " & StringReplace($MacroFile,"_FailSave_","") & " has been found and opened." & @CRLF & @CRLF & "It seems that " & $ProgramName & " or your computer has crashed.")
		ElseIf StringLen($MacroFile) > 0 And FileExists($MacroFile) Then	; open current macro file
			OpenMacro(False)
		Else
			NewMacro()
		EndIf
		GUICtrlSetState($Macros,$GUI_FOCUS)
		; interface loop
		While Not $CloseByMacro
			$aMsg = GUIGetMsg(1)
			If _WindowChanged($GUI,$GUIPosition,True) Then	; check and update for window repositioning and resizing
				$GUIPosition[2] = _Maximum($GUIPosition[2],$GUIWidth)
				$GUIPosition[3] = _Maximum($GUIPosition[3],$GUIHeight)
				; first set window for correct resizing
				WinMove($GUI,"",$GUIPosition[0],$GUIPosition[1],$GUIPosition[2]+_WindowBordersWidth(),$GUIPosition[3]+_WindowBordersHeight())
				WriteMainWindowPosition()
				GUICtrlSetPos($Processes,$Margin+$ProcessLabelWidth,$Margin-4,$GUIPosition[2]-2*$Margin-$ProcessLabelWidth-$ViewButtonMargin-$ViewButtonWidth)
				$ListWidth = $GUIPosition[2]-$ControlsWidth
				$ListHeight = $GUIPosition[3]-$ControlsHeight
				$PositionY = $Margin+1.7*$ControlHeight+$ListHeight
				GUICtrlSetPos($MacroStepsLabel,$Margin,$Margin+$ControlHeight,$ListWidth-5*$ViewButtonWidth-7*$ViewButtonMargin)
				GUICtrlSetPos($Macros,$Margin,$Margin+1.7*$ControlHeight,$ListWidth,$ListHeight)
				_GUICtrlListView_SetColumnWidth($Macros,0,$ListWidth-_WinAPI_GetSystemMetrics($SM_CXVSCROLL)-2*_WinAPI_GetSystemMetrics($SM_CXEDGE))
				GUICtrlSetPos($Play,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-$Margin,$ButtonWidth,$ButtonHeight)
				GUICtrlSetPos($LoopCount,$Margin+$ListWidth+$Margin+40,$PositionY-$ButtonHeight-$ControlHeight-$Margin-3,$ButtonWidth-50,$InputHeight)
				GUICtrlSetPos($LoopCountLabel,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-$ControlHeight-$Margin)
				GUICtrlSetPos($StepDelayAlways,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-2.1*$ControlHeight-$Margin)
				GUICtrlSetPos($StepDelayAlwaysLabel,$Margin+$ListWidth+$Margin+17,$PositionY-$ButtonHeight-2.1*$ControlHeight-$Margin)
				GUICtrlSetPos($StepDelay,$Margin+$ListWidth+$Margin+40,$PositionY-$ButtonHeight-2.8*$ControlHeight-$Margin-3,$ButtonWidth-50,$InputHeight)
				GUICtrlSetPos($StepDelayLabel,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-2.8*$ControlHeight-$Margin)
				GUICtrlSetPos($ShowBar,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-4*$ControlHeight-$Margin)
				GUICtrlSetPos($ShowBarLabel,$Margin+$ListWidth+$Margin+17,$PositionY-$ButtonHeight-4*$ControlHeight-$Margin)
				GUICtrlSetPos($ShowMain,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-5*$ControlHeight-$Margin,$ButtonWidth,$ButtonHeight)
				GUICtrlSetPos($ShowMainLabel,$Margin+$ListWidth+$Margin+17,$PositionY-$ButtonHeight-5*$ControlHeight-$Margin,$ButtonWidth,$ButtonHeight)
				GUICtrlSetPos($TestMacro,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-6*$ControlHeight-$Margin)
				GUICtrlSetPos($TestMacroLabel,$Margin+$ListWidth+$Margin+17,$PositionY-$ButtonHeight-6*$ControlHeight-$Margin)
				GUICtrlSetPos($Record,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-7*$ControlHeight-$Margin,$ButtonWidth,$ButtonHeight)
				GUICtrlSetPos($MoveTime,$Margin+$ListWidth+$Margin+40,$PositionY-$ButtonHeight-7.9*$ControlHeight-$Margin-3,$ButtonWidth-50,$InputHeight)
				GUICtrlSetPos($MoveTimeLabel,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-7.9*$ControlHeight-$Margin)
				GUICtrlSetPos($ShowRecording,$Margin+$ListWidth+$Margin,$PositionY-$ButtonHeight-9*$ControlHeight-$Margin)
				GUICtrlSetPos($ShowRecordingLabel,$Margin+$ListWidth+$Margin+17,$PositionY-$ButtonHeight-9*$ControlHeight-$Margin)
			EndIf
			If $ProcessRefresh And (WinActive($GUI) Or WinActive($GUIStep)) Then	; refresh if asked by program
				$ProcessRefresh = False
				FillProcesses()
				FillMacroSteps($MacroStep,True)
			EndIf
			If Not $MacroDetectClick And Not WinActive($GUI) And Not WinActive($GUIStep) And Not WinActive($GUICopy) Then $ProcessRefresh = True
			If $MacroDetectClick Then			; show tooltip for detecting a click during macro step editing
				$Tooltip = DetectCalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],_GUICtrlComboBox_GetCurSel($MacroMousePosition)) & ", " & DetectCalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],_GUICtrlComboBox_GetCurSel($MacroMousePosition))
				Switch _GUICtrlComboBox_GetCurSel($MacroStepType)
					Case $StepTypeMouse
						ToolTip($Tooltip & @CRLF & "Click on the desired mouse position")
					Case $StepTypeWindow
						ToolTip($Tooltip & @CRLF & "Click on the desired position for detection")
					Case $StepTypePixel
						ToolTip($Tooltip & @CRLF & "Click on the desired pixel area position")
					Case $StepTypeCapture
						If _GUICtrlComboBox_GetCurSel($MacroCaptureAction) = 1 Then
							$Tooltip = MouseGetPos(0) & ", " & MouseGetPos(1)
						Else
							$Tooltip = (MouseGetPos(0)-_WindowGetX($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)) & ", " & (MouseGetPos(1)-_WindowGetY($MacroSteps[$MacroStep][$MacroStepWindowColumn],True))
						EndIf
						ToolTip($Tooltip & @CRLF & "Click on the desired top left corner of the area")
				EndSwitch
			Else
				ToolTip("")
			EndIf
			If $GUIShowAreaShow > 0 And _Timer_Diff($GUIShowAreaTimer) > 3000 Then	; clear show area
				GUIDelete($GUIShowArea)
				If $GUIShowAreaShow = 1 Then
					; get macro step activated again
					WinSetOnTop($GUIStep,"",$WINDOWS_ONTOP)
					WinActivate($GUIStep)
					WinSetState($GUIStep,"",@SW_SHOW)
					GUISetState(@SW_SHOW,$GUIStep)
					WinSetOnTop($GUIStep,"",$WINDOWS_NOONTOP)
				EndIf
				$GUIShowAreaShow = 0
			EndIf
			If $MacroDetectClick And WinActive($GUIStep) Then
				$MacroDetectClick = False
				ToolTip("")
			ElseIf $MacroDetectClick And _IsPressed($MacroMouseButtons[_GUICtrlComboBox_GetCurSel($MacroMouseButton)][1],$User32DLL) Then	; detect mouse click during step editing and store
				If _GUICtrlComboBox_GetCurSel($MacroStepType) = $StepTypeCapture Then
					If _GUICtrlComboBox_GetCurSel($MacroCaptureAction) = 1 Then
						GUICtrlSetData($MacroCaptureScreenX,MouseGetPos(0))
						GUICtrlSetData($MacroCaptureScreenY,MouseGetPos(1))
					Else
						GUICtrlSetData($MacroCaptureWindowX,MouseGetPos(0)-_WindowGetX($MacroSteps[$MacroStep][$MacroStepWindowColumn],True))
						GUICtrlSetData($MacroCaptureWindowY,MouseGetPos(1)-_WindowGetY($MacroSteps[$MacroStep][$MacroStepWindowColumn],True))
					EndIf
				Else
					GUICtrlSetData($MacroPositionX,DetectCalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],_GUICtrlComboBox_GetCurSel($MacroMousePosition)))
					GUICtrlSetData($MacroPositionY,DetectCalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],_GUICtrlComboBox_GetCurSel($MacroMousePosition)))
				EndIf
				$MacroDetectClick = False
				ToolTip("")
				; get macro step activated again
				WinSetOnTop($GUIStep,"",$WINDOWS_ONTOP)
				WinActivate($GUIStep)
				WinSetState($GUIStep,"",@SW_SHOW)
				GUISetState(@SW_SHOW,$GUIStep)
				WinSetOnTop($GUIStep,"",$WINDOWS_NOONTOP)
				FillMacroSteps($MacroStep)
				$MacroPreviousStep = -1
			EndIf
			$MacroStepSelected = _GUICtrlListView_GetSelectedIndices($Macros)
			$MacroStep = $MacroStepSelected = "" ? -1 : Number($MacroStepSelected)
			$MacroStepCount = _GUICtrlListView_GetItemCount($Macros)
			; set row color of selected step
			If $MacroStep > -1 And $MacroPreviousStep <> $MacroStep Then
				If $MacroSteps[$MacroStep][$MacroStepEnableDisableColumn] = 1 Then
					GUICtrlSetData($MacroStepsLabel,"Macro steps - step " & ($MacroStep+1) & " disabled of " & $MacroStepCount)
				Else
					GUICtrlSetData($MacroStepsLabel,"Macro steps - step " & ($MacroStep+1) & " of " & $MacroStepCount)
				EndIf
				_GUICtrlListView_BeginUpdate($Macros)
				HighlightStep($MacroStep,$MacroPreviousStep)
				_GUICtrlListView_EndUpdate($Macros)
				$MacroPreviousStep = $MacroStep
			EndIf
			; process controls
			If $RecentFilesMenu Then	; check if user wants a recent macro to be opened
				For $File = 0 To UBound($RecentMenuItems)-1
					If $aMsg[0] = $RecentMenuItems[$File] Then
						If Not FileExists(_ShiftGet($RecentFiles,$File)) Then
							_Ok("File " & _ShiftGet($RecentFiles,$File) & " doesn't exist anymore.")
						Else
							$MacroFile = _ShiftGet($RecentFiles,$File)
							OpenMacro(False)
						EndIf
						ExitLoop
					EndIf
				Next
			EndIf
			Switch $aMsg[0]
				Case $MenuItemUndo
					If $Undo > 1 Then
						$Undo -= 1
						OpenMacro(False,_StackGet($Undos,$Undo-1))
						GUICtrlSetState($MenuItemRedo,$GUI_ENABLE)
						If $Undo = 1 Then GUICtrlSetState($MenuItemUndo,$GUI_DISABLE)
					EndIf
				Case $MenuItemRedo
					If $Undo < _StackSize($Undos) Then
						$Undo += 1
						OpenMacro(False,_StackGet($Undos,$Undo-1))
						GUICtrlSetState($MenuItemUndo,$GUI_ENABLE)
						If $Undo = _StackSize($Undos) Then GUICtrlSetState($MenuItemRedo,$GUI_DISABLE)
					EndIf
				Case $MenuItemCut
					If $MacroStep > -1 Then
						$BlockStep = $MacroStep
						If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) > 1 Then
							$BlockStep = FindBlockStep($MacroStep)
							If $BlockStep > -1 Then
								Switch _MessageBox("Cut step or whole block?","",0,"&Step|Cut selected step","&Block|Cut block if selected step is a block type","","",$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
									Case 0
										$BlockStep = -1
									Case 1
										$BlockStep = $MacroStep
									Case 2
										$BlockStep = $MacroStepsBlock[$BlockStep][2]
								EndSwitch
							Else
								$BlockStep = $MacroStep
							EndIf
						EndIf
						If $BlockStep > -1 Then
							CopyMacroSteps($MacroStep,$BlockStep)
							CutMacroSteps($MacroStep,$BlockStep)
							FillMacroSteps($MacroStep)
							GUICtrlSetState($MenuItemPaste,$GUI_ENABLE)
							GUICtrlSetState($MenuItemPasteAfter,$GUI_ENABLE)
							$MacroPreviousStep = -1
						EndIf
					EndIf
				Case $MenuItemCopy
					If $MacroStep > -1 Then
						$BlockStep = $MacroStep
						If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) > 1 Then
							$BlockStep = FindBlockStep($MacroStep)
							If $BlockStep > -1 Then
								Switch _MessageBox("Copy step or whole block?","",0,"&Step|Copy selected step","&Block|Copy block if selected step is a block type","","",$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
									Case 0
										$BlockStep = -1
									Case 1
										$BlockStep = $MacroStep
									Case 2
										$BlockStep = $MacroStepsBlock[$BlockStep][2]
								EndSwitch
							Else
								$BlockStep = $MacroStep
							EndIf
						EndIf
						If $BlockStep > -1 Then
							CopyMacroSteps($MacroStep,$BlockStep)
							GUICtrlSetState($MenuItemPaste,$GUI_ENABLE)
							GUICtrlSetState($MenuItemPasteAfter,$GUI_ENABLE)
						EndIf
					EndIf
				Case $MenuItemCutSteps
					WinSetTitle($GUICopy,"","Cut a range of steps")
					GUICtrlSetData($CopyFromStep,$MacroStep+1)
					GUICtrlSetData($CopyToStep,$MacroStep+1)
					GUICtrlSetState($CopyCut,$GUI_SHOW)
					GUICtrlSetState($CopyCopy,$GUI_HIDE)
					GUISetState(@SW_SHOW,$GUICopy)
				Case $MenuItemCopySteps
					WinSetTitle($GUICopy,"","Copy a range of steps")
					GUICtrlSetData($CopyFromStep,$MacroStep+1)
					GUICtrlSetData($CopyToStep,$MacroStep+1)
					GUICtrlSetState($CopyCopy,$GUI_SHOW)
					GUICtrlSetState($CopyCut,$GUI_HIDE)
					GUISetState(@SW_SHOW,$GUICopy)
				Case $CopyFromStep
					GUICtrlSetData($CopyFromStep,_Minimum(_Maximum(GUICtrlRead($CopyFromStep),1),UBound($MacroSteps)))
				Case $CopyToStep
					GUICtrlSetData($CopyToStep,_Minimum(_Maximum(GUICtrlRead($CopyToStep),1),UBound($MacroSteps)))
				Case $CopyCut
					If $MacroStep > -1 Then
						If GUICtrlRead($CopyFromStep) > GUICtrlRead($CopyToStep) Then
							$CopyStep = GUICtrlRead($CopyToStep)
							GUICtrlSetData($CopyToStep,GUICtrlRead($CopyFromStep))
							GUICtrlSetData($CopyFromStep,$CopyStep)
						EndIf
						GUISetState(@SW_HIDE,$GUICopy)
						CopyMacroSteps(GUICtrlRead($CopyFromStep)-1,GUICtrlRead($CopyToStep)-1)
						CutMacroSteps(GUICtrlRead($CopyFromStep)-1,GUICtrlRead($CopyToStep)-1)
						FillMacroSteps($MacroStep)
						GUICtrlSetState($MenuItemPaste,$GUI_ENABLE)
						GUICtrlSetState($MenuItemPasteAfter,$GUI_ENABLE)
						$MacroPreviousStep = -1
					EndIf
				Case $CopyCopy
					If $MacroStep > -1 Then
						If GUICtrlRead($CopyFromStep) > GUICtrlRead($CopyToStep) Then
							$CopyStep = GUICtrlRead($CopyToStep)
							GUICtrlSetData($CopyToStep,GUICtrlRead($CopyFromStep))
							GUICtrlSetData($CopyFromStep,$CopyStep)
						EndIf
						GUISetState(@SW_HIDE,$GUICopy)
						CopyMacroSteps(GUICtrlRead($CopyFromStep)-1,GUICtrlRead($CopyToStep)-1)
						GUICtrlSetState($MenuItemPaste,$GUI_ENABLE)
						GUICtrlSetState($MenuItemPasteAfter,$GUI_ENABLE)
					EndIf
				Case $MenuItemPaste
					PasteMacroSteps($MacroStep)
					$MacroPreviousStep = -1
				Case $MenuItemPasteAfter
					If $MacroStep > -1 Then
						PasteMacroSteps($MacroStep+1)
					Else
						PasteMacroSteps($MacroStep)
					EndIf
					$MacroPreviousStep = -1
				Case $MenuItemEnableDisableSteps
					GUICtrlSetData($EnableDisableFromStep,$MacroStep+1)
					GUICtrlSetData($EnableDisableToStep,$MacroStep+1)
					GUISetState(@SW_SHOW,$GUIEnableDisable)
				Case $EnableDisableFromStep
					GUICtrlSetData($EnableDisableFromStep,_Minimum(_Maximum(GUICtrlRead($EnableDisableFromStep),1),UBound($MacroSteps)))
				Case $EnableDisableToStep
					GUICtrlSetData($EnableDisableToStep,_Minimum(_Maximum(GUICtrlRead($EnableDisableToStep),1),UBound($MacroSteps)))
				Case $EnableDisableEnable
					EnableDisableMacroSteps(GUICtrlRead($EnableDisableFromStep)-1,GUICtrlRead($EnableDisableToStep)-1,0)
					GUISetState(@SW_HIDE,$GUIEnableDisable)
				Case $EnableDisableDisable
					EnableDisableMacroSteps(GUICtrlRead($EnableDisableFromStep)-1,GUICtrlRead($EnableDisableToStep)-1,1)
					GUISetState(@SW_HIDE,$GUIEnableDisable)
				Case $MenuItemFind
					If $MacroStep > -1 Then _GUICtrlComboBox_SetCurSel($FindStepType,$MacroSteps[$MacroStep][0])
					GUISetState(@SW_SHOW,$GUIFind)
					If _GUICtrlComboBox_GetCurSel($FindType) = 0 Then
						GUICtrlSetState($FindKeyword,$GUI_FOCUS)
					Else
						GUICtrlSetState($FindStepType,$GUI_FOCUS)
					EndIf
				Case $FindType
					If _GUICtrlComboBox_GetCurSel($FindType) = 0 Then
						GUICtrlSetState($FindKeywordLabel,$GUI_SHOW)
						GUICtrlSetState($FindKeyword,$GUI_SHOW)
						GUICtrlSetState($FindStepTypeLabel,$GUI_HIDE)
						GUICtrlSetState($FindStepType,$GUI_HIDE)
					Else
						GUICtrlSetState($FindStepTypeLabel,$GUI_SHOW)
						GUICtrlSetState($FindStepType,$GUI_SHOW)
						GUICtrlSetState($FindKeywordLabel,$GUI_HIDE)
						GUICtrlSetState($FindKeyword,$GUI_HIDE)
					EndIf
				Case $FindFind
					If _GUICtrlComboBox_GetCurSel($FindType) = 0 Then
						GUISetState(@SW_HIDE,$GUIFind)
						GUISetState(@SW_SHOW,$GUI)
						If $MacroStep > -1 And StringLen(GUICtrlRead($FindKeyword)) > 0 Then
							GUICtrlSetState($MenuItemFindNext,$GUI_ENABLE)
							GUICtrlSetState($MenuItemFindPrevious,$GUI_ENABLE)
							$FindStep = _GUICtrlListView_FindInText($Macros,GUICtrlRead($FindKeyword),$MacroStep)
							If $FindStep = -1 Then
								_Ok("A step with the keyword " & GUICtrlRead($FindKeyword) & " in its description wasn't found.")
							Else
								_GUICtrlListView_SetItemFocused($Macros,$MacroStep,False)
								$MacroStepPrevious = $MacroStep
								$MacroStep = $FindStep
								_GUICtrlListView_SetItemSelected($Macros,$MacroStep)
								_GUICtrlListView_EnsureVisible($Macros,$MacroStep)
							EndIf
						EndIf
					Else
						GUISetState(@SW_HIDE,$GUIFind)
						GUISetState(@SW_SHOW,$GUI)
						If $MacroStep > -1 Then
							GUICtrlSetState($MenuItemFindNext,$GUI_ENABLE)
							GUICtrlSetState($MenuItemFindPrevious,$GUI_ENABLE)
							FindByType(_GUICtrlComboBox_GetCurSel($FindStepType))
						EndIf
					EndIf
				Case $MenuItemFindNext
					If _GUICtrlComboBox_GetCurSel($FindType) = 0 Then
						$FindStep = _GUICtrlListView_FindInText($Macros,GUICtrlRead($FindKeyword),$MacroStep)
						If $FindStep = -1 Then
							_Ok("A next step with the keyword " & GUICtrlRead($FindKeyword) & " in its description wasn't found.")
						Else
							_GUICtrlListView_SetItemFocused($Macros,$MacroStep,False)
							$MacroStepPrevious = $MacroStep
							$MacroStep = $FindStep
							_GUICtrlListView_SetItemSelected($Macros,$MacroStep)
							_GUICtrlListView_EnsureVisible($Macros,$MacroStep)
						EndIf
					Else
						FindByType(_GUICtrlComboBox_GetCurSel($FindStepType))
					EndIf
				Case $MenuItemFindPrevious
					If _GUICtrlComboBox_GetCurSel($FindType) = 0 Then
						$FindStep = _GUICtrlListView_FindInText($Macros,GUICtrlRead($FindKeyword),$MacroStep,True,True)
						If $FindStep = -1 Then
							_Ok("A previous step with the keyword " & GUICtrlRead($FindKeyword) & " in its description wasn't found.")
						Else
							_GUICtrlListView_SetItemFocused($Macros,$MacroStep,False)
							$MacroStepPrevious = $MacroStep
							$MacroStep = $FindStep
							_GUICtrlListView_SetItemSelected($Macros,$MacroStep)
							_GUICtrlListView_EnsureVisible($Macros,$MacroStep)
						EndIf
					Else
						FindByType(_GUICtrlComboBox_GetCurSel($FindStepType),True)
					EndIf
				Case $MenuItemFindGoTo
					GUISetState(@SW_SHOW,$GUIGoTo)
					GUICtrlSetState($FindGoTo,$GUI_FOCUS)
				Case $FindGoToJump
					GUICtrlSetData($FindGoTo,_Maximum(_Minimum(GUICtrlRead($FindGoTo),UBound($MacroSteps)),1))
					GUISetState(@SW_HIDE,$GUIGoTo)
					GUISetState(@SW_SHOW,$GUI)
					If $MacroStep > -1 Then SelectStep(GUICtrlRead($FindGoTo)-1)
				Case $MenuItemMark
					If $MacroStep > -1 Then
						$MacroSteps[$MacroStep][$MacroStepMarkColumn] = $MacroSteps[$MacroStep][$MacroStepMarkColumn] = 0 ? 1 : 0
						FillMacroSteps($MacroStep)
						$MacroPreviousStep = -1	; force selected item refresh
						Changed()
					EndIf
				Case $MenuItemMarkNext
					FindMarker(False,$MacroStepMarkColumn,"You haven't set a bookmark yet.")
				Case $MenuItemMarkPrevious
					FindMarker(True,$MacroStepMarkColumn,"You haven't set a bookmark yet.")
				Case $MenuItemDisabledNext
					FindMarker(False,$MacroStepEnableDisableColumn,"There aren't any disabled steps to go to.")
				Case $MenuItemDisabledPrevious
					FindMarker(True,$MacroStepEnableDisableColumn,"There aren't any disabled steps to go to.")
				Case $MenuItemHelp
					Help($HelpFile)
				Case $MenuItemSite
				Case $MenuItemAbout
					_Ok($ProgramName & " is a macro builder to automate your computer." & @CRLF & @CRLF & "Version " & $Version & " released " & $ReleaseDate & @CRLF & @CRLF & "Copyright by Peter Verbeek")
				Case $Processes
					; remember current process and window
					$CurrentProcess = $ProcessList[SelectedProcessIndex()][3]
					$CurrentWindow = $ProcessList[SelectedProcessIndex()][2]
					FillMacroSteps($MacroStep)
					ShowButtonsToolTip()
					Changed()
				Case $WindowsVisible,$WindowsInvisible
					$ShowInvisible = Not $ShowInvisible
					If $ShowInvisible Then
						GUICtrlSetState($WindowsInvisible,$GUI_HIDE)
						GUICtrlSetState($WindowsVisible,$GUI_SHOW)
					Else
						GUICtrlSetState($WindowsInvisible,$GUI_SHOW)
						GUICtrlSetState($WindowsVisible,$GUI_HIDE)
					EndIf
					WriteSettings()
					FillProcesses()
				Case $StepViewGrid
					SetStepView(1)
				Case $StepViewReport
					SetStepView(2)
				Case $StepViewList
					SetStepView(3)
				Case $StepViewInverseColor
					$ViewInvertColors = Not $ViewInvertColors
					If $ViewInvertColors Then
						GUICtrlSetTip($StepViewInverseColor,"Set colors to normal")
					Else
						GUICtrlSetTip($StepViewInverseColor,"Invert colors")
					EndIf
					WriteSettings()
					If $MacroStep > -1 Then
						FillMacroSteps($MacroStep)
						$MacroPreviousStep = -1
					EndIf
				Case $StepViewNumbers
					$ViewNumbers = Not $ViewNumbers
					If $ViewNumbers Then
						GUICtrlSetTip($StepViewNumbers,"Hide the step line numbers")
					Else
						GUICtrlSetTip($StepViewNumbers,"Show the step line numbers")
					EndIf
					WriteSettings()
					If $MacroStep > -1 Then
						FillMacroSteps($MacroStep)
						$MacroPreviousStep = -1
					EndIf
				Case $Add
					$AddingStep = True
					InitializeMacroStepControls()
					ShowHideMacroStepControls()
					GUICtrlSetTip($MacroStepType,$MacroStepTypes[_GUICtrlComboBox_GetCurSel($MacroStepType)][1])
					GUISetState(@SW_HIDE,$GUI)
					WinSetTitle($GUIStep,"","Add macro step")
					GUISetState(@SW_SHOW,$GUIStep)
					SetF1($GUIStep,$HelpFile,$MacroStepTypes[_GUICtrlComboBox_GetCurSel($MacroStepType)][4])
				Case $Change
					$AddingStep = False
					InitializeMacroStepControls()
					_GUICtrlComboBox_SetCurSel($MacroStepType,$MacroSteps[$MacroStep][0])
					GUICtrlSetTip($MacroStepType,$MacroStepTypes[$MacroSteps[$MacroStep][0]][1])
					GUICtrlSetData($MacroDescription,$MacroSteps[$MacroStep][1])
					; fill appropriate step controls with stored macro step data
					Switch $MacroSteps[$MacroStep][0]
						Case $StepTypeMouse
							_GUICtrlComboBox_SetCurSel($MacroMouseButton,$MacroSteps[$MacroStep][2])
							GUICtrlSetData($MacroPositionX,$MacroSteps[$MacroStep][3])
							GUICtrlSetData($MacroPositionY,$MacroSteps[$MacroStep][4])
							_GUICtrlComboBox_SetCurSel($MacroMousePosition,$MacroSteps[$MacroStep][5])
							GUICtrlSetData($MacroPositionXRandom,$MacroSteps[$MacroStep][6])
							GUICtrlSetData($MacroPositionYRandom,$MacroSteps[$MacroStep][7])
							If $MacroSteps[$MacroStep][8] = 1 Then GUICtrlSetState($MacroMouseMoveInstantaneous,$GUI_CHECKED)
							GUICtrlSetData($MacroMouseClicks,$MacroSteps[$MacroStep][9])
						Case $StepTypeKey
							_GUICtrlComboBox_SetCurSel($MacroKeyMode,$MacroSteps[$MacroStep][2])
							_GUICtrlComboBox_SetCurSel($MacroKey,$MacroSteps[$MacroStep][3])
							_GUICtrlComboBox_SetCurSel($MacroTextSend,Number($MacroSteps[$MacroStep][5]))
							GUICtrlSetData($MacroText,$MacroSteps[$MacroStep][4])
						Case $StepTypeDelay
							_GUICtrlComboBox_SetCurSel($MacroDelayType,$MacroSteps[$MacroStep][2])
							If $MacroSteps[$MacroStep][2] = 0 Then
								GUICtrlSetData($MacroDelay,$MacroSteps[$MacroStep][3])
								GUICtrlSetData($MacroDelayRandom,$MacroSteps[$MacroStep][4])
							Else
								GUICtrlSetData($MacroDelayStep,$MacroSteps[$MacroStep][3])
							EndIf
						Case $StepTypeWindow
							_GUICtrlComboBox_SetCurSel($MacroWindowAction,$MacroSteps[$MacroStep][2])
							Switch $MacroSteps[$MacroStep][2]
								Case 3
									GUICtrlSetData($MacroPositionX,$MacroSteps[$MacroStep][3])
									GUICtrlSetData($MacroPositionY,$MacroSteps[$MacroStep][4])
									_GUICtrlComboBox_SetCurSel($MacroWindowPosition,$MacroSteps[$MacroStep][5])
									GUICtrlSetData($MacroPositionXRandom,$MacroSteps[$MacroStep][6])
									GUICtrlSetData($MacroPositionYRandom,$MacroSteps[$MacroStep][7])
									If $MacroSteps[$MacroStep][8] = 1 Then GUICtrlSetState($MacroWindowMoveInstantaneous,$GUI_CHECKED)
								Case 4
									GUICtrlSetData($MacroWidth,$MacroSteps[$MacroStep][3])
									GUICtrlSetData($MacroHeight,$MacroSteps[$MacroStep][4])
									_GUICtrlComboBox_SetCurSel($MacroWindowSize,$MacroSteps[$MacroStep][5])
									GUICtrlSetData($MacroPositionXRandom,$MacroSteps[$MacroStep][6])
									GUICtrlSetData($MacroPositionYRandom,$MacroSteps[$MacroStep][7])
									If $MacroSteps[$MacroStep][8] = 1 Then GUICtrlSetState($MacroWindowResizeInstantaneous,$GUI_CHECKED)
								Case 13,14,15,16
									GUICtrlSetData($MacroWindowWindow,$MacroSteps[$MacroStep][3])
							EndSwitch
						Case $StepTypeRandom
							GUICtrlSetData($MacroRandomMouse,$MacroSteps[$MacroStep][3])
							GUICtrlSetData($MacroRandomDelay,$MacroSteps[$MacroStep][4])
						Case $StepTypeJump
							_GUICtrlComboBox_SetCurSel($MacroJumpType,$MacroSteps[$MacroStep][2])
							FillJumpLabels($MacroJumpLabel,$MacroSteps[$MacroStep][3])
						Case $StepTypeRepeat
							_GUICtrlComboBox_SetCurSel($MacroRepeatType,$MacroSteps[$MacroStep][2])
							GUICtrlSetData($MacroRepeatCount,$MacroSteps[$MacroStep][3])
							GUICtrlSetData($MacroRepeatValue1,$MacroSteps[$MacroStep][3])
							GUICtrlSetData($MacroRepeatValue2,$MacroSteps[$MacroStep][4])
							If _GUICtrlComboBox_GetCurSel($MacroRepeatType) > 2 Then
								GUICtrlSetTip($MacroRepeatValue1,"Greater or equal than value")
							ElseIf _GUICtrlComboBox_GetCurSel($MacroRepeatType) = 1 Then
								GUICtrlSetTip($MacroRepeatValue1,"Not equal to value")
							Else
								GUICtrlSetTip($MacroRepeatValue1,"Equal to value")
							EndIf
						Case $StepTypeTime
							GUICtrlSetData($MacroTimeLoop,$MacroSteps[$MacroStep][3])
						Case $StepTypeInput
							_GUICtrlComboBox_SetCurSel($MacroInputType,$MacroSteps[$MacroStep][2])
							Switch $MacroSteps[$MacroStep][2]
								Case 0
									_GUICtrlComboBox_SetCurSel($MacroContinueKey,$MacroSteps[$MacroStep][3])
							EndSwitch
						Case $StepTypePixel
							_GUICtrlComboBox_SetCurSel($MacroPixelType,$MacroSteps[$MacroStep][2])
							GUICtrlSetData($MacroPositionX,$MacroSteps[$MacroStep][3])
							GUICtrlSetData($MacroPositionY,$MacroSteps[$MacroStep][4])
							_GUICtrlComboBox_SetCurSel($MacroMousePosition,$MacroSteps[$MacroStep][5])
							GUICtrlSetData($MacroPixelWidth,$MacroSteps[$MacroStep][6])
							GUICtrlSetData($MacroPixelHeight,$MacroSteps[$MacroStep][7])
							GUICtrlSetData($MacroPixelId,$MacroSteps[$MacroStep][8])
						Case $StepTypeAlarm
							_GUICtrlComboBox_SetCurSel($MacroAlarmType,$MacroSteps[$MacroStep][2])
							Switch $MacroSteps[$MacroStep][2]
								Case 0	; specific date and time
									SetControlTime($MacroAlarmTime,$MacroSteps[$MacroStep][3])
									SetControlDate($MacroAlarmDate,$MacroSteps[$MacroStep][4])
									_GUICtrlComboBox_SetCurSel($MacroAlarmSpecific,$MacroSteps[$MacroStep][5])
								Case 1	; minute
									SetControlTime($MacroAlarmTime)
									SetControlDate($MacroAlarmDate)
								Case 2	; hour
									SetControlTime($MacroAlarmTime,$MacroSteps[$MacroStep][3])
									SetControlDate($MacroAlarmDate)
								Case 3	; daily
									SetControlTime($MacroAlarmTime,$MacroSteps[$MacroStep][3])
									SetControlDate($MacroAlarmDate)
							EndSwitch
						Case $StepTypeSound
							_GUICtrlComboBox_SetCurSel($MacroSoundType,$MacroSteps[$MacroStep][2])
							Switch $MacroSteps[$MacroStep][2]
								Case 0
									GUICtrlSetData($MacroSoundFrequency,$MacroSteps[$MacroStep][3])
									GUICtrlSetData($MacroSoundDuration,$MacroSteps[$MacroStep][4])
								Case 1
									GUICtrlSetData($MacroSoundFile,$MacroSteps[$MacroStep][3])
								Case 3
									GUICtrlSetData($MacroSoundVolume,$MacroSteps[$MacroStep][3])
							EndSwitch
						Case $StepTypeCounter
							_GUICtrlComboBox_SetCurSel($MacroCounterType,$MacroSteps[$MacroStep][2])
							GUICtrlSetData($MacroCounterName,$MacroSteps[$MacroStep][3])
							FillCounterNames($MacroCounterNameList,$MacroSteps[$MacroStep][3])
							GUICtrlSetData($MacroCounterSet,$MacroSteps[$MacroStep][4])
							If $MacroSteps[$MacroStep][5] = 1 Then GUICtrlSetState($MacroCounterRandom,$GUI_CHECKED)
							GUICtrlSetData($MacroCounterValue1,$MacroSteps[$MacroStep][4])
							GUICtrlSetData($MacroCounterValue2,$MacroSteps[$MacroStep][5])
							SetCounterLabelsAndTips()
						Case $StepTypeFile
							_GUICtrlComboBox_SetCurSel($MacroFileType,$MacroSteps[$MacroStep][2])
							Switch $MacroSteps[$MacroStep][2]
								Case 0,1
									GUICtrlSetData($MacroFileFile,$MacroSteps[$MacroStep][3])
								Case 2,3
									GUICtrlSetData($MacroFilePath,$MacroSteps[$MacroStep][3])
								Case 4
									GUICtrlSetData($MacroFileStamp,$MacroSteps[$MacroStep][3])
								Case 5,6
									GUICtrlSetData($MacroFileChange,$MacroSteps[$MacroStep][3])
							EndSwitch
						Case $StepTypeControl
							_GUICtrlComboBox_SetCurSel($MacroControlType,$MacroSteps[$MacroStep][2])
						Case $StepTypeProcess
							_GUICtrlComboBox_SetCurSel($MacroProcessType,$MacroSteps[$MacroStep][2])
							GUICtrlSetData($MacroProcessShowProcessCurrent,$ProcessList[SelectedProcessIndex()][0] & " - " & $ProcessList[SelectedProcessIndex()][1])
							Switch $MacroSteps[$MacroStep][2]
								Case 0
									GUICtrlSetData($MacroProcessShowProcess,$MacroSteps[$MacroStep][5] & " - " & $MacroSteps[$MacroStep][6])
									_GUICtrlComboBox_SetCurSel($MacroProcessWindowType,$MacroSteps[$MacroStep][3])
									GUICtrlSetData($MacroProcessWindow,$MacroSteps[$MacroStep][4])
								Case 1
									GUICtrlSetData($MacroProcessProgram,$MacroSteps[$MacroStep][3])
									GUICtrlSetData($MacroProcessWindow,$MacroSteps[$MacroStep][4])
									GUICtrlSetData($MacroProcessWorkDir,$MacroSteps[$MacroStep][5])
									_GUICtrlComboBox_SetCurSel($MacroProcessWindowState,$MacroSteps[$MacroStep][6])
									If $MacroSteps[$MacroStep][7] = 1 Then GUICtrlSetState($MacroProcessWait,$GUI_CHECKED)
									GUICtrlSetData($MacroProcessTimeout,$MacroSteps[$MacroStep][8])
								Case 2,3
									GUICtrlSetData($MacroProcessProcess,$MacroSteps[$MacroStep][3])
							EndSwitch
						Case $StepTypeCapture
							_GUICtrlComboBox_SetCurSel($MacroCaptureAction,$MacroSteps[$MacroStep][2])
							Switch $MacroSteps[$MacroStep][2]
								Case 1
									GUICtrlSetData($MacroCaptureScreenX,$MacroSteps[$MacroStep][3])
									GUICtrlSetData($MacroCaptureScreenY,$MacroSteps[$MacroStep][4])
									GUICtrlSetData($MacroCaptureScreenWidth,$MacroSteps[$MacroStep][5])
									GUICtrlSetData($MacroCaptureScreenHeight,$MacroSteps[$MacroStep][6])
								Case 3
									GUICtrlSetData($MacroCaptureWindowX,$MacroSteps[$MacroStep][3])
									GUICtrlSetData($MacroCaptureWindowY,$MacroSteps[$MacroStep][4])
									GUICtrlSetData($MacroCaptureWindowWidth,$MacroSteps[$MacroStep][5])
									GUICtrlSetData($MacroCaptureWindowHeight,$MacroSteps[$MacroStep][6])
								Case 5
									GUICtrlSetData($MacroCapturePath,$MacroSteps[$MacroStep][3])
									If $MacroSteps[$MacroStep][4] = 1 Then
										GUICtrlSetState($MacroCapturePointer,$GUI_CHECKED)
									Else
										GUICtrlSetState($MacroCapturePointer,$GUI_UNCHECKED)
									EndIf
							EndSwitch
						Case $StepTypeShutdown
							_GUICtrlComboBox_SetCurSel($MacroShutDownAction,$MacroSteps[$MacroStep][2])
							If $MacroSteps[$MacroStep][3] = 1 Then GUICtrlSetState($MacroShutDownAsk,$GUI_CHECKED)
						Case $StepTypeComment
							_GUICtrlComboBox_SetCurSel($MacroCommentType,$MacroSteps[$MacroStep][2])
					EndSwitch
					ShowHideMacroStepControls()	; show and hide appropriate macro step controls
					GUISetState(@SW_HIDE,$GUI)
					WinSetTitle($GUIStep,"","Change macro step")
					GUISetState(@SW_SHOW,$GUIStep)
					SetF1($GUIStep,$HelpFile,$MacroStepTypes[_GUICtrlComboBox_GetCurSel($MacroStepType)][4])
					GUICtrlSetState($MacroDescription,$GUI_FOCUS)
				Case $Duplicate
					DuplicateMacroStep()
				Case $Delete
					DeleteMacroStep()
				Case $Up
					UpMacroStep()
				Case $Down
					DownMacroStep()
				Case $Enable
					$MacroSteps[$MacroStep][$MacroStepEnableDisableColumn] = 0
					If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) > 1 Then
						$BlockStep = FindBlockStep($MacroStep)
						If $BlockStep > -1 Then
							Switch _MessageBox("Enable step or whole block?","",0,"&Step|Enable selected step","&Block|Enable block if selected step is a block type","","",$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
								Case 0	; escaped
									$MacroSteps[$MacroStep][$MacroStepEnableDisableColumn] = 1	; reset
								Case 1
									$MacroSteps[$MacroStepsBlock[$BlockStep][2]][$MacroStepEnableDisableColumn] = 0
								Case 2
									For $Step = $MacroStep To $MacroStepsBlock[$BlockStep][2]
										$MacroSteps[$Step][$MacroStepEnableDisableColumn] = 0
									Next
							EndSwitch
							FillMacroSteps($MacroStep)
						EndIf
					EndIf
					$MacroPreviousStep = -1	; force selected item refresh
					Changed()
				Case $Disable
					$MacroSteps[$MacroStep][$MacroStepEnableDisableColumn] = 1
					If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) > 1 Then
						$BlockStep = FindBlockStep($MacroStep)
						If $BlockStep > -1 Then
							Switch _MessageBox("Disable step or whole block?","",0,"&Step|Disable selected step","&Block|Disable block if selected step is a block type","","",$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
								Case 0	; escaped
									$MacroSteps[$MacroStep][$MacroStepEnableDisableColumn] = 0	; reset
								Case 1
									$MacroSteps[$MacroStepsBlock[$BlockStep][2]][$MacroStepEnableDisableColumn] = 1
								Case 2
									For $Step = $MacroStep To $MacroStepsBlock[$BlockStep][2]
										$MacroSteps[$Step][$MacroStepEnableDisableColumn] = 1
									Next
							EndSwitch
							FillMacroSteps($MacroStep)
						EndIf
					EndIf
					$MacroPreviousStep = -1	; force selected item refresh
					Changed()
				Case $MoveTime
					If _Between(GUICtrlRead($MoveTime),1,50) And Not $HideMoveTimeWarning Then
						$HideMoveTimeWarning = HideMessage("A small interval time for mouse movement recording will result in many macro steps.","Don't show this warning again")
						WriteSettings()
					EndIf
					Changed()
				Case $ShowRecording,$ShowRecordingLabel
					If $aMsg[0] = $ShowRecordingLabel Then GUICtrlSetState($ShowRecording,GUICtrlRead($ShowRecording) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
					If GUICtrlRead($ShowRecording) = $GUI_CHECKED And Not $HideShowRecordWarning Then
						$HideShowRecordWarning = HideMessage("Mouse clicks and key presses can go undetected when the updating of the macro step list takes too long on a slow computer.","Don't show this warning again")
						WriteSettings()
					EndIf
					Changed()
				Case $Record
					If Not $HideRecordMessage Then
						$HideRecordMessage = HideMessage("Before recording select the right process from the processes list." & @CRLF & @CRLF & "And don't hold a key down too long. " & $ProgramName & " can't detect key repetition." & @CRLF & @CRLF & "To stop recording press " & $StartStopKeys[$StartStopKey][0] & ".","Don't show this message again")
						WriteSettings()
					EndIf
					RecordMacro()
				Case $LoopCount
					If Not $HideLoopWarning And (GUICtrlRead($LoopCount) = 0 Or GUICtrlRead($LoopCount) > 1) Then
						$HideLoopWarning = HideMessage("Be aware:" & @CRLF & @CRLF & "Playing a macro in a loop can be dangerous.","Don't show this warning again",True)
						WriteSettings()
					EndIf
					SetPlayTip()
					Changed()
				Case $StepDelay
					GUICtrlSetData($StepDelay,_Maximum(GUICtrlRead($StepDelay),1))
					If Not $HideStepDelayWarning And GUICtrlRead($StepDelay) < 10 Then
						$HideStepDelayWarning = HideMessage("Be aware:" & @CRLF & @CRLF & "A small delay, especialy lower than 10 milliseconds, may lead to macro playing problems.","Don't show this warning again",True)
						WriteSettings()
					EndIf
					SetPlayTip()
					Changed()
				Case $StepDelayAlways,$StepDelayAlwaysLabel
					If $aMsg[0] = $StepDelayAlwaysLabel Then GUICtrlSetState($StepDelayAlways,GUICtrlRead($StepDelayAlways) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
					Changed()
				Case $ShowMain
					Changed()
				Case $ShowBar,$ShowBarLabel
					If $aMsg[0] = $ShowBarLabel Then GUICtrlSetState($ShowBar,GUICtrlRead($ShowBar) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
					Changed()
				Case $TestMacro,$TestMacroLabel
					If $aMsg[0] = $TestMacroLabel Then GUICtrlSetState($TestMacro,GUICtrlRead($TestMacro) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
					Changed()
					SetPlayTip()
				Case $Play
					If $MacroStep = -1 Then
						_Ok("Add at least one macro step.")
					ElseIf _GUICtrlComboBox_GetCurSel($Processes) < 0 Then
						_Ok("Select a process to play macro on.")
					Else
						If Number(StringReplace($MacroFileVersion,".","")) <= 1020 And Not $HideEarlierVersionWarning Then
							$HideEarlierVersionWarning = HideMessage("Macros made with version 1.0.2.0 or earlier should be checked as they may fail." & @CRLF & @CRLF & "Current version runs a lot faster and the pixel checking has been changed too.","Don't show this warning again",True)
							WriteSettings()
						EndIf
						If Not $HidePlayMessage Then _Message("To stop the macro press " & $StartStopKeys[$StartStopKey][0] & ".",$HidePlayMessage,$ProgramName & " - Information","Don't show this info again","Ok",@ScriptDir & "\info.ico",0,0,0,$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
						WriteSettings()
						PlayMacro()
					EndIf
				Case $New,$MenuItemNew
					If Not $Changed Or _Sure("You've made changes but you haven't saved your macro." & @CRLF & @CRLF & "Create a new one?") Then NewMacro()
				Case $Open,$MenuItemOpen
					If Not $Changed Or _Sure("You've made changes but you haven't saved your macro." & @CRLF & @CRLF & "Open a macro?") Then OpenMacro()
				Case $Save,$MenuItemSave
					SaveMacro()
					FillProcesses()
					FillMacroSteps($MacroStep)
					$MacroPreviousStep = -1
				Case $SaveAs,$MenuItemSaveAs
					SaveMacro(True)
					FillProcesses()
					FillMacroSteps($MacroStep)
					$MacroPreviousStep = -1
				Case $MacroWidth,$MacroHeight,$MacroWindowSize
					If _GUICtrlComboBox_GetCurSel($MacroStepType) = $StepTypeWindow And _GUICtrlComboBox_GetCurSel($MacroWindowAction) = 4 And _GUICtrlComboBox_GetCurSel($MacroWindowSize) = 0 Then
						GUICtrlSetData($MacroWidth,_Maximum(1,GUICtrlRead($MacroWidth)))
						GUICtrlSetData($MacroHeight,_Maximum(1,GUICtrlRead($MacroHeight)))
					EndIf
				Case $MacroDetectPosition,$MacroCaptureDetect
					If Not $HideDetectionMessage Then
						$HideDetectionMessage = HideMessage("Be sure that the target process for the detection is selected." & @CRLF & @CRLF & $ProgramName & " will only do the detection on the selected process.","Don't show this warning again")
						WriteSettings()
					EndIf
					$MacroDetectClick = True
					If $MacroStep = -1 Then		; if there aren't any steps add one to have a window for further detection processing
						EditMacroStep(True,$StepTypeComment,"New macro",0)
						$MacroStep = 0
						$MacroStepCount = 1
					EndIf
					WinActivate($MacroSteps[$MacroStep][$MacroStepWindowColumn])
					WinWaitActive($MacroSteps[$MacroStep][$MacroStepWindowColumn])
				Case $Show,$MacroShowPosition,$MacroCaptureShow	; show area where macro step applies
					If Not $HideShowMessage Then
						$HideShowMessage = HideMessage("Be sure that the target process for the showing is selected." & @CRLF & @CRLF & $ProgramName & " will only show for the selected process.","Don't show this warning again")
						WriteSettings()
					EndIf
					If $MacroStep = -1 Then		; if there aren't any steps add one to have a window for further detection processing
						EditMacroStep(True,$StepTypeComment,"New macro",0)
						$MacroStep = 0
						$MacroStepCount = 1
					EndIf
					If $aMsg[0] = $Show Or _
					   ($aMsg[0] = $MacroShowPosition And _InList($MacroSteps[$MacroStep][0],$StepTypeMouse,$StepTypeWindow,$StepTypePixel)) Or _
					   ($aMsg[0] = $MacroCaptureShow And $MacroSteps[$MacroStep][0] = $StepTypeCapture) Then
						WinActivate($MacroSteps[$MacroStep][$MacroStepWindowColumn])
						WinWaitActive($MacroSteps[$MacroStep][$MacroStepWindowColumn])
						If $aMsg[0] <> $Show Then
							Switch _GUICtrlComboBox_GetCurSel($MacroStepType)
								Case $StepTypeMouse
									$MouseX = CalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],-1,GUICtrlRead($MacroPositionX),_GUICtrlComboBox_GetCurSel($MacroMousePosition))
									$MouseY = CalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],-1,GUICtrlRead($MacroPositionY),_GUICtrlComboBox_GetCurSel($MacroMousePosition))
									If $GUIShowAreaShow > 0 Then GUIDelete($GUIShowArea)
									$GUIShowArea = GUICreate("",10,10,$MouseX-5,$MouseY-5,$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
									GUISetBkColor($Color_Red,$GUIShowArea)
									WinSetTrans($GUIShowArea,"",200)
									GUISetState(@SW_SHOW,$GUIShowArea)
									$GUIShowAreaShow = 1
									$GUIShowAreaTimer = _Timer_Init()
									MouseMove($MouseX,$MouseY)
								Case $StepTypeWindow
									If _GUICtrlComboBox_GetCurSel($MacroWindowAction) = 3 Then
										If _GUICtrlComboBox_GetCurSel($MacroWindowPosition) = 0 Then
											MouseMove(GUICtrlRead($MacroPositionX),GUICtrlRead($MacroPositionY))
										Else
											MouseMove(_WindowGetX($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+GUICtrlRead($MacroPositionX),_WindowGetY($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+GUICtrlRead($MacroPositionY))
										EndIf
									EndIf
								Case $StepTypePixel
									If _GUICtrlComboBox_GetCurSel($MacroPixelType) = 0 Then
										If $GUIShowAreaShow > 0 Then GUIDelete($GUIShowArea)
										$GUIShowArea = GUICreate("",GUICtrlRead($MacroPixelWidth),GUICtrlRead($MacroPixelHeight),CalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],-1,GUICtrlRead($MacroPositionX),_GUICtrlComboBox_GetCurSel($MacroMousePosition)),CalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],-1,GUICtrlRead($MacroPositionY),_GUICtrlComboBox_GetCurSel($MacroMousePosition)),$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
										GUISetBkColor($Color_Red,$GUIShowArea)
										WinSetTrans($GUIShowArea,"",50)
										GUISetState(@SW_SHOW,$GUIShowArea)
										$GUIShowAreaShow = 1
										$GUIShowAreaTimer = _Timer_Init()
									ElseIf _GUICtrlComboBox_GetCurSel($MacroPixelType) = 2 Then
										MouseMove(CalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],-1,GUICtrlRead($MacroPositionX),_GUICtrlComboBox_GetCurSel($MacroMousePosition)),CalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],-1,GUICtrlRead($MacroPositionY),_GUICtrlComboBox_GetCurSel($MacroMousePosition)))
									EndIf
								Case $StepTypeCapture
									If _InList(_GUICtrlComboBox_GetCurSel($MacroCaptureAction),1,3) Then
										If $GUIShowAreaShow > 0 Then GUIDelete($GUIShowArea)
										If _GUICtrlComboBox_GetCurSel($MacroCaptureAction) = 1 Then
											$GUIShowArea = GUICreate("",GUICtrlRead($MacroCaptureScreenWidth),GUICtrlRead($MacroCaptureScreenHeight),GUICtrlRead($MacroCaptureScreenX),GUICtrlRead($MacroCaptureScreenY),$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
										Else
											$GUIShowArea = GUICreate("",GUICtrlRead($MacroCaptureWindowWidth),GUICtrlRead($MacroCaptureWindowHeight),_WindowGetX($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+GUICtrlRead($MacroCaptureWindowX),_WindowGetY($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+GUICtrlRead($MacroCaptureWindowY),$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
										EndIf
										GUISetBkColor($Color_Red,$GUIShowArea)
										WinSetTrans($GUIShowArea,"",50)
										GUISetState(@SW_SHOW,$GUIShowArea)
										$GUIShowAreaShow = 1
										$GUIShowAreaTimer = _Timer_Init()
									EndIf
							EndSwitch
						Else
							Switch $MacroSteps[$MacroStep][0]
								Case $StepTypeMouse
									$MouseX = CalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],$MacroStep)
									$MouseY = CalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],$MacroStep)
									If $GUIShowAreaShow > 0 Then GUIDelete($GUIShowArea)
									$GUIShowArea = GUICreate("",10,10,$MouseX-5,$MouseY-5,$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
									GUISetBkColor($Color_Red,$GUIShowArea)
									WinSetTrans($GUIShowArea,"",200)
									GUISetState(@SW_SHOW,$GUIShowArea)
									$GUIShowAreaShow = 2
									$GUIShowAreaTimer = _Timer_Init()
									MouseMove($MouseX,$MouseY)
								Case $StepTypeWindow
									If $MacroSteps[$MacroStep][2] = 2 Then
										If $MacroSteps[$MacroStep][5] = 0 Then
											MouseMove($MacroSteps[$MacroStep][3],$MacroSteps[$MacroStep][4])
										Else
											MouseMove(_WindowGetX($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+$MacroSteps[$MacroStep][3],_WindowGetY($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+$MacroSteps[$MacroStep][4])
										EndIf
									EndIf
								Case $StepTypePixel
									If $MacroSteps[$MacroStep][2] = 0 Then
										If $GUIShowAreaShow > 0 Then GUIDelete($GUIShowArea)
										$GUIShowArea = GUICreate("",$MacroSteps[$MacroStep][6],$MacroSteps[$MacroStep][7],CalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],$MacroStep),CalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],$MacroStep),$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
										GUISetBkColor($Color_Red,$GUIShowArea)
										WinSetTrans($GUIShowArea,"",50)
										GUISetState(@SW_SHOW,$GUIShowArea)
										$GUIShowAreaShow = 2
										$GUIShowAreaTimer = _Timer_Init()
									ElseIf $MacroSteps[$MacroStep][2] = 2 Then
										MouseMove(CalculateX($MacroSteps[$MacroStep][$MacroStepWindowColumn],$MacroStep),CalculateY($MacroSteps[$MacroStep][$MacroStepWindowColumn],$MacroStep))
									EndIf
								Case $StepTypeCapture
									If _InList($MacroSteps[$MacroStep][2],1,3) Then
										If $GUIShowAreaShow > 0 Then GUIDelete($GUIShowArea)
										If $MacroSteps[$MacroStep][2] = 1 Then
											$GUIShowArea = GUICreate("",$MacroSteps[$MacroStep][5],$MacroSteps[$MacroStep][6],$MacroSteps[$MacroStep][3],$MacroSteps[$MacroStep][4],$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
										Else
											$GUIShowArea = GUICreate("",$MacroSteps[$MacroStep][5],$MacroSteps[$MacroStep][6],_WindowGetX($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+$MacroSteps[$MacroStep][3],_WindowGetY($MacroSteps[$MacroStep][$MacroStepWindowColumn],True)+$MacroSteps[$MacroStep][4],$WS_BORDER+$WS_POPUP,$WS_EX_TOPMOST)
										EndIf
										GUISetBkColor($Color_Red,$GUIShowArea)
										WinSetTrans($GUIShowArea,"",50)
										GUISetState(@SW_SHOW,$GUIShowArea)
										$GUIShowAreaShow = 2
										$GUIShowAreaTimer = _Timer_Init()
									EndIf
							EndSwitch
						EndIf
					   EndIf
				Case $StepHelp
					Help($HelpFile,$MacroStepTypes[_GUICtrlComboBox_GetCurSel($MacroStepType)][4])
				Case $MacroStepType,$MacroMouseButton,$MacroKey,$MacroTextSend,$MacroWindowAction,$MacroDelayType,$MacroJumpType,$MacroInputType,$MacroPixelType,$MacroRepeatType,$MacroSoundType,$MacroCounterType,$MacroFileType,$MacroProcessType,$MacroProcessWindowType,$MacroCaptureAction
					GUICtrlSetTip($MacroStepType,$MacroStepTypes[_GUICtrlComboBox_GetCurSel($MacroStepType)][1])	; show tip belonging to step type
					ShowHideMacroStepControls()
					If _GUICtrlComboBox_GetCurSel($MacroRepeatType) > 2 Then
						GUICtrlSetTip($MacroRepeatValue1,"Greater or equal than value")
					ElseIf _GUICtrlComboBox_GetCurSel($MacroRepeatType) = 1 Then
						GUICtrlSetTip($MacroRepeatValue1,"Not equal to value")
					Else
						GUICtrlSetTip($MacroRepeatValue1,"Equal to value")
					EndIf
					SetCounterLabelsAndTips()
					SetF1($GUIStep,$HelpFile,$MacroStepTypes[_GUICtrlComboBox_GetCurSel($MacroStepType)][4])
				Case $MacroMouseMoveInstantaneousLabel
					GUICtrlSetState($MacroMouseMoveInstantaneous,GUICtrlRead($MacroMouseMoveInstantaneous) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $MacroWindowMoveInstantaneousLabel
					GUICtrlSetState($MacroWindowMoveInstantaneous,GUICtrlRead($MacroWindowMoveInstantaneous) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $MacroWindowResizeInstantaneousLabel
					GUICtrlSetState($MacroWindowResizeInstantaneous,GUICtrlRead($MacroWindowResizeInstantaneous) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $MacroAlarmType
					ShowHideMacroStepControls()
					Switch _GUICtrlComboBox_GetCurSel($MacroAlarmType)
						Case 0,3	; specific date and time or daily
							SetControlTime($MacroAlarmTime)
						Case 2	; hourly
							SetControlTime($MacroAlarmTime,0)
					EndSwitch
				Case $MacroAlarmTime
					If _GUICtrlComboBox_GetCurSel($MacroAlarmType) = 2 And GetControlTime($MacroAlarmTime) >= 10000 Then _Ok("Enter time in minutes and seconds, no hour for hourly alarm.")
				Case $MacroSoundSelect
					SelectFile("Select a sound file","Sound wave files (*.wav)|Mp3 files (*.mp3)",$MacroSoundFile)
				Case $MacroSoundVolume
					GUICtrlSetData($MacroSoundVolume,_Clamp(GUICtrlRead($MacroSoundVolume),0,100))
				Case $MacroCounterRandomLabel
					GUICtrlSetState($MacroCounterRandom,GUICtrlRead($MacroCounterRandom) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $MacroFileFileSelect
					SelectFile("Select a file to check","",$MacroFileFile)
				Case $MacroFilePathSelect
					SelectPath("Select a folder to check",$MacroFilePath)
				Case $MacroFileStampSelect
					SelectFile("Select a file to record change of","",$MacroFileStamp)
				Case $MacroFileChangeSelect
					SelectFile("Select a file to check if changed","",$MacroFileChange)
				Case $MacroProcessProgramSelect
					SelectFile("Select a program file","Executables (*.exe)|Batch files (*.bat)|Command files (*.com)",$MacroProcessProgram)
				Case $MacroProcessWorkDirSelect
					SelectPath("Select a working folder of selected program",$MacroProcessWorkDir)
				Case $MacroProcessProcessSelect
					If _GUICtrlComboBox_GetCurSel($MacroProcessType) = 1 Then
						SelectFile("Select a process executable file","Executables (*.exe)",$MacroProcessProgram)
					Else
						SelectFile("Select a process executable file","Executables (*.exe)",$MacroProcessProcess)
					EndIf
				Case $MacroProcessTest
					If StringLen(GUICtrlRead($MacroProcessProgram)) = 0 Then
						_Ok("Enter or select a program.")
					ElseIf Not FileExists(FilePathTranslateVariables(GUICtrlRead($MacroProcessProgram))) Then
						_Ok("Program file " & GUICtrlRead($MacroProcessProgram) & " doesn't exist.")
					Else
						Switch _GUICtrlComboBox_GetCurSel($MacroProcessWindowState)
							Case 0
								Run('"' & FilePathTranslateVariables(GUICtrlRead($MacroProcessProgram)) & '"',FilePathTranslateVariables(GUICtrlRead($MacroProcessWorkDir)),@SW_SHOWNORMAL)
							Case 1
								Run('"' & FilePathTranslateVariables(GUICtrlRead($MacroProcessProgram)) & '"',FilePathTranslateVariables(GUICtrlRead($MacroProcessWorkDir)),@SW_SHOWMAXIMIZED)
							Case 2
								Run('"' & FilePathTranslateVariables(GUICtrlRead($MacroProcessProgram)) & '"',FilePathTranslateVariables(GUICtrlRead($MacroProcessWorkDir)),@SW_SHOWMINIMIZED)
							Case 3
								Run('"' & FilePathTranslateVariables(GUICtrlRead($MacroProcessProgram)) & '"',FilePathTranslateVariables(GUICtrlRead($MacroProcessWorkDir)),@SW_HIDE)
						EndSwitch
					EndIf
				Case $MacroProcessWait,$MacroProcessWaitLabel
					If $aMsg[0] = $MacroProcessWaitLabel Then GUICtrlSetState($MacroProcessWait,GUICtrlRead($MacroProcessWait) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($MacroProcessTimeoutLabel,GUICtrlRead($MacroProcessWait) = $GUI_CHECKED ? $GUI_HIDE : $GUI_SHOW)
					GUICtrlSetState($MacroProcessTimeout,GUICtrlRead($MacroProcessWait) = $GUI_CHECKED ? $GUI_HIDE : $GUI_SHOW)
				Case $MacroProcessTimeout
					If GUICtrlRead($MacroProcessTimeout) < 5000 Then _Ok("A timeout of " & GUICtrlRead($MacroProcessTimeout) & " is small. The program might not be started yet.")
				Case $MacroCapturePathSelect
					SelectPath("Select a folder for the captured images",$MacroCapturePath)
				Case $MacroCapturePointerLabel
					GUICtrlSetState($MacroCapturePointer,GUICtrlRead($MacroCapturePointer) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $MacroShutDownAskLabel
					GUICtrlSetState($MacroShutDownAsk,GUICtrlRead($MacroShutDownAsk) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $MacroSave					; save macro step if everything is okay
					$MacroDetectClick = False
					$MacroSaveStep = True
					Switch _GUICtrlComboBox_GetCurSel($MacroStepType)
						Case $StepTypeMouse
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroMouseButton),GUICtrlRead($MacroPositionX),GUICtrlRead($MacroPositionY),_GUICtrlComboBox_GetCurSel($MacroMousePosition),GUICtrlRead($MacroPositionXRandom),GUICtrlRead($MacroPositionYRandom),GUICtrlRead($MacroMouseMoveInstantaneous) = $GUI_CHECKED ? 1 : 0,_Maximum(1,GUICtrlRead($MacroMouseClicks)))
						Case $StepTypeKey
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroKeyMode),_GUICtrlComboBox_GetCurSel($MacroKey),GUICtrlRead($MacroText),_GUICtrlComboBox_GetCurSel($MacroTextSend))
						Case $StepTypeDelay
							If _GUICtrlComboBox_GetCurSel($MacroDelayType) = 0 Then
								EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroDelayType),GUICtrlRead($MacroDelay),GUICtrlRead($MacroDelayRandom),GUICtrlRead($MacroDelayStep))
							Else
								EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroDelayType),_Maximum(GUICtrlRead($MacroDelayStep),10))
							EndIf
						Case $StepTypeWindow
							Switch _GUICtrlComboBox_GetCurSel($MacroWindowAction)
								Case 3
									If _GUICtrlComboBox_GetCurSel($MacroWindowPosition) = 0 Then
										If GUICtrlRead($MacroPositionX) < 0 Then _Ok("Notice, you've chosen a window x position smaller than 0.")
										If GUICtrlRead($MacroPositionY) < 0 Then _Ok("Notice, you've chosen a window y position smaller than 0.")
										If GUICtrlRead($MacroPositionX) > @DesktopWidth Then _Ok("Notice, you've chosen a window x position greater than the screen width.")
										If GUICtrlRead($MacroPositionY) > @DesktopHeight Then _Ok("Notice, you've chosen a window y position greater than the screen height.")
									EndIf
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroWindowAction),GUICtrlRead($MacroPositionX),GUICtrlRead($MacroPositionY),_GUICtrlComboBox_GetCurSel($MacroWindowPosition),GUICtrlRead($MacroPositionXRandom),GUICtrlRead($MacroPositionYRandom),GUICtrlRead($MacroWindowMoveInstantaneous) = $GUI_CHECKED ? 1 : 0)
								Case 4
									If _GUICtrlComboBox_GetCurSel($MacroWindowSize) = 0 Then
										If GUICtrlRead($MacroWidth) < 20 Then _Ok("Notice, you've chosen a very small window width.")
										If GUICtrlRead($MacroHeight) < 20 Then _Ok("Notice, you've chosen a very small window height.")
										If GUICtrlRead($MacroWidth) > @DesktopWidth Then _Ok("Notice, you've chosen a window width greater than the screen width.")
										If GUICtrlRead($MacroHeight) > @DesktopHeight Then _Ok("Notice, you've chosen a window height greater than the screen height.")
									EndIf
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroWindowAction),GUICtrlRead($MacroWidth),GUICtrlRead($MacroHeight),_GUICtrlComboBox_GetCurSel($MacroWindowSize),GUICtrlRead($MacroPositionXRandom),GUICtrlRead($MacroPositionYRandom),GUICtrlRead($MacroWindowResizeInstantaneous) = $GUI_CHECKED ? 1 : 0)
								Case 13,14,15,16
									If StringLen(GUICtrlRead($MacroWindowWindow)) = 0 Then
										If _GUICtrlComboBox_GetCurSel($MacroWindowAction) = 13 Then
											_Ok("Enter a keyword to match the window you want to switch to.")
										Else
											_Ok("Enter a keyword to match the window you want to check existence of.")
										EndIf
										$MacroSaveStep = False
									Else
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroWindowAction),GUICtrlRead($MacroWindowWindow))
										If _InList(_GUICtrlComboBox_GetCurSel($MacroWindowAction),14,15) Then AddEndStep($AddingStep)	; if window (not) exists
									EndIf
								Case Else
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroWindowAction))
									If _InList(_GUICtrlComboBox_GetCurSel($MacroWindowAction),6,7,8,9) Then AddEndStep($AddingStep)	; if (not) maximized/minimized
							EndSwitch
						Case $StepTypeRandom
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),0,GUICtrlRead($MacroRandomMouse),GUICtrlRead($MacroRandomDelay))
						Case $StepTypeLabel
							_GUICtrlTrimSpaces($MacroDescription)
							If _StringEmpty(GUICtrlRead($MacroDescription)) Then
								_Ok("You haven't entered a label in the description field.")
								$MacroSaveStep = False
							ElseIf MacroStepLabelExists(GUICtrlRead($MacroDescription),Not $AddingStep) Then
								_Ok("There's already a label called "  & GUICtrlRead($MacroDescription) & ".")
								$MacroSaveStep = False
							EndIf
							If $MacroSaveStep Then
								If $MacroStep > -1 Then ChangeJumpLabels($AddingStep,$MacroSteps[$MacroStep][1],GUICtrlRead($MacroDescription))
								EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription))
							EndIf
						Case $StepTypeJump
							If _GUICtrlComboBox_GetCurSel($MacroJumpType) = 0 And _StringEmpty(GUICtrlRead($MacroJumpLabel)) Then
								_Ok("Create a label first so it appears in the list of labels.")
							Else
								EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroJumpType),GUICtrlRead($MacroJumpLabel))
							EndIf
						Case $StepTypeEnd
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription))
						Case $StepTypeRepeat
							Switch _GUICtrlComboBox_GetCurSel($MacroRepeatType)
								Case 0
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroRepeatType),GUICtrlRead($MacroRepeatCount))
								Case 1,2,3	; = and >=
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroRepeatType),GUICtrlRead($MacroRepeatValue1))
								Case 4		; <=
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroRepeatType),0,GUICtrlRead($MacroRepeatValue2))
								Case 5,6	; between and not between
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroRepeatType),GUICtrlRead($MacroRepeatValue1),GUICtrlRead($MacroRepeatValue2))
								Case 7,8	; odd and even
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroRepeatType))
							EndSwitch
							AddEndStep($AddingStep)
						Case $StepTypeTime
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),0,GUICtrlRead($MacroTimeLoop))
							AddEndStep($AddingStep)
						Case $StepTypeInput
							Switch _GUICtrlComboBox_GetCurSel($MacroInputType)
								Case 0
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroInputType),_GUICtrlComboBox_GetCurSel($MacroContinueKey))
								Case 1
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroInputType))
								Case 2
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroInputType))
									AddEndStep($AddingStep)
							EndSwitch
						Case $StepTypePixel
							If Not $AddingStep And _GUICtrlComboBox_GetCurSel($MacroStepType) = $MacroSteps[$MacroStep][0] And GUICtrlRead($MacroPixelId) <> $MacroSteps[$MacroStep][8] Then
								If _InList(_GUICtrlComboBox_GetCurSel($MacroPixelType),0,2) And _Sure("You've changed the id number from " & $MacroSteps[$MacroStep][8] & " to " & GUICtrlRead($MacroPixelId) & "." & @CRLF & @CRLF & "If any, automatically change all accompanying If steps?") Then
									For $ChangeStep = 0 To UBound($MacroSteps)-1
										If _GUICtrlComboBox_GetCurSel($MacroPixelType) = 0 Then
											If $MacroSteps[$ChangeStep][0] = $StepTypePixel And $MacroSteps[$ChangeStep][2] = 1 And $MacroSteps[$ChangeStep][8] = $MacroSteps[$MacroStep][8] Then $MacroSteps[$ChangeStep][8] = GUICtrlRead($MacroPixelId)
										Else
											If $MacroSteps[$ChangeStep][0] = $StepTypePixel And $MacroSteps[$ChangeStep][2] = 3 And $MacroSteps[$ChangeStep][8] = $MacroSteps[$MacroStep][8] Then $MacroSteps[$ChangeStep][8] = GUICtrlRead($MacroPixelId)
										EndIf
									Next
								ElseIf _InList(_GUICtrlComboBox_GetCurSel($MacroPixelType),1,3) And _Sure("You've changed the id number from " & $MacroSteps[$MacroStep][8] & " to " & GUICtrlRead($MacroPixelId) & "." & @CRLF & @CRLF & "If any, automatically change all accompanying Pixel record steps?") Then
									For $ChangeStep = 0 To UBound($MacroSteps)-1
										If _GUICtrlComboBox_GetCurSel($MacroPixelType) = 1 Then
											If $MacroSteps[$ChangeStep][0] = $StepTypePixel And $MacroSteps[$ChangeStep][2] = 0 And $MacroSteps[$ChangeStep][8] = $MacroSteps[$MacroStep][8] Then $MacroSteps[$ChangeStep][8] = GUICtrlRead($MacroPixelId)
										Else
											If $MacroSteps[$ChangeStep][0] = $StepTypePixel And $MacroSteps[$ChangeStep][2] = 2 And $MacroSteps[$ChangeStep][8] = $MacroSteps[$MacroStep][8] Then $MacroSteps[$ChangeStep][8] = GUICtrlRead($MacroPixelId)
										EndIf
									Next
								EndIf
							EndIf
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroPixelType),GUICtrlRead($MacroPositionX),GUICtrlRead($MacroPositionY),_GUICtrlComboBox_GetCurSel($MacroMousePosition),GUICtrlRead($MacroPixelWidth),GUICtrlRead($MacroPixelHeight),GUICtrlRead($MacroPixelId))
							If _InList(_GUICtrlComboBox_GetCurSel($MacroPixelType),1,3) Then AddEndStep($AddingStep)
						Case $StepTypeAlarm
							Switch _GUICtrlComboBox_GetCurSel($MacroAlarmType)
								Case 0	; specific date and time
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroAlarmType),GetControlTime($MacroAlarmTime),GetControlDate($MacroAlarmDate),_GUICtrlComboBox_GetCurSel($MacroAlarmSpecific))
								Case 1 	; minute
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroAlarmType),"","")
								Case 2,3; hour and daily
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroAlarmType),GetControlTime($MacroAlarmTime),"")
							EndSwitch
							AddEndStep($AddingStep)
						Case $StepTypeSound
							Switch _GUICtrlComboBox_GetCurSel($MacroSoundType)
								Case 0
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroSoundType),GUICtrlRead($MacroSoundFrequency),GUICtrlRead($MacroSoundDuration))
								Case 1
									_GUICtrlTrimSpaces($MacroSoundFile)
									If StringLen(GUICtrlRead($MacroSoundFile)) = 0 Then
										_Ok("Enter or select a sound file for playing a sound.")
										$MacroSaveStep = False
									ElseIf Not FileExists(FilePathTranslateVariables((GUICtrlRead($MacroSoundFile)))) Then
										_Ok("File " & GUICtrlRead($MacroSoundFile) & " doesn't exist.")
										$MacroSaveStep = False
									Else
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroSoundType),GUICtrlRead($MacroSoundFile))
									EndIf
								Case 3
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroSoundType),GUICtrlRead($MacroSoundVolume))
								Case Else
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroSoundType))
							EndSwitch
						Case $StepTypeCounter
							_GUICtrlTrimSpaces($MacroCounterName)
							If _Inlist(_GUICtrlComboBox_GetCurSel($MacroCounterType),0,1) And _StringEmpty(GUICtrlRead($MacroCounterName)) Then
								_Ok("You haven't entered a counter name.")
								$MacroSaveStep = False
							ElseIf _GUICtrlComboBox_GetCurSel($MacroCounterType) > 1 And _StringEmpty(GUICtrlRead($MacroCounterNameList)) Then
								_Ok("First create a counter so it appears in the list of counters.")
								$MacroSaveStep = False
							Else
								Switch _GUICtrlComboBox_GetCurSel($MacroCounterType)
									Case 0	; set
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCounterType),GUICtrlRead($MacroCounterName),GUICtrlRead($MacroCounterSet),GUICtrlRead($MacroCounterRandom) = $GUI_CHECKED ? 1 : 0)
									Case 1	; set play
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCounterType),GUICtrlRead($MacroCounterName))
									Case 2	; add
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCounterType),GUICtrlRead($MacroCounterNameList),GUICtrlRead($MacroCounterSet),GUICtrlRead($MacroCounterRandom) = $GUI_CHECKED ? 1 : 0)
									Case 3,4,5	; =, not =,  >=
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCounterType),GUICtrlRead($MacroCounterNameList),GUICtrlRead($MacroCounterValue1))
									Case 6		; <=
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCounterType),GUICtrlRead($MacroCounterNameList),0,GUICtrlRead($MacroCounterValue2))
									Case 7,8	; between and not between
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCounterType),GUICtrlRead($MacroCounterNameList),GUICtrlRead($MacroCounterValue1),GUICtrlRead($MacroCounterValue2))
									Case 9,10	; odd and even
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCounterType),GUICtrlRead($MacroCounterNameList))
								EndSwitch
								If _GUICtrlComboBox_GetCurSel($MacroCounterType) > 2 Then AddEndStep($AddingStep)
							EndIf
						Case $StepTypeFile
							_GUICtrlTrimSpaces($MacroFileFile)
							_GUICtrlTrimSpaces($MacroFilePath)
							_GUICtrlTrimSpaces($MacroFileStamp)
							_GUICtrlTrimSpaces($MacroFileChange)
							Switch _GUICtrlComboBox_GetCurSel($MacroFileType)
								Case 0,1
									If StringLen(GUICtrlRead($MacroFileFile)) = 0 Then
										_Ok("Enter or select a file to check existence of.")
										$MacroSaveStep = False
									Else
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroFileType),GUICtrlRead($MacroFileFile))
										AddEndStep($AddingStep)
									EndIf
								Case 2,3
									If StringLen(GUICtrlRead($MacroFilePath)) = 0 Then
										_Ok("Enter or select a folder to check existence of.")
										$MacroSaveStep = False
									Else
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroFileType),GUICtrlRead($MacroFilePath))
										AddEndStep($AddingStep)
									EndIf
								Case 4
									If StringLen(GUICtrlRead($MacroFileStamp)) = 0 Then
										_Ok("Enter or select a file to record change of.")
										$MacroSaveStep = False
									ElseIf Not FileExists(FilePathTranslateVariables((GUICtrlRead($MacroFileStamp)))) Then
										_Ok("File " & GUICtrlRead($MacroFileStamp) & " doesn't exist.")
										$MacroSaveStep = False
									Else
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroFileType),GUICtrlRead($MacroFileStamp))
									EndIf
								Case 5,6
									If StringLen(GUICtrlRead($MacroFileChange)) = 0 Then
										_Ok("Enter or select a file to check if changed or not.")
										$MacroSaveStep = False
									ElseIf Not FileExists(FilePathTranslateVariables(GUICtrlRead($MacroFileChange))) Then
										_Ok("File " & GUICtrlRead($MacroFileChange) & " doesn't exist.")
										$MacroSaveStep = False
									Else
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroFileType),GUICtrlRead($MacroFileChange))
										AddEndStep($AddingStep)
									EndIf
							EndSwitch
						Case $StepTypeControl
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroControlType))
						Case $StepTypeProcess
							Switch _GUICtrlComboBox_GetCurSel($MacroProcessType)
								Case 0
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroProcessType),_GUICtrlComboBox_GetCurSel($MacroProcessWindowType),GUICtrlRead($MacroProcessWindow),$ProcessList[SelectedProcessIndex()][0],$ProcessList[SelectedProcessIndex()][1],$ProcessList[SelectedProcessIndex()][2],$ProcessList[SelectedProcessIndex()][3])
								Case 1
									_GUICtrlTrimSpaces($MacroProcessProgram)
									_GUICtrlTrimSpaces($MacroProcessWindow)
									_GUICtrlTrimSpaces($MacroProcessWorkDir)
									If StringLen(GUICtrlRead($MacroProcessProgram)) = 0 Then
										_Ok("Enter or select a program.")
										$MacroSaveStep = False
									ElseIf Not FileExists(FilePathTranslateVariables(GUICtrlRead($MacroProcessProgram))) Then
										_Ok("Program file " & GUICtrlRead($MacroProcessProgram) & " doesn't exist.")
										$MacroSaveStep = False
									ElseIf StringLen(GUICtrlRead($MacroProcessWorkDir)) > 0 And Not FileExists(FilePathTranslateVariables(GUICtrlRead($MacroProcessWorkDir))) Then
										_Ok("Working folder " & GUICtrlRead($MacroProcessWorkDir) & " doesn't exist.")
										$MacroSaveStep = False
									Else
										If StringLen(GUICtrlRead($MacroProcessWindow)) = 0 Then _Ok("Strongly recommended: Give a keyword to activate the right window." & @CRLF &@CRLF & "Without a keyword " & $ProgramName & " will activate the first window it finds. This might be a wrong one like a HIDDEN one.")
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroProcessType),GUICtrlRead($MacroProcessProgram),GUICtrlRead($MacroProcessWindow),GUICtrlRead($MacroProcessWorkDir),_GUICtrlComboBox_GetCurSel($MacroProcessWindowState),GUICtrlRead($MacroProcessWait) = $GUI_CHECKED ? 1 : 0,GUICtrlRead($MacroProcessTimeout))
									EndIf
								Case 2,3
									_GUICtrlTrimSpaces($MacroProcessProcess)
									If StringLen(GUICtrlRead($MacroProcessProcess)) = 0 Then
										_Ok("Enter or select a process executable.")
										$MacroSaveStep = False
									Else
										EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroProcessType),GUICtrlRead($MacroProcessProcess))
										AddEndStep($AddingStep)
									EndIf
								Case 4,5
									If _GUICtrlComboBox_GetCurSel($MacroProcessType) = 3 And Not $HideProcessCheckWarning Then
										$HideProcessCheckWarning = HideMessage("It can be useful to set the process checking off for a macro that runs on multiple processes." & @CRLF & @CRLF & "But it can also be dangerous if your macro gets out of the programs you want to macro to run on.","Don't show this warning again",True)
										WriteSettings()
									EndIf
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroProcessType))
								Case 6
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroProcessType))
							EndSwitch
						Case $StepTypeCapture
							Switch _GUICtrlComboBox_GetCurSel($MacroCaptureAction)
								Case 0,2,4
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCaptureAction))
								Case 1
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCaptureAction),GUICtrlRead($MacroCaptureScreenX),GUICtrlRead($MacroCaptureScreenY),GUICtrlRead($MacroCaptureScreenWidth),GUICtrlRead($MacroCaptureScreenHeight))
								Case 3
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCaptureAction),GUICtrlRead($MacroCaptureWindowX),GUICtrlRead($MacroCaptureWindowY),GUICtrlRead($MacroCaptureWindowWidth),GUICtrlRead($MacroCaptureWindowHeight))
								Case 5
									_GUICtrlTrimSpaces($MacroCapturePath)
									EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCaptureAction),GUICtrlRead($MacroCapturePath),GUICtrlRead($MacroCapturePointer) = $GUI_CHECKED ? 1 : 0)
							EndSwitch
						Case $StepTypeShutdown
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroShutDownAction),GUICtrlRead($MacroShutDownAsk) = $GUI_CHECKED ? 1 : 0)
						Case $StepTypeComment
							EditMacroStep($AddingStep,_GUICtrlComboBox_GetCurSel($MacroStepType),GUICtrlRead($MacroDescription),_GUICtrlComboBox_GetCurSel($MacroCommentType))
					EndSwitch
					If $MacroSaveStep Then
						GUISetState(@SW_SHOW,$GUI)
						GUISetState(@SW_HIDE,$GUIStep)
						$AddingStep = False
					EndIf
				Case $MacroStepHelp
					Help($HelpFile,"step.htm")
				Case $MacroCancel
					$MacroDetectClick = False
					GUISetState(@SW_SHOW,$GUI)
					GUISetState(@SW_HIDE,$GUIStep)
				Case $Help
					Help($HelpFile)
				Case $Settings,$MenuItemSettings
					$MacroDetectClick = False
					_GUICtrlComboBox_SetCurSel($SettingsStartStopKey,$StartStopKey)
					_GUICtrlComboBox_SetCurSel($SettingsPauseKey,$PauseKey)
					GUICtrlSetState($SettingsShowStartupMessage,$HideStartupMessage ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowDetectionMessage,$HideDetectionMessage ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowShowMessage,$HideShowMessage ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetData($SettingsDefaultStepDelay,$DefaultStepDelay)
					GUICtrlSetState($SettingsShowLoopWarning,$HideLoopWarning ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetData($SettingsLeftIndent,$LeftIndent)
					GUICtrlSetState($SettingsShowPlayToolTip,$HidePlayToolTip ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowPlayMessage,$HidePlayMessage ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowVersionWarning,$HideEarlierVersionWarning ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetData($SettingsActivateTimeOut,$ActivateTimeOut)
					GUICtrlSetData($SettingsActivateDelay,$ActivateDelay)
					GUICtrlSetState($SettingsShowProcessCheckWarning,$HideProcessCheckWarning ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowRecordMessage,$HideRecordMessage ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowRecordTooltip,$HideRecordToolTip ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowShowRecordWarning,$HideShowRecordWarning ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowMoveTimeWarning,$HideMoveTimeWarning ? $GUI_UNCHECKED : $GUI_CHECKED)
					GUICtrlSetState($SettingsShowStepDelayWarning,$HideStepDelayWarning ? $GUI_UNCHECKED : $GUI_CHECKED)
					_GUICtrlComboBox_SetCurSel($SettingsThemes,$ProgramTheme+1)
					GUISetState(@SW_HIDE,$GUI)
					GUISetState(@SW_SHOW,$GUISettings)
				Case $SettingsShowStartupMessageLabel
					GUICtrlSetState($SettingsShowStartupMessage,GUICtrlRead($SettingsShowStartupMessage) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowDetectionMessageLabel
					GUICtrlSetState($SettingsShowDetectionMessage,GUICtrlRead($SettingsShowDetectionMessage) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowShowMessageLabel
					GUICtrlSetState($SettingsShowShowMessage,GUICtrlRead($SettingsShowShowMessage) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowPlayMessageLabel
					GUICtrlSetState($SettingsShowPlayMessage,GUICtrlRead($SettingsShowPlayMessage) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowVersionWarningLabel
					GUICtrlSetState($SettingsShowVersionWarning,GUICtrlRead($SettingsShowVersionWarning) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowPlayToolTipLabel
					GUICtrlSetState($SettingsShowPlayToolTip,GUICtrlRead($SettingsShowPlayToolTip) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowLoopWarningLabel
					GUICtrlSetState($SettingsShowLoopWarning,GUICtrlRead($SettingsShowLoopWarning) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowStepDelayWarningLabel
					GUICtrlSetState($SettingsShowStepDelayWarning,GUICtrlRead($SettingsShowStepDelayWarning) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowProcessCheckWarningLabel
					GUICtrlSetState($SettingsShowProcessCheckWarning,GUICtrlRead($SettingsShowProcessCheckWarning) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowRecordMessageLabel
					GUICtrlSetState($SettingsShowRecordMessage,GUICtrlRead($SettingsShowRecordMessage) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowRecordToolTipLabel
					GUICtrlSetState($SettingsShowRecordToolTip,GUICtrlRead($SettingsShowRecordToolTip) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowShowRecordWarningLabel
					GUICtrlSetState($SettingsShowShowRecordWarning,GUICtrlRead($SettingsShowShowRecordWarning) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsShowMoveTimeWarningLabel
					GUICtrlSetState($SettingsShowMoveTimeWarning,GUICtrlRead($SettingsShowMoveTimeWarning) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
				Case $SettingsDefaultStepDelay
					GUICtrlSetData($SettingsDefaultStepDelay,_Maximum(GUICtrlRead($SettingsDefaultStepDelay),10))
				Case $SettingsLeftIndent
					GUICtrlSetData($SettingsLeftIndent,_Maximum(GUICtrlRead($SettingsLeftIndent),1))
				Case $SettingsActivateTimeOut
					GUICtrlSetData($SettingsActivateTimeOut,_Maximum(GUICtrlRead($SettingsActivateTimeOut),100))
				Case $SettingsActivateDelay
					GUICtrlSetData($SettingsActivateDelay,_Maximum(GUICtrlRead($SettingsActivateDelay),100))
				Case $SettingsReset
					If _Sure("Reset all settings to default values?" & @CRLF & @CRLF & "Click Save to save the reset values.") Then
						_GUICtrlComboBox_SetCurSel($SettingsStartStopKey,0)
						_GUICtrlComboBox_SetCurSel($SettingsPauseKey,0)
						GUICtrlSetState($SettingsShowStartupMessage,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowDetectionMessage,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowShowMessage,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowLoopWarning,$GUI_CHECKED)
						GUICtrlSetData($SettingsDefaultStepDelay,100)
						GUICtrlSetData($SettingsLeftIndent,4)
						GUICtrlSetState($SettingsShowPlayToolTip,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowPlayMessage,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowVersionWarning,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowProcessCheckWarning,$GUI_CHECKED)
						GUICtrlSetData($SettingsActivateTimeOut,5000)
						GUICtrlSetData($SettingsActivateDelay,500)
						GUICtrlSetState($SettingsShowRecordMessage,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowRecordTooltip,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowShowRecordWarning,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowMoveTimeWarning,$GUI_CHECKED)
						GUICtrlSetState($SettingsShowStepDelayWarning,$GUI_CHECKED)
						_GUICtrlComboBox_SetCurSel($SettingsThemes,0)
					EndIf
				Case $SettingsSave
					$StartStopKey = _GUICtrlComboBox_GetCurSel($SettingsStartStopKey)
					$PauseKey = _GUICtrlComboBox_GetCurSel($SettingsPauseKey)
					$HideStartupMessage = GUICtrlRead($SettingsShowStartupMessage) <> $GUI_CHECKED
					$HideDetectionMessage = GUICtrlRead($SettingsShowDetectionMessage) <> $GUI_CHECKED
					$HideShowMessage = GUICtrlRead($SettingsShowShowMessage) <> $GUI_CHECKED
					$DefaultStepDelay = Number(GUICtrlRead($SettingsDefaultStepDelay))
					$HideLoopWarning = GUICtrlRead($SettingsShowLoopWarning) <> $GUI_CHECKED
					$LeftIndent = Number(GUICtrlRead($SettingsLeftIndent))
					$HidePlayToolTip = GUICtrlRead($SettingsShowPlayToolTip) <> $GUI_CHECKED
					$HidePlayMessage = GUICtrlRead($SettingsShowPlayMessage) <> $GUI_CHECKED
					$HideEarlierVersionWarning = GUICtrlRead($SettingsShowVersionWarning) <> $GUI_CHECKED
					$ActivateTimeOut = GUICtrlRead($SettingsActivateTimeOut)
					$ActivateDelay = GUICtrlRead($SettingsActivateDelay)
					$HideProcessCheckWarning = GUICtrlRead($SettingsShowProcessCheckWarning) <> $GUI_CHECKED
					$HideRecordMessage = GUICtrlRead($SettingsShowRecordMessage) <> $GUI_CHECKED
					$HideRecordToolTip = GUICtrlRead($SettingsShowRecordTooltip) <> $GUI_CHECKED
					$HideShowRecordWarning = GUICtrlRead($SettingsShowShowRecordWarning) <> $GUI_CHECKED
					$HideMoveTimeWarning = GUICtrlRead($SettingsShowMoveTimeWarning) <> $GUI_CHECKED
					$HideStepDelayWarning = GUICtrlRead($SettingsShowStepDelayWarning) <> $GUI_CHECKED
					$ProgramTheme = _GUICtrlComboBox_GetCurSel($SettingsThemes)-1
					If $ProgramTheme > -1 Then
						GUISetBkColor($Themes[$ProgramTheme][0],$GUI)
						GUICtrlSetBkColor($Macros,$Themes[$ProgramTheme][2])
						GUICtrlSetColor($ProcessesLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($MacroStepsLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($ShowRecordingLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($MoveTimeLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($MoveTime,$Themes[$ProgramTheme][3])
						GUICtrlSetBkColor($MoveTime,$Themes[$ProgramTheme][2])
						GUICtrlSetColor($TestMacroLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($ShowBarLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($StepDelayLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($StepDelay,$Themes[$ProgramTheme][3])
						GUICtrlSetBkColor($StepDelay,$Themes[$ProgramTheme][2])
						GUICtrlSetColor($StepDelayAlwaysLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($LoopCountLabel,$Themes[$ProgramTheme][1])
						GUICtrlSetColor($LoopCount,$Themes[$ProgramTheme][3])
						GUICtrlSetBkColor($LoopCount,$Themes[$ProgramTheme][2])
					Else
						GUISetBkColor(_WindowGetBkColor(),$GUI)
						GUICtrlSetBkColor($Macros,0xFFFFFF)
					EndIf
					WriteSettings()
					FillMacroSteps($MacroStep)
					GUISetState(@SW_HIDE,$GUISettings)
					GUISetState(@SW_SHOW,$GUI)
				Case $SettingsCancel
					GUISetState(@SW_HIDE,$GUISettings)
					GUISetState(@SW_SHOW,$GUI)
				Case $CopyClose
					GUISetState(@SW_HIDE,$GUICopy)
					GUISetState(@SW_SHOW,$GUI)
				Case $FindClose
					GUISetState(@SW_HIDE,$GUIFind)
					GUISetState(@SW_SHOW,$GUI)
				Case $FindGoToClose
					GUISetState(@SW_HIDE,$GUIGoTo)
					GUISetState(@SW_SHOW,$GUI)
				Case $EnableDisableClose
					GUISetState(@SW_HIDE,$GUIEnableDisable)
					GUISetState(@SW_SHOW,$GUI)
				Case $GUI_EVENT_CLOSE,$Close,$MenuItemExit
					If $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $GUISettings Then
						GUISetState(@SW_HIDE,$GUISettings)
						GUISetState(@SW_SHOW,$GUI)
					ElseIf $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $GUICopy Then
						GUISetState(@SW_HIDE,$GUICopy)
						GUISetState(@SW_SHOW,$GUI)
					ElseIf $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $GUIFind Then
						GUISetState(@SW_HIDE,$GUIFind)
						GUISetState(@SW_SHOW,$GUI)
					ElseIf $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $GUIGoTo Then
						GUISetState(@SW_HIDE,$GUIGoTo)
						GUISetState(@SW_SHOW,$GUI)
					ElseIf $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $GUIEnableDisable Then
						GUISetState(@SW_HIDE,$GUIEnableDisable)
						GUISetState(@SW_SHOW,$GUI)
					ElseIf $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $GUIStep Then
						$MacroDetectClick = False
						GUISetState(@SW_HIDE,$GUIStep)
						GUISetState(@SW_SHOW,$GUI)
					Else
						If Not $Changed Then ExitLoop
						Switch _MessageBox("Exit " & $ProgramName & " by saving or not saving your macro?","",0,"&Save|Save macro and exit " & $ProgramName,"&Exit|Exit " & $ProgramName & " without saving","S&tay|Stay in " & $ProgramName,"",$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
							Case 1
								If SaveMacro() Then ExitLoop
							Case 2
								ExitLoop
						EndSwitch
					EndIf
			EndSwitch
			EnableDisableButtons()	; enable/disable buttons
			; focus on macro step list for arrows usage, rule out edit, settings, copy range windows
			If $aMsg[0] > 0 And Not _InList($aMsg[0],$Add,$Change,$MenuItemCopySteps,$MenuItemCutSteps,$MenuItemFind,$MenuItemFindGoTo,$MenuItemEnableDisableSteps,$Settings) And $aMsg[1] = $GUI Then GUICtrlSetState($Macros,$GUI_FOCUS)
		WEnd
	EndIf
	_GraphicButtons(False)
	_GDIPlus_ShutDown()
	GUIDelete($GUI)
	GUIDelete($GUIGoTo)
	GUIDelete($GUIFind)
	GUIDelete($GUICopy)
	GUIDelete($GUIStep)
	GUIDelete($GUISettings)
	DllClose($User32DLL)
	_Inputmask_MsgRegister(False)	; Unregister for _InputMask
	_Inputmask_close()				; Memory cleanup for Input masks
EndFunc

Func Help($HelpFile,$Topic = "",$Anchor = "")	; pop up manual
	If StringLen($Topic) = 0 Then
		ShellExecute(@WindowsDir & "\hh.exe","its:" & $HelpFile)
	ElseIf StringLen($Anchor) Then
		ShellExecute(@WindowsDir & "\hh.exe","its:" & $HelpFile & "::/" & $Topic)
	Else
		ShellExecute(@WindowsDir & "\hh.exe","its:" & $HelpFile & "::/" & $Topic & "#" & $Anchor)
	EndIf
EndFunc

Func SetF1($GUI,$HelpFile,$Topic = "")			; pop up manual by pressing F1
	If StringLen($Topic) = 0 Then
		GUISetHelp(@WindowsDir & "\hh.exe its:" & $HelpFile,$GUI)
	Else
		GUISetHelp(@WindowsDir & "\hh.exe its:" & $HelpFile & "::/" & $Topic,$GUI)
	EndIf
EndFunc

Func FindByType($StepType,$Reverse = False)		; find macro step by its type
	If $MacroStep = -1 Then Return False
	Local $StepFound = False
	If Not $Reverse Then
		If $MacroStep < UBound($MacroSteps)-1 Then
			For $Step = $MacroStep+1 To UBound($MacroSteps)-1
				If $MacroSteps[$Step][0] = $StepType Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
		If Not $StepFound Then
			For $Step = 0 To $MacroStep
				If $MacroSteps[$Step][0] = $StepType Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
		Else
		If $MacroStep > 0 Then
			For $Step = $MacroStep-1 To 0 Step -1
				If $MacroSteps[$Step][0] = $StepType Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
		If Not $StepFound Then
			For $Step = UBound($MacroSteps)-1 To $MacroStep Step -1
				If $MacroSteps[$Step][0] = $StepType Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
	EndIf
	If Not $StepFound Then _Ok("A step of type " & $MacroStepTypes[$StepType][0] & " wasn't found.")
	Return $StepFound
EndFunc

Func FindMarker($Reverse,$Column,$Text)			; find marked macro step
	If $MacroStep = -1 Then Return False
	Local $StepFound = False
	If Not $Reverse Then
		If $MacroStep < UBound($MacroSteps)-1 Then
			For $Step = $MacroStep+1 To UBound($MacroSteps)-1
				If $MacroSteps[$Step][$Column] = 1 Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
		If Not $StepFound Then
			For $Step = 0 To $MacroStep
				If $MacroSteps[$Step][$Column] = 1 Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
		Else
		If $MacroStep > 0 Then
			For $Step = $MacroStep-1 To 0 Step -1
				If $MacroSteps[$Step][$Column] = 1 Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
		If Not $StepFound Then
			For $Step = UBound($MacroSteps)-1 To $MacroStep Step -1
				If $MacroSteps[$Step][$Column] = 1 Then
					$StepFound = True
					SelectStep($Step)
					ExitLoop
				EndIf
			Next
		EndIf
	EndIf
	If Not $StepFound Then
		_Ok($Text)
	Else
		; redraw needed else arrow down causes listview to jump to first step
		FillMacroSteps($MacroStep)
		$MacroPreviousStep = -1	; force selected item refresh
	EndIf
	Return $StepFound
EndFunc

Func SelectStep($Step)
	_GUICtrlListView_SetItemFocused($Macros,$MacroStep,False)
	$MacroStepPrevious = $MacroStep
	$MacroStep = $Step
	_GUICtrlListView_SetItemSelected($Macros,$MacroStep)
	_GUICtrlListView_EnsureVisible($Macros,$MacroStep)
EndFunc

Func SetCounterLabelsAndTips()
	If _GUICtrlComboBox_GetCurSel($MacroCounterType) = 0 Then
		GUICtrlSetData($MacroCounterSetLabel,"Set counter value")
		GUICtrlSetTip($MacroCounterSet,"Set the initial value of the counter")
	ElseIf _GUICtrlComboBox_GetCurSel($MacroCounterType) = 2 Then
		GUICtrlSetData($MacroCounterSetLabel,"Add counter value")
		GUICtrlSetTip($MacroCounterSet,"Add a value to the counter")
	EndIf
	If _GUICtrlComboBox_GetCurSel($MacroCounterType) >= 5 Then
		GUICtrlSetTip($MacroCounterValue1,"Greater or equal than value")
	ElseIf _GUICtrlComboBox_GetCurSel($MacroCounterType) = 4 Then
		GUICtrlSetTip($MacroCounterValue1,"Not equal to value")
	Else
		GUICtrlSetTip($MacroCounterValue1,"Equal to value")
	EndIf
EndFunc

Func SetStepView($View)
	$StepView = $View
	WriteSettings()
	Switch $View
		Case 1
			_GUICtrlListView_SetView($Macros,1)
			_GUICtrlListView_SetExtendedListViewStyle($Macros,$LVS_EX_FULLROWSELECT+$WS_EX_CLIENTEDGE+$LVS_EX_GRIDLINES)
		Case 2
			_GUICtrlListView_SetView($Macros,1)
			_GUICtrlListView_SetExtendedListViewStyle($Macros,$LVS_EX_FULLROWSELECT+$WS_EX_CLIENTEDGE)
		Case 3
			_GUICtrlListView_SetView($Macros,2)
			_GUICtrlListView_SetExtendedListViewStyle($Macros,$LVS_EX_FULLROWSELECT+$WS_EX_CLIENTEDGE)
	EndSwitch
EndFunc

Func SetStepColor($Step,$Color,$BkColor)
	GUICtrlSetColor(_GUICtrlListView_GetItemParam($Macros,$Step),$Color)
	GUICtrlSetBkColor(_GUICtrlListView_GetItemParam($Macros,$Step),$BkColor)
EndFunc

Func HighlightStep($Step,$PreviousStep)
	Local $Image

	If $Step > -1 Then
		$Image = $MacroSteps[$Step][$MacroStepMarkColumn] = 1 ? 2 : 0
		If $MacroSteps[$Step][$MacroStepEnableDisableColumn] = 1 Then	; step disabled
			_GUICtrlListView_SetItemImage($Macros,$Step,$Image)
		Else
			_GUICtrlListView_SetItemImage($Macros,$Step,$Image-1)
		EndIf
		SetStepColor($Step,0xFFFFFF,$ColorHighLight)
	EndIf
	If $PreviousStep > -1 And $PreviousStep < _GUICtrlListView_GetItemCount($Macros) Then
		$Image = $MacroSteps[$PreviousStep][$MacroStepMarkColumn] = 1 ? 2 : 0
		If $MacroSteps[$PreviousStep][$MacroStepEnableDisableColumn] = 1 Then	; step disabled
			_GUICtrlListView_SetItemImage($Macros,$PreviousStep,$Image)
		Else
			_GUICtrlListView_SetItemImage($Macros,$PreviousStep,$Image-1)
		EndIf
		If $ViewInvertColors Then
			If $MacroSteps[$PreviousStep][$MacroStepEnableDisableColumn] = 1 Then	; step disabled
				SetStepColor($PreviousStep,0xFFFFFF,0xD0D0D0)
			Else
				SetStepColor($PreviousStep,0,$MacroSteps[$PreviousStep][$MacroStepColorColumn])
			EndIf
		ElseIf $MacroSteps[$PreviousStep][$MacroStepEnableDisableColumn] = 1 Then	; step disabled
			SetStepColor($PreviousStep,0xA0A0A0,0xFFFFFF)
		Else
			SetStepColor($PreviousStep,$MacroSteps[$PreviousStep][$MacroStepColorColumn],$ProgramTheme > -1 ? $Themes[$ProgramTheme][2]: 0xFFFFFF)
		EndIf
	EndIf
EndFunc

Func EnableDisableButtons()
	If $MacroStepCount > 0 Then
		If $Changed Then
			_GUICtrlChangeState($Save,$GUI_ENABLE)
		Else
			_GUICtrlChangeState($Save,$GUI_DISABLE)
		EndIf
		_GUICtrlChangeState($SaveAs,$GUI_ENABLE)
	Else
		_GUICtrlChangeState($Save,$GUI_DISABLE)
		_GUICtrlChangeState($SaveAs,$GUI_DISABLE)
	EndIf
	If $MacroStep = -1 Then
		_GUICtrlChangeState($Change,$GUI_DISABLE)
		_GUICtrlChangeState($Duplicate,$GUI_DISABLE)
		_GUICtrlChangeState($Delete,$GUI_DISABLE)
		_GUICtrlChangeState($Up,$GUI_DISABLE)
		_GUICtrlChangeState($Down,$GUI_DISABLE)
		_GUICtrlChangeState($Show,$GUI_DISABLE)
		_GUICtrlChangeState($Enable,$GUI_DISABLE)
		_GUICtrlChangeState($Enable,$GUI_HIDE)
		_GUICtrlChangeState($Disable,$GUI_DISABLE)
	Else
		_GUICtrlChangeState($Change,$GUI_ENABLE)
		_GUICtrlChangeState($Duplicate,$GUI_ENABLE)
		_GUICtrlChangeState($Delete,$GUI_ENABLE)
		If $MacroStepCount > 1 Then
			If $MacroStep > 0 Then
				_GUICtrlChangeState($Up,$GUI_ENABLE)
			Else
				_GUICtrlChangeState($Up,$GUI_DISABLE)
			EndIf
			If $MacroStep < $MacroStepCount-1 Then
				_GUICtrlChangeState($Down,$GUI_ENABLE)
			Else
				_GUICtrlChangeState($Down,$GUI_DISABLE)
			EndIf
		Else
			_GUICtrlChangeState($Up,$GUI_DISABLE)
			_GUICtrlChangeState($Down,$GUI_DISABLE)
		EndIf
		Switch $MacroSteps[$MacroStep][0]
			Case $StepTypeMouse
				_GUICtrlChangeState($Show,$GUI_ENABLE)
			Case $StepTypeWindow
				If $MacroSteps[$MacroStep][2] = 3 Then
					_GUICtrlChangeState($Show,$GUI_ENABLE)
				Else
					_GUICtrlChangeState($Show,$GUI_DISABLE)
				EndIf
			Case $StepTypePixel
				If _InList($MacroSteps[$MacroStep][2],0,2) Then
					_GUICtrlChangeState($Show,$GUI_ENABLE)
				Else
					_GUICtrlChangeState($Show,$GUI_DISABLE)
				EndIf
			Case $StepTypeCapture
				If _InList($MacroSteps[$MacroStep][2],1,3) Then
					_GUICtrlChangeState($Show,$GUI_ENABLE)
				Else
					_GUICtrlChangeState($Show,$GUI_DISABLE)
				EndIf
			Case Else
				_GUICtrlChangeState($Show,$GUI_DISABLE)
		EndSwitch
		If $MacroSteps[$MacroStep][$MacroStepEnableDisableColumn] = 1 Then	; step disabled
			_GUICtrlChangeState($Disable,$GUI_DISABLE)
			_GUICtrlChangeState($Disable,$GUI_HIDE)
			_GUICtrlChangeState($Enable,$GUI_ENABLE)
			_GUICtrlChangeState($Enable,$GUI_SHOW)
		Else
			_GUICtrlChangeState($Enable,$GUI_DISABLE)
			_GUICtrlChangeState($Enable,$GUI_HIDE)
			_GUICtrlChangeState($Disable,$GUI_ENABLE)
			_GUICtrlChangeState($Disable,$GUI_SHOW)
		EndIf
	EndIf
EndFunc

Func AddEndStep($AddingStep)
	If $AddingStep Then
		EditMacroStep(True,$StepTypeEnd)
		$MacroPreviousStep = $MacroStep		; set automatic highlighing off of end step
		$MacroStep -= 1
		_GUICtrlListView_SetItemSelected($Macros,$MacroStep,True,True)
	EndIf
EndFunc

Func SelectFile($Title,$Filter,$Control)
	If StringLen($Filter) = 0 Then $Filter = "All files (*.*)"
	Local $FileName = FileOpenDialog($Title,"",$Filter,BitOR($FD_FILEMUSTEXIST,$FD_PATHMUSTEXIST),GUICtrlRead($Control),$GUI)
	If @error Or StringLen(StringStripWS($FileName,$STR_STRIPLEADING+$STR_STRIPTRAILING)) = 0 Then Return False
	GUICtrlSetData($Control,$FileName)
	Return True
EndFunc

Func SelectPath($Title,$Control)
	Local $PathName = FileSelectFolder($Title,"",0,GUICtrlRead($Control),$GUI)
	If @error Or StringLen(StringStripWS($PathName,$STR_STRIPLEADING+$STR_STRIPTRAILING)) = 0 Then Return False
	GUICtrlSetData($Control,$PathName)
	Return True
EndFunc

Func FillCounterNames($Combo,$Name = "")
	_GUICtrlComboBox_BeginUpdate($Combo)
	_GUICtrlComboBox_ResetContent($Combo)
	For $Step = 0 To UBound($MacroSteps)-1
		If $MacroSteps[$Step][0] = $StepTypeCounter And _InList($MacroSteps[$Step][2],0,1) Then _GUICtrlComboBox_AddString($Combo,$MacroSteps[$Step][3])
	Next
	_GUICtrlComboBox_EndUpdate($Combo)
	If StringLen($Name) > 0 And _GUICtrlComboBox_FindString($Combo,$Name) > 0 Then
		_GUICtrlComboBox_SetCurSel($Combo,_GUICtrlComboBox_FindString($Combo,$Name))
	Else
		_GUICtrlComboBox_SetCurSel($Combo,0)
	EndIf
	Return _GUICtrlComboBox_GetCount($Combo)
EndFunc

Func FillJumpLabels($Combo,$Label = "")
	_GUICtrlComboBox_BeginUpdate($Combo)
	_GUICtrlComboBox_ResetContent($Combo)
	For $Step = 0 To UBound($MacroSteps)-1
		If $MacroSteps[$Step][0] = $StepTypeLabel And $MacroSteps[$Step][2] = 0 Then _GUICtrlComboBox_AddString($Combo,$MacroSteps[$Step][1])
	Next
	_GUICtrlComboBox_EndUpdate($Combo)
	If StringLen($Label) > 0  And _GUICtrlComboBox_FindString($Combo,$Label) > 0 Then
		_GUICtrlComboBox_SetCurSel($Combo,_GUICtrlComboBox_FindString($Combo,$Label))
	Else
		_GUICtrlComboBox_SetCurSel($Combo,0)
	EndIf
	Return _GUICtrlComboBox_GetCount($Combo)
EndFunc

Func ChangeJumpLabels($AddingStep,$OldLabel,$NewLabel)
	If Not $AddingStep And $OldLabel <> $NewLabel Then
		For $Step = 0 To UBound($MacroSteps)-1
			If $MacroSteps[$Step][0] = 6 And $MacroSteps[$Step][3] = $OldLabel Then $MacroSteps[$Step][3] = $NewLabel
		Next
	EndIf
EndFunc

Func MacroStepLabelExists($Label,$CheckStep = False)
	Local $ChangingStep

	If $MacroStepCount = 0 Then Return False
	If $CheckStep Then
		$ChangingStep = $MacroStep
	Else
		$ChangingStep = -1
	EndIf
	For $Step = 0 to UBound($MacroSteps)-1
		If $Step <> $ChangingStep And $MacroSteps[$Step][0] = 5 And $MacroSteps[$Step][1] = $Label Then Return True
	Next
	Return False
EndFunc

Func SetPlayTip()								; set play button tooltip according to set play way
	Local $Test = " Ctrl+P" & (GUICtrlRead($TestMacro) = $GUI_CHECKED ? " Testing" : "") & @CRLF & "Stop macro " & $StartStopKeys[$StartStopKey][0]

	If GUICtrlRead($LoopCount) > 1 Then
		GUICtrlSetTip($Play,"Play " & GUICtrlRead($LoopCount) & " times" & $Test)
	ElseIf GUICtrlRead($LoopCount) = 1 Then
		GUICtrlSetTip($Play,"Play once" & $Test)
	Else
		GUICtrlSetTip($Play,"Play indefinitely (looped)" & $Test)
	EndIf
EndFunc

Func SetStopKey($Set = True)
	For $Key = 0 To UBound($StartStopKeys)-1
		HotKeySet($StartStopKeys[$Key][1])
	Next
	If $Set Then
		HotKeySet($StartStopKeys[$StartStopKey][1],"StopMacro")
		$Stopped = 0
	EndIf
EndFunc

Func StopMacro()
	$Stopped = 1
EndFunc

Func SetPauseKey($Set = True)
	For $Key = 0 To UBound($PauseKeys)-1
		HotKeySet($PauseKeys[$Key][1])
	Next
	If $Set Then
		HotKeySet($PauseKeys[$PauseKey][1],"PauseMacro")
		$Paused = False
	EndIf
EndFunc

Func PauseMacro()
	$Paused = Not $Paused
EndFunc

Func InitializeMacroStepControls()				; initialize all macro step controls
	_GUICtrlComboBox_SetCurSel($MacroStepType,0)
	GUICtrlSetData($MacroDescription,"")
	_GUICtrlComboBox_SetCurSel($MacroMouseButton,0)
	GUICtrlSetData($MacroPositionX,0)
	GUICtrlSetData($MacroPositionY,0)
	_GUICtrlComboBox_SetCurSel($MacroMousePosition,0)
	GUICtrlSetData($MacroPositionXRandom,0)
	GUICtrlSetData($MacroPositionYRandom,0)
	GUICtrlSetState($MacroMouseMoveInstantaneous,$GUI_UNCHECKED)
	GUICtrlSetData($MacroMouseClicks,1)
	_GUICtrlComboBox_SetCurSel($MacroKey,0)
	_GUICtrlComboBox_SetCurSel($MacroKeyMode,0)
	_GUICtrlComboBox_SetCurSel($MacroTextSend,0)
	GUICtrlSetData($MacroText,"")
	_GUICtrlComboBox_SetCurSel($MacroDelayType,0)
	GUICtrlSetData($MacroDelay,0)
	GUICtrlSetData($MacroDelayRandom,0)
	GUICtrlSetData($MacroDelayStep,GUICtrlRead($StepDelay))
	_GUICtrlComboBox_SetCurSel($MacroWindowAction,0)
	_GUICtrlComboBox_SetCurSel($MacroWindowPosition,0)
	GUICtrlSetState($MacroWindowMoveInstantaneous,$GUI_UNCHECKED)
	GUICtrlSetState($MacroWindowResizeInstantaneous,$GUI_UNCHECKED)
	_GUICtrlComboBox_SetCurSel($MacroWindowWindow,"")
	GUICtrlSetData($MacroWidth,WindowWidth($ProcessList[SelectedProcessIndex()][2]))
	GUICtrlSetData($MacroHeight,WindowHeight($ProcessList[SelectedProcessIndex()][2]))
	_GUICtrlComboBox_SetCurSel($MacroWindowSize,0)
	GUICtrlSetData($MacroRandomMouse,0)
	GUICtrlSetData($MacroRandomDelay,0)
	_GUICtrlComboBox_SetCurSel($MacroJumpType,0)
	FillJumpLabels($MacroJumpLabel)
	_GUICtrlComboBox_SetCurSel($MacroRepeatType,0)
	GUICtrlSetData($MacroRepeatCount,1)
	GUICtrlSetData($MacroRepeatValue1,1)
	GUICtrlSetData($MacroRepeatValue2,1)
	GUICtrlSetData($MacroTimeLoop,1)
	_GUICtrlComboBox_SetCurSel($MacroInputType,0)
	_GUICtrlComboBox_SetCurSel($MacroContinueKey,0)
	_GUICtrlComboBox_SetCurSel($MacroPixelType,0)
	GUICtrlSetData($MacroPixelWidth,1)
	GUICtrlSetData($MacroPixelHeight,1)
	GUICtrlSetData($MacroPixelId,1)
	SetControlTime($MacroAlarmTime)
	SetControlDate($MacroAlarmDate)
	_GUICtrlComboBox_SetCurSel($MacroAlarmSpecific,0)
	_GUICtrlComboBox_SetCurSel($MacroSoundType,1)
	GUICtrlSetData($MacroSoundFile,"")
	GUICtrlSetData($MacroSoundVolume,100)
	GUICtrlSetData($MacroSoundFrequency,1000)
	GUICtrlSetData($MacroSoundDuration,500)
	GUICtrlSetData($MacroCounterName,"")
	FillCounterNames($MacroCounterNameList)
	GUICtrlSetData($MacroCounterSet,0)
	GUICtrlSetState($MacroCounterRandom,$GUI_UNCHECKED)
	GUICtrlSetData($MacroCounterValue1,0)
	GUICtrlSetData($MacroCounterValue2,1)
	GUICtrlSetData($MacroFileFile,"")
	GUICtrlSetData($MacroFilePath,"")
	GUICtrlSetData($MacroFileStamp,"")
	GUICtrlSetData($MacroFileChange,"")
	_GUICtrlComboBox_SetCurSel($MacroProcessType,0)
	GUICtrlSetData($MacroProcessShowProcessCurrent,$ProcessList[SelectedProcessIndex()][0] & " - " & $ProcessList[SelectedProcessIndex()][1])
	GUICtrlSetData($MacroProcessShowProcess,"None")
	_GUICtrlComboBox_SetCurSel($MacroProcessWindowType,0)
	GUICtrlSetData($MacroProcessProgram,"")
	GUICtrlSetData($MacroProcessWindow,"")
	GUICtrlSetData($MacroProcessWorkDir,"")
	_GUICtrlComboBox_SetCurSel($MacroProcessWindowState,0)
	GUICtrlSetState($MacroProcessWait,$GUI_UNCHECKED)
	GUICtrlSetData($MacroProcessTimeout,5000)
	GUICtrlSetData($MacroProcessProcess,"")
	_GUICtrlComboBox_SetCurSel($MacroCaptureAction,0)
	GUICtrlSetData($MacroCaptureScreenX,0)
	GUICtrlSetData($MacroCaptureScreenY,0)
	GUICtrlSetData($MacroCaptureScreenWidth,1)
	GUICtrlSetData($MacroCaptureScreenHeight,1)
	GUICtrlSetData($MacroCaptureWindowX,0)
	GUICtrlSetData($MacroCaptureWindowY,0)
	GUICtrlSetData($MacroCaptureWindowWidth,-1)
	GUICtrlSetData($MacroCaptureWindowHeight,-1)
	GUICtrlSetData($MacroCapturePath,"")
	GUICtrlSetState($MacroCapturePointer,$GUI_CHECKED)
	_GUICtrlComboBox_SetCurSel($MacroShutDownAction,0)
	GUICtrlSetState($MacroShutDownAsk,$GUI_UNCHECKED)
	_GUICtrlComboBox_SetCurSel($MacroCommentType,0)
EndFunc

Func AddTime($Time,$AddTime)					; add time period to given time
	Local $Hour = 0,$Minute = 0,$Second = 0,$AddHour = 0,$AddMinute = 0,$AddSecond

	If $Time >= 10000 Then
		$Hour = Floor($Time/10000)
		$Time -= $Hour*10000
	EndIf
	If $Time >= 100 Then
		$Minute = Floor($Time/100)
		$Time -= $Minute*100
	EndIf
	$Second = $Time
	If $AddTime >= 10000 Then
		$AddHour = Floor($AddTime/10000)
		$AddTime -= $AddHour*10000
	EndIf
	If $AddTime >= 100 Then
		$AddMinute = Floor($AddTime/100)
		$AddTime -= $AddMinute*100
	EndIf
	$AddSecond = $AddTime

	$Second += $AddSecond
	If $Second >= 60 Then
		$Second -= 60
		$Minute += 1
	EndIf
	$Minute += $AddMinute
	If $Minute >= 60 Then
		$Minute -= 60
		$Hour += 1
	EndIf
	$Hour += $AddHour
	If $Hour >= 24 Then
		$Hour -=24
	EndIf
	Return $Hour*10000+$Minute*100+$Second
EndFunc

Func CurrentTime()
	Return @HOUR*10000+@MIN*100+@SEC
EndFunc

Func CurrentDate()
	Return @YEAR*10000+@MON*100+@MDAY
EndFunc

Func DateString($Date)
	Local $Year = 1,$Month = 0,$Day = 0

	If $Date > 0 Then
		If $Date >= 10000 Then
			$Year = Floor($Date/10000)
			$Date -= $Year*10000
		EndIf
		If $Date >= 100 Then
			$Month = Floor($Date/100)
			$Date -= $Month*100
		EndIf
		$Day = $Date
	EndIf
	Return _DateTimeFormat(StringFormat("%04d/%02d/%02d",$Year,$Month,$Day),2)
EndFunc

Func TimeString($Time)
	Local $Hour = 0,$Minute = 0,$Second = 0

	If $Time >= 0 Then
		If $Time >= 10000 Then
			$Hour = Floor($Time/10000)
			$Time -= $Hour*10000
		EndIf
		If $Time >= 100 Then
			$Minute = Floor($Time/100)
			$Time -= $Minute*100
		EndIf
		$Second = $Time
	EndIf
	Return StringFormat("%02d:%02d:%02d",$Hour,$Minute,$Second)
EndFunc

Func SetControlDate($Handle,$Date = 0)
	Local $Year = @YEAR,$Month = @MON,$Day = @MDAY,$Hour = 0,$Minute = 0,$Second = 0

	If $Date > 0 Then
		If $Date >= 10000 Then
			$Year = Floor($Date/10000)
			$Date -= $Year*10000
		EndIf
		If $Date >= 100 Then
			$Month = Floor($Date/100)
			$Date -= $Month*100
		EndIf
		$Day = $Date
	EndIf
	SetControlDateTime($Handle,$Year,$Month,$Day,$Hour,$Minute,$Second)
EndFunc

Func SetControlTime($Handle,$Time = -1)
	Local $Year = @YEAR,$Month = @MON,$Day = @MDAY,$Hour = 0,$Minute = 0,$Second = 0

	If $Time >= 0 Then
		If $Time >= 10000 Then
			$Hour = Floor($Time/10000)
			$Time -= $Hour*10000
		EndIf
		If $Time >= 100 Then
			$Minute = Floor($Time/100)
			$Time -= $Minute*100
		EndIf
		$Second = $Time
	Else
		$Hour = @HOUR
		$Minute = @MIN
		$Second = @SEC
	EndIf
	SetControlDateTime($Handle,$Year,$Month,$Day,$Hour,$Minute,$Second)
EndFunc

Func SetControlDateTime($Handle,$Year = @YEAR,$Month = @MON,$Day = @MDAY,$Hour = @HOUR,$Minute = @MIN,$Second = @SEC)
	Local $DateTime = [false,$Year,$Month,$Day,$Hour,$Minute,$Second]
	_GUICtrlDTP_SetSystemTime(GUICtrlGetHandle($Handle),$DateTime)
EndFunc

Func GetControlDate($Handle)
	Local $DateTime = _GUICtrlDTP_GetSystemTime(GUICtrlGetHandle($Handle))
	Return $DateTime[0]*10000+$DateTime[1]*100+$DateTime[2]
EndFunc

Func GetControlTime($Handle)
	Local $DateTime = _GUICtrlDTP_GetSystemTime(GUICtrlGetHandle($Handle))
	Return $DateTime[3]*10000+$DateTime[4]*100+$DateTime[5]
EndFunc

Func ShowHideMacroStepControls()				; show and hide appropriate macro step controls
	GUICtrlSetState($MacroMouseButton,$GUI_HIDE)
	GUICtrlSetState($MacroMouseButtonLabel,$GUI_HIDE)
	GUICtrlSetState($MacroDetectPosition,$GUI_HIDE)
	GUICtrlSetState($MacroLabelX,$GUI_HIDE)
	GUICtrlSetState($MacroPositionX,$GUI_HIDE)
	GUICtrlSetState($MacroLabelY,$GUI_HIDE)
	GUICtrlSetState($MacroPositionY,$GUI_HIDE)
	GUICtrlSetState($MacroShowPosition,$GUI_HIDE)
	GUICtrlSetState($MacroPositionXRandomLabel,$GUI_HIDE)
	GUICtrlSetState($MacroPositionXRandom,$GUI_HIDE)
	GUICtrlSetState($MacroPositionYRandomLabel,$GUI_HIDE)
	GUICtrlSetState($MacroPositionYRandom,$GUI_HIDE)
	GUICtrlSetState($MacroMousePosition,$GUI_HIDE)
	GUICtrlSetState($MacroMouseMoveInstantaneous,$GUI_HIDE)
	GUICtrlSetState($MacroMouseMoveInstantaneousLabel,$GUI_HIDE)
	GUICtrlSetState($MacroMouseClicks,$GUI_HIDE)
	GUICtrlSetState($MacroMouseClicksLabel,$GUI_HIDE)
	GUICtrlSetState($MacroKeyLabel,$GUI_HIDE)
	GUICtrlSetState($MacroKey,$GUI_HIDE)
	GUICtrlSetState($MacroKeyModeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroKeyMode,$GUI_HIDE)
	GUICtrlSetState($MacroTextSendLabel,$GUI_HIDE)
	GUICtrlSetState($MacroTextSend,$GUI_HIDE)
	GUICtrlSetState($MacroTextLabel,$GUI_HIDE)
	GUICtrlSetState($MacroText,$GUI_HIDE)
	GUICtrlSetState($MacroDelayTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroDelayType,$GUI_HIDE)
	GUICtrlSetState($MacroDelayLabel,$GUI_HIDE)
	GUICtrlSetState($MacroDelay,$GUI_HIDE)
	GUICtrlSetState($MacroDelayRandomLabel,$GUI_HIDE)
	GUICtrlSetState($MacroDelayRandom,$GUI_HIDE)
	GUICtrlSetState($MacroDelayStepLabel,$GUI_HIDE)
	GUICtrlSetState($MacroDelayStep,$GUI_HIDE)
	GUICtrlSetState($MacroWindowActionLabel,$GUI_HIDE)
	GUICtrlSetState($MacroWindowAction,$GUI_HIDE)
	GUICtrlSetState($MacroWindowPosition,$GUI_HIDE)
	GUICtrlSetState($MacroWindowMoveInstantaneous,$GUI_HIDE)
	GUICtrlSetState($MacroWindowMoveInstantaneousLabel,$GUI_HIDE)
	GUICtrlSetState($MacroWindowResizeInstantaneous,$GUI_HIDE)
	GUICtrlSetState($MacroWindowResizeInstantaneousLabel,$GUI_HIDE)
	GUICtrlSetState($MacroWindowWindowLabel,$GUI_HIDE)
	GUICtrlSetState($MacroWindowWindow,$GUI_HIDE)
	GUICtrlSetState($MacroWidthLabel,$GUI_HIDE)
	GUICtrlSetState($MacroWidth,$GUI_HIDE)
	GUICtrlSetState($MacroHeightLabel,$GUI_HIDE)
	GUICtrlSetState($MacroHeight,$GUI_HIDE)
	GUICtrlSetState($MacroWindowSize,$GUI_HIDE)
	GUICtrlSetState($MacroRandomMouseLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRandomMouse,$GUI_HIDE)
	GUICtrlSetState($MacroRandomDelayLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRandomDelay,$GUI_HIDE)
	GUICtrlSetState($MacroJumpTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroJumpType,$GUI_HIDE)
	GUICtrlSetState($MacroJumpLabelLabel,$GUI_HIDE)
	GUICtrlSetState($MacroJumpLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatType,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatCountLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatCount,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatValue1aLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatValue1bLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatValue1cLabel,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatValue1,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatValue2Label,$GUI_HIDE)
	GUICtrlSetState($MacroRepeatValue2,$GUI_HIDE)
	GUICtrlSetState($MacroTimeLoopLabel,$GUI_HIDE)
	GUICtrlSetState($MacroTimeLoop,$GUI_HIDE)
	GUICtrlSetState($MacroPixelTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroPixelType,$GUI_HIDE)
	GUICtrlSetState($MacroPixelWidthLabel,$GUI_HIDE)
	GUICtrlSetState($MacroPixelWidth,$GUI_HIDE)
	GUICtrlSetState($MacroPixelHeightLabel,$GUI_HIDE)
	GUICtrlSetState($MacroPixelHeight,$GUI_HIDE)
	GUICtrlSetState($MacroPixelIdLabel1,$GUI_HIDE)
	GUICtrlSetState($MacroPixelIdLabel2,$GUI_HIDE)
	GUICtrlSetState($MacroPixelId,$GUI_HIDE)
	GUICtrlSetState($MacroInputTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroInputType,$GUI_HIDE)
	GUICtrlSetState($MacroContinueKeyLabel,$GUI_HIDE)
	GUICtrlSetState($MacroContinueKey,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmType,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmTimeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmTime,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmDateLabel,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmDate,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmSpecificLabel,$GUI_HIDE)
	GUICtrlSetState($MacroAlarmSpecific,$GUI_HIDE)
	GUICtrlSetState($MacroSoundTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroSoundType,$GUI_HIDE)
	GUICtrlSetState($MacroSoundFileLabel,$GUI_HIDE)
	GUICtrlSetState($MacroSoundSelect,$GUI_HIDE)
	GUICtrlSetState($MacroSoundFile,$GUI_HIDE)
	GUICtrlSetState($MacroSoundVolumeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroSoundVolume,$GUI_HIDE)
	GUICtrlSetState($MacroSoundVolumeUpDown,$GUI_HIDE)
	GUICtrlSetState($MacroSoundFrequencyLabel,$GUI_HIDE)
	GUICtrlSetState($MacroSoundFrequency,$GUI_HIDE)
	GUICtrlSetState($MacroSoundDurationLabel,$GUI_HIDE)
	GUICtrlSetState($MacroSoundDuration,$GUI_HIDE)
	GUICtrlSetState($MacroCounterTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCounterType,$GUI_HIDE)
	GUICtrlSetState($MacroCounterNameLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCounterName,$GUI_HIDE)
	GUICtrlSetState($MacroCounterNameList,$GUI_HIDE)
	GUICtrlSetState($MacroCounterSetLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCounterSet,$GUI_HIDE)
	GUICtrlSetState($MacroCounterRandom,$GUI_HIDE)
	GUICtrlSetState($MacroCounterRandomLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCounterValue1aLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCounterValue1bLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCounterValue1cLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCounterValue1,$GUI_HIDE)
	GUICtrlSetState($MacroCounterValue2Label,$GUI_HIDE)
	GUICtrlSetState($MacroCounterValue2,$GUI_HIDE)
	GUICtrlSetState($MacroFileTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroFileType,$GUI_HIDE)
	GUICtrlSetState($MacroFileFileLabel,$GUI_HIDE)
	GUICtrlSetState($MacroFileFile,$GUI_HIDE)
	GUICtrlSetState($MacroFileFileSelect,$GUI_HIDE)
	GUICtrlSetState($MacroFilePathLabel,$GUI_HIDE)
	GUICtrlSetState($MacroFilePath,$GUI_HIDE)
	GUICtrlSetState($MacroFilePathSelect,$GUI_HIDE)
	GUICtrlSetState($MacroFileStampLabel,$GUI_HIDE)
	GUICtrlSetState($MacroFileStamp,$GUI_HIDE)
	GUICtrlSetState($MacroFileStampSelect,$GUI_HIDE)
	GUICtrlSetState($MacroFileChangeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroFileChange,$GUI_HIDE)
	GUICtrlSetState($MacroFileChangeSelect,$GUI_HIDE)
	GUICtrlSetState($MacroControlTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroControlType,$GUI_HIDE)
	GUICtrlSetState($MacroProcessTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessType,$GUI_HIDE)
	GUICtrlSetState($MacroProcessShowProcessCurrentLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessShowProcessCurrent,$GUI_HIDE)
	GUICtrlSetState($MacroProcessShowProcessLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessShowProcess,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWindowTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWindowType,$GUI_HIDE)
	GUICtrlSetState($MacroProcessProgramLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessProgram,$GUI_HIDE)
	GUICtrlSetState($MacroProcessProgramSelect,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWorkDirLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWindowLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWindow,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWorkDir,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWorkDirSelect,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWait,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWaitLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessTest,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWindowStateLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessWindowState,$GUI_HIDE)
	GUICtrlSetState($MacroProcessTimeoutLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessTimeout,$GUI_HIDE)
	GUICtrlSetState($MacroProcessProcessLabel,$GUI_HIDE)
	GUICtrlSetState($MacroProcessProcess,$GUI_HIDE)
	GUICtrlSetState($MacroProcessProcessSelect,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureActionLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureAction,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureXLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureYLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureWidthLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureHeightLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureDetect,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureShow,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureScreenX,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureScreenY,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureScreenWidth,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureScreenHeight,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureWindowX,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureWindowY,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureWindowWidth,$GUI_HIDE)
	GUICtrlSetState($MacroCaptureWindowHeight,$GUI_HIDE)
	GUICtrlSetState($MacroCapturePathLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCapturePath,$GUI_HIDE)
	GUICtrlSetState($MacroCapturePathSelect,$GUI_HIDE)
	GUICtrlSetState($MacroCapturePointer,$GUI_HIDE)
	GUICtrlSetState($MacroCapturePointerLabel,$GUI_HIDE)
	GUICtrlSetState($MacroShutDownActionLabel,$GUI_HIDE)
	GUICtrlSetState($MacroShutDownAction,$GUI_HIDE)
	GUICtrlSetState($MacroShutDownAsk,$GUI_HIDE)
	GUICtrlSetState($MacroShutDownAskLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCommentTypeLabel,$GUI_HIDE)
	GUICtrlSetState($MacroCommentType,$GUI_HIDE)
	Switch _GUICtrlComboBox_GetCurSel($MacroStepType)
		Case $StepTypeMouse
			GUICtrlSetState($MacroMouseButton,$GUI_SHOW)
			GUICtrlSetState($MacroMouseButtonLabel,$GUI_SHOW)
			If _GUICtrlComboBox_GetCurSel($MacroMouseButton) <> 11 Then GUICtrlSetState($MacroMousePosition,$GUI_SHOW)
			If Not _Inlist(_GUICtrlComboBox_GetCurSel($MacroMouseButton),10,11) Then
				GUICtrlSetState($MacroDetectPosition,$GUI_SHOW)
				GUICtrlSetState($MacroLabelX,$GUI_SHOW)
				GUICtrlSetState($MacroPositionX,$GUI_SHOW)
				GUICtrlSetState($MacroLabelY,$GUI_SHOW)
				GUICtrlSetState($MacroPositionY,$GUI_SHOW)
				GUICtrlSetState($MacroShowPosition,$GUI_SHOW)
				GUICtrlSetState($MacroPositionXRandomLabel,$GUI_SHOW)
				GUICtrlSetState($MacroPositionXRandom,$GUI_SHOW)
				GUICtrlSetState($MacroPositionYRandomLabel,$GUI_SHOW)
				GUICtrlSetState($MacroPositionYRandom,$GUI_SHOW)
				GUICtrlSetState($MacroMouseMoveInstantaneous,$GUI_SHOW)
				GUICtrlSetState($MacroMouseMoveInstantaneousLabel,$GUI_SHOW)
				GUICtrlSetState($MacroMouseClicks,$GUI_SHOW)
				GUICtrlSetState($MacroMouseClicksLabel,$GUI_SHOW)
			EndIf
		Case $StepTypeKey
			GUICtrlSetState($MacroKeyLabel,$GUI_SHOW)
			GUICtrlSetState($MacroKey,$GUI_SHOW)
			GUICtrlSetState($MacroKeyModeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroKeyMode,$GUI_SHOW)
			If _GUICtrlComboBox_GetCurSel($MacroKey) = 0 Then
				GUICtrlSetState($MacroTextSendLabel,$GUI_SHOW)
				GUICtrlSetState($MacroTextSend,$GUI_SHOW)
				If _GUICtrlComboBox_GetCurSel($MacroTextSend) = 0 Then		; text
					GUICtrlSetState($MacroTextLabel,$GUI_SHOW)
					GUICtrlSetState($MacroText,$GUI_SHOW)
				EndIf
			EndIf
		Case $StepTypeDelay
			GUICtrlSetState($MacroDelayTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroDelayType,$GUI_SHOW)
			If _GUICtrlComboBox_GetCurSel($MacroDelayType) = 0 Then
				GUICtrlSetState($MacroDelayLabel,$GUI_SHOW)
				GUICtrlSetState($MacroDelay,$GUI_SHOW)
				GUICtrlSetState($MacroDelayRandomLabel,$GUI_SHOW)
				GUICtrlSetState($MacroDelayRandom,$GUI_SHOW)
			Else
				GUICtrlSetState($MacroDelayStepLabel,$GUI_SHOW)
				GUICtrlSetState($MacroDelayStep,$GUI_SHOW)
			EndIf
		Case $StepTypeWindow
			GUICtrlSetState($MacroWindowActionLabel,$GUI_SHOW)
			GUICtrlSetState($MacroWindowAction,$GUI_SHOW)
			Switch _GUICtrlComboBox_GetCurSel($MacroWindowAction)
				Case 3		; move
					GUICtrlSetState($MacroLabelX,$GUI_SHOW)
					GUICtrlSetState($MacroPositionX,$GUI_SHOW)
					GUICtrlSetState($MacroLabelY,$GUI_SHOW)
					GUICtrlSetState($MacroPositionY,$GUI_SHOW)
					GUICtrlSetState($MacroWindowPosition,$GUI_SHOW)
					GUICtrlSetState($MacroShowPosition,$GUI_SHOW)
					GUICtrlSetState($MacroPositionXRandomLabel,$GUI_SHOW)
					GUICtrlSetState($MacroPositionXRandom,$GUI_SHOW)
					GUICtrlSetState($MacroPositionYRandomLabel,$GUI_SHOW)
					GUICtrlSetState($MacroPositionYRandom,$GUI_SHOW)
					GUICtrlSetState($MacroWindowMoveInstantaneous,$GUI_SHOW)
					GUICtrlSetState($MacroWindowMoveInstantaneousLabel,$GUI_SHOW)
				Case 4 		; resize
					GUICtrlSetState($MacroWidthLabel,$GUI_SHOW)
					GUICtrlSetState($MacroWidth,$GUI_SHOW)
					GUICtrlSetState($MacroHeightLabel,$GUI_SHOW)
					GUICtrlSetState($MacroHeight,$GUI_SHOW)
					GUICtrlSetState($MacroWindowSize,$GUI_SHOW)
					GUICtrlSetState($MacroPositionXRandomLabel,$GUI_SHOW)
					GUICtrlSetState($MacroPositionXRandom,$GUI_SHOW)
					GUICtrlSetState($MacroPositionYRandomLabel,$GUI_SHOW)
					GUICtrlSetState($MacroPositionYRandom,$GUI_SHOW)
					GUICtrlSetState($MacroWindowResizeInstantaneous,$GUI_SHOW)
					GUICtrlSetState($MacroWindowResizeInstantaneousLabel,$GUI_SHOW)
				Case 13,14,15,16		; switch, (not) exists
					GUICtrlSetState($MacroWindowWindowLabel,$GUI_SHOW)
					GUICtrlSetState($MacroWindowWindow,$GUI_SHOW)
			EndSwitch
		Case $StepTypeRandom
			GUICtrlSetState($MacroRandomMouseLabel,$GUI_SHOW)
			GUICtrlSetState($MacroRandomMouse,$GUI_SHOW)
			GUICtrlSetState($MacroRandomDelayLabel,$GUI_SHOW)
			GUICtrlSetState($MacroRandomDelay,$GUI_SHOW)
		Case $StepTypeJump
			GUICtrlSetState($MacroJumpTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroJumpType,$GUI_SHOW)
			If _GUICtrlComboBox_GetCurSel($MacroJumpType) = 0 Then		; to label
				GUICtrlSetState($MacroJumpLabelLabel,$GUI_SHOW)
				GUICtrlSetState($MacroJumpLabel,$GUI_SHOW)
			EndIf
		Case $StepTypeRepeat
			GUICtrlSetState($MacroRepeatTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroRepeatType,$GUI_SHOW)
			Switch _GUICtrlComboBox_GetCurSel($MacroRepeatType)
				Case 0		; repeat loop
					GUICtrlSetState($MacroRepeatCountLabel,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatCount,$GUI_SHOW)
				Case 1		; if =
					GUICtrlSetState($MacroRepeatValue1aLabel,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatValue1,$GUI_SHOW)
				Case 2		; if not =
					GUICtrlSetState($MacroRepeatValue1bLabel,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatValue1,$GUI_SHOW)
				Case 3		; if >=
					GUICtrlSetState($MacroRepeatValue1cLabel,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatValue1,$GUI_SHOW)
				Case 4		; if <=
					GUICtrlSetState($MacroRepeatValue2Label,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatValue2,$GUI_SHOW)
				Case 5,6	; if between and not between
					GUICtrlSetState($MacroRepeatValue1cLabel,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatValue1,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatValue2Label,$GUI_SHOW)
					GUICtrlSetState($MacroRepeatValue2,$GUI_SHOW)
			EndSwitch
		Case $StepTypeTime
			GUICtrlSetState($MacroTimeLoopLabel,$GUI_SHOW)
			GUICtrlSetState($MacroTimeLoop,$GUI_SHOW)
		Case $StepTypePixel
			GUICtrlSetState($MacroPixelTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroPixelType,$GUI_SHOW)
			If _InList(_GUICtrlComboBox_GetCurSel($MacroPixelType),0,2) Then	; record
				GUICtrlSetState($MacroLabelX,$GUI_SHOW)
				GUICtrlSetState($MacroPositionX,$GUI_SHOW)
				GUICtrlSetState($MacroLabelY,$GUI_SHOW)
				GUICtrlSetState($MacroPositionY,$GUI_SHOW)
				GUICtrlSetState($MacroDetectPosition,$GUI_SHOW)
				GUICtrlSetState($MacroMousePosition,$GUI_SHOW)
				GUICtrlSetState($MacroShowPosition,$GUI_SHOW)
				If _GUICtrlComboBox_GetCurSel($MacroPixelType) = 0 Then
					GUICtrlSetState($MacroPixelWidthLabel,$GUI_SHOW)
					GUICtrlSetState($MacroPixelWidth,$GUI_SHOW)
					GUICtrlSetState($MacroPixelHeightLabel,$GUI_SHOW)
					GUICtrlSetState($MacroPixelHeight,$GUI_SHOW)
				EndIf
			EndIf
			If _InList(_GUICtrlComboBox_GetCurSel($MacroPixelType),0,1) Then
				GUICtrlSetState($MacroPixelIdLabel1,$GUI_SHOW)
			Else
				GUICtrlSetState($MacroPixelIdLabel2,$GUI_SHOW)
			EndIf
			GUICtrlSetState($MacroPixelId,$GUI_SHOW)
		Case $StepTypeInput
			GUICtrlSetState($MacroInputType,$GUI_SHOW)
			GUICtrlSetState($MacroInputTypeLabel,$GUI_SHOW)
			Switch  _GUICtrlComboBox_GetCurSel($MacroInputType)
				Case 0
					GUICtrlSetState($MacroContinueKeyLabel,$GUI_SHOW)
					GUICtrlSetState($MacroContinueKey,$GUI_SHOW)
			EndSwitch
		Case $StepTypeAlarm
			GUICtrlSetState($MacroAlarmTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroAlarmType,$GUI_SHOW)
			If _GUICtrlComboBox_GetCurSel($MacroAlarmType) = 0 Then		; specific date and time
				GUICtrlSetState($MacroAlarmTimeLabel,$GUI_SHOW)
				GUICtrlSetState($MacroAlarmTime,$GUI_SHOW)
				GUICtrlSetState($MacroAlarmDateLabel,$GUI_SHOW)
				GUICtrlSetState($MacroAlarmDate,$GUI_SHOW)
				GUICtrlSetState($MacroAlarmSpecificLabel,$GUI_SHOW)
				GUICtrlSetState($MacroAlarmSpecific,$GUI_SHOW)
				GUICtrlSetTip($MacroAlarmDate,"Set date to process block of macro steps")
				GUICtrlSetTip($MacroAlarmTime,"Set time to process block of macro steps")
			ElseIf _GUICtrlComboBox_GetCurSel($MacroAlarmType) = 1 Then	; minute
				GUICtrlSetState($MacroAlarmTimeLabel,$GUI_HIDE)
				GUICtrlSetState($MacroAlarmTime,$GUI_HIDE)
				GUICtrlSetState($MacroAlarmDateLabel,$GUI_HIDE)
				GUICtrlSetState($MacroAlarmDate,$GUI_HIDE)
			ElseIf _GUICtrlComboBox_GetCurSel($MacroAlarmType) >= 2 And _GUICtrlComboBox_GetCurSel($MacroAlarmType) <= 3 Then	; hour and daily
				If _GUICtrlComboBox_GetCurSel($MacroAlarmType) = 2 Then GUICtrlSetTip($MacroAlarmTime,"Set minute of hour to process block of macro steps")
				If _GUICtrlComboBox_GetCurSel($MacroAlarmType) = 3 Then GUICtrlSetTip($MacroAlarmTime,"Set time to process block of macro steps")
				GUICtrlSetState($MacroAlarmTimeLabel,$GUI_SHOW)
				GUICtrlSetState($MacroAlarmTime,$GUI_SHOW)
				GUICtrlSetState($MacroAlarmDateLabel,$GUI_HIDE)
				GUICtrlSetState($MacroAlarmDate,$GUI_HIDE)
			EndIf
		Case $StepTypeSound
			GUICtrlSetState($MacroSoundType,$GUI_SHOW)
			GUICtrlSetState($MacroSoundTypeLabel,$GUI_SHOW)
			Switch _GUICtrlComboBox_GetCurSel($MacroSoundType)
				Case 0
					GUICtrlSetState($MacroSoundFrequencyLabel,$GUI_SHOW)
					GUICtrlSetState($MacroSoundFrequency,$GUI_SHOW)
					GUICtrlSetState($MacroSoundDurationLabel,$GUI_SHOW)
					GUICtrlSetState($MacroSoundDuration,$GUI_SHOW)
				Case 1
					GUICtrlSetState($MacroSoundFileLabel,$GUI_SHOW)
					GUICtrlSetState($MacroSoundSelect,$GUI_SHOW)
					GUICtrlSetState($MacroSoundFile,$GUI_SHOW)
				Case 3
					GUICtrlSetState($MacroSoundVolumeLabel,$GUI_SHOW)
					GUICtrlSetState($MacroSoundVolume,$GUI_SHOW)
					GUICtrlSetState($MacroSoundVolumeUpDown,$GUI_SHOW)
			EndSwitch
		Case $StepTypeCounter
			GUICtrlSetState($MacroCounterTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroCounterType,$GUI_SHOW)
			GUICtrlSetState($MacroCounterNameLabel,$GUI_SHOW)
			Switch _GUICtrlComboBox_GetCurSel($MacroCounterType)
				Case 0		; set
					GUICtrlSetState($MacroCounterName,$GUI_SHOW)
					GUICtrlSetState($MacroCounterSetLabel,$GUI_SHOW)
					GUICtrlSetState($MacroCounterSet,$GUI_SHOW)
					GUICtrlSetState($MacroCounterRandom,$GUI_SHOW)
					GUICtrlSetState($MacroCounterRandomLabel,$GUI_SHOW)
				Case 1		; set play
					GUICtrlSetState($MacroCounterName,$GUI_SHOW)
				Case 2		; add
					GUICtrlSetState($MacroCounterNameList,$GUI_SHOW)
					GUICtrlSetState($MacroCounterSetLabel,$GUI_SHOW)
					GUICtrlSetState($MacroCounterSet,$GUI_SHOW)
					GUICtrlSetState($MacroCounterRandom,$GUI_SHOW)
					GUICtrlSetState($MacroCounterRandomLabel,$GUI_SHOW)
				Case 3		; if =
					GUICtrlSetState($MacroCounterNameList,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1aLabel,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1,$GUI_SHOW)
				Case 4		; if not =
					GUICtrlSetState($MacroCounterNameList,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1bLabel,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1,$GUI_SHOW)
				Case 5		; if >=
					GUICtrlSetState($MacroCounterNameList,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1cLabel,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1,$GUI_SHOW)
				Case 6		; <=
					GUICtrlSetState($MacroCounterNameList,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue2Label,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue2,$GUI_SHOW)
				Case 7,8	; between and not between
					GUICtrlSetState($MacroCounterNameList,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1cLabel,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue1,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue2Label,$GUI_SHOW)
					GUICtrlSetState($MacroCounterValue2,$GUI_SHOW)
				Case 9,10	; odd and even
					GUICtrlSetState($MacroCounterNameList,$GUI_SHOW)
			EndSwitch
		Case $StepTypeFile
			GUICtrlSetState($MacroFileTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroFileType,$GUI_SHOW)
			Switch _GUICtrlComboBox_GetCurSel($MacroFileType)
				Case 0,1	; check file
					GUICtrlSetState($MacroFileFileLabel,$GUI_SHOW)
					GUICtrlSetState($MacroFileFile,$GUI_SHOW)
					GUICtrlSetState($MacroFileFileSelect,$GUI_SHOW)
				Case 2,3	; check path
					GUICtrlSetState($MacroFilePathLabel,$GUI_SHOW)
					GUICtrlSetState($MacroFilePath,$GUI_SHOW)
					GUICtrlSetState($MacroFilePathSelect,$GUI_SHOW)
				Case 4		; record change
					GUICtrlSetState($MacroFileStampLabel,$GUI_SHOW)
					GUICtrlSetState($MacroFileStamp,$GUI_SHOW)
					GUICtrlSetState($MacroFileStampSelect,$GUI_SHOW)
				Case 5,6	; check changed
					GUICtrlSetState($MacroFileChangeLabel,$GUI_SHOW)
					GUICtrlSetState($MacroFileChange,$GUI_SHOW)
					GUICtrlSetState($MacroFileChangeSelect,$GUI_SHOW)
			EndSwitch
		Case $StepTypeControl
			GUICtrlSetState($MacroControlTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroControlType,$GUI_SHOW)
		Case $StepTypeProcess
			GUICtrlSetState($MacroProcessTypeLabel,$GUI_SHOW)
			GUICtrlSetState($MacroProcessType,$GUI_SHOW)
			Switch _GUICtrlComboBox_GetCurSel($MacroProcessType)
				Case 0		; select process
					GUICtrlSetState($MacroProcessShowProcessCurrentLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessShowProcessCurrent,$GUI_SHOW)
					GUICtrlSetState($MacroProcessShowProcessLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessShowProcess,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWindowTypeLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWindowType,$GUI_SHOW)
					If _GUICtrlComboBox_GetCurSel($MacroProcessWindowType) = 1 Then
						GUICtrlSetState($MacroProcessWindowLabel,$GUI_SHOW)
						GUICtrlSetState($MacroProcessWindow,$GUI_SHOW)
					EndIf
				Case 1		; start program
					GUICtrlSetState($MacroProcessProgramLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessProgram,$GUI_SHOW)
					GUICtrlSetState($MacroProcessProgramSelect,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWindowLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWindow,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWorkDirLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWorkDir,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWorkDirSelect,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWait,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWaitLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessTest,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWindowStateLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessWindowState,$GUI_SHOW)
					GUICtrlSetState($MacroProcessTimeoutLabel,GUICtrlRead($MacroProcessWait) = $GUI_CHECKED ? $GUI_HIDE : $GUI_SHOW)
					GUICtrlSetState($MacroProcessTimeout,GUICtrlRead($MacroProcessWait) = $GUI_CHECKED ? $GUI_HIDE : $GUI_SHOW)
				Case 2,3		; process (not) exists
					GUICtrlSetState($MacroProcessProcessLabel,$GUI_SHOW)
					GUICtrlSetState($MacroProcessProcess,$GUI_SHOW)
					GUICtrlSetState($MacroProcessProcessSelect,$GUI_SHOW)
			EndSwitch
		Case $StepTypeCapture
			GUICtrlSetState($MacroCaptureActionLabel,$GUI_SHOW)
			GUICtrlSetState($MacroCaptureAction,$GUI_SHOW)
			If _InList(_GUICtrlComboBox_GetCurSel($MacroCaptureAction),1,3) Then
				GUICtrlSetState($MacroCaptureXLabel,$GUI_SHOW)
				GUICtrlSetState($MacroCaptureYLabel,$GUI_SHOW)
				GUICtrlSetState($MacroCaptureWidthLabel,$GUI_SHOW)
				GUICtrlSetState($MacroCaptureHeightLabel,$GUI_SHOW)
				GUICtrlSetState($MacroCaptureDetect,$GUI_SHOW)
				GUICtrlSetState($MacroCaptureShow,$GUI_SHOW)
			EndIf
			Switch _GUICtrlComboBox_GetCurSel($MacroCaptureAction)
				Case 1
					GUICtrlSetState($MacroCaptureScreenX,$GUI_SHOW)
					GUICtrlSetState($MacroCaptureScreenY,$GUI_SHOW)
					GUICtrlSetState($MacroCaptureScreenWidth,$GUI_SHOW)
					GUICtrlSetState($MacroCaptureScreenHeight,$GUI_SHOW)
				Case 3
					GUICtrlSetState($MacroCaptureWindowX,$GUI_SHOW)
					GUICtrlSetState($MacroCaptureWindowY,$GUI_SHOW)
					GUICtrlSetState($MacroCaptureWindowWidth,$GUI_SHOW)
					GUICtrlSetState($MacroCaptureWindowHeight,$GUI_SHOW)
				Case 5
					GUICtrlSetState($MacroCapturePathLabel,$GUI_SHOW)
					GUICtrlSetState($MacroCapturePath,$GUI_SHOW)
					GUICtrlSetState($MacroCapturePathSelect,$GUI_SHOW)
					GUICtrlSetState($MacroCapturePointer,$GUI_SHOW)
					GUICtrlSetState($MacroCapturePointerLabel,$GUI_SHOW)
			EndSwitch
		Case $StepTypeShutdown
			GUICtrlSetState($MacroShutDownActionLabel,$GUI_SHOW)
			GUICtrlSetState($MacroShutDownAction,$GUI_SHOW)
			GUICtrlSetState($MacroShutDownAsk,$GUI_SHOW)
			GUICtrlSetState($MacroShutDownAskLabel,$GUI_SHOW)
		Case $StepTypeComment
			GUICtrlSetState($MacroCommentType,$GUI_SHOW)
			GUICtrlSetState($MacroCommentTypeLabel,$GUI_SHOW)
	EndSwitch
EndFunc

Func SetControlStates($Enable = True)
	If $Enable Then
		GUICtrlSetState($ProcessesLabel,$GUI_ENABLE)
		GUICtrlSetState($Processes,$GUI_ENABLE)
		GUICtrlSetState($WindowsVisible,$GUI_ENABLE)
		GUICtrlSetState($WindowsInvisible,$GUI_ENABLE)
		GUICtrlSetState($StepViewNumbers,$GUI_ENABLE)
		GUICtrlSetState($StepViewInverseColor,$GUI_ENABLE)
		GUICtrlSetState($StepViewList,$GUI_ENABLE)
		GUICtrlSetState($StepViewReport,$GUI_ENABLE)
		GUICtrlSetState($StepViewGrid,$GUI_ENABLE)
		GUICtrlSetState($Add,$GUI_ENABLE)
		GUICtrlSetState($Change,$GUI_ENABLE)
		GUICtrlSetState($Duplicate,$GUI_ENABLE)
		GUICtrlSetState($Delete,$GUI_ENABLE)
		GUICtrlSetState($Up,$GUI_ENABLE)
		GUICtrlSetState($Down,$GUI_ENABLE)
		GUICtrlSetState($Show,$GUI_ENABLE)
		GUICtrlSetState($Enable,$GUI_ENABLE)
		GUICtrlSetState($Disable,$GUI_ENABLE)
		GUICtrlSetState($MoveTimeLabel,$GUI_ENABLE)
		GUICtrlSetState($MoveTime,$GUI_ENABLE)
		GUICtrlSetState($ShowRecording,$GUI_ENABLE)
		GUICtrlSetState($Record,$GUI_ENABLE)
		GUICtrlSetState($Play,$GUI_ENABLE)
		GUICtrlSetState($LoopCountLabel,$GUI_ENABLE)
		GUICtrlSetState($LoopCount,$GUI_ENABLE)
		GUICtrlSetState($StepDelayLabel,$GUI_ENABLE)
		GUICtrlSetState($StepDelay,$GUI_ENABLE)
		GUICtrlSetState($StepDelayAlways,$GUI_ENABLE)
		GUICtrlSetState($ShowMain,$GUI_ENABLE)
		GUICtrlSetState($ShowBar,$GUI_ENABLE)
		GUICtrlSetState($TestMacro,$GUI_ENABLE)
		GUICtrlSetState($StartStopKey,$GUI_ENABLE)
		GUICtrlSetState($PauseKey,$GUI_ENABLE)
		GUICtrlSetState($New,$GUI_ENABLE)
		GUICtrlSetState($Open,$GUI_ENABLE)
		GUICtrlSetState($Save,$GUI_ENABLE)
		GUICtrlSetState($SaveAs,$GUI_ENABLE)
		GUICtrlSetState($Close,$GUI_ENABLE)
		GUICtrlSetState($Help,$GUI_ENABLE)
		GUICtrlSetState($Settings,$GUI_ENABLE)
		EnableDisableButtons()
	Else
		GUICtrlSetState($ProcessesLabel,$GUI_DISABLE)
		GUICtrlSetState($Processes,$GUI_DISABLE)
		GUICtrlSetState($WindowsVisible,$GUI_DISABLE)
		GUICtrlSetState($WindowsInvisible,$GUI_DISABLE)
		GUICtrlSetState($StepViewNumbers,$GUI_DISABLE)
		GUICtrlSetState($StepViewInverseColor,$GUI_DISABLE)
		GUICtrlSetState($StepViewList,$GUI_DISABLE)
		GUICtrlSetState($StepViewReport,$GUI_DISABLE)
		GUICtrlSetState($StepViewGrid,$GUI_DISABLE)
		GUICtrlSetState($Add,$GUI_DISABLE)
		GUICtrlSetState($Change,$GUI_DISABLE)
		GUICtrlSetState($Duplicate,$GUI_DISABLE)
		GUICtrlSetState($Delete,$GUI_DISABLE)
		GUICtrlSetState($Up,$GUI_DISABLE)
		GUICtrlSetState($Down,$GUI_DISABLE)
		GUICtrlSetState($Show,$GUI_DISABLE)
		GUICtrlSetState($Enable,$GUI_DISABLE)
		GUICtrlSetState($Disable,$GUI_DISABLE)
		GUICtrlSetState($MoveTimeLabel,$GUI_DISABLE)
		GUICtrlSetState($MoveTime,$GUI_DISABLE)
		GUICtrlSetState($ShowRecording,$GUI_DISABLE)
		GUICtrlSetState($Record,$GUI_DISABLE)
		GUICtrlSetState($Play,$GUI_DISABLE)
		GUICtrlSetState($LoopCountLabel,$GUI_DISABLE)
		GUICtrlSetState($LoopCount,$GUI_DISABLE)
		GUICtrlSetState($StepDelayLabel,$GUI_DISABLE)
		GUICtrlSetState($StepDelay,$GUI_DISABLE)
		GUICtrlSetState($StepDelayAlways,$GUI_DISABLE)
		GUICtrlSetState($ShowMain,$GUI_DISABLE)
		GUICtrlSetState($ShowBar,$GUI_DISABLE)
		GUICtrlSetState($TestMacro,$GUI_DISABLE)
		GUICtrlSetState($StartStopKey,$GUI_DISABLE)
		GUICtrlSetState($PauseKey,$GUI_DISABLE)
		GUICtrlSetState($New,$GUI_DISABLE)
		GUICtrlSetState($Open,$GUI_DISABLE)
		GUICtrlSetState($Save,$GUI_DISABLE)
		GUICtrlSetState($SaveAs,$GUI_DISABLE)
		GUICtrlSetState($Close,$GUI_DISABLE)
		GUICtrlSetState($Help,$GUI_DISABLE)
		GUICtrlSetState($Settings,$GUI_DISABLE)
	EndIf
EndFunc

Func RecordMacro()								; record macro by recording mouse clicks and position and by recording pressed keys
	Local $StopRecording = False,$Window = $ProcessList[SelectedProcessIndex()][2],$NoMoveTimer,$MoveTimer,$MoveX,$MoveY
	Local $LeftMouse = False,$RightMouse = False,$MiddleMouse = False,$Break = False,$ShiftKey = False,$ControlKey = False,$AltKey = False,$AlphaKey
	Local $Keys = [["30",0,"0",")"],["31",0,"1","!"],["32",0,"2","@"],["33",0,"3","#"],["34",0,"4","$"],["35",0,"5","%"],["36",0,"6","^"],["37",0,"7","&"],["38",0,"8","*"],["39",0,"9","("], _
	["BA",0,";",":"],["BB",0,"=","+"],["BC",0,",","<"],["BD",0,"-","_"],["BE",0,".",">"],["BF",0,"/","?"],["C0",0,"`","~"],["DB",0,"[","{"],["DC",0,"\","|"],["DD",0,"]","}"],["DE",0,"'",'"']]

	If $MacroStep > -1 Then $Window = $MacroSteps[$MacroStep][$MacroStepWindowColumn]
	; initialize keys in range
	For $Key = 1 To 67							; uptil 67 are normal keys, the rest are modifiers
		$MacroKeys[$Key][3] = False
	Next
	ReDim $Keys[UBound($Keys)+26][4]
	$AlphaKey = 0
	For $Key = 21 To 46	; a - z
		$Keys[$Key][0] = Hex(0x41+$AlphaKey,2)
		$Keys[$Key][1] = False
		$Keys[$Key][2] = Chr(0x61+$AlphaKey)
		$Keys[$Key][3] = Chr(0x41+$AlphaKey)
		$AlphaKey += 1
	Next
	SetControlStates(False)
	SetStopKey()
	WinActivate($Window)
	If Number(GUICtrlRead($MoveTime)) = 0 Then
		$NoMoveTimer = True
	Else
		$NoMoveTimer = False
		$MoveTimer = _Timer_Init()
		$MoveX = DetectCalculateX($Window,0)
		$MoveY = DetectCalculateY($Window,0)
	EndIf
	While $Stopped = 0
		If Not $HideRecordToolTip Then ToolTip($StartStopKeys[_GUICtrlComboBox_GetCurSel($StartStopKey)][0] & " to stop recording")
		If Not $ShiftKey And _IsPressed("10",$User32DLL) Then
			$ShiftKey = True
		ElseIf $ShiftKey And Not _IsPressed("10",$User32DLL) Then
			$ShiftKey = False
		EndIf
		If Not $ControlKey And _IsPressed("11",$User32DLL) Then
			$ControlKey = True
		ElseIf $ControlKey And Not _IsPressed("11",$User32DLL) Then
			$ControlKey = False
		EndIf
		If Not $AltKey And _IsPressed("12",$User32DLL) Then
			$AltKey = True
		ElseIf $AltKey And Not _IsPressed("12",$User32DLL) Then
			$AltKey = False
		EndIf
		For $Key = 0 To UBound($Keys)-1
			If $Keys[$Key][1] = 0 And _IsPressed($Keys[$Key][0],$User32DLL) Then
				$Keys[$Key][1] = $ShiftKey ? 2 : 1
			ElseIf $Keys[$Key][1] = 1 And Not _IsPressed($Keys[$Key][0],$User32DLL) Then
				$Keys[$Key][1] = 0
				If $MacroStep > -1 And $MacroSteps[$MacroStep][0] = 1 And $MacroSteps[$MacroStep][2] = 0 Then
					If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
						EditMacroStep(False,1,"",0,0,$MacroSteps[$MacroStep][4] & $Keys[$Key][2])
					Else
						AddChangeMacroStep(False,1,"",0,0,$MacroSteps[$MacroStep][4] & $Keys[$Key][2])
					EndIf
				Else
					If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
						EditMacroStep(True,1,"",0,0,$Keys[$Key][2])
					Else
						AddChangeMacroStep(True,1,"",0,0,$Keys[$Key][2])
					EndIf
				EndIf
			ElseIf $Keys[$Key][1] = 2 And Not _IsPressed($Keys[$Key][0],$User32DLL) Then
				$Keys[$Key][1] = 0
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(True,1,"",0,0,$Keys[$Key][3])
				Else
					AddChangeMacroStep(True,1,"",0,0,$Keys[$Key][3])
				EndIf
			EndIf
		Next
		For $Key = 1 To 67						; uptil 67 are normal keys, the rest are modifiers
			If Not $MacroKeys[$Key][3] And _IsPressed($MacroKeys[$Key][2],$User32DLL) Then
				$MacroKeys[$Key][3] = True
			ElseIf $MacroKeys[$Key][3] And Not _IsPressed($MacroKeys[$Key][2],$User32DLL) Then
				$MacroKeys[$Key][3] = False
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(True,1,"",0,$Key,"")
				Else
					AddChangeMacroStep(True,1,"",0,$Key,"")
				EndIf
			EndIf
		Next
		; left mouse click recording
		If Not $LeftMouse And _IsPressed("01",$User32DLL) Then
			$LeftMouse = True
			If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
				EditMacroStep(True,0,"",4,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
			Else
				AddChangeMacroStep(True,0,"",4,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
			EndIf
		ElseIf $LeftMouse And Not _IsPressed("01",$User32DLL) Then
			$LeftMouse = False
			If $MacroStep > -1 And $MacroSteps[$MacroStep][0] = 0 And $MacroSteps[$MacroStep][2] = 4 And _
			   $MacroSteps[$MacroStep][3] = DetectCalculateX($Window,0) And $MacroSteps[$MacroStep][4] = DetectCalculateY($Window,0) And _
			   $MacroSteps[$MacroStep][5] = 0 And $MacroSteps[$MacroStep][6] = 0 And $MacroSteps[$MacroStep][7] = 0 Then
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(False,0,"",0,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				Else
					AddChangeMacroStep(False,0,"",0,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				EndIf
			Else
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(True,0,"",7,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				Else
					AddChangeMacroStep(True,0,"",7,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				EndIf
			EndIf
		EndIf
		; right mouse click recording
		If Not $RightMouse And _IsPressed("02",$User32DLL) Then
			$RightMouse = True
			If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
				EditMacroStep(True,0,"",5,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
			Else
				AddChangeMacroStep(True,0,"",5,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
			EndIf
		ElseIf $RightMouse And Not _IsPressed("02",$User32DLL) Then
			$RightMouse = False
			If $MacroStep > -1 And $MacroSteps[$MacroStep][0] = 0 And $MacroSteps[$MacroStep][2] = 5 And _
			   $MacroSteps[$MacroStep][3] = DetectCalculateX($Window,0) And $MacroSteps[$MacroStep][4] = DetectCalculateY($Window,0) And _
			   $MacroSteps[$MacroStep][5] = 0 And $MacroSteps[$MacroStep][6] = 0 And $MacroSteps[$MacroStep][7] = 0 Then
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(False,0,"",0,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				Else
					AddChangeMacroStep(False,0,"",0,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				EndIf
			Else
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(True,0,"",8,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				Else
					AddChangeMacroStep(True,0,"",8,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				EndIf
			EndIf
		EndIf
		; middle mouse click recording
		If Not $MiddleMouse And _IsPressed("04",$User32DLL) Then
			$MiddleMouse = True
			If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
				EditMacroStep(True,0,"",6,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
			Else
				AddChangeMacroStep(True,0,"",6,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
			EndIf
		ElseIf $MiddleMouse And Not _IsPressed("04",$User32DLL) Then
			$MiddleMouse = False
			If $MacroStep > -1 And $MacroSteps[$MacroStep][0] = 0 And $MacroSteps[$MacroStep][2] = 6 And _
			   $MacroSteps[$MacroStep][3] = DetectCalculateX($Window,0) And $MacroSteps[$MacroStep][4] = DetectCalculateY($Window,0) And _
			   $MacroSteps[$MacroStep][5] = 0 And $MacroSteps[$MacroStep][6] = 0 And $MacroSteps[$MacroStep][7] = 0 Then
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(False,0,"",0,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				Else
					AddChangeMacroStep(False,0,"",0,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				EndIf
			Else
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(True,0,"",9,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				Else
					AddChangeMacroStep(True,0,"",9,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				EndIf
			EndIf
		EndIf
		If Not $NoMoveTimer And _Timer_Diff($MoveTimer) > Number(GUICtrlRead($MoveTime)) Then
			$MoveTimer = _Timer_Init()
			If $MoveX <> DetectCalculateX($Window,0) Or $MoveY <> DetectCalculateY($Window,0) Then
				$MoveX = DetectCalculateX($Window,0)
				$MoveY = DetectCalculateY($Window,0)
				If GUICtrlRead($ShowRecording) = $GUI_CHECKED Then
					EditMacroStep(True,0,"",3,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				Else
					AddChangeMacroStep(True,0,"",3,DetectCalculateX($Window,0),DetectCalculateY($Window,0),0,0,0)
				EndIf
			EndIf
		EndIf
		Sleep(10)	; delay against CPU clogging
	WEnd
	If Not $HideRecordToolTip Then ToolTip("")
	SetStopKey(False)
	GUISetState(@SW_RESTORE,$GUI)
	GUISetState(@SW_SHOWNORMAL,$GUI)
	SetControlStates()
	If GUICtrlRead($ShowRecording) = $GUI_UNCHECKED Then
		FillMacroSteps($MacroStep)
		Changed()
	EndIf
EndFunc

Func PlayMacro()								; play macro by processing macro steps
	Const $BarWidth = 200,$BarHeight = 15
	Local $GUIBar,$PlayBar,$PlayBarPercentage,$Delay,$DelayTimer,$DelayTime,$Window,$Initialize,$ProcessInitalize,$Count,$PerformDelay,$DelayInStep,$MouseStack,$MouseX,$MouseY,$MousePosition,$RandomMouse,$RandomDelay,$JumpStep,$BlockStep,$JumpBlock, _
		  $Sound,$RepeatStep,$Level,$PixelCheck,$Pixels[1][4],$PixelX,$PixelY,$PixelAdded,$MousePixel,$CounterValue,$Counters[1][3],$CounterAdded,$FileStamp,$FileStamps[1][5],$StampAdded,$WindowFound,$TimeOutTimer,$FindWindow, _
		  $Capture,$CapturePath,$CaptureLeft,$CaptureTop,$CaptureRight,$CaptureBottom,$CapturePoint,$ProcessId,$ProcessCheck,$ActiveProcess,$DelayStep,$AlwaysDelay

	If Not FillMacroSteps($MacroStep,True) And Not _Sure("Some issues has been found. Still run macro?") Then Return

	SetControlStates(False)
	SRandom((@HOUR+1)*(@MIN+1)*(@SEC+1))
	$MouseStack = _StackCreate()
	Opt("SendKeyDelay",Number(GUICtrlRead($StepDelay)))
	Opt("SendKeyDownDelay",Number(GUICtrlRead($StepDelay)))
	Opt("MouseClickDelay",Number(GUICtrlRead($StepDelay)))
	Opt("MouseClickDownDelay",Number(GUICtrlRead($StepDelay)))
	$AlwaysDelay = GUICtrlRead($StepDelayAlways) = $GUI_CHECKED
	; pre-process some maro step types
	For $Step = 0 To UBound($MacroSteps)-1
		Switch $MacroSteps[$Step][0]
			Case $StepTypeAlarm
				; set alarms
				$MacroSteps[$Step][6] = False
				Switch $MacroSteps[$Step][2]
					Case 1	; on the minute
						$MacroSteps[$Step][3] = Floor(CurrentTime()/100)*100
						$MacroSteps[$Step][4] = CurrentDate()
						$MacroSteps[$Step][6] = True
					Case 2	; hourly
						$MacroSteps[$Step][5] = AddTime(Floor(CurrentTime()/10000)*10000,$MacroSteps[$Step][3])
						$MacroSteps[$Step][4] = CurrentDate()
						$MacroSteps[$Step][6] = True
					Case 3	; daily
						$MacroSteps[$Step][4] = CurrentDate()
				EndSwitch
			Case $StepTypeSound
				If $MacroSteps[$Step][2] = 1 Then	; load sound file
					If StringLen($MacroSteps[$Step][3]) = 0 Then
						_Ok("No sound file is selected.")
						$MacroSteps[$Step][4] = False
					ElseIf Not FileExists(FilePathTranslateVariables($MacroSteps[$Step][3])) Then
						_Ok("Sound file " & $MacroSteps[$Step][3] & " doesn't exist.")
						$MacroSteps[$Step][4] = False
					Else
						$Sound = _SoundOpen(FilePathTranslateVariables($MacroSteps[$Step][3]))
						If @error Then
							If @extended <> 0 Then
								Local $iExtended = @extended ; Assign because @extended will be set after DllStructCreate().
								Local $tText = DllStructCreate("char[128]")
								DllCall("winmm.dll","short","mciGetErrorStringA","str",$iExtended,"struct*",$tText,"int",128)
								_Ok("Sound file " & $MacroSteps[$Step][3] & " couldn't be opened." & @CRLF & "Error number: " & $iExtended & @CRLF & "Error description: " & DllStructGetData($tText,1) & @CRLF & "Note: The sound may still play correctly.")
							EndIf
							$MacroSteps[$Step][4] = False
						Else
							$MacroSteps[$Step][4] = True
							$MacroSteps[$Step][5] = $Sound[0]
							$MacroSteps[$Step][6] = $Sound[1]
							$MacroSteps[$Step][7] = $Sound[2]
							$MacroSteps[$Step][8] = $Sound[3]
						EndIf
					EndIf
				EndIf
		EndSwitch
	Next

	Switch _GUICtrlComboBox_GetCurSel($ShowMain)
		Case 1
			GUISetState(@SW_MINIMIZE,$GUI)
		Case 2
			GUISetState(@SW_HIDE,$GUI)
	EndSwitch
	If GUICtrlRead($ShowBar) = $GUI_CHECKED Then
		$GUIBar = GUICreate("",$BarWidth,$BarHeight,0,0,$WS_POPUP+$WS_DLGFRAME,$WS_EX_TOPMOST)
		$PlayBar = GUICtrlCreateLabel("",0,0,0,$BarHeight)
		GUICtrlSetBkColor(-1,0x80B0FF)
		$PlayBarPercentage = GUICtrlCreateLabel("0%  ",0,1)
		GUICtrlSetFont(-1,8.5,$FW_SEMIBOLD)
		GUICtrlSetColor(-1,0xFFFFFF)
		GUICtrlSetBkColor(-1,$GUI_BKCOLOR_TRANSPARENT)
		GUISetState(@SW_SHOWNOACTIVATE)
	EndIf
	$Count = 1
	HighlightStep(-1,$MacroStep)
	$PreviousStep = -1
	SetStopKey()
	SetPauseKey()
	While 1										; process macro steps play loop
		$Initialize = True
		$ProcessInitalize = True
		For $Step = 0 To UBound($MacroSteps)-1	; process maco steps of list
			If $MacroSteps[$Step][$MacroStepEnableDisableColumn] = 1 Then ContinueLoop	; step disabled go to next step
			$PerformDelay = False
			$DelayInStep = False
			$DelayTimer = _Timer_Init()
			If $Initialize Then	; $Step will also be 0 at a restart step -> initialize again
				$Initialize = False
				$Delay = Number(GUICtrlRead($StepDelay))
				$ProcessCheck = True
				$RandomMouse = 0
				$RandomDelay = 0
				$PixelCheck = 0
				$FileStamp = 0
				$CounterValue = 0
				$Capture = 0
				$CaptureTop = 0
				$CaptureLeft = 0
				$CaptureRight = -1
				$CaptureBottom = -1
				$CapturePath = ""
				$CapturePoint = True
			EndIf
			If $ProcessInitalize And (_InList($MacroSteps[$Step][0],$StepTypeMouse,$StepTypeKey,$StepTypePixel,$StepTypeInput,$StepTypeCounter) Or _
			   ($MacroSteps[$Step][0] = $StepTypeWindow And $MacroSteps[$Step][2] <> 16) Or _
			   ($MacroSteps[$Step][0] = $StepTypeCapture And _InList($MacroSteps[$Step][2],0,1,2,3))) Then
				$ProcessInitalize = False
				$Window = $MacroSteps[$Step][$MacroStepWindowColumn]
				$ProcessId = $MacroSteps[$Step][$MacroStepWindowColumn+1]
				WinActivate($Window)
				If WinWaitActive($Window,"",$ActivateTimeOut/1000) = 0 Then
					_Ok("Macro can't be initiated." & @CRLF & @CRLF & "Window " & WinGetTitle($Window) & " couldn't be activated in " & $ActivateTimeOut & " ms.")
					$Stopped = 2
					ExitLoop
				Else
					Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
					SendKeepActive($Window)
				EndIf
			EndIf
			ToolTip($HidePlayToolTip ? "" : $StartStopKeys[_GUICtrlComboBox_GetCurSel($StartStopKey)][0] & " to stop macro")
			; highlight current step
			_GUICtrlListView_BeginUpdate($Macros)
			HighlightStep($Step,$PreviousStep)
			$PreviousStep = $Step
			_GUICtrlListView_EnsureVisible($Macros,$Step)
			_GUICtrlListView_EndUpdate($Macros)
			_GUICtrlListView_SetItemSelected($Macros,$Step,True,True)
			; process current macro step
			Switch $MacroSteps[$Step][0]
				Case $StepTypeMouse
					$MouseX = CalculateX($Window,$Step)+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse)
					$MouseY = CalculateY($Window,$Step)+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse)
					Switch $MacroSteps[$Step][2]
						Case 0,1,2		; left, right, middle click
							If GUICtrlRead($TestMacro) = $GUI_UNCHECKED Then
								$MacroSteps[$Step][9] = Number(_Maximum(1,$MacroSteps[$Step][9]))
								For $Click = 1 To $MacroSteps[$Step][9]
									MouseClick($MacroMouseButtons[$MacroSteps[$Step][2]][2],$MouseX,$MouseY,1,Number($MacroSteps[$Step][8]) = 1 ? 0 : _Clamp($Delay,1,100))
								Next
							Else
								MouseMove($MouseX,$MouseY,Number($MacroSteps[$Step][8]) = 1 ? 0 : _Clamp($Delay,1,100))
							EndIf
							$DelayInStep = True
						Case 3			; mouse move
							MouseMove($MouseX,$MouseY,Number($MacroSteps[$Step][8]) = 1 ? 0 : _Clamp($Delay-9,1,100))	; advised is 10 ms but mouse movement is to slow so subtract 9
							$DelayInStep = True
						Case 4,5,6		; left, right, middle down
							MouseMove($MouseX,$MouseY,Number($MacroSteps[$Step][8]) = 1 ? 0 : _Clamp($Delay-9,1,100))	; advised is 10 ms but mouse movement is to slow so subtract 9
							If GUICtrlRead($TestMacro) = $GUI_UNCHECKED Then MouseDown($MacroMouseButtons[$MacroSteps[$Step][2]][2])
							$DelayInStep = True
						Case 7,8,9		; left, right, middle up
							MouseMove($MouseX,$MouseY,Number($MacroSteps[$Step][8]) = 1 ? 0 : _Clamp($Delay-9,1,100))	; advised is 10 ms but mouse movement is to slow so subtract 9
							If GUICtrlRead($TestMacro) = $GUI_UNCHECKED Then MouseUp($MacroMouseButtons[$MacroSteps[$Step][2]][2])
							$DelayInStep = True
						Case 10
							_StackPush($MouseStack,MouseGetPos(0)-CalculateX($Window,-1,0,$MacroSteps[$Step][5]))
							_StackPush($MouseStack,MouseGetPos(1)-CalculateY($Window,-1,0,$MacroSteps[$Step][5]))
							_StackPush($MouseStack,$MacroSteps[$Step][5])
						Case 11
							If Not _StackIsEmpty($MouseStack) Then
								$MousePosition = _StackPop($MouseStack)
								$MouseY = _StackPop($MouseStack)
								$MouseX = _StackPop($MouseStack)
								$MouseX = CalculateX($Window,-1,$MouseX,$MousePosition)+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse)
								$MouseY = CalculateY($Window,-1,$MouseY,$MousePosition)+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse)
								MouseMove($MouseX,$MouseY,_Clamp($Delay-9,1,100))	; advised is 10 ms but mouse movement is to slow so subtract 9
								$DelayInStep = True
							EndIf
					EndSwitch
				Case $StepTypeKey
					If GUICtrlRead($TestMacro) = $GUI_UNCHECKED Then
						If $MacroSteps[$Step][3] = 0 Then
							If Number($MacroSteps[$Step][5]) = 0 Then
								Send($MacroKeyModes[$MacroSteps[$Step][2]][1] & $MacroSteps[$Step][4])
							Else
								Send($MacroKeyModes[$MacroSteps[$Step][2]][1] & RandomCharacter(Number($MacroSteps[$Step][5])))
							EndIf
						Else
							Send($MacroKeyModes[$MacroSteps[$Step][2]][1] & $MacroKeys[$MacroSteps[$Step][3]][1])
						EndIf
					EndIf
					$PerformDelay = True
				Case $StepTypeDelay
					If $MacroSteps[$Step][2] = 0 Then
						If $MacroSteps[$Step][4] > 0 Then
							Sleep(Number($MacroSteps[$Step][3])+Random(0,Number($MacroSteps[$Step][4]),1))
						Else
							Sleep(Number($MacroSteps[$Step][3]))
						EndIf
						$DelayInStep = True
					Else
						$Delay = Number($MacroSteps[$Step][3])
						Opt("SendKeyDelay",$Delay)
						Opt("SendKeyDownDelay",$Delay)
						Opt("MouseClickDelay",$Delay)
						Opt("MouseClickDownDelay",$Delay)
					EndIf
				Case $StepTypeWindow
					Switch $MacroSteps[$Step][2]
						Case 0	; maximize window
							WinSetState($Window,"",@SW_MAXIMIZE)
							$PerformDelay = True
						Case 1	; restore window
							WinSetState($Window,"",@SW_RESTORE)
							$PerformDelay = True
						Case 2	; unmaximize window
							WinSetState($Window,"",@SW_MINIMIZE)
							$PerformDelay = True
						Case 3	; move window to position
							If $MacroSteps[$Step][5] = 0 Then
								; note: coordinates difference between AutoIt and Desktop Window Manager is taken into account
								If $MacroSteps[$Step][8] = 0 Then
									WinMove($Window,"",_WindowGetX($Window)-_WindowGetX($Window,True)+$MacroSteps[$Step][3]+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse),_WindowGetY($Window)-_WindowGetY($Window,True)+$MacroSteps[$Step][4]+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse),WindowWidth($Window),WindowHeight($Window),_Clamp($Delay-9,1,100))
								Else
									WinMove($Window,"",_WindowGetX($Window)-_WindowGetX($Window,True)+$MacroSteps[$Step][3]+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse),_WindowGetY($Window)-_WindowGetY($Window,True)+$MacroSteps[$Step][4]+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse))
								EndIf
							ElseIf $MacroSteps[$Step][5] = 1 Then
								; use AutoIt coordinates instead of Desktop Window Manager ones as WinMove uses AutoIt (wrong) coordinates
								If $MacroSteps[$Step][8] = 0 Then
									WinMove($Window,"",_WindowGetX($Window)+$MacroSteps[$Step][3]+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse),_WindowGetY($Window)+$MacroSteps[$Step][4]+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse),WindowWidth($Window),WindowHeight($Window),_Clamp($Delay-9,1,100))
								Else
									WinMove($Window,"",_WindowGetX($Window)+$MacroSteps[$Step][3]+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse),_WindowGetY($Window)+$MacroSteps[$Step][4]+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse))
								EndIf
							EndIf
						Case 4	; resize window
							If $MacroSteps[$Step][5] = 0 Then
								WinMove($Window,"",Default,Default,$MacroSteps[$Step][3]+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse),$MacroSteps[$Step][4]+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse),_Clamp($Delay-9,1,100))
							ElseIf $MacroSteps[$Step][5] = 1 Then
								; Use AutoIt coordinates instead of Desktop Window Manager ones as WinMove uses AutoIt (wrong) coordinates
								WinMove($Window,"",Default,Default,WindowWidth($Window)+$MacroSteps[$Step][3]+RandomMinus($MacroSteps[$Step][6])+RandomMinus($RandomMouse),WindowHeight($Window)+$MacroSteps[$Step][4]+RandomMinus($MacroSteps[$Step][7])+RandomMinus($RandomMouse),_Clamp($Delay-9,1,100))
							EndIf
						Case 5	; close window
							WinClose($Window)
							$PerformDelay = True
						Case 6	; if window maximized
							If BitAnd(WinGetState($Window),$WIN_STATE_MAXIMIZED) <> $WIN_STATE_MAXIMIZED Then $Step = JumpToEndBlock($Step)
						Case 7	; if window not maximized
							If BitAnd(WinGetState($Window),$WIN_STATE_MAXIMIZED) = $WIN_STATE_MAXIMIZED Then $Step = JumpToEndBlock($Step)
						Case 8	; if window minimized
							If BitAnd(WinGetState($Window),$WIN_STATE_MINIMIZED) <> $WIN_STATE_MINIMIZED Then $Step = JumpToEndBlock($Step)
						Case 9	; if window not minimized
							If BitAnd(WinGetState($Window),$WIN_STATE_MINIMIZED) = $WIN_STATE_MINIMIZED Then $Step = JumpToEndBlock($Step)
						Case 10	; show window
							WinSetState($Window,"",@SW_SHOW)
							$PerformDelay = True
						Case 11	; hide window
							WinSetState($Window,"",@SW_HIDE)
							$PerformDelay = True
						Case 12	; kill window
							WinKill($Window)
							$PerformDelay = True
						Case 13	; activate window
							SendKeepActive("")
							$ProcessInitalize = False
							$WindowFound = False
							$TimeOutTimer = _Timer_Init()
							While _Timer_Diff($TimeOutTimer) < $ActivateTimeOut
								$FindWindow = _WindowFromProcessId($ProcessId,$MacroSteps[$Step][3])
								If $FindWindow <> Null Then
									$Window = $FindWindow
									While Not $WindowFound And _Timer_Diff($TimeOutTimer) < $ActivateTimeOut
										If WinActivate($Window) > 0 Then $WindowFound = True
									WEnd
									If $WindowFound Then
										Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
										SendKeepActive($Window)
									EndIf
									ExitLoop
								Else
									Sleep(100)
								EndIf
							WEnd
							If Not $WindowFound Then
								_Ok("Macro is stopped." & @CRLF & @CRLF & "A window in the selected process with keyword " & $MacroSteps[$Step][3] & " couldn't be activated in " & $ActivateTimeOut & " ms.")
								$Stopped = 2
							EndIf
							$PerformDelay = True
						Case 14
							If _WindowFromProcessId($ProcessId,$MacroSteps[$Step][3]) = Null Then $Step = JumpToEndBlock($Step)
						Case 15
							If _WindowFromProcessId($ProcessId,$MacroSteps[$Step][3]) <> Null Then $Step = JumpToEndBlock($Step)
						Case 16	; activate first found window
							SendKeepActive("")
							$ProcessInitalize = False
							$WindowFound = False
							AutoItSetOption("WinTitleMatchMode",2)
							$TimeOutTimer = _Timer_Init()
							While _Timer_Diff($TimeOutTimer) < $ActivateTimeOut
								If WinExists($MacroSteps[$Step][3]) Then
									While Not $WindowFound And _Timer_Diff($TimeOutTimer) < $ActivateTimeOut
										$FindWindow = WinActivate($MacroSteps[$Step][3])
										If $FindWindow > 0 Then
											$Window = $FindWindow
											$ProcessId = WinGetProcess($Window)
											$WindowFound = True
											ExitLoop
										EndIf
									WEnd
									If $WindowFound Then
										Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
										SendKeepActive($Window)
										ExitLoop
									EndIf
								EndIf
								Sleep(100)
							WEnd
							AutoItSetOption("WinTitleMatchMode",1)
							If Not $WindowFound Then
								_Ok("Macro is stopped." & @CRLF & @CRLF & "A window with keyword " & $MacroSteps[$Step][3] & " couldn't be activated in " & $ActivateTimeOut & " ms.")
								$Stopped = 2
							EndIf
							$PerformDelay = True
					EndSwitch
				Case $StepTypeRandom
					$RandomMouse = Number($MacroSteps[$Step][3])
					$RandomDelay = Number($MacroSteps[$Step][4])
				Case $StepTypeJump
					Switch $MacroSteps[$Step][2]
						Case 0	; jump
							$JumpStep = MacroStepGetLabelStep($MacroSteps[$Step][3])
							If $JumpStep > -1 Then $Step = $JumpStep-1
						Case 1	; restart
							$Step = -1
							$Initialize = True
						Case 2	; exit
							$Count = Number(GUICtrlRead($LoopCount))	; exit macro by setting count to max
							ExitLoop
						Case 3	; exit block
							For $ExitStep = $Step To UBound($MacroSteps)-1
								If $MacroSteps[$ExitStep][0] = $StepTypeEnd Then
									$Step = $ExitStep
									ExitLoop
								ElseIf UBound($MacroSteps)-1 Then		; no end block found
									$Step = UBound($MacroSteps)-1
								EndIf
							Next
						Case 4	; exit loop
							For $ExitStep = $Step To UBound($MacroSteps)-1
								If $MacroSteps[$ExitStep][0] = $StepTypeEnd Then
									If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) > 1 Then
										For $BlockStep = 0 To UBound($MacroStepsBlock)-1
											If $MacroStepsBlock[$BlockStep][2] = $ExitStep And _InList($MacroSteps[$MacroStepsBlock[$BlockStep][1]][0],$StepTypeRepeat,$StepTypeTime) Then	; repeat or time loop
												$Step = $ExitStep
												$ExitStep = UBound($MacroSteps)-1	; exit higher loop
												ExitLoop
											ElseIf UBound($MacroSteps)-1 Then		; no end loop found
												$Step = UBound($MacroSteps)-1
											EndIf
										Next
									EndIf
								EndIf
							Next
					EndSwitch
				Case $StepTypeEnd
					If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) > 1 Then
						For $BlockStep = 0 To UBound($MacroStepsBlock)-1
							If $MacroStepsBlock[$BlockStep][2] = $Step And $MacroSteps[$MacroStepsBlock[$BlockStep][1]][$MacroStepEnableDisableColumn] = 0 Then	; begin block step belonging to this end block
								Switch $MacroStepsBlock[$BlockStep][0]
									Case $StepTypeRepeat
										; repeat loop and (infinite or smaller than) -> loop
										If $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2] = 0 And ($MacroSteps[$MacroStepsBlock[$BlockStep][1]][3] = 0 Or $MacroStepsBlock[$BlockStep][3] < $MacroSteps[$MacroStepsBlock[$BlockStep][1]][3]) Then
											$MacroStepsBlock[$BlockStep][3] += 1
											$Step = $MacroStepsBlock[$BlockStep][1]		; jump to begin of block
										EndIf
									Case $StepTypeTime
										If TimerDiff($MacroStepsBlock[$BlockStep][3]) < $MacroSteps[$MacroStepsBlock[$BlockStep][1]][3] Then $Step = $MacroStepsBlock[$BlockStep][1]		; jump to begin of block
								EndSwitch
								ExitLoop
							EndIf
						Next
					EndIf
				Case $StepTypeRepeat
					$BlockStep = FindBlockStep($Step)
					If $BlockStep > -1 Then
						If $MacroSteps[$Step][2] = 0 Then ; repeat
							$MacroStepsBlock[$BlockStep][3] = 1
						Else
							$JumpBlock = True
							If $MacroStepsBlock[$BlockStep][3] > -1 Then	; is there a repeat belonging to this end step?
								$RepeatStep = FindBlockStep($MacroStepsBlock[$BlockStep][3])
								Switch $MacroSteps[$Step][2]
									Case 1	; =
										If $MacroStepsBlock[$RepeatStep][3] = $MacroSteps[$Step][3] Then $JumpBlock = False		; equal
									Case 2	; not =
										If $MacroStepsBlock[$RepeatStep][3] <> $MacroSteps[$Step][3] Then $JumpBlock = False	; not equal
									Case 3	; >=
										If $MacroStepsBlock[$RepeatStep][3] >= $MacroSteps[$Step][3] Then $JumpBlock = False	; greater
									Case 4	; <=
										If $MacroStepsBlock[$RepeatStep][3] <= $MacroSteps[$Step][4] Then $JumpBlock = False	; smaller
									Case 5	; between
										If _Between($MacroStepsBlock[$RepeatStep][3],$MacroSteps[$Step][3],$MacroSteps[$Step][4]) Then $JumpBlock = False	; between
									Case 6	; not between
										If Not _Between($MacroStepsBlock[$RepeatStep][3],$MacroSteps[$Step][3],$MacroSteps[$Step][4]) Then $JumpBlock = False	; not between
									Case 7	; odd
										If Mod($MacroStepsBlock[$RepeatStep][3],2) = 1 Then $JumpBlock = False	; odd
									Case 8	; even
										If Mod($MacroStepsBlock[$RepeatStep][3],2) = 0 Then $JumpBlock = False	; even
								EndSwitch
							EndIf
							If $JumpBlock Then $Step = JumpToEndBlock($Step)
						EndIf
					EndIf
				Case $StepTypeTime
					For $BlockStep = 0 To UBound($MacroStepsBlock)-1
						If $MacroStepsBlock[$BlockStep][1] = $Step Then
							$MacroStepsBlock[$BlockStep][3] = TimerInit()
							ExitLoop
						EndIf
					Next
				Case $StepTypePixel
					Switch $MacroSteps[$Step][2]
						Case 0,2		; record pixel area or color
							$PixelX = CalculateX($Window,$Step)
							$PixelY = CalculateY($Window,$Step)
							$PixelAdded = False
							If $PixelCheck > 0 Then
								For $PixelStep = 0 To $PixelCheck-1
									If $Pixels[$PixelStep][1] = $MacroSteps[$Step][2] And $Pixels[$PixelStep][2] = $MacroSteps[$Step][8] Then
										$PixelAdded = True
										$Pixels[$PixelStep][0] = $Step
										If $MacroSteps[$Step][2] = 0 Then
											$Pixels[$PixelStep][3] = PixelChecksum($PixelX,$PixelY,$PixelX+$MacroSteps[$Step][6],$PixelY+$MacroSteps[$Step][7])
										Else
											$Pixels[$PixelStep][3] = PixelGetColor($PixelX,$PixelY)
										EndIf
									EndIf
								Next
							EndIf
							If Not $PixelAdded Then
								$PixelCheck += 1
								ReDim $Pixels[$PixelCheck][4]
								$Pixels[$PixelCheck-1][0] = $Step
								$Pixels[$PixelCheck-1][1] = $MacroSteps[$Step][2]
								$Pixels[$PixelCheck-1][2] = $MacroSteps[$Step][8]	; pixel check id number
								If $MacroSteps[$Step][2] = 0 Then
									$Pixels[$PixelCheck-1][3] = PixelChecksum($PixelX,$PixelY,$PixelX+$MacroSteps[$Step][6],$PixelY+$MacroSteps[$Step][7])
								Else
									$Pixels[$PixelCheck-1][3] = PixelGetColor($PixelX,$PixelY)
								EndIf
							EndIf
						Case 1,3		; pixel area or color changed
							$JumpBlock = True
							If $PixelCheck > 0 Then
								For $PixelStep = 0 To $PixelCheck-1
									If (($Pixels[$PixelStep][1] = 0 And $MacroSteps[$Step][2] = 1) Or ($Pixels[$PixelStep][1] = 2 And $MacroSteps[$Step][2] = 3)) And $Pixels[$PixelStep][2] = $MacroSteps[$Step][8] Then
										$PixelX = CalculateX($Window,$Pixels[$PixelStep][0])
										$PixelY = CalculateY($Window,$Pixels[$PixelStep][0])
										If $MacroSteps[$Step][2] = 1 Then	; if checksum
											If $Pixels[$PixelStep][3] <> PixelChecksum($PixelX,$PixelY,$PixelX+$MacroSteps[$Pixels[$PixelStep][0]][6],$PixelY+$MacroSteps[$Pixels[$PixelStep][0]][7]) Then $JumpBlock = False
										Else
											If $Pixels[$PixelStep][3] <> PixelGetColor($PixelX,$PixelY) Then $JumpBlock = False
										EndIf
										$PerformDelay = True
										ExitLoop
									EndIf
								Next
							EndIf
							If $JumpBlock Then $Step = JumpToEndBlock($Step)
					EndSwitch
				Case $StepTypeInput
					Switch $MacroSteps[$Step][2]
						Case 0
							While Not _IsPressed($MacroContinueKeys[$MacroSteps[$Step][3]][1],$User32DLL)
								If $MacroSteps[$Step][3] <= 4 Then
									ToolTip("Click " & $MacroContinueKeys[$MacroSteps[$Step][3]][0] & " to continue macro")
								Else
									ToolTip("Press " & $MacroContinueKeys[$MacroSteps[$Step][3]][0] & " to continue macro")
								EndIf
								Sleep(10)
								If $Stopped > 0 Then ExitLoop
							WEnd
							While _IsPressed($MacroContinueKeys[$MacroSteps[$Step][3]][1],$User32DLL)
								Sleep(10)
								If $Stopped > 0 Then ExitLoop
							WEnd
						Case 1
							SendKeepActive("")
							_Ok($MacroSteps[$Step][1])
							WinActivate($Window)
							If WinWaitActive($Window,"",$ActivateTimeOut/1000) = 0 Then
								_Ok("Macro can't be initiated." & @CRLF & @CRLF & "Window " & WinGetTitle($Window) & " couldn't be activated in " & $ActivateTimeOut & " ms.")
								$Stopped = 2
								ExitLoop
							Else
								Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
								SendKeepActive($Window)
								$PerformDelay = True
							EndIf
						Case 2
							SendKeepActive("")
							If Not _Sure($MacroSteps[$Step][1]) Then $Step = JumpToEndBlock($Step)
							WinActivate($Window)
							If WinWaitActive($Window,"",$ActivateTimeOut/1000) = 0 Then
								_Ok("Macro can't be initiated." & @CRLF & @CRLF & "Window " & WinGetTitle($Window) & " couldn't be activated in " & $ActivateTimeOut & " ms.")
								$Stopped = 2
								ExitLoop
							Else
								Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
								SendKeepActive($Window)
								$PerformDelay = True
							EndIf
					EndSwitch
				Case $StepTypeAlarm
					Switch $MacroSteps[$Step][2]
						Case 0	; specific date and time
							If (($MacroSteps[$Step][5] = 0 And Not $MacroSteps[$Step][6]) Or $MacroSteps[$Step][5] = 1) And CurrentDate() = $MacroSteps[$Step][4] And CurrentTime() >= $MacroSteps[$Step][3] Then
								$MacroSteps[$Step][6] = True
							ElseIf (($MacroSteps[$Step][5] = 2 And Not $MacroSteps[$Step][6]) Or $MacroSteps[$Step][5] = 3) And CurrentDate() = $MacroSteps[$Step][4] And CurrentTime() < $MacroSteps[$Step][3] Then
								$MacroSteps[$Step][6] = True
							Else
								$Step = JumpToEndBlock($Step)
							EndIf
						Case 1	; on the minute
							If CurrentDate() > $MacroSteps[$Step][4] Then
								$MacroSteps[$Step][3] = Floor(CurrentTime()/100)*100
								$MacroSteps[$Step][4] = CurrentDate()
								$MacroSteps[$Step][6] = False
							ElseIf CurrentTime() >= AddTime($MacroSteps[$Step][3],100) Then
								$MacroSteps[$Step][3] = Floor(CurrentTime()/100)*100
								$MacroSteps[$Step][6] = False
							EndIf
							If Not $MacroSteps[$Step][6] Then
								$MacroSteps[$Step][6] = True
							Else
								$Step = JumpToEndBlock($Step)
							EndIf
						Case 2	; hourly
							If CurrentDate() > $MacroSteps[$Step][4] Then
								$MacroSteps[$Step][5] = AddTime(Floor(CurrentTime()/10000)*10000,$MacroSteps[$Step][3])
								$MacroSteps[$Step][4] = CurrentDate()
								$MacroSteps[$Step][6] = False
							ElseIf CurrentTime() > AddTime($MacroSteps[$Step][5],10000) Then
								$MacroSteps[$Step][5] = AddTime(AddTime(Floor(CurrentTime()/10000)*10000,10000),$MacroSteps[$Step][3])
								$MacroSteps[$Step][6] = False
							EndIf
							If Not $MacroSteps[$Step][6] Then
								$MacroSteps[$Step][6] = True
							Else
								$Step = JumpToEndBlock($Step)
							EndIf
						Case 3	; daily
							If CurrentDate() > $MacroSteps[$Step][4] Then
								$MacroSteps[$Step][4] = CurrentDate()
								$MacroSteps[$Step][6] = False
							EndIf
							If Not $MacroSteps[$Step][6] And CurrentDate() = $MacroSteps[$Step][4] And CurrentTime() >= $MacroSteps[$Step][3] Then
								$MacroSteps[$Step][6] = True
							Else
								$Step = JumpToEndBlock($Step)
							EndIf
					EndSwitch
				Case $StepTypeSound
					Switch $MacroSteps[$Step][2]
						Case 0	; beep
							Beep($MacroSteps[$Step][3],$MacroSteps[$Step][4])
						Case 1	; play sound
							If $MacroSteps[$Step][4] Then
								$Sound[0] = $MacroSteps[$Step][5]
								$Sound[1] = $MacroSteps[$Step][6]
								$Sound[2] = $MacroSteps[$Step][7]
								$Sound[3] = $MacroSteps[$Step][8]
								_SoundPlay($Sound)
							EndIf
						Case 2	; stop sounds
							For $SoundStep = 0 To UBound($MacroSteps)-1
								If $MacroSteps[$SoundStep][4] Then
									$Sound[0] = $MacroSteps[$SoundStep][5]
									$Sound[1] = $MacroSteps[$SoundStep][6]
									$Sound[2] = $MacroSteps[$SoundStep][7]
									$Sound[3] = $MacroSteps[$SoundStep][8]
									_SoundStop($Sound)
								EndIf
							Next
						Case 3	; set volume
							SoundSetWaveVolume($MacroSteps[$SoundStep][3])
						Case 4	; pause sounds
							For $SoundStep = 0 To UBound($MacroSteps)-1
								If $MacroSteps[$SoundStep][4] Then
									$Sound[0] = $MacroSteps[$SoundStep][5]
									$Sound[1] = $MacroSteps[$SoundStep][6]
									$Sound[2] = $MacroSteps[$SoundStep][7]
									$Sound[3] = $MacroSteps[$SoundStep][8]
									_SoundPause($Sound)
								EndIf
							Next
						Case 5	; resumes sounds
							For $SoundStep = 0 To UBound($MacroSteps)-1
								If $MacroSteps[$SoundStep][4] Then
									$Sound[0] = $MacroSteps[$SoundStep][5]
									$Sound[1] = $MacroSteps[$SoundStep][6]
									$Sound[2] = $MacroSteps[$SoundStep][7]
									$Sound[3] = $MacroSteps[$SoundStep][8]
									_SoundResume($Sound)
								EndIf
							Next
					EndSwitch
				Case $StepTypeCounter
					If _InList($MacroSteps[$Step][2],0,1,2) Then	; set, add
						$CounterAdded = False
						If $MacroSteps[$Step][2] = 1 Then
							$MacroSteps[$Step][4] = InputBox($ProgramName,"Enter a value for counter " & $MacroSteps[$Step][3],0)
							If @error <> 0 Then
								_Ok("Macro is stopped due to canceling or an error at the input of a counter value.")
								$Stopped = 2
							EndIf
							$MacroSteps[$Step][4]  = Number($MacroSteps[$Step][4])
							WinActivate($Window)
							If WinWaitActive($Window,"",$ActivateTimeOut/1000) = 0 Then
								_Ok("Macro is stopped." & @CRLF & @CRLF & "Window " & WinGetTitle($Window) & " couldn't be activated in " & $ActivateTimeOut & " ms.")
								$Stopped = 2
							Else
								Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
								SendKeepActive($Window)
								$PerformDelay = True
							EndIf
						EndIf
						If $CounterValue > 0 Then
							For $CounterStep = 0 To $CounterValue-1
								If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] Then	; counter name
									$CounterAdded = True
									$Counters[$CounterStep][0] = $Step
									If $MacroSteps[$Step][2] = 0 Then	; set
										If $MacroSteps[$Step][5] = 0 Then
											$Counters[$CounterStep][2] = $MacroSteps[$Step][4]
										Else	; random
											$Counters[$CounterStep][2] = Random(0,$MacroSteps[$Step][4],1)
										EndIf
									Else	; add
										If $MacroSteps[$Step][5] = 0 Then
											$Counters[$CounterStep][2] += $MacroSteps[$Step][4]
										Else	; random
											$Counters[$CounterStep][2] += Random(0,$MacroSteps[$Step][4],1)
										EndIf
									EndIf
								EndIf
							Next
						EndIf
						If Not $CounterAdded Then
							$CounterValue += 1
							ReDim $Counters[$CounterValue][3]
							$Counters[$CounterValue-1][0] = $Step
							$Counters[$CounterValue-1][1] = $MacroSteps[$Step][3]
							If $MacroSteps[$Step][5] = 0 Then
								$Counters[$CounterValue-1][2] = $MacroSteps[$Step][4]
							Else	; random
								$Counters[$CounterValue-1][2] = Random(0,$MacroSteps[$Step][4],1)
							EndIf
						EndIf
					Else
						$JumpBlock = True
						If $CounterValue > 0 Then
							For $CounterStep = 0 To $CounterValue-1
								Switch $MacroSteps[$Step][2]
									Case 3	; =
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And $Counters[$CounterStep][2] = $MacroSteps[$Step][4] Then $JumpBlock = False	; equal
									Case 4	; not =
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And $Counters[$CounterStep][2] <> $MacroSteps[$Step][4] Then $JumpBlock = False	; not equal
									Case 5	; >=
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And $Counters[$CounterStep][2] >= $MacroSteps[$Step][4] Then $JumpBlock = False	; greater and equal
									Case 6	; <=
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And $Counters[$CounterStep][2] <= $MacroSteps[$Step][5] Then $JumpBlock = False	; smaller and equal
									Case 7	; between
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And _Between($Counters[$CounterStep][2],$MacroSteps[$Step][4],$MacroSteps[$Step][5]) Then $JumpBlock = False	; between
									Case 8	; not between
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And Not _Between($Counters[$CounterStep][2],$MacroSteps[$Step][4],$MacroSteps[$Step][5]) Then $JumpBlock = False	; between
									Case 9	; odd
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And Mod($Counters[$CounterStep][2],2) = 1 Then $JumpBlock = False	; odd
									Case 10	; even
										If $Counters[$CounterStep][1] = $MacroSteps[$Step][3] And Mod($Counters[$CounterStep][2],2) = 0 Then $JumpBlock = False	; even
								EndSwitch
							Next
						EndIf
						If $JumpBlock Then $Step = JumpToEndBlock($Step)
					EndIf
				Case $StepTypeFile
					Switch $MacroSteps[$Step][2]
						Case 0,1,2,3
							$JumpBlock = True
							Switch $MacroSteps[$Step][2]
								Case 0
									If FileExists(FilePathTranslateVariables($MacroSteps[$Step][3])) Then $JumpBlock = False
								Case 1
									If Not FileExists(FilePathTranslateVariables($MacroSteps[$Step][3])) Then $JumpBlock = False
								Case 2
									If FileExists(FilePathTranslateVariables($MacroSteps[$Step][3])) Then $JumpBlock = False
								Case 3
									If Not FileExists(FilePathTranslateVariables($MacroSteps[$Step][3])) Then $JumpBlock = False
							EndSwitch
							If $JumpBlock Then $Step = JumpToEndBlock($Step)
						Case 4
							$StampAdded = False
							If $FileStamp > 0 Then
								For $StampStep = 0 To $FileStamp-1
									If $FileStamps[$StampStep][1] = $MacroSteps[$Step][3] Then	; file
										$StampAdded = True
										$FileStamps[$StampStep][0] = $Step
										$FileStamps[$StampStep][2] = FileGetSize(FilePathTranslateVariables($MacroSteps[$Step][3]))
										$FileStamps[$StampStep][3] = FileGetAttrib(FilePathTranslateVariables($MacroSteps[$Step][3]))
										$FileStamps[$StampStep][4] = FileGetTime(FilePathTranslateVariables($MacroSteps[$Step][3]),$FT_MODIFIED,$FT_STRING)
									EndIf
								Next
							EndIf
							If Not $StampAdded Then
								$FileStamp += 1
								ReDim $FileStamps[$FileStamp][5]
								$FileStamps[$FileStamp-1][0] = $Step
								$FileStamps[$FileStamp-1][1] = $MacroSteps[$Step][3]
								$FileStamps[$FileStamp-1][2] = FileGetSize(FilePathTranslateVariables($MacroSteps[$Step][3]))
								$FileStamps[$FileStamp-1][3] = FileGetAttrib(FilePathTranslateVariables($MacroSteps[$Step][3]))
								$FileStamps[$FileStamp-1][4] = FileGetTime(FilePathTranslateVariables($MacroSteps[$Step][3]),$FT_MODIFIED,$FT_STRING)
							EndIf
						Case 5
							$JumpBlock = True
							If $FileStamp > 0 Then
								For $StampStep = 0 To $FileStamp-1
									If $FileStamps[$FileStamp-1][1] = $MacroSteps[$Step][3] And _
									  ($FileStamps[$FileStamp-1][2] <> FileGetSize(FilePathTranslateVariables($MacroSteps[$FileStamps[$FileStamp-1][0]][3])) Or $FileStamps[$FileStamp-1][3] <> FileGetAttrib(FilePathTranslateVariables($MacroSteps[$FileStamps[$FileStamp-1][0]][3])) Or _
									   $FileStamps[$FileStamp-1][4] <> FileGetTime(FilePathTranslateVariables($MacroSteps[$FileStamps[$FileStamp-1][0]][3]),$FT_MODIFIED,$FT_STRING)) Then
										$JumpBlock = False
										ExitLoop
									EndIf
								Next
							EndIf
							If $JumpBlock Then $Step = JumpToEndBlock($Step)
						Case 6
							$JumpBlock = True
							If $FileStamp > 0 Then
								For $StampStep = 0 To $FileStamp-1
									If $FileStamps[$FileStamp-1][1] = $MacroSteps[$Step][3] And _
									  ($FileStamps[$FileStamp-1][2] = FileGetSize(FilePathTranslateVariables($MacroSteps[$FileStamps[$FileStamp-1][0]][3])) And $FileStamps[$FileStamp-1][3] = FileGetAttrib(FilePathTranslateVariables($MacroSteps[$FileStamps[$FileStamp-1][0]][3])) And _
									   $FileStamps[$FileStamp-1][4] = FileGetTime(FilePathTranslateVariables($MacroSteps[$FileStamps[$FileStamp-1][0]][3]),$FT_MODIFIED,$FT_STRING)) Then
										$JumpBlock = False
										ExitLoop
									EndIf
								Next
							EndIf
							If $JumpBlock Then $Step = JumpToEndBlock($Step)
					EndSwitch
				Case $StepTypeControl
					Switch $MacroSteps[$Step][2]
						Case 0
							$Paused = True
						Case 1
							$Stopped = 1
						Case 2
							$CloseByMacro = True
							$Stopped = 1
					EndSwitch
				Case $StepTypeProcess
					Switch $MacroSteps[$Step][2]
						Case 0
							SendKeepActive("")
							$Window = $MacroSteps[$Step][$MacroStepWindowColumn]
							$ProcessId = $MacroSteps[$Step][$MacroStepWindowColumn+1]
							$ProcessInitalize = False
							WinActivate($Window)
							If WinWaitActive($Window,"",$ActivateTimeOut/1000) = 0 Then
								_Ok("Macro is stopped." & @CRLF & @CRLF & "Window " & WinGetTitle($Window) & " couldn't be activated in " & $ActivateTimeOut & " ms.")
								$Stopped = 2
							Else
								Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
								SendKeepActive($Window)
								$PerformDelay = True
							EndIf
						Case 1
							If $MacroSteps[$Step][7] = 0 Then
								Switch $MacroSteps[$Step][6]
									Case 0
										$MacroSteps[$Step][9] = Run('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',FilePathTranslateVariables($MacroSteps[$Step][5]),@SW_SHOWNORMAL)
									Case 1
										$MacroSteps[$Step][9] = Run('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',FilePathTranslateVariables($MacroSteps[$Step][5]),@SW_SHOWMAXIMIZED)
									Case 2
										$MacroSteps[$Step][9] = Run('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',FilePathTranslateVariables($MacroSteps[$Step][5]),@SW_SHOWMINIMIZED)
									Case 3
										$MacroSteps[$Step][9] = Run('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',FilePathTranslateVariables($MacroSteps[$Step][5]),@SW_HIDE)
								EndSwitch
								If $MacroSteps[$Step][9] > 0 Then
									$ProcessId = $MacroSteps[$Step][9]
									$ProcessInitalize = False
									SendKeepActive("")
									$WindowFound = False
									$TimeOutTimer = _Timer_Init()
									While _Timer_Diff($TimeOutTimer) < $MacroSteps[$Step][8]
										$FindWindow = _WindowFromProcessId($MacroSteps[$Step][9],$MacroSteps[$Step][4])
										If $FindWindow <> Null Then	; process id is the same as run -> try to find window
											$Window = $FindWindow
											While Not $WindowFound And _Timer_Diff($TimeOutTimer) < $MacroSteps[$Step][8]
												If WinActivate($Window) > 0 Then $WindowFound = True
											WEnd
											If $WindowFound Then
												Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
												SendKeepActive($Window)
											EndIf
											ExitLoop
										Else
											Sleep(100)
										EndIf
									WEnd
									If Not $WindowFound Then
										If StringLen($MacroSteps[$Step][4]) = 0 Then
											_Ok("Macro is stopped." & @CRLF & @CRLF & "A window belonging to program " & $MacroSteps[$Step][3] & " couldn't be activated in " & $MacroSteps[$Step][8] & " ms.")
										Else
											_Ok("Macro is stopped." & @CRLF & @CRLF & "A window with keyword " & $MacroSteps[$Step][4] & " belonging to program " & $MacroSteps[$Step][3] & " couldn't be activated in " & $MacroSteps[$Step][8] & " ms.")
										EndIf
										$Stopped = 2
									Else
										$PerformDelay = True
									EndIf
								Else
									_Ok("Macro is stopped." & @CRLF & @CRLF & "Program " & $MacroSteps[$Step][3] & " couldn't be started. Does the program file exist?")
									$Stopped = 2
								EndIf
							Else
								Switch $MacroSteps[$Step][6]
									Case 0
										$MacroSteps[$Step][9] = RunWait('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',$MacroSteps[$Step][5],@SW_SHOWNORMAL)
									Case 1
										$MacroSteps[$Step][9] = RunWait('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',$MacroSteps[$Step][5],@SW_SHOWMAXIMIZED)
									Case 2
										$MacroSteps[$Step][9] = RunWait('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',$MacroSteps[$Step][5],@SW_SHOWMINIMIZED)
									Case 3
										$MacroSteps[$Step][9] = RunWait('"' & FilePathTranslateVariables($MacroSteps[$Step][3]) & '"',$MacroSteps[$Step][5],@SW_HIDE)
								EndSwitch
								$ProcessId = $MacroSteps[$Step][9]
								SendKeepActive("")
								WinActivate($Window)
								If WinWaitActive($Window,"",$ActivateTimeOut/1000) = 0 Then
									_Ok("Macro is stopped." & @CRLF & @CRLF & "Window " & WinGetTitle($Window) & " couldn't be activated in " & $ActivateTimeOut & " ms.")
									$Stopped = 2
								Else
									Sleep($ActivateDelay)	; small delay for preventing wrong process id by WinGetProcess("[ACTIVE]")
									SendKeepActive($Window)
									$PerformDelay = True
								EndIf
							EndIf
						Case 2
							If _ProcessGetProcessId(_FileName(FilePathTranslateVariables($MacroSteps[$Step][3]))) = 0 Then $Step = JumpToEndBlock($Step)
						Case 3
							If _ProcessGetProcessId(_FileName(FilePathTranslateVariables($MacroSteps[$Step][3]))) > 0 Then $Step = JumpToEndBlock($Step)
						Case 4
							$ProcessCheck = False
						Case 5
							$ProcessCheck = True
						Case 6
							$Window = WinGetHandle("[ACTIVE]")
							$ProcessId = WinGetProcess("[ACTIVE]")
							$ProcessInitalize = False
							$PerformDelay = True
					EndSwitch
				Case $StepTypeCapture
					If Stringlen($CapturePath) = 0 Then $CapturePath = @MyDocumentsDir & "\" & $ProgramName & "\Capture"
					If _Between($MacroSteps[$Step][2],0,3) And Not FileExists(FilePathTranslateVariables($CapturePath)) Then DirCreate(FilePathTranslateVariables($CapturePath))
					Switch $MacroSteps[$Step][2]
						Case 0
							$Capture = 1
						Case 1
							$Capture = 1
							$CaptureLeft = $MacroSteps[$Step][3]
							$CaptureTop = $MacroSteps[$Step][4]
							$CaptureRight = $CaptureLeft+$MacroSteps[$Step][5]
							$CaptureBottom = $CaptureTop+$MacroSteps[$Step][6]
						Case 2
							$Capture = 2
						Case 3
							$Capture = 2
							$CaptureLeft = $MacroSteps[$Step][3]
							$CaptureTop = $MacroSteps[$Step][4]
							$CaptureRight = $CaptureLeft+$MacroSteps[$Step][5]
							$CaptureBottom = $CaptureTop+$MacroSteps[$Step][6]
						Case 4
							$Capture = 0
						Case 5
							$CapturePath = FilePathTranslateVariables($MacroSteps[$Step][3])
							If Stringlen($CapturePath) = 0 Then $CapturePath = @MyDocumentsDir & "\" & $ProgramName & "\Capture"
							$CapturePoint = $MacroSteps[$Step][4] = 1
					EndSwitch
				Case $StepTypeShutdown
					Switch $MacroSteps[$Step][2]
						Case 0
							If $MacroSteps[$Step][3] = 0 Or _Sure("Log current user off?") Then
								Shutdown($SD_LOGOFF)
								$CloseByMacro = True
								$Stopped = 1
							EndIf
						Case 1
							If $MacroSteps[$Step][3] = 0 Or _Sure("Go into standby mode?") Then
								Shutdown($SD_STANDBY)
								$CloseByMacro = True
								$Stopped = 1
							EndIf
						Case 1
							If $MacroSteps[$Step][3] = 0 Or _Sure("Sleep by writing memory to disk?") Then
								Shutdown($SD_HIBERNATE)
								$CloseByMacro = True
								$Stopped = 1
							EndIf
						Case 3
							If $MacroSteps[$Step][3] = 0 Or _Sure("Restart the computer, all programs will be closed?") Then
								Shutdown($SD_REBOOT)
								$CloseByMacro = True
								$Stopped = 1
							EndIf
						Case 4
							If $MacroSteps[$Step][3] = 0 Or _Sure("Shut down computer by closing programs?") Then
								Shutdown($SD_SHUTDOWN)
								$CloseByMacro = True
								$Stopped = 1
							EndIf
						Case 5
							If $MacroSteps[$Step][3] = 0 Or _Sure("Shut down computer by forcing the closing of programs?") Then
								Shutdown(BitOr($SD_SHUTDOWN,$SD_FORCE))
								$CloseByMacro = True
								$Stopped = 1
							EndIf
						Case 6
							If $MacroSteps[$Step][3] = 0 Or _Sure("Power down computer, killing all processe immediately?") Then
								Shutdown($SD_POWERDOWN)
								$CloseByMacro = True
								$Stopped = 1
							EndIf
					EndSwitch
			EndSwitch
			If GUICtrlRead($ShowBar) = $GUI_CHECKED Then
				GUICtrlSetPos($PlayBar,0,0,$Step/(UBound($MacroSteps)-1)*$BarWidth,$BarHeight)
				GUICtrlSetPos($PlayBarPercentage,$BarWidth*$Step/(UBound($MacroSteps)-1)*0.5,1)
				GUICtrlSetData($PlayBarPercentage,String(Round($Step/(UBound($MacroSteps)-1)*100)) & "%")
			EndIf
			;$ActiveProcess = WinGetProcess("[ACTIVE]") -> doesn't work as it's suppose to
			$ActiveProcess = WinGetProcess(WinGetHandle(""))
			If Not $ProcessInitalize And $Stopped = 0 And $ProcessCheck And $ActiveProcess <> $ProcessId Then	; check if still in selected/started process unless user canceled checking
				$Stopped = 2
				_Ok("To prevent problems the macro has been stopped." & @CRLF & @CRLF & "A different process than the macro process " & _ProcessGetName($ProcessId) & " (" & $ProcessId & ") was being activated: " & _ProcessGetName($ActiveProcess) & " (" & $ActiveProcess & ")" & @CRLF & @CRLF & "You can disable this checking by setting a Process check off step however that's dangerous.")
			EndIf
			Switch $Capture
				Case 1
					_ScreenCapture_Capture($CapturePath & "\" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & "-" & @MSEC & ".jpg",$CaptureLeft,$CaptureTop,$CaptureRight,$CaptureBottom,$CapturePoint)
				Case 2
					_ScreenCapture_CaptureWnd($CapturePath & "\" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & "-" & @MSEC & ".jpg",$Window,$CaptureLeft,$CaptureTop,$CaptureRight,$CaptureBottom,$CapturePoint)
			EndSwitch
			If $Stopped > 0 Then ExitLoop
			If Not $Paused Then
				If Not $DelayInStep And ($PerformDelay Or $AlwaysDelay) Then
					$DelayTime = $Delay
					$DelayTime -= _Timer_Diff($DelayTimer)
					If $RandomDelay > 0 Then $DelayTime += Random(0,$RandomDelay,1)	; add random delay
					Sleep(_Maximum(1,$DelayTime))
				EndIf
			Else
				While $Paused And $Stopped = 0
					ToolTip("Macro is paused by " & $PauseKeys[$PauseKey][0] & ", continue it with same key. " & @CRLF & $StartStopKeys[$StartStopKey][0] & " to stop macro")
					Sleep(100)
				WEnd
			EndIf
		Next
		If $Stopped > 0 Or $Count = Number(GUICtrlRead($LoopCount)) Then ExitLoop	; not >=, loop count can be 0
		$Count += 1
	WEnd
	; reset/normalize things after playing
	SetStopKey(False)
	SetPauseKey(False)
	SendKeepActive("")
	GUISetState(@SW_RESTORE,$GUI)
	GUISetState(@SW_SHOWNORMAL,$GUI)
	ToolTip("")
	For $Step = 0 To UBound($MacroSteps)-1
		If $MacroSteps[$Step][0] = $StepTypeSound And $MacroSteps[$Step][2] = 1 And $MacroSteps[$Step][4] Then
			$Sound[0] = $MacroSteps[$Step][5]
			$Sound[1] = $MacroSteps[$Step][6]
			$Sound[2] = $MacroSteps[$Step][7]
			$Sound[3] = $MacroSteps[$Step][8]
			_SoundClose($Sound)
			$MacroSteps[$Step][4] = False
			$MacroSteps[$Step][5] = 0
			$MacroSteps[$Step][6] = 0
			$MacroSteps[$Step][7] = 0
			$MacroSteps[$Step][8] = 0
		EndIf
	Next
	If GUICtrlRead($ShowBar) = $GUI_CHECKED Then GUIDelete($GUIBar)
	FillProcesses()
	SetControlStates()
	HighlightStep($MacroStep,$PreviousStep)
	_GUICtrlListView_SetItemSelected($Macros,$MacroStep,True,True)
	_GUICtrlListView_EnsureVisible($Macros,$MacroStep)
	If Not $CloseByMacro Then
		If $Stopped = 1 Then
			_Ok("Macro is stopped by pressing " & $StartStopKeys[$StartStopKey][0] & " or by a corresponding macro step.")
		ElseIf $Stopped = 2 Then
			_Ok("Macro is stopped by " & $ProgramName & " due to a timeout, a fail-safe reason or canceling by you.")
		EndIf
	EndIf
EndFunc

Func FindBlockStep($Step)
	For $BlockStep = 0 To UBound($MacroStepsBlock)-1
		If $MacroStepsBlock[$BlockStep][1] = $Step Then Return $BlockStep
	Next
	Return -1
EndFunc

Func JumpToEndBlock($Step)
	Local $EndStep

	$EndStep = UBound($MacroSteps)-1	; if no end block exit this macro loop
	For $BlockStep = 0 To UBound($MacroStepsBlock)-1
		If $MacroStepsBlock[$BlockStep][1] = $Step Then
			$EndStep = $MacroStepsBlock[$BlockStep][2]
			ExitLoop
		EndIf
	Next
	Return $EndStep
EndFunc

Func MacroStepGetLabelStep($Label)
	If $MacroStepCount = 0 Then Return -1
	For $Step = 0 to UBound($MacroSteps)-1
		If $MacroSteps[$Step][0] = 5 And $MacroSteps[$Step][1] = $Label Then Return $Step
	Next
	Return -1
EndFunc

Func RandomCharacter($Type = 1)
	Local $Number

	Switch $Type
		Case 1	; a-z, A-Z, 0-9
			Do
				$Number = Random(48,122,1)
			Until $Number <= 57 Or ($Number >= 65 And $Number <= 90) Or ($Number >= 97 And $Number <= 122)
		Case 2	; a-z, A-Z
			Do
				$Number = Random(65,122,1)
			Until ($Number >= 65 And $Number <= 90) Or ($Number >= 97 And $Number <= 122)
		Case 3	; a-z
			$Number = Random(97,122,1)
		Case 4	; a-z, 0-9
			Do
				$Number = Random(48,122,1)
			Until $Number <= 57 Or ($Number >= 97 And $Number <= 122)
		Case 5	; 0-9
			$Number = Random(48,57,1)
	EndSwitch
	Return Chr($Number)
EndFunc

Func DetectCalculateX($Window,$Type)
	Switch $Type
		Case 0		; relative to window
			Return MouseGetPos(0)-_WindowGetX($Window,True)
		Case 1,2	; absolute screen coordinates or relative to mouse position
			Return MouseGetPos(0)
		Case 3		; relative to screen center
			Return Round(@DesktopWidth/2)-MouseGetPos(0)
		Case 4		; relative to window center
			Return _WindowGetX($Window,True)+Round(WindowWidth($Window)/2)-MouseGetPos(0)
		Case Else	; relative to program window
			Return MouseGetPos(0)-_WindowGetX($Window,True)
	EndSwitch
EndFunc

Func DetectCalculateY($Window,$Type)
	Switch $Type
		Case 0		; relative to window
			Return MouseGetPos(1)-_WindowGetY($Window,True)
		Case 1,2	; absolute screen coordinates or relative to mouse position
			Return MouseGetPos(1)
		Case 3		; relative to screen center
			Return Round(@DesktopHeight/2)-MouseGetPos(1)
		Case 4		; relative to window center
			Return _WindowGetY($Window,True)+Round(WindowHeight($Window)/2)-MouseGetPos(1)
		Case Else	; relative to program window
			Return MouseGetPos(1)-_WindowGetY($Window,True)
	EndSwitch
EndFunc

Func CalculateX($Window,$Step,$X = 0,$Type = 0)		; calculate x coordinate according to set position type
	Local $MouseX
	If $Step >= 0 Then
		$X = Number($MacroSteps[$Step][3])
		$Type = $MacroSteps[$Step][5]
	EndIf
	Switch $Type
		Case 1		; absolute screen coordinates
			$MouseX = $X
		Case 2		; relative to mouse position
			$MouseX = MouseGetPos(0)+$X
		Case 3		; relative to screen center
			$MouseX = Round(@DesktopWidth/2)+$X
		Case 4		; relative to window center
			$MouseX = _WindowGetX($Window,True)+Round(WindowWidth($Window)/2)+$X
		Case Else	; relative to program window
			$MouseX = _WindowGetX($Window,True)+$X
	EndSwitch
	Return $MouseX
EndFunc

Func CalculateY($Window,$Step,$Y = 0,$Type = 0)		; calculate y coordinate according to set position type
	Local $MouseY
	If $Step >= 0 Then
		$Y = Number($MacroSteps[$Step][4])
		$Type = $MacroSteps[$Step][5]
	EndIf
	Switch $Type
		Case 1		; absolute screen coordinates
			$MouseY = $Y
		Case 2		; relative to mouse position
			$MouseY = MouseGetPos(1)+$Y
		Case 3		; relative to screen center
			$MouseY = Round(@DesktopHeight/2)+$Y
		Case 4		; relative to window center
			$MouseY = _WindowGetY($Window,True)+Round(WindowHeight($Window)/2)+$Y
		Case Else	; relative to program window
			$MouseY = _WindowGetY($Window,True)+$Y
	EndSwitch
	Return $MouseY
EndFunc

Func InitializeMacro()
	$MacroStep = -1
	ReDim $MacroSteps[1][$MacroStepColumns]
	For $Column = 0 To $MacroStepColumns-1
		$MacroSteps[0][$Column] = $Column = 1 ? "" : 0
	Next
EndFunc

Func NewMacro()
	InitializeMacro()
	FillMacroSteps(-1)
	FileDelete(FileFailSave())
	$MacroFile = ""
	WriteFileName()
	Changed(False)
	FillProcesses()
	GUICtrlSetData($LoopCount,1)
	GUICtrlSetData($StepDelay,$DefaultStepDelay)
	GUICtrlSetState($StepDelayAlways,$GUI_UNCHECKED)
	GUICtrlSetState($TestMacro,$GUI_UNCHECKED)
	GUICtrlSetState($ShowBar,$GUI_UNCHECKED)
	GUICtrlSetState($ShowRecording,$GUI_UNCHECKED)
	GUICtrlSetState($MoveTime,0)
	_GUICtrlComboBox_SetCurSel($ShowMain,0)
	SetPlayTip()
	ResetUndos()
	ShowButtonsToolTip()
EndFunc

Func SaveMacro($Dialog = False,$UndoFile = "")	; save macro to file chosen by user
	Local $FileHandle,$FileName,$StepString

	If StringLen($UndoFile) > 0 Then
		$FileName = $UndoFile
	ElseIf Not $Dialog And StringLen($MacroFile) > 0 Then
		$FileName = $MacroFile
	Else
		$FileName = FileSaveDialog("Save macro","","Macro's (*" & $FileExtension & ")",BitOR($FD_PATHMUSTEXIST,$FD_PROMPTOVERWRITE),$MacroFile,$GUI)
		If @error Or StringLen(StringStripWS($FileName,$STR_STRIPLEADING+$STR_STRIPTRAILING)) = 0 Then Return False
	EndIf
	$FileHandle = FileOpen($FileName,$FO_OVERWRITE)
	If @error Then
		If StringLen($UndoFile) = 0 Then _Ok($MacroFile & " couldn't be saved.")
		Return False
	Else
		If StringLen($UndoFile) = 0 Then AddRecentFile($FileName)
		FileWriteLine($FileHandle,$Version)
		FileWriteLine($FileHandle,$ProcessList[SelectedProcessIndex()][0])
		FileWriteLine($FileHandle,$ProcessList[SelectedProcessIndex()][1])
		FileWriteLine($FileHandle,GUICtrlRead($StepDelayAlways) = $GUI_CHECKED ? 1 : 0)
		FileWriteLine($FileHandle,GUICtrlRead($LoopCount))
		FileWriteLine($FileHandle,GUICtrlRead($ShowRecording) = $GUI_CHECKED ? 1 : 0)
		FileWriteLine($FileHandle,GUICtrlRead($MoveTime))
		FileWriteLine($FileHandle,GUICtrlRead($StepDelay))
		FileWriteLine($FileHandle,_GUICtrlComboBox_GetCurSel($ShowMain))
		FileWriteLine($FileHandle,GUICtrlRead($ShowBar) = $GUI_CHECKED ? 1 : 0)
		FileWriteLine($FileHandle,GUICtrlRead($TestMacro) = $GUI_CHECKED ? 1 : 0)
		; write macro steps
		FileWriteLine($FileHandle,$MacroStep)	; current select step
		If $MacroStep > -1 Then
			For $Step = 0 To UBound($MacroSteps)-1
				$StepString = ""
				For $Column = 0 To $MacroStepEnableDisableColumn+1
					If ($Column = 7 Or $Column = 8) And $MacroSteps[$Step][0] = $StepTypeProcess And $MacroSteps[$Step][2] = 0 Then	; process -> don't save process id and window handle
						$StepString &= ""
					Else
						$StepString &= StringReplace(String($MacroSteps[$Step][$Column]),$FileAttributeDelimiter,"")
					EndIf
					$StepString &= $FileAttributeDelimiter
				Next
				FileWriteLine($FileHandle,StringTrimRight($StepString,StringLen($FileAttributeDelimiter)))
			Next
		EndIf
		FileClose($FileHandle)
		If StringLen($UndoFile) = 0 Then
			FileDelete(FileFailSave())
			$MacroFile = $FileName
			WriteFileName()
			Changed(False)
		EndIf
		Return True
	EndIf
EndFunc

Func OpenMacro($Dialog = True,$UndoFile = "")	; read macro from file selected by user
	Const $GUIReadWidth = 200,$GUIReadHeight = 40,$GUIReadMargin = 20
	Local $FileHandle,$FileName,$Process,$Window,$Step,$StepString,$GUIRead

	If StringLen($UndoFile) > 0 Then
		$FileName = $UndoFile
	ElseIf Not $Dialog Then
		If FileExists($MacroFile) Then $FileName = $MacroFile
	Else
		$FileName = FileOpenDialog("Open macro","","Macro's (*" & $FileExtension & ")",BitOR($FD_FILEMUSTEXIST,$FD_PATHMUSTEXIST),$MacroFile,$GUI)
		If @error Or StringLen(StringStripWS($FileName,$STR_STRIPLEADING+$STR_STRIPTRAILING)) = 0 Then Return False
	EndIf
	$FileHandle = FileOpen($FileName)
	If @error Then
		If StringLen($UndoFile) = 0 Then _Ok($MacroFile & " couldn't be opened.")
		Return False
	Else
		$Step = -1
		If StringLen($UndoFile) = 0 Then AddRecentFile($FileName)
		$MacroFileVersion = FileReadLine($FileHandle)
		InitializeMacro()
		$Process = FileReadLine($FileHandle)
		$Window = FileReadLine($FileHandle)
		GUICtrlSetState($StepDelayAlways,Number(FileReadLine($FileHandle)) = 1 ? $GUI_CHECKED : $GUI_UNCHECKED)
		GUICtrlSetData($LoopCount,FileReadLine($FileHandle))
		GUICtrlSetState($ShowRecording,Number(FileReadLine($FileHandle)) = 1 ? $GUI_CHECKED : $GUI_UNCHECKED)
		GUICtrlSetData($MoveTime,FileReadLine($FileHandle))
		GUICtrlSetData($StepDelay,FileReadLine($FileHandle))
		_GUICtrlComboBox_SetCurSel($ShowMain,Number(FileReadLine($FileHandle)))
		GUICtrlSetState($ShowBar,Number(FileReadLine($FileHandle)) = 1 ? $GUI_CHECKED : $GUI_UNCHECKED)
		GUICtrlSetState($TestMacro,Number(FileReadLine($FileHandle)) = 1 ? $GUI_CHECKED : $GUI_UNCHECKED)
		; read macro steps
		$MacroStep = Number(FileReadLine($FileHandle))
		If StringLen($UndoFile) = 0 Then
			$GUIRead = GUICreate("",$GUIReadWidth,$GUIReadHeight,$GUIPosition[0]+$GUIPosition[2]/2-$GUIReadWidth/2,$GUIPosition[1]+$GUIPosition[3]/2-$GUIReadHeight/2,$WS_POPUP+$WS_DLGFRAME,$WS_EX_TOPMOST)
			$ReadLabel =  GUICtrlCreateLabel("Reading steps ...",$GUIReadMargin,11,$GUIReadWidth-2*$GUIReadMargin,$GUIReadHeight)
			GUICtrlSetFont(-1,10)
		EndIf
		If $MacroStep > -1 Then
			GUISetState(@SW_SHOWNOACTIVATE)
			While 1
				$StepString = FileReadLine($FileHandle) & $FileAttributeDelimiter
				If @error <> 0 Then ExitLoop
				If $Step = -1 Then
					$Step = 0
				Else
					$Step += 1
					ReDim $MacroSteps[$Step+1][$MacroStepColumns]
				EndIf
				If StringLen($UndoFile) = 0 Then GUICtrlSetData($ReadLabel,"Reading step " & $Step)
				For $Column = 0 To $MacroStepEnableDisableColumn+1
					If StringLen($StepString) = 0 Then
						$MacroSteps[$Step][$Column] = 0
					Else
						$MacroSteps[$Step][$Column] = _StringToTag($StepString,$FileAttributeDelimiter)
						$StepString = _StringFromTag($StepString,$FileAttributeDelimiter)
					EndIf
				Next
			WEnd
		EndIf
		FileClose($FileHandle)
		If $MacroStep = -1 And $Step > -1 Then $MacroStep = 0
		$MacroPreviousStep = -1
		If StringLen($UndoFile) = 0 Then GUICtrlSetData($ReadLabel,"Filling processes list ...")
		FillProcesses()
		SelectProcess($Process,$Window,$Step = -1 Or ($MacroSteps[0][0] = $StepTypeProcess And _InList($MacroSteps[0][2],0,1)))
		If StringLen($UndoFile) = 0 Then GUICtrlSetData($ReadLabel,"Filling macro steps list ...")
		FillMacroSteps($MacroStep,True)
		If StringLen($UndoFile) = 0 Then GUIDelete($GUIRead)
		SetPlayTip()
		If StringLen($UndoFile) = 0 Then
			If $Step = -1 Then _Ok($FileName & " has no macro steps.")
			FileDelete(FileFailSave())
			$MacroFile = $FileName
			WriteFileName()
			Changed(False)
			ResetUndos()
		EndIf
		ShowButtonsToolTip()
		Return True
	EndIf
EndFunc

Func SelectProcess($ProcessName,$Window,$NoMessage)
	Local $ProcessFound = -1,$ProcessIndex = 0
	For $Process = 0 To UBound($ProcessList)-1
		If $ProcessList[$Process][0] = $ProcessName Then
			$CurrentProcess = $ProcessList[$Process][3]
			If $ProcessFound = -1 Then
				$ProcessFound = $Process
				_GUICtrlComboBox_SetCurSel($Processes,$ProcessIndex)
			EndIf
			If $ProcessList[$Process][1] = $Window Then
				_GUICtrlComboBox_SetCurSel($Processes,$ProcessIndex)
				$CurrentWindow = $ProcessList[$Process][2]
				Return
			EndIf
		EndIf
		If $ShowInvisible Or $ProcessList[$Process][4] Then $ProcessIndex += 1
	Next
	If $ProcessFound = -1 Then
		$CurrentProcess = -1
		$CurrentWindow = -1
		If Not $NoMessage Then _Ok("The main selected process " & $ProcessName & " of the macro wasn't found.")
		Changed()
	Else
		$CurrentProcess = $ProcessList[$ProcessFound][3]
		$CurrentWindow = -1
		If Not $NoMessage Then _Ok("The main selected process of the macro was found but not the selected window " & $Window & ".")
		Changed()
	EndIf
EndFunc

Func EnableDisableMacroSteps($FromStep,$ToStep,$EnableDisable)
	If $FromStep = -1 Then Return
	If $ToStep < $FromStep Then
		Local $Step = $FromStep
		$FromStep = $ToStep
		$ToStep =  $Step
	EndIf
	$ToStep = _Minimum($ToStep,UBound($MacroSteps)-1)
	For $Step = $FromStep To $ToStep
		$MacroSteps[$Step][$MacroStepEnableDisableColumn] = $EnableDisable
	Next
	FillMacroSteps($MacroStep)
	Changed()
EndFunc

Func CutMacroSteps($FromStep,$ToStep)
	If $FromStep = -1 Then Return
	If $ToStep < $FromStep Then
		Local $Step = $FromStep
		$FromStep = $ToStep
		$ToStep =  $Step
	EndIf
	$ToStep = _Minimum($ToStep,UBound($MacroSteps)-1)
	If $ToStep-$FromStep+1 = UBound($MacroSteps) Then
		$MacroStep = -1
		$MacroPreviousStep = -1
		ReDim $MacroSteps[1][$MacroStepColumns]
	Else
		For $Step = $ToStep+1 To UBound($MacroSteps)-1
			For $Column = 0 To $MacroStepColumns-1
				$MacroSteps[$Step+$FromStep-$ToStep-1][$Column] = $MacroSteps[$Step][$Column]
			Next
		Next
		ReDim $MacroSteps[UBound($MacroSteps)-$ToStep+$FromStep-1][$MacroStepColumns]
		If $MacroStep >= UBound($MacroSteps) Then
			$MacroStep = UBound($MacroSteps)-1
			$MacroPreviousStep = -1
		EndIf
		If $MacroStep <= 0 Then $MacroPreviousStep = -1
	EndIf
	FillMacroSteps($MacroStep)
	Changed()
EndFunc

Func CopyMacroSteps($FromStep,$ToStep)
	If $FromStep = -1 Then Return
	If $ToStep < $FromStep Then
		Local $Step = $FromStep
		$FromStep = $ToStep
		$ToStep =  $Step
	EndIf
	$ToStep = _Minimum($ToStep,UBound($MacroSteps)-1)
	ReDim $MacroStepsCopy[$ToStep-$FromStep+1][$MacroStepColumns]
	For $Step = $FromStep To $ToStep
		For $Column = 0 To $MacroStepColumns-1
			$MacroStepsCopy[$Step-$FromStep][$Column] = $MacroSteps[$Step][$Column]
		Next
	Next
EndFunc

Func PasteMacroSteps($InsertStep)
	If $InsertStep = -1 Then
		ReDim $MacroSteps[UBound($MacroStepsCopy)][$MacroStepColumns]
		For $Step = 0 To UBound($MacroStepsCopy)-1
			For $Column = 0 To $MacroStepColumns-1
				$MacroSteps[$Step][$Column] = $MacroStepsCopy[$Step][$Column]
			Next
		Next
		$MacroStep = 0
	Else
		; create gap for copied steps
		ReDim $MacroSteps[UBound($MacroSteps)+UBound($MacroStepsCopy)][$MacroStepColumns]
		For $Step = UBound($MacroSteps)-1-UBound($MacroStepsCopy) To $InsertStep Step -1
			For $Column = 0 To $MacroStepColumns-1
				$MacroSteps[$Step+UBound($MacroStepsCopy)][$Column] = $MacroSteps[$Step][$Column]
			Next
		Next
		; pasted copied steps
		For $Step = 0 To UBound($MacroStepsCopy)-1
			For $Column = 0 To $MacroStepColumns-1
				$MacroSteps[$InsertStep+$Step][$Column] = $MacroStepsCopy[$Step][$Column]
			Next
		Next
	EndIf
	FillMacroSteps($MacroStep)
	Changed()
EndFunc

Func UpMacroStep()
	Local $Value
	If UBound($MacroSteps) = 1 Or $MacroStep < 1 Then Return
	$MacroStep -= 1
	For $Column = 0 To $MacroStepColumns-1
		$Value = $MacroSteps[$MacroStep][$Column]
		$MacroSteps[$MacroStep][$Column] = $MacroSteps[$MacroStep+1][$Column]
		$MacroSteps[$MacroStep+1][$Column] = $Value
	Next
	FillMacroSteps($MacroStep)
	Changed()
EndFunc

Func DownMacroStep()
	Local $Value
	If UBound($MacroSteps) = 1 Or $MacroStep < 0 Or $MacroStep >= UBound($MacroSteps)-1 Then Return
	$MacroStep += 1
	For $Column = 0 To $MacroStepColumns-1
		$Value = $MacroSteps[$MacroStep][$Column]
		$MacroSteps[$MacroStep][$Column] = $MacroSteps[$MacroStep-1][$Column]
		$MacroSteps[$MacroStep-1][$Column] = $Value
	Next
	FillMacroSteps($MacroStep)
	Changed()
EndFunc

Func DeleteMacroStep()
	If $MacroStep > -1 Then
		If UBound($MacroSteps) = 1 Then
			$MacroStep = -1
			$MacroPreviousStep = -1
			ReDim $MacroSteps[1][$MacroStepColumns]
		Else
			If $MacroStep >= UBound($MacroSteps)-1 Then
				ReDim $MacroSteps[UBound($MacroSteps)-1][$MacroStepColumns]
				$MacroStep = UBound($MacroSteps)-1
			Else
				For $Step = $MacroStep To UBound($MacroSteps)-2
					For $Column = 0 To $MacroStepColumns-1
						$MacroSteps[$Step][$Column] = $MacroSteps[$Step+1][$Column]
					Next
				Next
				ReDim $MacroSteps[UBound($MacroSteps)-1][$MacroStepColumns]
			EndIf
			$MacroPreviousStep = -1
		EndIf
		FillMacroSteps($MacroStep)
		Changed()
	EndIf
EndFunc

Func DuplicateMacroStep()
	If $MacroStep = UBound($MacroSteps)-1 Then
		ReDim $MacroSteps[UBound($MacroSteps)+1][$MacroStepColumns]
		For $Column = 0 To $MacroStepColumns-1
			$MacroSteps[$MacroStep+1][$Column] = $MacroSteps[$MacroStep][$Column]
		Next
	Else
		ReDim $MacroSteps[UBound($MacroSteps)+1][$MacroStepColumns]
		For $Step = UBound($MacroSteps)-1 To $MacroStep+1 Step -1
			For $Column = 0 To $MacroStepColumns-1
				$MacroSteps[$Step][$Column] = $MacroSteps[$Step-1][$Column]
			Next
		Next
	EndIf
	$MacroStep += 1
	FillMacroSteps($MacroStep)
	Changed()
EndFunc

Func EditMacroStep($Add,$Type,$Value1 = "",$Value2 = 0,$Value3 = 0,$Value4 = 0,$Value5 = 0,$Value6 = 0,$Value7 = 0,$Value8 = 0,$Value9 = 0)
	AddChangeMacroStep($Add,$Type,$Value1,$Value2,$Value3,$Value4,$Value5,$Value6,$Value7,$Value8,$Value9)
	FillMacroSteps($MacroStep)
	Changed()
EndFunc

Func AddChangeMacroStep($Add,$Type,$Value1 = "",$Value2 = 0,$Value3 = 0,$Value4 = 0,$Value5 = 0,$Value6 = 0,$Value7 = 0,$Value8 = 0,$Value9 = 0)
	If $Add Then
		If $MacroStep = -1 Then
			$MacroPreviousStep = -1
			$MacroStep = 0
		Else
			If $MacroStep >= UBound($MacroSteps)-1 Then
				ReDim $MacroSteps[UBound($MacroSteps)+1][$MacroStepColumns]
				$MacroStep = UBound($MacroSteps)-1
			Else
				ReDim $MacroSteps[UBound($MacroSteps)+1][$MacroStepColumns]
				For $Step = UBound($MacroSteps)-1 To $MacroStep+2 Step -1
					For $Column = 0 To $MacroStepColumns-1
						$MacroSteps[$Step][$Column] = $MacroSteps[$Step-1][$Column]
					Next
				Next
				$MacroStep += 1
			EndIf
		EndIf
	Else
		$MacroPreviousStep = -1
	EndIf
	$MacroSteps[$MacroStep][0] = $Type
	$MacroSteps[$MacroStep][1] = $Value1
	$MacroSteps[$MacroStep][2] = $Value2
	$MacroSteps[$MacroStep][3] = $Value3
	$MacroSteps[$MacroStep][4] = $Value4
	$MacroSteps[$MacroStep][5] = $Value5
	$MacroSteps[$MacroStep][6] = $Value6
	$MacroSteps[$MacroStep][7] = $Value7
	$MacroSteps[$MacroStep][8] = $Value8
	$MacroSteps[$MacroStep][9] = $Value9
	For $Column = 10 To $MacroStepColumns-1
		$MacroSteps[$MacroStep][$Column] = 0
	Next
EndFunc

Func FillMacroSteps($SelectStep,$Report = False)
	Local $StepString,$Tab = "",$BlockStep,$Window,$ProcessId,$ProcessFound,$IssueFree = True

	$Window = $ProcessList[SelectedProcessIndex()][2]
	$ProcessId = $ProcessList[SelectedProcessIndex()][3]
	GUICtrlSetData($MacroStepsLabel,"Macro steps")
	_GUICtrlListView_BeginUpdate($Macros)
	_GUICtrlListView_DeleteAllItems($Macros)
	If $SelectStep > -1 Then
		ReDim $MacroStepsBlock[1]
		For $Step = 0 To UBound($MacroSteps)-1
			If $ViewInvertColors Then
				$MacroSteps[$Step][$MacroStepColorColumn] = $MacroStepTypes[$MacroSteps[$Step][0]][3]
			Else
				$MacroSteps[$Step][$MacroStepColorColumn] = $MacroStepTypes[$MacroSteps[$Step][0]][2]
			EndIf
			$BlockStep = -1
			If $MacroSteps[$Step][0] = $StepTypeEnd Then
				$Tab = StringTrimRight($Tab,$LeftIndent)
				If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) > 1 Then
					For $BlockStep = UBound($MacroStepsBlock)-1 To 0 Step -1
						If $MacroStepsBlock[$BlockStep][2] = -1 Then
							$MacroStepsBlock[$BlockStep][2] = $Step		; set corresponding end block step
							If $ViewInvertColors Then
								$MacroSteps[$Step][$MacroStepColorColumn] = $MacroStepTypes[$MacroSteps[$MacroStepsBlock[$BlockStep][1]][0]][3]
							Else
								$MacroSteps[$Step][$MacroStepColorColumn] = $MacroStepTypes[$MacroSteps[$MacroStepsBlock[$BlockStep][1]][0]][2]
							EndIf
							ExitLoop
						EndIf
					Next
				EndIf
			EndIf
			If $ViewNumbers Then
				GUICtrlCreateListViewItem(($step+1) & "  " & $Tab & MacroStepDescription($Step,$BlockStep),$Macros)
			Else
				GUICtrlCreateListViewItem($Tab & MacroStepDescription($Step,$BlockStep),$Macros)
			EndIf
			HighlightStep(-1,$Step)
			If ($MacroSteps[$Step][0] = $StepTypeWindow And _InList($MacroSteps[$Step][2],6,7,8,9,14,15)) Or $MacroSteps[$Step][0] = $StepTypeRepeat Or $MacroSteps[$Step][0] = $StepTypeTime Or _
			   ($MacroSteps[$Step][0] = $StepTypeInput And $MacroSteps[$Step][2] = 2) Or _
			   ($MacroSteps[$Step][0] = $StepTypePixel And _InList($MacroSteps[$Step][2],1,3)) Or $MacroSteps[$Step][0] = $StepTypeAlarm Or _
			   ($MacroSteps[$Step][0] = $StepTypeCounter And $MacroSteps[$Step][2] > 2) Or ($MacroSteps[$Step][0] = $StepTypeFile And _Between($MacroSteps[$Step][2],0,3)) Or _
			   ($MacroSteps[$Step][0] = $StepTypeFile And _InList($MacroSteps[$Step][2],5,6)) Or _
			   ($MacroSteps[$Step][0] = $StepTypeProcess And _InList($MacroSteps[$Step][2],2,3)) Then	; begin block step
				$Tab &= _StringRepeat(" ",$LeftIndent)
				If UBound($MacroStepsBlock,$UBOUND_DIMENSIONS) = 1 Then
					ReDim $MacroStepsBlock[1][5]
				Else
					ReDim $MacroStepsBlock[UBound($MacroStepsBlock)+1][5]
				EndIf
				$MacroStepsBlock[UBound($MacroStepsBlock)-1][0] = $MacroSteps[$Step][0]
				$MacroStepsBlock[UBound($MacroStepsBlock)-1][1] = $Step					; begin step of block
				$MacroStepsBlock[UBound($MacroStepsBlock)-1][2] = -1					; end block step position
				If $MacroSteps[$Step][0] = $StepTypeRepeat Then
					If $MacroSteps[$Step][2] = 0 Then
						$MacroStepsBlock[UBound($MacroStepsBlock)-1][3] = $MacroSteps[$Step][3]	; repeat counter
					Else
						$MacroStepsBlock[UBound($MacroStepsBlock)-1][3] = FindLoopStep($Step,$StepTypeRepeat,2)	; register repeat belonging to end step
					EndIf
				Else
					$MacroStepsBlock[UBound($MacroStepsBlock)-1][3] = -1	; timer handle
				EndIf
				$MacroStepsBlock[UBound($MacroStepsBlock)-1][4] = 0
			EndIf
			If $MacroSteps[$Step][0] = $StepTypeProcess And $MacroSteps[$Step][2] = 0 Then	; process the to selected process step
				If $MacroSteps[$Step][8] = 0 Or Not ProcessExists($MacroSteps[$Step][8]) Then
					$ProcessFound = False
					For $Process = 0 To UBound($ProcessList)-1
						If $ProcessList[$Process][0] = $MacroSteps[$Step][5] Then
							$MacroSteps[$Step][8] = $ProcessList[$Process][3]	; process id
							Switch $MacroSteps[$Step][3]
								Case 0
									If $ProcessList[$Process][1] = $MacroSteps[$Step][6] Then
										$MacroSteps[$Step][7] = $ProcessList[$Process][2]	; window handle
										$ProcessFound = True
										ExitLoop
									EndIf
								Case 1
									$MacroSteps[$Step][7] = _WindowFromProcessId($MacroSteps[$Step][8],$MacroSteps[$Step][4])
									If $MacroSteps[$Step][7] <> Null Then $ProcessFound = True
									ExitLoop
								Case 2
									$MacroSteps[$Step][7] = $ProcessList[$Process][2]
									$ProcessFound = True
									ExitLoop
							EndSwitch
						EndIf
					Next
					If Not $ProcessFound And $Report Then
						Switch $MacroSteps[$Step][3]
							Case 0
								_Ok("In step " & ($Step+1) & " the selected process " & $MacroSteps[$Step][5] & " and/or the selected window " & $MacroSteps[$Step][6] & " wasn't found." & @CRLF & @CRLF & "You have to start it now before it can be selected in the macro.")
								$IssueFree = False
							Case 1
								_Ok("In step " & ($Step+1) & " the selected process " & $MacroSteps[$Step][5] & " and/or the selected window with keyword " & $MacroSteps[$Step][4] & " wasn't found." & @CRLF & @CRLF & "You have to start it now before it can be selected in the macro.")
								$IssueFree = False
							Case 2
								_Ok("In step " & ($Step+1) & " the selected process " & $MacroSteps[$Step][5] & " wasn't found." & @CRLF & @CRLF & "You have to start it now before it can be selected in the macro.")
								$IssueFree = False
						EndSwitch
					EndIf
				EndIf
				If Number($MacroSteps[$Step][7]) > 0 Then
					$Window = $MacroSteps[$Step][7]
					$ProcessId = $MacroSteps[$Step][8]
				EndIf
			EndIf
			$MacroSteps[$Step][$MacroStepWindowColumn] = $Window	; set macro to window
			$MacroSteps[$Step][$MacroStepWindowColumn+1] = $ProcessId	; set macro to process id
		Next
		_GUICtrlListView_EndUpdate($Macros)
		$MacroStep = $SelectStep
		_GUICtrlListView_SetItemSelected($Macros,$MacroStep,True,True)
		HighlightStep($MacroStep,-1)
		_GUICtrlListView_EnsureVisible($Macros,$MacroStep)
	Else
		_GUICtrlListView_EndUpdate($Macros)
		$MacroStep = -1
		$MacroPreviousStep = -1
	EndIf
	Return $IssueFree
EndFunc

Func FindLoopStep($Step,$StepType,$StepSubType = -1)
	Local $Level = 0

	For $LoopStep = $Step To 0 Step -1	; search backwards to corresponding loop step
		If $MacroSteps[$LoopStep][0] = $StepType And ($StepSubType = -1 Or $MacroSteps[$LoopStep][$StepSubType] = 0) Then	; element 0 is loop command
			If $Level > 0 Then
				$Level -= 1
			Else
				For $BlockStep = 0 To UBound($MacroStepsBlock)-1
					If $MacroStepsBlock[$BlockStep][1] = $LoopStep Then Return $LoopStep
				Next
			EndIf
		ElseIf $MacroSteps[$LoopStep][0] = $StepTypeEnd	Then ; is end block belonging to a loop command
			For $BlockStep = 0 To UBound($MacroStepsBlock)-1
				If $MacroStepsBlock[$BlockStep][2] = $LoopStep And $MacroStepsBlock[$BlockStep][0] = $StepType And ($StepSubType = -1 Or $MacroSteps[$MacroStepsBlock[$BlockStep][1]][$StepSubType] = 0) Then ; begin block step belonging to this end block
					$Level += 1
					ExitLoop
				EndIf
			Next
		EndIf
	Next
	Return -1
EndFunc

Func MacroStepDescription($Step,$BlockStep = -1)	; create description for the macro step
	Local $StepString

	If _StringEmpty($MacroSteps[$Step][1]) Then
		$StepString = $MacroStepTypes[$MacroSteps[$Step][0]][0]
		Switch $MacroSteps[$Step][0]
			Case $StepTypeMouse
				If $MacroSteps[$Step][2] = 10 Then
					$StepString = "Push mouse position onto stack " & StringLower($MacroMousePositions[$MacroSteps[$Step][5]])
				ElseIf $MacroSteps[$Step][2] = 11 Then
					$StepString = "Pop mouse position from stack"
				Else
					$StepString = $MacroMouseButtons[$MacroSteps[$Step][2]][0] & " on position " & $MacroSteps[$Step][3] & "," & $MacroSteps[$Step][4] & " " & StringLower($MacroMousePositions[$MacroSteps[$Step][5]])
					If Number($MacroSteps[$Step][6]) > 0 Or Number($MacroSteps[$Step][7]) > 0 Then $StepString &= ", random " & $MacroSteps[$Step][6] & "," & $MacroSteps[$Step][7]
					If Number($MacroSteps[$Step][8]) = 1 Then $StepString &= ", instantaneously"
					If Number($MacroSteps[$Step][9]) > 1 Then $StepString &= ", " & Number($MacroSteps[$Step][9]) & " clicks"
				EndIf
			Case $StepTypeKey
				If $MacroSteps[$Step][2] > 0 Then
					$StepString &= " " & $MacroKeyModes[$MacroSteps[$Step][2]][0] & "+"
				Else
					$StepString &= " "
				EndIf
				If $MacroSteps[$Step][3] = 0 Then
					If Number($MacroSteps[$Step][5]) = 0 Then
						$StepString &= StringReplace($MacroSteps[$Step][4],"|","<pipeline>")
					Else
						$StepString &= $MacroTextMode[$MacroSteps[$Step][5]]
					EndIf
				Else
					$StepString &= $MacroKeys[$MacroSteps[$Step][3]][0]
				EndIf
			Case $StepTypeDelay
				If $MacroSteps[$Step][2] = 0 Then
					If Number($MacroSteps[$Step][3]) > 0 Then $StepString &= " " & $MacroSteps[$Step][3] & " ms"
					If Number($MacroSteps[$Step][4]) > 0 Then $StepString &= ", random " & $MacroSteps[$Step][4] & " ms"
				Else
					$StepString = "Set delay of each step to " & $MacroSteps[$Step][3] & " ms"
				EndIf
			Case $StepTypeWindow
				$StepString = $MacroWindowActions[$MacroSteps[$Step][2]]
				Switch $MacroSteps[$Step][2]
					Case 3
						$StepString &= " " & $MacroSteps[$Step][3] & "," & $MacroSteps[$Step][4] & " " & StringLower($MacroWindowPositions[$MacroSteps[$Step][5]])
						If Number($MacroSteps[$Step][6]) > 0 Or Number($MacroSteps[$Step][7]) > 0 Then $StepString &= ", random " & $MacroSteps[$Step][6] & "," & $MacroSteps[$Step][7]
						If Number($MacroSteps[$Step][8]) = 1 Then $StepString &= ", instantaneously"
					Case 4
						$StepString &= " to width " & $MacroSteps[$Step][3] & " height " & $MacroSteps[$Step][4] & " " & StringLower($MacroWindowSizes[$MacroSteps[$Step][5]])
						If Number($MacroSteps[$Step][6]) > 0 Or Number($MacroSteps[$Step][7]) > 0 Then $StepString &= ", random " & $MacroSteps[$Step][6] & "," & $MacroSteps[$Step][7]
						If Number($MacroSteps[$Step][8]) = 1 Then $StepString &= ", instantaneously"
					Case 13
						$StepString = "Switch to a window of selected process with keyword " & $MacroSteps[$Step][3]
					Case 14
						$StepString = "If a window with keyword " & $MacroSteps[$Step][3] & " exists"
					Case 15
						$StepString = "If a window with keyword " & $MacroSteps[$Step][3] & " not exists"
					Case 16
						$StepString = "Switch to first found window with keyword " & $MacroSteps[$Step][3]
				EndSwitch
			Case $StepTypeRandom
				If $MacroSteps[$Step][3] > 0 Then $StepString &= " mouse movement " & $MacroSteps[$Step][3]
				If $MacroSteps[$Step][4] > 0 Then $StepString &= " delay " & $MacroSteps[$Step][4] & " ms"
				If $MacroSteps[$Step][3] = 0 And $MacroSteps[$Step][4] = 0 Then $StepString &= " off"
			Case $StepTypeJump
				$StepString = $MacroJumpTypes[$MacroSteps[$Step][2]]
				If $MacroSteps[$Step][2] = 0 Then $StepString &= " " & $MacroSteps[$Step][3]
			Case $StepTypeEnd
				If $BlockStep >= 0 Then
					Switch $MacroStepsBlock[$BlockStep][0]
						Case $StepTypeWindow
							Switch $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2]
								Case 6
									$StepString = "End if window maximized"
								Case 7
									$StepString = "End if window not maximized"
								Case 8
									$StepString = "End if window minimized"
								Case 9
									$StepString = "End if window not minimized"
								Case 14
									$StepString = "End if a window exists"
								Case 15
									$StepString = "End if a window not exists"
							EndSwitch
						Case $StepTypeRepeat
							Switch $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2]
								Case 0
									$StepString = "End repeat"
								Case 1	; =
									$StepString = "End if repeat counter equal"
								Case 2	; not =
									$StepString = "End if repeat counter not equal"
								Case 3	; >
									$StepString = "End if repeat counter greater or equal"
								Case 4	; <
									$StepString = "End if repeat counter smaller or equal"
								Case 5	; between
									$StepString = "End if repeat counter between"
								Case 6	; not between
									$StepString = "End if repeat counter not beween"
								Case 7	; odd
									$StepString = "End if repeat counter odd"
								Case 8	; even
									$StepString = "End if repeat counter even"
							EndSwitch
						Case $StepTypeTime
							$StepString = "End time loop"
						Case $StepTypeAlarm
							$StepString = "End time alarm"
						Case $StepTypeInput
							$StepString = "End if confirmed"
						Case $StepTypePixel
							If $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2] = 1 Then
								$StepString = "End if pixel area changed"
							ElseIf $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2] = 3 Then
								$StepString = "End if pixel color changed"
							EndIf
						Case $StepTypeCounter
							Switch $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2]
								Case 2	; =
									$StepString = "End if counter equal"
								Case 3	; not =
									$StepString = "End if counter not equal"
								Case 4	; >=
									$StepString = "End if counter greater or equal"
								Case 5	; <=
									$StepString = "End if counter smaller or equal"
								Case 6	; between
									$StepString = "End if counter between"
								Case 7	; not between
									$StepString = "End if counter not beween"
								Case 8	; odd
									$StepString = "End if counter odd"
								Case 9	; even
									$StepString = "End if counter even"
							EndSwitch
						Case $StepTypeFile
							Switch $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2]
								Case 0
									$StepString = "End if file exists"
								Case 1
									$StepString = "End if file not exists"
								Case 2
									$StepString = "End if folder exists"
								Case 3
									$StepString = "End if folder not exist"
								Case 5
									$StepString = "End if file changed"
								Case 6
									$StepString = "End if file not changed"
							EndSwitch
						Case $StepTypeProcess
							Switch $MacroSteps[$MacroStepsBlock[$BlockStep][1]][2]
								Case 2
									$StepString = "End if process exists"
								Case 3
									$StepString = "End if process not exists"
							EndSwitch
					EndSwitch
				EndIf
			Case $StepTypeRepeat
				Switch $MacroSteps[$Step][2]
					Case 0
						If $MacroSteps[$Step][3] = 0 Then
							$StepString &= " infinitely"
						Else
							$StepString &= " " & $MacroSteps[$Step][3] & " times"
						EndIf
					Case 1
						$StepString = "If repeat counter is " & $MacroSteps[$Step][3]
					Case 2
						$StepString = "If repeat counter is not " & $MacroSteps[$Step][3]
					Case 3
						$StepString = "If repeat counter is greater or equal than " & $MacroSteps[$Step][3]
					Case 4
						$StepString = "If repeat counter is smaller or equal than " & $MacroSteps[$Step][4]
					Case 5
						$StepString = "If repeat counter is between " & $MacroSteps[$Step][3] & " and " & $MacroSteps[$Step][4]
					Case 6
						$StepString = "If repeat counter is not between " & $MacroSteps[$Step][3] & " and " & $MacroSteps[$Step][4]
					Case 7
						$StepString = "If repeat counter is odd"
					Case 8
						$StepString = "If repeat counter is even"
				EndSwitch
			Case $StepTypeTime
				$StepString &= " " & $MacroSteps[$Step][3] & " ms"
			Case $StepTypeInput
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString &= " " & $MacroContinueKeys[$MacroSteps[$Step][3]][0]
					Case 1
						$StepString = "Popup message"
					Case 2
						$StepString = "If confirm"
				EndSwitch
			Case $StepTypePixel
				If $MacroSteps[$Step][2] = 0 Then
					$StepString = $MacroPixelTypes[$MacroSteps[$Step][2]] & " on position " & $MacroSteps[$Step][3] & "," & $MacroSteps[$Step][4] & " (" & $MacroMousePositions[$MacroSteps[$Step][5]] & ") size " & $MacroSteps[$Step][6] & "," & $MacroSteps[$Step][7] & " id no " & $MacroSteps[$Step][8]
				ElseIf $MacroSteps[$Step][2] = 2 Then
					$StepString = $MacroPixelTypes[$MacroSteps[$Step][2]] & " on position " & $MacroSteps[$Step][3] & "," & $MacroSteps[$Step][4] & " (" & $MacroMousePositions[$MacroSteps[$Step][5]] & ") id no " & $MacroSteps[$Step][8]
				Else
					$StepString = $MacroPixelTypes[$MacroSteps[$Step][2]] & " id no " & $MacroSteps[$Step][8]
				EndIf
			Case $StepTypeAlarm
				Switch $MacroSteps[$Step][2]
					Case 0	; specific date and time
						Switch $MacroSteps[$Step][5]
							Case 0
								$StepString = "After time is " & DateString($MacroSteps[$Step][4]) & " " & TimeString($MacroSteps[$Step][3]) & " run block once"
							Case 1
								$StepString = "After time is " & DateString($MacroSteps[$Step][4]) & " " & TimeString($MacroSteps[$Step][3]) & " run block always"
							Case 2
								$StepString = "Before time is " & DateString($MacroSteps[$Step][4]) & " " & TimeString($MacroSteps[$Step][3]) & " run block once"
							Case 3
								$StepString = "Before time is " & DateString($MacroSteps[$Step][4]) & " " & TimeString($MacroSteps[$Step][3]) & " run block always"
						EndSwitch
					Case 1	; minute
						$StepString = "Run block on the minute"
					Case 2	; hour
						$StepString = "Run block hourly on " & TimeString($MacroSteps[$Step][3])
					Case 3	; dialy
						$StepString = "Run block every day on " & TimeString($MacroSteps[$Step][3])
				EndSwitch
			Case $StepTypeSound
				$StepString = $MacroSoundTypes[$MacroSteps[$Step][2]]
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString &= " at " & $MacroSteps[$Step][3] & " Hz for " & $MacroSteps[$Step][4] & " ms"
					Case 1
						$StepString &= " " & $MacroSteps[$Step][3]
					Case 3
						$StepString &= " to " & $MacroSteps[$Step][3] & "%"
				EndSwitch
			Case $StepTypeFile
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString = "If file " & $MacroSteps[$Step][3] & " exists"
					Case 1
						$StepString = "If file " & $MacroSteps[$Step][3] & " doesn't exist"
					Case 2
						$StepString = "If folder " & $MacroSteps[$Step][3] & " exists"
					Case 3
						$StepString = "If folder " & $MacroSteps[$Step][3] & " doesn't exist"
					Case 4
						$StepString = "Record stamp of file " & $MacroSteps[$Step][3]
					Case 5
						$StepString = "If file " & $MacroSteps[$Step][3] & " changed"
					Case 6
						$StepString = "If file " & $MacroSteps[$Step][3] & " isn't changed"
				EndSwitch
			Case $StepTypeCounter
				Switch $MacroSteps[$Step][2]
					Case 0
						If $MacroSteps[$Step][5] = 0 Then
							$StepString = "Set counter " & $MacroSteps[$Step][3] & " to " & $MacroSteps[$Step][4]
						Else
							$StepString = "Set counter " & $MacroSteps[$Step][3] & " to a random value between 0 and " & $MacroSteps[$Step][4]
						EndIf
					Case 1
						$StepString = "Set counter " & $MacroSteps[$Step][3] & " during macro playing"
					Case 2
						If $MacroSteps[$Step][5] = 0 Then
							$StepString = "Add " & $MacroSteps[$Step][4] & " to counter " & $MacroSteps[$Step][3]
						Else
							$StepString = "Add a random value between 0 and " & $MacroSteps[$Step][4] & " to counter " & $MacroSteps[$Step][3]
						EndIf
					Case 3
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is " & $MacroSteps[$Step][4]
					Case 4
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is not " & $MacroSteps[$Step][4]
					Case 5
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is greater or equal than " & $MacroSteps[$Step][4]
					Case 6
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is smaller or equal than " & $MacroSteps[$Step][5]
					Case 7
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is between " & $MacroSteps[$Step][4] & " and " & $MacroSteps[$Step][5]
					Case 8
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is not between " & $MacroSteps[$Step][4] & " and " & $MacroSteps[$Step][5]
					Case 9
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is odd"
					Case 10
						$StepString = "If counter " & $MacroSteps[$Step][3] & " is even"
				EndSwitch
			Case $StepTypeControl
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString = "Pause macro, continue it yourself"
					Case 1
						$StepString = "Stop macro"
					Case 2
						$StepString = "Stop macro and close " & $ProgramName
				EndSwitch
			Case $StepTypeProcess
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString = "Set to process " & $MacroSteps[$Step][5]
						Switch $MacroSteps[$Step][3]
							Case 0
								$StepString &= " - " & $MacroSteps[$Step][6]
							Case 1
								$StepString &= " - window by keyword " & $MacroSteps[$Step][4]
						EndSwitch
					Case 1
						$StepString = "Start program " & _FileName($MacroSteps[$Step][3])
						If $MacroSteps[$Step][7] = 1 Then $StepString &= " and pause macro"
						If StringLen($MacroSteps[$Step][4]) > 0 Then $StepString &= ", window keyword " & $MacroSteps[$Step][4]
						$StepString &= ", " & StringLower($MacroProcessWindowStates[$MacroSteps[$Step][6]])
						If $MacroSteps[$Step][8] > 0 Then $StepString &= ", timeout " & $MacroSteps[$Step][8] & " ms"
					Case 2
						$StepString = "If process " & $MacroSteps[$Step][3] & " exists"
					Case 3
						$StepString = "If process " & $MacroSteps[$Step][3] & " not exists"
					Case 4,5,6
						$StepString = $MacroProcessTypes[$MacroSteps[$Step][2]]
				EndSwitch
			Case $StepTypeCapture
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString = "Capture screen on"
					Case 1
						$StepString = "Capture screen on, area at " & $MacroSteps[$Step][3] & "," & $MacroSteps[$Step][4] & " size " & $MacroSteps[$Step][5] & "," & $MacroSteps[$Step][6]
					Case 2
						$StepString = "Capture window on"
					Case 3
						$StepString = "Capture window on, area at " & $MacroSteps[$Step][3] & "," & $MacroSteps[$Step][4] & " size " & $MacroSteps[$Step][5] & "," & $MacroSteps[$Step][6]
					Case 4
						$StepString = "Capture off"
					Case 5
						$StepString = "Capture to folder "
						If StringLen($MacroSteps[$Step][3]) = 0 Then
							$StepString &= @MyDocumentsDir & "\" & $ProgramName
						Else
							$StepString &= $MacroSteps[$Step][3]
						EndIf
						If $MacroSteps[$Step][4] = 1 Then $StepString &= ", capture pointer"
				EndSwitch
			Case $StepTypeShutdown
				$StepString = $MacroShutDownActions[$MacroSteps[$Step][2]]
				If $MacroSteps[$Step][3] = 1 Then $StepString &= ", confirm by asking"
			Case $StepTypeComment
				If $MacroSteps[$Step][2] = 1 Then
					$StepString = _StringRepeat("-",400)
				Else
					$StepString = ""
				EndIf
		EndSwitch
	Else
		Switch $MacroSteps[$Step][0]
			Case $StepTypeMouse
				$StepString = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroMouseButtons[$MacroSteps[$Step][2]][0]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroMousePositions[$MacroSteps[$Step][5]]),"$5",$MacroSteps[$Step][6]),"$6",$MacroSteps[$Step][7]),"$7",$MacroSteps[$Step][8]),"$8",$MacroSteps[$Step][9])
			Case $StepTypeKey
				$StepString = StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroKeyModes[$MacroSteps[$Step][2]][0] & "+"),"$2",$MacroKeys[$MacroSteps[$Step][3]][0]),"$3",$MacroSteps[$Step][4])
			Case $StepTypeDelay
				$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$2",$MacroSteps[$Step][3] & " ms"),"$3",$MacroSteps[$Step][4] & " ms")
			Case $StepTypeWindow
				Switch $MacroSteps[$Step][2]
					Case 3
						$StepString = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroWindowActions[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroWindowPositions[$MacroSteps[$Step][5]]),"$5",$MacroSteps[$Step][6]),"$6",$MacroSteps[$Step][7]),"$7",$MacroSteps[$Step][8])
					Case 4
						$StepString = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroWindowActions[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroWindowSizes[$MacroSteps[$Step][5]]),"$5",$MacroSteps[$Step][6]),"$6",$MacroSteps[$Step][7]),"$7",$MacroSteps[$Step][8])
					Case 13,14,15,16
						$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroWindowActions[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3])
				EndSwitch
			Case $StepTypeRandom
				$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroSteps[$Step][3]),"$2",$MacroSteps[$Step][4])
			Case $StepTypeLabel
				$StepString = $MacroStepTypes[$MacroSteps[$Step][0]][0] & " " & $MacroSteps[$Step][1]
			Case $StepTypeJump
				$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroJumpTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3])
			Case $StepTypeEnd
				$StepString = $MacroSteps[$Step][1]
			Case $StepTypeRepeat
				$StepString = StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroRepeatTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4])
			Case $StepTypeTime
				$StepString = StringReplace($MacroSteps[$Step][1],"$1",$MacroSteps[$Step][3] & " ms")
			Case $StepTypeInput
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString = StringReplace($MacroSteps[$Step][1],"$1",$MacroContinueKeys[$MacroSteps[$Step][2]][0])
					Case 1
						$StepString = "Popup message " & $MacroSteps[$Step][1]
					Case 2
						$StepString = "If confirmed " & $MacroSteps[$Step][1]
				EndSwitch
			Case $StepTypePixel
				If $MacroSteps[$Step][2] = 0 Or $MacroSteps[$Step][2] = 2 Then
					$StepString = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroPixelTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroMousePositions[$MacroSteps[$Step][5]]),"$5",$MacroSteps[$Step][6]),"$6",$MacroSteps[$Step][7]),"$7",$MacroSteps[$Step][8])
				Else
					$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroPixelTypes[$MacroSteps[$Step][2]]),"$5",$MacroSteps[$Step][8])
				EndIf
			Case $StepTypeAlarm
				$StepString = StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroAlarmTypes[$MacroSteps[$Step][2]]),"$2",TimeString($MacroSteps[$Step][3])),"$3",DateString($MacroSteps[$Step][4])),"$4",$MacroAlarmSpecifics[$MacroSteps[$Step][5]])
			Case $StepTypeSound
				$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroSoundTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3])
			Case $StepTypeCounter
				$StepString = StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroRepeatTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroSteps[$Step][5])
			Case $StepTypeFile
				$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroControlTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3])
			Case $StepTypeControl
				$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroControlTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3])
			Case $StepTypeProcess
				Switch $MacroSteps[$Step][2]
					Case 0
						$StepString = StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroProcessTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroProcessWindowStates[$MacroSteps[$Step][5]])
					Case 1
						$StepString = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroProcessTypes[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroSteps[$Step][5]),"$5",$MacroProcessWindowStates[$MacroSteps[$Step][6]]),"$6",$MacroSteps[$Step][7]),"$7",$MacroSteps[$Step][8])
				EndSwitch
			Case $StepTypeCapture
				StringReplace(StringReplace(StringReplace(StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroCaptureActions[$MacroSteps[$Step][2]]),"$2",$MacroSteps[$Step][3]),"$3",$MacroSteps[$Step][4]),"$4",$MacroSteps[$Step][5]),"$5",$MacroSteps[$Step][6])
			Case $StepTypeShutdown
				$StepString = StringReplace(StringReplace($MacroSteps[$Step][1],"$1",$MacroShutDownActions[$MacroSteps[$Step][2]]),"$2",$MacroShutDownActions[$MacroSteps[$Step][3]])
			Case $StepTypeComment
				If $MacroSteps[$Step][2] = 1 Then
					$StepString = "- " & $MacroSteps[$Step][1] & " " & _StringRepeat("-",400)
				Else
					$StepString = $MacroSteps[$Step][1]
				EndIf
		EndSwitch
	EndIf
	Return $StepString
EndFunc

Func ResetUndos()
	FileDelete($UndoTempFile & "*")	; delete old undo files
	_StackClear($Undos)
	$Undo = 1
	_StackPush($Undos,$UndoTempFile & $Undo)
	SaveMacro(False,$UndoTempFile & $Undo)
	GUICtrlSetState($MenuItemUndo,$GUI_DISABLE)
	GUICtrlSetState($MenuItemRedo,$GUI_DISABLE)
EndFunc

Func Changed($Change = True)
	$Changed = $Change
	If $Changed Then
		SaveMacro(False,FileFailSave())
		While _StackSize($Undos) > $Undo
			_StackPop($Undos)
		WEnd
		$Undo += 1
		_StackPush($Undos,$UndoTempFile & $Undo)
		SaveMacro(False,$UndoTempFile & $Undo)
		GUICtrlSetState($MenuItemUndo,$GUI_ENABLE)
		GUICtrlSetState($MenuItemRedo,$GUI_DISABLE)
	EndIf
	SetWindowTitle()
EndFunc

Func FileFailSave()
	If StringLen($MacroFile) > 0 Then
		Return _FilePath($MacroFile) & "\" & $FileFailSave & _FileName($MacroFile)
	Else
		Return @MyDocumentsDir & "\" & $FileFailSave & $FileExtension
	EndIf
EndFunc

Func SetWindowTitle()							; show program name/version and macro file in main window title bar
	Local $Title

	$Title = $ProgramName & " " & $Version
	If StringLen($MacroFile) > 0 Then
		$Title &= " - " & $MacroFile
	Else
		$Title &= " - <unsaved new macro>"
	EndIf
	If $Changed Then $Title &= "*"
	WinSetTitle($GUI,"",$Title)
EndFunc

Func FillProcesses()							; fill process list, if user have set, add processed without a window
	Local $PID,$ProcessesList,$ProcessCount,$ProcessName,$Process,$WindowList,$Select,$ProcessIndex

	; get processes with 1 or more windows
	$WindowList = WinList()
	If $WindowList[0][0] = 0 Then Return False
	$ProcessesList = ProcessList()
	If @error Then Return False
	$ProcessCount = -1
	For $Window = 1 To $WindowList[0][0]
		$PID = WinGetProcess($WindowList[$Window][1])
		If $PID <> -1 And StringLen($WindowList[$Window][0]) > 0 And _WindowIsState($WindowList[$Window][1],$WIN_STATE_EXISTS) Then
			$ProcessName = ""
			For $Process = 1 To $ProcessesList[0][0]
				If $ProcessesList[$Process][1] = $PID Then $ProcessName = $ProcessesList[$Process][0]
			Next
			$ProcessCount += 1
			ReDim $ProcessList[$ProcessCount+1][5]
			; add process executable
			$ProcessList[$ProcessCount][0] = $ProcessName
			$ProcessList[$ProcessCount][1] = $WindowList[$Window][0]	; window title
			$ProcessList[$ProcessCount][2] = $WindowList[$Window][1]	; window handle
			$ProcessList[$ProcessCount][3] = $PID
			$ProcessList[$ProcessCount][4] = _WindowIsState($WindowList[$Window][1],$WIN_STATE_VISIBLE)
		EndIf
	Next
	_ArraySort($ProcessList,0,0,0,0,0)	; sort process list
	$Select = 0
	; resolve for selection
	$ProcessIndex = 0
	If $CurrentProcess = -1 Or _GUICtrlComboBox_GetCurSel($Processes) = -1 Then
		$CurrentProcess = $ProcessList[0][3]
		$CurrentWindow = $ProcessList[0][2]
		For $Process = 0 To UBound($ProcessList)-1
			If $ProcessList[$Process][0] = "explorer.exe" And $ProcessList[$Process][1] = "Program Manager" Then
				$Select = $ProcessIndex
				$CurrentProcess = $ProcessList[$Process][3]
				$CurrentWindow = $ProcessList[$Process][2]
				ExitLoop
			EndIf
			If $ShowInvisible Or $ProcessList[$Process][4] Then $ProcessIndex += 1
		Next
	ElseIf $CurrentWindow = -1 And $CurrentProcess > -1 Then
		For $Process = 0 To UBound($ProcessList)-1
			If $ProcessList[$Process][3] = $CurrentProcess Then
				$Select = $ProcessIndex
				$CurrentWindow = $ProcessList[$Process][2]
				ExitLoop
			EndIf
			If $ShowInvisible Or $ProcessList[$Process][4] Then $ProcessIndex += 1
		Next
	Else
		For $Process = 0 To UBound($ProcessList)-1
			If $ProcessList[$Process][3] = $CurrentProcess And $ProcessList[$Process][2] = $CurrentWindow Then
				$Select = $ProcessIndex
				ExitLoop
			EndIf
			If $ShowInvisible Or $ProcessList[$Process][4] Then $ProcessIndex += 1
		Next
	EndIf
	; add to combo
	_GUICtrlComboBox_BeginUpdate($Processes)
	_GUICtrlComboBox_ResetContent($Processes)
	For $Process = 0 To UBound($ProcessList)-1
		If $ShowInvisible Or $ProcessList[$Process][4] Then _GUICtrlComboBox_AddString($Processes,$ProcessList[$Process][0] & " - " & $ProcessList[$Process][1])
	Next
	_GUICtrlComboBox_EndUpdate($Processes)
	_GUICtrlComboBox_SetCurSel($Processes,$Select)
	Return True
EndFunc

Func SelectedProcessIndex()
	If $ShowInvisible Then Return _GUICtrlComboBox_GetCurSel($Processes)	; all processes are shown -> index is equal to combo selection
	; only visible window process are shown -> calculate index
	Local $ProcessIndex  = -1
	For $Process = 0 To UBound($ProcessList)-1
		If $ProcessList[$Process][4] Then $ProcessIndex += 1
		If $ProcessIndex = _GUICtrlComboBox_GetCurSel($Processes) Then Return $Process
	Next
EndFunc

Func ProcessNameById($ProcessId)
	For $Process = 0 To UBound($ProcessList)-1
		If $ProcessList[$Process][3] = $ProcessId then Return $ProcessList[$Process][0]
	Next
	Return ""
EndFunc

Func RandomMinus($MaxValue)
	Return Random(0,Number($MaxValue),1)-Number($MaxValue)/2
EndFunc

Func HideMessage($Message,$HideText,$Icon = False)
	Return _HideMessage($Message,$ProgramName & " - Notice",$HideText,"I get it",$Icon ? @ScriptDir & "\important.ico" : "",0,0,0,$ProgramTheme > -1 ? $Themes[$ProgramTheme][1] : -1,$ProgramTheme > -1 ? $Themes[$ProgramTheme][0] : -1)
EndFunc

Func FilePathTranslateVariables($FileName)
	Return StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace($FileName,"%WINDIR%",@WindowsDir),"%SYSTEMROOT%",@WindowsDir),"%LOCALAPPDATA%",@LocalAppDataDir),"%APPDATA%",@LocalAppDataDir),"%Documents%",@MyDocumentsDir),"%Desktop%",@DesktopDir),"%Home%",@HomeDrive & @HomePath),"%StartMenu%",@StartMenuDir),"%system%",@SystemDir),"%temp%",@TempDir),"%programfiles%",@ProgramFilesDir),"%macronize%",@ScriptDir)
EndFunc

Func WindowWidth($Handle)								; window width
	Local $aWindow = WinGetPos($Handle)
	Return $aWindow[2]
EndFunc

Func WindowHeight($Handle)								; window height
	Local $aWindow = WinGetPos($Handle)
	Return $aWindow[3]
EndFunc