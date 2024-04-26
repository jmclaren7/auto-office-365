#include-once
;===============================================================================
; Function Name:   	_Log()
; Description:		Console & File Loging
; Call With:		_Log($Text, $iLevel)
; Parameter(s): 	$sMessage - Text to print
;					$iLevel - The level *this* message
;								Use 1 for critical or always shown (default), 2+ for debuging
;
; Return Value(s):  The original message, if $iLevel is greater than $LogLevel returns an empty string
; Notes:			Some options are configured with global variables
; Author(s):        JohnMC - JohnsCS.com
; Date/Last Change:	4/26/2024 -- Fixed global handling, added minimize window on start
;===============================================================================
; Write to the log, prepend a timestamp, create a custom log GUI
Func _Log($sMessage, $iLevel = 1)
	Static Local $_hLogFile

	; Global options
	Global $LogLevel, $LogTitle, $LogWindowStart, $LogWindowSize, $LogFullPath, $LogFileMaxSize, $LogFlushAlways


	; If $LogTitle is empty, skip the GUI
	If $LogLevel = "" Then $LogLevel = 1 ; Only show messages this level or below
	If $LogWindowStart = "" Then Global $LogWindowStart = -1 ; -1 for center, -# for minimized with position being the absolute value
	If $LogWindowSize = "" Then Global $LogWindowSize = 750 ; Starting width, height will be .6 of this value
	If $LogFullPath = "" Then Global $LogFullPath = "" ; The path of the log file, empty value will not log to file
	If $LogFileMaxSize = "" Then Global $LogFileMaxSize = 1024 ; Size limit for log in KB
	If $LogFlushAlways = "" Then Global $LogFlushAlways = False ; Flush log to disk after each update

	Local $LogFileMaxSize_Bytes = $LogFileMaxSize * 1024
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> "
	Local $sLogLine = $sTime & $sMessage
	Local $Minimize = False

	; Do not log this message if $iLevel is greater than global $LogLevel
	If $iLevel > $LogLevel Then Return ""

	; Send to console
	ConsoleWrite($sLogLine & @CRLF)

	; Append message to custom GUI if $LogTitle is set
	If $LogTitle <> "" Then
		If Not IsDeclared("_hLogEdit") Then
			; The GUI doesn't exist, create it
			If $LogWindowStart < -1 Then
				$LogWindowStart = Abs($LogWindowStart)
				$Minimize = True
			EndIf
			Global $_hLogWindow = GUICreate($LogTitle, $LogWindowSize, Round($LogWindowSize * 0.6), $LogWindowStart, $LogWindowStart, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX))
			Global $_hLogEdit = GUICtrlCreateEdit("", 0, 0, $LogWindowSize, Round($LogWindowSize * 0.6), BitOR($ES_MULTILINE, $ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL))
			GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
			GUICtrlSetColor(-1, 0xFFFFFF)
			GUICtrlSetBkColor(-1, 0x000000)
			GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
			GUISetState(@SW_SHOW, $_hLogWindow)
			If $Minimize Then GUISetState(@SW_MINIMIZE, $_hLogWindow)
			_GUICtrlEdit_AppendText($_hLogEdit, $sLogLine)
		Else
			; Update an existing GUI
			_GUICtrlEdit_BeginUpdate($_hLogEdit)
			_GUICtrlEdit_AppendText($_hLogEdit, @CRLF & $sLogLine)
			_GUICtrlEdit_LineScroll($_hLogEdit, - StringLen($sLogLine), _GUICtrlEdit_GetLineCount($_hLogEdit))
			_GUICtrlEdit_EndUpdate($_hLogEdit)
		EndIf
	EndIf

	; Append message to file
	If $LogFullPath <> "" Then
		If $_hLogFile = "" Then $_hLogFile = FileOpen($LogFullPath, $FO_APPEND)

		; Limit log size
		If $LogFileMaxSize > 0 Then
			Local $iCurrentSize = FileGetPos($_hLogFile) ; + StringLen($sLogLine)

			If $iCurrentSize > $LogFileMaxSize_Bytes Then
				; Rewrite desired data to begining of file, drop trailing data, flush to disk.
				FileSetPos($_hLogFile, 0, $FILE_BEGIN)
				$sLogLine = FileRead($_hLogFile) & $sLogLine
				$sLogLine = StringRight($sLogLine, $LogFileMaxSize_Bytes - 1024)
				FileSetPos($_hLogFile, 0, $FILE_BEGIN)
				FileWrite($_hLogFile, $sLogLine & @CRLF)
				FileSetEnd($_hLogFile)
				FileFlush($_hLogFile)

			Else
				; Write to file
				FileWrite($_hLogFile, $sLogLine & @CRLF)
				If $LogFlushAlways Then FileFlush($_hLogFile)

			EndIf

		EndIf

	EndIf

	Return $sMessage
EndFunc   ;==>_Log