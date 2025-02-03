#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AutoOffice365.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=n
#AutoIt3Wrapper_Res_Comment=https://github.com/jmclaren7/auto-office-365
#AutoIt3Wrapper_Res_Description=GUI For Office Deployment Tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.103
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=AutoOffice365
#AutoIt3Wrapper_Res_ProductVersion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=© John McLaren
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin


#include <File.au3>
#include <Misc.au3>
#include <Process.au3>
#include <EditConstants.au3>
#include <GuiEdit.au3>
#include <GuiListBox.au3>
#include <GuiComboBox.au3>
#include <InetConstants.au3>
#include <ListBoxConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

#include "include\JSON.au3"
#include <include\CommonFunctions.au3>
#include <include\Console.au3>

Global $Title = "AutoOffice365"
Global $Version = FileGetVersion(@ScriptFullPath)
Global $TitleVersion = $Title & " v" & StringTrimLeft($Version, StringInStr($Version, ".", 0, -1))
Global $TempPath = @TempDir & "\AutoOffice365"
Global $OfficeSetup = "OfficeDeploymentTool.exe"
Global $OfficeSetupFullPath = $TempPath & "\" & $OfficeSetup
Global $InstallerXML = "OfficeDeploymentTool_" & Random(1000, 9999, 1) & ".xml"
Global $InstallerXMLFullPath = $TempPath & "\" & $InstallerXML
Global $DownloadPID
Global $InstallPID
Global $XMLData = ""

; Setup Logging
_Console_Attach() ; If it was launched from a console, attach to that console
Global $LogFileMaxSize = 512
Global $LogLevel = 1
If @Compiled Then
	Global $LogFullPath = @TempDir & "\" & StringTrimRight(@ScriptName, 4) & ".log"
	_Console_Alloc()
Else
	$LogLevel = 3
	Global $LogTitle = $Title
	Global $LogWindowStart = 20
	Global $LogWindowSize = 700
EndIf

_Log("Starting " & $Title)

OnAutoItExitRegister("_Exit")

#Region ### START Koda GUI section ###
$Form1 = GUICreate("Title", 408, 194, -1, -1)
$Check_Arch32 = GUICtrlCreateCheckbox("Install 32-bit Version", 16, 13, 121, 17)
$Check_Access = GUICtrlCreateCheckbox("Also Install MS Access", 16, 88, 137, 17)
$Button_Install = GUICtrlCreateButton("Install", 224, 160, 75, 25)
$Button_Cancel = GUICtrlCreateButton("Cancel", 312, 160, 75, 25)
$Combo_Channel = GUICtrlCreateCombo("", 248, 48, 145, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL, $CBS_OEMCONVERT))
$Label1 = GUICtrlCreateLabel("Channel", 194, 52, 43, 17)
$Check_Shared = GUICtrlCreateCheckbox("Shared/RDS Licensing Mode", 16, 113, 161, 17)
$Combo_Build = GUICtrlCreateCombo("", 248, 88, 145, 25, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_OEMCONVERT))
GUICtrlSetTip(-1, "Format: xxxxx.yyyyy")
$Label2 = GUICtrlCreateLabel("Build", 209, 92, 27, 17)
GUICtrlSetTip(-1, "Format: xxxxx.yyyyy")
$Combo_ProductID = GUICtrlCreateCombo("", 248, 8, 145, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL, $CBS_OEMCONVERT))
$Check_EnableUpdates = GUICtrlCreateCheckbox("Enable Updates", 16, 138, 105, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$Check_ForceClose = GUICtrlCreateCheckbox("Force Close Office", 16, 63, 105, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$Label3 = GUICtrlCreateLabel("Product ID", 184, 12, 55, 17)
$Check_ReplaceArch = GUICtrlCreateCheckbox("Force Change 32-bit/64-bit", 16, 38, 161, 17)
$Label_FetchVersions = GUICtrlCreateLabel("Fetch Builds", 320, 120, 62, 17)
GUICtrlSetFont(-1, 8, 400, 4, "MS Sans Serif")
GUICtrlSetColor(-1, 0x000080)
GUICtrlSetCursor(-1, 0)
$Label_VisitGitHub = GUICtrlCreateLabel("Visit GitHub Page", 16, 170, 87, 17)
GUICtrlSetFont(-1, 8, 400, 4, "MS Sans Serif")
GUICtrlSetColor(-1, 0x000080)
GUICtrlSetCursor(-1, 0)
$Check_VersionUpdate = GUICtrlCreateCheckbox("Also Use Channel Selection for Updates", 120, 138, 245, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

WinSetTitle($Form1, "", $TitleVersion)

GUICtrlSetData($Combo_ProductID, "O365BusinessRetail|O365ProPlusRetail|O365BusinessEEANoTeamsRetail|O365ProPlusEEANoTeamsRetail", "O365BusinessRetail")
_GUICtrlComboBox_SetDroppedWidth($Combo_ProductID, 200)
GUICtrlSetData($Combo_Channel, "CurrentPreview|Current|SemiAnnualPreview|SemiAnnual|PerpetualVL2019|PerpetualVL2021|PerpetualVL2024|MonthlyEnterprise", "Current")
GUICtrlSetData($Combo_Build, "Latest", "Latest")
_GUICtrlComboBox_SetDroppedWidth($Combo_Build, 250)
_GUICtrlComboBox_SetMinVisible($Combo_Build, 15)

GUISetState(@SW_SHOW, $Form1)

$SelectedChannel = GUICtrlRead($Combo_Channel)
$SelectedChannelLast = $SelectedChannel

While 1
	$nMsg = GUIGetMsg()
	If StringInStr(@ScriptName, "[silent]") Then $nMsg = $Button_Install
	If $XMLData <> "" And _GUICtrlComboBox_GetDroppedState($Combo_Channel) = False Then
		$SelectedChannelLast = $SelectedChannel
		$SelectedChannel = GUICtrlRead($Combo_Channel)
		If $SelectedChannelLast <> $SelectedChannel Then $nMsg = $Label_FetchVersions
	EndIf

	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $Button_Cancel
			GUISetState(@SW_HIDE, $Form1)
			Exit

		Case $Label_VisitGitHub
			_Log("Opening GitHub page in default browser")
			ShellExecute("https://github.com/jmclaren7/auto-office-365")


		Case $Label_FetchVersions
			_Log("Fetching list of versions")
			GUISetCursor($MCID_WAIT, 1)

			$XMLFile = $TempPath & "\365ReleaseHistoryLocalCache.xml"
			$URL = "https://raw.githubusercontent.com/jmclaren7/auto-office-365/refs/heads/main/365ReleaseHistory.xml"

			$XMLData = FileRead($XMLFile)
			If @error Then
				_Log("Downloading file to " & $XMLFile)
				$XMLData = InetRead($URL, $INET_FORCERELOAD)
				$XMLData = BinaryToString($XMLData, $SB_UTF8)
				If @error Then
					_Error($Title, "Error downloading version information.")
					Exit
				EndIf
				FileDelete($XMLFile)
				DirCreate($TempPath)
				FileWrite($XMLFile, $XMLData)
			Else
				_Log("Using existing " & $XMLFile)
			EndIf

			$oXML = ObjCreate("Microsoft.XMLDOM")
			If @error Then
				_Error($Title, "XML error.", "Failed to create XMLDOM object")
				Exit
			EndIf

			; Fix unknown issue related to binary conversion and encoding
			$XMLData = StringTrimLeft($XMLData, StringInStr($XMLData, "<") - 1)

			$oXML.loadXML($XMLData)
			If $oXML.parseError.errorCode <> 0 Then
				;_Log("==========" & $XMLData & "==========")
				_Error($Title, "Error reading version information.", "XML Parse Error: " & $oXML.parseError.reason)
				Exit
			EndIf

			; Get all UpdateChannel nodes
			Local $oChannels = $oXML.selectNodes("/ReleaseHistory/UpdateChannel")
			Local $sChannelList = ""

			For $oChannel In $oChannels
				If $sChannelList <> "" Then $sChannelList &= "|"
				$sChannelList &= $oChannel.getAttribute("ID")
			Next
			_Log($sChannelList)

			$SelectedChannel = GUICtrlRead($Combo_Channel)
			If Not StringInStr($sChannelList & "|", $SelectedChannel & "|") Then $SelectedChannel = "Current"
			GUICtrlSetData($Combo_Channel, "")
			GUICtrlSetData($Combo_Channel, $sChannelList)
			_GUICtrlComboBox_SetCurSel($Combo_Channel, _GUICtrlComboBox_FindStringExact($Combo_Channel, $SelectedChannel))
			_GUICtrlComboBox_SetEditSel($Combo_Channel, 0, 0)

			; Get all update nodes for the selected channel
			Local $oUpdates = $oXML.selectNodes("/ReleaseHistory/UpdateChannel[@ID='" & $SelectedChannel & "']/Update")
			Local $sVersionList = "Latest"
			For $oUpdate In $oUpdates
				$Date = $oUpdate.getAttribute("PubTime")
				$Date = StringLeft($Date, StringInStr($Date, "T") - 1) ; Source: 2024-11-12T08:32:14.997Z

				$sVersionList &= "|" & $oUpdate.getAttribute("Build") & "   (" & $oUpdate.getAttribute("Version") & "  -  " & $Date & ")"
			Next

			_Log($sVersionList)

			GUICtrlSetData($Combo_Build, "")
			GUICtrlSetData($Combo_Build, $sVersionList, "Latest")
			_GUICtrlComboBox_SetEditSel($Combo_Build, 0, 0)
			_GUICtrlComboBox_ShowDropDown($Combo_Build, True)

			GUISetCursor()

		Case $Button_Install
			;If FileExists($TempPath) Then
			;	_Log("Removing existing temp folder at " & $TempPath)
			;	If Not FileDelete($TempPath) Then _Log("Remove FAILED")
			;	Sleep(1000)
			;EndIf

			If FileDelete($OfficeSetupFullPath) = 0 Then _Log("Could not remove ODT")
			If FileDelete($InstallerXMLFullPath) = 0 Then _Log("Could not remove XML")

			_Log("Creating temp folder at " & $TempPath)
			DirCreate($TempPath)

			_Log("Unpacking files to temp folder")
			FileInstall(".\Include\OfficeDeploymentTool.exe", $OfficeSetupFullPath, 1)
			FileInstall(".\Include\OfficeDeploymentTool.xml", $InstallerXMLFullPath, 1)

			If Not FileExists($OfficeSetupFullPath) Or Not FileExists($InstallerXMLFullPath) Then
				_Error($Title, "Error extracting files", Default, $TempPath)
				Exit
			EndIf

			; === Begin GUI option checks ===================
			If GUICtrlRead($Check_Arch32) = $GUI_CHECKED Then
				_Log("32 bit selected")
				_ReplaceStringInFile($InstallerXMLFullPath, "OfficeClientEdition=""64""", "OfficeClientEdition=""32""")
			EndIf

			If GUICtrlRead($Check_ReplaceArch) = $GUI_CHECKED Then
				_Log("Migrate arch selected")
				_ReplaceStringInFile($InstallerXMLFullPath, 'MigrateArch="FALSE"', 'MigrateArch="TRUE"')
			EndIf

			If GUICtrlRead($Check_Access) = $GUI_CHECKED Then
				_Log("Access selected")
				_ReplaceStringInFile($InstallerXMLFullPath, "<!--Access", "")
				_ReplaceStringInFile($InstallerXMLFullPath, "Access-->", "")
			EndIf

			If GUICtrlRead($Check_Shared) = $GUI_CHECKED Then
				_Log("Shared licnese selected")
				_ReplaceStringInFile($InstallerXMLFullPath, '"SharedComputerLicensing" Value="0"', '"SharedComputerLicensing" Value="1"')
			EndIf

			If GUICtrlRead($Check_EnableUpdates) <> $GUI_CHECKED Then
				_Log("Updates disabled")
				_ReplaceStringInFile($InstallerXMLFullPath, 'Updates Enabled="TRUE"', 'Updates Enabled="FALSE"')
			EndIf

			If GUICtrlRead($Check_ForceClose) = $GUI_CHECKED Then
				_Log("Force close selected")
				_ReplaceStringInFile($InstallerXMLFullPath, '"FORCEAPPSHUTDOWN" Value="FALSE"', '"FORCEAPPSHUTDOWN" Value="TRUE"')
			EndIf

			$ProductID = GUICtrlRead($Combo_ProductID)
			_Log("ProductID: " & $ProductID)
			_ReplaceStringInFile($InstallerXMLFullPath, 'Product ID=""', 'Product ID="' & $ProductID & '"')

			$Channel = GUICtrlRead($Combo_Channel)
			_Log("Channel: " & $Channel)
			_ReplaceStringInFile($InstallerXMLFullPath, 'Channel="Current"', 'Channel="' & $Channel & '"')

			If GUICtrlRead($Check_VersionUpdate) = $GUI_CHECKED Then
				_ReplaceStringInFile($InstallerXMLFullPath, 'Channel="Updates"', 'Channel="' & $Channel & '"')
			Else
				_ReplaceStringInFile($InstallerXMLFullPath, 'Channel="Updates"', '')
			EndIf

			; Build/Version
			$Version = GUICtrlRead($Combo_Build)
			_Log("Specified Version: " & $Version)

			If StringInStr($Version, "(") Then $Version = StringLeft($Version, StringInStr($Version, "(") - 1)
			$Version = StringStripWS($Version, 8)
			_Log("Updated Version: " & $Version)

			If StringRegExp($Version, "^\d{4,}\.\d{4,}\.\d{4,}\.\d{4,}$") Then ; Format is w.x.y.z
				_Log("Detected w.x.y.z")
				_ReplaceStringInFile($InstallerXMLFullPath, 'Version=""', 'Version="' & $Version & '"')

			ElseIf StringRegExp($Version, "^\d{4,}\.\d{4,}$") Then ; Format is y.z
				_Log("Detected y.z")
				_ReplaceStringInFile($InstallerXMLFullPath, 'Version=""', 'Version="16.0.' & $Version & '"')

			ElseIf $Version = "Latest" Or $Version = "" Then ; Latest or blank is specified
				_Log("Latest version selection")
				_ReplaceStringInFile($InstallerXMLFullPath, 'Version=""', '')

			Else ; Something else we didn't expect was specified
				_Log("Error in version selection")
				MsgBox(0, $Title, "Error in version selection.")
				ContinueLoop
			EndIf

			ExitLoop
	EndSwitch

	Sleep(10)
WEnd

GUISetState(@SW_HIDE)

If Not @Compiled And MsgBox(0 + 4, $Title, "Paused, Continue?") <> $IDYES Then Exit

_Log("Running Office setup download phase")
_Log("This can take a while and the indicated download progress is not accurate.")
_Log(" ")
$DownloadPID = ShellExecute($OfficeSetupFullPath, "/download " & $InstallerXML, $TempPath, Default, @SW_HIDE)

$LastDownloadSize = 0
$Progress = 1
AdlibRegister("_FileFlush", 3000)
While ProcessExists($DownloadPID)
	$DownloadSize = Round(DirGetSize($TempPath) / 1000 / 1000)
	If $DownloadSize <> $LastDownloadSize Then
		$LastDownloadSize = $DownloadSize
	EndIf

	Switch $Progress
		Case 1
			$ProgressMsg = "|  (" & $DownloadSize & "MB)"
		Case 2
			$ProgressMsg = "/  (" & $DownloadSize & "MB)"
		Case 3
			$ProgressMsg = "-  (" & $DownloadSize & "MB)"
		Case Else
			$ProgressMsg = "\  (" & $DownloadSize & "MB)"
			$Progress = 0
	EndSwitch
	$Progress += 1
	_Log($ProgressMsg, Default, True)

	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $Button_Cancel
			GUISetState(@SW_HIDE, $Form1)
			Exit
	EndSwitch

	Sleep(1 * 1000)
WEnd

Sleep(2 * 1000)

_Log("Running Office setup configure (install) phase")
$InstallPID = ShellExecuteWait($OfficeSetupFullPath, "/configure " & $InstallerXML, $TempPath)

;=====================================================================================
;=====================================================================================

Func _FileFlush()
	$aFiles = _FileListToArrayRec($TempPath, "*", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
	If @error Then _Log("Error flushing files")
	For $i = 1 To UBound($aFiles) - 1
		FileRead($aFiles[$i], 1)
	Next
EndFunc   ;==>_FileFlush

Func _Exit()
	_Log("Completed, removing temp folder")
	ProcessClose($DownloadPID)
	ProcessClose($InstallPID)
	FileDelete($TempPath)
	DirRemove($TempPath, 1)
EndFunc   ;==>_Exit
