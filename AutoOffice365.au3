#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AutoOffice365.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=n
#AutoIt3Wrapper_Res_Comment=https://github.com/jmclaren7/auto-office-365
#AutoIt3Wrapper_Res_Description=GUI For Office Deployment Tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.104
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

#include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiListBox.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <InetConstants.au3>
#include <ListBoxConstants.au3>
#include <ListViewConstants.au3>
#include <Misc.au3>
#include <Process.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

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
$Form1 = GUICreate("Title", 420, 292, -1, -1)
$Check_Arch32 = GUICtrlCreateCheckbox("Install 32-bit Version", 16, 13, 121, 17)
$Button_Install = GUICtrlCreateButton("Install", 240, 256, 75, 25)
$Button_Cancel = GUICtrlCreateButton("Cancel", 332, 256, 75, 25)
$Combo_Channel = GUICtrlCreateCombo("", 236, 40, 173, 25, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_OEMCONVERT))
$Label1 = GUICtrlCreateLabel("Channel", 182, 44, 43, 17)
$Check_Shared = GUICtrlCreateCheckbox("Shared/RDS Licensing Mode", 16, 88, 161, 17)
$Combo_Build = GUICtrlCreateCombo("", 236, 72, 173, 25, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_OEMCONVERT))
GUICtrlSetTip(-1, "Format: xxxxx.yyyyy")
$Label2 = GUICtrlCreateLabel("Build", 197, 76, 27, 17)
GUICtrlSetTip(-1, "Format: xxxxx.yyyyy")
$Combo_ProductID = GUICtrlCreateCombo("", 236, 8, 173, 25, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_OEMCONVERT))
$Check_EnableUpdates = GUICtrlCreateCheckbox("Enable Updates", 16, 113, 105, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$Check_ForceClose = GUICtrlCreateCheckbox("Force Close Office", 16, 63, 105, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$Label3 = GUICtrlCreateLabel("Product ID", 172, 12, 55, 17)
$Check_ReplaceArch = GUICtrlCreateCheckbox("Force Change 32-bit/64-bit", 16, 38, 161, 17)
$Label_VisitGitHub = GUICtrlCreateLabel("Visit GitHub Page", 12, 266, 87, 17)
GUICtrlSetFont(-1, 8, 400, 4, "MS Sans Serif")
GUICtrlSetColor(-1, 0x000080)
GUICtrlSetCursor(-1, 0)
$Check_VersionUpdate = GUICtrlCreateCheckbox("Apply Channel to Updates", 124, 113, 153, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$List_Exclude = GUICtrlCreateListView("", 128, 140, 278, 108, BitOR($GUI_SS_DEFAULT_LISTVIEW, $LVS_NOCOLUMNHEADER), BitOR($WS_EX_CLIENTEDGE, $LVS_EX_CHECKBOXES, $LVS_EX_TRACKSELECT))
$Label4 = GUICtrlCreateLabel("Exclude Apps", 32, 180, 69, 17)
$Check_ListBuilds = GUICtrlCreateCheckbox("Fetch List of Builds", 300, 100, 113, 17)
#EndRegion ### END Koda GUI section ###

WinSetTitle($Form1, "", $TitleVersion)

GUICtrlSetData($Combo_ProductID, "O365BusinessRetail|O365ProPlusRetail|O365BusinessEEANoTeamsRetail|O365ProPlusEEANoTeamsRetail", "O365BusinessRetail")
_GUICtrlComboBox_SetDroppedWidth($Combo_ProductID, 200)
_GUICtrlComboBox_SetEditSel($Combo_ProductID, 0, 0)

GUICtrlSetData($Combo_Channel, "CurrentPreview|Current|SemiAnnualPreview|SemiAnnual|PerpetualVL2019|PerpetualVL2021|PerpetualVL2024|MonthlyEnterprise", "Current")

GUICtrlSetData($Combo_Build, "Latest", "Latest")
_GUICtrlComboBox_SetDroppedWidth($Combo_Build, 250)
_GUICtrlComboBox_SetMinVisible($Combo_Build, 15)
_GUICtrlComboBox_SetEditSel($Combo_Build, 0, 0)

; Exclude apps list view setup
$aApps = StringSplit("Access|Bing*|Excel|Groove (Legacy OneDrive)*|Lync (Skype for Business)*|OneDrive|OneNote|Outlook|OutlookForWindows (New Outlook)*|PowerPoint|Publisher|Teams|Word", "|", 2)
_GUICtrlListView_AddColumn($List_Exclude, " ")
For $i = 0 To UBound($aApps) - 1 ; Rows
	$ListIndex = _GUICtrlListView_AddItem($List_Exclude, StringReplace($aApps[$i], "*", ""))
	If StringInStr($aApps[$i], "*") Then _GUICtrlListView_SetItemChecked($List_Exclude, $ListIndex, True)
Next
_GUICtrlListView_SetColumnWidth($List_Exclude, 0, $LVSCW_AUTOSIZE)

GUISetState(@SW_SHOW, $Form1)

$SelectedChannel = GUICtrlRead($Combo_Channel)
$SelectedChannelLast = $SelectedChannel

While 1
	$nMsg = GUIGetMsg()

	If StringInStr(@ScriptName, "[silent]") Then $nMsg = $Button_Install
	If $XMLData <> "" And _GUICtrlComboBox_GetDroppedState($Combo_Channel) = False Then
		$SelectedChannelLast = $SelectedChannel
		$SelectedChannel = GUICtrlRead($Combo_Channel)
		If $SelectedChannelLast <> $SelectedChannel Then $nMsg = $Check_ListBuilds ;$Label_FetchVersions
	EndIf

	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $Button_Cancel
			GUISetState(@SW_HIDE, $Form1)
			Exit

		Case $Combo_Build, $Combo_Channel, $Combo_ProductID
			_GUICtrlComboBox_SetEditSel($Combo_ProductID, 0, 0)
			_GUICtrlComboBox_SetEditSel($Combo_Channel, 0, 0)
			_GUICtrlComboBox_SetEditSel($Combo_Build, 0, 0)

		Case $Label_VisitGitHub
			_Log("Opening GitHub page in default browser")
			ShellExecute("https://github.com/jmclaren7/auto-office-365")

		Case $Check_ListBuilds
			_Log("Fetching list of versions")
			If GUICtrlRead($Check_ListBuilds) = $GUI_UNCHECKED Then ContinueLoop

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
				_Error($Title, "XML setup error.", Default, "Failed to create XMLDOM object")
				Exit
			EndIf

			; Fix unknown issue related to binary conversion and encoding
			$XMLData = StringTrimLeft($XMLData, StringInStr($XMLData, "<") - 1)

			$oXML.loadXML($XMLData)
			If $oXML.parseError.errorCode <> 0 Then
				;_Log("==========" & $XMLData & "==========")
				_Error($Title, "Error loading version information.", Default, "XML Parse Error: " & $oXML.parseError.reason)
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

				$sVersionList &= "|" & $oUpdate.getAttribute("Build") & "  (" & $oUpdate.getAttribute("Version") & "  " & $Date & ")"
			Next

			_Log($sVersionList)

			GUICtrlSetData($Combo_Build, "")
			GUICtrlSetData($Combo_Build, $sVersionList, "Latest")
			_GUICtrlComboBox_ShowDropDown($Combo_Build, True)

			GUISetCursor()

		Case $Button_Install
			If FileExists($OfficeSetupFullPath) And FileDelete($OfficeSetupFullPath) = 0 Then _Log("Could not remove ODT")

			_Log("Creating temp folder at " & $TempPath)
			DirCreate($TempPath)

			_Log("Unpacking ODT")
			FileInstall(".\Include\OfficeDeploymentTool.exe", $OfficeSetupFullPath, 1)

			If Not FileExists($OfficeSetupFullPath) Then
				_Error($Title, "Error extracting files", Default, $TempPath)
				Exit
			EndIf

			; Create XML Document
			$oXML = ObjCreate("Microsoft.XMLDOM")
			If @error Then Exit MsgBox(16, "Error", "Failed to create XMLDOM")

			; Create root element
			$rootNode = $oXML.createElement("Configuration")
			$oXML.appendChild($rootNode)

			; Create Add node
			$addNode = $oXML.createElement("Add")

			; Arch attribute
			If GUICtrlRead($Check_Arch32) = $GUI_CHECKED Then
				_Log("32 bit selected")
				$addNode.setAttribute("OfficeClientEdition", "32")
			EndIf

			; Channel attribute
			$Channel = GUICtrlRead($Combo_Channel)
			_Log("Channel: " & $Channel)
			$addNode.setAttribute("Channel", $Channel)

			; Version attribute
			$Version = GUICtrlRead($Combo_Build)
			_Log("Specified Version: " & $Version)

			If StringInStr($Version, "(") Then $Version = StringLeft($Version, StringInStr($Version, "(") - 1)
			$Version = StringStripWS($Version, 8)
			_Log("Updated Version: " & $Version)

			If StringRegExp($Version, "^\d{4,}\.\d{4,}\.\d{4,}\.\d{4,}$") Then ; Format is w.x.y.z
				_Log("Detected w.x.y.z")
				$addNode.setAttribute("Version", $Version)

			ElseIf StringRegExp($Version, "^\d{4,}\.\d{4,}$") Then ; Format is y.z
				_Log("Detected y.z")
				$addNode.setAttribute("Version", '16.0.' & $Version)

			ElseIf $Version = "Latest" Or $Version = "" Then ; Latest or blank is specified
				_Log("Latest version selection")

			Else ; Something else we didn't expect was specified
				_Log("Error in version selection")
				MsgBox(0, $Title, "Error in version selection.")
				ContinueLoop
			EndIf

			; MigrateArch attribute
			If GUICtrlRead($Check_ReplaceArch) = $GUI_CHECKED Then
				_Log("Migrate arch selected")
				_ReplaceStringInFile($InstallerXMLFullPath, 'MigrateArch="FALSE"', 'MigrateArch="TRUE"')
				$addNode.setAttribute("MigrateArch", "TRUE")
			EndIf

			$rootNode.appendChild($addNode)

			; Add Product node
			$ProductID = GUICtrlRead($Combo_ProductID)
			_Log("ProductID: " & $ProductID)
			$productNode = $oXML.createElement("Product")
			$productNode.setAttribute("ID", $ProductID)
			$addNode.appendChild($productNode)

			; Exclude apps
			For $i = 0 To _GUICtrlListView_GetItemCount($List_Exclude) - 1
				If _GUICtrlListView_GetItemChecked($List_Exclude, $i) Then
					$ItemText = _GUICtrlListView_GetItemText($List_Exclude, $i)
					If StringInStr($ItemText, " (") Then $ItemText = StringLeft($ItemText, StringInStr($ItemText, " (") - 1)
					_Log("Excluding: " & $ItemText)
					$ExcludeNode = $oXML.createElement("ExcludeApp")
					$ExcludeNode.setAttribute("ID", $ItemText)
					$productNode.appendChild($ExcludeNode)
				EndIf
			Next

			; Create Language node
			$langNode = $oXML.createElement("Language")
			$langNode.setAttribute("ID", "en-us")
			$productNode.appendChild($langNode)


			; SharedComputerLicensing
			If GUICtrlRead($Check_Shared) = $GUI_CHECKED Then
				_Log("Shared licnese selected")
				$propNode = $oXML.createElement("Property")
				$propNode.setAttribute("Name", "SharedComputerLicensing")
				$propNode.setAttribute("Value", "1")
				$rootNode.appendChild($propNode)
			EndIf

			If GUICtrlRead($Check_ForceClose) = $GUI_CHECKED Then
				_Log("Force close selected")
				$propNode = $oXML.createElement("Property")
				$propNode.setAttribute("Name", "FORCEAPPSHUTDOWN")
				$propNode.setAttribute("Value", "TRUE")
				$rootNode.appendChild($propNode)
			EndIf

			; Create Updates node
			$updatesNode = $oXML.createElement("Updates")

			; EnableUpdates attribute
			If GUICtrlRead($Check_EnableUpdates) = $GUI_UNCHECKED Then
				_Log("Updates disabled")
				$updatesNode.setAttribute("Enabled", "FALSE")
			Else
				$updatesNode.setAttribute("Enabled", "TRUE")
			EndIf

			; Updates Channel attribute
			If GUICtrlRead($Check_VersionUpdate) = $GUI_CHECKED Then
				_Log("Updates channel enabled: " & $Channel)
				$updatesNode.setAttribute("Channel", $Channel)
			EndIf

			; Add the updates node
			$rootNode.appendChild($updatesNode)

			; Add Display node
			$displayNode = $oXML.createElement("Display")
			$displayNode.setAttribute("AcceptEULA", "TRUE")
			$rootNode.appendChild($displayNode)

			; Save the XML file
			FileDelete($InstallerXMLFullPath)
			FileWrite($InstallerXMLFullPath, $oXML.xml)

			ExitLoop
	EndSwitch

	Sleep(10)
WEnd

GUISetState(@SW_HIDE)

If Not @Compiled Then
	ShellExecute($TempPath)
	If MsgBox(0 + 4, $Title, "Paused, Continue?") <> $IDYES Then Exit
EndIf

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
$InstallPID = ShellExecuteWait($OfficeSetupFullPath, "/configure " & $InstallerXML, $TempPath, Default, @SW_HIDE)
$DownloadPID = ""
$InstallPID = ""

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
