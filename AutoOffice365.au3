#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AutoOffice365.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=GUI For Office Deployment Tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.71
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#NoTrayIcon
;#RequireAdmin

#include <File.au3>
#include <Misc.au3>
#include <Process.au3>
#include <EditConstants.au3>
#Include <GuiEdit.au3>
#include <GuiListBox.au3>
#include <GuiComboBox.au3>
#include <InetConstants.au3>
#include <ListBoxConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "include\External.au3"
#include "include\JSON.au3"

Global $Title = "AutoOffice365"
Global $Version = FileGetVersion(@ScriptFullPath)
Global $TitleVersion = $Title & " v" & StringTrimLeft($Version, StringInStr($Version,".", 0, -1))
Global $LogTitle = $Title
Global $LogWindowStart = 20
Global $LogWindowSize = 700
Global $Path = @TempDir & "\AutoOffice365"
Global $OfficeSetup = "OfficeDeploymentTool.exe"
Global $OfficeSetupFullPath = $Path & "\" & $OfficeSetup
Global $InstallerXML = "OfficeDeploymentTool_"&Random(1000,9999,1)&".xml"
Global $InstallerXMLFullPath = $Path & "\" & $InstallerXML

_Log("Starting " & $Title)

OnAutoItExitRegister("_Exit")


#Region ### START Koda GUI section ###
$Form1 = GUICreate("Title", 408, 194, -1, -1)
$Check_Arch32 = GUICtrlCreateCheckbox("Install 32-bit Version", 16, 13, 121, 17)
$Check_Access = GUICtrlCreateCheckbox("Also Install MS Access", 16, 88, 137, 17)
$Button_Install = GUICtrlCreateButton("Install", 224, 160, 75, 25)
$Button_Cancel = GUICtrlCreateButton("Cancel", 312, 160, 75, 25)
$Combo_Channel = GUICtrlCreateCombo("Current", 248, 48, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL,$CBS_OEMCONVERT))
GUICtrlSetData(-1, "MonthlyEnterprise|SemiAnnual|CurrentPreview|SemiAnnualPreview|BetaChannel")
$Label1 = GUICtrlCreateLabel("Channel", 194, 52, 43, 17)
$Check_Shared = GUICtrlCreateCheckbox("Shared/RDS Licensing Mode", 16, 113, 161, 17)
$Combo_Build = GUICtrlCreateCombo("Latest", 248, 88, 145, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_OEMCONVERT))
GUICtrlSetTip(-1, "Format: xxxxx.yyyyy")
$Label2 = GUICtrlCreateLabel("Build", 209, 92, 27, 17)
GUICtrlSetTip(-1, "Format: xxxxx.yyyyy")
$Combo_ProductID = GUICtrlCreateCombo("O365BusinessRetail", 248, 8, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL,$CBS_OEMCONVERT))
GUICtrlSetData(-1, "O365ProPlusRetail")
$Check_EnableUpdates = GUICtrlCreateCheckbox("Enable Updates", 16, 138, 105, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$Check_ForceClose = GUICtrlCreateCheckbox("Force Close Office", 16, 63, 105, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$Label3 = GUICtrlCreateLabel("Product ID", 184, 12, 55, 17)
$Check_ReplaceArch = GUICtrlCreateCheckbox("Force Change 32-bit/64-bit", 16, 38, 161, 17)
$Label_FetchVersions = GUICtrlCreateLabel("Fetch Versions", 320, 120, 74, 17)
GUICtrlSetFont(-1, 8, 400, 4, "MS Sans Serif")
GUICtrlSetColor(-1, 0x000080)
GUICtrlSetCursor (-1, 0)
$Label_VisitGitHub = GUICtrlCreateLabel("Visit GitHub Page", 16, 170, 87, 17)
GUICtrlSetFont(-1, 8, 400, 4, "MS Sans Serif")
GUICtrlSetColor(-1, 0x000080)
GUICtrlSetCursor (-1, 0)
#EndRegion ### END Koda GUI section ###

WinSetTitle($Form1, "", $TitleVersion)
_GUICtrlComboBox_SetDroppedWidth($Combo_Build, 210)
_GUICtrlComboBox_SetMinVisible($Combo_Build, 15)
GUISetState(@SW_SHOW)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $Button_Cancel
			Exit

		Case $Label_VisitGitHub
			_Log("Opening GitHub page in default browser")
			ShellExecute("https://github.com/jmclaren7/auto-office-365")


		Case $Label_FetchVersions
			_Log("Fetching list of versions")

			$Channel = StringLower(GUICtrlRead($Combo_Channel))
			If $Channel = "MonthlyEnterprise" Then $Channel = "monthly"

			$URL = "https://functions.office365versions.com/api/getjson?name=" & $Channel
			_Log($URL)

			$dData = InetRead($URL, $INET_FORCERELOAD)
			If @error Then
				_Log("Error downloading version information: " & @error)
				MsgBox(0, $Title, "Error downloading version information.")
				GUICtrlSetData($Combo_Build, "")
				GUICtrlSetData($Combo_Build, "Latest", "Latest")
				ContinueLoop
			Endif

			$sJSON = BinaryToString($dData)

			$oJSON = _JSON_Parse($sJSON)
			If @error Then
				_Log("Error parsing JSON: " & @error)
				MsgBox(0, $Title, "Error parsing JSON.")
				GUICtrlSetData($Combo_Build, "")
				GUICtrlSetData($Combo_Build, "Latest", "Latest")
				ContinueLoop
			Endif

			$VersionList = "Latest"

			For $i=0 to 40
				$VersionList &= "|" & $oJSON.data[$i].build & "   (" & $oJSON.data[$i].version & "  -  " & $oJSON.data[$i].releaseDate & ")"

			Next

			GUICtrlSetData($Combo_Build, "")
			GUICtrlSetData($Combo_Build, $VersionList, "Latest")

			_GUICtrlComboBox_ShowDropDown($Combo_Build, True)

		Case $Button_Install
			If FileExists($Path) Then
				_Log("Removing existing temp folder at " & $Path)
				If NOT FileDelete($Path) Then _Log("Remove FAILED")
				Sleep(1000)
			Endif

			_Log("Creating temp folder at " & $Path)
			DirCreate($Path)

			_Log("Unpacking files to temp folder")
			FileInstall(".\Include\OfficeDeploymentTool.exe", $OfficeSetupFullPath, 1)
			FileInstall(".\Include\OfficeDeploymentTool.xml", $InstallerXMLFullPath, 1)

			If Not FileExists($OfficeSetupFullPath) Or Not FileExists($InstallerXMLFullPath) Then
				_Log("Extracted files not found" & $Path)
				MsgBox(0,$Title,"Error extracting files")
				Exit
			EndIf

			; Begin GUI option checks

			If GUICtrlRead($Check_Arch32) = $GUI_CHECKED Then
				_Log("32 bit selected")
				_ReplaceStringInFile($InstallerXMLFullPath, "OfficeClientEdition=""64""", "OfficeClientEdition=""32""")
			Endif


			If GUICtrlRead($Check_ReplaceArch) = $GUI_CHECKED Then
				_Log("Migrate arch selected")
				_ReplaceStringInFile($InstallerXMLFullPath, 'MigrateArch="FALSE"', 'MigrateArch="TRUE"')
			Endif


			If GUICtrlRead($Check_Access) = $GUI_CHECKED Then
				_Log("Access selected")
				_ReplaceStringInFile($InstallerXMLFullPath, "<!--Access", "")
				_ReplaceStringInFile($InstallerXMLFullPath, "Access-->", "")
			Endif


			If GUICtrlRead($Check_Shared) = $GUI_CHECKED Then
				_Log("Shared licnese selected")
				_ReplaceStringInFile($InstallerXMLFullPath, "<!--Shared", "")
				_ReplaceStringInFile($InstallerXMLFullPath, "Shared-->", "")
			Endif


			If GUICtrlRead($Check_EnableUpdates) <> $GUI_CHECKED Then
				_Log("Updates disabled")
				_ReplaceStringInFile($InstallerXMLFullPath, 'Updates Enabled="TRUE"', 'Updates Enabled="FALSE"')
			Endif


			If GUICtrlRead($Check_ForceClose) = $GUI_CHECKED Then
				_Log("Force close selected")
				_ReplaceStringInFile($InstallerXMLFullPath, "<!--ForceShutdown", "")
				_ReplaceStringInFile($InstallerXMLFullPath, "ForceShutdown-->", "")
			Endif

			$ProductID = GUICtrlRead($Combo_ProductID)
			_Log("ProductID: " & $ProductID)
			_ReplaceStringInFile($InstallerXMLFullPath, 'Product ID=""', 'Product ID="'&$ProductID&'"')


			$Channel = GUICtrlRead($Combo_Channel)
			_Log("Channel: " & $Channel)
			_ReplaceStringInFile($InstallerXMLFullPath, 'Channel=""', 'Channel="'&$Channel&'"')


			$Version = GUICtrlRead($Combo_Build)
			_Log("Specified Version: " & $Version)

			; Clean version selection string
			If StringInStr($Version, "(") Then $Version = StringLeft($Version, StringInStr($Version, "(") - 1)
			$Version = StringStripWS($Version, 8)
			_Log("Updated Version: " & $Version)

			; Format is w.x.y.z
			If StringRegExp($Version, "^\d{4,}\.\d{4,}\.\d{4,}\.\d{4,}$") Then
				_Log("Using Target Version w.x.y.z")
				_ReplaceStringInFile($InstallerXMLFullPath, 'TargetVersion=""', 'TargetVersion="'&$Version&'"')

			; Format is y.z
			ElseIf StringRegExp($Version, "^\d{4,}\.\d{4,}$") Then
				_Log("Using Target Version y.z")
				_ReplaceStringInFile($InstallerXMLFullPath, 'TargetVersion=""', 'TargetVersion="16.0.'&$Version&'"')

			; Latest or blank is specified
			Elseif $Version = "Latest" Or $Version = "" Then
				_Log("Latest version selection")
				_ReplaceStringInFile($InstallerXMLFullPath, 'TargetVersion=""', '')

			; Something else we didn't expect was specified
			Else
				_Log("Error in version selection")
				MsgBox(0, $Title, "Error in version selection.")
				ContinueLoop
			EndIf

			ExitLoop
	EndSwitch
WEnd

GUISetState(@SW_HIDE)

_Log("Running Office setup download phase")
_Log("This can take a while and the indicated download progress is not accurate.")
_Log("The blank console that appears is the Office Deployment Tool which doesn't provide any feedback.")
$pid = ShellExecute($OfficeSetupFullPath, "/download "&$InstallerXML, $Path)
$LastDownloadSize = 0
While ProcessExists($pid)
	Sleep(10*1000)
	$DownloadSize = Round(DirGetSize($Path) / 1000 / 1000)
	If $DownloadSize <> $LastDownloadSize Then
		_Log("Download progress: " & $DownloadSize & "MB")
		$LastDownloadSize = $DownloadSize
	EndIf
Wend

Sleep(2*1000)
_Log("Running Office setup configure (install) phase")
ShellExecuteWait($OfficeSetupFullPath, "/configure "&$InstallerXML, $Path)



Func _Exit()
	_Log("Completed, removing temp folder")
	FileDelete($Path)
	DirRemove($Path, 1)
EndFunc
