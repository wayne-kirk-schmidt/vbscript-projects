' HelpPage '
' Version: %%Version%% -- %%CommandLine%% [named-arguments]
' 
'	/verbose	/v		- turn on debugging statements within the script
'	/help		/h		- show this help message and exit
' 
' HelpPage '
'---

'---  
' This script gets the OS version and nothing more. This is boilerplate for:
'	+ Options Processing
'	+ Standard output
'---

' Constants and Variables
'---
Option Explicit
Dim strHelpFile     : strHelpFile   = wscript.ScriptName
Dim strScriptFile   : strScriptFile = wscript.scriptfullname
Dim strComputer     : strComputer = "."
Dim strResults
Dim strVersion		: strVersion = "1.0.0"
Dim strCommandExe	: strCommandExe = "getosversion"

Const HKCR          = &H80000000 'HKEY_CLASSES_ROOT
Const HKCU          = &H80000001 'HKEY_CURRENT_USER
Const HKLM          = &H80000002 'HKEY_LOCAL_MACHINE
Const HKU           = &H80000003 'HKEY_USERS
Const HKCC          = &H80000005 'HKEY_CURRENT_CONFIG

Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7

Const ForReading    = 1
Const ForWriting    = 2
Const ForAppending  = 8

Dim FileSystemObject
Dim objFileSys
Dim objFileData
Dim objFSO             	: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objWMIService      	: Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Dim objRegistry     	: Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

Dim objShell
Dim objExec	
Dim strOSVersion, arrOSVersion, strOSMajor, strOSMinor, strOSBuild
'---

'---
Dim strVerboseFlag		: strVerboseFlag = "no"
Dim colNamedArguments 	: Set colNamedArguments = WScript.Arguments.Named
'---

'---
Dim objRegexHelp 		: Set objRegexHelp = New RegExp	
With objRegexHelp
	  .IgnoreCase = true
	  .Global = true
    .Pattern = "\'\s+HelpPage\s+\'"
End With
'---

'---
Dim objRegexRem 		: Set objRegexRem = New RegExp
With objRegexRem
	  .IgnoreCase = true
	  .Global = true
	  .Pattern = "^(\')(\s+)?*"
End With
'---

' DisplayHelpPage
'---
Function DisplayHelpPage

	  Dim objTextFile       : Set objTextFile = objFSO.OpenTextFile (strScriptFile, ForReading)
	  Dim strScriptContents : strScriptContents = objTextFile.ReadAll
	  objTextFile.Close

	  Dim arrScriptText     : arrScriptText = Split(strScriptContents, vbCrLf)
	  Dim strScriptLine
	  Dim strScriptText
	  Dim strHelpPage  : strHelpPage = ""
	  Dim strPrintHelp : strPrintHelp = 0

	  for each strScriptLine in arrScriptText
		    if strPrintHelp <> 0 then
		        if  objRegexHelp.test(strScriptLine) then
			        strPrintHelp = 0
					Wscript.echo strHelpPage
				    Wscript.quit (0)
			    else
				    strScriptLine = Replace(strScriptLine, "'", "   ", 1, 1)
				    strScriptLine = Replace(strScriptLine, "%%CommandLine%%", strHelpFile, 1, 1)
					strScriptLine = Replace(strScriptLine, "%%Version%%", strVersion, 1, 1)
				    strHelpPage = strHelpPage & vbCrLf & strScriptLine
			    end if	
		    end if
		    if objRegexHelp.test(strScriptLine) then
		        strPrintHelp = 1
		    end if
	  next

End Function
'---

'---
function ProcessArguments

    If colNamedArguments.Exists("help") or colNamedArguments.Exists("h") Then
		DisplayHelpPage
	end if

    If colNamedArguments.Exists("verbose") or colNamedArguments.Exists("v") Then
		strVerboseFlag="yes"
	end if
	
End function
'---

' GetOsVersion
'---
Function GetOsVersion
	Dim oss : Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	Dim os, dtmInstallDate, osVersion
	For Each os in oss
		strOSVersion = os.Version
	Next
	arrOSVersion = split (strOSVersion, ".")
	strOSMajor = CInt(arrOSVersion(0))
	strOSMinor = CInt(arrOSVersion(1))
	strOSBuild = CInt(arrOSVersion(2))
	WScript.Echo "strOSVersion:" & vbTab & strOSVersion 
	WScript.Echo "strOSBuild: "	 & vbTab & strOSBuild 
	WScript.Echo "strOSMajor: "  & vbTab & strOSMajor
	WScript.Echo "strOSMinor: "  & vbTab & strOSMinor
End function
'---

' MainLogic
'---
sub MainLogic
	
	ProcessArguments
	GetOsVersion

end sub
'---

MainLogic