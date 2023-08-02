#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\ohiohealth.ico
#AutoIt3Wrapper_Outfile_x64=..\UserBackup.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Backup User's Profile
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         BiatuAutMiahn[@outlook.com]

 Script Function:
	Backup Various User Files

#ce ----------------------------------------------------------------------------

; Exclude SymLinks/Junctions/Shortcuts
; Exclude AppData
; Exclude Folders that are Empty
; Exclude Folders with only desktop.ini
; Exclude NTUSER.*
;#include <File.au3>
#include <Array.au3>
Local $sName, $aMatch
;$aFiles=_FileListToArrayRec(@UserProfileDir,"*|NTUSER*;desktop.ini;thumbs.db;*.lnk|AppData;Searches;IntelGraphicsProfiles",1,1)
;_ArrayDisplay($aFiles)
If Not $CmdLine[0] Then
    MsgBox(64,"","Usage: Simply drag and drop user's profile folder onto "&@ScriptName&" to backup the profile.")
    Exit 0
EndIf

$aMatch=StringRegExp(StringUpper($CmdLine[1]),"\\((?:[A-Z]\d)?[A-Z]{3,4}[0-9]{3,4}).*",1)
If @error Then
    $sName=StringTrimLeft($CmdLine[1],StringInStr($CmdLine[1],'\',0,-1))
Else
    $sName=$aMatch[0]
EndIf
$sArchive=@DesktopDir&'\Backups\'&@YEAR&'.'&@MON&'.'&@MDAY&'- '&$sName&'.tar'
If FileExists($sArchive) Then
    MsgBox(48,"UserBackup","Warning: Archive "&$sArchive&" already exists! Please rename or remove it then try again.")
    Exit 1
EndIf
If FileExists($CmdLine[1]&'\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks') Then FileCopy($CmdLine[1]&'\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks',$CmdLine[1]&'\_.EdgeBookmarks')
If FileExists($CmdLine[1]&'\AppData\Local\Google\Chrome\User Data\Default\Bookmarks') Then FileCopy($CmdLine[1]&'\AppData\Local\Google\Chrome\User Data\Default\Bookmarks',$CmdLine[1]&'\_.ChromeBookmarks')
RunWait('7zG.exe a -ttar "'&$sArchive&'" "'&$CmdLine[1]&'\*" -xr!AppData -xr!OneDrive -xr!Searches -xr!IntelGraphicsProfiles -xr!NTUSER* -xr!Desktop.ini -xr!thumbs.db -xr!*.lnk -xr!Cookies -xr!"Documents\My\ Music" -xr!"Documents\My Pictures"  -xr!"Documents\My Videos" -xr!"Application Data" -xr!"Local Settings" -xr!"My Documents" -xr!NetHood -xr!PrintHood -xr!Recent -xr!SendTo -xr!"Start Menu" -xr!Templates')

;"7z a -t7z -m0=lzma2 -mx=9 -aoa -mfb=64 -md=32m -ms=on -d=1024m -mhe"
;7za.exe a -ttar -so archive.tar source_files | 7za.exe a -si archive.tar.xz
