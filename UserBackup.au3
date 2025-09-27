#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\Infinity.ico
#AutoIt3Wrapper_Outfile_x64=..\_.rc\UserBackup.exe
#AutoIt3Wrapper_Compression=0
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Description=Creates a snapshot of current user's profile directory.
#AutoIt3Wrapper_Res_Fileversion=1.0.0.92
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1.0.0.0
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:        BiatuAutMiahn[@outlook.com]

 Script Function:
	-Backup all users' essential files.
    -Backup Chrome/Edge/Firefox bookmarks.
    -Backup Chrome/Edge/Firefox passwords.

#ce ----------------------------------------------------------------------------
#include <File.au3>
#include <Array.au3>
#include <Debug.au3>
#include <WinAPIFiles.au3>
#include <String.au3>
;#include "Includes\AD.au3"
;#include "Includes\Base64.au3"


Global Const $VERSION = "1.0.0.92"
Global $sAlias="UserBackup"
Global $sTitle=StringFormat("%s v%s",$sAlias,$VERSION)
ConsoleWrite($sTitle&@CRLF)
Global $sInfPath
$ctBin=EnvGet("SystemDrive")&"\@Corsica\_.UserBackup\"
DirCreate($ctBin)
FileInstall(".\libwim-15.dll",$ctBin,1)
FileInstall(".\wimlib-imagex.exe",$ctBin,1)
If Not FileExists($ctBin&"libwim-15.dll") Or Not FileExists($ctBin&"wimlib-imagex.exe") Then
  $ctBin=@ScriptDir&"\bin"
  If Not FileExists($ctBin&"libwim-15.dll") Or Not FileExists($ctBin&"wimlib-imagex.exe") Then
    ConsoleWrite('Error: Cannot find required files: "'&@CRLF&@ScriptDir&'\bin\libwim-15.dll"'&@CRLF&'"'&@ScriptDir&'\bin\wimlib-imagex.exe"')
    If @Compiled Then
      ConsoleRead()
        ConsoleWrite("Press any key to exit...")
        While Sleep(1)
            If ConsoleRead() <> "" Then ExitLoop
        WEnd
    EndIf
    Exit 1
  EndIf
EndIf

Func _cleanTemp()
    DirRemove($sInfPath,1)
EndFunc

Func CopyChromiumTracks($sSrc, $sDst, $sPath)
	Local $sComSrc=$sSrc&'\'&$sPath
	Local $sComDst=$sDst&'\'&$sPath
	Local $aProfPaths=StringSplit("Cookies,Bookmarks,History,AutoFill,Network\Cookies,Extension Cookies,Login Data,Safe Browsing Network\Safe Browsing Cookies",',')
	FileCopy($sComSrc &"\Last Version",$sComDst&"\Last Version")
	FileCopy($sComSrc &"\Local State",$sComDst&"\Local State")
	$hSearch = FileFindFirstFile($sSrc&'\'&$sPath&"\*")
	While 1
		$sName = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If Not @extended Then ContinueLoop
		For $i=1 to $aProfPaths[0]
			FileCopy($sComSrc&'\'&$sName&'\'&$aProfPaths[$i],$sComDst&'\'&$sName&'\'&$aProfPaths[$i],9)
		Next
	WEnd
EndFunc   ;==>CopyChromiumTracks

Func _delEmpty($dir)
	$folder_list = _FileListToArray($dir,'*',2)
	If @error <> 4 Then
		For $i = 1 To $folder_list[0]
			_delEmpty($dir & '\' & $folder_list[$i])
		Next
	EndIf
	FileFindFirstFile($dir & '\*')
	If @error Then DirRemove($dir)
	FileClose($dir)
EndFunc   ;==>_delEmpty

Func CopyUserDPAPITracks($sSrc,$sDest)
	Local $aAppData[]=[2,"Local","Roaming"]
	Local $aPaths[] = [12, _
	"IsolatedStorage", _
	"DataSharing", _
	"CrashDumps", _
	"Microsoft\NGC", _
	"Microsoft\Vault", _
	"Microsoft\Crypto", _
	"Microsoft\Protect", _
	"Microsoft\Credentials", _
	"Microsoft\SystemCertificates", _
	"Microsoft\Windows\CloudStore", _
	"Microsoft\Windows\CloudAPCache", _
	"ConnectedDevicesPlatform" _
	]
	For $a=1 to $aAppData[0]
		For $b=1 to $aPaths[0]
			DirCopy($sSrc&'\AppData\'&$aAppData[$a]&'\'&$aPaths[$b],$sDest&'\AppData\'&$aAppData[$a]&'\'&$aPaths[$b],1)
		Next
	Next
	DirCopy($sSrc&"\AppData\LocalLow\Microsoft\CryptnetUrlCache",$sDest&"\AppData\LocalLow\Microsoft\CryptnetUrlCache",1)
	FileCopy($sSrc&'\NTUSER.DAT*',$sDest&'\',1)
EndFunc

Func UserBackup($sUserDir, $sSysDrive = -1)
	$sInfPath = $sUserDir & '\_.Infinity.Backup'
	DirCreate($sInfPath)
  OnAutoItExitRegister("_cleanTemp")
	FileSetAttrib($sInfPath, "+ASHN")
	;
	; Need VSS for some of these. (WIP)
	;

	If $sSysDrive <> -1 And FileExists($sSysDrive) Then
		ConsoleWrite("Gathering System Data...")
		DirCopy($sSysDrive & "\Windows\System32\Microsoft\Protect", $sInfPath & "\Host\Windows\System32\Microsoft\Protect", 1)
		DirCopy($sSysDrive & "\Windows\System32\Config", $sInfPath & "\Host\Windows\System32\Config", 1)
    CopyUserDPAPITracks($sSysDrive&"\Windows\System32\Config\SystemProfile",$sInfPath&"\Host\Windows\System32\Config\SystemProfile")
    CopyUserDPAPITracks($sSysDrive&"\Windows\ServiceProfiles\LocalService",$sInfPath&"\Host\Windows\ServiceProfiles\LocalService")
    CopyUserDPAPITracks($sSysDrive&"\Windows\ServiceProfiles\NetworkService",$sInfPath&"\Host\Windows\ServiceProfiles\NetworkService")
		DirCopy($sSysDrive & "\ProgramData\Microsoft\Crypto", $sInfPath & "\Host\ProgramData\Microsoft\Crypto", 1)
		DirCopy($sSysDrive & "\ProgramData\Microsoft\Vault", $sInfPath & "\Host\ProgramData\Microsoft\Vault", 1)
		DirCopy($sSysDrive & "\ProgramData\Microsoft\Wlansvc", $sInfPath & "\Host\ProgramData\Microsoft\Wlansvc", 1)
		DirCopy($sSysDrive & "\ProgramData\Microsoft\Wwansvc", $sInfPath & "\Host\ProgramData\Microsoft\Wwansvc", 1)
		ConsoleWrite("Done" & @CRLF)
	EndIf
	ConsoleWrite("Gathering User Data...")
	DirCopy($sUserDir & "\AppData\Local\Microsoft\Vault", $sInfPath & "\User\AppData\Local\Microsoft\Vault", 1)
	DirCopy($sUserDir & "\AppData\Local\Microsoft\Protect", $sInfPath & "\User\AppData\Local\Microsoft\Protect", 1)
	DirCopy($sUserDir & "\AppData\Local\Microsoft\Credentials", $sInfPath & "\User\AppData\Local\Microsoft\Credentials", 1)
	DirCopy($sUserDir & "\AppData\Roaming\Microsoft\Vault", $sInfPath & "\User\AppData\Roaming\Microsoft\Vault", 1)
	DirCopy($sUserDir & "\AppData\Roaming\Microsoft\Protect", $sInfPath & "\User\AppData\Roaming\Microsoft\Protect", 1)
	DirCopy($sUserDir & "\AppData\Roaming\Microsoft\Credentials", $sInfPath & "\User\AppData\Roaming\Microsoft\Credentials", 1)
	ConsoleWrite("Done" & @CRLF)
	ConsoleWrite("Gathering User's Browser Data...")
    Local $aBrowsers[]=[0, _
        "Microsoft\Edge\User Data", _
        "Microsoft\Edge SxS\User Data", _
        "Google\Chrome\User Data", _
        "Google\Chrome SxS\User Data", _
        "Google\Chromium\User Data", _
        "Google\CocCoc\User Data", _
        "Google\Comodo\Dragon\User Data", _
        "Google\Elements Browser\User Data", _
        "Google\Epic Privacy\Browser\User Data", _
        "7Star\7Star\User Data", _
        "Amigo\User Data", _
        "BraveSoftware\Brave-Browser\User Data", _
        "CentBrowser\User Data", _
        "Chedot\User Data", _
        "Kometa\User Data", _
        "Opera Software\Opera Stable", _
        "Orbitum\User Data", _
        "Sputnik\User Data", _
        "Torch\User Data", _
        "Uran\User Data", _
        "Vivaldi\User Data", _
        "Yandex\YandexBrowser\User Data", _
        "UCBrowser" _
    ]
    $aBrowsers[0]=UBound($aBrowsers,1)-1
    For $i=1 To $aBrowsers[0]
        CopyChromiumTracks($sUserDir, $sInfPath & "\User", "\AppData\Local\"&$aBrowsers[$i])
    Next
	ConsoleWrite("Done"&@CRLF&@CRLF)
	_delEmpty($sInfPath)
;~ firefox_browsers = [
;~     (u'firefox', u'{APPDATA}\\Mozilla\\Firefox'),
;~     (u'blackHawk', u'{APPDATA}\\NETGATE Technologies\\BlackHawk'),
;~     (u'cyberfox', u'{APPDATA}\\8pecxstudios\\Cyberfox'),
;~     (u'comodo IceDragon', u'{APPDATA}\\Comodo\\IceDragon'),
;~     (u'k-Meleon', u'{APPDATA}\\K-Meleon'),
;~     (u'icecat', u'{APPDATA}\\Mozilla\\icecat'),
;~ ]

	;Get User SID
	Local $sProfList="HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	Local $sProfSID=""
	$aProfList=_RegKeys2Arr($sProfList)
	For $i=1 to $aProfList[0]
		If StringRegExp($aProfList[$i],"^S-1-5-\d{2}$") Then ContinueLoop
		$sProfDir=RegRead($sProfList&'\'&$aProfList[$i],"ProfileImagePath")
		If $sProfDir=$sUserDir Then
			$sProfSID=$aProfList[$i]
			ExitLoop
		EndIf
	Next
	If $sProfSID Then
		ConsoleWrite("Profile SID: "&$aProfList[$i]&@CRLF)
		ConsoleWrite("Export Desktop Icon Layout...")
		$iRet=RunWait('reg export HKU\'&$aProfList[$i]&'\Software\Microsoft\Windows\Shell\Bags\1\Desktop "'&$sInfPath&'\User\DesktopIcons.reg" /y',"",@SW_HIDE)
		$sRegDat=FileRead($sInfPath&"\User\DesktopIcons.reg")
		$sRegDat=StringReplace($sRegDat,"HKEY_USERS\"&$aProfList[$i],"HKEY_CURRENT_USER")
		$hFile=FileOpen($sInfPath&"\User\DesktopIcons.reg",2)
		FileWrite($hFile,$sRegDat)
		FileClose($hFile)
		;ConsoleWrite(($iRet=0?"Done":StringFormat("Failed (Error: %d)",$iRet))&@CRLF)
  Else
    ConsoleWrite('Profile SID "'&$aProfList[$i]&'" Not Found!'&@CRLF)
	EndIf

	;Backup user
	ConsoleWrite("!! This may take several minutes !!"&@CRLF)
	ConsoleWrite("Gathering User files list...")
  $hTimer=TimerInit()
	$aSearch = _FileListToArrayRec($sUserDir & '\', '*', 2, 1, 0, 1)
	ConsoleWrite(StringFormat("Done (%dms)\r\n",TimerDiff($hTimer)))

	ConsoleWrite("Generating Exclusion List...")
  $hTimer=TimerInit()
	$sConfigIni = "[ExclusionList]" & @CRLF
	; The following folders will be excluded
	Local $aDirFilter[] = [0, _
        "SendTo", _
        "Extensions", _
        "Recent", _
        "Edge Designer", _
        "Edge Kids Mode", _
        "Edge Shopping", _
        "Edge Tipping", _
        "Edge Travel", _
        "Edge Wallet", _
        "Windows\\WinX", _
        "Libraries", _
        "D3DSCache", _
        "Feeds Cache", _
        "CacheStorage", _
        "Feeds", _
        "ShaderCache", _
        "User Data\\Snapshots", _
        "Service Worker\\CacheStorage", _
        "Service Worker\\ScriptCache", _
        "Microsoft\\Templates", _
        "fontconfig\\cache", _
        "Code Cache", _
        "Cache\\Cache_Data", _
        "LocalLow\\Adobe", _
        "Links", _
        "Microsoft\\Office", _
        "PowerToys", _
        "Teams", _
        "RoamCache", _
        "pip\\cache", _
        "Package Cache", _
        "Microsoft\\Packages", _
        "Terminal Server Client", _
        "TokenBroker\\Cache", _
        "ActionCenterCache", _
        "Local\\Programs", _
        "Local\\Temp", _
        "Roaming\\Code", _
        "Microsoft\Installer", _
        "AppCenterCache", _
        "Burn", _
        "Windows\\Caches", _
        "TeamsMeetingAddin", _
        "TeamsPresenceAddin", _
        "thumbcache", _
        "WebCache", _
        "GPUCache", _
        "pyppeteer", _
        "iconcache", _
        "IntelGraphicsProfiles" _
    ]
	$aDirFilter[0] = UBound($aDirFilter, 1) - 1
	Local $sExclDirList = ''
	For $i = 1 To $aDirFilter[0]
		$sExclDirList &= $aDirFilter[$i]
		If $i <> $aDirFilter[0] Then $sExclDirList &= '|'
	Next
	$sFilter = "(?i)("&$sExclDirList&")"
  Local $iAttr
  Local $aDirExcl[]=[0]

	For $i = 1 To $aSearch[0]
    If Mod($i,100)=0 Then
      $sProgStr=StringFormat(" (%0.2f%%)",($i/$aSearch[0])*100)
      ConsoleWrite(StringFormat("\r\rGenerating Exclusion List...%s",$sProgStr))
    EndIf
    $vMatch=StringRegExp($aSearch[$i],$sFilter)
		If Not $vMatch Then
      $aDirFiles=_FileListToArrayRec($sUserDir&'\'&$aSearch[$i],'*',1,0,0,1)
      ;_DebugArrayDisplay($aDirFiles,$i&','&$sUserDir&'\'&$aSearch[$i])
      If @error Then ContinueLoop
      ;$aDirFiles=_FileListToArray($sUserDir&'\'&$aSearch[$i],'*',1,0)
      For $k=1 To $aDirFiles[0]
        $iAttr=_WinAPI_GetFileAttributes($sUserDir&'\'&$aSearch[$i]&'\'&$aDirFiles[$k])
        If BitAND($iAttr,0x00040000)=0x00040000 Or BitAND($iAttr,0x00400000)=0x00400000 Then
          ;ConsoleWrite('\'&$aSearch[$i]&'\'&$aDirFiles[$k]&@CRLF)
          $sConfigIni&='\'&$aSearch[$i]&@CRLF
        EndIf
      Next
      ContinueLoop
    EndIf
    For $k=1 To $aDirExcl[0]
      If $aSearch[$i]=$aDirExcl[$k] Then ContinueLoop 2
      If StringLeft($aSearch[$i],StringLen($aDirExcl[$k]))=$aDirExcl[$k] Then ContinueLoop 2
    Next
    $iMaxDE=UBound($aDirExcl,1)
    ReDim $aDirExcl[$iMaxDE+1]
    $aDirExcl[$iMaxDE]=$aSearch[$i]
    $aDirExcl[0]=$iMaxDE
		$sConfigIni&='\'&StringTrimRight($aSearch[$i],1)&@CRLF
  Next
	$sConfigIni &= @CRLF

	; The following files will be excluded
	Local $aFileFilter[] = [0, _
			"IconCache.db", _
			"bing.url", _
			"desktop.ini", _
			".igpi", _
			".regtrans-ms", _
			".search-ms", _
			".searchconnector-ms", _
			".LOG1", _
			".LOG2", _
			".blf" _
			]
	$aFileFilter[0] = UBound($aFileFilter, 1) - 1
	For $i = 1 To $aFileFilter[0]
		$sConfigIni &= $aFileFilter[$i]&@CRLF
		;If $i <> $aFileFilter[0] Then $sConfigIni &= @CRLF
	Next
	$sConfigIni &= @CRLF
	$sConfigIni &= @CRLF
	$sUserName = StringTrimLeft($sUserDir, StringInStr($sUserDir, '\', 0, -1))
	;$sDest = @WorkingDir & '\' & @YEAR & '.' & @MON & '.' & @MDAY & ',' & @HOUR & @MIN & ' ' & $sUserName & '.esd'
  $sDest=@WorkingDir&'\'&$sUserName&'.esd'
	$sConfigIni &= $sDest & @CRLF

	; The following extensions will not be compressed.
	$sExList = "7z,wim,esd,swm,apk,xapk,zst,rar,gzip,gz,pkg,xz,bz2,bzip,sit,dmg,vhdx,sqx,pit,xip,mpkg,dar,b1,uha,r00,zip,rev,wa,pak,rpm,deb,"
	$sExList &= "daa,cdz,lz4,lz,whl,warc,lzma,ipk,sfs,squashfs,ace,sitx,arc,jar,pea,war,zpaq,paq,zipx,tgz,bzip2,pup,arj,xar,paq8p,paq8f,"
	$sExList &= "jgz,b64,zim,bundle,cpgz,cpt,nsis,odex,sea,xpi,crx,docx,docm,dotx,dotm,wll,wwl,xlsx,xlsm,xltx,xltm,pptx,potx,potm,ppam,ppsx,ppsm,sidx,"
	$sExList &= "sldm,pa,pub,3g2,3gp,3gp2,3gpp,amv,asf,avi,bik,bin,crf,dav,divx,drc,dv,dvr-ms,evo,f4v,flv,gvi,gxf,m1v,m2v,m2t,m2ts,m4v,mkv,mov,mp2,"
	$sExList &= "mp2v,mp4,mp4v,mpe,mpeg,mpeg1,mpeg2,mpeg4,mpg,mpv2,mts,mtv,mxf,mxg,nsv,nuv,ogg,ogm,ogv,ogx,ps,rec,rm,rmvb,rpl,smk,thp,tod,tp,ts,tts,txd,"
	$sExList &= "vob,vro,webm,wm,wmv,wtv,xesc,3ga,a52,aac,ac3,adt,adts,aif,aifc,aiff,amb,amr,aob,ape,au,awb,caf,dts,flac,it,kar,m4a,m4b,m4p,m5p,mid,"
	$sExList &= "mka,mlp,mod,mpa,mp1,mp3,mpc,mpga,mus,oga,oma,opus,qcp,ra,rmi,s3m,sid,spx,tak,thd,tta,voc,vqf,w64,wav,wma,wv,xa,xm,appx,msix,appxupload,"
	$sExList &= "appxbundle,cab,txz,tbz2,tbz,z,taz,lzh,lhz,png,jpg,gze,lha,odg,odp,ods,odt,pdf,ott,jpeg"
	$aExcludeExts = StringSplit($sExList, ',')
	$sConfigIni &= "[CompressionExclusionList]" & @CRLF
	For $i = 1 To $aExcludeExts[0]
		$sConfigIni &= "*." & $aExcludeExts[$i]
		If $i <> $aExcludeExts[0] Then $sConfigIni &= @CRLF
	Next
	$sConfigIni &= @CRLF
	$hFile = FileOpen(@WorkingDir & "\Temp\config.ini", 2 + 8 + 128)
	FileWrite($hFile, $sConfigIni)
	FileClose($hFile)
	ConsoleWrite(StringFormat("\r\rGenerating Exclusion List...Done (%dms)\r\n",TimerDiff($hTimer)))
	ConsoleWrite("Archiving User Files..."&@CRLF)
	; 7z cannot create volume snapshots
	;RunWait('bin\7za.exe a -mx9 -myx9 -mmemusep80 -mtm -mtc -mtr -ms16g -stl -slp -i@"C:\temp\fnames.txt" "'&@YEAR&'.'&@MON&'.'&@MDAY&','&@HOUR&@MIN&' '&@UserName&'.7z"',$sUserDir)
	;ConsoleWrite('bin\wimlib-imagex.exe capture "' & $sUserDir & '" "' & $sDest & '" "' & @YEAR & '.' & @MON & '.' & @MDAY & ',' & @HOUR & @MIN & ' ' & $sUserName & '" ""  --solid-chunk-size="64M" --include-integrity  --snapshot  --solid  --solid-compress=LZMS:100  --no-acls --config ' & @WorkingDir & '\temp\config.ini' & @CRLF)
  $sOp="append"
	Local $sCmd = $ctBin&'\wimlib-imagex.exe '&$sOp&' "' & $sUserDir & '" "' & $sDest & '" "' & @YEAR & '.' & @MON & '.' & @MDAY & ',' & @HOUR & @MIN & ' ' & $sUserName & '" "" --create --include-integrity  --snapshot  --no-acls --compress=LZX:100 --config "' & @WorkingDir & '\temp\config.ini"'
	ConsoleWrite($sCmd & @CRLF)
	$iPid = Run($sCmd,@ScriptDir,@SW_SHOW,0x10)
    ProcessWait($iPid)
    ProcessSetPriority($iPid,0)
    ProcessWaitClose($iPid)
    $iRet=@Extended
	If $iRet <> 0 Then
		ConsoleWrite("Failed (" & $iRet & ')' & @CRLF)
	EndIf
	ConsoleWrite("Done" & @CRLF)
    _cleanTemp()
EndFunc   ;==>UserBackup

$sDrive = -1
$sUserPath = ""
If $CmdLine[0] = 0 Then
	$sDrive = EnvGet("SystemDrive")
	$sUserPath = @UserProfileDir
Else
	If Not FileExists($CmdLine[1]) Or Not StringInStr(FileGetAttrib($CmdLine[1]), 'D') Then
		MsgBox(48, '', 'Cannot Open "' & $CmdLine[1] & '"')
		Exit 2
	EndIf
	$sUserPath = $CmdLine[1]
	If $CmdLine[0] = 2 Then
		If FileExists($CmdLine[2]) And StringInStr(FileGetAttrib($CmdLine[2]), 'D') Then $sDrive = $CmdLine[2]
	EndIf
EndIf
;~ If Not FileExists($CmdLine[2]) Or Not StringInStr(FileGetAttrib($CmdLine[2]),'D') Then
;~     MsgBox(48,"","Nothing to do!")
;~     Exit 3
;~ EndIf
UserBackup($sUserPath, $sDrive)
If @Compiled Then
	ConsoleRead()
    ConsoleWrite("Press any key to exit...")
    While Sleep(1)
        If ConsoleRead() <> "" Then ExitLoop
    WEnd
EndIf

Func _RegKeys2Arr($sKey)
  Global $aKeys[1]
  $aKeys[0]=0
  $i=1
  Do
    $sVar=RegEnumKey($sKey,$i)
    If @error Then ExitLoop
    $iMax=UBound($aKeys,1)
    ReDim $aKeys[$iMax+1]
    $aKeys[$iMax]=$sVar
    $aKeys[0]=$iMax
    $i+=1
  Until False
  Return $aKeys
EndFunc
