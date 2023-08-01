#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\ohiohealth.ico
#AutoIt3Wrapper_Outfile_x64=..\_.rc\AddPrinterSimple.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Add a local printer
#AutoIt3Wrapper_Res_Fileversion=23.321.1638.15
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Fileversion_First_Increment=y
#AutoIt3Wrapper_Res_Fileversion_Use_Template=%YY.%MO%DD.%HH%MI.%SE
#AutoIt3Wrapper_Res_ProductName=AddPrinterSimple
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Au3Stripper=n
#AutoIt3Wrapper_Res_HiDpi=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         BiatuAutMiahn[@outlook.com]

 Script Function:
	Add Printer

#ce ----------------------------------------------------------------------------
Global $sAlias="AddPrinterSimple"
Global $sPr="", $iRet=0
Do
    $sPr=InputBox($sAlias,"Enter printer to add"&@LF&"eg; \\<redacted>\Secure Print")
Until $sPr<>"" Or @Error
If $sPr="" Then Exit 2
$iRet=ShellExecuteWait("rundll32.exe",'printui.dll,PrintUIEntry /in /n "'&$sPr&'"',@SystemDir)
If $iRet<>0 Then
    MsgBox(16,$sAlias,"Error: There was an error adding this printer. Error Code: "&$iRet)
    Exit 1
EndIf
$iRet=MsgBox(36,$sAlias,"Would you like to set this printer as default?")
If $iRet==6 Then
    $iRet=ShellExecuteWait("rundll32.exe",'printui.dll,PrintUIEntry /y /n "'&$sPr&'"',@SystemDir)
    If $iRet<>0 Then
        MsgBox(16,$sAlias,"Error: There was an error setting this printer as default. Error Code: "&$iRet)
        Exit 1
    EndIf
EndIf
Exit 0

