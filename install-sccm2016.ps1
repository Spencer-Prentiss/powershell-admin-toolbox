#region **** Script Information ****

<#

File name:		install-sccm2016.ps1
Description:	SCCM 2016 Client Installation Handler using imported common.psm1.

Author:			Spencer Prentiss

Versions:
	1.0			Base script with core functionality

#>

#endregion


#region **** Global "Const" Variables

[String]$Global:ScriptVersion = '1.0'
[String]$Global:PackageName = 'SCCM 2016'
[String]$Global:PackageFolder = 'SCCMClient_5.00.8577.1000' # Update before using
[String]$Global:HTTP_PORT = '777'							# Update before using
[String]$Global:HTTPS_PORT = '778'							# Update before using
[String]$Global:LOG_MAX_SIZE = '512000'						# Update before using
[String]$Global:SMS_SLP_SERVER = 'SMS.SLP.Server.com'		# Update before using
[String]$Global:SCCMArgs = "/noservice /UsePKICert SMSDIRECTORYLOOKUP=NOWINS CCMHTTPPORT=${Global:HTTP_PORT} CCMHTTPSPORT=${Global:HTTPS_PORT} SMSSITECODE=AUTO CCMLOGMAXHISTORY=1 CCMLOGMAXSIZE=${Global:LOG_MAX_SIZE} SMSSLP=${Global:SMS_SLP_SERVER}"

#endregion **** Global "Const" Variables


#region **** Setup Script ****

[String]$module = ((Split-Path -Parent $MyInvocation.MyCommand.Definition) + '\common.psm1')
If (Test-Path -Path $module)
{
	# [0]	Args
	# [1]	OSTypes
	# [2]	OSVersions
	Import-Module $module -ArgumentList (
		$Args,
		1,
		('6.1', '6.2', '6.3', '10.0')
	)
}
Else { Exit 2 }

#endregion **** Setup Script ****


#region **** Script Main ****

Function ScriptMain
{
	OutputLog -Text ('[FUNCTION]>  ScriptMain  [STATE]>  Start')
	OutputLog -Text ('[HEADER]>  Script Header  [STATE]>  Start')
	OutputLog -Text ('Script:   ' + $ScriptName)
	OutputLog -Text ('Version:  ' + $ScriptVersion)
	OutputLog -Text ('[HEADER]>  Script Header  [STATE]>  End')
	
	OutputLog -Text ('Disabling Program Compatibility Assistant service')
	[Void](RunProcess -FilePath 'sc.exe' -ArgList 'config pcasvc start= disabled')
	
	OutputLog -Text ('Stopping Program Compatibility Assistant service')
	[Void](RunProcess -FilePath 'sc.exe' -ArgList 'stop pcasvc')
	
	[String]$packagePath = ($Env:SystemRoot + '\Temp\' + $PackageFolder)
	[String]$packageArgs = ('/source:"' + $packagePath + '" ' + $Global:SCCMArgs)
	
	If (CachePackage -Source ($ScriptRoot + '\' + $PackageFolder) -Destination $packagePath -Mirror $True -Overwrite $True)
	{
		OutputLog -Text ('Installing ' + $PackageName)
		[Void](RunProcess -FilePath ($packagePath + '\ccmsetup.exe') -ArgList $packageArgs)
		OutputLog -Text ($PackageName + ' installed')
	}
	
	OutputLog -Text ('[FUNCTION]>  ScriptMain  [STATE]>  End')
	OutputLog -Text ('********************************************************************************')
}

#endregion **** Script Main ****


#region **** Script Main Call and Exit ****

ScriptMain
Exit 0

#endregion **** Script Main Call and Exit ****