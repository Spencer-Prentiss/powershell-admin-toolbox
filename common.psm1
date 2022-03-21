#region **** Script Information ****

<#	File Information

File name:		common.psm1
Description:	PowerShell Module with commonly used functions

Author:			Spencer Prentiss

Versions:
	1.0			Base module with core functionality
	1.1			Added NTDomain variable that is explicitly set based on NetBIOS domain name
	1.2			Restructured module to be more abstract so it can be used for script as well as packages
	
Import-Module ArgumentList
	Argument 1:  LaunchArgs					Type:  String Array			Default Value:  $Null			Mandatory:  False

Import-Module Examples:
	Example 1, Default Arguments:
		[String]$module = ((Split-Path -Parent $MyInvocation.MyCommand.Definition) + '\common.psm1')
		If (Test-Path -Path $module) { Import-Module $module }
		Else { Exit 2 }
	
	Example 2, Explicit Arguments:
		[String]$module = ((Split-Path -Parent $MyInvocation.MyCommand.Definition) + '\common.psm1')
		If (Test-Path -Path $module)
		{
			Import-Module $module -ArgumentList (
				$Args,
				(2, 3),
				(6.1, 6.2, 6.3, 10.0)
			)
		}
		Else { Exit 2 }

#>

#endregion


#region **** Import-Module ArgumentList ****

Param
(
	[Parameter(Position=0)][String[]]$Global:LaunchArgs = $Null,
	[Parameter(Position=1)][UInt32[]]$Global:SupportedOSTypes = 1,
	[Parameter(Position=2)][System.Version[]]$Global:SupportedOSVersions = ('6.1', '6.2', '6.3', '10.0')
)

#endregion **** Import-Module ArgumentList ****


#region **** Global "Const" Variables ****

[String]$Global:ModuleVersion = '1.2'
[String]$Global:DefaultString = 'NULL'
[String]$Global:NA = 'N/A'
[String]$Global:LogPath = $env:TEMP
[String]$Global:LogName = $DefaultString
[String]$Global:LogFile = $DefaultString
[UInt32]$Global:LogIndentSize = 4
[String]$Global:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
[String]$Global:ScriptFolder = Split-Path -Leaf $ScriptRoot
If (($Null -NE $MyInvocation.MyCommand) -And ($MyInvocation.MyCommand -NE [String]::Empty))
{
	[String]$Global:ModuleName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
	If (($Null -NE $MyInvocation.PSCommandPath) -And ($MyInvocation.PSCommandPath -NE [String]::Empty))
	{
		[String]$Global:ScriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.PSCommandPath)
	}
	Else { $Global:ScriptName = $ModuleName }
	
	$Global:LogName = $ScriptName
	$Global:LogFile = ($LogName + '.log')
}

#endregion **** Global "Const" Variables ****


#region **** Global Variables ****

[String]$Global:FullLogPath = $DefaultString
[Int32]$Global:LogIndentTier = 0
[System.Collections.Generic.List[String]]$Global:PreLogText = New-Object System.Collections.Generic.List[String]
[System.Collections.Generic.List[String]]$Global:PreWriteHostText = New-Object System.Collections.Generic.List[String]
[System.Collections.Generic.List[String]]$Global:PreWriteHostTextColor = New-Object System.Collections.Generic.List[String]
[System.Collections.Generic.List[String]]$Global:ExitWithErrorText = New-Object System.Collections.Generic.List[String]
[System.Collections.Generic.List[String]]$Global:ExitWithErrorCodes = New-Object System.Collections.Generic.List[String]
[String]$Global:Manufacturer = $DefaultString
[String]$Global:Model = $DefaultString
[String]$Global:ChassisType = 'Desktop'
[String]$Global:BIOSVersion = $DefaultString
[String]$Global:ServiceTag = $DefaultString
[String]$Global:AssetTag = $DefaultString
[Double]$Global:Processor = 0
[Double]$Global:Memory = 0
[Bool]$Global:OnACPower = $False
[String]$Global:OSName = $DefaultString
[UInt32]$Global:OSType = 0
[System.Version]$Global:OSVersion = $Null
[String]$Global:OSArch = $DefaultString
[String]$Global:UserName = $DefaultString
[String]$Global:Domain = $DefaultString
[String]$Global:NTDomain = $DefaultString
[String]$Global:ComputerName = $DefaultString
[String]$Global:TimeZone = $DefaultString
[Bool]$Global:LogToFile = $False
[Bool]$Global:LogToHost = $False
[String]$Global:ExecutionType = 'Silent'
[Bool]$Global:RebootSelected = $False
[Bool]$Global:ErrorState = $False
[Int32]$Global:ExitCode = 0

#endregion **** Global Variables ****


#region **** Additional Includes ****

Add-Type -AssemblyName System.Web.Extensions
Add-Type -AssemblyName System.Windows.Forms

#endregion **** Additional Includes ****


#region **** Module Main Function

Function ModuleMain
{
	Clear-Host
	
	OutputLog -Text ('********************************************************************************')
	OutputLog -Text ('[FUNCTION]>  ModuleMain  [STATE]>  Start')
	OutputLog -Text ('[HEADER]>  Module Header  [STATE]>  Start')
	OutputLog -Text ('Module:   ' + $ModuleName)
	OutputLog -Text ('Version:  ' + $ModuleVersion)
	OutputLog -Text ('[HEADER]>  Module Header  [STATE]>  End')
	
	CheckLaunchArgs
	GetSystemInfo
	InitSupportStructure
	ReportEnvironment
	
	OutputLog -Text ('[FUNCTION]>  ModuleMain  [STATE]>  End')
	If ($ErrorState)
	{
		OutputLog -Text ('********************************************************************************')
		ExitWithError -ShouldExitHere $True
	}
}

#endregion **** Module Main Function


#region **** Module Setup Functions

Function CheckLaunchArgs
{
	OutputLog -Text ('[FUNCTION]>  CheckLaunchArgs  [STATE]>  Start')
	
	OutputLog -Text ('Checking for launch arguments')
	If ($LaunchArgs.Length -GT 0)
	{
		If ($LaunchArgs.Length -EQ 1)
		{
			OutputLog -Text ($LaunchArgs.Length.ToString() + ' launch argument found')
			OutputLog -Text ('->  Argument:  ' + $LaunchArgs[0])
			OutputLog -Text ('Checking format of argument')
		}
		Else
		{
			OutputLog -Text ($LaunchArgs.Length.ToString() + ' launch arguments found')
			For ([UInt32]$i = 0; $i -NE $LaunchArgs.Length; ++$i)
			{
				OutputLog('->  Argument ' + ($i + 1) + ':  ' + $LaunchArgs[$i])
			}
			OutputLog -Text ('Checking format of arguments')
		}

		For ([UInt32]$i = 0; $i -NE $LaunchArgs.Length; ++$i)
		{
			OutputLog -Text ('Splitting argument "' + $LaunchArgs[$i] + '" with delimiter ":"')
			[String[]]$argSplit = $LaunchArgs[$i] -Split ':'
			If ($argSplit.Length -EQ 2)
			{
				[String]$subArg = $argSplit[0]
				While (($subArg.Length -GT 0) -And (!($subArg[0] -Match '[A-Z]'))) { $subArg = $subArg.Substring(1) }
				
				[Bool]$valueIsValid = $False
				OutputLog -Text ('Checking if "' + $subArg + '" is a valid settable variable')
				Switch ($subArg.ToUpper())
				{
					'EXECUTION'
					{
						OutputLog -Text ('Assuming "$Global:ExecutionType" is intended variable')
						$subArg = 'ExecutionType'
					}
					'LOG'
					{
						OutputLog -Text ('Assuming "$Global:LogToHost" is intended variable')
						$subArg = 'LogToHost'
					}
					'USER'
					{
						OutputLog -Text ('Assuming "$Global:UserName" is intended variable')
						$subArg = 'UserName'
					}
				}
				
				Switch ($subArg.ToUpper())
				{
					'EXECUTIONTYPE'
					{
						OutputLog -Text ('"$Global:ExecutionType" is a settable variable, checking if "' + $argSplit[1] + '" is a valid value')
						Switch ($argSplit[1].ToUpper())
						{
							'P'				{ $valueIsValid = $True; $Global:ExecutionType = 'Prompt' }
							'PROMPT'		{ $valueIsValid = $True; $Global:ExecutionType = 'Prompt' }
							'N'				{ $valueIsValid = $True; $Global:ExecutionType = 'Notify' }
							'NOTIFY'		{ $valueIsValid = $True; $Global:ExecutionType = 'Notify' }
							'NOTIFICATION'	{ $valueIsValid = $True; $Global:ExecutionType = 'Notify' }
							'S'				{ $valueIsValid = $True; $Global:ExecutionType = 'Silent' }
							'SILENT'		{ $valueIsValid = $True; $Global:ExecutionType = 'Silent' }
							'Q'				{ $valueIsValid = $True; $Global:ExecutionType = 'Silent' }
							'QUIET'			{ $valueIsValid = $True; $Global:ExecutionType = 'Silent' }
						}
						
						If ($valueIsValid) { OutputLog -Text ('Value "' + $argSplit[1] + '" is valid, value of "' + $ExecutionType + '" will be used') }
						Else { OutputLog -Text ('Value "' + $argSplit[1] + '" is not valid for "$Global:ExecutionType", default value of "' + $ExecutionType + '" will be used instead') -Severity 'Warning' }
					}
					'LOGTOHOST'
					{
						OutputLog -Text ('"$Global:LogToHost" is a settable variable, checking if "' + $argSplit[1] + '" is a valid value')
						Switch ($argSplit[1].ToUpper())
						{
							'T'		{ $valueIsValid = $True; $Global:LogToHost = $True }
							'TRUE'	{ $valueIsValid = $True; $Global:LogToHost = $True }
							'Y'		{ $valueIsValid = $True; $Global:LogToHost = $True }
							'YES'	{ $valueIsValid = $True; $Global:LogToHost = $True }
							'F'		{ $valueIsValid = $True; $Global:LogToHost = $False }
							'FALSE'	{ $valueIsValid = $True; $Global:LogToHost = $False }
							'N'		{ $valueIsValid = $True; $Global:LogToHost = $False }
							'NO'	{ $valueIsValid = $True; $Global:LogToHost = $False }
						}
						
						If ($valueIsValid) { OutputLog -Text ('Value "' + $argSplit[1] + '" is valid, value of "' + $LogToHost + '" will be used') }
						Else { OutputLog -Text ('Value "' + $argSplit[1] + '" is not valid for "$Global:LogToHost", default value of "' + $LogToHost + '" will be used instead') -Severity 'Warning' }
					}
					'USERNAME'
					{
						OutputLog -Text ('"$Global:UserName" is a settable variable, ' + $argSplit[1] + ' will be used')
						$Global:UserName = $argSplit[1]
					}
					'DOMAIN'
					{
						OutputLog -Text ('"$Global:Domain" is a settable variable, ' + $argSplit[1] + ' will be used')
						$Global:Domain = $argSplit[1]
					}
					'NTDOMAIN'
					{
						OutputLog -Text ('"$Global:NTDomain" is a settable variable, ' + $argSplit[1] + ' will be used')
						$Global:NTDomain = $argSplit[1]
					}
					Default { OutputLog -Text ('Variable not settable') -Severity 'Warning' }
				}
			}
			Else { OutputLog -Text ('"' + $LaunchArgs[$i] + '" does not match "Variable:Value" format, unable to determine variable') -Severity 'Warning' }
		}
	}
	Else { OutputLog -Text ('No launch arguments found') }
	
	OutputLog -Text ('[FUNCTION]>  CheckLaunchArgs  [STATE]>  End')
}

Function GetSystemInfo
{
	OutputLog -Text ('[FUNCTION]>  GetSystemInfo  [STATE]>  Start')
	
	[Object]$objWin32OS = $Null
	[Object]$objWin32Processor = $Null
	[Object]$objWin32Comp = $Null
	[Object]$objWin32Service = $Null
	[Object]$objWin32SysEnclosure = $Null
	[Object]$objWin32Battery = $Null
	[Object]$objWin32BIOS = $Null
	[Object]$objWin32TimeZone = $Null
	Try
	{
		$objWin32OS = Get-WMIObject -Class 'Win32_OperatingSystem' -ErrorAction SilentlyContinue
		$objWin32Processor = Get-WMIObject -Class 'Win32_Processor' -ErrorAction SilentlyContinue
		$objWin32Comp = Get-WMIObject -Class 'Win32_ComputerSystem' -ErrorAction SilentlyContinue
		$objWin32SysEnclosure = Get-WMIObject -Class 'Win32_SystemEnclosure' -ErrorAction SilentlyContinue
		$objWin32Battery = Get-WMIObject -Class 'Win32_Battery' -ErrorAction SilentlyContinue
		$objWin32BIOS = Get-WMIObject -Class 'Win32_BIOS' -ErrorAction SilentlyContinue
		$objWin32TimeZone = Get-WMIObject -Class 'Win32_TimeZone' -ErrorAction SilentlyContinue
	}
	Catch { AddErrorToList -Text ($_.Exception.Message) }
	
	If (ObjectIsValid -Object $objWin32OS)
	{
		$Global:OSName = $objWin32OS.Caption
		$Global:OSVersion = $objWin32OS.Version
		$Global:OSType = $objWin32OS.ProductType
	
		[Bool]$match = $False
		If ($OSVersion -NE $Null)
		{
			OutputLog -Text ('Checking if operating system version is supported by module')
			OutputLog -Text ('Supported operating system version(s):')
			For ([UInt32]$i = 0; $i -NE $SupportedOSVersions.Length; ++$i)
			{
				OutputLog -Text ('->  ' + $SupportedOSVersions[$i].ToString())
			}
			
			For ([UInt32]$i = 0; $i -NE $SupportedOSVersions.Length; ++$i)
			{
				If (($OSVersion.Major -EQ $SupportedOSVersions[$i].Major) -And ($OSVersion.Minor -EQ $SupportedOSVersions[$i].Minor))
				{
					$match = $True
					Break
				}
			}
			
			If ($match) { OutputLog -Text ('Supported operating system version found') }
			Else { AddErrorToList -Text ('Supported operating system version not found') }
		}
		Else { AddErrorToList -Text ('Operating system version not found') }
		
		$match = $False
		If ($OSType -NE 0)
		{
			OutputLog -Text ('Checking if operating system type is supported by module')
			OutputLog -Text ('Supported operating system type(s):')
			For ([UInt32]$i = 0; $i -NE $SupportedOSTypes.Length; ++$i)
			{
				OutputLog -Text ('->  ' + ($SupportedOSTypes[$i]).ToString())
			}
			
			For ([UInt32]$i = 0; $i -NE $SupportedOSTypes.Length; ++$i)
			{
				If ($OSType -EQ $SupportedOSTypes[$i])
				{
					$match = $True
					Break;
				}
			}
			
			If ($match) { OutputLog -Text ('Supported operating system type found') }
			Else { AddErrorToList -Text ('Supported operating system type not found') }
		}
		Else { AddErrorToList -Text ('Operating system type not found') }
	}
	Else { AddErrorToList -Text ('Win32_OperatingSystem WMI Class object not valid') }
	
	If (ObjectIsValid -Object $objWin32Processor)
	{
		$Global:Processor = [Math]::Round($objWin32Processor.MaxClockSpeed / 1000, 3)
		If ($objWin32Processor.OSArchitecture -EQ '32-bit') { $Global:OSArch = 'x86' }
		ElseIf ($objWin32Processor.OSArchitecture -EQ '64-bit') { $Global:OSArch = 'x64' }
		Else
		{
			If ($objWin32Processor.AddressWidth -EQ '32') { $Global:OSArch = 'x86' }
			ElseIf ($objWin32Processor.AddressWidth -EQ '64') { $Global:OSArch = 'x64' }
			Else
			{
				If (ObjectIsValid -Object ${Env:ProgramFiles(x86)}) { $Global:OSArch = 'x64' }
				Else { $Global:OSArch = 'x86' }
			}
		}
	}
	Else { AddErrorToList -Text ('Win32_Processor WMI Class object not valid') }
	
	Try
	{
		[String]$userLong = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
		[String[]]$userSplit = $userLong -Split '\\'
		If ((ObjectIsValid -Object $userSplit) -And ($userSplit.Length -GT 1))
		{
			If ($UserName -EQ $DefaultString) { $Global:UserName = $userSplit[1] }
		}
	}
	Catch { AddErrorToList -Text ($_.Exception.Message) }
	
	If (ObjectIsValid -Object $objWin32Comp)
	{
		$Global:ComputerName = $objWin32Comp.Name
		$Global:Manufacturer = $objWin32Comp.Manufacturer
		$Global:Model = $objWin32Comp.Model
		$Global:Memory = [Math]::Round($objWin32Comp.TotalPhysicalMemory / 1024 / 1024 / 1024, 3)
		If ($Domain -EQ $DefaultString) { $Global:Domain = $objWin32Comp.Domain }
		If ($NTDomain -EQ $DefaultString) { $Global:NTDomain = $Domain }
	}
	Else { AddErrorToList -Text ('Win32_ComputerSystem WMI Class object not valid') }
	
	If (ObjectIsValid -Object $objWin32SysEnclosure)
	{
		$Global:AssetTag = $objWin32SysEnclosure.SMBIOSAssetTag
		Switch ($objWin32SysEnclosure.ChassisTypes)
		{
			'8'  { $Global:ChassisType = 'Laptop' }
			'9'  { $Global:ChassisType = 'Laptop' }
			'10' { $Global:ChassisType = 'Laptop' }
			'11' { $Global:ChassisType = 'Laptop' }
			'12' { $Global:ChassisType = 'Laptop' }
			'14' { $Global:ChassisType = 'Laptop' }
			'18' { $Global:ChassisType = 'Laptop' }
			'21' { $Global:ChassisType = 'Laptop' }
		}
	}
	Else { OutputLog -Text ('Win32_SystemEnclosure WMI Class object not valid') -Severity 'Warning' }
	
	If ($ChassisType -EQ 'Laptop')
	{
		If (ObjectIsValid -Object $objWin32Battery)
		{
			If ($objWin32Battery.BatteryStatus -EQ 2) { $Global:OnACPower = $True }
		}
		Else { OutputLog -Text ('Win32_Battery WMI Class object not valid') -Severity 'Warning' }
	}
	
	If (ObjectIsValid -Object $objWin32BIOS)
	{
		$Global:ServiceTag = $objWin32BIOS.SerialNumber
		$Global:BIOSVersion = $objWin32BIOS.SMBIOSBIOSVersion
	}
	Else { OutputLog -Text ('Win32_BIOS WMI Class object not valid') -Severity 'Warning' }
	
	If (ObjectIsValid -Object $objWin32TimeZone)
	{
		$Global:TimeZone = $objWin32TimeZone.Caption
	}
	Else { OutputLog -Text ('Win32_TimeZone WMI Class object not valid') -Severity 'Warning' }
	
	OutputLog -Text ('[FUNCTION]>  GetSystemInfo  [STATE]>  End')
}

Function InitSupportStructure
{
	OutputLog -Text ('[FUNCTION]>  InitSupportStructure  [STATE]>  Start')
	
	If (ObjectExists -Path $LogPath -Item $LogFile -ItemType 'File' -CreateIfNotExist $True -LogOnlyErrors $False)
	{
		$Global:FullLogPath = ($LogPath + '\' + $LogFile)
		
		OutputLog -Text ('Checking if "' + $LogFile + '" is greater than 5 MBs and moving if so')
		If ((Get-Item -Path $FullLogPath -ErrorAction SilentlyContinue).Length -GT 5MB)
		{
			OutputLog -Text ('"' + $LogFile + '" is greater than 5 MBs, moving to "' + $LogName + '.Backup.Log"')
			Try
			{
				Move-Item -Path $FullLogPath -Destination ($LogPath + '\' + $LogName + '.Backup.log') -Force -ErrorAction SilentlyContinue
				OutputLog -Text ('"' + $LogFile + '" moved to "' + $LogName + '.Backup.log"')
			}
			Catch
			{
				OutputLog -Text ('Error while moving "' + $LogFile + '" to "' + $LogName + '.Backup.log') -Severity 'Error'
				OutputLog -Text ($_.Exception.Message) -Severity 'Error'
			}
		}
		Else { OutputLog -Text ('"' + $LogFile + '" is less than 5 MBs') }
		
		$Global:LogToFile = $True
	}
	Else { $Global:LogToFile = $False }
	
	OutputLog -Text ('[FUNCTION]>  InitSupportStructure  [STATE]>  End')
}

Function ReportEnvironment
{
	OutputLog -Text ('[FUNCTION]>  ReportEnvironment  [STATE]>  Start')
	
	OutputLog -Text ('PSVersion:                ' + ($PSVersionTable.PSVersion).ToString())
	OutputLog -Text ('Manufacturer:             ' + $Manufacturer)
	OutputLog -Text ('Model:                    ' + $Model)
	OutputLog -Text ('ChassisType:              ' + $ChassisType)
	If ($ChassisType -EQ 'Laptop') { OutputLog -Text ('OnACPower:                ' + $OnACPower) }
	OutputLog -Text ('ServiceTag:               ' + $ServiceTag)
	If (ObjectIsValid -Object $AssetTag) { OutputLog -Text ('AssetTag:                 ' + $AssetTag) }
	OutputLog -Text ('Processor:                ' + $Processor.ToString() + ' GHz')
	OutputLog -Text ('Memory:                   ' + $Memory.ToString() + ' GBs')
	OutputLog -Text ('OSName:                   ' + $OSName)
	OutputLog -Text ('OSType:                   ' + $OSType.ToString())
	OutputLog -Text ('OSVersion:                ' + $OSVersion)
	OutputLog -Text ('OSArch:                   ' + $OSArch)
	OutputLog -Text ('UserName:                 ' + $UserName)
	OutputLog -Text ('Domain:                   ' + $Domain)
	OutputLog -Text ('NTDomain:                 ' + $NTDomain)
	OutputLog -Text ('ComputerName:             ' + $ComputerName)
	OutputLog -Text ('TimeZone:                 ' + $TimeZone)
	OutputLog -Text ('ScriptFolder:             ' + $ScriptFolder)
	OutputLog -Text ('ScriptRoot:               ' + $ScriptRoot)
	OutputLog -Text ('LogName:                  ' + $LogName)
	OutputLog -Text ('LogFile:                  ' + $LogFile)
	OutputLog -Text ('LogPath:                  ' + $LogPath)
	OutputLog -Text ('FullLogPath:              ' + $FullLogPath)
	OutputLog -Text ('LogToFile:                ' + $LogToFile)
	OutputLog -Text ('LogToHost:                ' + $LogToHost)
	OutputLog -Text ('ExecutionType:            ' + $ExecutionType)
	
	OutputLog -Text ('[FUNCTION]>  ReportEnvironment  [STATE]>  End')
}

#endregion **** Module Setup Functions


#region **** Module Sub-Functions

Function ExecutionIsApproved([String]$name = $LogName, [String]$title = ($LogName + ' Execution'), [Bool]$packageAlreadyInstalled = $False)
{
	OutputLog -Text ('[FUNCTION]>  ExecutionIsApproved  [STATE]>  Start')

	[Bool]$isApproved = $False
	
	If (!(ObjectIsValid -Object $name)) { OutputLog -Text ('$Name is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $title)) { OutputLog -Text ('$Title is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $packageAlreadyInstalled)) { OutputLog -Text ('$PackageAlreadyInstalled is not valid') -Severity 'Error' }
	Else
	{
		[String]$text = (
							"Package not needed as " + $name + "`n" + `
							"was detected as installed, execution halted`n" + `
							"to prevent system issues.`n`n" + `
							"If your system does not have " + $name + "`n" + `
							"installed and this was detected in error,`n" + `
							"please contact the IT HelpDesk to investigate."
						)
		
		OutputLog -Text ('ExecutionType: ' + $ExecutionType)
		Switch ($ExecutionType)
		{
			'Prompt'
			{
				If ($packageAlreadyInstalled)
				{
					OutputLog -Text ('Package was detected as installed, displaying notification')
					[Windows.Forms.MessageBox]::Show($text, $title, 'OK', 64) | Out-Null
				}
				Else
				{
					OutputLog -Text ('Package is not already installed, proceeding with installation prompt')
					$text =	(
								"Proceed with " + $package + " Installation?`n`n" + `
								"Do not use any applications while this is running`n" + `
								"at any point! Do not try to halt/cancel`n" + `
								"the installation at any point, and do not reboot`n" + `
								"your machine until the installation finishes.`n" + `
								"If something goes wrong during the installation,`n" + `
								"please contact the HelpDesk to investigate.`n`n" + `
								"Please save all of your work and close all`n" + `
								"open applications BEFORE clicking Yes!"
							)
					
					Switch ([Windows.Forms.MessageBox]::Show($text, $title, 'YesNo', 32))
					{
						'Yes' { OutputLog -Text ('"Yes" selected, execution approved'); $isApproved = $True }
						'No' { OutputLog -Text ('"No" selected, execution denied') }
						Default { OutputLog -Text ('Unknown selection, assuming "No"') }
					}
				}
			}
			'Notify'
			{
				If ($packageAlreadyInstalled) { OutputLog -Text ('Package was detected as installed, displaying notification') }
				Else
				{
					$isApproved = $True
					$text = (
								$package + " Installation Has Started`n`n" + `
								"Do not use any applications while this is running`n" + `
								"at any point! Do not try to halt/cancel`n" + `
								"the installation at any point, and do not reboot`n" + `
								"your machine until the installation finishes.`n" + `
								"If something goes wrong during the installation,`n`n" + `
								"please contact the HelpDesk to investigate."
							)
				}
				
				Try { (New-Object -ComObject 'WScript.Shell' -ErrorAction SilentlyContinue).Popup($text, 30, $title, 64) | Out-Null }
				Catch
				{
					OutputLog -Text ('Error when trying to create WScript.Shell ComObject') -Severity 'Error'
					OutputLog -Text ($_.Exception.Message) -Severity 'Error'
				}
			}
			'Silent'
			{
				If ($packageAlreadyInstalled) { OutputLog -Text ('Package was detected as installed') }
				Else
				{
					OutputLog -Text ('ExecutionType set to Silent, execution approved and message box suppressed')
					$isApproved = $True
				}
			}
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  ExecutionIsApproved  [STATE]>  End')
	
	Return $isApproved
}

Function CompletionNotice([Bool]$success = $True, [Bool]$packageWasInstalled = $True, [Bool]$rebootNeeded = $False)
{
	OutputLog -Text ('[FUNCTION]>  CompletionNotice  [STATE]>  Start')
	
	If (!(ObjectIsValid -Object $success)) { OutputLog -Text ('$Success is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $packageWasInstalled)) { OutputLog -Text ('$PackageWasInstalled is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $rebootNeeded)) { OutputLog -Text ('$RebootNeeded is not valid') -Severity 'Error' }
	Else
	{
		[Int32]$icon = 64
		[Int32]$delay = 10
		[String]$text = 'Script execution has completed'
		
		OutputLog -Text ('ExecutionType: "' + $ExecutionType + '"')
		Switch ($ExecutionType)
		{
			'Prompt'
			{
				If (!$success)
				{
					OutputLog -Text ('Errors occurred during script execution, displaying notification') -Severity 'Error'
					$icon = 16
					$text = (
								"Errors occurred during script execution.`n" + `
								"See log for details:`n`n" + `
								"Log:  $FullLogPath"
							)
				}
				Else { OutputLog -Text ('Script execution appears successful, displaying notification') }
				
				[Windows.Forms.MessageBox]::Show($text, $title, 'OK', $icon) | Out-Null
				
				If (($success) -And ($packageWasInstalled) -And ($rebootNeeded))
				{
					OutputLog -Text ('Package requested a reboot upon successful completion, displaying prompt')
					$icon = 32
					$text = (
								"A system reboot is required to finalize`n" + `
								"installation. OK to reboot system now?`n" + `
								"If not, make sure you reboot your system`n" + `
								"at your earliest convenience as there may be`n" + `
								"usage issues until a reboot is performed!"
							)
					
					Switch ([Windows.Forms.MessageBox]::Show($text, $title, 'YesNo', $icon))
					{
						'Yes' { OutputLog -Text ('"Yes" selected'); $Global:RebootSelected = $True }
						'No' { OutputLog -Text ('"No" selected') }
						Default	{ OutputLog -Text ('Unknown selection, message box may have been closed or terminated by user, assuming "No"') }
					}
				}
			}
			'Notify'
			{
				If (!$success)
				{
					OutputLog -Text ('Errors occurred during script execution, displaying notification') -Severity 'Error'
					$icon = 16
					$delay = 30
					$text = (
								"Errors occurred during script execution.`n" + `
								"See log for details:`n`n" + `
								"Log:  $FullLogPath"
							)
				}
				Else { OutputLog -Text ('Script execution appears successful, displaying notification') }
				
				Try { (New-Object -ComObject 'WScript.Shell' -ErrorAction SilentlyContinue).Popup($text, $delay, $title, $icon) | Out-Null }
				Catch
				{
					OutputLog -Text ('Error when trying to create WScript.Shell ComObject') -Severity 'Error'
					OutputLog -Text ($_.Exception.Message) -Severity 'Error'
				}
				
				If (($success) -And ($packageWasInstalled) -And ($rebootNeeded))
				{
					OutputLog -Text ('Package requested a reboot upon successful completion, displaying notification')
					$icon = 64
					$delay = 20
					$text = (
								"A system reboot is required to finalize`n" + `
								"installation. Make sure you reboot your system`n" + `
								"at your earliest convenience as there may be`n" + `
								"usage issues until a reboot is performed!"
							)
							
					Try { (New-Object -ComObject 'WScript.Shell' -ErrorAction SilentlyContinue).Popup($text, $delay, $title, $icon) | Out-Null }
					Catch
					{
						OutputLog -Text ('Error when trying to create WScript.Shell ComObject') -Severity 'Error'
						OutputLog -Text ($_.Exception.Message) -Severity 'Error'
					}
				}
			}
			'Silent' { OutputLog -Text ('ExecutionType set to Silent, completion notice suppressed') }
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  CompletionNotice  [STATE]>  End')
}

#endregion **** Module Sub-Functions


#region **** Commonly Used Functions ****

Function OutputLog([String]$text, [Object]$severity = 1)
{
	If (($text -EQ $Null) -Or (!(ObjectIsValid -Object $severity))) { Return }
	
	Switch (($severity.ToString()).ToUpper())
	{
		'L'	 			{ $severity = 1 }
		'LOG' 			{ $severity = 1 }
		'W'				{ $severity = 2 }
		'WARN'			{ $severity = 2 }
		'WARNING' 		{ $severity = 2 }
		'E'				{ $severity = 3 }
		'ERR'			{ $severity = 3 }
		'ERROR' 		{ $severity = 3 }
		'V' 			{ $severity = 4 }
		'VERBOSE' 		{ $severity = 4 }
		'D' 			{ $severity = 5 }
		'DEBUG' 		{ $severity = 5 }
		'I' 			{ $severity = 6 }
		'INFO' 			{ $severity = 6 }
		'INFORMATION' 	{ $severity = 6 }
	}
	If (($severity -NE 1) -And ($severity -NE 2) -And ($severity -NE 3) -And ($severity -NE 4) -And ($severity -NE 5) -And ($severity -NE 6)) { $severity = 1 }
	
	[String]$context = $DefaultString
	[Object]$callStack = Get-PSCallStack -ErrorAction SilentlyContinue
	If (ObjectIsValid -Object $callStack)
	{
		If ($callStack.Length -GT 0)
		{
			If (ObjectIsValid -Object $callStack[1].Command)
			{
				If (($callStack[1].Command -NE '<ScriptBlock>') -And (ObjectIsValid -Object $MyInvocation.ScriptLineNumber))
				{
					$context = 'Function: ' + ($callStack[1].Command.ToString()) + ', Line: ' + ($MyInvocation.ScriptLineNumber.ToString())
				}
			}
		}
	}
	
	If ((((StripWhitespace -String $text).StartsWith('[FUNCTION]')) -Or ((StripWhitespace -String $text).StartsWith('[HEADER]'))) -And ((StripWhitespace -String $text) -Like '*STATE*>End*'))
	{
		$Global:LogIndentTier -= 1
	}
	If ($LogIndentTier -LT 0) { $Global:LogIndentTier = 0 }
	
	[String]$indent = [String]::Empty
	For ([UInt32]$i = 0; $i -NE ($LogIndentSize * $LogIndentTier); ++$i) { $indent += ' ' }
	
	$text = ($indent + $text)

	If (((StripWhitespace -String $text).StartsWith('[FUNCTION]')) -Or ((StripWhitespace -String $text).StartsWith('[HEADER]')))
	{
		OutputLog -Text ([String]::Empty)
		
		If ((StripWhitespace -String $text) -Like '*STATE*>Start*') { $Global:LogIndentTier += 1 }
	}

	[String]$logLine = 	'<![LOG[' + $text + ']LOG]!>' + `
						'<time="' + (Get-Date -Format HH:mm:ss) + '.' + ((Get-Date).Millisecond) + '+000' + '" ' + `
						'date="' + (Get-Date -Format MM-dd-yyyy) + '" ' + `
						'component="' + $context + '" ' + `
						'context="' + ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) + '" ' + `
						'type="' + $severity + '" ' + `
						'thread="' + $PID + '" ' + `
						'file="' + $LogFile + '">'
	
	If (($FullLogPath -NE $DefaultString) -And ($LogToFile))
	{
		If ($PreLogText.Count -GT 0)
		{
			For ([UInt32]$i = 0; $i -NE $PreLogText.Count; ++$i)
			{
				$PreLogText[$i] | Out-File -Append -Encoding UTF8 -FilePath $FullLogPath -ErrorAction SilentlyContinue
			}
			$Global:PreLogText.Clear()
		}
		$logLine | Out-File -Append -Encoding UTF8 -FilePath $FullLogPath -ErrorAction SilentlyContinue
	}
	Else { $PreLogText.Add($logLine) }
	
	[String]$color = 'White'
	Switch ($severity)
	{
		2 { $color = 'Yellow' }
		3 { $color = 'Red' }
		4 { $color = 'Green' }
		5 { $color = 'Magenta' }
		6 { $color = 'Cyan' }
	}
		
	If (!$LogToHost)
	{
		$Global:PreWriteHostText.Add($text)
		$Global:PreWriteHostTextColor.Add($color)
	}
	Else
	{
		If (($PreWriteHostText.Count -GT 0) -And ($PreWriteHostText.Count -EQ $PreWriteHostTextColor.Count))
		{
			For ([UInt32]$i = 0; $i -NE $PreWriteHostText.Count; ++$i)
			{
				Write-Host $PreWriteHostText[$i] -ForegroundColor $PreWriteHostTextColor[$i] -ErrorAction SilentlyContinue
			}
			$Global:PreWriteHostText.Clear()
			$Global:PreWriteHostTextColor.Clear()
		}
		Write-Host $text -ForegroundColor $color -ErrorAction SilentlyContinue
	}

	If (((StripWhitespace -String $text).StartsWith('[FUNCTION]')) -Or ((StripWhitespace -String $text).StartsWith('[HEADER]')))
	{
		OutputLog -Text ([String]::Empty)
	}
}

Function AddErrorToList([String]$text = $Global:NA, [Int32]$errorCode = -1)
{
	If ($text -EQ $Null) { OutputLog -Text ('$Text is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $errorCode)) { OutputLog -Text ('$ErrorCode is not valid') -Severity 'Error' }
	Else
	{
		$Global:ErrorState = $True
		$Global:ExitWithErrorText.Add($text)
		$Global:ExitWithErrorCodes.Add($errorCode)
	}
}

Function ObjectExists([String]$path, [String]$item = $DefaultString, [String]$itemType = $DefaultString, [Bool]$createIfNotExist = $False, [Bool]$logOnlyErrors = $True)
{
	[Bool]$exists = $False
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $item)) { OutputLog -Text ('$Item is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $itemType)) { OutputLog -Text ('$ItemType is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $createIfNotExist)) { OutputLog -Text ('$CreateIfNotExist is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $logOnlyErrors)) { OutputLog -Text ('$LogOnlyErrors is not valid') -Severity 'Error' }
	Else
	{
		If (!$logOnlyErrors) { OutputLog -Text ('Checking if "' + $path + '" exists') }
		If ($path -NE $DefaultString)
		{
			If (Test-Path -Path $path -ErrorAction SilentlyContinue)
			{
				If (!$logOnlyErrors) { OutputLog -Text ('"' + $path + '" exists') }
				$exists = $True
			}
			Else
			{
				OutputLog -Text ('"' + $path + '" does not exist')
				If ($createIfNotExist)
				{
					If (!$logOnlyErrors) { OutputLog -Text ('CreateIfNotExist set to True, attempting to create path') }
					Try
					{
						If (Test-Path -Path (New-Item -Path $path -ItemType 'Directory' -ErrorAction SilentlyContinue) -ErrorAction SilentlyContinue)
						{
							OutputLog -Text ('"' + $path + '" creation successful')
							$exists = $True
						}
						Else { OutputLog -Text ('"' + $path + '" creation unsuccessful') -Severity 'Error' }
					}
					Catch
					{
						OutputLog -Text ('Error creating directory') -Severity 'Error'
						OutputLog -Text ($_.Exception.Message) -Severity 'Error'
					}
				}
				Else { If (!$logOnlyErrors) { OutputLog -Text ('CreateIfNotExist set to False, path creation will not be attempted') } }
			}
			
			If (($exists) -And ($item -NE $DefaultString))
			{
				If ($itemType -EQ $DefaultString) { $itemType = 'File' }
			
				[String]$itemPath = $path
				If (!$itemPath.EndsWith('\')) { $itemPath += '\' }
				$itemPath += $item
				
				If (!$logOnlyErrors) { OutputLog -Text ('Checking if "' + $itemPath + '" exists') }
				If ($item -NE $DefaultString)
				{
					If (Test-Path -Path $itemPath -ErrorAction SilentlyContinue)
					{
						If (!$logOnlyErrors) { OutputLog -Text ('"' + $itemPath + '" exists') }
					}
					Else
					{
						OutputLog -Text ('"' + $itemPath + '" does not exist')
						If ($createIfNotExist)
						{
							If (!$logOnlyErrors) { OutputLog -Text ('CreateIfNotExist set to True, attempting to create item') }
							Try
							{
								If (Test-Path -Path (New-Item -Path $path -Name $item -ItemType $itemType -ErrorAction SilentlyContinue) -ErrorAction SilentlyContinue)
								{
									OutputLog -Text ('"' + $itemPath + '" creation successful')
								}
								Else
								{
									OutputLog -Text ('"' + $itemPath + '" creation unsuccessful') -Severity 'Error'
									$exists = $False
								}
							}
							Catch
							{
								OutputLog -Text ('Error creating item') -Severity 'Error'
								OutputLog -Text ($_.Exception.Message) -Severity 'Error'
							}
						}
						Else
						{
							If (!$logOnlyErrors) { OutputLog -Text ('CreateIfNotExist set to False, item creation will not be attempted') }
							$exists = $False
						}
					}
				}
			}
		}
		Else { OutputLog -Text ('Invalid Path: "' + $DefaultString + '"') -Severity 'Warning' }
	}
	
	Return $exists
}

Function ObjectIsValid([Object]$object)
{
	Return ($Null -NE $object)
}

Function SetNonNullString([String]$string)
{
	[String]$retString = $DefaultString
	If (ObjectIsValid -Object $string) { $retString = $string }
	
	Return $retString
}

Function PackageIsInstalled([String]$name = $LogName, [String[]]$packageObjectsToCheck, [Bool]$validateAll = $False)
{
	OutputLog -Text ('[FUNCTION]>  PackageIsInstalled  [STATE]>  Start')
	
	[Bool]$isInstalled = $False
	
	If (!(ObjectIsValid -Object $name)) { OutputLog -Text ('$Name is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $packageObjectsToCheck)) { OutputLog -Text ('$PackageObjectsToCheck is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $validateAll)) { OutputLog -Text ('$ValidateAll is not valid') -Severity 'Error' }
	ElseIf ($name -EQ $DefaultString) { OutputLog -Text ('$Name is not valid') -Severity 'Error' }
	ElseIf ($packageObjectsToCheck.Length -EQ 0) { OutputLog -Text ('$PackageObjectsToCheck is empty') -Severity 'Warning' }
	Else
	{
		[System.Collections.Generic.List[Bool]]$checkedObjectArray = New-Object System.Collections.Generic.List[Bool]
		
		For ([UInt32]$i = 0; $i -NE $packageObjectsToCheck.Length; ++$i)
		{
			OutputLog -Text ('Checking if "' + $packageObjectsToCheck[$i] + '" exists')
			If (ObjectExists -Path $packageObjectsToCheck[$i]) { $checkedObjectArray.Add($True) }
			ElseIf (($packageObjectsToCheck[$i].StartsWith('HKCU:')) -Or ($packageObjectsToCheck[$i].StartsWith('HKLM:')))
			{
				Try
				{
					[String]$parent = Split-Path -Path $packageObjectsToCheck[$i] -Parent -ErrorAction SilentlyContinue
					[String]$leaf = Split-Path -Path $packageObjectsToCheck[$i] -Leaf -ErrorAction SilentlyContinue
					OutputLog -Text ('Checking if "' + $parent + '" exists as a registry key and has value "' + $leaf + '"')
					If (RegKeyExists -RegKey $parent -RegValue $leaf) { $checkedObjectArray.Add($True) }
					Else
					{
						OutputLog -Text ('"' + $parent + '" not found with value "' + $leaf + '"')
						
						[String]$packageObjectToCheck2 = $DefaultString
						If ($packageObjectsToCheck[$i] -Like '*Wow6432Node*')
						{
							$packageObjectToCheck2 = ($packageObjectsToCheck[$i] -Replace 'Software\Wow6432Node', 'Software')
						}
						Else
						{
							$packageObjectToCheck2 = ($packageObjectsToCheck[$i] -Replace 'Software', 'Software\Wow6432Node')
						}
						
						OutputLog -Text ('Checking if "' + $packageObjectToCheck2 + '" exists')
						If (ObjectExists -Path $packageObjectToCheck2) { $checkedObjectArray.Add($True) }
						Else
						{
							$parent = Split-Path -Path $packageObjectToCheck2 -Parent -ErrorAction SilentlyContinue
							$leaf = Split-Path -Path $packageObjectToCheck2 -Leaf -ErrorAction SilentlyContinue
							OutputLog -Text ('Checking if "' + $parent + '" exists as a registry key and has value "' + $leaf + '"')
							If (RegKeyExists -RegKey $parent -RegValue $leaf) { $checkedObjectArray.Add($True) }
							Else { $checkedObjectArray.Add($False) }
						}
					}
				}
				Catch
				{
					OutputLog -Text ('Error while looking for "' + $packageObjectsToCheck[$i] + '"') -Severity 'Error'
					OutputLog -Text ($_.Exception.Message) -Severity 'Error'
					$checkedObjectArray.Add($False)
				}
			}
		}
		
		If ((ObjectIsValid -Object $checkedObjectArray) -And ($checkedObjectArray.Count -GT 0))
		{
			If ($validateAll)
			{
				$isInstalled = $True
				For ([UInt32]$i = 0; $i -NE $checkedObjectArray.Count; ++$i)
				{
					If (!$checkedObjectArray[$i])
					{
						$isInstalled = $False
						Break
					}
				}
			}
			Else
			{
				For ([UInt32]$i = 0; $i -NE $checkedObjectArray.Count; ++$i)
				{
					If ($checkedObjectArray[$i])
					{
						$isInstalled = $True
						Break
					}
				}
			}
		}
		
		If ($isInstalled) { OutputLog -Text ($name + ' installation found') }
		Else { OutputLog -Text ($name + ' installation not found') }
	}
	
	OutputLog -Text ('[FUNCTION]>  PackageIsInstalled  [STATE]>  End')
	
	Return $isInstalled
}

Function PackageIsValid([String]$parentPath = $DefaultString, [String[]]$pathArray)
{
	OutputLog -Text ('[FUNCTION]>  PackageIsValid  [STATE]>  Start')
	
	[Bool]$isValid = $True
	If (!(ObjectIsValid -Object $parentPath))
	{
		OutputLog -Text ('$ParentPath is not valid, unable to validate package') -Severity 'Warning'
		$isValid = $False
	}
	ElseIf (!(ObjectIsValid -Object $pathArray))
	{
		OutputLog -Text ('$PathArray is not valid, unable to validate package') -Severity 'Warning'
		$isValid = $False
	}
	Else
	{
		For ([UInt32]$i = 0; $i -NE $pathArray.Length; ++$i)
		{
			[String]$path = $pathArray[$i]
			If ($parentPath -NE $DefaultString)
			{
				If (!$parentPath.EndsWith('\')) { $parentPath += '\' }
				$path = $parentPath + $pathArray[$i]
			}
			If (!(ObjectExists -Path $path))
			{
				$isValid = $False
				Break
			}
		}
		
		If ($isValid) { OutputLog -Text ('Package is valid') }
		Else { OutputLog -Text ('Package is not valid') -Severity 'Error' }
	}
	
	OutputLog -Text ('[FUNCTION]>  PackageIsValid  [STATE]>  End')
	
	Return $isValid
}

Function ValidatePackageHashXML([String]$xmlPath = ($ScriptRoot + '\Hashes.xml'))
{
	OutputLog -Text ('[FUNCTION]>  ValidatePackageHashXML  [STATE]>  Start')
	
	[Bool]$hashesValid = $False
	[System.Collections.Generic.List[Bool]]$fileHashBoolArray = New-Object System.Collections.Generic.List[Bool]
	
	If (!(ObjectIsValid -Object $xmlPath)) { OutputLog -Text ('$XMLPath is not valid') -Severity 'Error' }
	Else
	{
		If ($PSVersionTable.PSVersion.Major -GT 2)
		{
			If ((!(ObjectIsValid -Object $xmlPath)) -OR ($xmlPath -EQ $DefaultString))
			{
				OutputLog -Text ('Unable to validate package hash XML as file path to hash table was not valid') -Severity 'Warning'
			}
			Else
			{
				OutputLog -Text ('Checking if "' + $xmlPath + '" exists')
				If (ObjectExists -Path $xmlPath)
				{
					OutputLog -Text ('"' + $hashXMLPath + '" found')
				
					Try
					{
						[XML]$hashXML = Get-Content -Path $hashXMLPath -ErrorAction SilentlyContinue
						If (ObjectIsValid -Object $hashXML)
						{
							For ([UInt32]$i = 0; $i -NE $hashXML.Package.File.Length; ++$i)
							{
								[String]$file = ($ScriptRoot + '\' + $hashXML.Package.File[$i].Name)
								If (ObjectExists -Path $file)
								{
									[Bool]$isHashValid = $True
									OutputLog -Text ('File:  ' + $hashXML.Package.File[$i].Name)
									For ([UInt32]$j = 0; $j -NE $hashXML.Package.File[$i].Hash.Length; ++$j)
									{
										[String]$returnedHash = (Get-FileHash -Path $file -Algorithm ($hashXML.Package.File[$i].Hash[$j].Algorithm) -ErrorAction SilentlyContinue).Hash
										[String]$expectedHash = ($hashXML.Package.File[$i].Hash[$j].'#text')
										If (($returnedHash) -EQ ($expectedHash))
										{
											OutputLog -Text ('->  ' + $hashXML.Package.File[$i].Hash[$j].Algorithm + ':  ' + $returnedHash)
										}
										Else
										{
											OutputLog -Text ('->  ' + $hashXML.Package.File[$i].Hash[$j].Algorithm + ':  ' + $returnedHash) -Severity 'Error'
											$isHashValid = $False
										}
										
										$fileHashBoolArray.Add($isHashValid)
									}
								}
								Else { OutputLog -Text ('Unable to validate hashes for "' + $file + '"') -Severity 'Warning' }
							}
						}
						Else { OutputLog -Text ('Get-Content for "' + $hashXMLPath + '" returned null, unable to validate hash XML') -Severity 'Warning' }
					}
					Catch
					{
						OutputLog -Text ('Error while getting content from "' + $hashXMLPath + '"') -Severity 'Error'
						OutputLog -Text ($_.Exception.Message) -Severity 'Error'
					}
				}
				Else { OutputLog -Text ('Unable to validate package hash XML as file was not found') -Severity 'Warning' }
			}
		}
		Else { OutputLog -Text ('Get-FileHash not supported below PowerShell version 3.0, unable to validate package hashes') -Severity 'Warning' }
		
		OutputLog -Text ('[FUNCTION]>  ValidatePackageHashXML  [STATE]>  End')
		
		If ((ObjectIsValid -Object $fileHashBoolArray) -And ($fileHashBoolArray.Count -GT 0))
		{
			$hashesValid = $True
			For ([UInt32]$i = 0; $i -NE $fileHashBoolArray.Count; ++$i)
			{
				If (!$fileHashBoolArray[$i]) { $hashesValid = $False; Break }
			}
		}
	}
	
	Return $hashesValid
}

Function ModifyVersionNumber([String]$versionString, [Int32]$major = $Null, [Int32]$minor = $Null, [Int32]$build = $Null, [Int32]$revision = $Null)
{
	[String]$retVar = $DefaultString
	
	If (!(ObjectIsValid -Object $versionString)) { OutputLog -Text ('$VersionString is not valid') -Severity 'Error' }
	Else
	{
		$retVar = $versionString
		
		Try
		{
			[Int32]$newMajor = 0
			[Int32]$newMinor = 0
			[Int32]$newBuild = 0
			[Int32]$newRevision = 0
			[System.Version]$versionObject = $versionString
			
			If (!(ObjectIsValid -Object $versionObject)) { OutputLog -Text ('$VersionObject is not valid') -Severity 'Error' }
			Else
			{
				If (ObjectIsValid -Object $major) { $newMajor = $major }
				Else { $newMajor = $versionObject.Major }
				
				If (ObjectIsValid -Object $minor) { $newMinor = $minor }
				Else { $newMinor = $versionObject.Minor }
				
				If (ObjectIsValid -Object $build) { $newBuild = $build }
				Else { $newBuild = $versionObject.Build }
				
				If (ObjectIsValid -Object $revision) { $newRevision = $revision }
				Else { $newRevision = $versionObject.Revision }
				
				$retVer = ($newMajor.ToString() + '.' + $newMinor.ToString() + '.' + $newBuild.ToString() + '.' + $newRevision.ToString())
			}
		}
		Catch
		{
			OutputLog -Text ('Error building new version object') -Severity 'Error'
			OutputLog -Text ($_.Exception.Message) -Severity 'Error'
		}
	}
	
	Return $retVar
}

Function ImportJSON([String]$path)
{
	[Object]$jsonData = $Null
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	Else
	{
		[Object]$jsonFile = $Null
		Try
		{
			If (ObjectExists -Path $path)
			{
				$jsonFile = Get-Content -Path $path -ErrorAction SilentlyContinue
				If (!(ObjectIsValid -Object $jsonFile)) { OutputLog -Text ('$JSONFile is not valid') -Severity 'Error' }
				Else
				{
					If ($PSVersionTable.PSVersion.Major -GT 2)
					{
						$jsonData = ConvertFrom-Json -InputObject $jsonFile -ErrorAction SilentlyContinue
					}
					Else { $jsonData = ((New-Object System.Web.Script.Serialization.JavaScriptSerializer).DeserializeObject($jsonFile)) }
				}
			}
			Else { OutputLog -Text ('Unable to import JSON file data') -Severity 'Error' }
		}
		Catch
		{
			OutputLog -Text ('Error importing JSON file') -Severity 'Error'
			OutputLog -Text ($_.Exception.Message) -Severity 'Error'
			$jsonData = $Null
		}
		
		$jsonFile = $Null
	}
	
	Return $jsonData
}

Function CachePackage([String]$source, [String]$destination = ($LogPath + '\' + $ScriptName), [Bool]$mirror = $False, [Bool]$overwrite = $False)
{
	OutputLog -Text ('[FUNCTION]>  CachePackage  [STATE]>  Start')
	
	[Bool]$isPackageCached = $False
	
	If (!(ObjectIsValid -Object $source)) { OutputLog -Text ('$Source is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $destination)) { OutputLog -Text ('$Destination is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $mirror)) { OutputLog -Text ('$Mirror is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $overwrite)) { OutputLog -Text ('$Overwrite is not valid') -Severity 'Error' }
	Else
	{
		OutputLog -Text ('Checking if package is cached')
		If ($destination -EQ $ScriptRoot)
		{
			OutputLog -Text ('Package is already cached')
			$isPackageCached = $True
		}
		Else
		{
			OutputLog -Text ('Package is not currently cached')
			If (RobocopyPath -Source $source -Destination $destination -Mirror $mirror -Overwrite $overwrite)
			{
				OutputLog -Text ('Package was successfully cached to "' + $destination + '"')
				$isPackageCached = $True
			}
			Else { OutputLog -Text ('Package was unsuccessfully cached to "' + $destination + '"') -Severity 'Error' }
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  CachePackage  [STATE]>  End')
	
	Return $isPackageCached
}

Function ChangeOwner([String]$path, [String]$newOwner = 'Administrators')
{
	OutputLog -Text ('[FUNCTION]>  ChangeOwner  [STATE]>  Start')
	
	[Bool]$ownerChanged = $False
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $newOwner)) { OutputLog -Text ('$NewOwner is not valid') -Severity 'Error' }
	Else
	{
		OutputLog -Text ('Checking if "' + $path + '" exists')
		If (ObjectExists -Path $path)
		{
			OutputLog -Text ('"' + $path + '" exists')
			OutputLog -Text ('Attempting to grant ownership to ' + $newOwner)
			If ((RunProcess -FilePath ('takeown.exe') -ArgList ('/f "' + $path + '"') -Attempts 1) -EQ 0)
			{
				If ((RunProcess -FilePath ('icacls.exe') -ArgList ('"' + $path + '" /Grant ' + $newOwner + ':(F)') -Attempts 1) -EQ 0)
				{
					OutputLog -Text ('Grant ownership successful')
					$ownerChanged = $True
				}
				Else { OutputLog -Text OutputLog -Text ('Grant ownership failed using "icacls.exe"') -Severity 'Error' }
			}
			Else { OutputLog -Text ('Grant ownership failed using "takeown.exe"') -Severity 'Error' }
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  ChangeOwner  [STATE]>  End')
	
	Return $ownerChanged
}

Function BackupPath([String]$path)
{
	[Bool]$wasBackedUp = $False
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	Else
	{
		OutputLog -Text ('Checking if "' + $path + '" exists')
		If (ObjectExists -Path $path)
		{
			OutputLog -Text ('"' + $path + '" exists')
			
			OutputLog -Text ('Attempting path backup')
			[String]$destination = ($path + '.' + ((Get-Date -Format 'yyyy-M-d_h-m-s-fff').ToString()) + '.Backup')
			OutputLog -Text ('Source:      ' + $path)
			OutputLog -Text ('Destination: ' + $destination)
			Try
			{
				Copy-Item -Path $path -Destination $destination -Recurse -Force -ErrorAction SilentlyContinue
				If (ObjectExists -Path $destination)
				{
					OutputLog -Text ('Backup successful')
					$wasBackedUp = $True
				}
				Else
				{
					OutputLog -Text ('Backup unsuccessful') -Severity 'Warning'
				}
			}
			Catch
			{
				OutputLog -Text ('Error backing up "' + $path + '"') -Severity 'Error'
				OutputLog -Text ($_.Exception.Message) -Severity 'Error'
			}
		}
		Else { OutputLog -Text ('Unable to perform backup') -Severity 'Warning' }
	}
	
	Return $wasBackedUp
}

Function DeletePath([String]$path)
{
	[Bool]$wasDeleted = $False
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	Else
	{
		OutputLog -Text ('Checking if "' + $path + '" exists')
		If (ObjectExists -Path $path)
		{
			OutputLog -Text ('"' + $path + '" exists')
			
			OutputLog -Text ('Attempting path deletion')
			Try
			{
				Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
				If (!(Test-Path -Path $path -ErrorAction SilentlyContinue))
				{
					OutputLog -Text ('Deletion successful')
					$wasDeleted = $True
				}
				Else
				{
					OutputLog -Text ('Deletion unsuccessful') -Severity 'Warning'
				}
			}
			Catch
			{
				OutputLog -Text ('Error deleting "' + $path + '"') -Severity 'Error'
				OutputLog -Text ($_.Exception.Message) -Severity 'Error'
			}
		}
	}
	
	Return $wasDeleted
}

Function CopyFile([String]$source, [String]$destination, [Bool]$overwrite = $False, [Bool]$removeReadOnly = $True)
{
	[Bool]$shouldCopy = $False
	[Bool]$wasCopied = $False

	If (!(ObjectIsValid -Object $source)) { OutputLog -Text ('$Source is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $destination)) { OutputLog -Text ('$Destination is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $overwrite)) { OutputLog -Text ('$Overwrite is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $removeReadOnly)) { OutputLog -Text ('$RemoveReadOnly is not valid') -Severity 'Error' }
	Else
	{
		OutputLog -Text ('Checking source file path "' + $source + '"')
		If (ObjectExists -Path $source)
		{
			OutputLog -Text ('Source file path found, checking destination path "' + $destination + '"')
			If (ObjectExists -Path $destination)
			{
				If ($overwrite)
				{
					OutputLog -Text ('Destination file "' + $destination + '" was found but overwrite set to true, attempting file copy')
					$shouldCopy = $True
				}
				Else
				{
					OutputLog -Text ('Destination file "' + $destination + '" was found and overwrite set to false, file copy will not be attempted') -Severity 'Warning'
					$shouldCopy = $False
				}
			}
			Else
			{
				OutputLog -Text ('Destination file "' + $destination + '" was not found, attempting file copy')
				$shouldCopy = $True
			}
		}
		Else
		{
			OutputLog -Text ('Source file "' + $source + '" was not found, unable to copy file') -Severity 'Warning'
			$shouldCopy = $False
		}
		
		If ($shouldCopy)
		{
			Try
			{
				[String]$parentPath = (Split-Path -Parent $destination)
				If (!(Test-Path -Path $parentPath))
				{
					New-Item -Path $parentPath -ItemType 'Directory' -Force -ErrorAction SilentlyContinue
				}
				
				If ($removeReadOnly)
				{
					Copy-Item -Path $source -Destination $destination -Recurse -Force -PassThru -ErrorAction SilentlyContinue | Set-ItemProperty -Name 'IsReadOnly' -Value $False -ErrorAction SilentlyContinue
				}
				Else
				{
					Copy-Item -Path $source -Destination $destination -Recurse -Force -PassThru -ErrorAction SilentlyContinue
				}
				
				If (ObjectExists -Path $destination) { $wasCopied = $True }
				Else { $wasCopied = $False }
			}
			Catch
			{
				OutputLog -Text ($_.Exception.Message) -Severity 'Error'
				$wasCopied = $False
			}
		}
		
		If ($wasCopied) { OutputLog -Text ('Copy successful') }
		Else { OutputLog -Text ('Copy unsuccessful') -Severity 'Warning' }
	}
	
	Return $wasCopied
}

Function ComparePaths([String]$source, [String]$destination, [String]$file = $DefaultString, [UInt32]$retries = 10, [UInt32]$wait = 5)
{
	OutputLog -Text ('[FUNCTION]>  ComparePaths  [STATE]>  Start')
	
	[Bool]$compareValid = $False
	[String]$type = 'Folder'
	
	If (!(ObjectIsValid -Object $source)) { OutputLog -Text ('$Source is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $destination)) { OutputLog -Text ('$Destination is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $file)) { OutputLog -Text ('$File is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $retries)) { OutputLog -Text ('$Retries is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $wait)) { OutputLog -Text ('$Wait is not valid') -Severity 'Error' }
	Else
	{
		While (($source.Length -GT 2) -And ($source.EndsWith('\'))) { $source = $source.Substring(0, $source.Length - 2) }
		While (($destination.Length -GT 2) -And ($destination.EndsWith('\'))) { $destination = $destination.Substring(0, $destination.Length - 2) }
		
		[String]$fullSource = $source
		[String]$fullDest = $destination
		If ($file -NE $DefaultString)
		{
			$fullSource = ($fullSource + '\' + $file)
			$fullDest = ($fullDest + '\' + $file)
			$type = 'File'
		}
		
		OutputLog -Text ('Checking if source ' + $type.ToLower() + ' "' + $fullSource + '" exists')
		If (ObjectExists -Path $fullSource)
		{
			OutputLog -Text ('Source ' + $type.ToLower() + ' exists')
			OutputLog -Text ('Checking if destination ' + $type.ToLower() + ' "' + $fullDest + '" exists')
			If (ObjectExists -Path $fullDest)
			{
				OutputLog -Text ('Destination ' + $type.ToLower() + ' exists')
				
				[String]$argList = ('/L "' + $source + '" "' + $destination + '" /R:' + $retries.ToString() + ' /W:' + $wait.ToString())
				If ($file -NE $DefaultString)
				{
					$argList = ('/L "' + $source + '" "' + $destination + '" "' + $file + '" /R:' + $retries.ToString() + ' /W:' + $wait.ToString())
				}
				
				[Int32]$retCode = (RunProcess -FilePath 'Robocopy.exe' -ArgList $argList -Attempts 1 -WaitTime 0)
				If ($retCode -NE 0)
				{
					OutputLog -Text ($type + ' compare not valid')
					$compareValid = $False
				}
				Else
				{
					OutputLog -Text ($type + ' compare valid')
					$compareValid = $True
				}
			}
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  ComparePaths  [STATE]>  End')
	
	Return $compareValid
}

Function RobocopyPath([String]$source, [String]$destination, [String]$file = $DefaultString, [UInt32]$retries = 10, [UInt32]$wait = 5, [Bool]$mirror = $False, [Bool]$overwrite = $False)
{
	OutputLog -Text ('[FUNCTION]>  RobocopyPath  [STATE]>  Start')
	
	[Bool]$shouldCopy = $False
	[Bool]$wasCopied = $False
	[String]$type = 'Folder'
	
	If (!(ObjectIsValid -Object $source)) { OutputLog -Text ('$Source is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $destination)) { OutputLog -Text ('$Destination is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $file)) { OutputLog -Text ('$File is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $retries)) { OutputLog -Text ('$Retries is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $wait)) { OutputLog -Text ('$Wait is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $mirror)) { OutputLog -Text ('$Mirror is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $overwrite)) { OutputLog -Text ('$Overwrite is not valid') -Severity 'Error' }
	Else
	{
		While (($source.Length -GT 2) -And ($source.EndsWith('\'))) { $source = $source.Substring(0, $source.Length - 2) }
		While (($destination.Length -GT 2) -And ($destination.EndsWith('\'))) { $destination = $destination.Substring(0, $destination.Length - 2) }
		
		[String]$fullSource = $source
		[String]$fullDest = $destination
		If ($file -NE $DefaultString)
		{
			$fullSource = ($fullSource + '\' + $file)
			$fullDest = ($fullDest + '\' + $file)
			$type = 'File'
		}
		
		OutputLog -Text ('Checking if source ' + $type.ToLower() + ' "' + $fullSource + '" exists')
		If (ObjectExists -Path $fullSource)
		{
			OutputLog -Text ('Source ' + $type.ToLower() + ' exists')
			OutputLog -Text ('Checking if destination ' + $type.ToLower() + ' "' + $fullDest + '" exists')
			If (ObjectExists -Path $fullDest)
			{
				OutputLog -Text ('Destination ' + $type.ToLower() + ' exists')
				If ($overwrite)
				{
					OutputLog -Text ('Overwrite set to True, ' + $type.ToLower() + ' copy will be attempted')
					$shouldCopy = $True
				}
				Else
				{
					OutputLog -Text ('Overwrite set to False, ' + $type.ToLower() + ' copy will not be attempted') -Severity 'Warning'
					$shouldCopy = $False
				}
			}
			Else { $shouldCopy = $True }
		}
		
		If ($shouldCopy)
		{
			[String]$argList = ('/E "' + $source + '" "' + $destination + '" /R:' + $retries.ToString() + ' /W:' + $wait.ToString())
			If ($mirror) { $argList = ('/MIR "' + $source + '" "' + $destination + '" /R:' + $retries.ToString() + ' /W:' + $wait.ToString()) }
			If ($file -NE $DefaultString)
			{
				$argList = ('"' + $source + '" "' + $destination + '" "' + $file + '" /R:' + $retries.ToString() + ' /W:' + $wait.ToString())
			}
			
			[Int32]$retCode = (RunProcess -FilePath 'Robocopy.exe' -ArgList $argList -Attempts 1 -WaitTime 0)
			If (($retCode -LT 0) -Or ($retCode -GT 7))
			{
				OutputLog -Text ($type + ' was copied unsuccessfully') -Severity 'Error'
				$wasCopied = $False
			}
			Else
			{
				OutputLog -Text ($type + ' was copied successfully')
				$wasCopied = $True
			}
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  RobocopyPath  [STATE]>  End')
	
	Return $wasCopied
}

Function AppendTextToFile([String]$path, [String]$text, [Bool]$force = $False)
{
	[Bool]$textAppended = $False
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	ElseIf ($text -EQ $Null) { OutputLog -Text ('$Text is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $force)) { OutputLog -Text ('$Force is not valid') -Severity 'Error' }
	Else
	{
		OutputLog -Text ('Checking if "' + $path + '" exists')
		If (ObjectExists -Path $path)
		{
			OutputLog -Text ('"' + $path + '" exists')
			OutputLog -Text ('Appending "' + $text + '" to end of file')
			Try
			{
				If ($Force)
				{
					If (Add-Content -Path $path -Value $text -PassThru -Force -ErrorAction SilentlyContinue)
					{
						OutputLog -Text ('Text append successful')
						$textAppended = $True
					}
					Else
					{
						OutputLog -Text ('Text append unsuccessful')
					}
				}
				Else
				{
					If (Add-Content -Path $path -Value $text -PassThru -ErrorAction SilentlyContinue)
					{
						OutputLog -Text ('Text append successful')
						$textAppended = $True
					}
					Else
					{
						OutputLog -Text ('Text append unsuccessful')
					}
				}
			}
			Catch
			{
				OutputLog -Text ('Error while appending text') -Severity 'Error'
				OutputLog -Text ($_.Exception.Message) -Severity 'Error'
			}
		}
		Else
		{
			OutputLog -Text ('Unable to append text to file') -Severity 'Error'
		}
	}
	
	Return $textAppended
}

Function GetValidRegTree([String]$regKey = $DefaultString, [Bool]$withoutColon = $False)
{
	If (!(ObjectIsValid -Object $regKey)) { OutputLog -Text ('$RegKey is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $withoutColon)) { OutputLog -Text ('$WithoutColon is not valid') -Severity 'Error' }
	Else
	{
		[String[]]$keySplit = $regKey -Split '\\'
		If ($keySplit.Length -GT 0)
		{
			Switch (($keySplit[0]).ToUpper())
			{
				'HKCU:'					{ $regKey = 'HKCU:' }
				'HKEY_CURRENT_USER'		{ $regKey = 'HKCU:' }
				'HKLM:'					{ $regKey = 'HKLM:' }
				'HKEY_LOCAL_MACHINE'	{ $regKey = 'HKLM:' }
				Default					{ $regKey = $DefaultString }
			}
		}
		Else { $regKey = $DefaultString }
		
		If ($regKey -NE $DefaultString)
		{
			If ($withoutColon) { $regKey = $regKey.Substring(0, $regKey.Length - 1) }
			For ([UInt32]$i = 1; $i -NE $keySplit.Length; ++$i) { $regKey += ('\' + $keySplit[$i]) }
		}
		Else { OutputLog -Text ('Registry tree is not valid') -Severity 'Warning' }
	}
	
	Return $regKey
}

Function RegKeyExists([String]$regKey, [String]$regValue = $DefaultString, [Bool]$logOnlyErrors = $True)
{
	[Bool]$exists = $False
	
	If (!(ObjectIsValid -Object $regKey)) { OutputLog -Text ('$RegKey is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $regValue)) { OutputLog -Text ('$RegValue is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $logOnlyErrors)) { OutputLog -Text ('$LogOnlyErrors is not valid') -Severity 'Error' }
	Else
	{
		$regKey = GetValidRegTree -RegKey $regKey
		If ($regKey -NE $DefaultString)
		{
			If (!$logOnlyErrors) { OutputLog -Text ('Checking if key "' + $regKey + '" exists') }
			Try
			{
				If (Test-Path -Path $regKey -ErrorAction SilentlyContinue)
				{
					If (!$logOnlyErrors) { OutputLog -Text ('"' + $regKey + '" exists') }
					$exists = $True
					
					If ($regValue -NE $DefaultString)
					{
						If (!$logOnlyErrors) { OutputLog -Text ('Checking if key "' + $regKey + '" has value "' + $regValue + '"') }
						If ((Get-ItemProperty -Path $regKey -ErrorAction SilentlyContinue).$regValue)
						{
							If (!$logOnlyErrors) { OutputLog -Text ('"' + $regValue + '" exists') }
							$exists = $True
						}
						Else
						{
							OutputLog -Text ('"' + $regValue + '" does not exist')
							$exists = $False
						}
					}
				}
				Else { OutputLog -Text ('"' + $regKey + '" does not exist') }
			}
			Catch
			{
				OutputLog -Text ('Error checking registry key') -Severity 'Error'
				OutputLog -Text ($_.Exception.Message) -Severity 'Error'
			}
		}
		Else { OutputLog -Text ('"' + $regKey + '" does not have a valid tree') -Severity 'Warning' }
	}
	
	Return $exists
}

Function GetRegValueData([String]$regKey, [String]$regValue)
{
	[Object]$data = $Null
	
	If (!(ObjectIsValid -Object $regKey)) { OutputLog -Text ('$RegKey is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $regValue)) { OutputLog -Text ('$RegValue is not valid') -Severity 'Error' }
	Else
	{
		$regKey = GetValidRegTree -RegKey $regKey
		If ($regKey -NE $DefaultString)
		{
			If (RegKeyExists -RegKey $regKey -RegValue $regValue)
			{
				Try { $data = (Get-ItemProperty -Path $regKey -ErrorAction SilentlyContinue).$regValue }
				Catch
				{
					OutputLog -Text ('Unable to get value data for "' + $regKey + '"') -Severity 'Error'
					OutputLog -Text ($_.Exception.Message) -Severity 'Error'
				}
			}
		}
		Else { OutputLog -Text ('"' + $regKey + '" does not have valid data') -Severity 'Warning' }
	}
	
	Return $data
}

Function SetRegKey([String]$regKey, [String]$regValue = $DefaultString, [String]$valueData = $DefaultString, [String]$valueType = $DefaultString, [Bool]$overwrite = $False)
{
	[Bool]$keySet = $False
	[Bool]$valueToBeSet = $False
	[Bool]$valueSet = $False
	
	If (!(ObjectIsValid -Object $regKey)) { OutputLog -Text ('$RegKey is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $regValue)) { OutputLog -Text ('$RegValue is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $valueType)) { OutputLog -Text ('$ValueType is not valid') -Severity 'Error' }
	ElseIf ($valueData -EQ $Null) { OutputLog -Text ('$ValueData is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $overwrite)) { OutputLog -Text ('$Overwrite is not valid') -Severity 'Error' }
	Else
	{
		$regKey = GetValidRegTree -RegKey $regKey
		If ($regKey -NE $DefaultString)
		{
			If (RegKeyExists -RegKey $regKey) { $keySet = $True }
			Else
			{
				OutputLog -Text ('Creating key "' + $regKey + '"')
				Try
				{
					If (RegKeyExists -RegKey ((New-Item -Path $regKey -Force -ErrorAction SilentlyContinue).Name) -LogOnlyErrors $False)
					{
						OutputLog -Text ('Key creation successful')
						$keySet = $True
					}
					Else { OutputLog -Text ('Key creation unsuccessful') -Severity 'Error' }
				}
				Catch
				{
					OutputLog -Text ('Error creating registry key') -Severity 'Error'
					OutputLog -Text ($_.Exception.Message) -Severity 'Error'
				}
			}
		
			If (($regValue -EQ $DefaultString) -And ($valueType -EQ $DefaultString) -And ($valueData -EQ $DefaultString))
			{
				$valueSet = $True
			}
			ElseIf ($keySet)
			{
				[Bool]$typeIsValid = $False
				OutputLog -Text ('Validating value type "' + $valueType + '" is valid')
				Switch ($valueType.ToUpper())
				{
					'STRING' 		{ $valueType = 'String'; 		$typeIsValid = $True }
					'EXPANDSTRING' 	{ $valueType = 'ExpandString'; 	$typeIsValid = $True }
					'BINARY' 		{ $valueType = 'Binary'; 		$typeIsValid = $True }
					'DWORD' 		{ $valueType = 'DWord'; 		$typeIsValid = $True }
					'MULTISTRING' 	{ $valueType = 'MultiString'; 	$typeIsValid = $True }
					'QWORD' 		{ $valueType = 'QWord'; 		$typeIsValid = $True }
					Default			{ OutputLog -Text ('"' + $valueType + '" is not valid') -Severity 'Warning' }
				}
				
				If ($typeIsValid)
				{
					OutputLog -Text ('"' + $valueType + '" is valid')
					If (RegKeyExists -RegKey $regKey -RegValue $regValue)
					{
						If ($overwrite) { $valueToBeSet = $True }
						Else
						{
							OutputLog -Text ('$Overwrite set to $False, key value will not be set') -Severity 'Warning'
							$valueToBeSet = $False
						}
					}
					Else { $valueToBeSet = $True }
					
					If ($valueToBeSet)
					{
						OutputLog -Text ('Setting key value:')
						OutputLog -Text ('->  Key:    ' + $regKey)
						OutputLog -Text ('->  Value:  ' + $regValue)
						OutputLog -Text ('->  Type:   ' + $valueType)
						OutputLog -Text ('->  Data:   ' + $valueData)
						Try
						{
							If ($valueData -EQ ((New-ItemProperty -Path $regKey -Name $regValue -PropertyType $valueType -Value $valueData -Force -ErrorAction SilentlyContinue).$regValue))
							{
								OutputLog -Text ('Key value set successfully')
								$valueSet = $True
							}
							Else { OutputLog -Text ('Key value set unsuccessfully') -Severity 'Error' }
						}
						Catch
						{
							OutputLog -Text ('Error creating value under registry key') -Severity 'Error'
							OutputLog -Text ($_.Exception.Message) -Severity 'Error'
						}
					}
				}
			}
		}
		Else { OutputLog -Text ('"' + $regKey + '" does not have a valid tree') -Severity 'Warning' }
	}
	
	Return (($keySet) -And ($valueSet))
}

Function DeleteRegKey([String]$regKey, [String]$regValue)
{
	[Bool]$wasDeleted = $False
	
	If (!(ObjectIsValid -Object $regKey)) { OutputLog -Text ('$RegKey is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $regValue)) { OutputLog -Text ('$RegValue is not valid') -Severity 'Error' }
	Else
	{
		$regKey = GetValidRegTree -RegKey $regKey
		If ($regKey -NE $DefaultString)
		{
			If (RegKeyExists -RegKey $regKey -RegValue $regValue)
			{
				[String]$deleteType = 'Key'
				If ($regValue -NE $DefaultString) { $deleteType = 'Key value' }
				
				OutputLog -Text ('Deleting registry ' + $deleteType.ToLower())
				Try
				{
					If ($regValue -EQ $DefaultString)
					{
						Remove-Item -Path $regKey -Recurse -Force -ErrorAction SilentlyContinue
					}
					Else
					{
						Remove-ItemProperty -Path $regKey -Name $regValue -Force -ErrorAction SilentlyContinue
					}
					
					If (!(RegKeyExists -RegKey $regKey -RegValue $regValue))
					{
						OutputLog -Text ($deleteType + ' deletion successful')
						$wasDeleted = $True
					}
					Else { OutputLog -Text ($deleteType + ' deletion unsuccessful') -Severity 'Error' }
				}
				Catch
				{
					OutputLog -Text ('Error deleting ' + $deleteType.ToLower()) -Severity 'Error'
					OutputLog -Text ($_.Exception.Message) -Severity 'Error'
				}
			}
		}
		Else { OutputLog -Text ('"' + $regKey + '" does not have a valid tree') -Severity 'Warning' }
	}
	
	Return $wasDeleted
}

Function ChangeRegHiveState([String]$state, [String]$hiveName, [String]$path = $DefaultString)
{
	OutputLog -Text ('[FUNCTION]>  ChangeRegHiveState  [STATE]>  Start')
	
	[Bool]$stateChanged = $False
	
	If (!(ObjectIsValid -Object $state)) { OutputLog -Text ('$State is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $hiveName)) { OutputLog -Text ('$HiveName is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	Else
	{
		$hiveName = GetValidRegTree -RegKey $hiveName
		[String]$hiveLoadName = GetValidRegTree -RegKey $hiveName -WithoutColon $True
		Switch ($state.ToUpper())
		{
			'LOAD'
			{
				If ($path -NE $DefaultString)
				{
					OutputLog -Text ('Checking if "' + $path + '" exists')
					If (ObjectExists -Path $path)
					{
						OutputLog -Text ('"' + $path + '" exists')
						
						OutputLog -Text ('Checking if "' + $hiveName + '" already exists')
						If (!(RegKeyExists -RegKey $hiveName))
						{
							OutputLog -Text ('Loading "' + $path + '" hive file to "' + $hiveLoadName + '"')
							If ((RunProcess -FilePath ('reg.exe') -ArgList ('load "' + $hiveLoadName + '" "' + $path + '"') -Attempts 1) -EQ 0)
							{
								OutputLog -Text ('Registry hive loaded successfully')
								$stateChanged = $True
							}
							Else { OutputLog -Text ('Error loading registry hive') -Severity 'Error' }
						}
						Else { OutputLog -Text ('"' + $hiveName + '" already exists, unable to load hive') -Severity 'Error' }
					}
				}
				Else { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
			}
			'UNLOAD'
			{
				OutputLog -Text ('Checking if "' + $hiveName + '" exists')
				If (RegKeyExists -RegKey $hiveName)
				{
					OutputLog -Text ('"' + $hiveName + '" exists')
					
					OutputLog -Text ('Closing "' + $hiveName + '" hive and running a garbage collection')
					Try
					{
						(Get-ChildItem -Path $hiveName -ErrorAction SilentlyContinue).Close()
						[GC]::Collect()
						Start-Sleep -S 10 -ErrorAction SilentlyContinue
						
						OutputLog -Text ('Unloading "' + $hiveLoadName + '" hive')
						If ((RunProcess -FilePath ('reg.exe') -ArgList ('unload "' + $hiveLoadName + '"') -Attempts 1) -EQ 0)
						{
							OutputLog -Text ('Registry hive unloaded successfully')
							$stateChanged = $True
						}
						Else { OutputLog -Text ('Error unloading registry hive') -Severity 'Error' }
					}
					Catch
					{
						OutputLog -Text ('Error closing mounted hive and running garbage collection') -Severity 'Error'
						OutputLog -Text ($_.Exception.Message) -Severity 'Error'
					}
				}
				Else { OutputLog -Text ('Unable to unmount hive') -Severity 'Error' }
			}
			Default { OutputLog -Text ('"' + $state + '" is an invalid hive state') -Severity 'Warning' }
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  ChangeRegHiveState  [STATE]>  End')
	
	Return $stateChanged
}

Function CreateShellShortcut([String]$path, [String]$target, [String]$arguments = [String]::Empty, [String]$icon, [String]$description = [String]::Empty, [String]$workingDir)
{
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $target)) { OutputLog -Text ('$Target is not valid') -Severity 'Error' }
	ElseIf ($arguments -EQ $Null) { OutputLog -Text ('$Arguments is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $icon)) { OutputLog -Text ('$Icon is not valid') -Severity 'Error' }
	ElseIf ($description -EQ $Null) { OutputLog -Text ('$Description is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $workingDir)) { OutputLog -Text ('$WorkingDir is not valid') -Severity 'Error' }
	Else
	{
		OutputLog -Text ('Creating shortcut to "' + $target + '" at location "' + $path + '"')
		Try
		{
			[Object]$objShell = New-Object -ComObject 'WScript.Shell' -ErrorAction SilentlyContinue
			[Object]$shortcut = $objShell.CreateShortcut($path)
			$shortcut.TargetPath = $target
			$shortcut.Arguments = $arguments
			$shortcut.IconLocation = $icon
			$shortcut.Description = $description
			$shortcut.WorkingDirectory = $workingDir
			$shortcut.Save()
		}
		Catch
		{
			OutputLog -Text ('Error creating shortcut') -Severity 'Error'
			OutputLog -Text ($_.Exception.Message) -Severity 'Error'
		}
	}
}

Function GetProductVersion([String]$path)
{
	[String]$productVersion = $DefaultString
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	Else
	{
		If (ObjectExists -Path $path)
		{
			If ($path.EndsWith('.msi'))
			{
				[Object]$winInstaller = New-Object -ComObject 'WindowsInstaller.Installer' -ErrorAction SilentlyContinue
				Try
				{
					$msiDatabase = $winInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $winInstaller, @($path, 0))
					$query = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
					$view = $msiDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $msiDatabase, ($query))
					$view.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $view, $Null)
					$record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $Null, $view, $Null)
					$productVersion = $record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 1)

					$msiDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $Null, $msiDatabase, $Null)
					$view.GetType().InvokeMember("Close", "InvokeMethod", $Null, $view, $Null)           
					$msiDatabase = $Null
					$view = $Null
				}
				Catch
				{
					OutputLog -Text ('Error encountered while checking ProductVersion for "' + $path + '"') -Severity 'Error'
					OutputLog -Text ('Error: ' + $_.Exception.Message) -Severity 'Error'
				}
				
				[System.Runtime.Interopservices.Marshal]::ReleaseComObject($winInstaller) | Out-Null
				[System.GC]::Collect()
			}
			ElseIf ($path.EndsWith('.exe'))
			{
				$productVersion = (Get-Item -Path $path -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
			}
			Else { OutputLog -Text ('"' + $path + '" is not supported for checking ProductVersion') -Severity 'Warning' }
		}
	}

	Return $productVersion
}

Function StripWhitespace([String]$string)
{
	[String]$str = [String]::Empty
	
	If (!(ObjectIsValid -Object $string)) { OutputLog -Text ('$String is not valid') -Severity 'Error' }
	Else
	{
		For ([UInt32]$i = 0; $i -NE $string.Length; ++$i)
		{
			If ($string[$i] -NE ' ') { $str += $string[$i] }
		}
	}
	
	Return $str
}

Function FindFileWithStartStringAndExtension([String]$path, [String]$startString, [String]$extension)
{
	[String]$found = $DefaultString
	
	If (!(ObjectIsValid -Object $path)) { OutputLog -Text ('$Path is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $startString)) { OutputLog -Text ('$StartString is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $extension)) { OutputLog -Text ('$Extension is not valid') -Severity 'Error' }
	Else
	{
		[String]$fullPath = $path
		If (!$fullPath.EndsWith('\')) { $fullPath += '\' }
		If (!$startString.StartsWith('*')) { $startString = ('*' + $startString) }
		If (!$startString.EndsWith('*')) { $startString += '*' }
		$fullPath += ($startString + '.' + $extension)
		
		[Object[]]$files = Get-Item -Path $fullPath -ErrorAction SilentlyContinue
		If (!(ObjectIsValid -Object $files)) { OutputLog -Text ('Unable to find any files of type "' + $extension + '" starting with "' + $startString + '" at the path "' + $path + '"') -Severity 'Warning' }
		Else { $found = [String]$files[0].Name }
	}
	
	Return $found
}

Function TerminateProcesses([String[]]$processArray)
{
	OutputLog -Text ('[FUNCTION]>  TerminateProcesses  [STATE]>  Start')
	
	If (!(ObjectIsValid -Object $processArray)) { OutputLog -Text ('$ProcessArray is not valid') -Severity 'Error' }
	Else
	{
		If ($processArray.Length -EQ 1) { OutputLog -Text ('Terminating process') }
		Else { OutputLog -Text ('Terminating processes') }
		For ([UInt32]$i = 0; $i -NE $processArray.Length; ++$i)
		{
			OutputLog -Text ('Terminating all instances of "' + $processArray[$i] + '"')
			[UInt32]$count = 0
			While (Stop-Process -ProcessName $processArray[$i] -PassThru -Force -ErrorAction SilentlyContinue)
			{
				Start-Sleep -M 500 -ErrorAction SilentlyContinue
				++$count
			}
			If ($count -EQ 1) { OutputLog -Text ($count.ToString() + ' instance of "' + $processArray[$i] + '" terminated') }
			Else { OutputLog -Text ($count.ToString() + ' instances of "' + $processArray[$i] + '" terminated') }
		}
		If ($processArray.Length -EQ 1) { OutputLog -Text ('Process terminated') }
		Else { OutputLog -Text ('Processes terminated') }
	}
	
	OutputLog -Text ('[FUNCTION]>  TerminateProcesses  [STATE]>  End')
}

Function RunProcess([String]$filePath, [String]$argList, [UInt32]$attempts = 5, [UInt32]$waitTime = 10, [Bool]$waitForExit = $True)
{
	OutputLog -Text ('[FUNCTION]>  RunProcess  [STATE]>  Start')
	
	[Int32]$retCode = -1
	
	If (!(ObjectIsValid -Object $filePath)) { OutputLog -Text ('$FilePath is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $argList)) { OutputLog -Text ('$ArgList is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $attempts)) { OutputLog -Text ('$Attempts is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $waitTime)) { OutputLog -Text ('$WaitTime is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $waitForExit)) { OutputLog -Text ('$WaitForExit is not valid') -Severity 'Error' }
	Else
	{
		If ($attempts -EQ 0) { $attempts = 1 }
		OutputLog -Text ('Running process')
		OutputLog -Text ('Process: ' + $filePath)
		OutputLog -Text ('Argument List: ' + $argList)
		OutputLog -Text ('Attempts: ' + $attempts)
		OutputLog -Text ('Wait time between attempts: ' + $waitTime + ' seconds')
		OutputLog -Text ('Wait for exit: ' + $waitForExit.ToString())
		Try
		{
			For ([UInt32]$i = 1; $i -LE $attempts; ++$i)
			{
				OutputLog -Text ('Process attempt ' + $i + ' of ' + $attempts)
				[System.Diagnostics.ProcessStartInfo]$processInfo = New-Object System.Diagnostics.ProcessStartInfo -ErrorAction SilentlyContinue
				$processInfo.FileName = $filePath
				$processInfo.Arguments = $argList
				$processInfo.WindowStyle = 'Hidden'
				$processInfo.UseShellExecute = $False
				$processInfo.RedirectStandardError = $True
				$processInfo.RedirectStandardOutput = $True
				[System.Diagnostics.Process]$process = New-Object System.Diagnostics.Process -ErrorAction SilentlyContinue
				$process.StartInfo = $processInfo
				OutputLog -Text ('Starting process: ' + $processInfo.FileName + ' ' + $argList)
				$process.Start() | Out-Null
				[String]$stdOut = $process.StandardOutput.ReadToEnd()
				[String]$stdErr = $process.StandardError.ReadToEnd()
				If ($waitForExit) { $process.WaitForExit() }
				$retCode = $process.ExitCode
				OutputLog -Text ('Return code from process: ' + $retCode)
				If ($retCode -EQ 1618)
				{
					If ($i -EQ $attempts)
					{
						OutputLog -Text ('Maximum process attempts reached with another installation still in progress') -Severity 'Error'
						Break
					}
					Else
					{
						OutputLog -Text ('Another installation is in progress, waiting ' + $waitTime.ToString() + ' seconds and trying again')
						Start-Sleep -S $waitTime -ErrorAction SilentlyContinue
						TerminateProcesses
					}
				}
				Else { Break }
			}
		}
		Catch
		{
			OutputLog -Text ('Error while trying to start process "' + $filepath + '" with argument list "' + $argList + '"') -Severity 'Error'
			OutputLog -Text ($_.Exception.Message) -Severity 'Error'
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  RunProcess  [STATE]>  End')
	
	Return $retCode
}

Function WaitForProcessToExit([String]$name, [UInt32]$maxWaitTime)
{
	OutputLog -Text ('[FUNCTION]>  WaitForProcessToExit  [STATE]>  Start')
	
	If (!(ObjectIsValid -Object $name)) { OutputLog -Text ('$Name is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $maxWaitTime)) { OutputLog -Text ('$MaxWaitTime is not valid') -Severity 'Error' }
	Else
	{
		$name = [IO.Path]::GetFileNameWithoutExtension($name)
		
		Try
		{
			OutputLog -Text ('Checking if process "' + $name + '" exists')
			[Int32]$processID = (Get-Process -Name $name -ErrorAction SilentlyContinue).ID
			If ($Null -NE $processID)
			{
				OutputLog -Text ('"' + $name + '" exists')
				
				[UInt32]$minCounter = 0
				If ($maxWaitTime -EQ 1) { OutputLog -Text ('Looping until "' + $name + '" has exited or ' + $maxWaitTime.ToString() + ' minute reached') }
				Else { OutputLog -Text ('Looping until "' + $name + '" has exited or ' + $maxWaitTime.ToString() + ' minutes reached') }
				
				[Object]$sw = [System.Diagnostics.Stopwatch]::StartNew()
				While ($Null -NE (Get-Process -ID $processID -ErrorAction SilentlyContinue))
				{
					If ($sw.Elapsed.Minutes -GT $minCounter)
					{
						$minCounter = $sw.Elapsed.Minutes
						If ($minCounter -EQ 1) { OutputLog -Text ('Elapsed Time: ' + $minCounter + ' minute') }
						Else { OutputLog -Text ('Elapsed Time: ' + $minCounter + ' minutes') }
					}
					
					If ($minCounter -GE $maxWaitTime)
					{
						OutputLog -Text ('Process maximum wait time reached, stopping loop') -Severity 'Warning'
						Break
					}
				}
				
				$sw.Stop()
				
				If (($sw.Elapsed.Minutes -EQ 1) -And ($sw.Elapsed.Seconds -EQ 1))
				{
					OutputLog -Text ('Total elapsed time: ' + $sw.Elapsed.Minutes + ' minute, ' + $sw.Elapsed.Seconds + ' second')
				}
				ElseIf ($sw.Elapsed.Minutes -EQ 1)
				{
					OutputLog -Text ('Total elapsed time: ' + $sw.Elapsed.Minutes + ' minute, ' + $sw.Elapsed.Seconds + ' seconds')
				}
				ElseIf ($sw.Elapsed.Seconds -EQ 1)
				{
					OutputLog -Text ('Total elapsed time: ' + $sw.Elapsed.Minutes + ' minutes, ' + $sw.Elapsed.Seconds + ' second')
				}
				ElseIf (($sw.Elapsed.Minutes -EQ 0) -And ($sw.Elapsed.Seconds -EQ 1))
				{
					OutputLog -Text ('Total elapsed time: ' + $sw.Elapsed.Seconds + ' second')
				}
				ElseIf ($sw.Elapsed.Minutes -EQ 0)
				{
					OutputLog -Text ('Total elapsed time: ' + $sw.Elapsed.Seconds + ' seconds')
				}
				Else { OutputLog -Text ('Total elapsed time: ' + $sw.Elapsed.Minutes + ' minutes, ' + $sw.Elapsed.Seconds + ' seconds') }
			}
			Else { OutputLog -Text ('Process "' + $name + '" does not exist') }
		}
		Catch
		{
			OutputLog -Text ('Error waiting for process to exit') -Severity 'Error'
			OutputLog -Text ($_.Exception.Message) -Severity 'Error'
		}
	}
	
	OutputLog -Text ('[FUNCTION]>  WaitForProcessToExit  [STATE]>  End')
}

Function ExitWithError([Int32]$errorCode = -1, [Bool]$shouldExitHere = $False)
{
	OutputLog -Text ('[FUNCTION]>  ExitWithError  [STATE]>  Start')
	
	If (!(ObjectIsValid -Object $ExitWithErrorText)) { OutputLog -Text ('$ExitWithErrorText is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $ExitWithErrorCodes)) { OutputLog -Text ('$ExitWithErrorCodes is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $errorCode)) { OutputLog -Text ('$ErrorCode is not valid') -Severity 'Error' }
	ElseIf (!(ObjectIsValid -Object $shouldExitHere)) { OutputLog -Text ('$ShouldExitHere is not valid') -Severity 'Error' }
	Else
	{
		If (($ExitWithErrorText.Count -GT 0) -And ($ExitWithErrorCodes.Count -GT 0) -And ($ExitWithErrorText.Count -EQ $ExitWithErrorCodes.Count))
		{
			If ($ExitWithErrorText.Count -EQ 1) { OutputLog -Text ('Exiting script due to critical runtime error') -Severity 'Error' }
			Else { OutputLog -Text ('Exiting script due to critical runtime errors') -Severity 'Error' }
			For ([UInt32]$i = 0; $i -NE $ExitWithErrorText.Count; ++$i)
			{
				OutputLog -Text ('->  Error Text:  ' + $ExitWithErrorText[$i] + ',  Error Code:  ' + $ExitWithErrorCodes[$i]) -Severity 'Error'
			}
		}
		
		$Global:ExitCode = $errorCode
		
		OutputLog -Text ('[FUNCTION]>  ExitWithError  [STATE]>  End')
	
		If ($shouldExitHere -EQ $True) { Exit $ExitCode }
	}
}

#endregion **** Commonly Used Functions


#region **** Module Main Call ****

ModuleMain

#endregion **** Module Main Call ****