#region **** Script Information ****

<#

File name:		get-all-users-from-ad.ps1
Description:	Gets all active directory users in targeted OU and all child OUs.
                Outputs all users found to Users.csv file to be used elsewhere.
                Requires optional RSAT AD module installed to access AD cmdlets.

Author:			Spencer Prentiss

Versions:
1.0		Base script with core functionality

#>

#endregion


#region **** Global Variables ****

[System.Collections.Concurrent.ConcurrentDictionary[String, PSCustomObject]]$Global:ConcurrentDict = [System.Collections.Concurrent.ConcurrentDictionary[String, PSCustomObject]]::new()
[String]$Global:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
[String]$Global:UsersCSV = $ScriptRoot + '\Users.csv'
[String]$Global:OURoot = 'OU=User,OU=Accounts,OU=Root,DC=Domain,DC=com'     # Update based on target OU root
[Int32]$Global:MaxThreads = 5
[String[]]$Global:ADProps = @(
    'SamAccountName',
    'DistinguishedName',
    'GivenName',
    'Surname',
    'DisplayName',
    'EmailAddress',
    'Enabled',
    'Title',
    'Department',
    'HROrganization',
    'HROrganizationDesc',
    'StreetAddress',
    'City',
    'State',
    'PostalCode',
    'Office',
    'co',
    'HRLocation',
    'HRRegion'
)

#endregion



#region **** Main ****

Function Main() {
    Clear-Host
    Write-Host ([String]::Empty)
    Write-Host ([String]::Empty)
    Write-Host ('Getting all child OUs under "' + $OURoot + '"')
    [String[]]$ouList = (Get-ADOrganizationalUnit -Filter * -SearchBase $ouRoot -SearchScope Subtree).DistinguishedName
    Write-Host ([String]::Empty)
    Write-Host ('OU count: ' + $ouList.Count.ToString())
    Write-Host ([String]::Empty)

    [System.Diagnostics.Stopwatch]$sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ouList | ForEach-Object -ThrottleLimit $MaxThreads -Parallel {
        $dict = $using:ConcurrentDict
        $props = $using:ADProps
        [PSCustomObject[]]$users = Get-ADUser -Filter { ObjectClass -EQ 'user' } -SearchBase $_ -SearchScope OneLevel -Properties $props |
        Select-Object $props |
        Where-Object { ((ObjectIsValid -ObjectToCheck $_.GivenName) -And (ObjectIsValid -ObjectToCheck $_.Surname) -And (ObjectIsValid -ObjectToCheck $_.EmailAddress)) }
        Write-Host ('User count for "' + $_ + '": ' + $users.Length.ToString())
        $users | ForEach-Object {
            $Null = $dict.TryAdd($_.EmailAddress, $_)
        }

        $Null = $users
    }
    $sw.Stop()

    Write-Host ([String]::Empty)
    Write-Host ('Dictionary Size: ' + $ConcurrentDict.Count.ToString())
    Write-Host ('Time to find all users: ' + $sw.Elapsed.ToString())
    Write-Host ('Outputting all users to "' + $UsersCSV + '"')

    $sw.Reset()
    $sw.Start()
    [System.Collections.Generic.List[PSCustomObject]]$users = [System.Collections.Generic.List[PSCustomObject]]::new()
    $Null = $ConcurrentDict.Keys | Where-Object { $Null = $users.Add($ConcurrentDict.$_) }
    $Null = $users | Export-Csv -Path $UsersCSV -NoTypeInformation -Delimiter ',' -Encoding UTF8 -Force

    $sw.Stop()
    Write-Host ('Time to output CSV: ' + $sw.Elapsed.ToString())

    $Null = $sw

    Return 0
}

#endregion


#region **** Helper Functions ****

Function ObjectIsValid([Object]$object) {
    [Boolean]$valid = $False

    If ($Null -NE $object) {
        $valid = $True
        If (($object.GetType() -EQ [String]) -And ($object -EQ [String]::Empty)) {
            $valid = $False
        }
    }

    Return ($valid)
}

Function GetNonNullString([Object]$objectToCheck) {
    If ($Null -EQ $objectToCheck) {
        Return [String]::Empty
    }

    Return $objectToCheck.ToString()
}

#endregion


#region **** Main Call and Exit ****

[Int32]$retCode = Main
Write-Host ([String]::Empty)
Write-Host ([String]::Empty)
Exit $retCode

#endregion
