#region **** Script Information ****

<#

File name:		validate-users-and-email-licenses.ps1
Description:	Validates users using RSAT AD module and optionally emails license details.
                This builds on get-all-users-from-ad.ps1 and adds email functionality.

Author:			Spencer Prentiss

Versions:
1.0		Base script with core functionality

#>

#endregion


#region **** Global "Const" Variables

[String]$Global:SCRIPT_ROOT = Split-Path -Parent $MyInvocation.MyCommand.Definition
[String]$Global:CUSTOMER_CSV = $SCRIPT_ROOT + '\CustomerExport.csv'
[String]$Global:LICENSE_CSV = $SCRIPT_ROOT + '\ExportLicenses.csv'
[String]$Global:USERS_CSV = $SCRIPT_ROOT + '\Users.csv'
[String]$Global:AD_SERVER = 'AD.Server.com'                     #   Update with AD server to query
[String]$Global:EMAIL_DOMAIN = 'Email.Domain.com'               #   Update with AD domain email postfix
[String]$Global:DOMAIN_SMTP_SERVER = 'Domain.SMTP.Server.com'   #   Server used for SMTP when emailing licenses
[String]$Global:NA = 'N/A'

[Boolean]$Global:VALIDATE_WITH_LOCAL_USERS_FILE = $True     #   Set to $True to use Users.csv from ActiveDirectory rather than check every user based on their email
[Boolean]$Global:OUTPUT_VALIDATED_RECORDS = $True           #	Set to $True to output validation to "CustomerExport_Validated.csv", $False to suppress output -- This is good for fixing invalid user records
[Boolean]$Global:HALT_ON_INVALID_RECORDS = $True            #	Set to $True to halt the script if any invalid records are found (ideally this is always $True), $False to continue with invalid records (not ideal!)
[Boolean]$Global:EMAIL_USERS = $True                        #	Set to $True to email user their license details (you will still be prompted to respond with Yes or No to ensure you want to send potentially thousands of emails, $False to not prompt to email license details

[String[]]$Global:CUSTOMER_EXPORT_PROPS = @(
    'CustomerID',
    'CompanyName',
    'FirstName',
    'LastName',
    'Country',
    'EMail'
)
[String[]]$Global:LICENSE_EXPORT_PROPS = @(
    'CustomerID',
    'CompanyName',
    'FirstName',
    'LastName',
    'Country',
    'EMail',
    'Enabled',
    'LicenseID',
    'ActivationPassword',
    'Status',
    'ProductName',
    'OptionName',
    'UnlocksLeft'
)
[String[]]$Global:AD_PROPS = @(
    'SamAccountName',
    'DistinguishedName',
    'GivenName',
    'Surname',
    'EmailAddress',
    'Enabled',
    'Title',
    'HROrganizationDesc',
    'Office',
    'City',
    'State',
    'co'
)

#endregion


#region **** Global Variables

[Hashtable]$Global:ExportedCustomerRecords = @{ }
[Hashtable]$Global:ADUsers = @{ }
[Hashtable]$Global:ValidUserObjects = @{ }
[Hashtable]$Global:UsersWithLicenses = @{ }
[Boolean]$Global:ShouldContinue = $True

#endregion



#region **** Main ****

Function Main {
    Clear-Host

    [String[]]$functionNames = @(
        'PrerequisitesPassed',
        'CheckCustomerRecordsAreValid',
        'CheckUserEmailsAreValid',
        'QueryUsersWithLicenses',
        'EmailLicensesToUsers'
    )

    For ([UInt32]$i = 0; $i -NE $functionNames.Length; ++$i) {
        &($functionNames[$i])
        If (!($Global:ShouldContinue)) { Return 2 }
    }

    Return 0
}

#endregion


#region **** Primary Functions ****

Function PrerequisitesPassed() {
    Set-Location -Path $Global:SCRIPT_ROOT
    Write-Host ('Working directory:  ' + $Global:SCRIPT_ROOT)
    Write-Host ([String]::Empty)

    If (!$Global:VALIDATE_WITH_LOCAL_USERS_FILE) {
        $Global:ShouldContinue = IsModuleInstalled -Module 'ActiveDirectory'

        If ($Global:ShouldContinue) {
            $Global:ShouldContinue = IsOnDomainNetwork -ServerName $Global:AD_SERVER
        }
        If ($Global:ShouldContinue -And (!(Test-Path -Path $Global:USERS_CSV -ErrorAction SilentlyContinue))) {
            $Host.UI.WriteErrorLine('[ERROR]  "' + $Global:USERS_CSV + '" not found, script halting...')
            $Global:ShouldContinue = $False
        }
    }

    If ($Global:ShouldContinue -And (!(Test-Path -Path $Global:CUSTOMER_CSV -ErrorAction SilentlyContinue))) {
        $Host.UI.WriteErrorLine('[ERROR]  "' + $Global:CUSTOMER_CSV + '" not found, script halting...')
        $Global:ShouldContinue = $False
    }
    If ($Global:ShouldContinue -And (!(Test-Path -Path $Global:LICENSE_CSV -ErrorAction SilentlyContinue))) {
        $Host.UI.WriteErrorLine('[ERROR]  "' + $Global:LICENSE_CSV + '" not found, script halting...')
        $Global:ShouldContinue = $False
    }
}

Function CheckCustomerRecordsAreValid() {
    Write-Host ([String]::Empty)
    Write-Host ('Querying records in "' + (Split-Path -Leaf $Global:CUSTOMER_CSV) + '", please wait...')

    [System.Collections.Generic.List[Hashtable]]$invalids = [System.Collections.Generic.List[Hashtable]]::new()

    $Null = Import-Csv -Path $Global:CUSTOMER_CSV -Delimiter ',' -Encoding UTF8 |
    Where-Object { $_.Enabled -EQ $True } |
    Select-Object $Global:CUSTOMER_EXPORT_PROPS | ForEach-Object {
        [Hashtable]$hash = @{ CustomerID = $_.CustomerID }
        [Boolean]$valid = $True

        [String]$firstName = GetNonNullString -ObjectToCheck $_.FirstName
        If ($firstName -EQ [String]::Empty) {
            $hash.Add('FirstName', $firstName)
            $valid = $False
        }

        [String]$lastName = GetNonNullString -ObjectToCheck $_.LastName
        If ($lastName -EQ [String]::Empty) {
            $hash.Add('LastName', $lastName)
            $valid = $False
        }

        [String]$country = GetNonNullString -ObjectToCheck $_.Country
        If (($country -EQ [String]::Empty) -Or (($country -NE 'UNITED STATES') -And ($country -NE 'CANADA'))) {
            $hash.Add('Country', $country)
            $valid = $False
        }

        [String]$email = GetNonNullString -ObjectToCheck $_.EMail
        If (($email -EQ [String]::Empty) -Or (!($email.ToLower().EndsWith('@' + $Global:EMAIL_DOMAIN)))) {
            $hash.Add('EMail', $email)
            $valid = $False
        }
        ElseIf ($Global:ExportedCustomerRecords.ContainsKey($email)) {
            $hash.Add('EMail', ($email + ' is a duplicate'))
            $valid = $False
        }

        If ($valid) {
            $Global:ExportedCustomerRecords.Add($_.EMail, (PopulateADUserObject -Record $_))
        }
        Else {
            $invalids.Add($hash)
        }
    }

    Write-Host ($Global:ExportedCustomerRecords.Count.ToString() + ' valid records found')

    If ($invalids.Count -EQ 0) { Write-Host ($invalids.Count.ToString() + ' invalid records found') }
    Else {
        $Host.UI.WriteErrorLine($invalids.Count.ToString() + ' invalid records found')
        For ([Int32]$i = 0; $i -NE $invalids.Count; ++$i) {
            Write-Host ([String]::Empty)
            $Host.UI.WriteErrorLine(($i + 1).ToString() + ': ' + $invalids[$i].CustomerID)
            ForEach ($key in $invalids[$i].Keys) {
                If ($key -NE 'CustomerID') {
                    $Host.UI.WriteErrorLine('    -> ' + $key + ': ' + $invalids[$i].$key)
                }
            }
        }
        Write-Host ([String]::Empty)
        $Host.UI.WriteErrorLine('These records need to be corrected on the server before proceeding!')
        $Host.UI.WriteErrorLine('Once these records are corrected, export customer records once again and re-run this.')
        $Global:ShouldContinue = $False
    }
}

Function CheckUserEmailsAreValid() {
    Write-Host ([String]::Empty)

    [System.Collections.Generic.List[Hashtable]]$invalids = [System.Collections.Generic.List[Hashtable]]::new()

    If ($Global:VALIDATE_WITH_LOCAL_USERS_FILE) {
        #   Get all users from AD with desired properties from large AD CSV export
        Write-Host ('Validating user emails across "' + (Split-Path -Leaf $Global:USERS_CSV) + '" exported from Active Directory...')

        Import-CSV -Path $Global:USERS_CSV -Delimiter ',' -Encoding UTF8 |
        Select-Object $Global:AD_PROPS |
        Where-Object { (GetNonNullString -ObjectToCheck $_.EmailAddress) -NE '' } |
        ForEach-Object { Try { $Global:ADUsers.Add($_.EmailAddress, $_) } Catch { Write-Host $_.EMail } }

        #   Get all user records with valid email addresses from large AD CSV export
        $Global:ExportedCustomerRecords.Keys | ForEach-Object {
            Try {
                If ($Global:ADUsers.ContainsKey($_)) {
                    $Global:ValidUserObjects.Add($_, (PopulateADUserObject -Record $Global:ExportedCustomerRecords.$_ -OldRecord $Global:ADUsers.$_))
                }
                Else {
                    $invalids.Add(@{
                            CustomerID   = $Global:ExportedCustomerRecords.$_.CustomerID;
                            EmailAddress = $Global:ExportedCustomerRecords.$_.EMail;
                        })
                }
            }
            Catch { }
        }
    }
    Else {
        #   Get all user records with valid email addresses from AD by querying each email address individually
        Write-Host ('Validating user emails across Active Directory...')

        $Global:ExportedCustomerRecords.Keys | ForEach-Object {
            [PSCustomObject]$user = PopulateADUserObject -Record (GetADUserFromEmail -Record $Global:ExportedCustomerRecords.$_)
            If ($user.EmailAddress -NE $Global:NA) { $Global:ValidUserObjects.Add($user.EMail, $user) }
            Else { $invalids.Add($user) }
        }
    }

    Write-Host ($Global:ValidUserObjects.Count.ToString() + ' valid emails found')

    If ($invalids.Count -EQ 0) { Write-Host ($invalids.Count.ToString() + ' invalid emails found') }
    Else {
        $Host.UI.WriteErrorLine($invalids.Count.ToString() + ' invalid emails found')
        Write-Host ([String]::Empty)
        For ([Int32]$i = 0; $i -NE $invalids.Count; ++$i) {
            $Host.UI.WriteErrorLine(($i + 1).ToString() + ':' + (GetAlignSpacing -CurrentNum $i -MaxNum $invalids.Count) + 'CustomerID: ' + $invalids[$i].CustomerID + '  |  EmailAddress: ' + $invalids[$i].EmailAddress)
        }
        Write-Host ([String]::Empty)
        $Host.UI.WriteErrorLine('These emails need to be corrected on the server before proceeding!')
        $Host.UI.WriteErrorLine('Once these emails are corrected, export customer records once again and re-run this.')
        $Global:ShouldContinue = $False
    }
}

Function QueryUsersWithLicenses() {
    Write-Host ([String]::Empty)
    Write-Host ('Querying users with licenses in "' + (Split-Path -Leaf $Global:LICENSE_CSV) + '", please wait...')

    $Null = Import-Csv -Path $Global:LICENSE_CSV -Delimiter ',' -Encoding UTF8 |
    Select-Object $Global:LICENSE_EXPORT_PROPS | ForEach-Object {
        If (((GetNonNullString -ObjectToCheck $_.EMail) -NE [String]::Empty) -And ($Global:ValidUserObjects.ContainsKey($_.EMail)) -And ($_.Enabled -EQ $True) -And ($_.Status.StartsWith('OK'))) {
            $Global:ValidUserObjects.($_.EMail).Licenses.Add($_.LicenseID, $_)
        }
    }

    $Global:ValidUserObjects.Keys |
    Where-Object { $Global:ValidUserObjects.$_.Licenses.Keys.Count -GT 0 } |
    ForEach-Object { $Global:UsersWithLicenses.Add($_, $Global:ValidUserObjects.$_) }

    [Int32]$usersWithoutLicenses = $Global:ValidUserObjects.Count - $Global:UsersWithLicenses.Count
    Write-Host ($Global:UsersWithLicenses.Count.ToString() + ' users with licenses')
    If ($usersWithoutLicenses -EQ 0) {
        Write-Host ($usersWithoutLicenses.ToString() + ' users without licenses')
    }
    Else {
        $Host.UI.WriteWarningLine($usersWithoutLicenses.ToString() + ' users without licenses')
        $Host.UI.WriteWarningLine('You may want to disable any user customer records on the server that do not have licenses attached')
    }
    Write-Host ([String]::Empty)

    [Hashtable]$warnings = @{ }

    $Global:UsersWithLicenses.Keys | ForEach-Object {
        [Hashtable]$warn = @{ }
        [Boolean]$shouldWarn = $False
        If ($Global:UsersWithLicenses.$_.Licenses.Count -GT 1) {
            $warn.Add('LicenseCount', $Global:UsersWithLicenses.$_.Licenses.Count)
            $shouldWarn = $True
        }
        ForEach ($license in $Global:UsersWithLicenses.$_.Licenses.Keys) {
            If (!($warn.ContainsKey('Licenses'))) {
                $warn.Add('Licenses', @{ })
            }
            $warn.Licenses.Add($license, @{ LicenseID = $license })
            If (([Int32]($Global:UsersWithLicenses.$_.Licenses.$license.UnlocksLeft)) -NE 3) {
                $warn.Licenses.$license.Add('UnlocksLeft', $Global:UsersWithLicenses.$_.Licenses.$license.UnlocksLeft)
                $shouldWarn = $True
            }
        }

        If ($shouldWarn) { $warnings.Add($_, $warn) }
    }

    $warnings.Keys | ForEach-Object {
        $Host.UI.WriteWarningLine('Warnings for ' + $_ + ':')
        If ($warnings.$_.LicenseCount -GT 1) {
            $Host.UI.WriteWarningLine('-> License Count: ' + $warnings.$_.LicenseCount.ToString())
        }
        ForEach ($license in $warnings.$_.Licenses.Keys) {
            $Host.UI.WriteWarningLine('-> LicenseID: ' + $license)
            If ($warnings.$_.Licenses.$license.ContainsKey('UnlocksLeft')) {
                $Host.UI.WriteWarningLine('   -> Unlocks Left: ' + $warnings.$_.Licenses.$license.UnlocksLeft)
            }
        }
        Write-Host ('')
    }

    [String]$response = Read-Host -Prompt ('Acknowledge warnings and continue (Y/N)')
    While (($response.ToLower() -NE 'y') -And ($response.ToLower() -NE 'yes') -And ($response.ToLower() -NE 'n') -And ($response.ToLower() -NE 'no')) {
        Write-Host ('"' + $response + '" is not a valid response!')
        $response = Read-Host -Prompt ('Acknowledge warnings and continue (Y/N)')
    }
    If (($response.ToLower() -NE 'y') -And ($response.ToLower() -NE 'yes')) {
        $Global:ShouldContinue = $False
    }
}

Function EmailLicensesToUsers() {
    If ($Global:EMAIL_USERS) {
        Write-Host ([String]::Empty)
        [String]$response = Read-Host -Prompt ('Email ' + $Global:UsersWithLicenses.Count.ToString() + ' users their license details (Y/N)')
        While (($response.ToLower() -NE 'y') -And ($response.ToLower() -NE 'yes') -And ($response.ToLower() -NE 'n') -And ($response.ToLower() -NE 'no')) {
            Write-Host ('"' + $response + '" is not a valid response!')
            $response = Read-Host -Prompt ('Email ' + $Global:UsersWithLicenses.Count.ToString() + ' users their license details (Y/N)')
        }
        If (($response.ToLower() -EQ 'y') -Or ($response.ToLower() -EQ 'yes')) {
            Write-Host ('Emailing users...')
            $Global:UsersWithLicenses.Keys |
            ForEach-Object {
                [PSCustomObject]$user = $Global:UsersWithLicenses.$_
                [Hashtable]$licenses = $Global:UsersWithLicenses.$_.Licenses
                [String]$postfix = [String]::Empty
                If ($licenses.Count -GT 1) { $postfix = 's' }
                [String]$userString = ($user.FirstName + ' ' + $user.LastName + ' <' + $user.EMail + '>')
                Write-Host ('Emailing ' + $licenses.Count + ' license' + $postfix + ' to "' + $userString + '"...')
                $licenses.Keys | ForEach-Object {
                    [String]$body = GetEmailBodyAsHtml -FirstName $licenses.$_.FirstName -ProductName $licenses.$_.ProductName -OptionName $licenses.$_.OptionName -LicenseID $_ -ActivationPassword $licenses.$_.ActivationPassword -UnlocksLeft $licenses.$_.UnlocksLeft
                    If (([Int32]($licenses.$_.UnlocksLeft)) -GT 3) {
                        $Null = Send-MailMessage -From ('Generic.Company.Email@' + $Global:EMAIL_DOMAIN) -To $userString -Subject 'License Details' -Body $body -BodyAsHtml -SmtpServer $Global:DOMAIN_SMTP_SERVER -WarningAction Ignore
                    }
                }
            }
            Write-Host ('Emails sent')
        }
    }
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

Function GetAlignSpacing([Int32]$currentNum, [Int32]$maxNum) {
    [Int32]$spaceCount = 1
    If ($maxNum -GE 9) { $spaceCount = 2 }
    ElseIf ($maxNum -GE 99) { $spaceCount = 3 }
    ElseIf ($maxNum -GE 999) { $spaceCount = 4 }

    If ($currentNum -GE 9) { $spaceCount -= 1 }
    If ($currentNum -GE 99) { $spaceCount -= 1 }
    If ($currentNum -GE 999) { $spaceCount -= 1 }

    [String]$space = ' '
    For ([Int32]$i = 0; $i -NE $spaceCount; ++$i) { $space += ' ' }

    Return $space
}

Function IsModuleInstalled([String]$module) {
    Write-Host ('Checking if ' + $module + ' is available...')
    If ($Null -EQ (Import-Module -Name $module -ErrorAction SilentlyContinue)) {
        Write-Host ($module + ' may already be available')
    }
    Else {
        Write-Host ($module + ' is available')
    }

    Write-Host ([String]::Empty)
    Write-Host ('Checking if ' + $module + ' is installed...')
    [PSCustomObject[]]$modules = Get-Module | Select-Object Name | Where-Object { $_.Name -EQ $module }
    If ($modules.Length -GT 0) {
        Write-Host ($module + ' module is installed')
        Write-Host ([String]::Empty)
        Return $True
    }

    $Host.UI.WriteErrorLine($module + ' is not installed, script halting')
    Return $False
}

Function IsOnDomainNetwork([String]$serverName, [Boolean]$continueOnError = $True) {
    Write-Host ('Checking if domain network is available...')

    [System.Diagnostics.ProcessStartInfo]$startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = 'ping.exe'
    $startInfo.Arguments = ($serverName + ' -4 -n 1')
    $startInfo.WindowStyle = 'Hidden'
    $startInfo.UseShellExecute = $False
    $startInfo.RedirectStandardError = $True
    $startInfo.RedirectStandardOutput = $True

    [System.Diagnostics.Process]$ping = [System.Diagnostics.Process]::new()
    $ping.StartInfo = $startInfo
    $Null = $ping.Start()
    $Null = $ping.StandardOutput.ReadToEnd()
    $Null = $ping.StandardError.ReadToEnd()
    $ping.WaitForExit()
    [Int32]$pingRetCode = $ping.ExitCode
    If ($pingRetCode -EQ 0) {
        Write-Host ('Domain network available')
        Return $True
    }

    If (!($continueOnError)) {
        $Host.UI.WriteErrorLine('Domain network not available, script halting')
        Return $False
    }

    $Host.UI.WriteWarningLine('Domain network not available')
    Return $True
}

Function PopulateADUserObject([PSCustomObject]$record, [PSCustomObject]$oldRecord = $Null) {
    [PSCustomObject]$newRecord = [PSCustomObject]@{
        CustomerID         = $Global:NA;
        CompanyName        = $Global:NA;
        FirstName          = $Global:NA;
        GivenName          = $Global:NA;
        LastName           = $Global:NA;
        Surname            = $Global:NA;
        EMail              = $Global:NA;
        EmailAddress       = $Global:NA;
        Enabled            = $True;
        SamAccountName     = $Global:NA;
        DistinguishedName  = $Global:NA;
        Title              = $Global:NA;
        HROrganizationDesc = $Global:NA;
        Office             = $Global:NA;
        City               = $Global:NA;
        State              = $Global:NA;
        Country            = $Global:NA;
        co                 = $Global:NA;
        Licenses           = @{ }
    }

    If (ObjectIsValid -Object $oldRecord) {
        $oldRecord.PSObject.Members |
        Where-Object { ($_.MemberType -Like 'NoteProperty') -And ($_.Value -NE $Global:NA) } |
        ForEach-Object { $newRecord.($_.Name) = $_.Value }
    }

    $record.PSObject.Members |
    Where-Object { ($_.MemberType -Like 'NoteProperty') -And ($_.Value -NE $Global:NA) } |
    ForEach-Object { $newRecord.($_.Name) = $_.Value }

    Return $newRecord
}

Function GetADUserFromEmail([PSCustomObject]$record) {
    Try {
        [PSCustomObject[]]$user = Get-ADUser -Filter ('EmailAddress -EQ "' + $record.EMail + '"') -Properties $Global:AD_PROPS | Select-Object $Global:AD_PROPS
        If (($Null -NE $user) -And ($user.Length -GT 0)) {
            $user[0].PSObject.Members | Where-Object MemberType -Like 'NoteProperty' | ForEach-Object { $record.($_.Name) = $_.Value }
        }
    }
    Catch { }

    Return $record
}

Function GetEmailBodyAsHtml([String]$firstName, [String]$productName, [String]$optionName, [String]$licenseID, [String]$activationPassword, [Int32]$unlocksLeft) {
    Return @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<body>
		Populate license details here.
	</body>
</html>
"@
}

#endregion


#region Main Call and Exit

[Int32]$Global:RetCode = Main
Write-Host ([String]::Empty)
Write-Host ([String]::Empty)
If ($Global:RetCode -EQ 0) {
    Write-Host ('Exiting with return code: ' + $Global:RetCode.ToString())
}
Else {
    $Host.UI.WriteErrorLine('Exiting with return code: ' + $Global:RetCode.ToString())
}
Write-Host ([String]::Empty)
Write-Host ([String]::Empty)
Exit $Global:RetCode

#endregion
