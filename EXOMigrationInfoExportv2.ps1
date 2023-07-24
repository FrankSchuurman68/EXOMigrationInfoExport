# Verify that the ExchangeOnlineManagement module is installed
$requiredVersion = [System.Version]::new(3, 2, 0)
$moduleName = "ExchangeOnlineManagement"

# Locate all versions of the ExchangeOnlineManagement module
$moduleVersions = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName } | Sort-Object Version -Descending

if ($moduleVersions.Count -gt 0) {
    $latestVersion = $moduleVersions[0].Version

    if ($latestVersion -ge $requiredVersion) {
        Write-Host "ExchangeOnlineManagement module version $($latestVersion.ToString()) is installed and meets the minimum required version of $($requiredVersion.ToString())."
    } else {
        Write-Host "The installed version of ExchangeOnlineManagement module ($($latestVersion.ToString())) is not compatible with the required version of $($requiredVersion.ToString())."
    }
} else {
    Write-Host "The ExchangeOnlineManagement module is not installed."
    if (Check-ModuleInstallationPermission) {
        Install-Module powershellget -force
        Install-Module -Name ExchangeOnlineManagement -force
        Import-Module ExchangeOnlineManagement
    }
}

Connect-ExchangeOnline
Connect-MsolService

# Array with email addresses
$UserMailBoxAddress = @(
 "email address1"
 "email address2"
 "email address3"
)

# Archive Status Check


$ArchiveResult1 = Get-EXOMailbox -Identity $usermailbox1 -PropertySets Archive
$ArchiveResult2 = Get-EXOMailbox -Identity $usermailbox2 -PropertySets Archive
$ArchiveResult3 = Get-EXOMailbox -Identity $usermailbox3 -PropertySets Archive

# Get the user's licenses

$licenses1 = Get-MsolUser -UserPrincipalName $usermailbox1 | Select-Object -ExpandProperty Licenses
$licenses2 = Get-MsolUser -UserPrincipalName $usermailbox2 | Select-Object -ExpandProperty Licenses
$licenses3 = Get-MsolUser -UserPrincipalName $usermailbox3 | Select-Object -ExpandProperty Licenses

# Export to Log File

Start-Transcript -Path "c:\temp\archivestatus.txt"

$ArchiveResult1
$ArchiveResult2
$ArchiveResult3

Stop-Transcript

Start-Transcript -Path "c:\temp\archiveuserlicenses.txt"

$usermailbox1
$licenses1 | Format-Table -Property AccountSkuId, SkuPartNumber, ConsumedUnits, ActiveUnits
$usermailbox2
$licenses2 | Format-Table -Property AccountSkuId, SkuPartNumber, ConsumedUnits, ActiveUnits
$usermailbox3
$licenses3 | Format-Table -Property AccountSkuId, SkuPartNumber, ConsumedUnits, ActiveUnits

Stop-Transcript

Start-Transcript -Path "c:\temp\$usermailbox1.txt"
Get-Mailbox $usermailbox1 | Format-List pers*
Get-Mailbox $usermailbox1 | Format-List *quota*,*xloc*
Get-Mailbox $usermailbox1 -Archive | Format-List *AutoExpandingArchiveEnabled*
Get-MailboxStatistics $usermailbox1 -Archive | Format-List name,*itemsize*
Stop-Transcript

Start-Transcript -Path "c:\temp\$usermailbox2.txt"
Get-Mailbox $usermailbox2 | Format-List pers*
Get-Mailbox $usermailbox2 | Format-List *quota*,*xloc*
Get-Mailbox $usermailbox2 -Archive | Format-List *AutoExpandingArchiveEnabled*
Get-MailboxStatistics $usermailbox2 -Archive | Format-List name,*itemsize*
Stop-Transcript

Start-Transcript -Path "c:\temp\$usermailbox3.txt"
Get-Mailbox $usermailbox3 | Format-List pers*
Get-Mailbox $usermailbox3 | Format-List *quota*,*xloc*
Get-Mailbox $usermailbox3 -Archive | Format-List *AutoExpandingArchiveEnabled*
Get-MailboxStatistics $usermailbox3 -Archive | Format-List name,*itemsize*
Stop-Transcript


function Check-ModuleInstallationPermission {
    # Check the execution policy setting
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -eq "Restricted" -or $executionPolicy -eq "AllSigned") {
        Write-Host "Je hebt momenteel geen toestemming om modules te installeren vanwege de huidige uitvoeringsbeleidsinstelling ($executionPolicy)."
        return $false
    }

    # Verify that the user has administrative privileges
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentUserPrincipal = New-Object Security.Principal.WindowsPrincipal($currentUser)

    if (!$currentUserPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "You need administrator privileges to install modules. Run Powershell as Administrator"
        return $false
    }

    return $true
}

# Call the function to check if the user is allowed to install a module
$canInstallModule = Check-ModuleInstallationPermission

if ($canInstallModule) {
    Write-Host "You have permission to install modules. Enter the installation code here."
    # Place the code here to install the desired module
    # Example: Install-Module -Name ModuleNaam -Scope CurrentUser
} else {
    Write-Host "You currently do not have permission to install modules. Run Powershell as Administrator"
}