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
 "email_address1"
 "email_address2"
 "email_address3"
)

foreach ($user in $UserMailBoxAddress) {

    Start-Transcript -Path "c:\temp\archivestatus$user.txt"

    # Archive Status Controleren 
   $ArchiveResult = Get-EXOMailbox -Identity $user -PropertySets Archive
   $ArchiveResult

   Stop-Transcript

   Start-Transcript -Path "c:\temp\archiveuserlicenses$user.txt"

   # Haal de licenties van de gebruiker op
   $licenses = Get-MsolUser -UserPrincipalName $user | Select-Object -ExpandProperty Licenses
   $licenses | Format-Table -Property AccountSkuId, SkuPartNumber, ConsumedUnits, ActiveUnits

   Stop-Transcript

   Start-Transcript -Path "c:\temp\$user.txt"
   
   echo "Generate Get-Mailbox Pers:"
   Get-Mailbox $user | Format-List pers*
   Get-Mailbox $user | Format-List *quota*,*xloc*
   Get-Mailbox $user -Archive | Format-List *AutoExpandingArchiveEnabled*
   
   echo "Generate Mailbox Statistics:"
   Get-MailboxStatistics $user -Archive | Format-List name,*itemsize*

   Get-MigrationUser $user | fl
   Get-SyncRequest  -Mailbox $user | fl
   Get-SyncRequest  -Mailbox $user | Get-SyncRequestStatistics -IncludeReport -DiagnosticInfo "showtimeslots, verbose"
   Get-SyncRequest  -Mailbox $user | Get-SyncRequestStatistics -IncludeReport -DiagnosticInfo "showtimeslots, verbose" | export-clixml C:\temp\SyncRequest$user.xml
   Get-MoveRequestStatistics $user -IncludeReport | Export-CliXml C:\temp\MoveRequest$user.xml
   Get-MigrationUserStatistics $user -IncludeReport |fl
 
   $stats = Get-MoveRequestStatistics -Identity $user -DiagnosticInfo "verbose,showtimeslots,showtimeline" 
   $stats.DiagnosticInfo
 
   Get-MoveRequestStatistics -Identity $user -DiagnosticInfo 'verbose,showtimeslots' | export-clixml C:\temp\EXO_MoveReqStats$user.xml

   Stop-Transcript

}

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
