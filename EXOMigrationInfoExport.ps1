# Controleer of de ExchangeOnlineManagement-module is geïnstalleerd
$requiredVersion = [System.Version]::new(3, 2, 0)
$moduleName = "ExchangeOnlineManagement"

# Zoek alle versies van de ExchangeOnlineManagement-module
$moduleVersions = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName } | Sort-Object Version -Descending

if ($moduleVersions.Count -gt 0) {
    $latestVersion = $moduleVersions[0].Version

    if ($latestVersion -ge $requiredVersion) {
        Write-Host "ExchangeOnlineManagement module versie $($latestVersion.ToString()) is geïnstalleerd en voldoet aan de minimale vereiste versie van $($requiredVersion.ToString())."
    } else {
        Write-Host "De geïnstalleerde versie van ExchangeOnlineManagement module ($($latestVersion.ToString())) is niet compatibel met de vereiste versie van $($requiredVersion.ToString())."
    }
} else {
    Write-Host "De ExchangeOnlineManagement-module is niet geïnstalleerd."
    if (Check-ModuleInstallationPermission) {
        Install-Module powershellget -force
        Install-Module -Name ExchangeOnlineManagement -force
        Import-Module ExchangeOnlineManagement
    }
}

Connect-ExchangeOnline
Connect-MsolService

$usermailbox1 = "email address1"
$usermailbox2 = "email address2"
$usermailbox3 = "email address3"


# Archive Status Controleren 


$ArchiveResult1 = Get-EXOMailbox -Identity $usermailbox1 -PropertySets Archive
$ArchiveResult2 = Get-EXOMailbox -Identity $usermailbox2 -PropertySets Archive
$ArchiveResult3 = Get-EXOMailbox -Identity $usermailbox3 -PropertySets Archive


# Haal de licenties van de gebruiker op

$licenses1 = Get-MsolUser -UserPrincipalName $usermailbox1 | Select-Object -ExpandProperty Licenses
$licenses2 = Get-MsolUser -UserPrincipalName $usermailbox2 | Select-Object -ExpandProperty Licenses
$licenses3 = Get-MsolUser -UserPrincipalName $usermailbox3 | Select-Object -ExpandProperty Licenses


# Exporteren naar Log FIle

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
    # Controleer de uitvoeringsbeleidsinstelling
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -eq "Restricted" -or $executionPolicy -eq "AllSigned") {
        Write-Host "Je hebt momenteel geen toestemming om modules te installeren vanwege de huidige uitvoeringsbeleidsinstelling ($executionPolicy)."
        return $false
    }

    # Controleer of de gebruiker beheerdersrechten heeft
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentUserPrincipal = New-Object Security.Principal.WindowsPrincipal($currentUser)

    if (!$currentUserPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "Je hebt beheerdersrechten nodig om modules te installeren.Start Powershell als Administrator"
        return $false
    }

    return $true
}

# Roep de functie aan om te controleren of de gebruiker een module mag installeren
$canInstallModule = Check-ModuleInstallationPermission

if ($canInstallModule) {
    Write-Host "Je hebt toestemming om modules te installeren. Voer hier de installatiecode in."
    # Plaats hier de code om de gewenste module te installeren
    # Bijvoorbeeld: Install-Module -Name ModuleNaam -Scope CurrentUser
} else {
    Write-Host "Je hebt momenteel geen toestemming om modules te installeren. Start Powershell als Administrator"
}




