<#
.SYNOPSIS
    Exports Microsoft Teams memberships and roles for a selected user to Excel.

.DESCRIPTION
    Microsoft Teams Membership Export is an interactive PowerShell tool for Microsoft 365 / Teams administrators.

    The script:
    - asks for the administrator UPN/e-mail address,
    - securely asks for the administrator password,
    - asks for the target user's UPN/e-mail address,
    - connects to Microsoft Teams,
    - reads all Teams memberships of the target user,
    - determines the role in each Team (Owner, Member, Guest),
    - creates an .xlsx report on the current user's Desktop,
    - opens the generated report automatically.

    If username/password authentication is blocked by MFA or Conditional Access,
    the script falls back to interactive Microsoft authentication and finally
    device-code authentication.

.NOTES
    Project        : Microsoft Teams Membership Export
    Script         : Teams-Membership-Export.ps1
    Author         : Abdelrahman Al-Hnedi
    Version        : 1.0.0
    Last Updated   : 2026-08-27

    Platform       : Windows
    PowerShell     : Windows PowerShell 5.1 or PowerShell 7.2+
    Dependencies   : MicrosoftTeams, ImportExcel

    Security:
    - No password is stored in the script.
    - The password is held only in a PSCredential/SecureString during execution.
    - The password is never exported to Excel.
    - No tenant-specific account is hard-coded.

    Repository     : https://github.com/<YOUR-GITHUB-USERNAME>/Microsoft-Teams-Membership-Export

.EXAMPLE
    PS C:\Scripts> .\Teams-Membership-Export.ps1

    The script opens dialogs for the administrator account and target user
    and creates an Excel report on the Desktop.

.LINK
    https://learn.microsoft.com/powershell/module/teams/connect-microsoftteams
    https://learn.microsoft.com/microsoftteams/teams-powershell-install
#>

#Requires -Version 5.1

$ErrorActionPreference = "Stop"

# ============================================================
# Configuration / master data
# ============================================================

$ScriptName    = "Microsoft Teams Membership Export"
$ScriptVersion = "1.0.0"
$ScriptAuthor  = "Abdelrahman Al-Hnedi"

# ============================================================
# Load GUI components
# ============================================================

try {
    Add-Type -AssemblyName Microsoft.VisualBasic
    Add-Type -AssemblyName System.Windows.Forms
}
catch {
    Write-Host "GUI-Komponenten konnten nicht geladen werden." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

# ============================================================
# Helper functions
# ============================================================

function Show-Info {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [string]$Title = "Information"
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Show-Warning {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [string]$Title = "Warnung"
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

function Show-ErrorMessage {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [string]$Title = "Fehler"
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Disconnect-TeamsSafe {
    try {
        Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        # Disconnect errors are intentionally ignored.
    }
}

function Test-UpnLikeValue {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    return ($Value.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Request-Upn {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [Parameter(Mandatory)]
        [string]$Title,

        [string]$Example = "user@contoso.com"
    )

    while ($true) {
        $Value = [Microsoft.VisualBasic.Interaction]::InputBox(
            "$Prompt`n`nBeispiel:`n$Example",
            $Title,
            ""
        )

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return $null
        }

        $Value = $Value.Trim()

        if (Test-UpnLikeValue -Value $Value) {
            return $Value
        }

        Show-Warning `
            -Message "Die Eingabe sieht nicht wie eine gültige E-Mail-Adresse / UPN aus.`n`nBitte erneut versuchen." `
            -Title "Ungültige Eingabe"
    }
}

function Ensure-Module {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Purpose
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $Answer = [System.Windows.Forms.MessageBox]::Show(
            "Das PowerShell-Modul '$Name' ist nicht installiert.`n`n$Purpose`n`nSoll es jetzt für den aktuellen Benutzer installiert werden?",
            "$Name installieren",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            throw "Das benötigte Modul '$Name' wurde nicht installiert."
        }

        Write-Host "Installiere Modul '$Name'..." -ForegroundColor Yellow

        Install-Module `
            -Name $Name `
            -Scope CurrentUser `
            -Force `
            -AllowClobber `
            -ErrorAction Stop
    }

    Import-Module $Name -ErrorAction Stop
}

# ============================================================
# 1. Check PowerShell / platform
# ============================================================

if ($PSVersionTable.PSVersion -lt [Version]"5.1") {
    Show-ErrorMessage `
        -Message "Mindestens Windows PowerShell 5.1 ist erforderlich." `
        -Title "PowerShell-Version nicht unterstützt"
    return
}

if ($PSVersionTable.PSEdition -eq "Core" -and $PSVersionTable.PSVersion -lt [Version]"7.2") {
    Show-ErrorMessage `
        -Message "Bei PowerShell 7 wird mindestens Version 7.2 benötigt.`n`nAktuell: $($PSVersionTable.PSVersion)" `
        -Title "PowerShell-Version nicht unterstützt"
    return
}

# ============================================================
# 2. Request administrator account
# ============================================================

$AdminAccount = Request-Upn `
    -Prompt "Bitte die E-Mail-Adresse / UPN des Microsoft Teams Admin-Kontos eingeben." `
    -Title "$ScriptName - Admin-Konto" `
    -Example "admin@contoso.com"

if (-not $AdminAccount) {
    Show-Info -Message "Der Vorgang wurde abgebrochen." -Title "Abgebrochen"
    return
}

# ============================================================
# 3. Request administrator password securely
# ============================================================

try {
    $Credential = Get-Credential `
        -UserName $AdminAccount `
        -Message "Bitte das Passwort für das Microsoft Teams Admin-Konto eingeben."
}
catch {
    Show-ErrorMessage `
        -Message "Die Anmeldedaten konnten nicht abgefragt werden.`n`n$($_.Exception.Message)" `
        -Title "Credential-Fehler"
    return
}

if ($null -eq $Credential) {
    Show-Info -Message "Die Passwortabfrage wurde abgebrochen." -Title "Abgebrochen"
    return
}

# ============================================================
# 4. Request target user
# ============================================================

$TargetUser = Request-Upn `
    -Prompt "Bitte die E-Mail-Adresse / UPN des Benutzers eingeben, dessen Teams-Mitgliedschaften exportiert werden sollen." `
    -Title "$ScriptName - Zielbenutzer" `
    -Example "user@contoso.com"

if (-not $TargetUser) {
    Show-Info -Message "Der Vorgang wurde abgebrochen." -Title "Abgebrochen"
    return
}

# ============================================================
# 5. Prepare modules
# ============================================================

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " $ScriptName v$ScriptVersion" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Autor       : $ScriptAuthor" -ForegroundColor DarkGray
Write-Host "PowerShell  : $($PSVersionTable.PSVersion)" -ForegroundColor DarkGray
Write-Host ""

try {
    Ensure-Module `
        -Name "MicrosoftTeams" `
        -Purpose "Das Modul wird benötigt, um Microsoft Teams abzufragen."

    Ensure-Module `
        -Name "ImportExcel" `
        -Purpose "Das Modul wird benötigt, um den Bericht als XLSX-Datei zu erstellen."
}
catch {
    Show-ErrorMessage `
        -Message "Die Vorbereitung ist fehlgeschlagen.`n`n$($_.Exception.Message)" `
        -Title "Modulfehler"
    return
}

$TeamsModule = Get-Module MicrosoftTeams |
    Sort-Object Version -Descending |
    Select-Object -First 1

Write-Host "MicrosoftTeams: $($TeamsModule.Version)" -ForegroundColor DarkGray
Write-Host ""

# ============================================================
# 6. Connect to Microsoft Teams
# ============================================================

$Connected          = $false
$ConnectionInfo     = $null
$ActualAdminAccount = $AdminAccount
$ConnectCommand     = Get-Command Connect-MicrosoftTeams -ErrorAction Stop
$SupportsDisableWAM = $ConnectCommand.Parameters.ContainsKey("DisableWAM")
$SupportsDeviceAuth = $ConnectCommand.Parameters.ContainsKey("UseDeviceAuthentication")

Write-Host "Microsoft Teams Anmeldung..." -ForegroundColor Cyan
Write-Host "Admin-Konto: $AdminAccount" -ForegroundColor Yellow
Write-Host ""

# Attempt 1: Explicit credential + DisableWAM, if available.
# This is the path that actually uses the entered password.
if ($SupportsDisableWAM -and -not $Connected) {
    try {
        Write-Host "[1/4] Benutzername + Passwort (ohne WAM)..." -ForegroundColor Cyan

        $ConnectionInfo = Connect-MicrosoftTeams `
            -Credential $Credential `
            -DisableWAM `
            -ErrorAction Stop

        $Connected = $true
    }
    catch {
        Write-Host "Fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor DarkYellow
        Write-Host ""
    }
}

# Attempt 2: Credential using the module's normal sign-in behavior.
if (-not $Connected) {
    try {
        Write-Host "[2/4] Benutzername + Passwort (Standard)..." -ForegroundColor Cyan

        $ConnectionInfo = Connect-MicrosoftTeams `
            -Credential $Credential `
            -ErrorAction Stop

        $Connected = $true
    }
    catch {
        Write-Host "Fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor DarkYellow
        Write-Host ""
    }
}

# Attempt 3: Interactive authentication.
# Required in many environments with MFA / Conditional Access.
if (-not $Connected) {
    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Die Anmeldung mit Benutzername und Passwort war nicht erfolgreich.`n`nDas ist bei MFA oder Conditional Access normal.`n`nSoll jetzt die interaktive Microsoft-Anmeldung gestartet werden?",
        "Interaktive Anmeldung",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    if ($Answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Write-Host "[3/4] Interaktive Microsoft-Anmeldung..." -ForegroundColor Cyan

            $ConnectionInfo = Connect-MicrosoftTeams `
                -AccountId $AdminAccount `
                -ErrorAction Stop

            $Connected = $true
        }
        catch {
            Write-Host "Fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor DarkYellow
            Write-Host ""
        }
    }
}

# Attempt 4: Device-code authentication.
if (-not $Connected -and $SupportsDeviceAuth) {
    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Die interaktive Anmeldung war nicht erfolgreich.`n`nSoll als letzte Alternative die Gerätecode-Anmeldung gestartet werden?",
        "Gerätecode-Anmeldung",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($Answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Write-Host "[4/4] Gerätecode-Anmeldung..." -ForegroundColor Cyan

            $ConnectionInfo = Connect-MicrosoftTeams `
                -AccountId $AdminAccount `
                -UseDeviceAuthentication `
                -ErrorAction Stop

            $Connected = $true
        }
        catch {
            Write-Host "Fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor DarkYellow
            Write-Host ""
        }
    }
}

if (-not $Connected) {
    Show-ErrorMessage `
        -Message "Es konnte keine Verbindung zu Microsoft Teams hergestellt werden.`n`nPrüfe Admin-Konto, MFA/Conditional Access, Netzwerkzugriff und die MicrosoftTeams-Modulversion." `
        -Title "Anmeldefehler"
    return
}

if ($null -ne $ConnectionInfo -and $ConnectionInfo.Account) {
    $ActualAdminAccount = $ConnectionInfo.Account
}

Write-Host "Anmeldung erfolgreich: $ActualAdminAccount" -ForegroundColor Green
Write-Host ""

# ============================================================
# 7. Resolve target user
# ============================================================

$UserUPN     = $TargetUser
$DisplayName = $TargetUser

try {
    $UserInfo = Get-CsOnlineUser `
        -Identity $TargetUser `
        -ErrorAction Stop

    if ($UserInfo.UserPrincipalName) {
        $UserUPN = $UserInfo.UserPrincipalName
    }

    if ($UserInfo.DisplayName) {
        $DisplayName = $UserInfo.DisplayName
    }
}
catch {
    Write-Warning "Get-CsOnlineUser konnte den Benutzer nicht auflösen. Der Export wird mit der eingegebenen UPN fortgesetzt."
}

Write-Host "Zielbenutzer : $DisplayName" -ForegroundColor Green
Write-Host "UPN          : $UserUPN" -ForegroundColor Green
Write-Host ""

# ============================================================
# 8. Read Teams
# ============================================================

try {
    $Teams = @(
        Get-Team `
            -User $UserUPN `
            -ErrorAction Stop
    )
}
catch {
    Show-ErrorMessage `
        -Message "Die Teams des Benutzers konnten nicht abgerufen werden.`n`n$($_.Exception.Message)" `
        -Title "Teams konnten nicht abgerufen werden"

    Disconnect-TeamsSafe
    return
}

if ($Teams.Count -eq 0) {
    Show-Info `
        -Message "Für den Benutzer wurden keine Teams gefunden.`n`n$DisplayName`n$UserUPN" `
        -Title "Keine Teams gefunden"

    Disconnect-TeamsSafe
    return
}

Write-Host "Gefundene Teams: $($Teams.Count)" -ForegroundColor Green
Write-Host ""

# ============================================================
# 9. Determine role in each Team
# ============================================================

$Result  = @()
$Counter = 0

foreach ($Team in $Teams) {
    $Counter++

    Write-Host "[$Counter/$($Teams.Count)] $($Team.DisplayName)" -ForegroundColor DarkGray

    try {
        $Membership = @(
            Get-TeamUser `
                -GroupId $Team.GroupId `
                -ErrorAction Stop |
                Where-Object {
                    $_.User -ieq $UserUPN
                }
        )

        if ($Membership.Count -gt 0) {
            $RawRole = [string]$Membership[0].Role

            switch ($RawRole.ToLowerInvariant()) {
                "owner"  { $Role = "Owner" }
                "member" { $Role = "Member" }
                "guest"  { $Role = "Guest" }
                default  { $Role = $RawRole }
            }
        }
        else {
            $Role = "Unknown"
        }

        $Result += [PSCustomObject]@{
            TeamName = $Team.DisplayName
            Role     = $Role
            GroupId  = $Team.GroupId
        }
    }
    catch {
        Write-Warning "Team '$($Team.DisplayName)' konnte nicht vollständig ausgelesen werden."

        $Result += [PSCustomObject]@{
            TeamName = $Team.DisplayName
            Role     = "Error"
            GroupId  = $Team.GroupId
        }
    }
}

# ============================================================
# 10. Sort and summarize
# ============================================================

$SortedResult = $Result |
    Sort-Object `
        @{
            Expression = {
                switch ($_.Role) {
                    "Owner"   { 1 }
                    "Member"  { 2 }
                    "Guest"   { 3 }
                    "Unknown" { 4 }
                    "Error"   { 5 }
                    default   { 6 }
                }
            }
        },
        TeamName

$OwnerCount   = @($Result | Where-Object Role -eq "Owner").Count
$MemberCount  = @($Result | Where-Object Role -eq "Member").Count
$GuestCount   = @($Result | Where-Object Role -eq "Guest").Count
$UnknownCount = @($Result | Where-Object Role -eq "Unknown").Count
$ErrorCount   = @($Result | Where-Object Role -eq "Error").Count

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " ERGEBNIS" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""

$SortedResult |
    Format-Table TeamName, Role, GroupId -AutoSize

Write-Host ""
Write-Host "Teams insgesamt : $($Result.Count)" -ForegroundColor Green
Write-Host "Owner           : $OwnerCount" -ForegroundColor Green
Write-Host "Member          : $MemberCount" -ForegroundColor Green

if ($GuestCount -gt 0) {
    Write-Host "Guest           : $GuestCount" -ForegroundColor Green
}

if ($UnknownCount -gt 0) {
    Write-Host "Unknown         : $UnknownCount" -ForegroundColor Yellow
}

if ($ErrorCount -gt 0) {
    Write-Host "Fehler          : $ErrorCount" -ForegroundColor Red
}

# ============================================================
# 11. Determine Desktop path and output file
# ============================================================

$Desktop = [Environment]::GetFolderPath("Desktop")

if ([string]::IsNullOrWhiteSpace($Desktop) -or -not (Test-Path $Desktop)) {
    Show-ErrorMessage `
        -Message "Der Desktop-Pfad konnte nicht ermittelt werden." `
        -Title "Desktop nicht gefunden"

    Disconnect-TeamsSafe
    return
}

$SafeName = $DisplayName -replace '[\\/:*?"<>|]', "_"
$Date     = Get-Date -Format "yyyy-MM-dd"
$Time     = Get-Date -Format "HHmmss"
$FileName = "$SafeName`_Teams_$Date`_$Time.xlsx"
$Path     = Join-Path $Desktop $FileName

# ============================================================
# 12. Create Excel report
# ============================================================

try {
    $SortedResult |
        Export-Excel `
            -Path $Path `
            -WorksheetName "Teams" `
            -TableName "TeamsMembership" `
            -AutoSize `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow `
            -ErrorAction Stop

    $Info = @(
        [PSCustomObject]@{
            Information = "Report"
            Value       = $ScriptName
        }
        [PSCustomObject]@{
            Information = "Script Version"
            Value       = $ScriptVersion
        }
        [PSCustomObject]@{
            Information = "Target User"
            Value       = $DisplayName
        }
        [PSCustomObject]@{
            Information = "Target UPN"
            Value       = $UserUPN
        }
        [PSCustomObject]@{
            Information = "Admin Account"
            Value       = $ActualAdminAccount
        }
        [PSCustomObject]@{
            Information = "Teams Total"
            Value       = $Result.Count
        }
        [PSCustomObject]@{
            Information = "Owner"
            Value       = $OwnerCount
        }
        [PSCustomObject]@{
            Information = "Member"
            Value       = $MemberCount
        }
        [PSCustomObject]@{
            Information = "Guest"
            Value       = $GuestCount
        }
        [PSCustomObject]@{
            Information = "Unknown"
            Value       = $UnknownCount
        }
        [PSCustomObject]@{
            Information = "Errors"
            Value       = $ErrorCount
        }
        [PSCustomObject]@{
            Information = "Export Date"
            Value       = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        [PSCustomObject]@{
            Information = "PowerShell"
            Value       = $PSVersionTable.PSVersion.ToString()
        }
        [PSCustomObject]@{
            Information = "MicrosoftTeams Module"
            Value       = $TeamsModule.Version.ToString()
        }
    )

    $Info |
        Export-Excel `
            -Path $Path `
            -WorksheetName "Info" `
            -AutoSize `
            -BoldTopRow `
            -Append `
            -ErrorAction Stop
}
catch {
    Show-ErrorMessage `
        -Message "Beim Erstellen der Excel-Datei ist ein Fehler aufgetreten.`n`n$($_.Exception.Message)" `
        -Title "Excel-Exportfehler"

    Disconnect-TeamsSafe
    return
}

# ============================================================
# 13. Disconnect and finish
# ============================================================

Disconnect-TeamsSafe

$Message = @"
Export erfolgreich!

Benutzer:
$DisplayName

E-Mail / UPN:
$UserUPN

Teams insgesamt: $($Result.Count)
Owner: $OwnerCount
Member: $MemberCount
Guest: $GuestCount

Excel-Datei:
$Path
"@

Show-Info `
    -Message $Message `
    -Title "$ScriptName - Erfolgreich"

try {
    Invoke-Item -Path $Path -ErrorAction Stop
}
catch {
    Show-Info `
        -Message "Die Excel-Datei wurde erstellt, konnte aber nicht automatisch geöffnet werden.`n`n$Path" `
        -Title "Excel-Datei erstellt"
}
