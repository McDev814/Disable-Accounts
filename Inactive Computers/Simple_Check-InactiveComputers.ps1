<#
.SYNOPSIS
    Reports - and optionally disables - enabled computer accounts that have not authenticated
    to Active Directory in N days (default 90). Supports CMMC IA.L2-3.5.6.

.DESCRIPTION
    A computer is inactive when BOTH its LastLogonDate and its machine PasswordLastSet are older
    than the threshold (or empty). Skipped automatically: disabled computers, domain controllers,
    managed service accounts, the Entra Seamless SSO account (AZUREADSSOACC), computers created
    within the threshold, and anything in -ExcludeOUs / -ExcludeComputers.

    Without -Commit : report only. Nothing in AD is changed.
    With    -Commit : each inactive computer is disabled, its description is stamped, and it is
                      moved to the OU named 'Auto Disabled'.

    Every run writes a CSV to the report folder (default C:\Reports).

.PARAMETER Commit
    Disable and move the inactive computers.

.PARAMETER Force
    With -Commit: skip the "type YES to continue" prompt (use for scheduled tasks).

.PARAMETER Credential
    Optional AD credential (e.g. your admin account when running as your regular user).

.EXAMPLE
    .\Check-InactiveComputers.ps1
    Report only -> C:\Reports\InactiveComputers_<date>_ReportOnly.csv

.EXAMPLE
    .\Check-InactiveComputers.ps1 -Commit
    Disable and move inactive computers (asks for confirmation first).

.EXAMPLE
    .\Check-InactiveComputers.ps1 -Commit -Force
    Same, without the prompt - for a scheduled task.

.NOTES
    Exit codes: 0 = success, 1 = one or more errors.
#>
[CmdletBinding()]
param(
    [switch]$Commit,
    [switch]$Force,
    [int]$DaysInactive = 90,

    # Leave empty to search the whole domain
    [string]$SearchBase,

    # Name of the OU disabled computers are moved to. Must exist exactly once in the domain.
    [string]$DisabledOUName = 'Auto Disabled',

    # ===== OUs to skip =====
    # Use just the OU name ('OU=Test Servers') to match it anywhere, or a full DN to match one
    # specific OU. Computers in sub-OUs of an excluded OU are skipped too.
    [string[]]$ExcludeOUs = @(
        'OU=Servers,OU=Acutec,DC=Acutec,DC=local'      # includes OU=Dev Servers
        # Add more OUs here
    ),

    # ===== Individual computers to skip (names, without the trailing $) =====
    [string[]]$ExcludeComputers = @(
        # 'VEEAM-SCESX01B'
    ),

    [string]$ReportFolder = 'C:\Reports',
    [System.Management.Automation.PSCredential]$Credential,
    [string]$Server
)

Import-Module ActiveDirectory -ErrorAction Stop

$runStart  = Get-Date
$mode      = if ($Commit) { 'Commit' } else { 'ReportOnly' }
$cutoff    = $runStart.AddDays(-$DaysInactive)
$exitCode  = 0
$eventSource = 'Check-InactiveComputers'   # optional - see the guide, "Windows Event Log"

$ad = @{}
if ($Server)     { $ad.Server = $Server }
if ($Credential) { $ad.Credential = $Credential }

# ---------- helpers ----------
function Get-ParentDN([string]$DistinguishedName) {
    $DistinguishedName -replace '^.*?(?<!\\),', ''          # first unescaped comma
}

function Test-ExcludedOU([string]$DistinguishedName) {
    $parent = ',' + (Get-ParentDN $DistinguishedName) + ','
    foreach ($ou in $ExcludeOUs) {
        if ([string]::IsNullOrWhiteSpace($ou)) { continue }
        $needle = ',' + ($ou.Trim() -replace ',\s+', ',') + ','
        if ($parent.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

function Write-RunEvent([int]$EventId, [string]$EntryType, [string]$Message) {
    # Writes to the Application log only if the event source has been registered (see guide)
    try {
        if ([System.Diagnostics.EventLog]::SourceExists($eventSource)) {
            Write-EventLog -LogName Application -Source $eventSource -EventId $EventId -EntryType $EntryType -Message $Message
        }
    } catch { }
}

# ---------- setup ----------
try {
    $null = New-Item -ItemType Directory -Path $ReportFolder -Force -ErrorAction Stop
} catch {
    Write-Error "Cannot create or access the report folder '$ReportFolder': $($_.Exception.Message)"
    exit 1
}
$reportFile = Join-Path $ReportFolder ("InactiveComputers_{0}_{1}.csv" -f $runStart.ToString('yyyyMMdd_HHmmss'), $mode)

if (-not $SearchBase) { $SearchBase = (Get-ADDomain @ad).DistinguishedName }

$targetOU = $null
if ($Commit) {
    $ous = @(Get-ADOrganizationalUnit -Filter "Name -eq '$DisabledOUName'" @ad)
    if ($ous.Count -ne 1) {
        Write-Error "Expected exactly one OU named '$DisabledOUName' but found $($ous.Count). Create it (or remove the duplicate) and run again. Nothing was changed."
        exit 1
    }
    $targetOU = $ous[0].DistinguishedName
}

Write-Host "Mode: $mode | Threshold: $DaysInactive days (before $($cutoff.ToString('yyyy-MM-dd'))) | Search: $SearchBase"
if ($Commit) { Write-Host "Disabled computers will be moved to: $targetOU" }

# ---------- find inactive computers ----------
$excluded = 0
$results = foreach ($c in Get-ADComputer -SearchBase $SearchBase -Filter 'Enabled -eq $true' @ad `
        -Properties Enabled, LastLogonDate, PasswordLastSet, whenCreated, OperatingSystem, Description, PrimaryGroupID, ObjectClass) {

    if (-not ($c.Enabled -eq $true -and
              $c.ObjectClass -eq 'computer' -and                 # skip gMSAs
              $c.PrimaryGroupID -notin 516, 521 -and             # skip domain controllers
              $c.Name -ne 'AZUREADSSOACC' -and                   # skip Entra Seamless SSO account
              $c.whenCreated -lt $cutoff -and                    # skip newly created computers
              ($null -eq $c.LastLogonDate   -or $c.LastLogonDate   -lt $cutoff) -and
              ($null -eq $c.PasswordLastSet -or $c.PasswordLastSet -lt $cutoff))) { continue }

    if ((Test-ExcludedOU $c.DistinguishedName) -or ($ExcludeComputers -contains $c.Name)) { $excluded++; continue }

    [pscustomobject]@{
        Name                = $c.Name
        OperatingSystem     = $c.OperatingSystem
        LastLogonDate       = $c.LastLogonDate
        PasswordLastSet     = $c.PasswordLastSet
        WhenCreated         = $c.whenCreated
        DaysSinceLogon      = if ($c.LastLogonDate) { ($runStart - $c.LastLogonDate).Days } else { 'Never' }
        OriginalOU          = Get-ParentDN $c.DistinguishedName
        OriginalDescription = $c.Description
        DistinguishedName   = $c.DistinguishedName
        Action              = if ($Commit) { 'Pending' } else { 'WouldDisable' }
    }
}
$results = @($results | Sort-Object LastLogonDate)

$results | Format-Table Name, OperatingSystem, LastLogonDate, DaysSinceLogon, OriginalOU -AutoSize |
    Out-String -Width 220 | Write-Host
Write-Host "$($results.Count) inactive enabled computer(s) found ($excluded more skipped by exclusions)."

# ---------- disable and move ----------
if ($Commit -and $results.Count -gt 0) {
    $proceed = $true
    if (-not $Force) {
        $answer = Read-Host "Type YES to disable these $($results.Count) computer(s) and move them to '$DisabledOUName'"
        $proceed = ($answer -ceq 'YES')
    }

    if (-not $proceed) {
        foreach ($r in $results) { $r.Action = 'Cancelled' }
        Write-Host 'Cancelled - nothing was changed.' -ForegroundColor Yellow
    }
    else {
        $today = $runStart.ToString('yyyy-MM-dd')
        foreach ($r in $results) {
            try {
                $last = if ($r.LastLogonDate) { $r.LastLogonDate.ToString('yyyy-MM-dd') } else { 'never' }
                $note = "Auto disabled $today - inactive $DaysInactive+ days (last logon $last)"
                if ($r.OriginalDescription) { $note = "$note | $($r.OriginalDescription)" }
                if ($note.Length -gt 1024)  { $note = $note.Substring(0, 1024) }

                Set-ADComputer    -Identity $r.DistinguishedName -Description $note @ad
                Disable-ADAccount -Identity $r.DistinguishedName @ad
                if ($r.OriginalOU -ne $targetOU) {
                    Move-ADObject -Identity $r.DistinguishedName -TargetPath $targetOU @ad
                }
                $r.Action = 'Disabled and moved'
                Write-Host "Disabled $($r.Name)"
                Write-RunEvent 1001 Information "Disabled computer $($r.Name) (last logon $last) and moved it from $($r.OriginalOU) to $targetOU."
            }
            catch {
                $r.Action = "ERROR: $($_.Exception.Message)"
                $exitCode = 1
                Write-Host "FAILED $($r.Name): $($_.Exception.Message)" -ForegroundColor Red
                Write-RunEvent 1002 Error "Failed to disable computer $($r.Name): $($_.Exception.Message)"
            }
        }
    }
}

# ---------- report ----------
try {
    $results | Select-Object Name, OperatingSystem, LastLogonDate, PasswordLastSet, WhenCreated, DaysSinceLogon,
                             Action, OriginalOU, OriginalDescription, DistinguishedName |
        Export-Csv $reportFile -NoTypeInformation -ErrorAction Stop
    Write-Host "Report saved to $reportFile"
} catch {
    Write-Warning "Could not save the report: $($_.Exception.Message)"
    $exitCode = 1
}

$summary = "Inactive computer check ($mode): $($results.Count) inactive, $excluded excluded, " +
           "$(@($results | Where-Object Action -eq 'Disabled and moved').Count) disabled, " +
           "$(@($results | Where-Object Action -like 'ERROR*').Count) errors. Report: $reportFile"
Write-RunEvent 1000 $(if ($exitCode) { 'Warning' } else { 'Information' }) $summary

exit $exitCode