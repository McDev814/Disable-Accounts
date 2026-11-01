#Requires -Version 5.1
#Requires -Modules ActiveDirectory, Microsoft.Graph.Authentication, Microsoft.Graph.Users
<#
.SYNOPSIS
    Finds and (optionally) disables user and computer accounts that have been inactive
    for a defined period. Supports CMMC Level 2 practice IA.L2-3.5.6.

.DESCRIPTION
    For every enabled account in scope, the script determines the most recent activity from:
      * On-prem AD  : lastLogonTimestamp (LastLogonDate)
      * Entra ID    : signInActivity.lastSuccessfulSignInDateTime (Microsoft Graph, GCC High)
    and uses whichever date is LATER. An account is inactive only when that effective date is
    older than DaysInactive. Accounts that have never signed in anywhere are judged by their
    creation date (NewAccountGraceDays).

    Scope:
      * AD users (synced or not)      -> disabled in on-prem AD (sync carries it to Entra)
      * AD computers                  -> AD data only (LastLogonDate AND PasswordLastSet)
      * Entra cloud-only users        -> disabled through Microsoft Graph

    The script runs in REPORT-ONLY mode unless -Commit is specified. Every run writes a CSV
    (evidence) and a transcript to the log directory.

.PARAMETER Commit
    Actually disable and move the accounts. Without it, nothing is changed.

.PARAMETER AuthMode
    Interactive  - you sign in to Graph in a browser (workstation use).
    Certificate  - app-only sign-in with the certificate in the configuration (server use).

.PARAMETER ADCredential
    Optional credential for AD reads/writes (e.g. your admin account when running from a
    workstation as your regular user). Not used when running as the gMSA.

.PARAMETER Server
    Optional domain controller to target for all AD operations.

.PARAMETER IgnoreSafetyLimit
    Allow a -Commit run to disable more than MaxDisablePerRun accounts (initial cleanup only).

.EXAMPLE
    .\Disable-InactiveAccounts.ps1
    Report-only run, interactive Graph sign-in.

.EXAMPLE
    .\Disable-InactiveAccounts.ps1 -AuthMode Certificate -Commit
    Scheduled production run on the server.

.NOTES
    Exit codes: 0 = success, 1 = error (nothing further changed), 2 = stopped by safety limit.
#>
[CmdletBinding()]
param(
    [switch]$Commit,
    [ValidateSet('Interactive', 'Certificate')]
    [string]$AuthMode = 'Interactive',
    [System.Management.Automation.PSCredential]$ADCredential,
    [string]$Server,
    [switch]$IgnoreSafetyLimit
)

#region ======================= CONFIGURATION - EDIT THIS SECTION =======================

$Config = @{

    # --- Policy -------------------------------------------------------------------------
    DaysInactive         = 90      # Inactivity threshold (CMMC IA.L2-3.5.6 defined period)
    NewAccountGraceDays  = 90      # Never-signed-in accounts are left alone this long after creation
    ReEnableGraceDays    = 14      # Skip accounts whose enabled/disabled state changed this recently
                                   # (gives a re-enabled user time to sign in)

    # --- What to process ----------------------------------------------------------------
    IncludeUsers          = $true
    IncludeComputers      = $true
    IncludeCloudOnlyUsers = $true  # Entra accounts that are not synced from AD
    IncludeGuests         = $true  # Entra B2B guest accounts (cloud-only)

    # --- Where to look in AD ------------------------------------------------------------
    # Leave empty to search the whole domain. Otherwise list the OUs to search.
    SearchBases = @(
        # 'OU=Acutec Users,DC=yourdomain,DC=local'
        # 'OU=Acutec Computers,DC=yourdomain,DC=local'
    )

    # Where disabled objects are moved. Set to '' to disable without moving.
    DisabledUsersOU     = 'OU=Disabled Users,DC=yourdomain,DC=local'
    DisabledComputersOU = 'OU=Disabled Computers,DC=yourdomain,DC=local'

    # --- EXEMPTIONS ---------------------------------------------------------------------
    # Any object located in (or below) one of these OUs is never disabled.
    # You may enter just the OU name ('OU=Service Accounts') to match that OU anywhere in
    # the domain, or a full DN ('OU=Service Accounts,OU=Acutec,DC=yourdomain,DC=local')
    # to match one specific OU.
    ExemptOUs = @(
        'OU=Shared Inbox'
        'OU=CMM Logins'
        'OU=Service Accounts'
        # Add more exempt OUs here
    )

    # Members of this AD group (including nested members) are never disabled.
    # Set to '' if you do not use an exemption group. If set, the group MUST exist.
    ExemptGroup = 'AD-Inactivity-Exempt'

    # AD sAMAccountNames that are always exempt (computer accounts end in $).
    ExemptSamAccountNames = @(
        'Administrator'
        'krbtgt'
        'Guest'
        'AZUREADSSOACC$'      # Entra Seamless SSO computer account - never disable
    )

    # Cloud-only Entra accounts that are always exempt (break-glass / emergency access, etc.).
    # Use object IDs (preferred) and/or UPNs.
    ExemptCloudObjectIds = @(
        # '00000000-0000-0000-0000-000000000000'
    )
    ExemptCloudUPNs = @(
        # 'breakglass01@acutec.onmicrosoft.us'
    )

    # --- Safety -------------------------------------------------------------------------
    MaxDisablePerRun = 25          # A -Commit run that would disable more than this stops without
                                   # changing anything (override with -IgnoreSafetyLimit)

    # --- Microsoft Graph (GCC High) -----------------------------------------------------
    Graph = @{
        Environment           = 'USGov'                                  # GCC High
        TenantId              = '44e9b2c4-4a43-4cac-bdb1-d719bca04d42'
        ClientId              = ''   # App registration (Application) ID - Certificate mode only
        CertificateThumbprint = ''   # Certificate in LocalMachine\My or CurrentUser\My - Certificate mode only
        DisableWAM            = $true  # Interactive mode: use a browser instead of the Windows broker
    }

    # --- Logging ------------------------------------------------------------------------
    LogDirectory     = 'C:\Scripts\InactiveAccounts\Logs'
    LogAllAccounts   = $true       # $true = CSV lists every evaluated account (best audit evidence)
    LogRetentionDays = 0           # 0 = never delete old logs

    # --- Email summary (optional) -------------------------------------------------------
    Mail = @{
        Enabled    = $false
        SmtpServer = 'smtp.yourdomain.local'
        Port       = 25
        UseSsl     = $false
        From       = 'inactive-accounts@acutecprecision.com'
        To         = @('it@acutecprecision.com')
    }
}

#endregion ===================================================================================


#region ----------------------------------- Helpers -----------------------------------------

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Warning $line }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function ConvertTo-LocalDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    $dt = [datetime]$Value
    if ($dt.Year -le 1601) { return $null }
    if ($dt.Kind -eq [DateTimeKind]::Utc) { return $dt.ToLocalTime() }
    return $dt
}

function Get-LatestDate {
    param([object[]]$Dates)
    $valid = @($Dates | Where-Object { $null -ne $_ })
    if ($valid.Count -eq 0) { return $null }
    return ($valid | Sort-Object -Descending | Select-Object -First 1)
}

function Format-Date {
    param($Value)
    if ($null -eq $Value) { return '' }
    return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm')
}

function Get-ParentDN {
    param([string]$DistinguishedName)
    # Split on the first comma that is not escaped (handles names like "Smith\, John")
    $m = [regex]::Match($DistinguishedName, '(?<!\\),')
    if ($m.Success) { return $DistinguishedName.Substring($m.Index + 1) }
    return ''
}

function Get-ExemptOUMatch {
    param([string]$DistinguishedName)
    $parent = ',' + (Get-ParentDN $DistinguishedName) + ','
    foreach ($ou in $Config.ExemptOUs) {
        if ([string]::IsNullOrWhiteSpace($ou)) { continue }
        $needle = ',' + (($ou.Trim()) -replace ',\s+', ',') + ','
        if ($parent.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $ou.Trim() }
    }
    return $null
}

function ConvertTo-LdapFilterValue {
    param([string]$Value)
    return ($Value -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29')
}

function Get-InactivityVerdict {
    param($LastActivity, $Created)
    if ($null -ne $LastActivity) {
        $days = [int][math]::Floor(((Get-Date) - $LastActivity).TotalDays)
        if ($LastActivity -lt $script:Cutoff) {
            return @{ Inactive = $true; Days = $days; Reason = "No activity since $(Format-Date $LastActivity) ($days days)" }
        }
        return @{ Inactive = $false; Days = $days; Reason = 'Active' }
    }
    if ($null -ne $Created -and $Created -gt $script:GraceCutoff) {
        return @{ Inactive = $false; Days = $null; Reason = "Never signed in; created $(Format-Date $Created) (within $($Config.NewAccountGraceDays)-day grace period)" }
    }
    $createdText = if ($Created) { Format-Date $Created } else { 'unknown date' }
    return @{ Inactive = $true; Days = $null; Reason = "Never signed in; created $createdText" }
}

function Get-EnabledStateLastChanged {
    # When was userAccountControl last changed? Used to give re-enabled accounts a grace period.
    param([string]$DistinguishedName)
    try {
        $p = @{ Object = $DistinguishedName; Server = $script:MetadataServer; Properties = 'userAccountControl' }
        if ($ADCredential) { $p.Credential = $ADCredential }
        $md = Get-ADReplicationAttributeMetadata @p -ErrorAction Stop |
              Where-Object { $_.AttributeName -eq 'userAccountControl' } | Select-Object -First 1
        if ($md) { return (ConvertTo-LocalDate $md.LastOriginatingChangeTime) }
    } catch {
        Write-Log "Could not read replication metadata for $DistinguishedName : $($_.Exception.Message)" 'WARN'
    }
    return $null
}

function New-Record {
    param(
        [string]$ObjectType, [string]$Name, [string]$UPN, [string]$DN, [string]$ObjectId,
        [bool]$Synced, $Created, $ADLastLogon, $ADPasswordLastSet, $EntraLastSignIn,
        $EffectiveLastActivity, $DaysInactive, [string]$Status, [string]$Reason, [string]$Description
    )
    [pscustomobject]@{
        RunTime               = $script:RunStart.ToString('yyyy-MM-dd HH:mm:ss')
        Mode                  = $script:Mode
        ObjectType            = $ObjectType
        Name                  = $Name
        UserPrincipalName     = $UPN
        DistinguishedName     = $DN
        OriginalOU            = if ($DN) { Get-ParentDN $DN } else { '' }
        ObjectId              = $ObjectId
        SyncedToEntra         = $Synced
        Created               = Format-Date $Created
        ADLastLogon           = Format-Date $ADLastLogon
        ADPasswordLastSet     = Format-Date $ADPasswordLastSet
        EntraLastSuccessfulSignIn = Format-Date $EntraLastSignIn
        EffectiveLastActivity = Format-Date $EffectiveLastActivity
        DaysInactive          = $DaysInactive
        Status                = $Status       # Active | Exempt | Skipped | Candidate
        Reason                = $Reason
        Action                = 'None'
        OriginalDescription   = $Description
    }
}

function Send-RunSummary {
    param([string]$Subject, [string]$Body, [string[]]$Attachments)
    if (-not $Config.Mail.Enabled) { return }
    try {
        $p = @{
            SmtpServer = $Config.Mail.SmtpServer; Port = $Config.Mail.Port; From = $Config.Mail.From
            To = $Config.Mail.To; Subject = $Subject; Body = $Body; UseSsl = $Config.Mail.UseSsl
        }
        $files = @($Attachments | Where-Object { $_ -and (Test-Path $_) })
        if ($files.Count) { $p.Attachments = $files }
        Send-MailMessage @p -WarningAction SilentlyContinue -ErrorAction Stop
        Write-Log "Summary email sent to $($Config.Mail.To -join ', ')"
    } catch {
        Write-Log "Failed to send summary email: $($_.Exception.Message)" 'WARN'
    }
}

function Connect-GraphForRun {
    $g = $Config.Graph
    if ($AuthMode -eq 'Certificate') {
        if ([string]::IsNullOrWhiteSpace($g.ClientId) -or [string]::IsNullOrWhiteSpace($g.CertificateThumbprint)) {
            throw 'Certificate mode requires Graph.ClientId and Graph.CertificateThumbprint in the configuration.'
        }
        $cert = Get-ChildItem "Cert:\LocalMachine\My\$($g.CertificateThumbprint)" -ErrorAction SilentlyContinue
        if (-not $cert) { $cert = Get-ChildItem "Cert:\CurrentUser\My\$($g.CertificateThumbprint)" -ErrorAction SilentlyContinue }
        if (-not $cert) { throw "Certificate $($g.CertificateThumbprint) not found in LocalMachine\My or CurrentUser\My." }
        Connect-MgGraph -Environment $g.Environment -TenantId $g.TenantId -ClientId $g.ClientId `
            -Certificate $cert -ContextScope Process -NoWelcome -ErrorAction Stop
    }
    else {
        $scopes = @('User.Read.All', 'AuditLog.Read.All')
        if ($Commit -and $Config.IncludeCloudOnlyUsers) { $scopes += 'User.EnableDisableAccount.All' }
        if ($g.DisableWAM) {
            try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction Stop } catch { }
        }
        Connect-MgGraph -Environment $g.Environment -TenantId $g.TenantId -Scopes $scopes `
            -ContextScope Process -NoWelcome -ErrorAction Stop
    }
    $ctx = Get-MgContext
    $who = if ($ctx.Account) { $ctx.Account } else { "app $($ctx.ClientId)" }
    Write-Log "Connected to Microsoft Graph ($($ctx.Environment)) as $who"
}

#endregion


#region ------------------------------------- Main ------------------------------------------

$ErrorActionPreference = 'Stop'
$script:RunStart    = Get-Date
$script:Mode        = if ($Commit) { 'Commit' } else { 'ReportOnly' }
$script:Cutoff      = $RunStart.AddDays(-$Config.DaysInactive)
$script:GraceCutoff = $RunStart.AddDays(-$Config.NewAccountGraceDays)
$reEnableCutoff     = $RunStart.AddDays(-$Config.ReEnableGraceDays)
$exitCode           = 0

$stamp          = $RunStart.ToString('yyyyMMdd_HHmmss')
$null           = New-Item -ItemType Directory -Path $Config.LogDirectory -Force
$csvPath        = Join-Path $Config.LogDirectory "InactiveAccounts_${stamp}_$Mode.csv"
$transcriptPath = Join-Path $Config.LogDirectory "InactiveAccounts_${stamp}_$Mode.log"
Start-Transcript -Path $transcriptPath -Force | Out-Null

$records = New-Object System.Collections.Generic.List[object]

try {
    Write-Log "===== Inactive account job starting - mode: $Mode ====="
    Write-Log "Threshold: $($Config.DaysInactive) days (cutoff $(Format-Date $Cutoff)); new-account grace: $($Config.NewAccountGraceDays) days; re-enable grace: $($Config.ReEnableGraceDays) days"
    Write-Log "Running as Windows user: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"

    # ---------------- AD connection settings ----------------
    $adParams = @{}
    if ($Server)       { $adParams.Server = $Server }
    if ($ADCredential) { $adParams.Credential = $ADCredential }

    $domain = Get-ADDomain @adParams
    $searchBases = @($Config.SearchBases | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($searchBases.Count -eq 0) { $searchBases = @($domain.DistinguishedName) }
    Write-Log "AD domain: $($domain.DNSRoot); search bases: $($searchBases -join ' | ')"

    if ($Server) { $script:MetadataServer = $Server }
    else {
        $dcParams = @{ Discover = $true; DomainName = $domain.DNSRoot }
        $dc = Get-ADDomainController @dcParams
        $script:MetadataServer = [string]($dc.HostName | Select-Object -First 1)
    }

    # Validate destination OUs before any change is possible
    if ($Commit) {
        foreach ($ou in @($Config.DisabledUsersOU, $Config.DisabledComputersOU)) {
            if (-not [string]::IsNullOrWhiteSpace($ou)) {
                $null = Get-ADOrganizationalUnit -Identity $ou @adParams   # throws if missing
            }
        }
    }

    # ---------------- Exemption group ----------------
    $exemptGroupDNs = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if (-not [string]::IsNullOrWhiteSpace($Config.ExemptGroup)) {
        $grp = Get-ADGroup -Identity $Config.ExemptGroup @adParams    # throws if missing - by design
        $filter = '(memberOf:1.2.840.113556.1.4.1941:={0})' -f (ConvertTo-LdapFilterValue $grp.DistinguishedName)
        Get-ADObject -LDAPFilter $filter @adParams | ForEach-Object { [void]$exemptGroupDNs.Add($_.DistinguishedName) }
        Write-Log "Exemption group '$($Config.ExemptGroup)': $($exemptGroupDNs.Count) member object(s)"
    }
    $exemptSams = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $Config.ExemptSamAccountNames | Where-Object { $_ } | ForEach-Object { [void]$exemptSams.Add($_) }

    # ---------------- Microsoft Graph ----------------
    $graphBySid = @{}
    $graphUsers = @()
    if ($Config.IncludeUsers -or $Config.IncludeCloudOnlyUsers) {
        Connect-GraphForRun
        $props = @('id', 'displayName', 'userPrincipalName', 'accountEnabled', 'userType', 'onPremisesSyncEnabled',
                   'onPremisesSecurityIdentifier', 'onPremisesSamAccountName', 'createdDateTime', 'signInActivity')
        $graphUsers = @(Get-MgUser -All -Property $props -ErrorAction Stop)
        Write-Log "Graph returned $($graphUsers.Count) Entra user object(s)"

        $syncedGraph = @($graphUsers | Where-Object { $_.OnPremisesSyncEnabled })
        foreach ($gu in $syncedGraph) {
            if ($gu.OnPremisesSecurityIdentifier) { $graphBySid[$gu.OnPremisesSecurityIdentifier] = $gu }
        }

        # Sanity check: if sign-in data is missing entirely, abort rather than risk mass-disabling
        $withSignIn = @($graphUsers | Where-Object { $_.SignInActivity -and $_.SignInActivity.LastSuccessfulSignInDateTime })
        if ($graphUsers.Count -eq 0 -or $withSignIn.Count -eq 0) {
            throw 'Microsoft Graph returned no users or no sign-in activity data. Aborting to avoid disabling accounts based on AD data alone. Check AuditLog.Read.All permission and Entra ID P1 licensing.'
        }
        Write-Log "$($withSignIn.Count) Entra user(s) have a recorded successful sign-in"
    }

    # ---------------- AD users ----------------
    if ($Config.IncludeUsers) {
        $adUsers = foreach ($sb in $searchBases) {
            Get-ADUser -SearchBase $sb -Filter 'Enabled -eq $true' @adParams `
                -Properties LastLogonDate, whenCreated, Description, SID, UserPrincipalName
        }
        $adUsers = @($adUsers | Sort-Object DistinguishedName -Unique)
        Write-Log "Evaluating $($adUsers.Count) enabled AD user(s)"

        foreach ($u in $adUsers) {
            $gu       = $graphBySid[$u.SID.Value]
            $entra    = if ($gu -and $gu.SignInActivity) { ConvertTo-LocalDate $gu.SignInActivity.LastSuccessfulSignInDateTime } else { $null }
            $adLast   = ConvertTo-LocalDate $u.LastLogonDate
            $created  = ConvertTo-LocalDate $u.whenCreated
            $effective = Get-LatestDate @($adLast, $entra)
            $verdict  = Get-InactivityVerdict -LastActivity $effective -Created $created

            $status = 'Active'; $reason = $verdict.Reason
            $exOU = Get-ExemptOUMatch $u.DistinguishedName
            if     ($u.SID.Value -match '-500$')                     { $status = 'Exempt'; $reason = 'Built-in Administrator (RID 500)' }
            elseif ($exemptSams.Contains($u.SamAccountName))         { $status = 'Exempt'; $reason = 'Exempt account name' }
            elseif ($exOU)                                           { $status = 'Exempt'; $reason = "Exempt OU: $exOU" }
            elseif ($exemptGroupDNs.Contains($u.DistinguishedName))  { $status = 'Exempt'; $reason = "Member of $($Config.ExemptGroup)" }
            elseif ($verdict.Inactive) {
                $status = 'Candidate'
                $uacChanged = Get-EnabledStateLastChanged $u.DistinguishedName
                if ($uacChanged -and $uacChanged -gt $reEnableCutoff) {
                    $status = 'Skipped'
                    $reason = "Account state changed $(Format-Date $uacChanged) (within $($Config.ReEnableGraceDays)-day re-enable grace). $($verdict.Reason)"
                }
            }

            $records.Add((New-Record -ObjectType 'User' -Name $u.SamAccountName -UPN $u.UserPrincipalName `
                -DN $u.DistinguishedName -ObjectId $(if ($gu) { $gu.Id } else { '' }) -Synced ([bool]$gu) `
                -Created $created -ADLastLogon $adLast -ADPasswordLastSet $null -EntraLastSignIn $entra `
                -EffectiveLastActivity $effective -DaysInactive $verdict.Days -Status $status -Reason $reason `
                -Description $u.Description))
        }
    }

    # ---------------- AD computers ----------------
    if ($Config.IncludeComputers) {
        $adComputers = foreach ($sb in $searchBases) {
            Get-ADComputer -SearchBase $sb -Filter 'Enabled -eq $true' @adParams `
                -Properties LastLogonDate, PasswordLastSet, whenCreated, Description, ObjectClass, PrimaryGroupID, ServicePrincipalName
        }
        $adComputers = @($adComputers | Sort-Object DistinguishedName -Unique)
        Write-Log "Evaluating $($adComputers.Count) enabled AD computer object(s)"

        foreach ($c in $adComputers) {
            $adLast   = ConvertTo-LocalDate $c.LastLogonDate
            $pwdLast  = ConvertTo-LocalDate $c.PasswordLastSet
            $created  = ConvertTo-LocalDate $c.whenCreated
            # A computer is only as stale as its most recent logon OR machine password change
            $effective = Get-LatestDate @($adLast, $pwdLast)
            $verdict  = Get-InactivityVerdict -LastActivity $effective -Created $created

            $status = 'Active'; $reason = $verdict.Reason
            $exOU = Get-ExemptOUMatch $c.DistinguishedName
            $isCluster = @($c.ServicePrincipalName | Where-Object { $_ -like 'MSClusterVirtualServer/*' }).Count -gt 0
            if     ($c.ObjectClass -ne 'computer')                   { $status = 'Exempt'; $reason = "Managed service account ($($c.ObjectClass))" }
            elseif ($c.PrimaryGroupID -in 516, 521)                  { $status = 'Exempt'; $reason = 'Domain controller' }
            elseif ($isCluster)                                      { $status = 'Exempt'; $reason = 'Failover cluster name object' }
            elseif ($exemptSams.Contains($c.SamAccountName))         { $status = 'Exempt'; $reason = 'Exempt account name' }
            elseif ($exOU)                                           { $status = 'Exempt'; $reason = "Exempt OU: $exOU" }
            elseif ($exemptGroupDNs.Contains($c.DistinguishedName))  { $status = 'Exempt'; $reason = "Member of $($Config.ExemptGroup)" }
            elseif ($verdict.Inactive) {
                $status = 'Candidate'
                $uacChanged = Get-EnabledStateLastChanged $c.DistinguishedName
                if ($uacChanged -and $uacChanged -gt $reEnableCutoff) {
                    $status = 'Skipped'
                    $reason = "Account state changed $(Format-Date $uacChanged) (within $($Config.ReEnableGraceDays)-day re-enable grace). $($verdict.Reason)"
                }
            }

            $records.Add((New-Record -ObjectType 'Computer' -Name $c.SamAccountName -UPN '' `
                -DN $c.DistinguishedName -ObjectId '' -Synced $false -Created $created `
                -ADLastLogon $adLast -ADPasswordLastSet $pwdLast -EntraLastSignIn $null `
                -EffectiveLastActivity $effective -DaysInactive $verdict.Days -Status $status -Reason $reason `
                -Description $c.Description))
        }
    }

    # ---------------- Entra cloud-only users ----------------
    if ($Config.IncludeCloudOnlyUsers) {
        $exemptIds  = @($Config.ExemptCloudObjectIds | Where-Object { $_ })
        $exemptUpns = @($Config.ExemptCloudUPNs      | Where-Object { $_ })
        $cloudUsers = @($graphUsers | Where-Object { -not $_.OnPremisesSyncEnabled -and $_.AccountEnabled })
        if (-not $Config.IncludeGuests) { $cloudUsers = @($cloudUsers | Where-Object { $_.UserType -ne 'Guest' }) }
        Write-Log "Evaluating $($cloudUsers.Count) enabled cloud-only Entra user(s)"
        if ($exemptIds.Count -eq 0 -and $exemptUpns.Count -eq 0) {
            Write-Log 'No cloud-only exemptions configured. Add your break-glass accounts to ExemptCloudObjectIds before using -Commit.' 'WARN'
        }

        foreach ($cu in $cloudUsers) {
            $entra    = if ($cu.SignInActivity) { ConvertTo-LocalDate $cu.SignInActivity.LastSuccessfulSignInDateTime } else { $null }
            $created  = ConvertTo-LocalDate $cu.CreatedDateTime
            $verdict  = Get-InactivityVerdict -LastActivity $entra -Created $created

            $status = 'Active'; $reason = $verdict.Reason
            if     ($exemptIds -contains $cu.Id)                 { $status = 'Exempt'; $reason = 'Exempt cloud object ID' }
            elseif ($exemptUpns -contains $cu.UserPrincipalName) { $status = 'Exempt'; $reason = 'Exempt cloud UPN' }
            elseif ($verdict.Inactive)                           { $status = 'Candidate' }

            $type = if ($cu.UserType -eq 'Guest') { 'CloudGuest' } else { 'CloudUser' }
            $records.Add((New-Record -ObjectType $type -Name $cu.DisplayName -UPN $cu.UserPrincipalName `
                -DN '' -ObjectId $cu.Id -Synced $false -Created $created -ADLastLogon $null `
                -ADPasswordLastSet $null -EntraLastSignIn $entra -EffectiveLastActivity $entra `
                -DaysInactive $verdict.Days -Status $status -Reason $reason -Description ''))
        }
    }

    # ---------------- Decide and act ----------------
    $candidates = @($records | Where-Object { $_.Status -eq 'Candidate' })
    Write-Log "Accounts meeting the inactivity criteria: $($candidates.Count)"

    $safetyTripped = $Commit -and -not $IgnoreSafetyLimit -and ($candidates.Count -gt $Config.MaxDisablePerRun)

    if (-not $Commit) {
        foreach ($r in $candidates) { $r.Action = 'WouldDisable' }
        Write-Log 'REPORT-ONLY mode: no changes were made. Re-run with -Commit to disable the candidates.'
    }
    elseif ($safetyTripped) {
        foreach ($r in $candidates) { $r.Action = 'NotProcessed-SafetyLimit' }
        Write-Log "SAFETY LIMIT: $($candidates.Count) candidates exceeds MaxDisablePerRun ($($Config.MaxDisablePerRun)). No changes were made. Review the CSV; use -IgnoreSafetyLimit only for an intentional cleanup." 'ERROR'
        $exitCode = 2
    }
    else {
        if ($IgnoreSafetyLimit) { Write-Log "Safety limit overridden with -IgnoreSafetyLimit ($($candidates.Count) candidates)" 'WARN' }
        $today = $RunStart.ToString('yyyy-MM-dd')

        foreach ($r in $candidates) {
            try {
                if ($r.ObjectType -in 'User', 'Computer') {
                    $lastText = if ($r.EffectiveLastActivity) { $r.EffectiveLastActivity } else { 'never' }
                    $note = "Disabled $today by inactivity job ($($Config.DaysInactive)+ days, last activity: $lastText)"
                    if ($r.OriginalDescription) { $note = "$note | $($r.OriginalDescription)" }
                    if ($note.Length -gt 1024) { $note = $note.Substring(0, 1024) }

                    Set-ADObject      -Identity $r.DistinguishedName -Replace @{ description = $note } @adParams
                    Disable-ADAccount -Identity $r.DistinguishedName @adParams

                    $target = if ($r.ObjectType -eq 'User') { $Config.DisabledUsersOU } else { $Config.DisabledComputersOU }
                    if (-not [string]::IsNullOrWhiteSpace($target)) {
                        Move-ADObject -Identity $r.DistinguishedName -TargetPath $target @adParams
                        $r.Action = "Disabled; moved to $target"
                    } else {
                        $r.Action = 'Disabled'
                    }
                }
                else {
                    Update-MgUser -UserId $r.ObjectId -AccountEnabled:$false -ErrorAction Stop
                    $r.Action = 'Disabled (Entra)'
                }
                Write-Log "Disabled $($r.ObjectType) $($r.Name)"
            }
            catch {
                $r.Action = "ERROR: $($_.Exception.Message)"
                Write-Log "Failed to process $($r.ObjectType) $($r.Name): $($_.Exception.Message)" 'ERROR'
                if ($exitCode -eq 0) { $exitCode = 1 }
            }
        }
    }
}
catch {
    Write-Log "Run aborted: $($_.Exception.Message)" 'ERROR'
    $exitCode = 1
}
finally {
    # ---------------- Evidence CSV ----------------
    try {
        $toWrite = if ($Config.LogAllAccounts) { $records } else { @($records | Where-Object { $_.Status -ne 'Active' }) }
        if (@($toWrite).Count -gt 0) {
            $toWrite | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Log "Evidence CSV written: $csvPath"
        }
    } catch {
        Write-Log "Failed to write CSV: $($_.Exception.Message)" 'ERROR'
        $exitCode = 1
    }

    # ---------------- Summary ----------------
    $summary = $records | Group-Object ObjectType, Status | Sort-Object Name |
               ForEach-Object { '{0,-30} {1,6}' -f $_.Name, $_.Count }
    $summaryText = @(
        "Inactive account job - $Mode - $(Format-Date $RunStart)"
        "Threshold: $($Config.DaysInactive) days. Exit code: $exitCode"
        ''
        'Type, Status                     Count'
        $summary
        ''
        "Actions: " + (($records | Where-Object { $_.Action -ne 'None' } | Group-Object { ($_.Action -split ';')[0] } |
                        ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', ')
        "CSV: $csvPath"
    ) -join [Environment]::NewLine
    Write-Host ''
    Write-Host $summaryText

    $subjectState = switch ($exitCode) { 0 { 'OK' } 2 { 'STOPPED - safety limit' } default { 'ERRORS' } }
    Send-RunSummary -Subject "[Inactive Accounts] $Mode - $subjectState" -Body $summaryText -Attachments @($csvPath)

    # ---------------- Log retention ----------------
    if ($Config.LogRetentionDays -gt 0) {
        Get-ChildItem $Config.LogDirectory -File -Filter 'InactiveAccounts_*' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$Config.LogRetentionDays) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    Stop-Transcript | Out-Null
}

exit $exitCode

#endregion