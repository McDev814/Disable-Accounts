# Disable-InactiveAccounts.ps1 – Windows Server Setup Guide

Use this guide to set up the script as a **daily scheduled task** on a domain-joined Windows Server. This is the production setup for CMMC IA.L2-3.5.6.

## How it fits together

| Part | Purpose |
|---|---|
| **Scheduled task**, daily | Runs the script without anyone signed in |
| **gMSA**: `gmsa-InactAcct$` | The identity the task runs as. Its password is managed by AD and never seen by anyone. Holds only the AD rights it needs. |
| **App registration and certificate** | App-only access to Microsoft Graph (GCC High) for reading sign-in data and disabling cloud-only accounts. No client secret. |
| **Log folder** | CSV evidence and a transcript for every run |

Throughout this guide, replace these placeholders:

| Placeholder | Replace with |
|---|---|
| `YOURDOMAIN` | Your domain's NetBIOS name |
| `DC=yourdomain,DC=local` | Your domain DN |
| `SERVER01` | Name of the server that runs the task |

---

## Step 1 – Prepare the server

**Choose the server.** Use a domain-joined Windows Server 2019, 2022 or 2025 in your management tier. Don't use a domain controller, and don't use a server end users sign in to. Run everything in this guide in an **elevated** PowerShell window on that server, signed in with a domain admin account.

**Install the AD module:**

```powershell
Install-WindowsFeature RSAT-AD-PowerShell
```

**Install the Graph modules for all users, so the gMSA can load them.** Match the version on your workstation:

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users -RequiredVersion 2.41.1 -Scope AllUsers
```

If the server has no internet access:

1. On a machine that does, run `Save-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users -RequiredVersion 2.41.1 -Path C:\Temp\Modules`.
2. Copy both module folders to `C:\Program Files\WindowsPowerShell\Modules` on the server.

---

## Step 2 – Create the AD objects

```powershell
$domainDN = (Get-ADDomain).DistinguishedName

New-ADOrganizationalUnit -Name 'Disabled Users'     -Path $domainDN -ProtectedFromAccidentalDeletion $true
New-ADOrganizationalUnit -Name 'Disabled Computers' -Path $domainDN -ProtectedFromAccidentalDeletion $true

New-ADGroup -Name 'AD-Inactivity-Exempt' -GroupScope Global -GroupCategory Security `
    -Path "OU=Groups,$domainDN" `
    -Description 'Members are exempt from automatic inactivity disable (CMMC IA.L2-3.5.6). Document the reason for each member.'
```

Skip any object that already exists. **Optional:** link a GPO to the two Disabled OUs that denies interactive and RDP logon.

---

## Step 3 – Create the gMSA

```powershell
# 1. A KDS root key must exist (one per forest)
Get-KdsRootKey
# If nothing is returned, create one. AD waits up to 10 hours before gMSAs can use it.
Add-KdsRootKey -EffectiveImmediately

# 2. Create the gMSA and allow only this server to use it
New-ADServiceAccount -Name 'gmsa-InactAcct' `
    -DNSHostName "gmsa-InactAcct.$((Get-ADDomain).DNSRoot)" `
    -PrincipalsAllowedToRetrieveManagedPassword 'SERVER01$' `
    -Description 'Runs the inactive account disable job (CMMC IA.L2-3.5.6)'

# 3. On SERVER01: install and test it
Install-ADServiceAccount -Identity 'gmsa-InactAcct'
Test-ADServiceAccount   -Identity 'gmsa-InactAcct'    # must return True
```

If `Test-ADServiceAccount` returns False, reboot the server, or run `klist purge -li 0x3e7`, so the server picks up its new permission. Then test again.

### Grant "Log on as a batch job"

The gMSA needs this right on SERVER01:

- **If a GPO manages this right** (common in CMMC baselines): add `YOURDOMAIN\gmsa-InactAcct$` in that GPO under **Computer Configuration → Policies → Windows Settings → Security Settings → Local Policies → User Rights Assignment → Log on as a batch job**.
- **Otherwise:** add it locally in `secpol.msc`, at the same path.

Run `gpupdate /force` afterwards.

---

## Step 4 – Delegate AD permissions to the gMSA

The gMSA needs exactly these rights:

| Where | Right | Why |
|---|---|---|
| Every OU it searches | Write `userAccountControl` on user and computer objects | Disable the account |
| Every OU it searches | Write `description` on user and computer objects | Stamp the disable note |
| Every OU it searches | Delete child user and computer objects; write `name` and `cn` | Required to *move* an object out |
| The two Disabled OUs | Create child user and computer objects | Required to *move* an object in |

**Run this once for each top-level source OU.** If `SearchBases` is empty, the source is the domain root, so delegate on your top-level user and computer OUs instead.

```powershell
$gmsa = 'YOURDOMAIN\gmsa-InactAcct$'
$sourceOUs = @(
    'OU=Acutec Users,DC=yourdomain,DC=local'
    'OU=Acutec Computers,DC=yourdomain,DC=local'
)

foreach ($ou in $sourceOUs) {
    foreach ($class in 'user', 'computer') {
        dsacls $ou /I:S /G "${gmsa}:WP;userAccountControl;$class"
        dsacls $ou /I:S /G "${gmsa}:WP;description;$class"
        dsacls $ou /I:S /G "${gmsa}:WP;name;$class"
        dsacls $ou /I:S /G "${gmsa}:WP;cn;$class"
        dsacls $ou /I:T /G "${gmsa}:DC;$class"
    }
}

dsacls 'OU=Disabled Users,DC=yourdomain,DC=local'     /I:T /G "${gmsa}:CC;user"
dsacls 'OU=Disabled Computers,DC=yourdomain,DC=local' /I:T /G "${gmsa}:CC;computer"
```

> **Privileged accounts:** Accounts in protected groups (Domain Admins, Enterprise Admins, Administrators, and so on) have `adminCount=1` and **ignore** delegated permissions. The job will log `ERROR: Access is denied` for those accounts. Don't grant the gMSA rights through AdminSDHolder. Review inactive privileged accounts by hand instead. An assessor will expect that anyway.

Reading `lastLogonTimestamp`, `pwdLastSet` and replication metadata needs no extra rights.

---

## Step 5 – Create the Graph certificate

Run this on SERVER01 in an elevated window:

```powershell
$cert = New-SelfSignedCertificate -Subject 'CN=InactiveAccounts-Graph' `
    -CertStoreLocation 'Cert:\LocalMachine\My' `
    -KeyExportPolicy NonExportable -KeySpec Signature `
    -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(1)

$cert.Thumbprint                                   # record this
Export-Certificate -Cert $cert -FilePath C:\Temp\InactiveAccounts-Graph.cer   # public key only
```

If your organization has an internal PKI, you can request a certificate from it instead.

**Let the gMSA read the private key:**

1. Open `certlm.msc` and go to **Personal → Certificates**.
2. Right-click **InactiveAccounts-Graph** and choose **All Tasks → Manage Private Keys**.
3. Select **Add**, then **Object Types**, and tick **Service Accounts**.
4. Enter `gmsa-InactAcct$` and grant **Read** only.

**Set a calendar reminder about 30 days before the certificate expires.** Renewal steps are in Step 11.

---

## Step 6 – Register the app in Entra (GCC High)

1. Sign in to **https://entra.microsoft.us** as a Global Administrator.
2. Go to **Identity → Applications → App registrations → New registration**.
   - **Name:** `Inactive Account Automation`
   - **Supported account types:** *Accounts in this organizational directory only*
   - Select **Register**.
3. On the **Overview** page, copy the **Application (client) ID**. The tenant ID is `44e9b2c4-4a43-4cac-bdb1-d719bca04d42`.
4. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions** and add:
   - `User.Read.All`
   - `AuditLog.Read.All`
   - `User.EnableDisableAccount.All`, needed only if `IncludeCloudOnlyUsers = $true`
5. Select **Grant admin consent for Acutec Precision Aerospace** and confirm. Every permission must show a green check.
6. Go to **Certificates & secrets → Certificates → Upload certificate** and upload `InactiveAccounts-Graph.cer`.
   **Don't create a client secret.**
7. **Optional cleanup:** under **API permissions**, remove the default delegated `User.Read` permission.

**Limitations to know about:**

- **Cloud-only admins:** An app with `User.EnableDisableAccount.All` can't disable cloud-only accounts that hold admin roles. The job logs those as `ERROR:`. Review them by hand.
- **Privileged app:** Treat this app registration as privileged. Assign it an owner, and include it in your access reviews.

---

## Step 7 – Install and configure the script

```powershell
New-Item -ItemType Directory -Path 'C:\Scripts\InactiveAccounts\Logs' -Force
# Copy Disable-InactiveAccounts.ps1 into C:\Scripts\InactiveAccounts, then:
Unblock-File 'C:\Scripts\InactiveAccounts\Disable-InactiveAccounts.ps1'

# The gMSA may read and run the script, but it must not be able to change it
icacls 'C:\Scripts\InactiveAccounts' /inheritance:r `
    /grant 'BUILTIN\Administrators:(OI)(CI)F' 'NT AUTHORITY\SYSTEM:(OI)(CI)F' `
    /grant 'YOURDOMAIN\gmsa-InactAcct$:(OI)(CI)RX'
icacls 'C:\Scripts\InactiveAccounts\Logs' /grant 'YOURDOMAIN\gmsa-InactAcct$:(OI)(CI)M'
```

Edit the configuration block in the script. It's the same block described in Workstation-Guide.md, section 4. On the server, also set these values:

| Setting | Value |
|---|---|
| `Graph.ClientId` | Application (client) ID from Step 6 |
| `Graph.CertificateThumbprint` | Thumbprint from Step 5 |
| `LogDirectory` | `C:\Scripts\InactiveAccounts\Logs` |
| `Mail.Enabled`, `SmtpServer`, `From`, `To` | Your internal SMTP relay, if you want summary emails. The relay must accept mail from SERVER01. |
| `ExemptCloudObjectIds` | Your break-glass account object IDs. **Required before you enable `-Commit`.** |

**Optional:** if your execution policy is `AllSigned`, sign the script with your code-signing certificate after every edit.

---

## Step 8 – Create the scheduled task (report-only first)

```powershell
$script = 'C:\Scripts\InactiveAccounts\Disable-InactiveAccounts.ps1'
$taskArgs = "-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File `"$script`" -AuthMode Certificate"

$action    = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
                 -Argument $taskArgs -WorkingDirectory 'C:\Scripts\InactiveAccounts'
$trigger   = New-ScheduledTaskTrigger -Daily -At 5:45AM
$principal = New-ScheduledTaskPrincipal -UserId 'YOURDOMAIN\gmsa-InactAcct$' -LogonType Password -RunLevel Limited
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
                 -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskPath '\Acutec\' -TaskName 'Disable Inactive Accounts' `
    -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
    -Description 'CMMC IA.L2-3.5.6 - disables accounts inactive 90+ days. Logs: C:\Scripts\InactiveAccounts\Logs'
```

`-LogonType Password` is correct for a gMSA, even though no password is ever entered.

**Test it now:**

```powershell
Start-ScheduledTask -TaskPath '\Acutec\' -TaskName 'Disable Inactive Accounts'
# Wait a minute or two, then:
Get-ScheduledTaskInfo -TaskPath '\Acutec\' -TaskName 'Disable Inactive Accounts' |
    Select-Object LastRunTime, LastTaskResult
```

`LastTaskResult` is the script's exit code:

| Code | Meaning |
|---|---|
| `0` | Success |
| `1` | Error. Read the newest `.log` file in the Logs folder. |
| `2` | The safety limit stopped the run |

Also turn on **Enable All Tasks History** in Task Scheduler's **Actions** pane.

---

## Step 9 – Rollout

**Phase 1: report-only, 1–2 weeks**

- Leave the task as it is, without `-Commit`.
- Review the daily CSV and refine the exemptions until the candidate list contains only accounts that should be disabled.

**Phase 2: one-time initial cleanup**

The first real run will probably exceed the 25-account safety limit. Run it manually on SERVER01 while signed in as an admin who has AD rights. Local admins can use the certificate.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
cd C:\Scripts\InactiveAccounts
.\Disable-InactiveAccounts.ps1 -AuthMode Certificate                                  # final report-only check
.\Disable-InactiveAccounts.ps1 -AuthMode Certificate -Commit -IgnoreSafetyLimit
```

Keep that commit CSV somewhere safe. It's your record of the cleanup and your guide if anything needs to be undone.

**Phase 3: turn on daily enforcement**

```powershell
$task = Get-ScheduledTask -TaskPath '\Acutec\' -TaskName 'Disable Inactive Accounts'
$task.Actions[0].Arguments = $task.Actions[0].Arguments + ' -Commit'
Set-ScheduledTask -InputObject $task
```

From then on, the job disables accounts as they pass 90 days. The safety limit catches anything unusual, such as a sync outage that makes everyone look inactive.

---

## Step 10 – Evidence for your CMMC assessment (IA.L2-3.5.6)

Keep these together for the assessor:

1. **Policy:** the written statement that defines the inactivity period as 90 days, and the account types it covers.
2. **Mechanism:**
   - this script, including its configuration block
   - a screenshot of the scheduled task and its run history
   - the app registration's permissions
3. **Records:** the daily CSVs. They show every account evaluated, the dates used, and the action taken.
4. **Exceptions:** the documented reason for every exempt OU and every member of `AD-Inactivity-Exempt`.
5. **Sample:** one disabled account that shows the job's description note and its location in the Disabled OU.

---

## Step 11 – Maintenance

| Task | When |
|---|---|
| Review the email summaries and investigate any `ERROR` rows | Daily or weekly |
| Review the exempt OUs, exempt group members and break-glass list | Quarterly |
| Renew the certificate: repeat Step 5, upload the new `.cer` in Step 6, update `CertificateThumbprint`, then remove the old certificate | Before it expires (1 year) |
| Update the Graph modules on the server and your workstation together. Test report-only first. | As needed |
| Archive logs according to your retention policy. `LogRetentionDays = 0` keeps everything. | Yearly |

---

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| Task result `0x80070569` or "logon failure" | The gMSA lacks **Log on as a batch job** (Step 3). |
| Task never starts, or `Test-ADServiceAccount` returns False | SERVER01 isn't in `PrincipalsAllowedToRetrieveManagedPassword`, or hasn't picked up the change yet. Reboot it. |
| `Certificate ... not found` | The thumbprint is wrong, or the certificate isn't in `LocalMachine\My`. |
| Graph `Insufficient privileges` / 403 | Admin consent wasn't granted (Step 6.5), or permissions are missing. |
| `Keyset does not exist` / private key errors | The gMSA lacks Read on the private key (Step 5). |
| `Microsoft Graph returned no users or no sign-in activity data` | The abort is deliberate. Check `AuditLog.Read.All` consent and licensing. |
| `ERROR: Access is denied` on specific accounts | A protected (`adminCount=1`) account, or an OU missing the delegation from Step 4. |
| `The term 'Connect-MgGraph' is not recognized` | The Graph modules were installed `-Scope CurrentUser`. Reinstall them with `-Scope AllUsers`. |
