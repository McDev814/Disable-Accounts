# Check-InactiveComputers.ps1 – How to Run It

This script finds **enabled** computer accounts that haven't authenticated to Active Directory in **90 days**. It writes a report, and if you ask it to, it disables those computers and moves them to the **Auto Disabled** OU.

| Command | What happens |
|---|---|
| `.\Check-InactiveComputers.ps1` | **Report only.** Writes a CSV to `C:\Reports`. Nothing in AD changes. |
| `.\Check-InactiveComputers.ps1 -Commit` | Disables each inactive computer, stamps its description, and moves it to **Auto Disabled**. Asks you to type `YES` first. |
| `.\Check-InactiveComputers.ps1 -Commit -Force` | Same as above, without the prompt. Use this for the scheduled task. |

### How a computer counts as inactive

A computer is inactive when **both** of these dates are more than 90 days ago, or empty:

- `LastLogonDate`, the last time the computer authenticated to a DC
- `PasswordLastSet`, the last time Windows changed the machine password, which happens about every 30 days

### What it skips

The script always skips:

- disabled computers
- domain controllers
- managed service accounts
- `AZUREADSSOACC`
- computers created in the last 90 days
- anything in `-ExcludeOUs` or `-ExcludeComputers`

### Optional parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-DaysInactive` | `90` | Inactivity threshold |
| `-SearchBase` | whole domain | Limit the search to one OU |
| `-ExcludeOUs` | `OU=Servers,OU=Acutec,...` | OUs to skip, including all their sub-OUs |
| `-ExcludeComputers` | none | Individual computer names to skip |
| `-DisabledOUName` | `Auto Disabled` | Name of the destination OU |
| `-ReportFolder` | `C:\Reports` | Where the CSV is written |
| `-Credential` | current user | AD account used to make changes |

---

## Part 1 – One-time setup (do this once, from either machine)

### 1. Create the Auto Disabled OU

Create it once. The script finds it by name, so it can live anywhere in the domain. The only rule is that there must be exactly one OU with that name.

```powershell
New-ADOrganizationalUnit -Name 'Auto Disabled' -Path 'OU=Acutec,DC=Acutec,DC=local' `
    -Description 'Computers disabled automatically for 90+ days of inactivity (CMMC IA.L2-3.5.6)' `
    -ProtectedFromAccidentalDeletion $true
```

**Optional:** link a GPO to this OU, or leave GPOs off it, depending on how you want disabled machines treated.

### 2. Review the exclusions

Open the script and look at the `ExcludeOUs` and `ExcludeComputers` lists near the top.

**What's excluded now:** `OU=Servers,OU=Acutec,DC=Acutec,DC=local` and every OU below it, including Dev Servers, HyperV, RDS, DMZ and Non-Windows. That removes about 25 of the 85 computers in your last report.

**Servers that aren't under `OU=Servers`, so they're still in scope:**

- `VEEAM-SCESX01B` (CN=Computers)
- `ESXI2HYPERV` (CN=Computers)
- `SLDC` (CN=Computers)
- `VEEAM-PROXY4` (CN=Computers)
- `RADIUSTESTVM` (OU=TestOU)

For each one, either move it into the right Servers OU or add its name to `ExcludeComputers`.

Use just the OU name (`'OU=Kiosk Computers'`) to match that OU anywhere in the domain. Use a full DN to match exactly one OU.

### 3. Unblock the file

Do this on each machine you copy the script to. Windows blocks downloaded scripts until they're unblocked.

```powershell
Unblock-File .\Check-InactiveComputers.ps1
```

---

## Part 2 – Running it from your workstation

**Prerequisite:** the RSAT Active Directory module. You already have it. If you ever need to reinstall it, run this as administrator:

```powershell
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

**Create the report folder.** Run this once as administrator, because `C:\` needs admin rights to create folders:

```powershell
New-Item -ItemType Directory -Path C:\Reports -Force
icacls C:\Reports /grant 'ACUTEC\matt.mccain:(OI)(CI)M'    # your regular account
```

**Report only.** Open a normal (non-elevated) PowerShell window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
cd "C:\Users\matt.mccain\OneDrive - Acutec Precision Aerospace, Inc\Scripts\Disable Accounts"
.\Check-InactiveComputers.ps1
```

**Commit.** Your regular account can read AD but can't disable or move computers, so pass your admin credentials:

```powershell
$cred = Get-Credential -Message 'AD admin account (ACUTEC\username)'
.\Check-InactiveComputers.ps1 -Commit -Credential $cred
```

The script lists the computers, then waits for you to type `YES`, in capitals. Anything else cancels the run without changing anything.

---

## Part 3 – Running it from the Falcon Log Collector server

### Before you start

The Falcon Log Collector server works for this, with two things to keep in mind:

- **It becomes a server that can change AD.** Run the job under a **gMSA** that has only the computer-account rights below. Don't use a domain admin account, and limit who can sign in to the server.
- **Its actions can go to Falcon LogScale.** You can send the job's own events to the Application log, and the collector can forward them alongside your other logs. See step 6.

### Step 1 – Install the AD module

Run this on the server as administrator:

```powershell
Install-WindowsFeature RSAT-AD-PowerShell
```

### Step 2 – Create the folders

```powershell
New-Item -ItemType Directory -Path C:\Scripts\InactiveComputers, C:\Reports -Force
# Copy Check-InactiveComputers.ps1 into C:\Scripts\InactiveComputers, then:
Unblock-File C:\Scripts\InactiveComputers\Check-InactiveComputers.ps1
```

### Step 3 – Create a gMSA for the job

Replace `FALCONLC01` with the server's computer name.

```powershell
# One KDS root key per forest. If this returns nothing, run Add-KdsRootKey -EffectiveImmediately
# and wait up to 10 hours before continuing.
Get-KdsRootKey

New-ADServiceAccount -Name 'gmsa-InactComp' `
    -DNSHostName 'gmsa-InactComp.Acutec.local' `
    -PrincipalsAllowedToRetrieveManagedPassword 'FALCONLC01$' `
    -Description 'Runs Check-InactiveComputers.ps1 (CMMC IA.L2-3.5.6)'

# On the Falcon server:
Install-ADServiceAccount -Identity 'gmsa-InactComp'
Test-ADServiceAccount   -Identity 'gmsa-InactComp'     # must return True (reboot the server if False)
```

Then give the gMSA the **Log on as a batch job** right on the server:

- **If a GPO controls this right:** add `ACUTEC\gmsa-InactComp$` in that GPO.
- **Otherwise:** add it locally in `secpol.msc` under **Local Policies → User Rights Assignment → Log on as a batch job**.

Run `gpupdate /force` afterwards.

### Step 4 – Give the gMSA only the rights it needs

```powershell
$gmsa = 'ACUTEC\gmsa-InactComp$'

# Where computers live today: disable, stamp the description, and move them out
foreach ($ou in 'OU=Acutec,DC=Acutec,DC=local', 'CN=Computers,DC=Acutec,DC=local') {
    dsacls $ou /I:S /G "${gmsa}:WP;userAccountControl;computer"
    dsacls $ou /I:S /G "${gmsa}:WP;description;computer"
    dsacls $ou /I:S /G "${gmsa}:WP;name;computer"
    dsacls $ou /I:S /G "${gmsa}:WP;cn;computer"
    dsacls $ou /I:T /G "${gmsa}:DC;computer"
}

# Where they go: allow computers to be moved in
dsacls 'OU=Auto Disabled,OU=Acutec,DC=Acutec,DC=local' /I:T /G "${gmsa}:CC;computer"

# Let the gMSA write reports
icacls C:\Reports /grant "${gmsa}:(OI)(CI)M"
icacls C:\Scripts\InactiveComputers /grant "${gmsa}:(OI)(CI)RX"
```

If `Auto Disabled` sits somewhere other than `OU=Acutec`, change its path in the `dsacls` line above.

### Step 5 – Create the scheduled task

Start in **report-only** mode:

```powershell
$script    = 'C:\Scripts\InactiveComputers\Check-InactiveComputers.ps1'
$taskArgs  = "-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File `"$script`""

$action    = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
                 -Argument $taskArgs -WorkingDirectory 'C:\Scripts\InactiveComputers'
$trigger   = New-ScheduledTaskTrigger -Daily -At 6:10AM
$principal = New-ScheduledTaskPrincipal -UserId 'ACUTEC\gmsa-InactComp$' -LogonType Password -RunLevel Limited
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
                 -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask -TaskPath '\Acutec\' -TaskName 'Check Inactive Computers' `
    -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
    -Description 'Disables computers inactive 90+ days (CMMC IA.L2-3.5.6). Reports: C:\Reports'

# Test it now
Start-ScheduledTask -TaskPath '\Acutec\' -TaskName 'Check Inactive Computers'
Get-ScheduledTaskInfo -TaskPath '\Acutec\' -TaskName 'Check Inactive Computers' | Select-Object LastRunTime, LastTaskResult
```

`LastTaskResult` of `0` means success. Check that a new CSV appeared in `C:\Reports`.

**Switch on enforcement when you're happy with the report.** The task can't answer the prompt, so `-Force` is required:

```powershell
$task = Get-ScheduledTask -TaskPath '\Acutec\' -TaskName 'Check Inactive Computers'
$task.Actions[0].Arguments += ' -Commit -Force'
Set-ScheduledTask -InputObject $task
```

For the **first** cleanup of roughly 60 computers, consider running `-Commit` by hand from the server or your workstation, so you can watch it. Turn on the scheduled `-Commit` after that.

### Step 6 – Windows Event Log and Falcon LogScale (optional)

Register the event source once, as administrator on the server:

```powershell
New-EventLog -LogName Application -Source 'Check-InactiveComputers'
```

The script then writes these events to the **Application** log:

| Event ID | Level | Written when |
|---|---|---|
| 1000 | Information or Warning | Run summary: counts and report path |
| 1001 | Information | A computer was disabled and moved |
| 1002 | Error | A computer failed to disable |

If the Log Collector on this server doesn't already read the Application log, add it to the collector's `config.yaml`. Match the sink name to your existing config:

```yaml
sources:
  windows_application:
    type: wineventlog
    channels:
      - name: Application
    sink: logscale
```

You can then search or alert in LogScale on `SourceName=Check-InactiveComputers`, for example to alert on event 1002. Your DCs also log their own events when computer accounts are changed or moved, such as 4742 and, if directory service change auditing is on, 5139. Those reach LogScale if you already collect DC Security logs.

---

## Undoing a disable

The CSV from the `-Commit` run has `OriginalOU` and `OriginalDescription` for every computer.

```powershell
$name = 'L-SPARE06'
$row  = Import-Csv 'C:\Reports\InactiveComputers_<date>_Commit.csv' | Where-Object Name -eq $name

$c = Get-ADComputer $name
Enable-ADAccount -Identity $c
Set-ADComputer   -Identity $c -Description $row.OriginalDescription
Move-ADObject    -Identity $c -TargetPath $row.OriginalOU
```

If you're running this from your workstation, add `-Credential $cred` to the three AD commands.

The re-enabled computer needs to connect to the network and sign in to a DC soon. Otherwise the next run will catch it again, because its dates are still old. If it's going to stay offline, add its name to `ExcludeComputers`.

---

## Troubleshooting

| Message | Fix |
|---|---|
| `running scripts is disabled on this system` | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned` first. |
| `...is not digitally signed` | Run `Unblock-File` on the script. |
| `Cannot create or access the report folder 'C:\Reports'` | Create the folder and grant write access (Part 2 or Part 3, step 4). |
| `Expected exactly one OU named 'Auto Disabled' but found 0` | Create the OU (Part 1). If it says 2 or more, rename or delete the extra one. |
| `FAILED <name>: Access is denied` | The account running the script lacks the rights from Part 3, step 4. On the workstation, use `-Credential`. |
| Task result `0x80070569` | The gMSA is missing **Log on as a batch job**. |
