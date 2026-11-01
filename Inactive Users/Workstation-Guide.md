# Disable-InactiveAccounts.ps1 – Workstation Guide

Use this guide to run the script from your own workstation. Report-only runs are the main use. You can also use it, with care, for the one-time initial cleanup.

For the scheduled production job, see **Server-Guide.md**.

---

## 1. What the script does

For every enabled account in scope, the script finds the **most recent activity** from two sources and uses the **later** date:

| Account type | Activity sources |
|---|---|
| AD users (synced or not) | AD `LastLogonDate` **and** Entra `lastSuccessfulSignInDateTime` |
| AD computers | AD `LastLogonDate` **and** `PasswordLastSet` |
| Entra cloud-only users and guests | Entra `lastSuccessfulSignInDateTime` |

If that date is more than **90 days** old, the account is a **Candidate**.

Accounts that have never signed in anywhere are judged by their creation date. They are left alone for 90 days after they are created.

**The script is safe by default:**

- **Report-only unless you add `-Commit`.** Without `-Commit`, nothing is changed.
- **Stops if Graph data is missing.** If Microsoft Graph returns no sign-in data, the run aborts instead of judging users on AD data alone. AD-only data would have flagged about 421 active users.
- **Safety limit.** A `-Commit` run that would disable more than 25 accounts stops without changing anything.
- **Re-enable grace period.** An account that was re-enabled in the last 14 days is skipped, which gives the person time to sign in.
- **Evidence for every run.** Each run writes a CSV listing every account evaluated, plus a transcript `.log`.

---

## 2. One-time prerequisites

Open **PowerShell as administrator** for this section only.

### 2a. AD PowerShell module (RSAT)

```powershell
Get-Module -ListAvailable ActiveDirectory
# If nothing is returned:
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

### 2b. Microsoft Graph modules

You already have `Microsoft.Graph.Authentication` 2.41.1. The script also needs `Microsoft.Graph.Users`, and the two should be the same version.

```powershell
Get-Module -ListAvailable Microsoft.Graph.Authentication, Microsoft.Graph.Users |
    Select-Object Name, Version

# If Microsoft.Graph.Users is missing or a different version:
Install-Module Microsoft.Graph.Users -RequiredVersion 2.41.1 -Scope AllUsers
```

Close the administrator window. **Everything from here on runs in a normal, non-elevated PowerShell window.** An elevated window can hang at the Graph sign-in.

---

## 3. Put the script in place

1. Create a folder, for example `C:\Scripts\InactiveAccounts`, and save `Disable-InactiveAccounts.ps1` there.
2. Unblock the file. Windows marks downloaded files, and `RemoteSigned` blocks them until you do:

   ```powershell
   Unblock-File C:\Scripts\InactiveAccounts\Disable-InactiveAccounts.ps1
   ```

---

## 4. Edit the configuration block

Open the script in VS Code or the PowerShell ISE. Edit **only** the section between
`#region ===== CONFIGURATION` and `#endregion`.

| Setting | What to set |
|---|---|
| `SearchBases` | Leave empty to search the whole domain, or list the OUs that hold users and computers. |
| `DisabledUsersOU` / `DisabledComputersOU` | Full DN of the OUs that disabled accounts move to. Replace `DC=yourdomain,DC=local` with your domain. Set to `''` to disable without moving. |
| `ExemptOUs` | Already contains `OU=Shared Inbox`, `OU=CMM Logins` and `OU=Service Accounts`. Add more as you find them. A bare `OU=Name` matches that OU anywhere in the domain. A full DN matches only that one OU. Sub-OUs are covered either way. |
| `ExemptGroup` | Name of an AD group whose members are never disabled. **The group must exist.** Create it (see below) or set this to `''`. |
| `ExemptSamAccountNames` | Individual accounts to exempt. Computer names end in `$`. |
| `ExemptCloudObjectIds` / `ExemptCloudUPNs` | **Add your break-glass and emergency-access accounts here.** The script warns you until you do. |
| `LogDirectory` | A folder you can write to. On a workstation, `"$env:USERPROFILE\Documents\InactiveAccounts\Logs"` works well. |

> Note: you had `OU-CMM Logins` with a hyphen in an earlier message. The script uses `OU=CMM Logins`. Check the real OU name in ADUC (View → Advanced Features → the OU's Properties → Attribute Editor → `distinguishedName`).

To create the exemption group (this needs an account with rights in AD):

```powershell
New-ADGroup -Name 'AD-Inactivity-Exempt' -GroupScope Global -GroupCategory Security `
    -Path 'OU=Groups,DC=yourdomain,DC=local' `
    -Description 'Members are exempt from automatic inactivity disable (CMMC IA.L2-3.5.6). Document the reason for each member.'
```

---

## 5. Run in report-only mode

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
cd C:\Scripts\InactiveAccounts
.\Disable-InactiveAccounts.ps1
```

- A browser tab opens for the Graph sign-in. Sign in as **`mmccain@acutec.onmicrosoft.us`**. If the browser signs you in as `matt.mccain` automatically, choose **Use another account**.
- Your regular Windows account is enough for the AD side, because reading logon data doesn't need admin rights.
- When the run finishes, it prints a summary table and the path to the CSV.

---

## 6. Review the CSV

Open the newest `InactiveAccounts_<date>_ReportOnly.csv` in Excel, then select **Data → Filter**.

| Column | Meaning |
|---|---|
| `Status` | `Active`, `Exempt`, `Skipped` (re-enable grace), or `Candidate` |
| `Action` | `WouldDisable` in report-only mode |
| `EntraLastSuccessfulSignIn` / `ADLastLogon` / `ADPasswordLastSet` | The raw dates |
| `EffectiveLastActivity` | The later of the dates, which is what the decision is based on |
| `Reason` | Why the account got its status |

Filter `Status = Candidate` and check:

1. **No real, working people are listed.** If someone is, find out why both sources show them as inactive before you go further.
2. **Service, shared, kiosk and machine-login accounts aren't listed.** If one is, add its OU to `ExemptOUs`, or add the account to the exempt group.
3. **Break-glass accounts show `Exempt`.**
4. **The totals are roughly what we expect.** The earlier analysis predicts up to about 230 synced users, before exemptions, plus a handful of cloud-only accounts and some computers.

Adjust the configuration and re-run until the candidate list is correct. Report-only runs are free to repeat.

---

## 7. Initial cleanup from the workstation (optional)

The first `-Commit` run will probably exceed the 25-account safety limit. The preferred approach is to do this cleanup on the server (see Server-Guide.md, Phase 2). If you do it from your workstation instead:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
cd C:\Scripts\InactiveAccounts

# An AD account that can disable and move users and computers
$cred = Get-Credential -Message 'AD admin account (DOMAIN\username)'

.\Disable-InactiveAccounts.ps1 -Commit -ADCredential $cred -IgnoreSafetyLimit
```

- **Extra Graph permission.** The Graph sign-in also asks for `User.EnableDisableAccount.All`, so the script can disable cloud-only accounts. Approve it with your Global Admin account.
- **Re-check right before committing.** Run report-only again immediately before the commit, because the candidate list changes daily.
- **Expected errors on privileged accounts.** Protected accounts, such as members of Domain Admins, may fail with "Access is denied" if `$cred` lacks rights. They show as `ERROR:` in the `Action` column. Handle them by hand.

---

## 8. Re-enable an account that was disabled by mistake

The commit CSV records each account's original OU and description.

```powershell
$dn = 'CN=Jane Doe,OU=Disabled Users,DC=yourdomain,DC=local'       # where it is now
$originalOU = 'OU=Staff,DC=yourdomain,DC=local'                     # OriginalOU column in the CSV
$originalDescription = 'Machinist - 2nd shift'                       # OriginalDescription column

Enable-ADAccount -Identity $dn -Credential $cred
Set-ADUser       -Identity $dn -Description $originalDescription -Credential $cred
Move-ADObject    -Identity $dn -TargetPath $originalOU -Credential $cred
```

**After re-enabling:**

- **Grace period.** The re-enabled account has **14 days** to sign in before it can be flagged again.
- **If it should never be flagged again,** add it to `AD-Inactivity-Exempt` and record why.
- **Cloud-only account:** run `Update-MgUser -UserId <ObjectId> -AccountEnabled:$true`. The `ObjectId` column in the CSV has the ID.

---

## 9. Troubleshooting

| Symptom | Fix |
|---|---|
| `running scripts is disabled on this system` | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned` in that window. |
| `...is not digitally signed` | Run `Unblock-File` on the script (step 3). |
| Window freezes at sign-in, no browser opens | You're in an elevated window. Use a normal window. The script already turns off the Windows sign-in broker. |
| `User is not in the allowed roles` (403) | You signed in to Graph as `matt.mccain`. Sign in as `mmccain@acutec.onmicrosoft.us`. |
| `Microsoft Graph returned no users or no sign-in activity data` | The account lacks `AuditLog.Read.All` consent or a reporting role. The run aborted on purpose. |
| `Cannot find an object with identity: 'AD-Inactivity-Exempt'` | Create the group, or set `ExemptGroup = ''`. |
| `Directory object not found` on a Disabled OU (only with `-Commit`) | Create the OU, or fix its DN in the configuration. |
| Exit code `2` | The safety limit stopped the run. Review the CSV, then use `-IgnoreSafetyLimit` only if the list is correct. |
