# Check-InactiveComputers.ps1 - read-only report of stale, ENABLED computer accounts
param(
    [int]$DaysInactive = 90,
    [string]$SearchBase = (Get-ADDomain).DistinguishedName,
    [string]$OutFolder = "C:\Users\matt.mccain\OneDrive - Acutec Precision Aerospace, Inc\Scripts\Disable Accounts"
)

Import-Module ActiveDirectory
$cutoff  = (Get-Date).AddDays(-$DaysInactive)
$OutFile = Join-Path $OutFolder "InactiveComputers_$(Get-Date -Format yyyyMMdd).csv"

$stale = Get-ADComputer -SearchBase $SearchBase -Filter 'Enabled -eq $true' `
    -Properties Enabled, LastLogonDate, PasswordLastSet, whenCreated, OperatingSystem, PrimaryGroupID, ObjectClass |
  Where-Object {
    $_.Enabled -eq $true -and                           # enabled computers only
    $_.ObjectClass -eq 'computer' -and                  # skip gMSAs
    $_.PrimaryGroupID -notin 516, 521 -and              # skip domain controllers
    $_.Name -ne 'AZUREADSSOACC' -and                    # skip Entra Seamless SSO account
    $_.whenCreated -lt $cutoff -and                     # skip newly created computers
    ($null -eq $_.LastLogonDate   -or $_.LastLogonDate   -lt $cutoff) -and
    ($null -eq $_.PasswordLastSet -or $_.PasswordLastSet -lt $cutoff)
  } |
  Select-Object Name, Enabled, OperatingSystem, LastLogonDate, PasswordLastSet, whenCreated,
    @{ n = 'DaysSinceLogon'; e = { if ($_.LastLogonDate) { ((Get-Date) - $_.LastLogonDate).Days } else { 'Never' } } },
    @{ n = 'OU'; e = { ($_.DistinguishedName -split ',', 2)[1] } } |
  Sort-Object LastLogonDate

$stale | Format-Table -AutoSize
Write-Host "$(@($stale).Count) inactive enabled computer(s) found."

try {
    New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null
    $stale | Export-Csv $OutFile -NoTypeInformation -ErrorAction Stop
    Write-Host "Report saved to $OutFile"
} catch {
    Write-Warning "Could not save the report: $($_.Exception.Message)"
}