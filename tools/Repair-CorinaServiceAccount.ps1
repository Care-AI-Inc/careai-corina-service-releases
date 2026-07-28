# Repair-CorinaServiceAccount.ps1
#
# One-time remediation for Corina clinics whose Windows service was reset to
# LocalSystem by the v1.3.4 signed-migration installer. Re-binds the service to
# the site's own service account so it can authenticate to SMB/NAS shares again.
#
# Run by the MSP (as Administrator), one push across their managed sites via
# their RMM. The MSP supplies their own service-account credential; CareAI never
# holds it. Idempotent and safe:
#   - acts ONLY when the service is currently LocalSystem (so it never overrides
#     an account an MSP has already fixed or deliberately set to something else);
#   - if the service is already the target account, it exits 0 without touching
#     anything;
#   - the signed daily updater never modifies the service account, so this fix
#     sticks through all future auto-updates.
#
# Examples:
#   .\Repair-CorinaServiceAccount.ps1 -ServiceAccount '.\caregp-uploader'
#       (prompts securely for the password)
#   .\Repair-CorinaServiceAccount.ps1 -ServiceAccount 'CLINIC\svc-corina' -Password 'P@ss' -Instance level-3
#       (non-interactive form for RMM; pass -WhatIf first to preview)

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # Local (.\user) or domain (DOMAIN\user) account the service must run as.
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^(?:\.|[A-Za-z0-9_.-]+)\\[A-Za-z0-9 ._-]+$')]
    [string]$ServiceAccount,

    # Password for that account. Omit to be prompted securely.
    [string]$Password,

    # Tagged/multi-site installs only. Matches the CorinaService-<Instance> name.
    [ValidatePattern('^[A-Za-z0-9](?:[A-Za-z0-9_-]*[A-Za-z0-9])?$')]
    [string]$Instance
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this script as Administrator.'
}

$serviceName = if ($Instance) { "CorinaService-$Instance" } else { 'CorinaService' }

$service = Get-CimInstance Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue
if (-not $service) {
    Write-Host "[skip] Service '$serviceName' is not installed on this machine." -ForegroundColor Yellow
    return
}

$currentAccount = [string]$service.StartName
Write-Host "[info] '$serviceName' currently runs as: $currentAccount"

# Normalise for comparison (LocalSystem has several spellings).
$isLocalSystem = $currentAccount -in @('LocalSystem', 'NT AUTHORITY\SYSTEM', '.\LocalSystem', '')
$targetMatches = $currentAccount.TrimStart('.') -ieq $ServiceAccount.TrimStart('.')

if ($targetMatches) {
    Write-Host "[done] Already running as '$ServiceAccount'. No change needed." -ForegroundColor Green
    return
}
if (-not $isLocalSystem) {
    Write-Host "[skip] Service runs as '$currentAccount', not LocalSystem. Leaving it untouched (nothing to repair, or already set by you)." -ForegroundColor Yellow
    return
}

if (-not $Password) {
    $secure = Read-Host -AsSecureString "Password for $ServiceAccount"
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { $Password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}
if ([string]::IsNullOrEmpty($Password)) { throw 'No password supplied.' }

if (-not $PSCmdlet.ShouldProcess($serviceName, "Set service account to $ServiceAccount and restart")) {
    return
}

# sc.exe escapes an embedded quote/backslash poorly, so pass the account/password
# as native arguments; PowerShell quotes them for us.
Write-Host "[*] Setting '$serviceName' to run as '$ServiceAccount'"
& sc.exe config $serviceName obj= $ServiceAccount password= $Password | Out-Null
if ($LASTEXITCODE -ne 0) { throw "sc.exe config failed (exit $LASTEXITCODE)." }

Write-Host "[*] Restarting the service"
try { Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue } catch { }
Start-Sleep -Seconds 2
try {
    Start-Service -Name $serviceName -ErrorAction Stop
}
catch {
    # Roll the binding back to LocalSystem so a wrong credential does not leave
    # the service unstartable until the next visit.
    & sc.exe config $serviceName obj= LocalSystem | Out-Null
    Start-Service -Name $serviceName -ErrorAction SilentlyContinue
    throw ("Service would not start as '$ServiceAccount' - reverted to LocalSystem. " +
        "Most likely a wrong password, or the account lacks the 'Log on as a service' right. " +
        "Grant it (secpol.msc -> Local Policies -> User Rights Assignment -> Log on as a service) and re-run. Underlying error: $_")
}

$after = Get-CimInstance Win32_Service -Filter "Name='$serviceName'"
if ($after.State -ne 'Running' -or ($after.StartName.TrimStart('.') -ine $ServiceAccount.TrimStart('.'))) {
    throw "Post-change verification failed: state=$($after.State) account=$($after.StartName)."
}

Write-Host "[OK] '$serviceName' is running as '$ServiceAccount'. Share access should be restored within a minute." -ForegroundColor Green
