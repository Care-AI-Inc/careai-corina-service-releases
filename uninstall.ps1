# Corina Service production uninstaller.
# Final release copies are Authenticode-signed after all build substitutions.
# The repository source intentionally contains no stale signature block.

[CmdletBinding()]
param(
    [ValidatePattern('^[A-Za-z0-9](?:[A-Za-z0-9_-]*[A-Za-z0-9])?$')]
    [string]$Instance,
    [string[]]$TrustedSignerThumbprints
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$builtInTrustedSignerThumbprints = @('__CORINA_RELEASE_SIGNER_THUMBPRINTS__')
$script:CorinaServiceBaseName = 'CorinaService'
$script:CorinaTaskBaseName = 'CorinaProdDailyUpdater'
$script:CorinaProgramFilesLeaf = 'CorinaService'
$script:CorinaProgramDataRoot = 'CareAI\CorinaService'
$script:CorinaRegistryRoot = 'HKLM:\SOFTWARE\CareAI\CorinaService'
$script:CorinaLegacyShimBaseName = 'run-daily-updater-prod'

# Assign the fallback to a separate variable: writing $null back into the
# [ValidatePattern] parameter variable re-triggers validation and always throws
# on default-instance machines where the environment variable is unset.
$corinaRegistryInstance = $Instance
if ([string]::IsNullOrWhiteSpace($corinaRegistryInstance)) { $corinaRegistryInstance = [Environment]::GetEnvironmentVariable('CorinaRegistryInstance', [EnvironmentVariableTarget]::Process) }
if ($corinaRegistryInstance -and $corinaRegistryInstance -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9_-]*[A-Za-z0-9])?$') { throw 'Invalid Corina registry instance.' }
$registryPath = $script:CorinaRegistryRoot
if ($corinaRegistryInstance) { $registryPath = Join-Path $registryPath $corinaRegistryInstance }

# Wrapping this parameter directly in @( ) is NOT a safe array-wrap: an unbound
# typed [string[]] parameter is a null whose array wrap stays $null, so the
# .Count below threw under Set-StrictMode and made every bare `.\uninstall.ps1`
# fail before the registry and built-in trust fallbacks could be reached.
$trustInput = @()
if ($null -ne $TrustedSignerThumbprints) {
    $trustInput = @($TrustedSignerThumbprints | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
}
if ($trustInput.Count -eq 0 -and (Test-Path -LiteralPath $registryPath)) {
    # Blank entries must not count as trust state, or a registry value holding an
    # empty string would suppress the built-in fallback and block the uninstall.
    try { $trustInput = @((Get-ItemPropertyValue -LiteralPath $registryPath -Name TrustedSignerThumbprints -ErrorAction Stop) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) }
    catch { $trustInput = @() }
}
if ($trustInput.Count -eq 0) { $trustInput = $builtInTrustedSignerThumbprints }
$invalidTrust = @($trustInput | Where-Object {
    -not [string]::IsNullOrWhiteSpace([string]$_) -and
    ([string]$_).Replace(' ', '').ToUpperInvariant() -notmatch '^[0-9A-F]{40}$' -and
    [string]$_ -ne '__CORINA_RELEASE_SIGNER_THUMBPRINTS__'
})
if ($invalidTrust.Count -gt 0) { throw 'A trusted signer thumbprint is invalid.' }
$trusted = @($trustInput | ForEach-Object { ([string]$_).Replace(' ', '').ToUpperInvariant() } | Where-Object { $_ -match '^[0-9A-F]{40}$' } | Select-Object -Unique)
if ($trusted.Count -eq 0) { throw 'No trusted signer thumbprint is available; uninstall is blocked (fail closed).' }

if ([string]::IsNullOrWhiteSpace($PSCommandPath)) { throw 'uninstall.ps1 must run from a signed file on disk.' }
$signature = Get-AuthenticodeSignature -FilePath $PSCommandPath
if ($signature.Status -ne 'Valid' -or -not $signature.SignerCertificate) { throw "Uninstaller Authenticode signature is not valid: $($signature.Status)." }
$certificate = $signature.SignerCertificate
$thumbprint = $certificate.Thumbprint.Replace(' ', '').ToUpperInvariant()
if ($thumbprint -notin $trusted) { throw 'Uninstaller signer thumbprint is not trusted.' }
$simpleName = ($certificate.GetNameInfo([Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false) -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
if ($simpleName -cne 'CAREAIPTYLTD') { throw 'Uninstaller publisher common name mismatch.' }
$subject = (([string]$certificate.Subject) -replace '[^A-Za-z0-9=,.]', '').ToUpperInvariant()
if ($subject -notmatch '(?:^|,)O=CAREAIPTYLTD(?:,|$)' -or
    $subject -notmatch '(?:^|,)C=AU(?:,|$)' -or
    $subject -notmatch '(?:^|,)(?:SERIALNUMBER|OID\.2\.5\.4\.5)=38681904512(?:,|$)') {
    throw 'Uninstaller publisher identity mismatch.'
}
$hasCodeSigningEku = $false
foreach ($extension in $certificate.Extensions) {
    if ($extension.Oid.Value -eq '2.5.29.37') {
        $eku = [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]$extension
        if ($eku.EnhancedKeyUsages | Where-Object { $_.Value -eq '1.3.6.1.5.5.7.3.3' }) { $hasCodeSigningEku = $true }
    }
}
if (-not $hasCodeSigningEku) { throw 'Uninstaller certificate lacks the code-signing EKU.' }

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'You must run uninstall.ps1 as Administrator.' }

$serviceName = if ($corinaRegistryInstance) { "$($script:CorinaServiceBaseName)-$corinaRegistryInstance" } else { $script:CorinaServiceBaseName }
$taskName = if ($corinaRegistryInstance) { "$($script:CorinaTaskBaseName)-$corinaRegistryInstance" } else { $script:CorinaTaskBaseName }
$installDir = if ($corinaRegistryInstance) { Join-Path (Join-Path $env:ProgramFiles $script:CorinaProgramFilesLeaf) $corinaRegistryInstance } else { Join-Path $env:ProgramFiles $script:CorinaProgramFilesLeaf }
$stateRoot = if ($corinaRegistryInstance) { Join-Path (Join-Path $env:ProgramData $script:CorinaProgramDataRoot) $corinaRegistryInstance } else { Join-Path (Join-Path $env:ProgramData $script:CorinaProgramDataRoot) 'default' }

Write-Host "[*] Uninstalling $serviceName"
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "    -> Removed scheduled task $taskName"
}

$service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
if ($service) {
    Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    try {
        $cimService = Get-CimInstance Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue
        if ($cimService -and $cimService.ProcessId -gt 0) { Stop-Process -Id $cimService.ProcessId -Force -ErrorAction SilentlyContinue }
    } catch { }
    & sc.exe delete $serviceName | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "sc.exe could not delete '$serviceName' (exit $LASTEXITCODE)." }
    Write-Host "    -> Removed Windows service $serviceName"
}

# The signed uninstaller ships inside the tree it deletes, so an administrator
# who ran it from that folder holds the directory open as their working
# directory and Remove-Item fails with "because it is in use". Step out first.
Set-Location -LiteralPath "$env:SystemDrive\"
if (Test-Path -LiteralPath $installDir) {
    Remove-Item -LiteralPath $installDir -Recurse -Force
    Write-Host "    -> Removed $installDir"
}
if (Test-Path -LiteralPath $stateRoot) {
    Remove-Item -LiteralPath $stateRoot -Recurse -Force
    Write-Host "    -> Removed updater cache, logs and rollback data at $stateRoot"
}

$legacyShim = if ($corinaRegistryInstance) { "C:\Scripts\$($script:CorinaLegacyShimBaseName)-$corinaRegistryInstance.ps1" } else { "C:\Scripts\$($script:CorinaLegacyShimBaseName).ps1" }
if (Test-Path -LiteralPath $legacyShim -PathType Leaf) {
    Remove-Item -LiteralPath $legacyShim -Force
    Write-Host "    -> Removed obsolete downloader shim $legacyShim"
}

# Registry enrolment/token state is deliberately retained so a reinstall does not
# silently destroy clinic identity. Remove it only through an explicit data-reset
# procedure.
Write-Host "SUCCESS: $serviceName was uninstalled. Registry enrolment state was retained."
