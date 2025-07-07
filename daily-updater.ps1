# daily-updater.ps1 — for production

if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "❌ You must run this script as Administrator."
    exit 1
}

$logPath = "C:\Scripts\corina-prod-update-log.txt"
"[$(Get-Date)] Running production updater..." | Out-File -Append $logPath

try {
    irm https://raw.githubusercontent.com/Care-AI-Inc/careai-corina-service-production-releases/main/install.ps1 | iex
    "[$(Get-Date)] ✅ Update succeeded" | Out-File -Append $logPath
} catch {
    "[$(Get-Date)] ❌ Update failed: $_" | Out-File -Append $logPath
}
