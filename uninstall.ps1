# uninstall.ps1 — Corina Service Uninstaller (Production)

# Ensure Admin
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "❌ You must run this script as Administrator."
    exit 1
}

$serviceName = "CorinaService"
$installDir  = Join-Path ${env:ProgramFiles} "CorinaService"
$taskName    = "CorinaProdDailyUpdater"

Write-Host "🛑 Uninstalling $serviceName ..."

# Stop service if it exists
if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
    try {
        Write-Host "➡ Stopping service..."
        Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    } catch {
        Write-Warning "⚠ Could not stop $serviceName: $_"
    }

    Write-Host "➡ Deleting service..."
    sc.exe delete $serviceName | Out-Null
    Start-Sleep -Seconds 2
}

# Kill any lingering process
Get-Process careai-corina-service -ErrorAction SilentlyContinue | Stop-Process -Force

# Remove scheduled task
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    try {
        Write-Host "➡ Removing scheduled task $taskName..."
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    } catch {
        Write-Warning "⚠ Could not remove scheduled task: $_"
    }
}

# Remove install directory
if (Test-Path $installDir) {
    try {
        Write-Host "➡ Removing install directory: $installDir"
        Remove-Item -Recurse -Force $installDir
    } catch {
        Write-Warning "⚠ Could not fully delete $installDir: $_"
    }
}

Write-Host "✅ $serviceName has been uninstalled successfully."
