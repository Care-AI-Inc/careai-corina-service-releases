# install.ps1 (Production Installer)

# Ensure Admin
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "❌ You must run this script as Administrator."
    exit 1
}

Write-Host "✅ Running as Administrator"

# Get latest production release from GitHub
$repo = "Care-AI-Inc/careai-corina-service-releases"
$apiUrl = "https://api.github.com/repos/$repo/releases/latest"
$headers = @{ "User-Agent" = "CorinaServiceInstaller" }

try {
    $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers
    $latestTag = $response.tag_name
    $zipAsset = $response.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1
    $zipUrl = $zipAsset.browser_download_url
    $zipName = $zipAsset.name
} catch {
    Write-Error "❌ Failed to fetch release or asset info from GitHub"
    exit 1
}

Write-Host "⬇ Downloading $zipName from $zipUrl"

# Download the ZIP
$zipPath = "$env:TEMP\$zipName"
Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath

# Define install path and service name
$installDir = Join-Path ${env:ProgramFiles} "CorinaService"
$serviceName = "CorinaService"

# Stop and remove existing service if running
if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
    Write-Host "🛑 Stopping existing service..."
    Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Write-Host "🧹 Deleting existing service..."
    sc.exe delete $serviceName | Out-Null
    Start-Sleep -Seconds 2

    # Kill any lingering process
    Get-Process careai-corina-service -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 1
}

# Remove old install dir
if (Test-Path $installDir) {
    try {
        Write-Host "🧼 Removing old install directory: $installDir"
        Remove-Item -Recurse -Force $installDir
    } catch {
        Write-Warning "⚠️ Could not fully delete $installDir, retrying in 5 seconds..."
        Start-Sleep -Seconds 5
        Remove-Item -Recurse -Force $installDir -ErrorAction SilentlyContinue
    }
}

# Extract new version
Expand-Archive -Path $zipPath -DestinationPath $installDir

# Install as Windows Service
$exePath = Join-Path $installDir "careai-corina-service.exe"

if (-not (Test-Path $exePath)) {
    Write-Error "❌ Failed to find service executable at $exePath"
    exit 1
}

# Remove old service if exists
if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
    Stop-Service -Name $serviceName -Force
    sc.exe delete $serviceName | Out-Null
    Start-Sleep -Seconds 2
}

# Register service
sc.exe create $serviceName binPath= "`"$exePath`"" start= auto obj= "LocalSystem" DisplayName= "Corina Service (Production)"

# Set recovery options for Production (same as Staging)
Write-Host "🔁 Configuring service recovery options for Production..."
sc.exe failure CorinaService reset= 86400 actions= restart/5000/restart/5000/restart/5000 | Out-Null
sc.exe failureflag CorinaService 1 | Out-Null
Write-Host "✅ Service will auto-restart on failure (3x retries, 5s wait, reset every 1 day)"

# Start service
Start-Service -Name $serviceName

Write-Host "🎉 Corina Service (Production) installed and started successfully!"

# === [ Setup Dynamic Daily Auto-Updater - Production ] ===
$scriptDir = "C:\Scripts"
$shimPath = "$scriptDir\run-daily-updater-prod.ps1"
$taskName = "CorinaProdDailyUpdater"

# Ensure script directory exists
if (-not (Test-Path $scriptDir)) {
    New-Item -ItemType Directory -Path $scriptDir | Out-Null
}

# Write shim script that always fetches latest updater
@'
# run-daily-updater-prod.ps1
try {
    Invoke-Expression (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Care-AI-Inc/careai-corina-service-releases/main/daily-updater.ps1" -UseBasicParsing).Content
} catch {
    "`n[$(Get-Date)] ❌ Failed to fetch and run latest prod updater: $_" | Out-File -Append "C:\Scripts\corina-prod-update-log.txt"
}
'@ | Set-Content -Path $shimPath -Encoding UTF8

# Define task components
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File `"$shimPath`""
$trigger1 = New-ScheduledTaskTrigger -Daily -At 7am
$trigger2 = New-ScheduledTaskTrigger -Daily -At 9am
$trigger3 = New-ScheduledTaskTrigger -Daily -At 11am
$trigger4 = New-ScheduledTaskTrigger -Daily -At 1pm
$trigger5 = New-ScheduledTaskTrigger -Daily -At 3pm
$trigger6 = New-ScheduledTaskTrigger -Daily -At 5pm
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Remove old task if needed
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Start-Sleep -Seconds 1
}

# Register the new production auto-updater task
Register-ScheduledTask -TaskName $taskName `
    -Action $action `
    -Trigger $trigger1, $trigger2, $trigger3, $trigger4, $trigger5, $trigger6 `
    -Principal $principal

Write-Host "📅 Scheduled task '$taskName' created with 6 daily triggers."
