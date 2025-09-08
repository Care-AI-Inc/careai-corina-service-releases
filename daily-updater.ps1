# daily-updater-prod.ps1 — for production

# Ensure Admin
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "❌ You must run this script as Administrator."
    exit 1
}

# Log path
$logPath = "C:\Scripts\corina-prod-update-log.txt"
"[$(Get-Date)] 🔄 Starting Corina Production update..." | Out-File -Append $logPath

# GitHub release info
$repo = "Care-AI-Inc/careai-corina-service-releases"
$apiUrl = "https://api.github.com/repos/$repo/releases/latest"
$headers = @{ "User-Agent" = "CorinaProdUpdater" }

try {
    $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers
    $zipAsset = $response.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1
    $zipUrl = $zipAsset.browser_download_url
    $zipName = $zipAsset.name

    # Use ProgramData instead of Windows\Temp to avoid ACL/AV issues
    $tempZip   = "C:\ProgramData\CorinaService\latest.zip"
    $extractDir = "C:\ProgramData\CorinaService\Extract"
    
    # Ensure dirs exist
    New-Item -ItemType Directory -Path (Split-Path $tempZip) -Force | Out-Null
    if (Test-Path $extractDir) { Remove-Item -Recurse -Force $extractDir }
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    
    # Download and extract
    Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip
    Expand-Archive -Path $tempZip -DestinationPath $extractDir -Force

    # Stop service and wait until process fully exits
    $serviceName = "CorinaService"
    $procName = "careai-corina-service"
    
    if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
        Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
    }
    
    # Wait up to 30s for the process to release the DLL
    $timeout = [DateTime]::UtcNow.AddSeconds(30)
    while ((Get-Process $procName -ErrorAction SilentlyContinue) -and ([DateTime]::UtcNow -lt $timeout)) {
        Write-Host "⏳ Waiting for $procName to exit..."
        Start-Sleep -Seconds 1
    }
    
    # Overwrite files once unlocked
    $installDir = Join-Path ${env:ProgramFiles} "CorinaService"
    Copy-Item -Path "$extractDir\*" -Destination $installDir -Recurse -Force

    # Restart service
    Start-Service -Name $serviceName

    "[$(Get-Date)] ✅ Corina Production updated and restarted successfully." | Out-File -Append $logPath
}
catch {
    "[$(Get-Date)] ❌ Update failed: $_" | Out-File -Append $logPath
}

# === Ensure Scheduled Task has all desired run times ===
$taskName = "CorinaProdDailyUpdater"
$desiredTimes = @("07:00", "09:00", "11:00", "13:00", "15:00", "17:00")

try {
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
    $existingTriggers = $existingTask.Triggers

    $existingTimes = $existingTriggers | ForEach-Object {
        ([DateTime]::Parse($_.StartBoundary)).ToString("HH:mm")
    }

    $missingTimes = $desiredTimes | Where-Object { $_ -notin $existingTimes }

    if ($missingTimes.Count -gt 0) {
        Write-Host "🕐 Adding missing production times: $($missingTimes -join ', ')"
        $newTriggers = @($existingTriggers)
        foreach ($time in $missingTimes) {
            $dt = [datetime]::ParseExact($time, "HH:mm", $null)
            $newTriggers += New-ScheduledTaskTrigger -Daily -At $dt
        }
        Set-ScheduledTask -TaskName $taskName -Trigger $newTriggers
        Write-Host "✅ Production task triggers updated."
    }
    else {
        Write-Host "✅ All production task times exist."
    }
}
catch {
    Write-Error "❌ Failed to update production task schedule: $_"
}
