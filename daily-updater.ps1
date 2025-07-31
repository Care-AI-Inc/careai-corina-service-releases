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

    $tempZip = "$env:TEMP\$zipName"
    $extractDir = "$env:TEMP\CorinaProdExtract"

    # Download and extract
    Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip
    if (Test-Path $extractDir) { Remove-Item -Recurse -Force $extractDir }
    Expand-Archive -Path $tempZip -DestinationPath $extractDir

    # Stop service
    $serviceName = "CorinaService"
    if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
        Stop-Service -Name $serviceName -Force
        Start-Sleep -Seconds 2
    }

    # Overwrite files
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
