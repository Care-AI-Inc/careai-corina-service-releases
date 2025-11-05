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
 
# Concurrency guard to avoid overlapping runs
$m = New-Object Threading.Mutex($false, "Global\CorinaDailyUpdater")
if (-not $m.WaitOne([TimeSpan]::FromMinutes(30))) {
    "[$(Get-Date)] ❌ Another updater instance is running." | Out-File -Append $logPath
    exit 1
}
try {
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

    # Stop service
    $serviceName = "CorinaService"
    if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
        Stop-Service -Name $serviceName -Force
        Start-Sleep -Seconds 2
    }

    # Download and extract
    Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip
    Expand-Archive -Path $tempZip -DestinationPath $extractDir -Force

    # Wait until extracted files are readable (handle AV scans)
    function Wait-FileReadable([string]$path, [int]$timeoutSec = 60) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt $timeoutSec) {
            try {
                $fs = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
                $fs.Dispose()
                return $true
            } catch {
                Start-Sleep -Milliseconds 500
            }
        }
        return $false
    }
    Get-ChildItem -Path $extractDir -Recurse -File | ForEach-Object {
        if (-not (Wait-FileReadable $_.FullName 60)) {
            throw "Source not readable after wait: $($_.FullName)"
        }
    }

    # Overwrite files
    $installDir = Join-Path ${env:ProgramFiles} "CorinaService"
    # Use robocopy for resilient copying with retries
    & robocopy "$extractDir" "$installDir" * /E /COPY:DAT /R:5 /W:3 /NFL /NDL /NP /NJH /NJS | Out-Null
    $rc = $LASTEXITCODE
    if ($rc -ge 8) { throw "Robocopy failed with exit code $rc" }

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

    # Ensure only one instance runs at a time
    $settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew
    Set-ScheduledTask -TaskName $taskName -Settings $settings

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
}
finally {
    if ($m) { $m.ReleaseMutex(); $m.Dispose() }
}
