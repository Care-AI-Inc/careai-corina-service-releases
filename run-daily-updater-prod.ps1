# run-daily-updater-prod.ps1
# Safer and more reliable version with TLS 1.2, retry logic, and logging.

$ErrorActionPreference = 'Stop'
$env:DOTNET_ENVIRONMENT = 'Production'
$_inst   = $env:CorinaRegistryInstance
$LogPath = if ($_inst) { "C:\Scripts\corina-prod-update-log-$_inst.txt" } else { 'C:\Scripts\corina-prod-update-log.txt' }
$Url     = 'https://raw.githubusercontent.com/Care-AI-Inc/careai-corina-service-releases/main/daily-updater.ps1'

# 1) Force TLS 1.2 (required for GitHub)
try {
    $proto = [System.Net.ServicePointManager]::SecurityProtocol
    $tls12 = [System.Net.SecurityProtocolType]::Tls12
    if (($proto -band $tls12) -eq 0) {
        [System.Net.ServicePointManager]::SecurityProtocol = $proto -bor $tls12
    }
} catch {
    "`n[$(Get-Date)] ⚠️ Failed to enable TLS 1.2: $_" | Out-File -Append $LogPath
}

# 2) Simple retry helper
function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$MaxRetries = 3,
        [int]$DelaySec   = 5
    )
    $attempt = 0
    while ($true) {
        try {
            $attempt++
            return & $Action
        } catch {
            if ($attempt -ge $MaxRetries) { throw }
            Start-Sleep -Seconds $DelaySec
        }
    }
}

# 3) Download, save, and run
try {
    $Headers = @{ 'User-Agent' = 'PowerShell/5.1 CareAI-Updater' }
    $TmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'daily-updater.ps1')

    $content = Invoke-WithRetry {
        (Invoke-WebRequest -Uri $Url -Headers $Headers -UseBasicParsing -TimeoutSec 30).Content
    }

    if ([string]::IsNullOrWhiteSpace($content)) {
        throw "Downloaded content is empty."
    }
    if ($content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) {
        $content = $content.Substring(1)
    }

    $content | Set-Content -LiteralPath $TmpFile -Encoding UTF8

    # Run the downloaded script in a new process
    & powershell -NoProfile -ExecutionPolicy Bypass -File $TmpFile
}
catch {
    "`n[$(Get-Date)] ❌ Failed to fetch and run latest prod updater: $_" | Out-File -Append $LogPath
    exit 1
}
