# run-daily-updater-prod.ps1
# Safer and more reliable version with TLS 1.2, retry logic, and logging.

$ErrorActionPreference = 'Stop'
$LogPath = 'C:\Scripts\corina-prod-update-log.txt'
$Url     = 'https://raw.githubusercontent.com/Care-AI-Inc/careai-corina-service-releases/main/daily-updater.ps1'

# Optional config via env vars:
# - CORINA_PROXY: explicit proxy URL (e.g., http://proxy:8080)
# - CORINA_SKIP_TLS_VERIFY=1: emergency bypass for TLS certificate validation (INSECURE)
$CorinaProxy = $env:CORINA_PROXY
$SkipTlsVerify = ($env:CORINA_SKIP_TLS_VERIFY -eq '1')

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

function Get-ExceptionText([Exception]$ex) {
    $parts = New-Object System.Collections.Generic.List[string]
    $i = 0
    while ($ex -and $i -lt 10) {
        $parts.Add(("{0}: {1}" -f $ex.GetType().FullName, $ex.Message))
        $ex = $ex.InnerException
        $i++
    }
    return ($parts -join " | ")
}

function Get-ProxyInfo([string]$UriString) {
    try {
        $u = [Uri]$UriString
        $p = [System.Net.WebRequest]::DefaultWebProxy
        if (-not $p) { return "Proxy: <none>" }
        $pu = $p.GetProxy($u)
        if (-not $pu) { return "Proxy: <unknown>" }
        return "Proxy: $($pu.AbsoluteUri)"
    } catch {
        return "Proxy: <error>"
    }
}

# 3) Download, save, and run
try {
    $Headers = @{ 'User-Agent' = 'PowerShell/5.1 CareAI-Updater' }
    $TmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'daily-updater.ps1')

    $prevCb = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
    try {
        if ($SkipTlsVerify) {
            "`n[$(Get-Date)] ⚠️ CORINA_SKIP_TLS_VERIFY=1 enabled. TLS cert validation is being bypassed (INSECURE)." | Out-File -Append $LogPath
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        }

        $iwParams = @{
            Uri            = $Url
            Headers        = $Headers
            UseBasicParsing = $true
            TimeoutSec     = 30
        }
        if ($CorinaProxy) { $iwParams.Proxy = $CorinaProxy }

        $content = Invoke-WithRetry { (Invoke-WebRequest @iwParams).Content }
    } finally {
        if ($SkipTlsVerify) {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $prevCb
        }
    }

    if ([string]::IsNullOrWhiteSpace($content)) {
        throw "Downloaded content is empty."
    }

    $content | Set-Content -LiteralPath $TmpFile -Encoding UTF8

    # Run the downloaded script in a new process
    & powershell -NoProfile -ExecutionPolicy Bypass -File $TmpFile
}
catch {
    $exText = Get-ExceptionText $_.Exception
    "`n[$(Get-Date)] ❌ Failed to fetch and run latest prod updater: $Url" | Out-File -Append $LogPath
    "[$(Get-Date)] ℹ️ SecurityProtocol: $([System.Net.ServicePointManager]::SecurityProtocol)" | Out-File -Append $LogPath
    "[$(Get-Date)] ℹ️ $((Get-ProxyInfo $Url))" | Out-File -Append $LogPath
    "[$(Get-Date)] ℹ️ Exception: $exText" | Out-File -Append $LogPath
    exit 1
}
# SIG # Begin signature block
# MIImbQYJKoZIhvcNAQcCoIImXjCCJloCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBp4CIIIYnVQRD+
# ijsEDI5/7ShbwpoMPE7F5VI/q6wJlqCCE7QwggXrMIID06ADAgECAghWtinNNLx4
# 9jANBgkqhkiG9w0BAQsFADCBgjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFz
# MRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xNzA1
# BgNVBAMMLlNTTC5jb20gRVYgUm9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBS
# U0EgUjIwHhcNMTcwNTMxMTgxNDM3WhcNNDIwNTMwMTgxNDM3WjCBgjELMAkGA1UE
# BhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQK
# DA9TU0wgQ29ycG9yYXRpb24xNzA1BgNVBAMMLlNTTC5jb20gRVYgUm9vdCBDZXJ0
# aWZpY2F0aW9uIEF1dGhvcml0eSBSU0EgUjIwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQCPNmVA4dZNwNe06Ubaa+ozR81M+X19vr0tPfDbeOGGpdm6CVdo
# 7Vc+oNAIQYPnKEEkH+NyFdABGvtecCOyy58548/FTsaSbSbGe7uz2iedCobpgTcF
# /vBxcezDHOljohcUne8bZ9OFVQIC1knJzFrhsfdvMp/J1DuIQaicvcur2217CR+i
# THKQ2isI/M88VM5nD6jPXZYZC8Tjcuut0X0dJ++S6xC/W+s7r8+A3cHSlgRben6k
# qTw4dqRijqA5Xup3z10AWY9mLD4HoqMFJhFpl+qFtw+WC0vIQOFQui6Ky/cPmiLn
# f5o3E83yTRNrIdHAzCLyoUb2RGmcymE1BwBv1mEIEeq6uPbps2DlTbnsnxRmyVdY
# 282HafiKhhIDR79mE3asd300JIWDzdeqnJAanyEsf3i3ZLjY6Kb0eLNVy4TSMsR4
# rqOPYd3OCFOt7Ij8FeSaDeafGnfOTI+4FBU9YpyGOAYAZhLkWXZaU8ACmKIQK2hE
# e455zjNKdqpbgRYbtYrY0AB7XmK0CdaGYw6mBZVJuiiLiJOyNBzYpFVutxzQ3plV
# OyP0IuD5KWYm7CBQd9tKC4++5QJgcEFe1K5QOSIUJsuyO3N0VUcHeYE5qDATROUE
# iq6WEyVCD7lTxJv8zeQc3jz6q9YGSh9nppgwHN0s29wYlVdmxv9ci1b1dwIDAQAB
# o2MwYTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFPlgu9Tj1TT2uPUGgCWn
# c9tGaaieMB0GA1UdDgQWBBT5YLvU49U09rj1BoAlp3PbRmmonjAOBgNVHQ8BAf8E
# BAMCAYYwDQYJKoZIhvcNAQELBQADggIBAFazjssKnUmOv6TEkbtmFwVRmHX75VAs
# ep7xFPqr04o+/5Epj2OL2LSpVAENvpOGL/lKbcde9Vf5ylUcEr5HDzbF32q323XC
# RyV/ufFj+GgtVQTR8o2wpM+8PF4feOeloCBwsATFt/dyp94iDb0zJUaMZJIm4z4u
# Y5bam4w9+BgJ1wPMfYaC4MoEB1FQ1/+S1Qzv2oafmdfrt69o4jkmlLpot7+D0+p6
# Zz1iZ64l5XLo4uTsrhL2Sys8n+mwQPM4VLP9t2jI2saPUTyy+5HcHOebneG3DXKP
# 4qTEqXj56xSsxkMFwmU5KBgCw4KynQW+Ze2WX2V0PPsJNS57nBP9Gw9dx22BOlYP
# zDvhrwIvIqxGykY8oBxM1kS0Xi5cFWYJ4SYp/sZSYbqxc//DDJzlbGqUPxTKQBaV
# hPNZqaxfTGGTbdE7zKKVDCKmZ2dELrnZ0opBs2YLWvt9I6XyGrD/3puDlC7RP9+S
# t5GvBTtlx6Bssc1iEsOQG+MlzjS8b3d2sRDD9wUawNavdGJIF3eSaZBhHN6VgHRU
# jxgcw/MD0L+kQ3WGUxh6Ci4JHDafkf2CiiJL0Q5QJd3LAwwXyYMACE41TYqL7fAC
# lGYsRH/LlSeWF60JMKy2cRduixf2HAnULTuYpXHTVBPZYPP1S2ZP+vHuIBKNtKxX
# sUVjoax2qcL7MIIG3jCCBMagAwIBAgIQYlOvpTOs2ZD3RuJtlmsHrDANBgkqhkiG
# 9w0BAQsFADB7MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcM
# B0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMTcwNQYDVQQDDC5TU0wuY29tIEVW
# IENvZGUgU2lnbmluZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIzMB4XDTI1MDYxMjA1
# MzAyNloXDTI2MDYxMjA1MzAyNlowgbwxCzAJBgNVBAYTAkFVMRgwFgYDVQQIDA9O
# ZXcgU291dGggV2FsZXMxEjAQBgNVBAcMCUhheW1hcmtldDEYMBYGA1UECgwPQ0FS
# RSBBSSBQVFkgTFREMRcwFQYDVQQFEw4zOCA2ODEgOTA0IDUxMjEYMBYGA1UEAwwP
# Q0FSRSBBSSBQVFkgTFREMR0wGwYDVQQPDBRQcml2YXRlIE9yZ2FuaXphdGlvbjET
# MBEGCysGAQQBgjc8AgEDEwJBVTCCAaIwDQYJKoZIhvcNAQEBBQADggGPADCCAYoC
# ggGBAONL4V8ZpSB+jIpLtPqh0yBdLBw3/xkrJmahkSjz53GiZDjwUQNUnNT8bbFP
# pjdmGuSZXveze2vKY3TQ+Imc6Gw/MjsoqrzLQSQTa1S8ZKkc6Vlsph2YmendKlUe
# Q2UiwMyyxQqafZm5yiUdc645EN7y3C7kcvfgs+C2PynqIFRGPCuHLs5lls07TxXh
# dxeAAxv2U+Rq2PZkJ7VHtSpNQex7RwO9QxBlVW69olQHaJ5z2DQs4p7/nFA3YPEL
# d+LFkL4l3SLJS4JLHfzKr2/rkft6h7KZLqAtgP7KzlWVOoDRZMMnQnedv/fvCIKG
# jK4ZoV+7Ym7MBPlTDDeDZOtWuX1qqZL610ydowjm0jYriagJyGpGpL3K1cl3WcwA
# I11ctyX8vsXicAOu5NxTPcbv664UuBCT1VsRdoynTeqF2rE6PYUT9dSzxeGQ7MoV
# WUqmeKRhhleh8f3Ssc+LXRg3zNL1eNk6C8i/FngdcmrOJ9HHBTKhkOBo88H3r9NZ
# sPN4iwIDAQABo4IBmjCCAZYwDAYDVR0TAQH/BAIwADAfBgNVHSMEGDAWgBQ2vUn/
# MSzrr2pA/pnAFu26/EjdXzB9BggrBgEFBQcBAQRxMG8wSwYIKwYBBQUHMAKGP2h0
# dHA6Ly9jZXJ0LnNzbC5jb20vU1NMY29tLVN1YkNBLUVWLUNvZGVTaWduaW5nLVJT
# QS00MDk2LVIzLmNlcjAgBggrBgEFBQcwAYYUaHR0cDovL29jc3BzLnNzbC5jb20w
# UAYDVR0gBEkwRzAHBgVngQwBAzA8BgwrBgEEAYKpMAEDAwIwLDAqBggrBgEFBQcC
# ARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0b3J5MBMGA1UdJQQMMAoGCCsG
# AQUFBwMDMFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9jcmxzLnNzbC5jb20vU1NM
# Y29tLVN1YkNBLUVWLUNvZGVTaWduaW5nLVJTQS00MDk2LVIzLmNybDAdBgNVHQ4E
# FgQUHhCKPf1OVJHGT158hUd4N7CJAkAwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3
# DQEBCwUAA4ICAQCocrpSfj/1MdSXtCeWXJegBOU0DkAxa5hXy9A0C93mlKWDjWGZ
# ir5lpbEoHkWo1dsJOefuh3bqqD2dCFw0yz1rAqKJgfIrhpKUZkEIphJzpiddRwA7
# zavrWaDdhVGTfXKVEbiy4Hf8EEKtkELhD4uw8E4gMM62pn0+XLmhJtJSrWYmvnvD
# kSLTmScCH0wjGKNoUuBiTsHz6vua5c0ei/ahLWgd46ByI2vNBtN9MiyUNyjc5AX/
# 9igp7/QPKsZMxze1bYHDroUadSh/NetufK3s4jtPfvxvt0qOgmIwrF5X4XK1OtA/
# V7WypiwPuZallFgTfyPhsxKZbm1NsRSdmL+tmQefElctHgj9DzlwdPDn0uUR1yPX
# 3lbvu36GCEoqLTFGAQucH/JaVFboKnXwyVv5eyNIauUMJjpadkhAJhiA82rPFvkd
# LUyyZAyROY9yjHthMfhQ3TStmQM6tMDFZ1v376yf7/YJPtE6lM2V2s2nxx6nYr8g
# xOJOo2xRo6bFhKVU6J2utdTMld/xrYpDcm5qm8CsTKVmS+ZAImylkwT3DFKBo8vI
# HUqde9R2XDxPgPn6y6wxp5zewljWsHw7tE/b5fCZd8C9tfnduY0kovqr5gk3SzaF
# QUW5lkRuVyFiul6QM9v4VdjAiPkTkP77v5UWuw1KXtd0oaRqOY1uR8zmBzCCBt8w
# ggTHoAMCAQICEEJLalPOx2YUHCpjsaUcQQQwDQYJKoZIhvcNAQELBQAwgYIxCzAJ
# BgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYG
# A1UECgwPU1NMIENvcnBvcmF0aW9uMTcwNQYDVQQDDC5TU0wuY29tIEVWIFJvb3Qg
# Q2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUlNBIFIyMB4XDTE5MDMyNjE3NDQyM1oX
# DTM0MDMyMjE3NDQyM1owezELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAw
# DgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE3MDUGA1UEAwwuU1NM
# LmNvbSBFViBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBSMzCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAPCqN/crIZEgZzo5jhXlIbpaEyqN
# UdBzPb2uAYwTe9Z++GZj7fR7K5LxvmNr96g3ZXYNfDSnUrHYePBf6Z93emE+g2bs
# Zs63pD19GpWvHV1/xZVJoNjqvSPmlD+ZbiVGOMRVmDg8qfTlrnna+3VuAB8QP7GP
# Av9CrpL89dNaCSVSY4jdX/SRKBYVq1QunPHe4NvSMmkhZ0ZtV1+bytE3f6dpJx6u
# O2pessYKoD1gHnx2xRyjAmVzhDFl7f5VaJusIdGdhH7qAc/k50tMGF1kgXc2aMcD
# +MrENvafEmzdRBkL6WB+CSvbmjw2z46hHAH3dbX2b4cLA1rPmNfLKFCXpaHyqCEc
# +7FMNeoYWxbHRVwAIHlviNNQb3D3xdJDHxeSfjGWqUG6Q/K50Y3GaJLgm4qA1nnW
# KV/mwIGK8ssOTRg2C3WqSTbtI84XzlGHKdDYDKKiZv/b55MTi3yUyWtRjVLWO++K
# DeS9/jihWmhZ2AfntTWwkDg8Wy0iEJcHO7KyMmBhxjgVbLC6tX6D+TyyKh6/rc1Y
# p49vO2w3366ILEffER2o1xS0Za9P9qJJsmFwCv7ZThd4V16JJdLEHkrTnnPqFGgp
# AiJR/c8UBC7/HvOUlJ1zUKyqqStDcSGOdjKWKBBZK+w/IOku5tPjZiUROJxpQ+rT
# JKT/oiXqCA4oWJzpAgMBAAGjggFVMIIBUTASBgNVHRMBAf8ECDAGAQH/AgEAMB8G
# A1UdIwQYMBaAFPlgu9Tj1TT2uPUGgCWnc9tGaaieMHwGCCsGAQUFBwEBBHAwbjBK
# BggrBgEFBQcwAoY+aHR0cDovL3d3dy5zc2wuY29tL3JlcG9zaXRvcnkvU1NMY29t
# LVJvb3RDQS1FVi1SU0EtNDA5Ni1SMi5jcnQwIAYIKwYBBQUHMAGGFGh0dHA6Ly9v
# Y3Nwcy5zc2wuY29tMBEGA1UdIAQKMAgwBgYEVR0gADATBgNVHSUEDDAKBggrBgEF
# BQcDAzBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3Jscy5zc2wuY29tL1NTTGNv
# bS1Sb290Q0EtRVYtUlNBLTQwOTYtUjIuY3JsMB0GA1UdDgQWBBQ2vUn/MSzrr2pA
# /pnAFu26/EjdXzAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZIhvcNAQELBQADggIBAHKP
# +oFIgpHiYIMlW3uPL5QPg1jOiCT6mUJOLU43ififsR6udEB5+d7L9/8sJRBSmECP
# VDj/XdEqqVrmtwK7yH/uKtP/f8w2PFUpQ102SZYmXXDn8isFZ0dMmVgZCPaxxk9g
# 0vw4vgKsJdGIDaUs4d3TfVfPasMZYNJtql17ROhaW4PbyBs2Cn4K9QpSNnjimvsT
# VMycyUe/Yk41rz7hug/Jk+7VILeWt1B2UjV6naE7JmQ3H868A3vEYYFSicx7/loF
# Gkeu5BLKjlTjWp+wwYry+V9GaLmvx9k+hNErJRI4PbuaAerfzGaotsUfapNHsM4G
# koStQ4NqhjlcTOICS3hzrkso5qT4YWmAzP806LAvZAJJDY0uH33roYYFD+1ecDTl
# GAIA62O+dSZtpxyQVweumaWON9Knw1hspfTnUiI1p1u7butI25py3qpaYkkJnpAr
# Eg/IOtuvaHOd2eN5ypj5aB3q5lguqRhszZk6ms0mcETmZpicJR4ZasfY8+f/pjV3
# +/V9u4yCx299VDK76pkLOeggURUvieMq4cUg83p4Tj2vF2KSVI0njJA33OMp6EKT
# tvg7KwuZULjkNAaYI+7q37VUu67b8erdcvlF7bHaQzuA/G9s39yRbbil1O91zWVM
# ZCxZ3xMuAhtL+gSTwLs3HR+yINNPM68WoRzAqqiIMYISDzCCEgsCAQEwgY8wezEL
# MAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREw
# DwYDVQQKDAhTU0wgQ29ycDE3MDUGA1UEAwwuU1NMLmNvbSBFViBDb2RlIFNpZ25p
# bmcgSW50ZXJtZWRpYXRlIENBIFJTQSBSMwIQYlOvpTOs2ZD3RuJtlmsHrDANBglg
# hkgBZQMEAgEFAKCBrzAUBgorBgEEAYI3AgEMMQYwBKECgAAwGQYJKoZIhvcNAQkD
# MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLQYJ
# KoZIhvcNAQk0MSAwHjANBglghkgBZQMEAgEFAKENBgkqhkiG9w0BAQsFADAvBgkq
# hkiG9w0BCQQxIgQgU7YeUmA2KpZUiiuiqa9EyYw9D/iHezDxYwF2sJeIzogwDQYJ
# KoZIhvcNAQELBQAEggGAR7A6dTWXXIVfS876Tnxsl+e3PxeZiVRaGVA/rZ/I2+O/
# KotFu/53ebOAjpe+HfrZmOymM8UWEZ7hysgHhUFsK/j/rnGWmOo1FRdjliwdJd0T
# Wcs82siv30DFjW+JvHcoxakFD/4vseIBCEOG5TZQkdrqLsHlapYpw0u+7IcxYOX3
# eXShQs7R3ygol8sdyKCNtUr9BlVHQ8Ei6yD2S/8oUXn4DS22hgZJi3LThejpDQEi
# smfK2AfV5AOFD35FWf6ri8NsxrZvN6ZdEX3/0uEPknc8Kute5WGPhI79lmvOXlFp
# Ys4MIqsA6RXbLwYY03362d1ZurXYycz3i2AkUDMcTi1BdUiLO2czWS1a173TJhk7
# 0pB6uRCC8Q0v7phbcTernrmAi4j1cVXPcGxsW9CymFTPnuSFeVCFw4ozaazrzyXU
# SWgk6AN9n84QV+fOLhdrKSB0z4RFJuJXTURdD5dBQU/EBVcGXyZHHCrI3dhNjR8Y
# lRHGsO41s12HFcl+6QoaoYIPHjCCDxoGCisGAQQBgjcDAwExgg8KMIIPBgYJKoZI
# hvcNAQcCoIIO9zCCDvMCAQMxDTALBglghkgBZQMEAgEwfwYLKoZIhvcNAQkQAQSg
# cARuMGwCAQEGDCsGAQQBgqkwAQMGATAxMA0GCWCGSAFlAwQCAQUABCB4KDPLmSWD
# tAcOHsIsb/l5p48ch0RaTb0LzWMUNsuvqgIIfwNLmqq5qj4YDzIwMjUxMjAxMDQ0
# ODE4WjADAgEBAgYBmtg9d3agggwAMIIE/DCCAuSgAwIBAgIQWlqs6Bo1brRiho1X
# feA9xzANBgkqhkiG9w0BAQsFADBzMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4
# YXMxEDAOBgNVBAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS8wLQYDVQQD
# DCZTU0wuY29tIFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSBSMTAeFw0yNDAy
# MTkxNjE4MTlaFw0zNDAyMTYxNjE4MThaMG4xCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxKjAo
# BgNVBAMMIVNTTC5jb20gVGltZXN0YW1waW5nIFVuaXQgMjAyNCBFMTBZMBMGByqG
# SM49AgEGCCqGSM49AwEHA0IABKdhcvUw6XrEgxSWBULj3Oid25Rt2TJvSmLLaLy3
# cmVATADvhyMryD2ZELwYfVwABUwivwzYd1mlWCRXUtcEsHyjggFaMIIBVjAfBgNV
# HSMEGDAWgBQMnRAljpqnG5mHQ88IfuG9gZD0zzBRBggrBgEFBQcBAQRFMEMwQQYI
# KwYBBQUHMAKGNWh0dHA6Ly9jZXJ0LnNzbC5jb20vU1NMLmNvbS10aW1lU3RhbXBp
# bmctSS1SU0EtUjEuY2VyMFEGA1UdIARKMEgwPAYMKwYBBAGCqTABAwYBMCwwKgYI
# KwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNzbC5jb20vcmVwb3NpdG9yeTAIBgZngQwB
# BAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwRgYDVR0fBD8wPTA7oDmgN4Y1aHR0
# cDovL2NybHMuc3NsLmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVJTQS1SMS5j
# cmwwHQYDVR0OBBYEFFBPJKzvtT5jEyMJkibsujqW5F0iMA4GA1UdDwEB/wQEAwIH
# gDANBgkqhkiG9w0BAQsFAAOCAgEAmKCPAwCRvKvEZEF/QiHiv6tsIHnuVO7BWILq
# cfZ9lJyIyiCmpLOtJ5VnZ4hvm+GP2tPuOpZdmfTYWdyzhhOsDVDLElbfrKMLiOXn
# 9uwUJpa5fMZe3Zjoh+n/8DdnSw1MxZNMGhuZx4zeyqei91f1OhEU/7b2vnJCc9yB
# FMjY++tVKovFj0TKT3/Ry+Izdbb1gGXTzQQ1uVFy7djxGx/NG1VP/aye4OhxHG9F
# iZ3RM9oyAiPbEgjrnVCc+nWGKr3FTQDKi8vNuyLnCVHkiniL+Lz7H4fBgk163Llx
# i11Ynu5A/phpm1b+M2genvqo1+2r8iVLHrERgFGMUHEdKrZ/OFRDmgFrCTY6xnaP
# TA5/ursCqMK3q3/59uZaOsBZhZkaP9EuOW2p0U8Gkgqp2GNUjFoaDNWFoT/EcoGD
# iTgN8VmQFgn0Fa4/3dOb6lpYEPBcjsWDdqUaxugStY9aW/AwCal4lSN4otljbok8
# u31lZx5NVa4jK6N6upvkgyZ6osmbmIWr9DLhg8bI+KiXDnDWT0547gSuZLYUq+TV
# 6O/DhJZH5LVXJaeS1jjjZZqhK3EEIJVZl0xYV4H4Skvy6hA2rUyFK3+whSNS52TJ
# kshsxVCOPtvqA9ecPqZLwWBaIICG4zVr+GAD7qjWwlaLMd2ZylgOHI3Oit/0pVET
# qJHutyYwggb8MIIE5KADAgECAhBtUhhwh+gjTYVgANCAj5NWMA0GCSqGSIb3DQEB
# CwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0aW9uMTEwLwYDVQQDDChTU0wuY29t
# IFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUlNBMB4XDTE5MTExMzE4NTAw
# NVoXDTM0MTExMjE4NTAwNVowczELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFz
# MRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwm
# U1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgUjEwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQCuURAT0vk8IKAghd7JUBxkyeH9xek0/wp/
# MUjoclrFXqhh/fGH91Fc+7fm0MHCE7A+wmOiqBj9ODrJAYGq3rm33jCnHSsCBNWA
# QYyoauLq8IjqsS1JlXL29qDNMMdwZ8UNzQS7vWZMDJ40JSGNphMGTIA2qn2bohGt
# gRc4p1395ESypUOaGvJ3t0FNL3BuKmb6YctMcQUF2sqooMzd89h0E6ujdvBDo6Zw
# NnWoxj7YmfWjSXg33A5GuY9ym4QZM5OEVgo8ebz/B+gyhyCLNNhh4Mb/4xvCTCMV
# mNYrBviGgdPZYrym8Zb84TQCmSuX0JlLLa6WK1aO6qlwISbb9bVGh866ekKblC/X
# RP20gAu1CjvcYciUgNTrGFg8f8AJgQPOCc1/CCdaJSYwhJpSdheKOnQgESgNmYZP
# hFOC6IKaMAUXk5U1tjTcFCgFvvArXtK4azAWUOO1Y3fdldIBL6LjkzLUCYJNkFXq
# hsBVcPMuB0nUDWvLJfPimstjJ8lF4S6ECxWnlWi7OElVwTnt1GtRqeY9ydvvGLnt
# U+FecK7DbqHDUd366UreMkSBtzevAc9aqoZPnjVMjvFqV1pYOjzmTiVHZtAc80bA
# fFe5LLfJzPI6DntNyqobpwTevQpHqPDN9qqNO83r3kaw8A9j+HZiSw2AX5cGdQP0
# kG0vhzfgBwIDAQABo4IBgTCCAX0wEgYDVR0TAQH/BAgwBgEB/wIBADAfBgNVHSME
# GDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2mWTCBgwYIKwYBBQUHAQEEdzB1MFEGCCsG
# AQUFBzAChkVodHRwOi8vd3d3LnNzbC5jb20vcmVwb3NpdG9yeS9TU0xjb21Sb290
# Q2VydGlmaWNhdGlvbkF1dGhvcml0eVJTQS5jcnQwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMD8GA1UdIAQ4MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIB
# Fh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYB
# BQUHAwgwOwYDVR0fBDQwMjAwoC6gLIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wu
# Y29tLXJzYS1Sb290Q0EuY3JsMB0GA1UdDgQWBBQMnRAljpqnG5mHQ88IfuG9gZD0
# zzAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZIhvcNAQELBQADggIBAJIZdQ2mWkLPGQfZ
# 8vyU+sCb8BXpRJZaL3Ez3VDlE3uZk3cPxPtybVfLuqaci0W6SB22JTMttCiQMnIV
# OsXWnIuAbD/aFTcUkTLBI3xys+wEajzXaXJYWACDS47BRjDtYlDW14gLJxf8W6DQ
# oH3jHDGGy8kGJFOlDKG7/YrK7UGfHtBAEDVe6lyZ+FtCsrk7dD/IiL/+Q3Q6SFAS
# JLQ2XI89ihFugdYL77CiDNXrI2MFspQGswXEAGpHuaQDTHUp/LdR3TyrIsLlnzoL
# skUGswF/KF8+kpWUiKJNC4rPWtNrxlbXYRGgdEdx8SMjUTDClldcrknlFxbqHsVm
# r9xkT2QtFmG+dEq1v5fsIK0vHaHrWjMMmaJ9i+4qGJSD0stYfQ6v0PddT7EpGxGd
# 867Ada6FZyHwbuQSadMb0K0P0OC2r7rwqBUe0BaMqTa6LWzWItgBjGcObXeMxmbQ
# qlEz2YtAcErkZvh0WABDDE4U8GyV/32FdaAvJgTfe9MiL2nSBioYe/g5mHUSWAay
# /Ip1RQmQCvmF9sNfqlhJwkjy/1U1ibUkTIUBX3HgymyQvqQTZLLys6pL2tCdWcjI
# 9YuLw30rgZm8+K387L7ycUvqrmQ3ZJlujHl3r1hgV76s3WwMPgKk1bAEFMj+rRXi
# mSC+Ev30hXZdqyMdl/il5Ksd0vhGMYICWDCCAlQCAQEwgYcwczELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhT
# U0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNzdWluZyBS
# U0EgQ0EgUjECEFparOgaNW60YoaNV33gPccwCwYJYIZIAWUDBAIBoIIBYTAaBgkq
# hkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI1MTIwMTA0
# NDgxOFowKAYJKoZIhvcNAQk0MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIw
# LwYJKoZIhvcNAQkEMSIEINncPxNN4gag8SFel9XLQASMGVLYvo4WmgRlfpK9Ttof
# MIHJBgsqhkiG9w0BCRACLzGBuTCBtjCBszCBsAQgnXF/jcI3ZarOXkqw4fV115oX
# 1Bzu2P2v7wP9Pb2JR+cwgYswd6R1MHMxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVU
# ZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLzAtBgNV
# BAMMJlNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIFIxAhBaWqzo
# GjVutGKGjVd94D3HMAoGCCqGSM49BAMCBEcwRQIhALWNl9UXQgG+9QVx43V6vm5c
# WBufs1Z44aS7Fk0b/C0LAiBpS5453TpAoVENdzP56l9srgL7JWbCfRIAt8W1Q7F+
# cw==
# SIG # End signature block
