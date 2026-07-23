$repoRoot = Split-Path -Parent $PSScriptRoot
$runtimeScripts = @(
    'install.ps1',
    'daily-updater.ps1',
    'ensure-updater-task.ps1',
    'uninstall.ps1',
    'run-daily-updater-prod.ps1'
)

Describe 'Corina release script static security policy' {
    foreach ($name in $runtimeScripts) {
        It "$name parses without errors" {
            $tokens = $null
            $errors = $null
            [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot $name), [ref]$tokens, [ref]$errors)
            $errors.Count | Should Be 0
        }
    }

    It 'contains no network-to-expression or execution-policy bypass flow' {
        $all = ($runtimeScripts | ForEach-Object { Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw }) -join "`n"
        $all | Should Not Match '(?i)raw\.githubusercontent\.com'
        $all | Should Not Match '(?i)Invoke-Expression'
        $all | Should Not Match '(?i)-ExecutionPolicy\s+Bypass'
        $all | Should Not Match '(?i)Add-MpPreference|Unblock-File|CORINA_DISABLE_CRL'
        $all | Should Not Match '(?i)api\.github\.com/.*/releases/latest'
    }

    It 'pins installer discovery and limits mutable discovery to the updater' {
        $installer = Get-Content -LiteralPath (Join-Path $repoRoot 'install.ps1') -Raw
        $updater = Get-Content -LiteralPath (Join-Path $repoRoot 'daily-updater.ps1') -Raw
        $installer | Should Not Match 'releases/latest/download'
        $installer | Should Match 'CorinaInstallerReleaseVersion'
        $installer | Should Match 'ExpectedSequence'
        $updater | Should Match 'releases/latest/download/.*ManifestFileName'
        $installer | Should Match 'SafeGetValue'
        $updater | Should Match 'SafeGetValue'
        $installer | Should Match 'Get-AuthenticodeSignature'
        $updater | Should Match 'Get-FileHash'
    }

    It 'preflights the updater before the SYSTEM task invokes it' {
        $helper = Get-Content -LiteralPath (Join-Path $repoRoot 'ensure-updater-task.ps1') -Raw
        $helper | Should Match 'Get-AuthenticodeSignature'
        $helper | Should Match 'TrustedSignerThumbprints'
        $helper | Should Match 'CAREAIPTYLTD'
        $helper | Should Not Match 'ExecutionPolicy'
        $helper | Should Not Match 'Invoke-WebRequest|Invoke-RestMethod'
    }

    It 'uses bounded signer rotation rather than accumulating historical trust' {
        foreach ($name in @('install.ps1','daily-updater.ps1')) {
            $content = Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw
            $content | Should Match '\$manifest\._VerifiedSignerThumbprint \+ @\(\$manifest\.NextSignerThumbprints\)'
            $content | Should Not Match '\$trusted \+ @\(\$manifest\.NextSignerThumbprints\)'
        }
    }

    It 'limits ZIP entry count and expanded size before extraction' {
        foreach ($name in @('install.ps1','daily-updater.ps1')) {
            $content = Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw
            $content | Should Match 'Entries.Count -gt 20000'
            $content | Should Match '4294967296L'
            $content | Should Match 'duplicate path'
        }
    }

    It 'keeps repository sources unsigned for final build-time signing' {
        foreach ($name in @('install.ps1','daily-updater.ps1','ensure-updater-task.ps1','uninstall.ps1')) {
            (Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw) | Should Not Match '# SIG # Begin signature block'
        }
    }
}

Describe 'Signed data-only manifest contract' {
    $installerPath = Join-Path $repoRoot 'install.ps1'
    $tokens = $null
    $errors = $null
    $installerAst = [Management.Automation.Language.Parser]::ParseFile($installerPath, [ref]$tokens, [ref]$errors)
    $functionNames = @('ConvertTo-CorinaThumbprintList','Get-CorinaCertificateSubjectAttribute','ConvertTo-CorinaIdentityValue','Assert-CorinaAssetDefinition','Read-CorinaReleaseManifest')
    $definitions = foreach ($functionName in $functionNames) {
        $node = $installerAst.Find({ param($ast) $ast -is [Management.Automation.Language.FunctionDefinitionAst] -and $ast.Name -eq $functionName }, $true)
        if (-not $node) { throw "Test could not locate $functionName in install.ps1." }
        $node.Extent.Text
    }
    $moduleText = @"
`$script:CorinaReleaseChannel = 'production'
`$script:CorinaReleaseRepository = 'Care-AI-Inc/careai-corina-service-releases'
`$script:CorinaServiceSourceRepository = 'Care-AI-Inc/careai-corina-service'
`$script:CorinaScriptsSourceRepository = 'Care-AI-Inc/careai-corina-service-releases'
function Assert-CorinaSignedFile {
    param([string]`$Path,[string[]]`$AllowedThumbprints)
    `$certificate = [pscustomobject]@{
        Subject = 'CN=CARE AI PTY LTD, SERIALNUMBER=38 681 904 512, O=CARE AI PTY LTD, C=AU'
        Thumbprint = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    }
    return [pscustomobject]@{ SignerCertificate = `$certificate; TimeStamperCertificate = [object]::new() }
}
$($definitions -join "`n")
Export-ModuleMember -Function Read-CorinaReleaseManifest,Get-CorinaCertificateSubjectAttribute,ConvertTo-CorinaIdentityValue
"@
    $contractModule = New-Module -ScriptBlock ([scriptblock]::Create($moduleText))
    Import-Module $contractModule -Force
    $fixture = Join-Path $PSScriptRoot 'fixtures\corina-production.ps1'

    It 'accepts the build-pipeline schema fixture' {
        $manifest = Read-CorinaReleaseManifest -Path $fixture -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 1
        $manifest.ReleaseVersion | Should Be '9.8.7'
        $manifest.Assets.Count | Should Be 5
        $manifest.NextSignerThumbprints[0] | Should Be 'BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB'
        $manifest._VerifiedSignerThumbprint | Should Be 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    }

    It 'rejects an executable manifest even after the signature stub succeeds' {
        $badPath = Join-Path $TestDrive 'executable-manifest.ps1'
        "Get-Process`n" | Set-Content -LiteralPath $badPath
        { Read-CorinaReleaseManifest -Path $badPath -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 1 } | Should Throw
    }

    It 'rejects an asset URL outside the exact immutable release tag' {
        $badPath = Join-Path $TestDrive 'mutable-url.ps1'
        (Get-Content -LiteralPath $fixture -Raw).Replace('/releases/download/v9.8.7/install.ps1', '/raw/main/install.ps1') | Set-Content -LiteralPath $badPath
        { Read-CorinaReleaseManifest -Path $badPath -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 1 } | Should Throw
    }

    It 'rejects rollback sequences' {
        { Read-CorinaReleaseManifest -Path $fixture -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 988 } | Should Throw
    }

    It 'normalizes the real SSL.com CARE AI X.500 subject rendering' {
        $realSubject = (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'fixtures\care-ai-real-subject.txt') -Raw).Trim()
        (ConvertTo-CorinaIdentityValue (Get-CorinaCertificateSubjectAttribute -Subject $realSubject -Names @('O'))) | Should Be 'CAREAIPTYLTD'
        (ConvertTo-CorinaIdentityValue (Get-CorinaCertificateSubjectAttribute -Subject $realSubject -Names @('C'))) | Should Be 'AU'
        (ConvertTo-CorinaIdentityValue (Get-CorinaCertificateSubjectAttribute -Subject $realSubject -Names @('SERIALNUMBER','OID.2.5.4.5'))) | Should Be '38681904512'
    }
}
