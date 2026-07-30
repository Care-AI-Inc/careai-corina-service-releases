Describe 'Corina release script static security policy' {
    BeforeAll {
        $repoRoot = Split-Path -Parent $PSScriptRoot
        $runtimeScripts = @(
            'install.ps1',
            'daily-updater.ps1',
            'ensure-updater-task.ps1',
            'uninstall.ps1',
            'run-daily-updater-prod.ps1'
        )

        function Assert-CorinaEqual {
            param($Actual, $Expected)
            if ($Actual -ne $Expected) {
                throw "Expected '$Expected', got '$Actual'."
            }
        }

        function Assert-CorinaMatch {
            param([AllowNull()]$Actual, [string] $Pattern)
            if ([string]$Actual -notmatch $Pattern) {
                throw "Expected content to match '$Pattern'."
            }
        }

        function Assert-CorinaNotMatch {
            param([AllowNull()]$Actual, [string] $Pattern)
            if ([string]$Actual -match $Pattern) {
                throw "Expected content not to match '$Pattern'."
            }
        }
    }

    foreach ($scriptName in @(
            'install.ps1',
            'daily-updater.ps1',
            'ensure-updater-task.ps1',
            'uninstall.ps1',
            'run-daily-updater-prod.ps1'
        )) {
        It "$scriptName parses without errors" -TestCases @{ ScriptName = $scriptName } {
            param($ScriptName)
            $tokens = $null
            $errors = $null
            [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot $ScriptName), [ref]$tokens, [ref]$errors)
            Assert-CorinaEqual -Actual $errors.Count -Expected 0
        }
    }

    It 'contains no network-to-expression or execution-policy bypass flow' {
        $all = ($runtimeScripts | ForEach-Object { Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw }) -join "`n"
        Assert-CorinaNotMatch -Actual $all -Pattern '(?i)raw\.githubusercontent\.com'
        Assert-CorinaNotMatch -Actual $all -Pattern '(?i)Invoke-Expression'
        Assert-CorinaNotMatch -Actual $all -Pattern '(?i)-ExecutionPolicy\s+Bypass'
        Assert-CorinaNotMatch -Actual $all -Pattern '(?i)Add-MpPreference|Unblock-File|CORINA_DISABLE_CRL'
        Assert-CorinaNotMatch -Actual $all -Pattern '(?i)api\.github\.com/.*/releases/latest'
    }

    It 'pins installer discovery and limits mutable discovery to the updater' {
        $installer = Get-Content -LiteralPath (Join-Path $repoRoot 'install.ps1') -Raw
        $updater = Get-Content -LiteralPath (Join-Path $repoRoot 'daily-updater.ps1') -Raw
        Assert-CorinaNotMatch -Actual $installer -Pattern 'releases/latest/download'
        Assert-CorinaMatch -Actual $installer -Pattern 'CorinaInstallerReleaseVersion'
        Assert-CorinaMatch -Actual $installer -Pattern 'ExpectedSequence'
        Assert-CorinaMatch -Actual $updater -Pattern 'releases/latest/download/.*ManifestFileName'
        Assert-CorinaMatch -Actual $installer -Pattern 'SafeGetValue'
        Assert-CorinaMatch -Actual $updater -Pattern 'SafeGetValue'
        Assert-CorinaMatch -Actual $installer -Pattern 'Get-AuthenticodeSignature'
        Assert-CorinaMatch -Actual $updater -Pattern 'Get-FileHash'
    }

    It 'migrates the legacy updater through authenticated release assets before signed self-validation' {
        $installer = Get-Content -LiteralPath (Join-Path $repoRoot 'install.ps1') -Raw
        $updater = Get-Content -LiteralPath (Join-Path $repoRoot 'daily-updater.ps1') -Raw
        $migrationIndex = $updater.IndexOf('if ($script:IsLegacyUnsignedBootstrap -and')
        $signedSelfValidationIndex = $updater.IndexOf('$null = Assert-CorinaSignedFile -Path $PSCommandPath')

        if ($migrationIndex -lt 0 -or $signedSelfValidationIndex -lt 0 -or $migrationIndex -gt $signedSelfValidationIndex) {
            throw 'Legacy migration must run before the signed updater self-validation path.'
        }
        Assert-CorinaMatch -Actual $updater -Pattern 'LegacyBootstrapMinimumSequence\s*=\s*\[UInt64\]1000003000004'
        Assert-CorinaMatch -Actual $updater -Pattern "LegacySignerPlaceholder = '__CORINA_RELEASE_' \+ 'SIGNER_THUMBPRINTS__'"
        Assert-CorinaMatch -Actual $updater -Pattern 'SignatureStatus\]::NotSigned'
        Assert-CorinaMatch -Actual $updater -Pattern 'Read-CorinaReleaseManifest'
        Assert-CorinaMatch -Actual $updater -Pattern 'Receive-CorinaAsset'
        Assert-CorinaMatch -Actual $updater -Pattern 'RequireAuthenticode'
        Assert-CorinaMatch -Actual $updater -Pattern 'CorinaAgentToken is missing; the running service was not changed'
        Assert-CorinaMatch -Actual $updater -Pattern '\[IO\.FileShare\]::Read'
        Assert-CorinaNotMatch -Actual $installer -Pattern 'Ensure-CorinaUpdaterTask[^\r\n]+-ForceRecreate'
    }

    It 'never destroys existing registry enrolment state' {
        $installer = Get-Content -LiteralPath (Join-Path $repoRoot 'install.ps1') -Raw
        # Registry-provider New-Item -Force REPLACES an existing key, wiping
        # CorinaAgentToken/HaloGuid and every per-instance subkey. Creation of
        # the enrolment key must always be guarded by an existence check, and
        # exactly one (guarded) creation site may exist.
        Assert-CorinaMatch -Actual $installer -Pattern '(?s)if \(-not \(Test-Path -LiteralPath \$regPath\)\) \{\s*New-Item -Path \$regPath -Force'
        $creationSites = [regex]::Matches($installer, 'New-Item -Path \$regPath -Force').Count
        Assert-CorinaEqual -Actual $creationSites -Expected 1
    }

    It 'preserves the existing service account on reinstall and migration' {
        $installer = Get-Content -LiteralPath (Join-Path $repoRoot 'install.ps1') -Raw
        # Clinics with credentialed SMB/NAS shares run the service as a
        # per-site user; only a brand-new service may default to LocalSystem.
        Assert-CorinaMatch -Actual $installer -Pattern 'sc\.exe create [^\r\n]*obj= LocalSystem'
        Assert-CorinaNotMatch -Actual $installer -Pattern 'sc\.exe config [^\r\n]*obj='
    }

    It 'never reassigns a validated Instance parameter variable' {
        # Assigning $null back into a [ValidatePattern] parameter re-triggers
        # validation and throws on every default-instance machine. This list
        # previously omitted run-daily-updater-prod.ps1, which is exactly how
        # that file kept the defect after 0c955ea fixed it everywhere else:
        # it must cover EVERY runtime script that declares an $Instance param.
        foreach ($name in $runtimeScripts) {
            $content = Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw
            Assert-CorinaNotMatch -Actual $content -Pattern '(?m)^\s*\$Instance\s*='
        }
    }

    It 'never treats a bare @() wrap of a nullable value as a countable list' {
        # @($x) is not a safe array-wrap in this codebase. An untyped $null wraps
        # to a ONE-element array holding $null (so 'Count -eq 0' fallbacks never
        # fire), and an unbound typed [string[]] parameter wraps to $null itself
        # (so '.Count' throws outright under Set-StrictMode). Every nullable
        # source must be filtered through the pipeline before it is counted.
        foreach ($name in $runtimeScripts) {
            $content = Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw
            Assert-CorinaNotMatch -Actual $content -Pattern '@\(\$TrustedSignerThumbprints\)'
            Assert-CorinaNotMatch -Actual $content -Pattern '@\(\(Get-CorinaRegistryValue[^|\r\n]*\)\)'
            Assert-CorinaNotMatch -Actual $content -Pattern '(?<!@)\(Get-ChildItem[^\r\n]*\)\.Count'
        }
    }

    It 'reaches the trust fallbacks when no thumbprints are supplied' {
        # The uninstaller crashed at its first .Count before the registry and
        # built-in fallbacks could be consulted, so a bare `.\uninstall.ps1`
        # could never succeed on any clinic.
        $uninstaller = Get-Content -LiteralPath (Join-Path $repoRoot 'uninstall.ps1') -Raw
        Assert-CorinaMatch -Actual $uninstaller -Pattern '(?s)if \(\$null -ne \$TrustedSignerThumbprints\)[^\r\n]*\r?\n[^\r\n]*Where-Object'
        $updater = Get-Content -LiteralPath (Join-Path $repoRoot 'daily-updater.ps1') -Raw
        Assert-CorinaMatch -Actual $updater -Pattern 'storedTrusted[^\r\n]*Where-Object \{ -not \[string\]::IsNullOrWhiteSpace'
        Assert-CorinaMatch -Actual $updater -Pattern 'if \(\$storedTrusted\.Count -eq 0\) \{ \$storedTrusted = \$script:BuiltInTrustedSignerThumbprints \}'
    }

    It 'steps out of the install tree before deleting it' {
        # The signed uninstaller ships inside the directory it removes; an
        # administrator running it from that folder holds it open and the
        # delete fails with "because it is in use".
        $uninstaller = Get-Content -LiteralPath (Join-Path $repoRoot 'uninstall.ps1') -Raw
        $stepOutIndex = $uninstaller.IndexOf('Set-Location -LiteralPath "$env:SystemDrive\')
        $removeIndex = $uninstaller.IndexOf('Remove-Item -LiteralPath $installDir')
        if ($stepOutIndex -lt 0 -or $removeIndex -lt 0 -or $stepOutIndex -gt $removeIndex) {
            throw 'uninstall.ps1 must leave the install tree before removing it.'
        }
    }

    It 'preflights the updater before the SYSTEM task invokes it' {
        $helper = Get-Content -LiteralPath (Join-Path $repoRoot 'ensure-updater-task.ps1') -Raw
        Assert-CorinaMatch -Actual $helper -Pattern 'Get-AuthenticodeSignature'
        Assert-CorinaMatch -Actual $helper -Pattern 'TrustedSignerThumbprints'
        Assert-CorinaMatch -Actual $helper -Pattern 'CAREAIPTYLTD'
        Assert-CorinaNotMatch -Actual $helper -Pattern 'ExecutionPolicy'
        Assert-CorinaNotMatch -Actual $helper -Pattern 'Invoke-WebRequest|Invoke-RestMethod'
    }

    It 'uses bounded signer rotation rather than accumulating historical trust' {
        foreach ($name in @('install.ps1','daily-updater.ps1')) {
            $content = Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw
            # The merge must start from the signer that actually authenticated
            # this manifest, never from the previously trusted set.
            Assert-CorinaMatch -Actual $content -Pattern '@\(@\(\$manifest\._VerifiedSignerThumbprint\) \+ @\(\$manifest\.NextSignerThumbprints\)\)'
            Assert-CorinaNotMatch -Actual $content -Pattern '\$trusted \+ @\(\$manifest\.NextSignerThumbprints\)'
            # The earlier revision of this test pinned the unwrapped left operand
            # verbatim, which locked in a '[string] + [string[]]' concatenation
            # that fuses two thumbprints into one 80-character value. Assert the
            # broken shape can never come back.
            Assert-CorinaNotMatch -Actual $content -Pattern '@\(\$manifest\._VerifiedSignerThumbprint \+ '
        }
    }

    It 'merges the current and successor signers into separate thumbprints' {
        # Behavioural guard for the concatenation defect above: with a successor
        # announced, the merge must yield TWO 40-hex entries, not one 80-char one.
        $current = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
        $successors = [string[]]@('BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB')
        $merged = @(@($current) + @($successors))
        Assert-CorinaEqual -Actual $merged.Count -Expected 2
        Assert-CorinaEqual -Actual $merged[0] -Expected $current
        Assert-CorinaEqual -Actual $merged[1] -Expected $successors[0]
        foreach ($entry in $merged) {
            if ($entry -notmatch '^[0-9A-F]{40}$') { throw "Merged trust entry '$entry' is not a single 40-hex thumbprint." }
        }
    }

    It 'limits ZIP entry count and expanded size before extraction' {
        foreach ($name in @('install.ps1','daily-updater.ps1')) {
            $content = Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw
            Assert-CorinaMatch -Actual $content -Pattern 'Entries.Count -gt 20000'
            Assert-CorinaMatch -Actual $content -Pattern '4294967296L'
            Assert-CorinaMatch -Actual $content -Pattern 'duplicate path'
        }
    }

    It 'keeps repository sources unsigned for final build-time signing' {
        foreach ($name in @('install.ps1','daily-updater.ps1','ensure-updater-task.ps1','uninstall.ps1')) {
            Assert-CorinaNotMatch `
                -Actual (Get-Content -LiteralPath (Join-Path $repoRoot $name) -Raw) `
                -Pattern '# SIG # Begin signature block'
        }
    }
}

Describe 'Signed data-only manifest contract' {
    BeforeAll {
        $repoRoot = Split-Path -Parent $PSScriptRoot

        function Assert-CorinaEqual {
            param($Actual, $Expected)
            if ($Actual -ne $Expected) {
                throw "Expected '$Expected', got '$Actual'."
            }
        }

        function Assert-CorinaThrows {
            param([scriptblock] $Action)
            try {
                & $Action
            }
            catch {
                return
            }
            throw 'Expected the action to throw, but it completed successfully.'
        }

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
    }

    AfterAll {
        if ($contractModule) { Remove-Module $contractModule -Force -ErrorAction SilentlyContinue }
    }

    It 'accepts the build-pipeline schema fixture' {
        $manifest = Read-CorinaReleaseManifest -Path $fixture -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 1
        Assert-CorinaEqual -Actual $manifest.ReleaseVersion -Expected '9.8.7'
        Assert-CorinaEqual -Actual $manifest.Assets.Count -Expected 5
        Assert-CorinaEqual -Actual $manifest.NextSignerThumbprints[0] -Expected 'BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB'
        Assert-CorinaEqual -Actual $manifest._VerifiedSignerThumbprint -Expected 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    }

    It 'rejects an executable manifest even after the signature stub succeeds' {
        $badPath = Join-Path $TestDrive 'executable-manifest.ps1'
        "Get-Process`n" | Set-Content -LiteralPath $badPath
        Assert-CorinaThrows -Action {
            Read-CorinaReleaseManifest -Path $badPath -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 1
        }
    }

    It 'rejects an asset URL outside the exact immutable release tag' {
        $badPath = Join-Path $TestDrive 'mutable-url.ps1'
        (Get-Content -LiteralPath $fixture -Raw).Replace('/releases/download/v9.8.7/install.ps1', '/raw/main/install.ps1') | Set-Content -LiteralPath $badPath
        Assert-CorinaThrows -Action {
            Read-CorinaReleaseManifest -Path $badPath -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 1
        }
    }

    It 'rejects rollback sequences' {
        Assert-CorinaThrows -Action {
            Read-CorinaReleaseManifest -Path $fixture -AllowedThumbprints @('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA') -MinimumSequence 988
        }
    }

    It 'normalizes the real SSL.com CARE AI X.500 subject rendering' {
        $realSubject = (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'fixtures\care-ai-real-subject.txt') -Raw).Trim()
        Assert-CorinaEqual `
            -Actual (ConvertTo-CorinaIdentityValue (Get-CorinaCertificateSubjectAttribute -Subject $realSubject -Names @('O'))) `
            -Expected 'CAREAIPTYLTD'
        Assert-CorinaEqual `
            -Actual (ConvertTo-CorinaIdentityValue (Get-CorinaCertificateSubjectAttribute -Subject $realSubject -Names @('C'))) `
            -Expected 'AU'
        Assert-CorinaEqual `
            -Actual (ConvertTo-CorinaIdentityValue (Get-CorinaCertificateSubjectAttribute -Subject $realSubject -Names @('SERIALNUMBER','OID.2.5.4.5'))) `
            -Expected '38681904512'
    }
}
