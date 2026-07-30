# Corina Service secure production updater.
#
# This installed, byte-stable script is the scheduled task target. The task
# preflights its Authenticode signature before PowerShell loads it; this script
# then repeats that check and authenticates every downloaded byte before use.

[CmdletBinding()]
param(
    [ValidatePattern('^[A-Za-z0-9](?:[A-Za-z0-9_-]*[A-Za-z0-9])?$')]
    [string]$Instance
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:CorinaReleaseChannel = 'production'
$script:CorinaReleaseRepository = 'Care-AI-Inc/careai-corina-service-releases'
$script:CorinaServiceSourceRepository = 'Care-AI-Inc/careai-corina-service'
$script:CorinaScriptsSourceRepository = 'Care-AI-Inc/careai-corina-service-releases'
$script:CorinaServiceBaseName = 'CorinaService'
$script:CorinaTaskBaseName = 'CorinaProdDailyUpdater'
$script:CorinaProgramDataRoot = 'CareAI\CorinaService'
$script:CorinaRegistryRoot = 'HKLM:\SOFTWARE\CareAI\CorinaService'
$script:CorinaDotNetEnvironment = 'Production'
$script:CorinaLegacyShimBaseName = 'run-daily-updater-prod'
$script:CorinaManifestFileName = "corina-$($script:CorinaReleaseChannel).ps1"
$script:CorinaManifestUri = "https://github.com/$($script:CorinaReleaseRepository)/releases/latest/download/$($script:CorinaManifestFileName)"
$script:BuiltInTrustedSignerThumbprints = @('__CORINA_RELEASE_SIGNER_THUMBPRINTS__')
$script:LegacySignerPlaceholder = '__CORINA_RELEASE_' + 'SIGNER_THUMBPRINTS__'
$script:IsLegacyUnsignedBootstrap = (
    $script:BuiltInTrustedSignerThumbprints.Count -eq 1 -and
    $script:BuiltInTrustedSignerThumbprints[0] -ceq $script:LegacySignerPlaceholder
)
# Existing clinics run a legacy task that downloads this source file. This
# public certificate identity is used once to authenticate the first signed
# release. Signed release builds replace the sentinel above and never enter
# this compatibility path.
$script:LegacyBootstrapTrustedSignerThumbprints = @('FEA8C8EF4EB6E9525D1303D5CC2CC7B4F3447810')
$script:LegacyBootstrapMinimumSequence = [UInt64]1000003000004
$script:CareAiPublisher = @{
    CommonName   = 'CARE AI PTY LTD'
    Organisation = 'CARE AI PTY LTD'
    Country      = 'AU'
    SerialNumber = '38681904512'
}

function Get-CorinaCertificateSubjectAttribute {
    param([Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string[]]$Names)
    foreach ($component in [regex]::Split($Subject, '(?<!\\),')) {
        $separator = $component.IndexOf('=')
        if ($separator -gt 0 -and $component.Substring(0, $separator).Trim().ToUpperInvariant() -in $Names) {
            return $component.Substring($separator + 1).Trim()
        }
    }
    return $null
}

function ConvertTo-CorinaIdentityValue {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return (($Value -replace '[^A-Za-z0-9]', '').ToUpperInvariant())
}

function ConvertTo-CorinaThumbprintList {
    param([string[]]$Values, [switch]$AllowEmpty)
    $invalid = @($Values | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_) -and
        ([string]$_).Replace(' ', '').ToUpperInvariant() -notmatch '^[0-9A-F]{40}$' -and
        [string]$_ -ne '__CORINA_RELEASE_SIGNER_THUMBPRINTS__'
    })
    if ($invalid.Count -gt 0) { throw 'Trusted signer state contains an invalid thumbprint.' }
    $result = @($Values | ForEach-Object {
        if ($null -ne $_) { ([string]$_).Replace(' ', '').ToUpperInvariant() }
    } | Where-Object { $_ -match '^[0-9A-F]{40}$' } | Select-Object -Unique)
    if (-not $AllowEmpty -and $result.Count -eq 0) { throw 'Trusted signer state is empty; update is blocked (fail closed).' }
    return ,([string[]]$result)
}

function Test-CorinaCertificatePublisher {
    param([Parameter(Mandatory)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)
    $simpleName = $Certificate.GetNameInfo([Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
    if ((ConvertTo-CorinaIdentityValue $simpleName) -cne (ConvertTo-CorinaIdentityValue $script:CareAiPublisher.CommonName)) { return $false }
    $organisation = Get-CorinaCertificateSubjectAttribute -Subject $Certificate.Subject -Names @('O')
    $country = Get-CorinaCertificateSubjectAttribute -Subject $Certificate.Subject -Names @('C')
    $serial = Get-CorinaCertificateSubjectAttribute -Subject $Certificate.Subject -Names @('SERIALNUMBER','OID.2.5.4.5','2.5.4.5')
    if ((ConvertTo-CorinaIdentityValue $organisation) -cne (ConvertTo-CorinaIdentityValue $script:CareAiPublisher.Organisation) -or (ConvertTo-CorinaIdentityValue $country) -cne (ConvertTo-CorinaIdentityValue $script:CareAiPublisher.Country) -or (ConvertTo-CorinaIdentityValue $serial) -cne $script:CareAiPublisher.SerialNumber) { return $false }
    $hasCodeSigningEku = $false
    foreach ($extension in $Certificate.Extensions) {
        if ($extension.Oid.Value -eq '2.5.29.37') {
            $eku = [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]$extension
            $hasCodeSigningEku = [bool]($eku.EnhancedKeyUsages | Where-Object { $_.Value -eq '1.3.6.1.5.5.7.3.3' })
        }
    }
    return $hasCodeSigningEku
}

function Assert-CorinaSignedFile {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string[]]$AllowedThumbprints)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Signed file not found: $Path" }
    $signature = Get-AuthenticodeSignature -FilePath $Path
    if ($signature.Status -ne [Management.Automation.SignatureStatus]::Valid -or -not $signature.SignerCertificate) {
        throw "Authenticode validation failed for '$Path': $($signature.Status) $($signature.StatusMessage)"
    }
    $thumbprint = $signature.SignerCertificate.Thumbprint.Replace(' ', '').ToUpperInvariant()
    if ($thumbprint -notin $AllowedThumbprints) { throw "Signer thumbprint $thumbprint is not trusted for '$Path'." }
    if (-not (Test-CorinaCertificatePublisher -Certificate $signature.SignerCertificate)) { throw "CARE AI publisher identity check failed for '$Path'." }
    return $signature
}

function Get-CorinaSha256 {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Get-CorinaRegistryValue {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name)
    try { return Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop }
    catch { return $null }
}

function Assert-CorinaAssetDefinition {
    param(
        [Parameter(Mandatory)][hashtable]$Asset,
        [Parameter(Mandatory)][string]$Role,
        [Parameter(Mandatory)][string]$ReleaseVersion,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ExpectedFileName,
        [long]$MaximumSize
    )
    $unexpected = @($Asset.Keys | Where-Object { [string]$_ -notin @('FileName','Url','Sha256','Size') })
    if ($unexpected.Count) { throw "Asset '$Role' has unexpected field(s): $($unexpected -join ', ')." }
    $fileName = [string]$Asset.FileName
    if ([string]::IsNullOrWhiteSpace($fileName) -or $fileName -cne [IO.Path]::GetFileName($fileName) -or $fileName -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { throw "Asset '$Role' has an unsafe filename." }
    if ($ExpectedFileName -and $fileName -cne $ExpectedFileName) { throw "Asset '$Role' must be named '$ExpectedFileName'." }
    if ([string]$Asset.Sha256 -notmatch '^[0-9A-Fa-f]{64}$') { throw "Asset '$Role' has an invalid SHA-256." }
    $size = 0L
    if (-not [long]::TryParse([string]$Asset.Size, [ref]$size) -or $size -le 0 -or $size -gt $MaximumSize) { throw "Asset '$Role' has an invalid size." }
    $releaseTag = if ($script:CorinaReleaseChannel -eq 'staging') { "staging-v$ReleaseVersion" } else { "v$ReleaseVersion" }
    $expectedUri = "https://github.com/$($script:CorinaReleaseRepository)/releases/download/$releaseTag/$fileName"
    if ([string]$Asset.Url -cne $expectedUri) { throw "Asset '$Role' URL is not the exact immutable release URL '$expectedUri'." }
}

function Read-CorinaReleaseManifest {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string[]]$AllowedThumbprints, [UInt64]$MinimumSequence)
    $manifestSignature = Assert-CorinaSignedFile -Path $Path -AllowedThumbprints $AllowedThumbprints
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Manifest parse error: $($parseErrors[0].Message)" }
    $cleanBlockProperty = $ast.PSObject.Properties['CleanBlock']
    $hasCleanBlock = [bool]($cleanBlockProperty -and $cleanBlockProperty.Value)
    if ($ast.BeginBlock -or $ast.ProcessBlock -or $hasCleanBlock -or $ast.ParamBlock -or -not $ast.EndBlock -or $ast.EndBlock.Statements.Count -ne 1) { throw 'Manifest must contain exactly one literal hashtable.' }
    $statement = $ast.EndBlock.Statements[0]
    if ($statement -isnot [Management.Automation.Language.PipelineAst] -or $statement.PipelineElements.Count -ne 1 -or
        $statement.PipelineElements[0] -isnot [Management.Automation.Language.CommandExpressionAst] -or
        $statement.PipelineElements[0].Expression -isnot [Management.Automation.Language.HashtableAst]) { throw 'Manifest root must be a literal hashtable.' }
    try { $manifest = [hashtable]$statement.PipelineElements[0].Expression.SafeGetValue() }
    catch { throw "Manifest contains a non-literal value: $_" }

    $unexpectedTop = @($manifest.Keys | Where-Object { [string]$_ -notin @('SchemaVersion','Channel','Repository','ReleaseVersion','Sequence','PublishedUtc','Source','Signer','NextSignerThumbprints','Assets') })
    if ($unexpectedTop.Count) { throw "Manifest has unexpected field(s): $($unexpectedTop -join ', ')." }
    $requiredTop = @('SchemaVersion','Channel','Repository','ReleaseVersion','Sequence','PublishedUtc','Source','Signer','NextSignerThumbprints','Assets')
    $missingTop = @($requiredTop | Where-Object { -not $manifest.ContainsKey($_) })
    if ($missingTop.Count) { throw "Manifest is missing field(s): $($missingTop -join ', ')." }
    if ([int]$manifest.SchemaVersion -ne 1) { throw 'Unsupported manifest schema.' }
    if ([string]$manifest.Channel -cne $script:CorinaReleaseChannel -or [string]$manifest.Repository -cne $script:CorinaReleaseRepository) { throw 'Manifest channel/repository does not match this signed updater build.' }
    if ($manifest.Source -isnot [hashtable]) { throw 'Manifest Source must be a hashtable.' }
    $sourceKeys = @('Repository','Commit','Ref','ScriptsRepository','ScriptsCommit')
    if (@($manifest.Source.Keys | Where-Object { [string]$_ -notin $sourceKeys }).Count -gt 0 -or @($sourceKeys | Where-Object { -not $manifest.Source.ContainsKey($_) }).Count -gt 0) { throw 'Manifest Source fields do not match schema 1.' }
    if ([string]$manifest.Source.Repository -cne $script:CorinaServiceSourceRepository -or [string]$manifest.Source.ScriptsRepository -cne $script:CorinaScriptsSourceRepository) { throw 'Manifest source repositories do not match this signed updater build.' }
    if ([string]$manifest.Source.Commit -cnotmatch '^[0-9a-f]{40}$' -or [string]$manifest.Source.ScriptsCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Manifest source commits are invalid.' }
    if ([string]::IsNullOrWhiteSpace([string]$manifest.Source.Ref) -or [string]$manifest.Source.Ref -match '[\x00-\x1F\x7F]' -or ([string]$manifest.Source.Ref).Length -gt 255) { throw 'Manifest source ref is invalid.' }
    if ($manifest.Signer -isnot [hashtable]) { throw 'Manifest Signer must be a hashtable.' }
    $signerKeys = @('Subject','CertificateThumbprint','TimestampRequired')
    if (@($manifest.Signer.Keys | Where-Object { [string]$_ -notin $signerKeys }).Count -gt 0 -or @($signerKeys | Where-Object { -not $manifest.Signer.ContainsKey($_) }).Count -gt 0) { throw 'Manifest Signer fields do not match schema 1.' }
    $actualManifestThumbprint = $manifestSignature.SignerCertificate.Thumbprint.Replace(' ', '').ToUpperInvariant()
    if ([string]$manifest.Signer.Subject -cne [string]$manifestSignature.SignerCertificate.Subject -or ([string]$manifest.Signer.CertificateThumbprint).Replace(' ', '').ToUpperInvariant() -cne $actualManifestThumbprint) { throw 'Manifest Signer metadata does not match its signature.' }
    if ($manifest.Signer.TimestampRequired -isnot [bool] -or -not [bool]$manifest.Signer.TimestampRequired -or -not $manifestSignature.TimeStamperCertificate) { throw 'Manifest must declare and contain a valid timestamp.' }
    $version = [string]$manifest.ReleaseVersion
    if ($version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$') { throw 'Manifest ReleaseVersion is not strict three-part SemVer.' }
    $sequence = [UInt64]0
    if (-not [UInt64]::TryParse([string]$manifest.Sequence, [ref]$sequence) -or $sequence -eq 0 -or $sequence -lt $MinimumSequence) { throw 'Manifest sequence is invalid or is a rollback.' }
    $published = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse([string]$manifest.PublishedUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$published) -or $published -gt [DateTimeOffset]::UtcNow.AddHours(24)) { throw 'Manifest PublishedUtc is invalid.' }
    if ($manifest.Assets -isnot [hashtable]) { throw 'Manifest Assets must be a hashtable.' }
    $required = @{
        InstallerScript='install.ps1'; UpdaterScript='daily-updater.ps1'; TaskHelperScript='ensure-updater-task.ps1';
        UninstallerScript='uninstall.ps1'; ServicePackage=''
    }
    $extraAssets = @($manifest.Assets.Keys | Where-Object { [string]$_ -notin $required.Keys })
    if ($extraAssets.Count) { throw "Manifest has unexpected asset role(s): $($extraAssets -join ', ')." }
    foreach ($role in $required.Keys) {
        if (-not $manifest.Assets.ContainsKey($role) -or $manifest.Assets[$role] -isnot [hashtable]) { throw "Manifest is missing '$role'." }
        Assert-CorinaAssetDefinition -Asset $manifest.Assets[$role] -Role $role -ReleaseVersion $version -ExpectedFileName $required[$role] -MaximumSize $(if ($role -eq 'ServicePackage') { 2147483648L } else { 5242880L })
    }
    if ([IO.Path]::GetExtension([string]$manifest.Assets.ServicePackage.FileName) -cne '.zip') { throw 'ServicePackage must be a .zip.' }
    $rawNextSigners = @($manifest.NextSignerThumbprints)
    if (@($rawNextSigners | Where-Object { [string]$_ -notmatch '^[0-9A-Fa-f]{40}$' }).Count -gt 0 -or $rawNextSigners.Count -gt 3) { throw 'Manifest NextSignerThumbprints must contain at most three valid thumbprints.' }
    $manifest.NextSignerThumbprints = ConvertTo-CorinaThumbprintList -Values $rawNextSigners -AllowEmpty
    $manifest['_VerifiedSignerThumbprint'] = $actualManifestThumbprint
    return $manifest
}

function Receive-CorinaFile {
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$Destination)
    $directory = Split-Path -Parent $Destination
    if (-not (Test-Path -LiteralPath $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
    Invoke-WebRequest -Uri $Uri -OutFile $Destination -UseBasicParsing -TimeoutSec 600 -Headers @{ 'User-Agent'='CareAI-Corina-SecureUpdater/2' }
    if (-not (Test-Path -LiteralPath $Destination -PathType Leaf)) { throw "Download produced no file: $Uri" }
}

function Receive-CorinaAsset {
    param([Parameter(Mandatory)][hashtable]$Asset, [Parameter(Mandatory)][string]$Destination, [Parameter(Mandatory)][string[]]$AllowedThumbprints, [switch]$RequireAuthenticode)
    Receive-CorinaFile -Uri ([string]$Asset.Url) -Destination $Destination
    if ((Get-Item -LiteralPath $Destination).Length -ne [long]$Asset.Size) { throw "Size verification failed for '$($Asset.FileName)'." }
    if ((Get-CorinaSha256 -Path $Destination) -cne ([string]$Asset.Sha256).ToUpperInvariant()) { throw "SHA-256 verification failed for '$($Asset.FileName)'." }
    if ($RequireAuthenticode) { $null = Assert-CorinaSignedFile -Path $Destination -AllowedThumbprints $AllowedThumbprints }
}

function Expand-CorinaArchiveSafely {
    param([Parameter(Mandatory)][string]$ArchivePath, [Parameter(Mandatory)][string]$Destination)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Recurse -Force }
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    $root = [IO.Path]::GetFullPath($Destination).TrimEnd('\') + '\'
    $archive = [IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        if ($archive.Entries.Count -lt 1) { throw 'Service package is empty.' }
        if ($archive.Entries.Count -gt 20000) { throw 'Service package contains more than 20,000 entries.' }
        $totalUncompressed = 0L
        $seenTargets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $archive.Entries) {
            $normalised = $entry.FullName.Replace('/', '\')
            if ([string]::IsNullOrWhiteSpace($normalised) -or [IO.Path]::IsPathRooted($normalised) -or $normalised -match '(^|\\)\.\.(\\|$)' -or $normalised.Contains(':')) { throw "Unsafe archive path '$($entry.FullName)'." }
            $target = [IO.Path]::GetFullPath((Join-Path $Destination $normalised))
            if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { throw "Archive path escapes staging: '$($entry.FullName)'." }
            if (-not $seenTargets.Add($target)) { throw "Service package contains a duplicate path: '$($entry.FullName)'." }
            if ($entry.Length -gt 2147483648L -or $totalUncompressed -gt (4294967296L - $entry.Length)) { throw 'Service package exceeds the 4 GiB uncompressed safety limit.' }
            $totalUncompressed += $entry.Length
            $externalAttributes = [UInt32]([Int64]$entry.ExternalAttributes -band 0xFFFFFFFFL)
            $unixFileType = (($externalAttributes -shr 16) -band 0xF000)
            if ($unixFileType -eq 0xA000) { throw "Service package contains a symbolic link: '$($entry.FullName)'." }
        }
    } finally { $archive.Dispose() }
    Expand-Archive -LiteralPath $ArchivePath -DestinationPath $Destination -Force
}

function Copy-CorinaTree {
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination, [switch]$Mirror)
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    $mode = if ($Mirror) { '/MIR' } else { '/E' }
    # /IS re-copies unchanged same-size files in general, but it is NOT reliable
    # for the dotfile '.version' (robocopy's wildcard + fixed 1980 ZIP timestamp
    # skip it as "Same"), so that marker is copied explicitly below.
    & robocopy $Source $Destination '*' $mode /IS /COPY:DAT /R:10 /W:5 /NFL /NDL /NP /NJH /NJS | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed (exit $LASTEXITCODE) copying '$Source' to '$Destination'." }
    # Deterministically overwrite the version marker; robocopy cannot be trusted
    # to re-copy the same-size/same-timestamp '.version' dotfile.
    $sourceVersionMarker = Join-Path $Source '.version'
    if (Test-Path -LiteralPath $sourceVersionMarker -PathType Leaf) {
        Copy-Item -LiteralPath $sourceVersionMarker -Destination $Destination -Force -ErrorAction Stop
    }
    # The apphost 'careai-corina-service.exe' hits the same problem: it is a near-constant
    # -size native stub, so with the fixed 1980 ZIP timestamp robocopy skips it as "Same"
    # and leaves a stale version resource on an otherwise-updated install. Force-copy it too.
    $sourceExe = Join-Path $Source 'careai-corina-service.exe'
    if (Test-Path -LiteralPath $sourceExe -PathType Leaf) {
        Copy-Item -LiteralPath $sourceExe -Destination $Destination -Force -ErrorAction Stop
    }
}

function Stop-CorinaServiceProcess {
    param([Parameter(Mandatory)][string]$Name)
    try {
        $service = Get-CimInstance Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue
        if ($service -and $service.ProcessId -gt 0) { Stop-Process -Id $service.ProcessId -Force -ErrorAction SilentlyContinue }
    } catch { }
}

function Set-CorinaServiceEnvironment {
    param([Parameter(Mandatory)][string]$Name, [string]$RegistryInstance)
    $values = @("DOTNET_ENVIRONMENT=$($script:CorinaDotNetEnvironment)")
    if ($RegistryInstance) { $values += "CorinaRegistryInstance=$RegistryInstance" }
    New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$Name" -Name Environment -PropertyType MultiString -Value $values -Force | Out-Null
}

$corinaRegistryInstance = $Instance
if ([string]::IsNullOrWhiteSpace($corinaRegistryInstance)) { $corinaRegistryInstance = [Environment]::GetEnvironmentVariable('CorinaRegistryInstance', [EnvironmentVariableTarget]::Process) }
if ($corinaRegistryInstance) {
    $corinaRegistryInstance = $corinaRegistryInstance.Trim()
    if ($corinaRegistryInstance -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9_-]*[A-Za-z0-9])?$') { throw 'Invalid Corina registry instance.' }
    $env:CorinaRegistryInstance = $corinaRegistryInstance
}

$serviceName = if ($corinaRegistryInstance) { "$($script:CorinaServiceBaseName)-$corinaRegistryInstance" } else { $script:CorinaServiceBaseName }
$taskName = if ($corinaRegistryInstance) { "$($script:CorinaTaskBaseName)-$corinaRegistryInstance" } else { $script:CorinaTaskBaseName }
$regPath = $script:CorinaRegistryRoot
if ($corinaRegistryInstance) { $regPath = Join-Path $regPath $corinaRegistryInstance }
$stateRoot = if ($corinaRegistryInstance) { Join-Path (Join-Path $env:ProgramData $script:CorinaProgramDataRoot) $corinaRegistryInstance } else { Join-Path (Join-Path $env:ProgramData $script:CorinaProgramDataRoot) 'default' }
$logDir = Join-Path $stateRoot 'Logs'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logPath = Join-Path $logDir 'update.log'

function Write-CorinaLog {
    param([Parameter(Mandatory)][string]$Message, [ValidateSet('INFO','STEP','OK','WARN','FAIL')][string]$Level='INFO')
    $line = "[$([DateTimeOffset]::Now.ToString('o'))] [$Level] $Message"
    $line | Out-File -LiteralPath $logPath -Append -Encoding utf8
    Write-Host $line
}

function Invoke-CorinaLegacySignedMigration {
    param(
        [Parameter(Mandatory)][string]$ServiceName,
        [Parameter(Mandatory)][string]$RegistryPath,
        [Parameter(Mandatory)][string]$StateRoot,
        [string]$RegistryInstance
    )

    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'The Corina updater migration must run as Administrator or SYSTEM.'
    }

    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $service -or $service.Status -ne 'Running') {
        Write-CorinaLog "Automatic secure-updater migration deferred because service '$ServiceName' is not running." WARN
        return
    }
    $token = Get-CorinaRegistryValue -Path $RegistryPath -Name CorinaAgentToken
    if ([string]::IsNullOrWhiteSpace([string]$token)) {
        Write-CorinaLog 'Automatic secure-updater migration deferred because CorinaAgentToken is missing; the running service was not changed.' WARN
        return
    }

    $bootstrapTrusted = ConvertTo-CorinaThumbprintList -Values $script:LegacyBootstrapTrustedSignerThumbprints
    $migrationRoot = Join-Path $StateRoot ("Migration\" + [guid]::NewGuid().ToString('N'))
    $manifestPath = Join-Path $migrationRoot $script:CorinaManifestFileName
    $manifestLock = $null
    $installerLock = $null
    try {
        New-Item -ItemType Directory -Path $migrationRoot -Force | Out-Null
        Write-CorinaLog 'Authenticating the signed release used to migrate the legacy updater.' STEP
        Receive-CorinaFile -Uri $script:CorinaManifestUri -Destination $manifestPath
        $manifestLock = [IO.File]::Open(
            $manifestPath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::Read
        )

        $minimumSequence = $script:LegacyBootstrapMinimumSequence
        $storedSequence = Get-CorinaRegistryValue -Path $RegistryPath -Name AcceptedManifestSequence
        $parsedStoredSequence = [UInt64]0
        if ([UInt64]::TryParse([string]$storedSequence, [ref]$parsedStoredSequence) -and
            $parsedStoredSequence -gt $minimumSequence) {
            $minimumSequence = $parsedStoredSequence
        }
        $manifest = Read-CorinaReleaseManifest `
            -Path $manifestPath `
            -AllowedThumbprints $bootstrapTrusted `
            -MinimumSequence $minimumSequence
        if ([version]$manifest.ReleaseVersion -lt [version]'1.3.4') {
            throw "The signed migration release v$($manifest.ReleaseVersion) is older than the minimum secure release v1.3.4."
        }

        $releaseSigner = @([string]$manifest._VerifiedSignerThumbprint)
        $installerAsset = [hashtable]$manifest.Assets.InstallerScript
        $installerPath = Join-Path $migrationRoot ([string]$installerAsset.FileName)
        Receive-CorinaAsset `
            -Asset $installerAsset `
            -Destination $installerPath `
            -AllowedThumbprints $releaseSigner `
            -RequireAuthenticode
        $installerLock = [IO.File]::Open(
            $installerPath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::Read
        )

        Write-CorinaLog "Running authenticated installer v$($manifest.ReleaseVersion); the existing token will be preserved." STEP
        if ($RegistryInstance) {
            & $installerPath -Instance $RegistryInstance -TrustedSignerThumbprints $bootstrapTrusted
        }
        else {
            & $installerPath -TrustedSignerThumbprints $bootstrapTrusted
        }
        if (-not $?) { throw 'The authenticated installer returned a failure status.' }

        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if (-not $service -or $service.Status -ne 'Running') {
            throw "Service '$ServiceName' was not running after secure-updater migration."
        }
        Write-CorinaLog "Legacy updater migration to signed release v$($manifest.ReleaseVersion) completed." OK
    }
    finally {
        if ($installerLock) { $installerLock.Dispose() }
        if ($manifestLock) { $manifestLock.Dispose() }
        Remove-Item -LiteralPath $migrationRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path -LiteralPath $regPath)) { throw "Corina registry state is missing: $regPath" }
if ([string]::IsNullOrWhiteSpace($PSCommandPath)) { throw 'daily-updater.ps1 must run from a file on disk.' }
$currentScriptSignature = Get-AuthenticodeSignature -FilePath $PSCommandPath
if ($script:IsLegacyUnsignedBootstrap -and
    $currentScriptSignature.Status -eq [Management.Automation.SignatureStatus]::NotSigned) {
    try {
        Invoke-CorinaLegacySignedMigration `
            -ServiceName $serviceName `
            -RegistryPath $regPath `
            -StateRoot $stateRoot `
            -RegistryInstance $corinaRegistryInstance
    }
    catch {
        Write-CorinaLog "Automatic secure-updater migration failed before completion: $_" FAIL
        try {
            $current = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($current -and $current.Status -ne 'Running') {
                Start-Service -Name $serviceName -ErrorAction SilentlyContinue
            }
        }
        catch { }
    }
    return
}

$storedTrusted = @((Get-CorinaRegistryValue -Path $regPath -Name TrustedSignerThumbprints))
if ($storedTrusted.Count -eq 0) { $storedTrusted = $script:BuiltInTrustedSignerThumbprints }
$trusted = ConvertTo-CorinaThumbprintList -Values $storedTrusted
$null = Assert-CorinaSignedFile -Path $PSCommandPath -AllowedThumbprints $trusted
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'The Corina updater must run as Administrator or SYSTEM.' }

try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 }
catch { throw "TLS 1.2 could not be enabled: $_" }

$mutex = [Threading.Mutex]::new($false, $(if ($corinaRegistryInstance) { "Global\CorinaDailyUpdater-$corinaRegistryInstance" } else { 'Global\CorinaDailyUpdater' }))
$mutexAcquired = $false
try { $mutexAcquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5)) }
catch [Threading.AbandonedMutexException] { $mutexAcquired = $true }
if (-not $mutexAcquired) { Write-CorinaLog 'Another updater run is active; this trigger is redundant.' WARN; $mutex.Dispose(); return }

$stagingRoot = Join-Path $stateRoot ("Staging\" + [guid]::NewGuid().ToString('N'))
$extractDir = Join-Path $stagingRoot 'service'
$backupDir = Join-Path $stateRoot 'Backup'
$deploymentStarted = $false
$haveBackup = $false
$servicePath = $null
$installDir = $null
$trustStateWritten = $false

try {
    Write-CorinaLog 'Secure update started.' STEP
    New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null
    $manifestPath = Join-Path $stagingRoot $script:CorinaManifestFileName
    Receive-CorinaFile -Uri $script:CorinaManifestUri -Destination $manifestPath
    $minimumSequence = [UInt64]0
    $storedSequence = Get-CorinaRegistryValue -Path $regPath -Name AcceptedManifestSequence
    [UInt64]::TryParse([string]$storedSequence, [ref]$minimumSequence) | Out-Null
    $manifest = Read-CorinaReleaseManifest -Path $manifestPath -AllowedThumbprints $trusted -MinimumSequence $minimumSequence
    $releaseSigner = @([string]$manifest._VerifiedSignerThumbprint)
    Write-CorinaLog "Authenticated manifest v$($manifest.ReleaseVersion), sequence $($manifest.Sequence)." OK

    $installedVersionText = [string](Get-CorinaRegistryValue -Path $regPath -Name InstalledReleaseVersion)
    if ($installedVersionText -match '^\d+\.\d+\.\d+$' -and [version]$manifest.ReleaseVersion -lt [version]$installedVersionText) { throw "Release v$($manifest.ReleaseVersion) is older than installed v$installedVersionText." }
    if ([UInt64]$manifest.Sequence -eq $minimumSequence -and $installedVersionText -ceq [string]$manifest.ReleaseVersion) {
        Write-CorinaLog "Already on authenticated release v$installedVersionText; no deployment is needed." OK
        return
    }

    $downloads = @{}
    foreach ($role in @('InstallerScript','UpdaterScript','TaskHelperScript','UninstallerScript')) {
        $asset = [hashtable]$manifest.Assets[$role]
        $destination = Join-Path $stagingRoot ([string]$asset.FileName)
        Receive-CorinaAsset -Asset $asset -Destination $destination -AllowedThumbprints $releaseSigner -RequireAuthenticode
        $downloads[$role] = $destination
    }
    $packageAsset = [hashtable]$manifest.Assets.ServicePackage
    $packagePath = Join-Path $stagingRoot ([string]$packageAsset.FileName)
    Receive-CorinaAsset -Asset $packageAsset -Destination $packagePath -AllowedThumbprints $trusted
    Expand-CorinaArchiveSafely -ArchivePath $packagePath -Destination $extractDir

    $service = Get-CimInstance Win32_Service -Filter "Name='$serviceName'" -ErrorAction Stop
    if (-not $service) { throw "Service '$serviceName' is not installed." }
    $match = [regex]::Match([string]$service.PathName, '^[\s"]*(?<exe>[^"\r\n]+?\.exe)')
    if (-not $match.Success) { throw "Could not parse service executable path '$($service.PathName)'." }
    $servicePath = $match.Groups['exe'].Value.Trim()
    $installDir = Split-Path -Parent $servicePath
    $exeName = Split-Path -Leaf $servicePath
    $stagedExe = Join-Path $extractDir $exeName
    if (-not (Test-Path -LiteralPath $stagedExe -PathType Leaf)) { throw "Service package is missing '$exeName'." }
    $versionMarker = Join-Path $extractDir '.version'
    if (-not (Test-Path -LiteralPath $versionMarker -PathType Leaf) -or
        (Get-Content -LiteralPath $versionMarker -Raw).Trim() -cne [string]$manifest.ReleaseVersion) {
        throw 'Service package .version does not match the signed manifest release version.'
    }
    $null = Assert-CorinaSignedFile -Path $stagedExe -AllowedThumbprints $releaseSigner
    if ((Get-ChildItem -LiteralPath $extractDir -File -Recurse).Count -lt 5) { throw 'Service package is unexpectedly incomplete.' }
    $stagedUpdateDir = Join-Path $extractDir 'Update'
    if (Test-Path -LiteralPath $stagedUpdateDir) { Remove-Item -LiteralPath $stagedUpdateDir -Recurse -Force }
    New-Item -ItemType Directory -Path $stagedUpdateDir -Force | Out-Null
    Copy-Item -LiteralPath $downloads.UpdaterScript -Destination (Join-Path $stagedUpdateDir 'daily-updater.ps1')
    Copy-Item -LiteralPath $downloads.TaskHelperScript -Destination (Join-Path $stagedUpdateDir 'ensure-updater-task.ps1')
    Copy-Item -LiteralPath $downloads.UninstallerScript -Destination (Join-Path $stagedUpdateDir 'uninstall.ps1')

    $token = Get-CorinaRegistryValue -Path $regPath -Name CorinaAgentToken
    if ([string]::IsNullOrWhiteSpace([string]$token)) { Write-CorinaLog 'CorinaAgentToken is still missing; the service may not authenticate.' WARN }

    if (Test-Path -LiteralPath $backupDir) { Remove-Item -LiteralPath $backupDir -Recurse -Force }
    Copy-CorinaTree -Source $installDir -Destination $backupDir
    $haveBackup = $true

    Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Stop-CorinaServiceProcess -Name $serviceName
    $deploymentStarted = $true
    Copy-CorinaTree -Source $extractDir -Destination $installDir -Mirror
    $installedVersionMarker = Join-Path $installDir '.version'
    if (-not (Test-Path -LiteralPath $installedVersionMarker -PathType Leaf) -or
        (Get-Content -LiteralPath $installedVersionMarker -Raw).Trim() -cne [string]$manifest.ReleaseVersion) {
        throw 'Deployed service .version does not match the signed manifest release version.'
    }
    $null = Assert-CorinaSignedFile -Path $servicePath -AllowedThumbprints $releaseSigner
    # Verify the ACTUAL deployed exe version, not just the '.version' marker. A stale
    # apphost that slipped past the copy would otherwise pass every check and leave the
    # service exe reporting the wrong version.
    $deployedExeVersion = [string]([Diagnostics.FileVersionInfo]::GetVersionInfo($servicePath).FileVersion)
    if ([string]::IsNullOrWhiteSpace($deployedExeVersion)) {
        # Unreadable version metadata must not hard-block a clinic's updates forever: an
        # empty value would never equal the expected string and would roll back every run.
        # The signature and '.version' assertions above still gate this deployment.
        Write-CorinaLog 'Deployed exe reports no FileVersion; skipping the exe version comparison.' WARN
    } elseif ($deployedExeVersion -ne "$([string]$manifest.ReleaseVersion).0") {
        throw "Deployed exe FileVersion '$deployedExeVersion' does not match release v$($manifest.ReleaseVersion)."
    } else {
        Write-CorinaLog "Deployed exe FileVersion $deployedExeVersion matches release v$($manifest.ReleaseVersion)." OK
    }
    foreach ($role in @('UpdaterScript','TaskHelperScript','UninstallerScript')) {
        $targetName = [string]$manifest.Assets[$role].FileName
        $targetPath = Join-Path (Join-Path $installDir 'Update') $targetName
        if ((Get-Item -LiteralPath $targetPath).Length -ne [long]$manifest.Assets[$role].Size -or
            (Get-CorinaSha256 -Path $targetPath) -cne ([string]$manifest.Assets[$role].Sha256).ToUpperInvariant()) {
            throw "Deployed $role failed its post-copy size/hash check."
        }
        $null = Assert-CorinaSignedFile -Path $targetPath -AllowedThumbprints $releaseSigner
    }
    Set-CorinaServiceEnvironment -Name $serviceName -RegistryInstance $corinaRegistryInstance
    Start-Service -Name $serviceName

    $healthy = $false
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt 30) {
        Start-Sleep -Seconds 2
        $current = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($current -and $current.Status -eq 'Running') { $healthy = $true; break }
    }
    if ($healthy) { Start-Sleep -Seconds 5; $current = Get-Service -Name $serviceName -ErrorAction SilentlyContinue; $healthy = [bool]($current -and $current.Status -eq 'Running') }
    if (-not $healthy) { throw "Service '$serviceName' failed its post-update health check." }

    $updateDir = Join-Path $installDir 'Update'
    . (Join-Path $updateDir 'ensure-updater-task.ps1')
    $legacyTasks = @()
    $legacyShims = @($(if ($corinaRegistryInstance) {
        "C:\Scripts\$($script:CorinaLegacyShimBaseName)-$corinaRegistryInstance.ps1"
    } else { "C:\Scripts\$($script:CorinaLegacyShimBaseName).ps1" }))
    if ($corinaRegistryInstance -and -not (Get-Service -Name $script:CorinaServiceBaseName -ErrorAction SilentlyContinue)) { $legacyTasks += $script:CorinaTaskBaseName; $legacyShims += "C:\Scripts\$($script:CorinaLegacyShimBaseName).ps1" }
    $logCallback = { param($Message) Write-CorinaLog $Message INFO }
    Ensure-CorinaUpdaterTask -Instance $corinaRegistryInstance -TaskName $taskName -UpdateRoot $updateDir -RegistryPath $regPath -LegacyTaskNames $legacyTasks -LegacyShimPaths $legacyShims -Log $logCallback

    # Bounded rotation drops historical certificates after the next signer has
    # successfully authenticated a complete release.
    $mergedTrusted = ConvertTo-CorinaThumbprintList -Values @($manifest._VerifiedSignerThumbprint + @($manifest.NextSignerThumbprints))
    New-ItemProperty -Path $regPath -Name TrustedSignerThumbprints -PropertyType MultiString -Value $mergedTrusted -Force | Out-Null
    $trustStateWritten = $true
    New-ItemProperty -Path $regPath -Name AcceptedManifestSequence -PropertyType QWord -Value ([UInt64]$manifest.Sequence) -Force | Out-Null
    New-ItemProperty -Path $regPath -Name InstalledReleaseVersion -PropertyType String -Value ([string]$manifest.ReleaseVersion) -Force | Out-Null
    New-ItemProperty -Path $regPath -Name AcceptedManifestSha256 -PropertyType String -Value (Get-CorinaSha256 -Path $manifestPath) -Force | Out-Null
    Write-CorinaLog "Update to v$($manifest.ReleaseVersion) completed and passed health checks." OK

}
catch {
    Write-CorinaLog "Update failed: $_" FAIL
    if ($trustStateWritten) {
        try { New-ItemProperty -Path $regPath -Name TrustedSignerThumbprints -PropertyType MultiString -Value $trusted -Force | Out-Null }
        catch { Write-CorinaLog "Could not restore previous signer trust state: $_" WARN }
    }
    if ($deploymentStarted -and $haveBackup -and $installDir) {
        try {
            Write-CorinaLog 'Rolling back the complete previous installation.' WARN
            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
            Stop-CorinaServiceProcess -Name $serviceName
            Copy-CorinaTree -Source $backupDir -Destination $installDir -Mirror
            Set-CorinaServiceEnvironment -Name $serviceName -RegistryInstance $corinaRegistryInstance
            Start-Service -Name $serviceName -ErrorAction Stop
            Write-CorinaLog 'Rollback completed and previous service restarted.' OK
        } catch { Write-CorinaLog "Rollback failed: $_" FAIL }
    } else {
        try {
            $current = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($current -and $current.Status -ne 'Running') { Start-Service -Name $serviceName -ErrorAction SilentlyContinue }
        } catch { }
    }
    throw
}
finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
    if ($mutexAcquired) { try { $mutex.ReleaseMutex() } catch { }; $mutex.Dispose() }
}
