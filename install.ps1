# Corina Service production installer.
#
# SECURITY CONTRACT
# - This file is a byte-stable release asset. Download it to disk, verify its
#   release-published SHA-256 and Authenticode signer, then invoke it with `&`.
# - It deliberately refuses network-pipeline/in-memory execution.
# - All network-delivered code and packages are authenticated by a signed,
#   data-only release manifest and exact size/SHA-256 checks before use.

[CmdletBinding()]
param(
    [ValidatePattern('^[A-Za-z0-9](?:[A-Za-z0-9_-]*[A-Za-z0-9])?$')]
    [string]$Instance,

    [string[]]$TrustedSignerThumbprints
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# These two values are intentionally compile-time constants. The release build
# may replace them with the staging pair before signing; callers cannot supply a
# repository or arbitrary manifest URL.
$script:CorinaReleaseChannel = 'production'
$script:CorinaReleaseRepository = 'Care-AI-Inc/careai-corina-service-releases'
$script:CorinaServiceSourceRepository = 'Care-AI-Inc/careai-corina-service'
$script:CorinaScriptsSourceRepository = 'Care-AI-Inc/careai-corina-service-releases'
$script:CorinaServiceBaseName = 'CorinaService'
$script:CorinaTaskBaseName = 'CorinaProdDailyUpdater'
$script:CorinaDisplayName = 'Corina Service (Production)'
$script:CorinaProgramFilesLeaf = 'CorinaService'
$script:CorinaProgramDataRoot = 'CareAI\CorinaService'
$script:CorinaRegistryRoot = 'HKLM:\SOFTWARE\CareAI\CorinaService'
$script:CorinaBackendBaseUrl = 'https://backend.agent.caregp.com.au'
$script:CorinaDotNetEnvironment = 'Production'
$script:CorinaLegacyShimBaseName = 'run-daily-updater-prod'
$script:CorinaManifestFileName = "corina-$($script:CorinaReleaseChannel).ps1"
$script:CorinaInstallerReleaseVersion = '__CORINA_RELEASE_VERSION__'
$script:CorinaInstallerReleaseSequence = '__CORINA_RELEASE_SEQUENCE__'
if ($script:CorinaInstallerReleaseVersion -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$' -or
    $script:CorinaInstallerReleaseSequence -notmatch '^[1-9]\d*$') {
    throw 'install.ps1 contains unresolved or invalid release identity placeholders.'
}
$script:CorinaInstallerReleaseTag = if ($script:CorinaReleaseChannel -eq 'staging') {
    "staging-v$($script:CorinaInstallerReleaseVersion)"
} else { "v$($script:CorinaInstallerReleaseVersion)" }
$script:CorinaManifestUri = "https://github.com/$($script:CorinaReleaseRepository)/releases/download/$($script:CorinaInstallerReleaseTag)/$($script:CorinaManifestFileName)"

# The publish job replaces the sentinel with one or more current certificate
# thumbprints in final release assets. Source remains fail-closed until either
# that happens or a previously verified bootstrap passes -TrustedSignerThumbprints.
$script:BuiltInTrustedSignerThumbprints = @('__CORINA_RELEASE_SIGNER_THUMBPRINTS__')

$script:CareAiPublisher = @{
    CommonName   = 'CARE AI PTY LTD'
    Organisation = 'CARE AI PTY LTD'
    Country      = 'AU'
    SerialNumber = '38681904512'
}

function Get-CorinaCertificateSubjectAttribute {
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string[]]$Names
    )
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

    $result = @($Values | ForEach-Object {
        if ($null -ne $_) { ([string]$_).Replace(' ', '').ToUpperInvariant() }
    } | Where-Object { $_ -match '^[0-9A-F]{40}$' } | Select-Object -Unique)

    $invalid = @($Values | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_) -and
        ([string]$_).Replace(' ', '').ToUpperInvariant() -notmatch '^[0-9A-F]{40}$' -and
        [string]$_ -ne '__CORINA_RELEASE_SIGNER_THUMBPRINTS__'
    })
    if ($invalid.Count -gt 0) {
        throw "Invalid trusted signer thumbprint value(s). Expected 40 hexadecimal characters."
    }
    if (-not $AllowEmpty -and $result.Count -eq 0) {
        throw 'No trusted release signer thumbprints were supplied or embedded. Installation is blocked (fail closed).'
    }
    return ,([string[]]$result)
}

function Test-CorinaCertificatePublisher {
    param([Parameter(Mandatory)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    $simpleName = $Certificate.GetNameInfo(
        [Security.Cryptography.X509Certificates.X509NameType]::SimpleName,
        $false
    )
    if ((ConvertTo-CorinaIdentityValue $simpleName) -cne (ConvertTo-CorinaIdentityValue $script:CareAiPublisher.CommonName)) { return $false }
    $organisation = Get-CorinaCertificateSubjectAttribute -Subject $Certificate.Subject -Names @('O')
    $country = Get-CorinaCertificateSubjectAttribute -Subject $Certificate.Subject -Names @('C')
    $serial = Get-CorinaCertificateSubjectAttribute -Subject $Certificate.Subject -Names @('SERIALNUMBER','OID.2.5.4.5','2.5.4.5')
    if ((ConvertTo-CorinaIdentityValue $organisation) -cne (ConvertTo-CorinaIdentityValue $script:CareAiPublisher.Organisation) -or
        (ConvertTo-CorinaIdentityValue $country) -cne (ConvertTo-CorinaIdentityValue $script:CareAiPublisher.Country) -or
        (ConvertTo-CorinaIdentityValue $serial) -cne $script:CareAiPublisher.SerialNumber) { return $false }

    $codeSigningOid = '1.3.6.1.5.5.7.3.3'
    $hasCodeSigningEku = $false
    foreach ($extension in $Certificate.Extensions) {
        if ($extension.Oid.Value -eq '2.5.29.37') {
            $eku = [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]$extension
            $hasCodeSigningEku = [bool]($eku.EnhancedKeyUsages | Where-Object { $_.Value -eq $codeSigningOid })
        }
    }
    return $hasCodeSigningEku
}

function Assert-CorinaSignedFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$AllowedThumbprints
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Signed file not found: $Path"
    }
    $signature = Get-AuthenticodeSignature -FilePath $Path
    if ($signature.Status -ne [Management.Automation.SignatureStatus]::Valid -or -not $signature.SignerCertificate) {
        throw "Authenticode validation failed for '$Path': $($signature.Status) $($signature.StatusMessage)"
    }
    $thumbprint = $signature.SignerCertificate.Thumbprint.Replace(' ', '').ToUpperInvariant()
    if ($thumbprint -notin $AllowedThumbprints) {
        throw "The signer of '$Path' is not in the trusted release signer allowlist (thumbprint $thumbprint)."
    }
    if (-not (Test-CorinaCertificatePublisher -Certificate $signature.SignerCertificate)) {
        throw "The signer of '$Path' does not match the exact CARE AI publisher identity."
    }
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
        [long]$MaximumSize = 2147483648
    )

    $allowedKeys = @('FileName', 'Url', 'Sha256', 'Size')
    $unexpected = @($Asset.Keys | Where-Object { [string]$_ -notin $allowedKeys })
    if ($unexpected.Count -gt 0) { throw "Manifest asset '$Role' has unexpected field(s): $($unexpected -join ', ')." }

    $fileName = [string]$Asset.FileName
    if ([string]::IsNullOrWhiteSpace($fileName) -or $fileName -ne [IO.Path]::GetFileName($fileName) -or $fileName -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        throw "Manifest asset '$Role' has an unsafe FileName."
    }
    if ($ExpectedFileName -and $fileName -cne $ExpectedFileName) {
        throw "Manifest asset '$Role' must be named '$ExpectedFileName', not '$fileName'."
    }
    if ([string]$Asset.Sha256 -notmatch '^[0-9A-Fa-f]{64}$') {
        throw "Manifest asset '$Role' has an invalid SHA-256."
    }
    $size = 0L
    if (-not [long]::TryParse([string]$Asset.Size, [ref]$size) -or $size -le 0 -or $size -gt $MaximumSize) {
        throw "Manifest asset '$Role' has an invalid size."
    }

    $releaseTag = if ($script:CorinaReleaseChannel -eq 'staging') { "staging-v$ReleaseVersion" } else { "v$ReleaseVersion" }
    $expectedUri = "https://github.com/$($script:CorinaReleaseRepository)/releases/download/$releaseTag/$fileName"
    if ([string]$Asset.Url -cne $expectedUri) {
        throw "Manifest asset '$Role' must use its exact immutable release URL '$expectedUri'."
    }
}

function Read-CorinaReleaseManifest {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$AllowedThumbprints,
        [UInt64]$MinimumSequence = 0,
        [string]$ExpectedReleaseVersion,
        [UInt64]$ExpectedSequence = 0
    )

    $manifestSignature = Assert-CorinaSignedFile -Path $Path -AllowedThumbprints $AllowedThumbprints

    # Parse only after Authenticode succeeds. SafeGetValue accepts a literal
    # hashtable but refuses commands, variables, subexpressions and executable
    # statements, so the signed .ps1 manifest is treated as data, not code.
    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw "Release manifest has parse errors: $($parseErrors[0].Message)" }
    $cleanBlockProperty = $ast.PSObject.Properties['CleanBlock']
    $hasCleanBlock = [bool]($cleanBlockProperty -and $cleanBlockProperty.Value)
    if ($ast.BeginBlock -or $ast.ProcessBlock -or $hasCleanBlock -or $ast.ParamBlock -or
        -not $ast.EndBlock -or $ast.EndBlock.Statements.Count -ne 1) {
        throw 'Release manifest must contain exactly one literal hashtable and no executable blocks.'
    }
    $statement = $ast.EndBlock.Statements[0]
    if ($statement -isnot [Management.Automation.Language.PipelineAst] -or
        $statement.PipelineElements.Count -ne 1 -or
        $statement.PipelineElements[0] -isnot [Management.Automation.Language.CommandExpressionAst] -or
        $statement.PipelineElements[0].Expression -isnot [Management.Automation.Language.HashtableAst]) {
        throw 'Release manifest root must be one literal hashtable.'
    }
    try { $manifest = [hashtable]$statement.PipelineElements[0].Expression.SafeGetValue() }
    catch { throw "Release manifest contains a non-literal value: $_" }

    $allowedTop = @('SchemaVersion', 'Channel', 'Repository', 'ReleaseVersion', 'Sequence', 'PublishedUtc', 'Source', 'Signer', 'NextSignerThumbprints', 'Assets')
    $unexpectedTop = @($manifest.Keys | Where-Object { [string]$_ -notin $allowedTop })
    if ($unexpectedTop.Count -gt 0) { throw "Release manifest has unexpected field(s): $($unexpectedTop -join ', ')." }
    $missingTop = @($allowedTop | Where-Object { -not $manifest.ContainsKey($_) })
    if ($missingTop.Count -gt 0) { throw "Release manifest is missing field(s): $($missingTop -join ', ')." }
    if ([int]$manifest.SchemaVersion -ne 1) { throw "Unsupported release manifest schema '$($manifest.SchemaVersion)'." }
    if ([string]$manifest.Channel -cne $script:CorinaReleaseChannel) { throw "Release manifest channel mismatch." }
    if ([string]$manifest.Repository -cne $script:CorinaReleaseRepository) { throw "Release manifest repository mismatch." }

    if ($manifest.Source -isnot [hashtable]) { throw 'Manifest Source must be a hashtable.' }
    $sourceKeys = @('Repository','Commit','Ref','ScriptsRepository','ScriptsCommit')
    $unexpectedSource = @($manifest.Source.Keys | Where-Object { [string]$_ -notin $sourceKeys })
    if ($unexpectedSource.Count -gt 0 -or @($sourceKeys | Where-Object { -not $manifest.Source.ContainsKey($_) }).Count -gt 0) {
        throw 'Manifest Source fields do not match schema 1.'
    }
    if ([string]$manifest.Source.Repository -cne $script:CorinaServiceSourceRepository -or
        [string]$manifest.Source.ScriptsRepository -cne $script:CorinaScriptsSourceRepository) {
        throw 'Manifest source repositories do not match this signed installer build.'
    }
    if ([string]$manifest.Source.Commit -cnotmatch '^[0-9a-f]{40}$' -or [string]$manifest.Source.ScriptsCommit -cnotmatch '^[0-9a-f]{40}$') {
        throw 'Manifest source commits must be lowercase 40-character Git object IDs.'
    }
    if ([string]::IsNullOrWhiteSpace([string]$manifest.Source.Ref) -or [string]$manifest.Source.Ref -match '[\x00-\x1F\x7F]' -or ([string]$manifest.Source.Ref).Length -gt 255) {
        throw 'Manifest source ref is invalid.'
    }

    if ($manifest.Signer -isnot [hashtable]) { throw 'Manifest Signer must be a hashtable.' }
    $signerKeys = @('Subject','CertificateThumbprint','TimestampRequired')
    $unexpectedSigner = @($manifest.Signer.Keys | Where-Object { [string]$_ -notin $signerKeys })
    if ($unexpectedSigner.Count -gt 0 -or @($signerKeys | Where-Object { -not $manifest.Signer.ContainsKey($_) }).Count -gt 0) {
        throw 'Manifest Signer fields do not match schema 1.'
    }
    $actualManifestThumbprint = $manifestSignature.SignerCertificate.Thumbprint.Replace(' ', '').ToUpperInvariant()
    if ([string]$manifest.Signer.Subject -cne [string]$manifestSignature.SignerCertificate.Subject -or
        ([string]$manifest.Signer.CertificateThumbprint).Replace(' ', '').ToUpperInvariant() -cne $actualManifestThumbprint) {
        throw 'Manifest Signer metadata does not match its actual Authenticode signature.'
    }
    if ($manifest.Signer.TimestampRequired -isnot [bool] -or -not [bool]$manifest.Signer.TimestampRequired -or -not $manifestSignature.TimeStamperCertificate) {
        throw 'Manifest must declare and contain a valid Authenticode timestamp.'
    }

    $version = [string]$manifest.ReleaseVersion
    if ($version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$') { throw "ReleaseVersion must be a three-part SemVer value." }
    $sequence = [UInt64]0
    if (-not [UInt64]::TryParse([string]$manifest.Sequence, [ref]$sequence) -or $sequence -eq 0 -or $sequence -lt $MinimumSequence) {
        throw "Manifest sequence '$($manifest.Sequence)' is invalid or older than accepted sequence '$MinimumSequence'."
    }
    if ($ExpectedReleaseVersion -and $version -cne $ExpectedReleaseVersion) {
        throw "Manifest version '$version' does not match this immutable installer release '$ExpectedReleaseVersion'."
    }
    if ($ExpectedSequence -gt 0 -and $sequence -ne $ExpectedSequence) {
        throw "Manifest sequence '$sequence' does not match this immutable installer sequence '$ExpectedSequence'."
    }
    $published = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse([string]$manifest.PublishedUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$published)) {
        throw 'Manifest PublishedUtc is not a valid ISO-8601 timestamp.'
    }
    if ($published -gt [DateTimeOffset]::UtcNow.AddHours(24)) { throw 'Manifest PublishedUtc is unreasonably far in the future.' }

    if ($manifest.Assets -isnot [hashtable]) { throw 'Manifest Assets must be a hashtable.' }
    $requiredAssets = @{
        InstallerScript   = 'install.ps1'
        UpdaterScript     = 'daily-updater.ps1'
        TaskHelperScript  = 'ensure-updater-task.ps1'
        UninstallerScript = 'uninstall.ps1'
        ServicePackage    = ''
    }
    $unexpectedAssets = @($manifest.Assets.Keys | Where-Object { [string]$_ -notin $requiredAssets.Keys })
    if ($unexpectedAssets.Count -gt 0) { throw "Manifest has unexpected asset role(s): $($unexpectedAssets -join ', ')." }
    foreach ($role in $requiredAssets.Keys) {
        if (-not $manifest.Assets.ContainsKey($role) -or $manifest.Assets[$role] -isnot [hashtable]) {
            throw "Manifest is missing required asset '$role'."
        }
        $maxSize = if ($role -eq 'ServicePackage') { 2147483648L } else { 5242880L }
        Assert-CorinaAssetDefinition -Asset $manifest.Assets[$role] -Role $role -ReleaseVersion $version -ExpectedFileName $requiredAssets[$role] -MaximumSize $maxSize
    }
    if ([IO.Path]::GetExtension([string]$manifest.Assets.ServicePackage.FileName) -cne '.zip') {
        throw 'ServicePackage must be a .zip file.'
    }

    $rawNextSigners = @($manifest.NextSignerThumbprints)
    if (@($rawNextSigners | Where-Object { [string]$_ -notmatch '^[0-9A-Fa-f]{40}$' }).Count -gt 0 -or $rawNextSigners.Count -gt 3) {
        throw 'Manifest NextSignerThumbprints must contain at most three valid thumbprints.'
    }
    $manifest.NextSignerThumbprints = ConvertTo-CorinaThumbprintList -Values $rawNextSigners -AllowEmpty
    # Internal, non-manifest state. The key is added only after the strict schema
    # check, so a publisher cannot smuggle it into the data document.
    $manifest['_VerifiedSignerThumbprint'] = $actualManifestThumbprint
    return $manifest
}

function Receive-CorinaFile {
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$Destination)

    $parent = Split-Path -Parent $Destination
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
    Invoke-WebRequest -Uri $Uri -OutFile $Destination -UseBasicParsing -TimeoutSec 600 -Headers @{ 'User-Agent' = 'CareAI-Corina-SecureInstaller/2' }
    if (-not (Test-Path -LiteralPath $Destination -PathType Leaf)) { throw "Download produced no file: $Uri" }
}

function Receive-CorinaManifestAsset {
    param(
        [Parameter(Mandatory)][hashtable]$Asset,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string[]]$AllowedThumbprints,
        [switch]$RequireAuthenticode
    )

    Receive-CorinaFile -Uri ([string]$Asset.Url) -Destination $Destination
    $actualSize = (Get-Item -LiteralPath $Destination).Length
    if ($actualSize -ne [long]$Asset.Size) {
        throw "Size verification failed for '$($Asset.FileName)': expected $($Asset.Size), got $actualSize."
    }
    $actualHash = Get-CorinaSha256 -Path $Destination
    if ($actualHash -cne ([string]$Asset.Sha256).ToUpperInvariant()) {
        throw "SHA-256 verification failed for '$($Asset.FileName)'."
    }
    if ($RequireAuthenticode) {
        $null = Assert-CorinaSignedFile -Path $Destination -AllowedThumbprints $AllowedThumbprints
    }
}

function Expand-CorinaArchiveSafely {
    param([Parameter(Mandatory)][string]$ArchivePath, [Parameter(Mandatory)][string]$Destination)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Recurse -Force }
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    $root = [IO.Path]::GetFullPath($Destination).TrimEnd('\') + '\'
    $archive = [IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        if ($archive.Entries.Count -lt 1) { throw 'Service package archive is empty.' }
        if ($archive.Entries.Count -gt 20000) { throw 'Service package contains more than 20,000 entries.' }
        $totalUncompressed = 0L
        $seenTargets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $archive.Entries) {
            if ([string]::IsNullOrWhiteSpace($entry.FullName)) { throw 'Service package contains an unnamed entry.' }
            $normalised = $entry.FullName.Replace('/', '\')
            if ([IO.Path]::IsPathRooted($normalised) -or $normalised -match '(^|\\)\.\.(\\|$)' -or $normalised.Contains(':')) {
                throw "Service package contains an unsafe path: $($entry.FullName)"
            }
            $target = [IO.Path]::GetFullPath((Join-Path $Destination $normalised))
            if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Service package entry escapes the staging directory: $($entry.FullName)"
            }
            if (-not $seenTargets.Add($target)) { throw "Service package contains a duplicate path: $($entry.FullName)" }
            if ($entry.Length -gt 2147483648L -or $totalUncompressed -gt (4294967296L - $entry.Length)) {
                throw 'Service package exceeds the 4 GiB uncompressed safety limit.'
            }
            $totalUncompressed += $entry.Length
            $externalAttributes = [UInt32]([Int64]$entry.ExternalAttributes -band 0xFFFFFFFFL)
            $unixFileType = (($externalAttributes -shr 16) -band 0xF000)
            if ($unixFileType -eq 0xA000) { throw "Service package contains a symbolic link: $($entry.FullName)" }
        }
    } finally { $archive.Dispose() }
    Expand-Archive -LiteralPath $ArchivePath -DestinationPath $Destination -Force
}

function Get-CorinaRegistryInstance {
    param([string]$ExplicitInstance)
    $value = $ExplicitInstance
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable('CorinaRegistryInstance', [EnvironmentVariableTarget]::Process)
    }
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $value = $value.Trim()
    if ($value -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9_-]*[A-Za-z0-9])?$') { throw "Invalid Corina registry instance '$value'." }
    $env:CorinaRegistryInstance = $value
    return $value
}

function Test-CorinaAdministrator {
    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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

function Copy-CorinaTree {
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination, [switch]$Mirror)
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    $mode = if ($Mirror) { '/MIR' } else { '/E' }
    # Release ZIP entries use deterministic 1980 timestamps. /IS prevents
    # robocopy from skipping changed same-size files such as .version.
    & robocopy $Source $Destination '*' $mode /IS /COPY:DAT /R:5 /W:3 /NFL /NDL /NP /NJH /NJS | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed copying '$Source' to '$Destination' (exit $LASTEXITCODE)." }
}

$trusted = if ($TrustedSignerThumbprints -and $TrustedSignerThumbprints.Count -gt 0) {
    ConvertTo-CorinaThumbprintList -Values $TrustedSignerThumbprints
} else {
    ConvertTo-CorinaThumbprintList -Values $script:BuiltInTrustedSignerThumbprints
}

if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
    throw 'install.ps1 must be executed from a signed file on disk; in-memory execution is not supported.'
}
$null = Assert-CorinaSignedFile -Path $PSCommandPath -AllowedThumbprints $trusted
if (-not (Test-CorinaAdministrator)) { throw 'You must run install.ps1 as Administrator.' }

try {
    $protocol = [Net.ServicePointManager]::SecurityProtocol
    [Net.ServicePointManager]::SecurityProtocol = $protocol -bor [Net.SecurityProtocolType]::Tls12
} catch { throw "TLS 1.2 could not be enabled: $_" }

$corinaRegistryInstance = Get-CorinaRegistryInstance -ExplicitInstance $Instance
$serviceName = if ($corinaRegistryInstance) { "$($script:CorinaServiceBaseName)-$corinaRegistryInstance" } else { $script:CorinaServiceBaseName }
$taskName = if ($corinaRegistryInstance) { "$($script:CorinaTaskBaseName)-$corinaRegistryInstance" } else { $script:CorinaTaskBaseName }
$serviceDisplayName = if ($corinaRegistryInstance) { $script:CorinaDisplayName.TrimEnd(')') + " - $corinaRegistryInstance)" } else { $script:CorinaDisplayName }
$installDir = if ($corinaRegistryInstance) {
    Join-Path (Join-Path $env:ProgramFiles $script:CorinaProgramFilesLeaf) $corinaRegistryInstance
} else { Join-Path $env:ProgramFiles $script:CorinaProgramFilesLeaf }
$updateDir = Join-Path $installDir 'Update'
$exeName = 'careai-corina-service.exe'
$exePath = Join-Path $installDir $exeName
$regPath = $script:CorinaRegistryRoot
if ($corinaRegistryInstance) { $regPath = Join-Path $regPath $corinaRegistryInstance }
$stateRoot = if ($corinaRegistryInstance) {
    Join-Path (Join-Path $env:ProgramData $script:CorinaProgramDataRoot) $corinaRegistryInstance
} else { Join-Path (Join-Path $env:ProgramData $script:CorinaProgramDataRoot) 'default' }
$stagingRoot = Join-Path $stateRoot ("Staging\" + [guid]::NewGuid().ToString('N'))
$extractDir = Join-Path $stagingRoot 'service'
$backupDir = Join-Path $stateRoot 'Backup'
$existingService = [bool](Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
$serviceWasStopped = $false
$haveBackup = $false
$previousRegistryTrusted = @()
if (Test-Path -LiteralPath $regPath) {
    $previousRegistryTrusted = @((Get-CorinaRegistryValue -Path $regPath -Name TrustedSignerThumbprints))
}
$trustStateWritten = $false

Write-Host "[*] Secure Corina Service installer ($($script:CorinaReleaseChannel))"
Write-Host "    -> Instance: $(if ($corinaRegistryInstance) { $corinaRegistryInstance } else { '<default>' })"

try {
    New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null
    $manifestPath = Join-Path $stagingRoot $script:CorinaManifestFileName
    Write-Host '[*] Download and authenticate the release manifest'
    Receive-CorinaFile -Uri $script:CorinaManifestUri -Destination $manifestPath

    $minimumSequence = [UInt64]0
    if (Test-Path -LiteralPath $regPath) {
        $storedSequence = Get-CorinaRegistryValue -Path $regPath -Name AcceptedManifestSequence
        [UInt64]::TryParse([string]$storedSequence, [ref]$minimumSequence) | Out-Null
    }
    $manifest = Read-CorinaReleaseManifest -Path $manifestPath -AllowedThumbprints $trusted -MinimumSequence $minimumSequence -ExpectedReleaseVersion $script:CorinaInstallerReleaseVersion -ExpectedSequence ([UInt64]$script:CorinaInstallerReleaseSequence)
    $releaseSigner = @([string]$manifest._VerifiedSignerThumbprint)
    $null = Assert-CorinaSignedFile -Path $PSCommandPath -AllowedThumbprints $releaseSigner
    Write-Host "    -> Authenticated release v$($manifest.ReleaseVersion), sequence $($manifest.Sequence)"

    $downloads = @{}
    foreach ($role in @('InstallerScript', 'UpdaterScript', 'TaskHelperScript', 'UninstallerScript')) {
        $asset = [hashtable]$manifest.Assets[$role]
        if ($role -eq 'InstallerScript') {
            if ((Get-Item -LiteralPath $PSCommandPath).Length -ne [long]$asset.Size -or
                (Get-CorinaSha256 -Path $PSCommandPath) -cne ([string]$asset.Sha256).ToUpperInvariant()) {
                throw 'This installer does not match InstallerScript in its immutable signed manifest.'
            }
            $downloads[$role] = $PSCommandPath
            continue
        }
        $destination = Join-Path $stagingRoot ([string]$asset.FileName)
        Write-Host "    -> Downloading and verifying $role"
        Receive-CorinaManifestAsset -Asset $asset -Destination $destination -AllowedThumbprints $releaseSigner -RequireAuthenticode
        $downloads[$role] = $destination
    }
    $packageAsset = [hashtable]$manifest.Assets.ServicePackage
    $packagePath = Join-Path $stagingRoot ([string]$packageAsset.FileName)
    Write-Host '    -> Downloading and verifying ServicePackage'
    Receive-CorinaManifestAsset -Asset $packageAsset -Destination $packagePath -AllowedThumbprints $trusted
    Expand-CorinaArchiveSafely -ArchivePath $packagePath -Destination $extractDir

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

    # Certificate rotation is accepted only after the current signer authenticates
    # the complete manifest and every executable artifact used by this install.
    # Bounded two-release rotation: retain the signer that authenticated this
    # release plus its announced successor(s), not every historical certificate.
    $mergedTrusted = ConvertTo-CorinaThumbprintList -Values @($manifest._VerifiedSignerThumbprint + @($manifest.NextSignerThumbprints))

    # Never recreate an existing key: the registry provider's New-Item -Force
    # REPLACES the key, destroying enrolment state (CorinaAgentToken, HaloGuid,
    # SamanthaBaseUrl) and every per-instance subkey beneath it.
    if (-not (Test-Path -LiteralPath $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    $defaultBackend = $script:CorinaBackendBaseUrl
    $baseUrl = Get-CorinaRegistryValue -Path $regPath -Name SamanthaBaseUrl
    if ([string]::IsNullOrWhiteSpace([string]$baseUrl)) {
        New-ItemProperty -Path $regPath -Name SamanthaBaseUrl -PropertyType String -Value $defaultBackend -Force | Out-Null
    }
    $token = Get-CorinaRegistryValue -Path $regPath -Name CorinaAgentToken
    if ([string]::IsNullOrWhiteSpace([string]$token)) {
        Write-Warning 'CorinaAgentToken is not configured; the service may not authenticate until enrolment is completed.'
    }

    if (Test-Path -LiteralPath $backupDir) { Remove-Item -LiteralPath $backupDir -Recurse -Force }
    if (Test-Path -LiteralPath $installDir) {
        Write-Host '[*] Backing up the current installation'
        Copy-CorinaTree -Source $installDir -Destination $backupDir
        $haveBackup = $true
    }

    if ($existingService) {
        Write-Host '[*] Stopping the existing service'
        Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Stop-CorinaServiceProcess -Name $serviceName
        $serviceWasStopped = $true
    }

    Write-Host '[*] Deploying verified service and update files'
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    Copy-CorinaTree -Source $extractDir -Destination $installDir -Mirror
    $installedVersionMarker = Join-Path $installDir '.version'
    if (-not (Test-Path -LiteralPath $installedVersionMarker -PathType Leaf) -or
        (Get-Content -LiteralPath $installedVersionMarker -Raw).Trim() -cne [string]$manifest.ReleaseVersion) {
        throw 'Deployed service .version does not match the signed manifest release version.'
    }
    foreach ($scriptName in @('daily-updater.ps1','ensure-updater-task.ps1','uninstall.ps1')) {
        $null = Assert-CorinaSignedFile -Path (Join-Path $updateDir $scriptName) -AllowedThumbprints $releaseSigner
    }
    $null = Assert-CorinaSignedFile -Path $exePath -AllowedThumbprints $releaseSigner

    if (-not $existingService) {
        & sc.exe create $serviceName binPath= "`"$exePath`"" start= auto obj= LocalSystem DisplayName= $serviceDisplayName | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "sc.exe create failed for '$serviceName' (exit $LASTEXITCODE)." }
    } else {
        # Never pass obj= for an existing service: clinics with credentialed
        # SMB/NAS shares run the service as a per-site user, and resetting it
        # to LocalSystem breaks their share access.
        & sc.exe config $serviceName binPath= "`"$exePath`"" start= auto DisplayName= $serviceDisplayName | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "sc.exe config failed for '$serviceName' (exit $LASTEXITCODE)." }
    }
    Set-CorinaServiceEnvironment -Name $serviceName -RegistryInstance $corinaRegistryInstance
    & sc.exe failure $serviceName reset= 86400 actions= restart/5000/restart/5000/restart/5000 | Out-Null
    & sc.exe failureflag $serviceName 1 | Out-Null

    Start-Service -Name $serviceName
    $healthy = $false
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt 30) {
        Start-Sleep -Seconds 2
        $current = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($current -and $current.Status -eq 'Running') { $healthy = $true; break }
    }
    if ($healthy) {
        Start-Sleep -Seconds 5
        $current = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        $healthy = [bool]($current -and $current.Status -eq 'Running')
    }
    if (-not $healthy) { throw "Service '$serviceName' did not stay running after installation." }

    # Store the already authenticated bounded trust set before registering the
    # first task, so an immediately due trigger can pass its preflight.
    New-ItemProperty -Path $regPath -Name TrustedSignerThumbprints -PropertyType MultiString -Value $mergedTrusted -Force | Out-Null
    $trustStateWritten = $true

    New-ItemProperty -Path $regPath -Name AcceptedManifestSequence -PropertyType QWord -Value ([UInt64]$manifest.Sequence) -Force | Out-Null
    New-ItemProperty -Path $regPath -Name InstalledReleaseVersion -PropertyType String -Value ([string]$manifest.ReleaseVersion) -Force | Out-Null
    New-ItemProperty -Path $regPath -Name ReleaseChannel -PropertyType String -Value $script:CorinaReleaseChannel -Force | Out-Null
    New-ItemProperty -Path $regPath -Name ReleaseRepository -PropertyType String -Value $script:CorinaReleaseRepository -Force | Out-Null
    New-ItemProperty -Path $regPath -Name AcceptedManifestSha256 -PropertyType String -Value (Get-CorinaSha256 -Path $manifestPath) -Force | Out-Null

    # Load only the already authenticated, exact-hash helper from disk.
    # The task is changed last so a failed install cannot strand an existing
    # clinic with a task that points at files subsequently rolled back.
    . (Join-Path $updateDir 'ensure-updater-task.ps1')
    $legacyTaskNames = @()
    $legacyShimPaths = @($(if ($corinaRegistryInstance) {
        "C:\Scripts\$($script:CorinaLegacyShimBaseName)-$corinaRegistryInstance.ps1"
    } else { "C:\Scripts\$($script:CorinaLegacyShimBaseName).ps1" }))
    if ($corinaRegistryInstance -and -not (Get-Service -Name $script:CorinaServiceBaseName -ErrorAction SilentlyContinue)) {
        $legacyTaskNames += $script:CorinaTaskBaseName
        $legacyShimPaths += "C:\Scripts\$($script:CorinaLegacyShimBaseName).ps1"
    }
    Ensure-CorinaUpdaterTask -Instance $corinaRegistryInstance -TaskName $taskName -UpdateRoot $updateDir -RegistryPath $regPath -LegacyTaskNames $legacyTaskNames -LegacyShimPaths $legacyShimPaths

    Write-Host "SUCCESS: Corina Service v$($manifest.ReleaseVersion) installed and the secure updater task is configured."

}
catch {
    Microsoft.PowerShell.Utility\Write-Error "Secure installation failed: $_" -ErrorAction Continue
    if ($trustStateWritten) {
        try {
            if ($previousRegistryTrusted.Count -gt 0) {
                New-ItemProperty -Path $regPath -Name TrustedSignerThumbprints -PropertyType MultiString -Value $previousRegistryTrusted -Force | Out-Null
            } else {
                Remove-ItemProperty -LiteralPath $regPath -Name TrustedSignerThumbprints -ErrorAction SilentlyContinue
            }
        } catch { Write-Warning "Could not restore previous signer trust state: $_" }
    }
    if ($haveBackup) {
        try {
            Write-Warning 'Restoring the previous installation from backup.'
            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
            Stop-CorinaServiceProcess -Name $serviceName
            Copy-CorinaTree -Source $backupDir -Destination $installDir -Mirror
            if ($existingService) {
                Set-CorinaServiceEnvironment -Name $serviceName -RegistryInstance $corinaRegistryInstance
                Start-Service -Name $serviceName -ErrorAction SilentlyContinue
            }
        } catch { Write-Warning "Rollback also encountered an error: $_" }
    } elseif (-not $existingService) {
        try {
            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
            & sc.exe delete $serviceName | Out-Null
        } catch { }
    } elseif ($serviceWasStopped) {
        try { Start-Service -Name $serviceName -ErrorAction SilentlyContinue } catch { }
    }
    throw
}
finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
