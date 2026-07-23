# Corina Service release scripts

This repository contains the source for the static Windows installer, updater,
scheduled-task helper and uninstaller used by Corina Service. Repository copies
are deliberately unsigned. The protected release workflow copies them to an
isolated output directory, applies channel/release constants, Authenticode-signs
the final bytes with timestamping, verifies the signatures, calculates hashes,
creates a data-only signed manifest, and publishes immutable release assets.

Do not execute scripts directly from `main`, pipe network content to `iex`, or
commit signed files back to Git. An Authenticode signature covers the exact byte
sequence; any edit, line-ending conversion or token substitution after signing
invalidates it.

## Runtime trust chain

1. The analytics-generated clinic wrapper is unique because it contains clinic
   enrolment data. Approve that wrapper by its exact SHA-256. It downloads the
   immutable `install.ps1` asset for an explicit release tag to disk.
2. The wrapper verifies the published SHA-256, a `Valid` timestamped
   Authenticode signature, and an allowed CARE AI leaf certificate thumbprint. It
   invokes the file with `&`, passing `-TrustedSignerThumbprints` and optional
   `-Instance`. It never uses `irm | iex`.
3. `install.ps1` repeats its own signature check. Its release version and manifest
   sequence are baked into the signed final file. It downloads the manifest from
   the same immutable release tag and requires the exact version and sequence.
4. The signed manifest is parsed only after Authenticode validation. The parser
   accepts one literal hashtable via `SafeGetValue`; commands, expressions,
   variables and unknown schema fields are rejected.
5. Every asset URL must exactly match the expected repository, channel tag and
   filename. Downloads are checked against manifest size and SHA-256. Each script
   and the extracted service executable must also have a valid CARE AI
   Authenticode signature from a currently trusted thumbprint.
6. The service package is inspected before extraction: rooted/traversal/ADS paths,
   duplicate case-insensitive paths, symbolic links, more than 20,000 entries and
   more than 4 GiB uncompressed data are rejected.
7. The installed task runs as SYSTEM, but its action first checks the installed
   updater signature, thumbprint, code-signing EKU and CARE AI publisher identity.
   Only then does it invoke `daily-updater.ps1`. There is no execution-policy
   bypass and no downloaded/generated shim.
8. The updater uses the signed `releases/latest/download/corina-<channel>.ps1`
   pointer only for discovery. Registry state enforces a monotonic sequence and
   release version. It verifies all assets before stopping the service, mirrors
   the complete staged tree so stale DLLs cannot survive, health-checks the new
   service, and restores the previous mirrored backup on failure.

TLS protects transport. The signed manifest, hashes, signatures and pinned
publisher/thumbprints provide the artifact trust boundary; a GitHub release or
account change without a trusted signing key cannot introduce executable bytes.

## Manifest schema (version 1)

The final Authenticode-signed `corina-production.ps1` or
`corina-staging.ps1` contains exactly one literal hashtable:

```powershell
@{
    SchemaVersion = 1
    Channel = 'production'
    Repository = 'Care-AI-Inc/careai-corina-service-releases'
    ReleaseVersion = '1.4.0'
    Sequence = 1000004000000L
    PublishedUtc = '2026-07-22T00:00:00.000Z'
    Source = @{
        Repository = 'Care-AI-Inc/careai-corina-service'
        Commit = '<40 lowercase hex>'
        Ref = 'refs/heads/main'
        ScriptsRepository = 'Care-AI-Inc/careai-corina-service-releases'
        ScriptsCommit = '<40 lowercase hex>'
    }
    Signer = @{
        Subject = '<exact manifest certificate subject>'
        CertificateThumbprint = '<40 hex>'
        TimestampRequired = $true
    }
    NextSignerThumbprints = @('<optional next 40-hex thumbprint>')
    Assets = @{
        ServicePackage = @{ FileName = 'corina-1.4.0-win-x64.zip'; Url = '<immutable tag URL>'; Sha256 = '<64 hex>'; Size = 1L }
        InstallerScript = @{ FileName = 'install.ps1'; Url = '<immutable tag URL>'; Sha256 = '<64 hex>'; Size = 1L }
        UpdaterScript = @{ FileName = 'daily-updater.ps1'; Url = '<immutable tag URL>'; Sha256 = '<64 hex>'; Size = 1L }
        TaskHelperScript = @{ FileName = 'ensure-updater-task.ps1'; Url = '<immutable tag URL>'; Sha256 = '<64 hex>'; Size = 1L }
        UninstallerScript = @{ FileName = 'uninstall.ps1'; Url = '<immutable tag URL>'; Sha256 = '<64 hex>'; Size = 1L }
    }
}
```

Production tags are `v<version>` in
`Care-AI-Inc/careai-corina-service-releases`. Staging tags are
`staging-v<version>` in
`Care-AI-Inc/careai-corina-service-staging-releases`. The protected build uses
this production-named repository as the script source for both channels, records
that source commit in the manifest, transforms the channel constants, rejects all
unresolved placeholders/production leakage, and signs only the transformed copy.

## Installed layout and state

Production defaults:

- Service: `CorinaService` or `CorinaService-<instance>`
- Task: `CorinaProdDailyUpdater` or `CorinaProdDailyUpdater-<instance>`
- Files: `C:\Program Files\CorinaService[\<instance>]`
- Signed management files: the `Update` child directory
- Update state/log/backup: `C:\ProgramData\CareAI\CorinaService\default` or
  `C:\ProgramData\CareAI\CorinaService\<instance>`
- Registry: `HKLM\SOFTWARE\CareAI\CorinaService[\<instance>]`

Staging final assets are build-transformed to the corresponding
`CorinaService-Staging`, `CorinaStagingDailyUpdater`,
`CorinaService-Staging` filesystem/registry roots, staging backend and
`DOTNET_ENVIRONMENT=Staging`. Callers cannot select a channel or repository at
runtime.

Registry security state includes `TrustedSignerThumbprints` (REG_MULTI_SZ),
`AcceptedManifestSequence` (REG_QWORD), `InstalledReleaseVersion`, channel,
repository and accepted manifest hash. The uninstaller retains clinic enrolment
registry data deliberately but removes the service, task, installed files and
update/rollback state.

## Certificate rotation

Rotation is bounded and requires overlap:

1. Publish a fully verified release signed by the currently trusted certificate
   with the new certificate in `NextSignerThumbprints`. All executable assets in
   this release remain signed by the current certificate.
2. After the whole release verifies and passes its service health check, clients
   store only the actual manifest signer plus the announced successor(s).
3. Publish the next release signed by the new certificate. Once it verifies,
   clients store the new signer plus any newly announced successor; the old
   certificate drops out automatically.

Do not switch signing certificates without the overlap release. Do not use an
environment-variable trust override. Emergency revocation of a compromised
certificate requires a separately hash-pinned migration package/policy change;
a manifest signed only by an untrusted replacement is correctly rejected.

## Release invariants

- Build once, transform once, sign final bytes, then hash those signed bytes.
- Timestamp every Authenticode signature and fail if the timestamp is missing.
- Publish script, package and manifest assets under one immutable tag.
- Never replace/clobber an existing published asset or move a published tag.
- `latest` is only the updater discovery pointer; bootstrap installers use their
  immutable same-tag manifest.
- Keep release sequence numbers strictly increasing across a channel.
- Preserve the assurance pack described in [THREATLOCKER.md](THREATLOCKER.md).

## Tests

`tests/Security.Tests.ps1` parses every runtime script, rejects dangerous legacy
patterns, checks scheduled-task preflight controls, consumes the build-compatible
manifest fixture through the real installer parser, and tests mutable URL and
sequence rejection. On a Windows host with Pester:

```powershell
Invoke-Pester -Script .\tests\Security.Tests.ps1
```

The local shell must permit the test file under the organisation's normal policy;
do not use `ExecutionPolicy Bypass` as a test workaround.
