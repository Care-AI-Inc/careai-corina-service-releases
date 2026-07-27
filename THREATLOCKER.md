# ThreatLocker and clinic allowlisting guide

This is a deployment input for each clinic/MSP, not a request to disable threat
controls. Build the narrowest policy possible and validate it in ThreatLocker's
audit/test mode before enforcement. ThreatLocker configuration and CDN endpoints
can change; the clinic's security administrator remains the policy owner.

## Two distinct approval layers

1. **Clinic wrapper — exact hash.** Each enrolment wrapper is unique and therefore
   cannot share a publisher hash. Approve only that supplied file's SHA-256, its
   intended administrator, machine/group and short deployment window. The wrapper
   may start the immutable CARE AI signed `install.ps1`; it must not confer a
   general rule for PowerShell or other clinic scripts.
2. **Static CARE AI release files — certificate plus constrained identity.** Approve
   the exact CARE AI code-signing certificate/publisher only when the file is under
   the locked Corina Program Files path (or is the one immutable installer staged
   by the approved wrapper). Where supported, combine certificate, filename,
   product metadata, owner/ACL and parent-process constraints. The service package
   and release manifest hashes remain release-specific evidence even when a
   certificate rule is used operationally.

ThreatLocker Application Control approval does not automatically bypass
PowerShell Ringfencing. Review existing Ringfencing policies for PowerShell's
network, registry, service-control and child-process operations. Permit only the
behaviour below for the approved wrapper/static scripts. A higher-priority deny
can still override a lower computer/group allow rule.

ThreatLocker approval is also separate from Windows PowerShell execution policy.
On an `AllSigned` estate, validate the assurance-pack thumbprint and deploy the
CARE AI signing certificate through the clinic's normal Trusted Publishers/GPO
process so the non-interactive SYSTEM updater can run without a trust prompt. Do
not work around either control with `ExecutionPolicy Bypass`.

## Production inventory

| Item | Expected value |
|---|---|
| Service | `CorinaService` or `CorinaService-<instance>`; LocalSystem; automatic |
| Scheduled task | `CorinaProdDailyUpdater` or `CorinaProdDailyUpdater-<instance>`; SYSTEM/highest |
| Service binary | `C:\Program Files\CorinaService[\<instance>]\careai-corina-service.exe` |
| Static scripts | `C:\Program Files\CorinaService[\<instance>]\Update\{daily-updater.ps1,ensure-updater-task.ps1,uninstall.ps1}` |
| Update state | `C:\ProgramData\CareAI\CorinaService\{default|<instance>}\{Staging,Backup,Logs}` |
| Registry | `HKLM\SOFTWARE\CareAI\CorinaService[\<instance>]` and the service's SCM key |
| Task host | `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe` with `-NoLogo -NoProfile -NonInteractive`; no bypass |
| Native children | `sc.exe` and `robocopy.exe` during install/update; the verified service executable via SCM |

Staging uses the separately constrained `CorinaService-Staging`,
`CorinaStagingDailyUpdater`, `C:\Program Files\CorinaService-Staging`,
`C:\ProgramData\CareAI\CorinaService-Staging` and corresponding registry root.
Do not let a staging approval automatically match production.

The scheduled-task command itself performs a signature/thumbprint/publisher/EKU
preflight before invoking the updater. Preserve this command when reviewing task
changes; do not replace it with a direct unverified script action.

## Network destinations

The scripts make initial HTTPS requests only to immutable or signed-discovery URLs
on `github.com` in the applicable CARE AI release repository. GitHub release asset
downloads redirect to GitHub-operated object/CDN hosts—commonly
`release-assets.githubusercontent.com`, with the exact redirect hostname subject
to GitHub change. Inspect the release URL redirect and clinic proxy logs for the
current endpoint; constrain it to HTTPS, the updater/installer process, CARE AI
release traffic and the deployment window. A CDN hostname allow alone is not an
artifact trust decision; manifest signatures and hashes remain mandatory.

The installer and updater do not change or migrate existing clinic application
credentials. This release is deliberately limited to artifact signing,
version/hash verification, installation, update and rollback behavior.

## Rules not to create

Do not broadly allow any of the following:

- `powershell.exe` or all signed/unsigned PowerShell scripts
- `%TEMP%\*.ps1`, `C:\Scripts\*` or user-writable script locations
- `ExecutionPolicy Bypass`, encoded commands or `Invoke-Expression`
- all GitHub/raw/CDN content regardless of process and artifact verification
- Defender exclusions or unrestricted child-process/service creation
- the CARE AI certificate at arbitrary writable paths

The deprecated `run-daily-updater-prod.ps1` repository reference is not deployed
or scheduled. New policy should not approve the old `C:\Scripts` downloader shim;
the installer removes it during migration where it is no longer legitimately used.

## Certificate rotation in policy

CARE AI uses a two-release overlap. Before certificate renewal, the assurance pack
will list both current and next thumbprints. Add the next certificate rule with the
same path/process constraints while retaining the old rule. After a release signed
by the new certificate is deployed and verified fleet-wide, remove the old
certificate rule. Never approve by publisher display name alone, and do not remove
the old certificate before the overlap release reaches offline clinics.

For suspected key compromise, disable the affected rule immediately and require a
new exact-hash migration package through the clinic's change process. Automatic
manifest rotation intentionally cannot trust a release signed only by an unknown
key.

## Per-release assurance pack

Provide the clinic/MSP with:

- channel, immutable tag, release version, monotonic sequence and publication time
- Git commit IDs for service source and release-script source
- SHA-256 and byte size for the signed installer, updater, helper, uninstaller,
  service ZIP and signed channel manifest
- Authenticode status, exact signer subject, leaf thumbprint, issuer, validity,
  code-signing EKU and timestamp details for every executable artifact
- manifest signature verification result and a copy of the literal manifest
- extracted service executable hash/signature and deterministic package hash
- service, task, filesystem, registry, process-child and network inventory
- silent invocation syntax, instance value (if any), expected installer parent and
  deployment window
- release notes, health-check evidence, rollback instructions and known domains
- certificate-overlap notice showing current/next thumbprints and planned removal
- confirmation that release assets/tag are immutable and no post-sign edit occurred

Retain the approved wrapper hash and the assurance pack with the clinic change
ticket. A later release is a new change: script/package hashes must not be silently
reused.
