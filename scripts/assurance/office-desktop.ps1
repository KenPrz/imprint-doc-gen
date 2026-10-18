#requires -Version 5.1
<#
Run on a dedicated, interactive Windows VM with licensed desktop Word and Excel.
Trusted operator fixtures only. Disconnect networking and close Office first.
This collects opening/export evidence; it never approves a release or visual baseline.
Example: .\scripts\assurance\office-desktop.ps1 -Manifest C:\fixtures\manifest.json -OutputDirectory C:\evidence\run1
#>
[CmdletBinding()]
param(
    [string]$Manifest,
    [string]$OutputDirectory,
    [ValidateRange(5, 600)][int]$TimeoutSeconds = 90,
    [string]$WorkerInput
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT' -or $PSVersionTable.PSEdition -ne 'Desktop') { throw 'Windows PowerShell 5.1 and desktop Office are required.' }

function Write-Json($Value, [string]$Path) {
    $Value | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding UTF8
}
function Hash([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }

if ($WorkerInput) {
    $job = Get-Content -LiteralPath $WorkerInput -Raw | ConvertFrom-Json
    $result = [ordered]@{
        id = $job.id; inputPath = $job.path; inputSha256 = $job.sha256
        application = $null; applicationVersion = $null; applicationBuild = $null
        status = 'failed'; pdfPath = $job.pdf; pdfSha256 = $null; error = $null
        readonly = $false; macrosDisabled = $false; externalLinkUpdatesDisabled = $false
        repairRequested = $false; interactiveWarnings = 'unknown'
    }
    $app = $null; $document = $null; $originalWordUpdateLinks = $null
    try {
        if ((Hash $job.path) -ne $job.sha256) { throw 'Input hash changed before opening.' }
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OfficeWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
        $isWord = [IO.Path]::GetExtension($job.path) -eq '.docx'
        $progid = if ($isWord) { 'Word.Application' } else { 'Excel.Application' }
        $result.application = $progid
        $app = New-Object -ComObject $progid
        $app.AutomationSecurity = 3
        $app.Visible = $false
        # Word exposes HWND on Window, not Application. A blank document supplies
        # a window before the actual fixture can block in Open.
        if ($isWord) {
            $document = $app.Documents.Add()
            $officeWindowHandle = $document.Windows.Item(1).Hwnd
        } else { $officeWindowHandle = $app.Hwnd }
        # Record only this COM instance's process, never kill Office by image name.
        [uint32]$officeProcessId = 0
        [void][OfficeWindow]::GetWindowThreadProcessId([IntPtr]$officeWindowHandle, [ref]$officeProcessId)
        if ($officeProcessId -eq 0) { throw 'Unable to identify the Office process.' }
        $officeProcess = Get-Process -Id $officeProcessId
        Write-Json @{ id = $officeProcessId; started = $officeProcess.StartTime.ToUniversalTime().ToString('o') } $job.process
        if ($isWord) {
            $document.Close(0)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
            $document = $null
        }
        $result.applicationVersion = [string]$app.Version
        $result.applicationBuild = [string]$app.Build
        $app.Visible = $false
        $app.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
        $missing = [Type]::Missing
        # Explicit nonempty random password prevents an interactive password request.
        # Encrypted inputs are deliberately unsupported, never decrypted by this runner.
        $unavailablePassword = [Guid]::NewGuid().ToString('N')
        if ($isWord) {
            $app.DisplayAlerts = 0
            $originalWordUpdateLinks = $app.Options.UpdateLinksAtOpen
            $app.Options.UpdateLinksAtOpen = $false
            $document = $app.Documents.OpenNoRepairDialog(
                $job.path, $false, $true, $false, $unavailablePassword,
                $unavailablePassword, $false, $unavailablePassword, $unavailablePassword,
                $missing, $missing, $false, $false, $missing, $true, $missing)
            $result.readonly = [bool]$document.ReadOnly; $result.macrosDisabled = $true; $result.externalLinkUpdatesDisabled = $true
            $document.ExportAsFixedFormat($job.pdf, 17, $false)
        } else {
            $app.DisplayAlerts = $false
            $app.EnableEvents = $false
            $app.AskToUpdateLinks = $false
            # UpdateLinks=0, ReadOnly=true, Notify=false, CorruptLoad=xlNormalLoad.
            $document = $app.Workbooks.Open(
                $job.path, 0, $true, $missing, $unavailablePassword, $unavailablePassword,
                $true, $missing, $missing, $false, $false, $missing, $false, $false, 0)
            $result.readonly = [bool]$document.ReadOnly; $result.macrosDisabled = $true; $result.externalLinkUpdatesDisabled = $true
            $document.ExportAsFixedFormat(0, $job.pdf)
        }
        if (!$result.readonly) { throw 'Office did not open the fixture read-only.' }
        if (!(Test-Path -LiteralPath $job.pdf) -or (Get-Item -LiteralPath $job.pdf).Length -lt 5) { throw 'No PDF was exported.' }
        if ((Hash $job.path) -ne $job.sha256) { throw 'Input changed during Office processing.' }
        $result.pdfSha256 = Hash $job.pdf
        $result.status = 'opened-and-exported'
    } catch { $result.error = $_.Exception.Message }
    finally {
        # Parent bounds this cleanup too: modal Office dialogs can also block Close/Quit.
        try { if ($null -ne $document) { $document.Close($false) } } catch { $result.status = 'failed'; $result.error = 'Office document cleanup failed: ' + $_.Exception.Message }
        try {
            if ($null -ne $app) {
                if ($null -ne $originalWordUpdateLinks) { $app.Options.UpdateLinksAtOpen = $originalWordUpdateLinks }
                $app.Quit()
            }
        } catch { $result.status = 'failed'; $result.error = 'Office application cleanup failed: ' + $_.Exception.Message }
        if ($null -ne $document) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
        if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
        Write-Json $result $job.result
    }
    exit
}

if (!$Manifest -or !$OutputDirectory) { throw 'Manifest and OutputDirectory are required.' }
if (@(Get-Process -Name WINWORD, EXCEL -ErrorAction SilentlyContinue).Count -gt 0) { throw 'Close Word and Excel before running on the dedicated VM.' }
$manifestPath = (Resolve-Path -LiteralPath $Manifest).Path
if ((Get-Item -LiteralPath $manifestPath).Length -gt 4MB) { throw 'Manifest exceeds 4 MiB.' }
$manifestDirectory = [IO.Path]::GetDirectoryName($manifestPath)
$manifestPrefix = $manifestDirectory.TrimEnd('\') + '\'
$definition = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ($definition.schemaVersion -ne 1 -or $definition.fixtures -isnot [Array] -or $definition.fixtures.Count -lt 1 -or $definition.fixtures.Count -gt 1000) { throw 'Expected schemaVersion 1 and a fixtures array with 1 to 1000 entries.' }
$seen = @{}
$fixtures = foreach ($fixture in $definition.fixtures) {
    if ($fixture.id -isnot [string] -or $fixture.id -notmatch '^[a-zA-Z0-9][a-zA-Z0-9_-]{0,79}$' -or $seen.ContainsKey($fixture.id)) { throw 'Invalid or duplicate fixture id.' }
    $seen[$fixture.id] = $true
    if ($fixture.sha256 -isnot [string] -or $fixture.sha256 -cnotmatch '^[a-f0-9]{64}$') { throw 'Each fixture requires a lowercase SHA-256.' }
    if ($fixture.path -isnot [string] -or !$fixture.path -or $fixture.path.Length -gt 240) { throw 'Fixture path must be a nonempty string of at most 240 characters.' }
    if ($fixture.path -match '^[a-zA-Z]:[\\/]') {
        if ($fixture.path.Substring(2).Contains(':')) { throw 'Alternate data streams are not supported.' }
        $candidate = [IO.Path]::GetFullPath($fixture.path)
    } else {
        # Portable manifest paths are relative to the manifest, never the process cwd.
        # Reject traversal lexically, then check normalized containment as defense in depth.
        if ([IO.Path]::IsPathRooted($fixture.path) -or $fixture.path.Contains(':') -or (($fixture.path -split '[\\/]') -contains '..')) { throw 'Relative fixture paths must stay inside the manifest directory.' }
        $candidate = [IO.Path]::GetFullPath((Join-Path $manifestDirectory $fixture.path))
        if (!$candidate.StartsWith($manifestPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture escapes the manifest directory.' }
    }
    $resolved = (Resolve-Path -LiteralPath $candidate).Path
    $drive = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($resolved))
    if ($drive.DriveType -eq [IO.DriveType]::Network) { throw 'Network fixture drives are not supported.' }
    # Resolve-Path alone does not resolve Windows junctions/symlinks. Reject every
    # reparse point in the chain, including the manifest directory's ancestors.
    $item = Get-Item -LiteralPath $resolved -Force
    if ($item.PSIsContainer -or $item.Length -gt 256MB) { throw 'Fixture must be a file of at most 256 MiB.' }
    $ancestor = $item
    while ($null -ne $ancestor) {
        if (($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Fixture path contains a symlink or junction.' }
        if ($ancestor -is [IO.FileInfo]) { $ancestor = $ancestor.Directory } else { $ancestor = $ancestor.Parent }
    }
    if ([IO.Path]::GetExtension($resolved) -notin @('.docx', '.xlsx')) { throw 'Only DOCX and XLSX are supported.' }
    if ((Hash $resolved) -ne $fixture.sha256) { throw "Fixture hash mismatch: $($fixture.id)" }
    @{ id = $fixture.id; path = $resolved; sha256 = $fixture.sha256 }
}
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'OutputDirectory must not already exist.' }
[void][IO.Directory]::CreateDirectory($output)
$report = [ordered]@{
    schemaVersion = 1; kind = 'office-desktop-run'; startedAt = [DateTime]::UtcNow.ToString('o')
    completedAt = $null; platform = 'win32'; osVersion = [Environment]::OSVersion.VersionString
    manifestSha256 = Hash $manifestPath; scriptSha256 = Hash $PSCommandPath
    releaseApproved = $false; humanReviewRequired = $true; results = @()
}
foreach ($fixture in $fixtures) {
    $prefix = Join-Path $output $fixture.id
    $job = @{
        id = $fixture.id; path = $fixture.path; sha256 = $fixture.sha256
        pdf = "$prefix.pdf"; process = "$prefix.process.json"; result = "$prefix.result.json"
    }
    $jobPath = "$prefix.input.json"
    Write-Json $job $jobPath
    # Encode the invocation as UTF-16: no interpolated paths are executable shell text.
    $escapedScript = $PSCommandPath.Replace("'", "''")
    $escapedJob = $jobPath.Replace("'", "''")
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes("& '$escapedScript' -WorkerInput '$escapedJob'"))
    $child = Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList @('-NoProfile', '-NonInteractive', '-STA', '-EncodedCommand', $encoded) -PassThru
    $finished = $child.WaitForExit($TimeoutSeconds * 1000)
    if (!$finished) { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue }
    # Cleanup only the recorded PID with matching start time to avoid PID reuse.
    if (Test-Path -LiteralPath $job.process) {
        $owned = Get-Content -LiteralPath $job.process -Raw | ConvertFrom-Json
        $process = Get-Process -Id $owned.id -ErrorAction SilentlyContinue
        if ($null -ne $process -and $process.StartTime.ToUniversalTime().ToString('o') -eq $owned.started) {
            Stop-Process -Id $process.Id -Force
        }
    }
    if ($finished -and $child.ExitCode -eq 0 -and (Test-Path -LiteralPath $job.result)) {
        $result = Get-Content -LiteralPath $job.result -Raw | ConvertFrom-Json
    } else {
        $result = @{
            id = $fixture.id; inputPath = $fixture.path; inputSha256 = $fixture.sha256
            application = $null; applicationVersion = $null; applicationBuild = $null
            status = $(if ($finished) { 'failed' } else { 'timeout' })
            pdfPath = $null; pdfSha256 = $null
            readonly = $false; macrosDisabled = $false; externalLinkUpdatesDisabled = $false
            repairRequested = $false; interactiveWarnings = 'unknown'
            error = $(if ($finished) { 'Worker failed without a completed result.' } else { 'Office exceeded the per-fixture deadline; operator investigation required.' })
        }
    }
    $report.results += $result
    Write-Json $report (Join-Path $output 'report.json')
}
$report.completedAt = [DateTime]::UtcNow.ToString('o')
Write-Json $report (Join-Path $output 'report.json')
if (@($report.results | Where-Object { $_.status -ne 'opened-and-exported' }).Count -gt 0) { exit 1 }
