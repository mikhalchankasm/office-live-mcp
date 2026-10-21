<#
.SYNOPSIS
    Install or update Office Live MCP from the latest GitHub release, silently.

.DESCRIPTION
    Downloads the latest *-setup.exe, checks its SHA-256 against the digest GitHub publishes for the asset
    (when available) and runs it with /VERYSILENT. An update keeps the install folder, connected agents and their
    access mode. Close the agents that use the server (Claude, Cursor, Codex...) first: a busy server stops the
    installer with exit code 7 before any file is copied.

    Run from a clone:      powershell -ExecutionPolicy Bypass -File scripts\update.ps1
    Run without a clone:   & ([scriptblock]::Create((irm https://raw.githubusercontent.com/mikhalchankasm/office-live-mcp/main/scripts/update.ps1)))

.PARAMETER Clients
    Passed as /CLIENTS: detected | none | a list such as codex,claude-code. Omitted: an update keeps every
    registration, a fresh install connects nobody.

.PARAMETER Access
    full | readonly, for agents added by this run only.

.PARAMETER Log
    Installer log file (/LOG). Default: a log in %TEMP%.

.PARAMETER DryRun
    Resolve the release and print what would run, without downloading.
#>
param(
    [string]$Clients,
    [ValidateSet('', 'full', 'readonly')][string]$Access = '',
    [string]$Log,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
# Windows PowerShell 5.1 may default to TLS 1.0, which GitHub refuses.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'  # the 5.1 progress bar slows Invoke-WebRequest down many times

$release = Invoke-RestMethod 'https://api.github.com/repos/mikhalchankasm/office-live-mcp/releases/latest' `
    -Headers @{ 'User-Agent' = 'office-live-mcp-update' }
$asset = @($release.assets | Where-Object { $_.name -like '*-setup.exe' })[0]
if (-not $asset) { throw "Release $($release.tag_name) has no *-setup.exe asset." }

# Inno Setup key of the permanent AppId (packaging/office-live-mcp.iss); per-user install.
$installed = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{47C3867B-C095-4C76-AACD-DA6A268CF854}_is1' `
    -ErrorAction SilentlyContinue
$current = if ($installed) { $installed.DisplayVersion } else { 'none' }
Write-Host "Installed by setup.exe: $current; latest: $($release.tag_name)"

$arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART')
if ($Clients) { $arguments += "/CLIENTS=$Clients" }
if ($Access) { $arguments += "/ACCESS=$Access" }
if ($Log) { $arguments += "/LOG=`"$Log`"" }

$file = Join-Path $env:TEMP $asset.name
if ($DryRun) {
    Write-Host "Would download $($asset.browser_download_url)"
    Write-Host "Would run: $file $($arguments -join ' ')"
    return
}

Invoke-WebRequest $asset.browser_download_url -OutFile $file -UseBasicParsing
if ($asset.digest -like 'sha256:*') {
    $expected = $asset.digest.Substring(7)
    $actual = (Get-FileHash $file -Algorithm SHA256).Hash
    if ($actual -ne $expected) {
        Remove-Item $file
        throw "SHA-256 mismatch for $($asset.name): expected $expected, got $actual."
    }
    Write-Host "SHA-256 verified."
}

$process = Start-Process $file -ArgumentList $arguments -Wait -PassThru
switch ($process.ExitCode) {
    0 { Write-Host "Office Live MCP $($release.tag_name) installed. Restart your agents." }
    7 { Write-Warning 'The server is in use: close the connected agents (Claude, Cursor, Codex...) and run again.' }
    default { Write-Warning "setup.exe exited with code $($process.ExitCode); rerun with -Log to see why." }
}
# `exit` inside a script block run via irm would close the caller's PowerShell window.
if ($MyInvocation.MyCommand.Path) { exit $process.ExitCode }
$global:LASTEXITCODE = $process.ExitCode
