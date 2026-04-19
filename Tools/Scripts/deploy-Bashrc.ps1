# Deploy-Bashrc.ps1
# - Liest Host-Aliase aus OpenSSH config inkl. Include-Dateien
# - Kopiert .bashrc per scp auf jeden Host
# - Backup + chmod + bash -n auf dem Ziel
# - Include-/Exclude-Listen fuer gezielte Steuerung

[CmdletBinding()]
param(
  [Parameter()][ValidateNotNullOrEmpty()][string]$SshConfigPath = "C:\Users\AlessandroBello\.ssh\config",
  [Parameter()][string[]]$IncludeHosts = @(),
  [Parameter()][string[]]$ExcludeHosts = @("ccu", "homeassistant"),
  [Parameter()][switch]$ShowConfigFiles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$helperPath = Join-Path $PSScriptRoot 'Helpers\SshHelpers.ps1'
. $helperPath

$TemplatePath = "C:\Users\AlessandroBello\Workspaces\ServiceConfigs\ServiceConfigs\Linux\Bash\.bashrc"

Assert-SshCommand -RequireScp

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
  throw "Template nicht gefunden: $TemplatePath"
}

$resolvedHosts = Resolve-SshHosts -SshConfigPath $SshConfigPath -IncludeHosts $IncludeHosts -ExcludeHosts $ExcludeHosts
$configFiles = @($resolvedHosts.ConfigFiles)
$hosts = @($resolvedHosts.Hosts)
$SshConfigPath = $resolvedHosts.SshConfigPath

if ($ShowConfigFiles) {
  Write-Host "SSH-Config Dateien (inkl. Include):"
  foreach ($configFile in $configFiles) {
    Write-Host "- $configFile"
  }
  Write-Host ""
}

if (-not $hosts -or $hosts.Count -eq 0) {
  throw "Keine Host-Aliase gefunden. Pruefe Host-Definitionen und Include/Exclude-Filter."
}

Write-Host "Gefundene Hosts (nach Filter):"
foreach ($h in $hosts) {
  Write-Host "- $h"
}
Write-Host ""

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

foreach ($h in $hosts) {
  Write-Host "==> $h"

  $backupScript = @'
set -e
if [ -f "$HOME/.bashrc" ]; then
  cp -p "$HOME/.bashrc" "$HOME/.bashrc.bak.__TS__"
fi
'@ -replace '__TS__', $timestamp

  Invoke-SshBashScript -HostAlias $h -SshConfigPath $SshConfigPath -Script $backupScript

  Copy-SshFile -HostAlias $h -SshConfigPath $SshConfigPath -LocalPath $TemplatePath -RemotePath "~/.bashrc"

  $validateScript = @'
set -e
chmod 600 "$HOME/.bashrc"
bash -n "$HOME/.bashrc"
'@

  Invoke-SshBashScript -HostAlias $h -SshConfigPath $SshConfigPath -Script $validateScript

  Write-Host "OK: .bashrc deployed. Backup: ~/.bashrc.bak.$timestamp"
  Write-Host ""
}

Write-Host "Fertig."
