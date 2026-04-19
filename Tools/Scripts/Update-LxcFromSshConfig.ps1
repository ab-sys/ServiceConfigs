[CmdletBinding()]
param(
  [Parameter()][ValidateNotNullOrEmpty()][string]$SshConfigPath = "C:\Users\AlessandroBello\Proton Drive\masbel2000\My Files\Tools\ssh\config",
  [Parameter()][ValidateRange(1, 300)][int]$ConnectTimeoutSeconds = 10,
  [Parameter()][ValidateNotNullOrEmpty()][string]$TranscriptFile = ".\lxc_update_transcript_$(Get-Date -Format yyyyMMdd_HHmmss).log",
  [Parameter()][switch]$PauseBetweenHosts,
  [Parameter()][string[]]$IncludeHosts = @(),
  [Parameter()][string[]]$ExcludeHosts = @("ccu", "docker-prod", "homeassistant", "npm", "pbs", "motu-online", "motu-lan-server", "proxmox"),
  [Parameter()][switch]$ShowConfigFiles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Clear-Host

$helperPath = Join-Path $PSScriptRoot 'Helpers\SshHelpers.ps1'
. $helperPath

if (-not $PSBoundParameters.ContainsKey('PauseBetweenHosts')) {
  $PauseBetweenHosts = $true
}

$transcriptPath = [System.IO.Path]::GetFullPath($TranscriptFile)
$transcriptDir = Split-Path -Parent $transcriptPath
if ($transcriptDir -and -not (Test-Path -LiteralPath $transcriptDir)) {
  New-Item -ItemType Directory -Path $transcriptDir -Force | Out-Null
}

$resolvedHosts = Resolve-SshHosts -SshConfigPath $SshConfigPath -IncludeHosts $IncludeHosts -ExcludeHosts $ExcludeHosts
$hosts = @($resolvedHosts.Hosts)
$SshConfigPath = $resolvedHosts.SshConfigPath

if ($ShowConfigFiles) {
  Write-Host "SSH-Config Dateien (inkl. Include):"
  foreach ($configFile in $resolvedHosts.ConfigFiles) {
    Write-Host "- $configFile"
  }
  Write-Host ""
}

if (-not $hosts -or $hosts.Count -eq 0) {
  throw "Keine Host-Aliase gefunden. Pruefe Host-Definitionen und Include/Exclude-Filter."
}

Write-Host "Gefundene Hosts (nach Filter):"
foreach ($hostName in $hosts) {
  Write-Host "- $hostName"
}
Write-Host ""

# "update" so ausfuehren, dass typische Bash-Aliases/Funktionen verfuegbar sind:
# bash -lic: login + interactive + command
$remoteCmd = @"
command -v bash >/dev/null 2>&1 && bash -lic 'update' || sh -c 'update'
"@.Trim()

Start-Transcript -Path $transcriptPath -Append | Out-Null

try {
  Write-Host "START: $($hosts.Count) Hosts aus SSH config: $SshConfigPath" -ForegroundColor Cyan
  Write-Host "Transcript: $transcriptPath" -ForegroundColor DarkGray
  Write-Host ""

  foreach ($h in $hosts) {
    Write-Host "============================================================" -ForegroundColor DarkGray
    Write-Host "HOST: $h" -ForegroundColor Yellow
    Write-Host "Action: update (interaktiv)" -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor DarkGray
    Write-Host ""

    try {
      Invoke-SshCommand -HostAlias $h -SshConfigPath $SshConfigPath -Command $remoteCmd -ConnectTimeoutSeconds $ConnectTimeoutSeconds -ForceTty -StrictHostKeyChecking accept-new
      $exit = $LASTEXITCODE
    }
    catch {
      $exit = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { 1 }
      Write-Warning "HOST: $h - Fehler: $($_.Exception.Message)"
    }

    Write-Host ""
    $exitColor = if ($exit -eq 0) { "Green" } else { "Red" }
    Write-Host "HOST: $h - ExitCode: $exit" -ForegroundColor $exitColor
    Write-Host ""

    if ($PauseBetweenHosts) {
      Read-Host "Weiter mit Enter (naechster Host)"
      Write-Host ""
    }
  }

  Write-Host "END: Fertig." -ForegroundColor Cyan
}
finally {
  Stop-Transcript | Out-Null
}
