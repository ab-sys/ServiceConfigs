[CmdletBinding()]
param(
  [Parameter()][ValidateNotNullOrEmpty()][string]$SshConfigPath = "C:\Users\AlessandroBello\OneDrive - AB-Systems\Tools\ssh\config",
  [Parameter()][ValidateRange(1,300)][int]$ConnectTimeoutSeconds = 10,
  [Parameter()][ValidateNotNullOrEmpty()][string]$TranscriptFile = ".\lxc_update_transcript_$(Get-Date -Format yyyyMMdd_HHmmss).log",
  [Parameter()][switch]$PauseBetweenHosts,
  [Parameter()][string[]]$ExcludeHosts = @("ccu", "docker-prod", "homeassistant")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Clear-Host

if (-not $PSBoundParameters.ContainsKey('PauseBetweenHosts')) {
  $PauseBetweenHosts = $true
}

if (-not (Get-Command ssh -ErrorAction SilentlyContinue)) {
  throw "ssh.exe nicht gefunden. Installiere den 'OpenSSH Client' unter Windows Features."
}
if (-not (Test-Path -LiteralPath $SshConfigPath)) {
  throw "SSH config nicht gefunden: $SshConfigPath"
}

$transcriptPath = [System.IO.Path]::GetFullPath($TranscriptFile)
$transcriptDir = Split-Path -Parent $transcriptPath
if ($transcriptDir -and -not (Test-Path -LiteralPath $transcriptDir)) {
  New-Item -ItemType Directory -Path $transcriptDir -Force | Out-Null
}

function Resolve-IncludePaths([string]$baseFile, [string]$includeValue) {
  $baseDir = Split-Path -Parent $baseFile
  $raw = $includeValue.Trim()
  if (-not $raw) { return @() }

  $patterns = $raw -split "\s+"
  $paths = @()

  foreach ($p in $patterns) {
    $expanded = $p

    if ($expanded.StartsWith("~")) {
      $userHome = $env:USERPROFILE
      $expanded = Join-Path $userHome ($expanded.Substring(1).TrimStart("/","\")) 
    }

    if (-not [System.IO.Path]::IsPathRooted($expanded)) {
      $expanded = Join-Path $baseDir $expanded
    }

    $includeMatches = Get-ChildItem -Path $expanded -File -ErrorAction SilentlyContinue
    foreach ($m in $includeMatches) { $paths += $m.FullName }
  }

  return $paths
}

function Get-SshConfigLines([string]$path, [System.Collections.Generic.HashSet[string]]$visited) {
  $full = (Resolve-Path -LiteralPath $path).Path
  if ($visited.Contains($full)) { return @() }
  $visited.Add($full) | Out-Null

  $lines = Get-Content -LiteralPath $full -ErrorAction Stop
  $all = New-Object System.Collections.Generic.List[string]

  foreach ($l in $lines) {
    $trim = $l.Trim()

    $includeMatch = [regex]::Match($trim, '^(?i)Include\s+(.+)$')
    if ($includeMatch.Success) {
      $incValue = $includeMatch.Groups[1].Value
      $incPaths = Resolve-IncludePaths -baseFile $full -includeValue $incValue
      foreach ($ip in $incPaths) {
        (Get-SshConfigLines -path $ip -visited $visited) | ForEach-Object { $all.Add($_) }
      }
      continue
    }

    $all.Add($l)
  }

  return $all.ToArray()
}

function Get-HostAliasesFromConfig([string]$path) {
  $visited = New-Object "System.Collections.Generic.HashSet[string]"
  $lines = Get-SshConfigLines -path $path -visited $visited

  $set = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)

  foreach ($line in $lines) {
    $noComment = ($line -split "#", 2)[0].Trim()
    if (-not $noComment) { continue }

    $hostMatch = [regex]::Match($noComment, '^(?i)Host\s+(.+)$')
    if ($hostMatch.Success) {
      $names = $hostMatch.Groups[1].Value.Trim() -split "\s+"
      foreach ($n in $names) {
        if (-not $n) { continue }
        if ($n.StartsWith("!")) { continue }   # negierte patterns ignorieren
        if ($n -match '[\*\?]') { continue }   # wildcard patterns ignorieren
        $null = $set.Add($n)
      }
    }
  }

  return $set | Sort-Object
}

$hosts = Get-HostAliasesFromConfig -path $SshConfigPath
if (-not $hosts -or $hosts.Count -eq 0) {
  throw "Keine expliziten Host-Aliases gefunden (nur Wildcards oder leere Config)."
}

if ($ExcludeHosts -and $ExcludeHosts.Count -gt 0) {
  $excludeSet = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($ex in $ExcludeHosts) {
    if (-not [string]::IsNullOrWhiteSpace($ex)) { $null = $excludeSet.Add($ex.Trim()) }
  }
  $hosts = $hosts | Where-Object { -not $excludeSet.Contains($_) }
}

if (-not $hosts -or $hosts.Count -eq 0) {
  throw "Alle Hosts wurden ausgeschlossen. Bitte Exclude-Liste pruefen."
}

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

    # -tt erzwingt TTY, damit Interaktivitaet sicher funktioniert
    # Kein BatchMode, damit ggf. auch Passwort/Keyboard-Interactive moeglich ist
    $sshArgs = @(
      "-tt",
      "-F", $SshConfigPath,
      "-o", "StrictHostKeyChecking=accept-new",
      "-o", "ConnectTimeout=$ConnectTimeoutSeconds",
      $h,
      $remoteCmd
    )

    try {
      & ssh @sshArgs
      $exit = $LASTEXITCODE
    }
    catch {
      $exit = $LASTEXITCODE
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
