param(
  [string]$SshConfigPath = "C:\Users\AlessandroBello\OneDrive - AB-Systems\Tools\ssh\config",
  [int]$ConnectTimeoutSeconds = 10,
  [string]$TranscriptFile = ".\lxc_update_transcript_$(Get-Date -Format yyyyMMdd_HHmmss).log"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Get-Command ssh -ErrorAction SilentlyContinue)) {
  throw "ssh.exe nicht gefunden. Installiere den 'OpenSSH Client' unter Windows Features."
}
if (-not (Test-Path -LiteralPath $SshConfigPath)) {
  throw "SSH config nicht gefunden: $SshConfigPath"
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
      $home = $env:USERPROFILE
      $expanded = Join-Path $home ($expanded.Substring(1).TrimStart("/","\"))
    }

    if (-not [System.IO.Path]::IsPathRooted($expanded)) {
      $expanded = Join-Path $baseDir $expanded
    }

    $matches = Get-ChildItem -Path $expanded -File -ErrorAction SilentlyContinue
    foreach ($m in $matches) { $paths += $m.FullName }
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

    if ($trim -match '^(?i)Include\s+(.+)$') {
      $incPaths = Resolve-IncludePaths -baseFile $full -includeValue $Matches[1]
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

    if ($noComment -match '^(?i)Host\s+(.+)$') {
      $names = $Matches[1].Trim() -split "\s+"
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

# "update" so ausführen, dass typische Bash-Aliases/Funktionen verfügbar sind:
# bash -lic: login + interactive + command
$remoteCmd = @"
command -v bash >/dev/null 2>&1 && bash -lic 'update' || sh -c 'update'
"@.Trim()

Start-Transcript -Path $TranscriptFile -Append | Out-Null

try {
  Write-Host "START: $($hosts.Count) Hosts aus SSH config: $SshConfigPath"
  Write-Host "Transcript: $TranscriptFile"
  Write-Host ""

  foreach ($h in $hosts) {
    Write-Host "============================================================"
    Write-Host "HOST: $h"
    Write-Host "Action: update (interaktiv)"
    Write-Host "Hinweis: Wenn 'update' beendet ist, geht es automatisch zum nächsten Host."
    Write-Host "============================================================"
    Write-Host ""

    # -tt erzwingt TTY, damit Interaktivität sicher funktioniert
    # Kein BatchMode, damit ggf. auch Passwort/Keyboard-Interactive möglich ist
    $sshArgs = @(
      "-tt",
      "-F", $SshConfigPath,
      "-o", "StrictHostKeyChecking=accept-new",
      "-o", "ConnectTimeout=$ConnectTimeoutSeconds",
      $h,
      $remoteCmd
    )

    & ssh @sshArgs
    $exit = $LASTEXITCODE

    Write-Host ""
    Write-Host "HOST: $h - ExitCode: $exit"
    Write-Host ""
  }

  Write-Host "END: Fertig."
}
finally {
  Stop-Transcript | Out-Null
}
