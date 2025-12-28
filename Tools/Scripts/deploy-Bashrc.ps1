# Deploy-Bashrc.ps1
# - Liest Host-Aliase aus OpenSSH config inkl. Include-Dateien
# - Kopiert .bashrc per scp auf jeden Host
# - Backup + chmod + bash -n auf dem Ziel
# - Include-/Exclude-Listen für gezielte Steuerung

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$SshConfigPath = "C:\Users\AlessandroBello\.ssh\config"
$TemplatePath  = "C:\Users\AlessandroBello\Workspaces\ServiceConfigs\ServiceConfigs\Linux\Bash\.bashrc"

# Host-Steuerung
$IncludeHosts = @()               # leer = alle Hosts
$ExcludeHosts = @("ccu", "homeassistant")          # hier weitere Hosts ergänzen

# Optional: Diagnostik
$ShowConfigFiles = $true

function Assert-Command {
  param([Parameter(Mandatory)][string]$Name)
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Command '$Name' nicht gefunden. OpenSSH Client installieren oder ssh/scp in den PATH aufnehmen."
  }
}

function Get-Tokens {
  param([Parameter(Mandatory)][string]$Text)

  $tokens = New-Object System.Collections.Generic.List[string]
  $matches = [regex]::Matches($Text, '("([^"]+)"|''([^'']+)''|(\S+))')
  foreach ($m in $matches) {
    if ($m.Groups[2].Success) { $tokens.Add($m.Groups[2].Value); continue }
    if ($m.Groups[3].Success) { $tokens.Add($m.Groups[3].Value); continue }
    if ($m.Groups[4].Success) { $tokens.Add($m.Groups[4].Value); continue }
  }
  $tokens
}

function Expand-IncludePath {
  param(
    [Parameter(Mandatory)][string]$BaseDir,
    [Parameter(Mandatory)][string]$Spec
  )

  $p = [Environment]::ExpandEnvironmentVariables($Spec)

  if ($p.StartsWith("~")) {
    $p = Join-Path $env:USERPROFILE $p.Substring(1).TrimStart('\','/')
  }

  $p = $p -replace '/', '\'

  if (-not [System.IO.Path]::IsPathRooted($p)) {
    $p = Join-Path $BaseDir $p
  }

  $p
}

function Get-OpenSshConfigFiles {
  param([Parameter(Mandatory)][string]$RootConfigPath)

  if (-not (Test-Path -LiteralPath $RootConfigPath -PathType Leaf)) {
    throw "SSH config nicht gefunden: $RootConfigPath"
  }

  $seen  = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
  $queue = New-Object System.Collections.Generic.Queue[string]
  $result = New-Object System.Collections.Generic.List[string]

  $rootFull = (Get-Item -LiteralPath $RootConfigPath -Force).FullName
  [void]$queue.Enqueue($rootFull)

  while ($queue.Count -gt 0) {
    $file = $queue.Dequeue()
    $full = (Get-Item -LiteralPath $file -Force).FullName

    if (-not $seen.Add($full)) { continue }

    $result.Add($full)

    $baseDir = Split-Path -Parent $full
    $lines = Get-Content -LiteralPath $full -ErrorAction Stop

    foreach ($line in $lines) {
      $t = $line.Trim()
      if ($t.Length -eq 0) { continue }
      if ($t.StartsWith("#")) { continue }

      $noComment = ($t -split '\s+#', 2)[0].Trim()

      if ($noComment -match '^(?i)Include\s+(.+)$') {
        $rest = $Matches[1].Trim()
        $specs = Get-Tokens -Text $rest

        foreach ($spec in $specs) {
          $expanded = Expand-IncludePath -BaseDir $baseDir -Spec $spec

          $candidates = @()
          if ($expanded -match '[\*\?]') {
            $candidates = Get-ChildItem -Path $expanded -File -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }
          } else {
            if (Test-Path -LiteralPath $expanded -PathType Leaf) {
              $candidates = @((Get-Item -LiteralPath $expanded -Force).FullName)
            }
          }

          foreach ($c in $candidates) {
            [void]$queue.Enqueue($c)
          }
        }
      }
    }
  }

  $result | Sort-Object -Unique
}

function Get-SshHostAliases {
  param([Parameter(Mandatory)][string[]]$ConfigFiles)

  $aliases = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)

  foreach ($file in $ConfigFiles) {
    $lines = Get-Content -LiteralPath $file -ErrorAction Stop

    foreach ($line in $lines) {
      $t = $line.Trim()
      if ($t.Length -eq 0) { continue }
      if ($t.StartsWith("#")) { continue }

      $noComment = ($t -split '\s+#', 2)[0].Trim()

      if ($noComment -match '^(?i)Host\s+(.+)$') {
        $rest = $Matches[1].Trim()
        if ($rest -eq "*") { continue }

        $parts = $rest -split '\s+'
        foreach ($p in $parts) {
          if ([string]::IsNullOrWhiteSpace($p)) { continue }
          if ($p.StartsWith("!")) { continue }
          if ($p -match '[\*\?]') { continue }
          [void]$aliases.Add($p)
        }
      }
    }
  }

  $aliases | Sort-Object
}

function Invoke-RemoteBash {
  param(
    [Parameter(Mandatory)][string]$HostAlias,
    [Parameter(Mandatory)][string]$Script
  )

  $tmp = Join-Path $env:TEMP ("bash-" + [guid]::NewGuid().ToString("N") + ".sh")

  try {
    # CRLF -> LF, UTF-8 ohne BOM
    $lf = $Script -replace "`r", ""
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tmp, $lf, $utf8NoBom)

    $p = Start-Process -FilePath "ssh" `
      -ArgumentList @("-T", $HostAlias, "bash", "-s") `
      -RedirectStandardInput $tmp `
      -NoNewWindow `
      -Wait `
      -PassThru

    if ($p.ExitCode -ne 0) {
      throw "SSH fehlgeschlagen auf '$HostAlias' (ExitCode $($p.ExitCode))"
    }
  }
  finally {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  }
}

function Copy-ToRemote {
  param(
    [Parameter(Mandatory)][string]$HostAlias,
    [Parameter(Mandatory)][string]$LocalPath,
    [Parameter(Mandatory)][string]$RemotePath
  )

  & scp -p $LocalPath "${HostAlias}:$RemotePath"
  if ($LASTEXITCODE -ne 0) {
    throw "SCP fehlgeschlagen nach '$HostAlias' (ExitCode $LASTEXITCODE)"
  }
}

Assert-Command -Name "ssh"
Assert-Command -Name "scp"

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
  throw "Template nicht gefunden: $TemplatePath"
}

$configFiles = Get-OpenSshConfigFiles -RootConfigPath $SshConfigPath

if ($ShowConfigFiles) {
  Write-Host "SSH-Config Dateien (inkl. Include):"
  foreach ($f in $configFiles) { Write-Host "- $f" }
  Write-Host ""
}

$hosts = Get-SshHostAliases -ConfigFiles $configFiles

# Include/Exclude anwenden
if ($IncludeHosts.Count -gt 0) {
  $hosts = $hosts | Where-Object { $_ -in $IncludeHosts }
}
if ($ExcludeHosts.Count -gt 0) {
  $hosts = $hosts | Where-Object { $_ -notin $ExcludeHosts }
}

if (-not $hosts -or $hosts.Count -eq 0) {
  throw "Keine Host-Aliase gefunden. Prüfe Host-Definitionen und Include-Pfade oder Include/Exclude-Filter."
}

Write-Host "Gefundene Hosts (nach Filter):"
foreach ($h in $hosts) { Write-Host "- $h" }
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

  Invoke-RemoteBash -HostAlias $h -Script $backupScript

  Copy-ToRemote -HostAlias $h -LocalPath $TemplatePath -RemotePath "~/.bashrc"

  $validateScript = @'
set -e
chmod 600 "$HOME/.bashrc"
bash -n "$HOME/.bashrc"
'@

  Invoke-RemoteBash -HostAlias $h -Script $validateScript

  Write-Host "OK: .bashrc deployed. Backup: ~/.bashrc.bak.$timestamp"
  Write-Host ""
}

Write-Host "Fertig."
