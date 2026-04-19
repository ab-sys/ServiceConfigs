function Assert-Command {
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Name
  )

  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Command '$Name' nicht gefunden. OpenSSH Client installieren oder '$Name' in den PATH aufnehmen."
  }
}

function New-CaseInsensitiveStringSet {
  return ,([System.Collections.Generic.HashSet[string]]::new([System.Collections.Generic.IEqualityComparer[string]][System.StringComparer]::OrdinalIgnoreCase))
}

function Assert-SshCommand {
  param(
    [Parameter()][switch]$RequireScp
  )

  Assert-Command -Name "ssh"
  if ($RequireScp) {
    Assert-Command -Name "scp"
  }
}

function Get-SshTokens {
  param(
    [Parameter(Mandatory)][string]$Text
  )

  $tokens = New-Object System.Collections.Generic.List[string]
  $matches = [regex]::Matches($Text, '("([^"]+)"|''([^'']+)''|(\S+))')

  foreach ($match in $matches) {
    if ($match.Groups[2].Success) {
      $tokens.Add($match.Groups[2].Value)
      continue
    }
    if ($match.Groups[3].Success) {
      $tokens.Add($match.Groups[3].Value)
      continue
    }
    if ($match.Groups[4].Success) {
      $tokens.Add($match.Groups[4].Value)
      continue
    }
  }

  return $tokens
}

function Expand-SshIncludePath {
  param(
    [Parameter(Mandatory)][string]$BaseDir,
    [Parameter(Mandatory)][string]$Spec
  )

  $pathSpec = [Environment]::ExpandEnvironmentVariables($Spec)

  if ($pathSpec.StartsWith("~")) {
    $pathSpec = Join-Path $env:USERPROFILE $pathSpec.Substring(1).TrimStart('\', '/')
  }

  $pathSpec = $pathSpec -replace '/', '\'

  if (-not [System.IO.Path]::IsPathRooted($pathSpec)) {
    $pathSpec = Join-Path $BaseDir $pathSpec
  }

  return $pathSpec
}

function Get-OpenSshConfigFiles {
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RootConfigPath
  )

  if (-not (Test-Path -LiteralPath $RootConfigPath -PathType Leaf)) {
    throw "SSH config nicht gefunden: $RootConfigPath"
  }

  $seen = New-CaseInsensitiveStringSet
  $queue = New-Object 'System.Collections.Generic.Queue[string]'
  $result = New-Object System.Collections.Generic.List[string]

  $rootFull = (Get-Item -LiteralPath $RootConfigPath -Force).FullName
  [void]$queue.Enqueue($rootFull)

  while ($queue.Count -gt 0) {
    $file = $queue.Dequeue()
    $full = (Get-Item -LiteralPath $file -Force).FullName

    if (-not $seen.Add($full)) {
      continue
    }

    $result.Add($full)

    $baseDir = Split-Path -Parent $full
    $lines = Get-Content -LiteralPath $full -ErrorAction Stop

    foreach ($line in $lines) {
      $trimmed = $line.Trim()
      if ($trimmed.Length -eq 0) { continue }
      if ($trimmed.StartsWith("#")) { continue }

      $noComment = ($trimmed -split '\s+#', 2)[0].Trim()
      if ($noComment -notmatch '^(?i)Include\s+(.+)$') { continue }

      $specs = Get-SshTokens -Text $Matches[1].Trim()
      foreach ($spec in $specs) {
        $expanded = Expand-SshIncludePath -BaseDir $baseDir -Spec $spec

        if ($expanded -match '[\*\?]') {
          foreach ($candidate in (Get-ChildItem -Path $expanded -File -ErrorAction SilentlyContinue | Sort-Object FullName)) {
            [void]$queue.Enqueue($candidate.FullName)
          }
          continue
        }

        if (Test-Path -LiteralPath $expanded -PathType Leaf) {
          [void]$queue.Enqueue((Get-Item -LiteralPath $expanded -Force).FullName)
        }
      }
    }
  }

  return $result.ToArray()
}

function Get-SshHostAliases {
  param(
    [Parameter(Mandatory)][string[]]$ConfigFiles
  )

  $aliases = New-CaseInsensitiveStringSet

  foreach ($file in $ConfigFiles) {
    $lines = Get-Content -LiteralPath $file -ErrorAction Stop

    foreach ($line in $lines) {
      $trimmed = $line.Trim()
      if ($trimmed.Length -eq 0) { continue }
      if ($trimmed.StartsWith("#")) { continue }

      $noComment = ($trimmed -split '\s+#', 2)[0].Trim()
      if ($noComment -notmatch '^(?i)Host\s+(.+)$') { continue }

      foreach ($hostAlias in ($Matches[1].Trim() -split '\s+')) {
        if ([string]::IsNullOrWhiteSpace($hostAlias)) { continue }
        if ($hostAlias.StartsWith("!")) { continue }
        if ($hostAlias -match '[\*\?]') { continue }
        [void]$aliases.Add($hostAlias)
      }
    }
  }

  return $aliases | Sort-Object
}

function New-SshHostSet {
  param(
    [Parameter()][AllowNull()][string[]]$Hosts
  )

  $set = New-CaseInsensitiveStringSet

  foreach ($hostName in @($Hosts)) {
    if ([string]::IsNullOrWhiteSpace($hostName)) { continue }
    [void]$set.Add($hostName.Trim())
  }

  return ,$set
}

function Resolve-SshHosts {
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SshConfigPath,
    [Parameter()][string[]]$IncludeHosts = @(),
    [Parameter()][string[]]$ExcludeHosts = @()
  )

  $rootFull = (Get-Item -LiteralPath $SshConfigPath -Force).FullName
  $configFiles = Get-OpenSshConfigFiles -RootConfigPath $rootFull
  $hosts = @(Get-SshHostAliases -ConfigFiles $configFiles)

  $includeSet = New-SshHostSet -Hosts $IncludeHosts
  if ($includeSet.Count -gt 0) {
    $hosts = @($hosts | Where-Object { $includeSet.Contains($_) })
  }

  $excludeSet = New-SshHostSet -Hosts $ExcludeHosts
  if ($excludeSet.Count -gt 0) {
    $hosts = @($hosts | Where-Object { -not $excludeSet.Contains($_) })
  }

  return [pscustomobject]@{
    SshConfigPath = $rootFull
    ConfigFiles   = @($configFiles)
    Hosts         = @($hosts)
  }
}

function New-SshArgumentList {
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SshConfigPath,
    [Parameter()][ValidateRange(0, 300)][int]$ConnectTimeoutSeconds = 0,
    [Parameter()][switch]$ForceTty,
    [Parameter()][AllowNull()][string]$StrictHostKeyChecking
  )

  $arguments = New-Object System.Collections.Generic.List[string]

  if ($ForceTty) {
    $arguments.Add("-tt")
  }

  $arguments.Add("-F")
  $arguments.Add((Get-Item -LiteralPath $SshConfigPath -Force).FullName)

  if (-not [string]::IsNullOrWhiteSpace($StrictHostKeyChecking)) {
    $arguments.Add("-o")
    $arguments.Add("StrictHostKeyChecking=$StrictHostKeyChecking")
  }

  if ($ConnectTimeoutSeconds -gt 0) {
    $arguments.Add("-o")
    $arguments.Add("ConnectTimeout=$ConnectTimeoutSeconds")
  }

  return $arguments.ToArray()
}

function Invoke-SshCommand {
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$HostAlias,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SshConfigPath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Command,
    [Parameter()][ValidateRange(0, 300)][int]$ConnectTimeoutSeconds = 0,
    [Parameter()][switch]$ForceTty,
    [Parameter()][ValidateSet('accept-new', 'ask', 'no', 'yes')][string]$StrictHostKeyChecking
  )

  Assert-SshCommand

  $sshArgs = New-SshArgumentList -SshConfigPath $SshConfigPath -ConnectTimeoutSeconds $ConnectTimeoutSeconds -ForceTty:$ForceTty -StrictHostKeyChecking $StrictHostKeyChecking
  $sshArgs += $HostAlias
  $sshArgs += $Command

  & ssh @sshArgs
}

function Invoke-SshBashScript {
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$HostAlias,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SshConfigPath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Script,
    [Parameter()][ValidateRange(0, 300)][int]$ConnectTimeoutSeconds = 0,
    [Parameter()][ValidateSet('accept-new', 'ask', 'no', 'yes')][string]$StrictHostKeyChecking
  )

  Assert-SshCommand

  $sshArgs = New-SshArgumentList -SshConfigPath $SshConfigPath -ConnectTimeoutSeconds $ConnectTimeoutSeconds -StrictHostKeyChecking $StrictHostKeyChecking
  $sshArgs = @("-T") + $sshArgs + @($HostAlias, "bash", "-s")

  $lfScript = $Script -replace "`r", ""
  $lfScript | & ssh @sshArgs

  if ($LASTEXITCODE -ne 0) {
    throw "SSH fehlgeschlagen auf '$HostAlias' (ExitCode $LASTEXITCODE)"
  }
}

function Copy-SshFile {
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$HostAlias,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SshConfigPath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LocalPath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RemotePath
  )

  Assert-SshCommand -RequireScp

  if (-not (Test-Path -LiteralPath $LocalPath -PathType Leaf)) {
    throw "Lokale Datei nicht gefunden: $LocalPath"
  }

  $scpArgs = @(
    "-F", (Get-Item -LiteralPath $SshConfigPath -Force).FullName,
    "-p",
    (Get-Item -LiteralPath $LocalPath -Force).FullName,
    "${HostAlias}:$RemotePath"
  )

  & scp @scpArgs
  if ($LASTEXITCODE -ne 0) {
    throw "SCP fehlgeschlagen nach '$HostAlias' (ExitCode $LASTEXITCODE)"
  }
}
