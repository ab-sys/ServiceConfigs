Describe 'SSH script integration' {
  $scriptsRoot = Split-Path -Parent $PSScriptRoot
  $updatePath = Join-Path $scriptsRoot 'Update-LxcFromSshConfig.ps1'
  $deployPath = Join-Path $scriptsRoot 'deploy-Bashrc.ps1'

  function Get-ParameterNames {
    param([string]$Path)

    $errors = $null
    $tokens = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) {
      throw "Parserfehler in $Path"
    }

    return @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
  }

  It 'declares the shared host selection parameters on both scripts' {
    $updateParameters = Get-ParameterNames -Path $updatePath
    $deployParameters = Get-ParameterNames -Path $deployPath

    foreach ($name in @('SshConfigPath', 'IncludeHosts', 'ExcludeHosts', 'ShowConfigFiles')) {
      ($updateParameters -contains $name) | Should Be $true
      ($deployParameters -contains $name) | Should Be $true
    }

    ($updateParameters -contains 'FirstOutputTimeoutSeconds') | Should Be $false
  }

  It 'dot-sources the shared helper from both scripts' {
    (Get-Content -LiteralPath $updatePath -Raw) | Should Match 'Helpers[\\/]+SshHelpers\.ps1'
    (Get-Content -LiteralPath $deployPath -Raw) | Should Match 'Helpers[\\/]+SshHelpers\.ps1'
  }

  It 'passes SshConfigPath through to the shared helper calls' {
    $updateContent = Get-Content -LiteralPath $updatePath -Raw
    $deployContent = Get-Content -LiteralPath $deployPath -Raw

    $updateContent | Should Match 'Resolve-SshHosts\s+-SshConfigPath\s+\$SshConfigPath'
    $updateContent | Should Match 'Invoke-SshCommand\s+-HostAlias\s+\$h\s+-SshConfigPath\s+\$SshConfigPath'
    $updateContent | Should Match 'Invoke-SshCommand[\s\S]+-ForceTty'
    $updateContent | Should Not Match 'FirstOutputTimeoutSeconds'

    $deployContent | Should Match 'Resolve-SshHosts\s+-SshConfigPath\s+\$SshConfigPath'
    $deployContent | Should Match 'Invoke-SshBashScript\s+-HostAlias\s+\$h\s+-SshConfigPath\s+\$SshConfigPath'
    $deployContent | Should Match 'Copy-SshFile\s+-HostAlias\s+\$h\s+-SshConfigPath\s+\$SshConfigPath'
  }

  It 'does not use the reserved automatic variable Host as a local variable name' {
    $updateContent = Get-Content -LiteralPath $updatePath -Raw
    $deployContent = Get-Content -LiteralPath $deployPath -Raw

    $updateContent | Should Not Match '(?i)\$host\b'
    $deployContent | Should Not Match '(?i)\$host\b'
  }
}
