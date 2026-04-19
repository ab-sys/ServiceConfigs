$helperPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Helpers\SshHelpers.ps1'
. $helperPath

Describe 'Resolve-SshHosts' {
  It 'resolves includes recursively with quoted specs, globs, and spaces' {
    $rootDir = Join-Path $TestDrive 'ssh'
    $includeDir = Join-Path $rootDir 'configs with spaces'
    $nestedDir = Join-Path $rootDir 'nested'

    $null = New-Item -ItemType Directory -Path $includeDir -Force
    $null = New-Item -ItemType Directory -Path $nestedDir -Force

    Set-Content -LiteralPath (Join-Path $rootDir 'config') -Value @(
      'Host root-host'
      'Include "configs with spaces\*.conf" nested\extra.conf'
    )

    Set-Content -LiteralPath (Join-Path $includeDir '01-base.conf') -Value @(
      'Host app1'
      'Include ..\nested\deep.conf'
    )

    Set-Content -LiteralPath (Join-Path $nestedDir 'extra.conf') -Value 'Host app2'
    Set-Content -LiteralPath (Join-Path $nestedDir 'deep.conf') -Value 'Host app3'

    $result = Resolve-SshHosts -SshConfigPath (Join-Path $rootDir 'config')

    (@($result.ConfigFiles) | ForEach-Object { Split-Path -Leaf $_ } | Sort-Object) | Should Be @('01-base.conf', 'config', 'deep.conf', 'extra.conf')
    @($result.Hosts) | Should Be @('app1', 'app2', 'app3', 'root-host')
  }

  It 'ignores wildcard and negated host patterns' {
    $rootDir = Join-Path $TestDrive 'wildcards'
    $null = New-Item -ItemType Directory -Path $rootDir -Force

    Set-Content -LiteralPath (Join-Path $rootDir 'config') -Value @(
      'Host *'
      'Host valid *.example !skip alias-two'
      'Host exact'
    )

    $result = Resolve-SshHosts -SshConfigPath (Join-Path $rootDir 'config')

    @($result.Hosts) | Should Be @('alias-two', 'exact', 'valid')
  }

  It 'applies include filters before exclude filters case-insensitively' {
    $rootDir = Join-Path $TestDrive 'filters'
    $null = New-Item -ItemType Directory -Path $rootDir -Force

    Set-Content -LiteralPath (Join-Path $rootDir 'config') -Value @(
      'Host One one TWO three'
    )

    $result = Resolve-SshHosts -SshConfigPath (Join-Path $rootDir 'config') -IncludeHosts @('one', 'TWO', 'missing') -ExcludeHosts @('two')

    @($result.Hosts) | Should Be @('One')
  }
}

Describe 'Invoke-SshCommand implementation' {
  It 'uses a direct ssh invocation without the watchdog runner' {
    $helperContent = Get-Content -LiteralPath $helperPath -Raw

    $helperContent | Should Match 'function\s+Invoke-SshCommand'
    $helperContent | Should Match '&\s+ssh\s+@sshArgs'
    $helperContent | Should Not Match 'Invoke-ConsoleCommandWithStartWatchdog'
    $helperContent | Should Not Match 'FirstOutputTimeoutSeconds'
  }
}
