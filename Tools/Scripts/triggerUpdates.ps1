<#
.SYNOPSIS
    System maintenance script: Chocolatey, Defender, Windows Update, Microsoft Store.
.DESCRIPTION
    - Runs with proper error handling, logging, and idempotent checks
    - Enhanced visual output with colors and progress indicators
    - Installs PSWindowsUpdate if missing and uses it to install available updates
    - Defender signature update is OPTIONAL (only when -EnableDefender is specified)
    - Updates Microsoft Store apps via winget when available, otherwise opens Store UI
    - Provides switches to control behavior
.PARAMETER SkipStore
    Skip Microsoft Store app updates.
.PARAMETER EnableDefender
    Enable Microsoft Defender signature updates (default: off).
.PARAMETER LogPath
    Optional path for transcript log. Defaults to %TEMP%\Maintenance-YYYYMMDD-HHMMSS.log
.EXAMPLE
    .\triggerUpdates.ps1 -Verbose
.EXAMPLE
    .\triggerUpdates.ps1 -EnableDefender
.EXAMPLE
    .\triggerUpdates.ps1 -SkipStore -WhatIf
    Shows what actions would be performed without actually executing them.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [switch]$SkipStore,
    [switch]$EnableDefender,
    [string]$LogPath
)

#region Setup
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'
$script:StartTime = Get-Date
$script:StepCount = 0
$script:TotalSteps = 4

# Enhanced color scheme
$Colors = @{
    Title       = 'Magenta'
    Success     = 'Green'
    Warning     = 'Yellow'
    Error       = 'Red'
    Info        = 'Cyan'
    Progress    = 'Blue'
    Separator   = 'Gray'
    Highlight   = 'White'
}

function Write-ColoredOutput {
    param(
        [string]$Message,
        [string]$Color = 'White',
        [switch]$NoNewline,
        [string]$Prefix = ''
    )
    if ($Prefix) {
        Write-Host "$Prefix " -ForegroundColor $Colors.Progress -NoNewline
    }
    if ($NoNewline) {
        Write-Host $Message -ForegroundColor $Colors[$Color] -NoNewline
    } else {
        Write-Host $Message -ForegroundColor $Colors[$Color]
    }
}

function Write-Section {
    param(
        [string]$Title,
        [string]$Description = ''
    )
    $script:StepCount++
    Write-Host
    Write-Host ('═' * 60) -ForegroundColor $Colors.Separator
    Write-ColoredOutput "[$script:StepCount/$script:TotalSteps] $Title" -Color 'Title'
    if ($Description) {
        Write-ColoredOutput "    $Description" -Color 'Info'
    }
    Write-Host ('═' * 60) -ForegroundColor $Colors.Separator
    Write-Host
}

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Progress')]
        [string]$Type = 'Info',
        [string]$Icon = ''
    )

    $icons = @{
        Info     = 'ℹ️'
        Success  = '✅'
        Warning  = '⚠️'
        Error    = '❌'
        Progress = '⏳'
    }

    if (-not $Icon) {
        $Icon = $icons[$Type]
    }

    $timestamp = Get-Date -Format 'HH:mm:ss'
    Write-ColoredOutput "[$timestamp] $Icon $Message" -Color $Type
}

function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete = 0,
        [int]$Id = 1
    )
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete -Id $Id
}

function Measure-ExecutionTime {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName
    )

    Write-Status "Starte: $OperationName" -Type 'Progress'
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        & $ScriptBlock
        $stopwatch.Stop()
        Write-Status "Abgeschlossen: $OperationName (${stopwatch.ElapsedMilliseconds}ms)" -Type 'Success'
        return $true
    }
    catch {
        $stopwatch.Stop()
        Write-Status "Fehler in: $OperationName nach ${stopwatch.ElapsedMilliseconds}ms - $($_.Exception.Message)" -Type 'Error'
        return $false
    }
}

function Ensure-Admin {
    Write-Status "Prüfe administrative Rechte..." -Type 'Progress'
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-Status "Administrative Rechte erforderlich!" -Type 'Error'
        throw 'Dieses Skript muss mit administrativen Rechten ausgeführt werden.'
    }
    Write-Status "Administrative Rechte bestätigt" -Type 'Success'
}

function Start-Logging {
    if (-not $LogPath) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $LogPath = Join-Path $env:TEMP "Maintenance-$stamp.log"
    }
    Write-Status "Starte Protokollierung: $LogPath" -Type 'Info'
    if (-not $WhatIfPreference) {
        try {
            Start-Transcript -Path $LogPath -Append -IncludeInvocationHeader | Out-Null
        } catch {
            Write-Status "Konnte Transcript nicht starten: $($_.Exception.Message)" -Type 'Warning'
        }
    } else {
        Write-Status "WhatIf aktiv: Transcript wird nicht gestartet" -Type 'Info'
    }
    Write-Status "Skript gestartet um $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Type 'Info'
}

function Stop-Logging {
    if ($WhatIfPreference) { return }
    try {
        Stop-Transcript | Out-Null
        Write-Status "Protokollierung beendet" -Type 'Info'
    } catch { }
}

function Unregister-EventSubscribersForSource {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourceObject
    )
    try {
        Get-EventSubscriber | Where-Object { $_.SourceObject -eq $SourceObject } | ForEach-Object {
            try { Unregister-Event -SubscriptionId $_.SubscriptionId -Force -ErrorAction SilentlyContinue } catch { }
        }
    } catch { }
}

function Complete-Progress {
    param([string]$Activity)
    try {
        Write-Progress -Activity $Activity -Completed
    } catch { }
}

function Get-PSWindowsUpdateUpdateCapability {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ModuleInfo
    )

    # Update-Module only works if installed via PowerShellGet/Install-Module (typical paths below).
    # If module was copied/preinstalled via other mechanisms, skip Update-Module and proceed with Import-Module.
    $base = $null
    try { $base = $ModuleInfo.ModuleBase } catch { $base = $null }

    if (-not $base) { return $false }

    $paths = @(
        (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules\PSWindowsUpdate'),
        (Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules\PSWindowsUpdate')
    )

    foreach ($p in $paths) {
        if ($base -like "$p*") { return $true }
    }
    return $false
}

# Header anzeigen
Clear-Host
Write-Host ('═' * 80) -ForegroundColor $Colors.Title
Write-ColoredOutput "    SYSTEM WARTUNGSSKRIPT" -Color 'Title'
Write-ColoredOutput "    $(Get-Date -Format 'dddd, dd. MMMM yyyy - HH:mm:ss')" -Color 'Info'
Write-Host ('═' * 80) -ForegroundColor $Colors.Title

Ensure-Admin
Start-Logging
#endregion Setup

$script:chocoProcess = $null
try {
    #region Chocolatey
    function Update-Chocolatey {
        Show-Progress -Activity "Chocolatey Update" -Status "Prüfe Chocolatey Installation..." -PercentComplete 0

        $choco = Get-Command choco -ErrorAction SilentlyContinue
        if (-not $choco) {
            Write-Status "Chocolatey ist nicht installiert" -Type 'Warning'
            if ($PSCmdlet.ShouldProcess("Chocolatey", "Installieren")) {
                if ($WhatIfPreference) {
                    Write-Status "WhatIf: Chocolatey Installation wird übersprungen" -Type 'Info'
                } else {
                    Show-Progress -Activity "Chocolatey Update" -Status "Installiere Chocolatey..." -PercentComplete 25
                    Write-Status "Installiere Chocolatey..." -Type 'Progress'

                    Measure-ExecutionTime -OperationName "Chocolatey Installation" -ScriptBlock {
                        Set-ExecutionPolicy Bypass -Scope Process -Force
                        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
                    } | Out-Null

                    $choco = Get-Command choco -ErrorAction SilentlyContinue
                    if ($choco) {
                        Write-Status "Chocolatey erfolgreich installiert" -Type 'Success'
                    }
                }
            }
        } else {
            Write-Status "Chocolatey ist bereits installiert" -Type 'Success'
        }

        if ($choco) {
            Show-Progress -Activity "Chocolatey Update" -Status "Aktualisiere Chocolatey Core..." -PercentComplete 50

            Measure-ExecutionTime -OperationName "Chocolatey Core Update" -ScriptBlock {
                if ($PSCmdlet.ShouldProcess('Chocolatey', 'Upgrade chocolatey core')) {
                    if ($WhatIfPreference) {
                        Write-Status "WhatIf: choco upgrade chocolatey wird übersprungen" -Type 'Info'
                    } else {
                        $output = choco upgrade chocolatey -y 2>&1
                        Write-Verbose ($output -join "`n")
                    }
                }
            } | Out-Null

			Measure-ExecutionTime -OperationName "Chocolatey Packages Update" -ScriptBlock {
				if ($PSCmdlet.ShouldProcess('Chocolatey packages', 'Upgrade all')) {

					if ($WhatIfPreference) {
						Write-Status "WhatIf: Chocolatey Paket-Updates werden übersprungen" -Type 'Info'
						return
					}

					Write-Status "Starte Chocolatey-Paket Updates..." -Type 'Progress'

					$upgradedCount  = 0
					$currentPackage = $null

					# Chocolatey live ausführen und Output parsen
					& choco upgrade all -y --except=chocolatey 2>&1 | ForEach-Object {
						$line = $_.ToString()
						if ([string]::IsNullOrWhiteSpace($line)) { return }

						Write-Verbose $line

						$pkg   = $null
						$state = $null

						# Typische Chocolatey-Muster (best effort)
						if ($line -match '^(\S+)\s+v[\d\.]+.*\s+to\s+v[\d\.]+') {
							$pkg = $matches[1]
							$state = 'Wird aktualisiert'
							$upgradedCount++
						}
						elseif ($line -match '^(\S+)\s+has been successfully upgraded') {
							$pkg = $matches[1]
							$state = 'Aktualisiert'
						}
						elseif ($line -match '^(\S+)\s+v[\d\.]+ is the latest version available') {
							$pkg = $matches[1]
							$state = 'Bereits aktuell'
						}
						elseif ($line -match 'Installing|Downloading|Extracting') {
							$state = 'In Bearbeitung'
						}

						if ($pkg) { $currentPackage = $pkg }

						if ($state) {
							$pkgDisplay = if ($currentPackage) { $currentPackage } else { '-' }
							Show-Progress `
								-Activity "Chocolatey Update" `
								-Status ("Paket: {0} | Status: {1}" -f $pkgDisplay, $state) `
								-PercentComplete 75

							switch ($state) {
								'Wird aktualisiert' { Write-Status "Aktualisiere: $currentPackage" -Type 'Progress' }
								'Aktualisiert'     { Write-Status "Abgeschlossen: $currentPackage" -Type 'Success' }
								'Bereits aktuell'  { Write-Status "Bereits aktuell: $currentPackage" -Type 'Info' }
							}
						}
					}

					$exitCode = $LASTEXITCODE
					if ($exitCode -eq 0) {
						if ($upgradedCount -gt 0) {
							Write-Status "Chocolatey-Update abgeschlossen: $upgradedCount Pakete aktualisiert" -Type 'Success'
						} else {
							Write-Status "Alle Chocolatey-Pakete sind bereits aktuell" -Type 'Info'
						}
					} else {
						Write-Status "Chocolatey-Update mit Fehlern beendet (Exit Code: $exitCode)" -Type 'Warning'
					}
				}
			} | Out-Null
        }

        Show-Progress -Activity "Chocolatey Update" -Status "Abgeschlossen" -PercentComplete 100
        Start-Sleep -Milliseconds 300
        Complete-Progress -Activity "Chocolatey Update"
    }

    Write-Section "Chocolatey Apps Update" "Aktualisiere Chocolatey und installierte Pakete"
    Update-Chocolatey
    #endregion Chocolatey

    #region Defender
    function Invoke-DefenderUpdate {
        [CmdletBinding()]
        param()

        if ($WhatIfPreference) {
            Write-Status "WhatIf: Defender Signatur-Update wird übersprungen" -Type 'Info'
            return $true
        }

        Show-Progress -Activity "Microsoft Defender Update" -Status "Prüfe aktuelle Signaturen..." -PercentComplete 0

        $success = $false

        try {
            $before = (Get-MpComputerStatus).AntispywareSignatureVersion
            Write-Status "Aktuelle Signatur-Version: $before" -Type 'Info'
        } catch {
            $before = $null
            Write-Status "Konnte aktuelle Signatur-Version nicht ermitteln" -Type 'Warning'
        }

        Show-Progress -Activity "Microsoft Defender Update" -Status "MpCmdRun.exe" -PercentComplete 75

        $mpPath = Join-Path $env:ProgramFiles 'Windows Defender'
        if (-not (Test-Path $mpPath)) { $mpPath = Join-Path $env:ProgramFiles 'Microsoft Defender' }
        $mpExe = Join-Path $mpPath 'MpCmdRun.exe'

        if (Test-Path $mpExe) {
            $success = Measure-ExecutionTime -OperationName "MpCmdRun.exe -SignatureUpdate" -ScriptBlock {
                if ($PSCmdlet.ShouldProcess($mpExe, 'SignatureUpdate')) {
                    $process = Start-Process -FilePath $mpExe -ArgumentList "-SignatureUpdate -MMPC" -PassThru -Wait -NoNewWindow
                    if ($process.ExitCode -ne 0) {
                        throw "MpCmdRun.exe beendete mit Exit-Code $($process.ExitCode)"
                    }
                }
            }
        } else {
            Write-Status "MpCmdRun.exe nicht gefunden in $mpPath" -Type 'Warning'
            $success = $false
        }

        try {
            $after = (Get-MpComputerStatus).AntispywareSignatureVersion
            if ($before -and $after) {
                if ($before -eq $after) {
                    Write-Status "Signaturen waren bereits aktuell (Version: $after)" -Type 'Info'
                } else {
                    Write-Status "Signaturen aktualisiert: $before → $after" -Type 'Success'
                }
            } else {
                Write-Status "Signatur-Update durchgeführt" -Type 'Success'
            }
        } catch {
            if ($success) {
                Write-Status "Defender-Update abgeschlossen" -Type 'Success'
            }
        }

        Show-Progress -Activity "Microsoft Defender Update" -Status "Abgeschlossen" -PercentComplete 100
        Start-Sleep -Milliseconds 300
        Complete-Progress -Activity "Microsoft Defender Update"

        return $success
    }

    if ($EnableDefender) {
        Write-Section "Microsoft Defender Signaturen" "Aktualisiere Antivirus-Signaturen"
        $defenderSuccess = Invoke-DefenderUpdate

        if (-not $defenderSuccess) {
            Write-Status "Defender-Update mit Fehlern beendet. Überprüfe:" -Type 'Warning'
            Write-ColoredOutput "  • Konkurrierende Antivirus-Software" -Color 'Warning'
            Write-ColoredOutput "  • Windows Defender Ereignisprotokoll" -Color 'Warning'
            Write-ColoredOutput "  • Internetverbindung" -Color 'Warning'
        }
    } else {
        Write-Status "Defender-Update übersprungen (kein -EnableDefender angegeben)" -Type 'Info'
        $defenderSuccess = $false
    }
    #endregion Defender

    #region PSWindowsUpdate
    function Update-Windows {
        Show-Progress -Activity "Windows Update" -Status "Bereite Windows Update vor..." -PercentComplete 0

        Write-Status "Prüfe PSGallery Konfiguration" -Type 'Progress'
        try {
            $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
            Write-Status "PSGallery gefunden (Policy: $($repo.InstallationPolicy))" -Type 'Info'
        } catch {
            Write-Status "PSGallery nicht verfügbar oder nicht lesbar: $($_.Exception.Message)" -Type 'Warning'
        }

        Show-Progress -Activity "Windows Update" -Status "Prüfe PSWindowsUpdate Modul..." -PercentComplete 20

        $mod = Get-Module -ListAvailable PSWindowsUpdate | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $mod) {
            Write-Status "PSWindowsUpdate nicht gefunden - installiere Modul" -Type 'Warning'
            Show-Progress -Activity "Windows Update" -Status "Installiere PSWindowsUpdate..." -PercentComplete 40

            if ($PSCmdlet.ShouldProcess('PSWindowsUpdate', 'Install-Module (CurrentUser)')) {
                if ($WhatIfPreference) {
                    Write-Status "WhatIf: Install-Module PSWindowsUpdate wird übersprungen" -Type 'Info'
                } else {
                    Measure-ExecutionTime -OperationName "PSWindowsUpdate Installation" -ScriptBlock {
                        Install-Module PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    } | Out-Null
                    Write-Status "PSWindowsUpdate erfolgreich installiert" -Type 'Success'
                }
            }
        } else {
            Write-Status "PSWindowsUpdate vorhanden (Version: $($mod.Version))" -Type 'Info'
            Show-Progress -Activity "Windows Update" -Status "Optionales Modul-Update..." -PercentComplete 40

            $canUpdate = Get-PSWindowsUpdateUpdateCapability -ModuleInfo $mod
            if ($canUpdate) {
                if ($PSCmdlet.ShouldProcess('PSWindowsUpdate', 'Update-Module')) {
                    if ($WhatIfPreference) {
                        Write-Status "WhatIf: Update-Module PSWindowsUpdate wird übersprungen" -Type 'Info'
                    } else {
                        Measure-ExecutionTime -OperationName "PSWindowsUpdate Update" -ScriptBlock {
                            Update-Module PSWindowsUpdate -Force -ErrorAction Stop
                        } | Out-Null
                    }
                }
            } else {
                Write-Status "Überspringe Modul-Update (nicht über Install-Module installiert)" -Type 'Warning'
            }
        }

        Show-Progress -Activity "Windows Update" -Status "Lade PSWindowsUpdate Modul..." -PercentComplete 60
        try {
            Import-Module PSWindowsUpdate -ErrorAction Stop
            Write-Status "PSWindowsUpdate Modul geladen" -Type 'Success'
        } catch {
            Write-Status "PSWindowsUpdate konnte nicht geladen werden: $($_.Exception.Message)" -Type 'Error'
            throw
        }

        Show-Progress -Activity "Windows Update" -Status "Suche nach verfügbaren Updates..." -PercentComplete 70
        Write-Status "Suche nach verfügbaren Windows Updates..." -Type 'Progress'

        if ($WhatIfPreference) {
            Write-Status "WhatIf: Get-WindowsUpdate Suche/Installation wird übersprungen" -Type 'Info'
            Show-Progress -Activity "Windows Update" -Status "Abgeschlossen" -PercentComplete 100
            Start-Sleep -Milliseconds 200
            Complete-Progress -Activity "Windows Update"
            return
        }

        $availableUpdates = $null
        try {
            $availableUpdates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -ErrorAction Stop
        } catch {
            Write-Status "Windows Update Suche fehlgeschlagen: $($_.Exception.Message)" -Type 'Error'
            throw
        }

        if ($availableUpdates) {
            Write-Status "Gefundene Updates ($($availableUpdates.Count)):" -Type 'Info'
            foreach ($update in $availableUpdates) {
                Write-ColoredOutput "  • $($update.Title)" -Color 'Info'
            }

            Show-Progress -Activity "Windows Update" -Status "Installiere Updates..." -PercentComplete 85

            $wuParams = @{
                Install         = $true
                AcceptAll       = $true
                IgnoreReboot    = $true
                MicrosoftUpdate = $true
                Verbose         = $false
            }

            if ($PSCmdlet.ShouldProcess('Windows Update', 'Get-WindowsUpdate -Install')) {
                Measure-ExecutionTime -OperationName "Windows Updates Installation" -ScriptBlock {
                    $results = Get-WindowsUpdate @wuParams
                    if ($results) {
                        Write-Status "Updates verarbeitet:" -Type 'Success'
                        foreach ($result in $results) {
                            if ($result.Result -eq 'Installed') {
                                Write-ColoredOutput "  ✅ $($result.Title)" -Color 'Success'
                            } else {
                                Write-ColoredOutput "  ⚠️  $($result.Title) - Status: $($result.Result)" -Color 'Warning'
                            }
                        }
                    } else {
                        Write-Status "Keine installierbaren Updates (oder keine Rückgabe vom Cmdlet)" -Type 'Info'
                    }
                } | Out-Null
            }
        } else {
            Write-Status "Keine Windows Updates verfügbar" -Type 'Info'
        }

        Show-Progress -Activity "Windows Update" -Status "Abgeschlossen" -PercentComplete 100
        Start-Sleep -Milliseconds 300
        Complete-Progress -Activity "Windows Update"
    }

    Write-Section "Windows Updates" "Suche und installiere verfügbare Windows Updates"
    Update-Windows
    #endregion PSWindowsUpdate

    #region Store
    function Update-StoreApps {
        if ($SkipStore) {
            Write-Status "Store-Updates wurden übersprungen (-SkipStore Parameter)" -Type 'Warning'
            return
        }

        if ($WhatIfPreference) {
            Write-Status "WhatIf: Store-Update wird vollständig übersprungen" -Type 'Info'
            return
        }

        Show-Progress -Activity "Microsoft Store Update" -Status "Prüfe winget Installation..." -PercentComplete 0

        $winget = Get-Command winget -ErrorAction SilentlyContinue
        if ($winget) {
            $wgVer = $null
            try { $wgVer = (winget --version) } catch { $wgVer = $null }

            # FIX: avoid nested quotes inside interpolation
            $wgSuffix = ''
            if ($wgVer) { $wgSuffix = " - Version: $wgVer" }

            Write-Status ("winget gefunden{0}" -f $wgSuffix) -Type 'Success'

            Show-Progress -Activity "Microsoft Store Update" -Status "Suche nach Store App Updates..." -PercentComplete 25

            Write-Status "Suche nach verfügbaren Store-App Updates..." -Type 'Progress'

            Measure-ExecutionTime -OperationName "Store Apps Update Check" -ScriptBlock {
                if ($PSCmdlet.ShouldProcess('winget', 'upgrade --all --source msstore')) {
                    Show-Progress -Activity "Microsoft Store Update" -Status "Aktualisiere Store Apps..." -PercentComplete 75

                    $upgradeList = winget upgrade --source msstore 2>$null
                    $upgradeableApps = ($upgradeList | Where-Object { $_ -match 'Available' }).Count

                    if ($upgradeableApps -gt 0) {
                        Write-Status "Aktualisiere $upgradeableApps Store-Apps..." -Type 'Progress'
                        $null = winget upgrade --all --source msstore --accept-package-agreements --accept-source-agreements --silent

                        if ($LASTEXITCODE -eq 0) {
                            Write-Status "Store-Apps erfolgreich aktualisiert" -Type 'Success'
                        } else {
                            Write-Status "Einige Store-Apps konnten nicht aktualisiert werden (Exit Code: $LASTEXITCODE)" -Type 'Warning'
                        }
                    } else {
                        Write-Status "Alle Store-Apps sind bereits aktuell" -Type 'Info'
                    }
                }
            } | Out-Null
        } else {
            Write-Status "winget nicht gefunden - versuche Installation..." -Type 'Warning'
            Show-Progress -Activity "Microsoft Store Update" -Status "Installiere winget..." -PercentComplete 50

            if ($PSCmdlet.ShouldProcess("winget", "Installieren (via Chocolatey)")) {
                $choco = Get-Command choco -ErrorAction SilentlyContinue
                if (-not $choco) {
                    Write-Status "Chocolatey nicht verfügbar. Öffne Microsoft Store UI." -Type 'Warning'
                    Start-Process 'ms-windows-store://downloadsandupdates'
                } else {
                    $chocoResult = Measure-ExecutionTime -OperationName "winget Installation via Chocolatey" -ScriptBlock {
                        choco install winget -y
                    }

                    if ($chocoResult) {
                        $winget = Get-Command winget -ErrorAction SilentlyContinue
                        if ($winget) {
                            Write-Status "winget erfolgreich installiert" -Type 'Success'
                            Show-Progress -Activity "Microsoft Store Update" -Status "Aktualisiere Store Apps..." -PercentComplete 75
                            $null = winget upgrade --all --source msstore --accept-package-agreements --accept-source-agreements --silent
                        } else {
                            Write-Status "winget Installation fehlgeschlagen - öffne Microsoft Store UI" -Type 'Warning'
                            Start-Process 'ms-windows-store://downloadsandupdates'
                        }
                    } else {
                        Write-Status "winget Installation fehlgeschlagen - öffne Microsoft Store UI" -Type 'Warning'
                        Start-Process 'ms-windows-store://downloadsandupdates'
                    }
                }
            } else {
                Write-Status "Öffne Microsoft Store Updates UI" -Type 'Info'
                Start-Process 'ms-windows-store://downloadsandupdates'
            }
        }

        Show-Progress -Activity "Microsoft Store Update" -Status "Abgeschlossen" -PercentComplete 100
        Start-Sleep -Milliseconds 300
        Complete-Progress -Activity "Microsoft Store Update"
    }

    Write-Section "Microsoft Store Apps" "Aktualisiere Apps aus dem Microsoft Store"
    Update-StoreApps
    #endregion Store

    # Zusammenfassung
    $endTime = Get-Date
    $totalDuration = New-TimeSpan -Start $script:StartTime -End $endTime

    Write-Host
    Write-Host ('═' * 60) -ForegroundColor $Colors.Title
    Write-ColoredOutput "WARTUNG ABGESCHLOSSEN" -Color 'Title'
    Write-Host ('═' * 60) -ForegroundColor $Colors.Title

	$defenderStatus =
	if (-not $EnableDefender) {
		"ℹ️  Übersprungen"
	}
	elseif ($defenderSuccess) {
		"✅ Signaturen aktualisiert"
	}
	else {
		"⚠️  Mit Fehlern"
	}

	$summary = [ordered]@{
		'Chocolatey'       = if (Get-Command choco -ErrorAction SilentlyContinue) { "✅ Geprüft/aktualisiert" } else { "⚠️  Übersprungen (nicht installiert)" }
		'Windows Defender' = $defenderStatus
		'Windows Updates'  = "✅ Geprüft und verarbeitet"
		'Microsoft Store'  = if ($SkipStore) { "⚠️  Übersprungen (-SkipStore)" } else { if ($WhatIfPreference) { "ℹ️  WhatIf (übersprungen)" } else { "✅ Geprüft" } }
	}

    Write-Host
    foreach ($item in $summary.GetEnumerator()) {
        $color =
            if ($item.Value.StartsWith('✅')) { 'Success' }
            elseif ($item.Value.StartsWith('⚠️')) { 'Warning' }
            else { 'Info' }
        Write-ColoredOutput "  $($item.Key.PadRight(20)): $($item.Value)" -Color $color
    }

    Write-Host
    Write-ColoredOutput "⏱️  Gesamtdauer: $([int]$totalDuration.TotalMinutes) Min $($totalDuration.Seconds) Sek" -Color 'Info'
    Write-ColoredOutput "📅 Abgeschlossen: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Color 'Info'
}
catch {
    Write-Status "Kritischer Fehler: $($_.Exception.Message)" -Type 'Error'
    Write-ColoredOutput "Fehlerdetails:" -Color 'Error'
    Write-ColoredOutput $_.ScriptStackTrace -Color 'Error'
}
finally {
    try {
        if ($script:chocoProcess) {
            Unregister-EventSubscribersForSource -SourceObject $script:chocoProcess
        }
    } catch { }

    try { Complete-Progress -Activity "Chocolatey Update" } catch { }
    try { Complete-Progress -Activity "Microsoft Defender Update" } catch { }
    try { Complete-Progress -Activity "Windows Update" } catch { }
    try { Complete-Progress -Activity "Microsoft Store Update" } catch { }

    Stop-Logging

    Write-Host
    Write-ColoredOutput "Drücke eine beliebige Taste zum Beenden..." -Color 'Info'
    if (-not $WhatIfPreference) {
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}
