# Konfiguration
$NasServers = @(
    [PSCustomObject]@{
        IPAddress     = "192.168.178.151"
        Credentials   = New-Object System.Management.Automation.PSCredential("admin", (ConvertTo-SecureString "WiuPh3MiUcPY43QlWRGr" -AsPlainText -Force))
        ShareMappings = @(
            [PSCustomObject]@{
                Drive  = 'M'
                Share  = 'data'
            },
            [PSCustomObject]@{
                Drive  = 'N'
                Share  = 'private'
            },
            [PSCustomObject]@{
                Drive  = 'Y'
                Share  = 'backup_copy'
            }
        )
    },
    [PSCustomObject]@{
        IPAddress     = "192.168.178.150"
        Credentials   = New-Object System.Management.Automation.PSCredential("admin", (ConvertTo-SecureString "4hBJKwHeqQE)DAKQ" -AsPlainText -Force))
        ShareMappings = @(
            [PSCustomObject]@{
                Drive  = 'Z'
                Share  = 'backup_copy'
            }
        )
    }
)

# Netzwerklaufwerk-Mapping-Verarbeitung

function Show-CustomErrorMessage {
    param(
        [string]$Message
    )

    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

foreach ($NasServer in $NasServers) {
    if (Test-Connection -ComputerName $NasServer.IPAddress -ErrorAction SilentlyContinue -Count 1) {
        Write-Host "Server $($NasServer.IPAddress) ist erreichbar." -ForegroundColor Green

        foreach ($ShareMapping in $NasServer.ShareMappings) {
            try {
				New-PSDrive -PSProvider FileSystem -Root "\\$($NasServer.IPAddress)\$($ShareMapping.Share)" -Name $ShareMapping.Drive -Persist -Credential $NasServer.Credentials -Scope "Global" -ErrorAction SilentlyContinue |Out-Null

                if ($? -eq $true) {
                    Write-Host " [SUCCESS] Laufwerk $($ShareMapping.Drive): verbunden mit \\$($NasServer.IPAddress)\$($ShareMapping.Share)" -ForegroundColor Green
                }
            }
            catch {
                Show-CustomErrorMessage "Fehler beim Erstellen des Mappings für Laufwerk $($ShareMapping.Drive): $_"
            }
        }

        Write-Host ""
    }
    else {
        Write-Host "Server $($NasServer.IPAddress) ist nicht erreichbar. Überprüfen Sie die Netzwerkverbindung." -ForegroundColor Red
    }
}