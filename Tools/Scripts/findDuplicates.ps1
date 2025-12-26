# Erweitertes Skript zum Finden und Löschen von Duplikaten basierend auf Dateiinhalt
# Ignoriert den Ordner ".dtrash", zeigt das erste gefundene File in grün und fragt nach Bestätigung

Write-Host "Suche nach Duplikaten anhand des Dateiinhalts..." -ForegroundColor Green

# Hole das aktuelle Verzeichnis
$currentPath = Get-Location

# Funktion zur Berechnung des SHA256-Hashes
function Get-FileHash {
    param(
        [string]$FilePath
    )
    
    $fileStream = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    try {
        $hasher = New-Object System.Security.Cryptography.SHA256Managed
        $bytes = $hasher.ComputeHash($fileStream)
        $hashString = [BitConverter]::ToString($bytes) -replace '-',''
        return $hashString
    }
    finally {
        $fileStream.Close()
    }
}

# Sammle alle Dateien rekursiv, aber ignoriere den .dtrash-Ordner
$files = Get-ChildItem -Path $currentPath -Recurse -File | Where-Object { $_.Directory.Name -ne ".dtrash" }

# Erstelle ein Array mit Dateinamen und Hashes
$fileHashes = @()

Write-Host "Verarbeite Dateien..." -ForegroundColor Yellow

foreach ($file in $files) {
    try {
        $hash = Get-FileHash -FilePath $file.FullName
        $fileHashes += [PSCustomObject]@{
            Name = $file.Name
            FullName = $file.FullName
            Hash = $hash
        }
    }
    catch {
        Write-Warning "Fehler beim Verarbeiten von $($file.FullName): $($_.Exception.Message)"
    }
}

# Gruppiere nach Hash (Duplikate)
$hashGroups = $fileHashes | Group-Object Hash

# Finde Duplikate
$duplicates = $hashGroups | Where-Object { $_.Count -gt 1 }

if ($duplicates.Count -eq 0) {
    Write-Host "Keine Duplikate gefunden." -ForegroundColor Green
} else {
    Write-Host "`nDuplikate basierend auf Dateiinhalt:" -ForegroundColor Yellow
    Write-Host "===================================" -ForegroundColor Yellow
    
    $duplicateCount = 0
    $filesToDelete = @()
    $totalSizeDeleted = 0

    foreach ($group in $duplicates) {
        if ($group.Count -gt 1) {
            $duplicateCount++
            Write-Host "`nGruppe $($duplicateCount):" -ForegroundColor Red
            Write-Host "Hash: $($group.Name)" -ForegroundColor Blue
            
            # Zeige alle Dateien mit diesem Hash an
            $firstFile = $true
            $filesInGroup = @()
            
            $group.Group | ForEach-Object {
                if ($firstFile) {
                    # Erstes File in grün anzeigen (wird behalten)
                    Write-Host "  [BEHALTEN] $($_.FullName)" -ForegroundColor Green
                    $firstFile = $false
                    $filesInGroup += $_
                } else {
                    # Alle anderen in rot anzeigen (werden gelöscht)
                    Write-Host "  [ZU LÖSCHEN] $($_.FullName)" -ForegroundColor Red
                    $filesInGroup += $_
                }
            }
            
            # Füge die zu löschenden Dateien zur Liste hinzu
            if ($filesInGroup.Count -gt 1) {
                # Das erste File wird behalten, alle anderen werden gelöscht
                $filesToDelete += $filesInGroup[1..($filesInGroup.Count - 1)]
            }
        }
    }
    
    if ($filesToDelete.Count -gt 0) {
        Write-Host "`nZu löschende Dateien:" -ForegroundColor Yellow
        
        # Berechne den Gesamtspeicherplatz
        $totalSizeDeleted = 0
        $filesToDelete | ForEach-Object {
            $fileInfo = Get-Item $_.FullName
            $totalSizeDeleted += $fileInfo.Length
            Write-Host "  $($_.FullName) ($([Math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Red
        }
        
        # Formatierung des Speicherplatzes
        $sizeInMB = [Math]::Round($totalSizeDeleted / 1MB, 2)
        $sizeInGB = [Math]::Round($totalSizeDeleted / (1024*1024*1024), 2)
        
        Write-Host "`nGesamter Speicherplatz: $($sizeInMB) MB ($($sizeInGB) GB)" -ForegroundColor Yellow
        
        Write-Host ""
        $confirmation = Read-Host "Möchten Sie wirklich diese $(($filesToDelete.Count)) Datei(en) löschen? (ja/nein)"
        
        if ($confirmation -eq "ja" -or $confirmation -eq "j") {
            $deletedCount = 0
            foreach ($file in $filesToDelete) {
                try {
                    Remove-Item $file.FullName -Force
                    Write-Host "Gelöscht: $($file.FullName)" -ForegroundColor Cyan
                    $deletedCount++
                }
                catch {
                    Write-Warning "Fehler beim Löschen von $($file.FullName): $($_.Exception.Message)"
                }
            }
            Write-Host "`nErfolgreich $(($filesToDelete.Count)) Datei(en) gelöscht." -ForegroundColor Green
            Write-Host "Freigegebener Speicherplatz: $($sizeInMB) MB ($($sizeInGB) GB)" -ForegroundColor Green
        } else {
            Write-Host "Löschung abgebrochen." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Keine Dateien zum Löschen gefunden." -ForegroundColor Green
    }
}

Write-Host "`nSkript abgeschlossen." -ForegroundColor Green

# Optional: Pause, um Ergebnisse zu sehen
Read-Host "Drücken Sie eine beliebige Taste zum Beenden..."
