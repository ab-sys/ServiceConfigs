$root = 'C:\Users\AlessandroBello\OneDrive - AB-Systems'
$root2 = 'C:\Users\AlessandroBello\OneDrive'

$xsizetot = 0

get-childitem $root -Force -File -Recurse |
where Attributes -eq 'Archive, ReparsePoint' |
foreach {
    $xpath = $_.fullname
    $xsize = [Math]::Round($_.Length / 1024 / 1024, 2)
    $xsizetot += $_.Length

    Write-Host "$xpath ($xsize MB)"
    attrib $_.fullname +U -P
}

get-childitem $root2 -Force -File -Recurse |
where Attributes -eq 'Archive, ReparsePoint' |
foreach {
    $xpath = $_.fullname
    $xsize = [Math]::Round($_.Length / 1024 / 1024, 2)
    $xsizetot += $_.Length

    Write-Host "$xpath ($xsize MB)"
    attrib $_.fullname +U -P
}

$xsizetot = [Math]::Round($xsizetot / 1024 / 1024, 2)

Write-Host
Write-Host "Total $xsizetot MB freed on $root"