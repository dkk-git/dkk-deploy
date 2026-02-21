# --- 1. PŘÍPRAVA PROSTŘEDÍ ---
$tempDir = "C:\Temp"
if (!(Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

$outFile = "$tempDir\office_lic.txt"
$osppPaths = @(
    "C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
    "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"
)

$found = $false

# --- 2. LOGIKA KONTROLY A AKTIVACE ---
foreach ($path in $osppPaths) {
    if (Test-Path $path) {
        Write-Host "Nalezen OSPP.VBS na: $path" -ForegroundColor Yellow
        
        # Spuštění aktivace a uložení výstupu do souboru
        # Přidali jsme parametr /act pro provedení aktivace
        cscript //nologo $path /act > $outFile 2>&1
        
        # Přidání výpisu statusu do logu
        cscript //nologo $path /dstatus >> $outFile 2>&1
        
        $found = $true
        break 
    }
}

if (-not $found) {
    "OSPP.vbs nebyl nalezen. Office 2016 pravděpodobně není nainstalován." | Out-File -FilePath $outFile
    Write-Host "Chyba: OSPP.vbs nebyl nalezen." -ForegroundColor Red
    exit 1
}

Write-Host "Hotovo. Výsledek uložen v $outFile" -ForegroundColor Green
exit 0
