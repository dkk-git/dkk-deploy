# --- 1. PŘÍPRAVA PROSTŘEDÍ ---
$tempDir = "C:\Temp"
if (!(Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

$logFile = "$tempDir\office_lic.log"
$outFile = "$tempDir\office_lic.txt"
$osppPaths = @(
    "C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
    "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"
)

# Funkce pro logování
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp  $Message" | Add-Content -Path $logFile
}

Write-Log "========== SPUŠTĚNÍ SKRIPTU =========="

$found = $false

# --- 2. LOGIKA KONTROLY A AKTIVACE ---
foreach ($path in $osppPaths) {
    if (Test-Path $path) {
        Write-Host "Nalezen OSPP.VBS na: $path" -ForegroundColor Yellow
        Write-Log "Nalezen OSPP.VBS na: $path"

        # Spuštění aktivace a uložení výstupu do souboru
        cscript //nologo $path /act > $outFile 2>&1
        Write-Log "Aktivace provedena (/act)"

        # Přidání výpisu statusu do logu
        cscript //nologo $path /dstatus >> $outFile 2>&1
        Write-Log "Status zjištěn (/dstatus)"

        # Přidání obsahu výstupu do logu
        Get-Content $outFile | ForEach-Object { Write-Log "  $_" }

        $found = $true
        break
    }
}

if (-not $found) {
    $msg = "OSPP.vbs nebyl nalezen. Office 2016 pravděpodobně není nainstalován."
    $msg | Out-File -FilePath $outFile
    Write-Host "Chyba: OSPP.vbs nebyl nalezen." -ForegroundColor Red
    Write-Log "CHYBA: OSPP.vbs nebyl nalezen."
    Write-Log "========== KONEC SKRIPTU (chyba) =========="
    exit 1
}

Write-Host "Hotovo. Výsledek uložen v $outFile" -ForegroundColor Green
Write-Log "Hotovo. Výsledek uložen v $outFile"
Write-Log "========== KONEC SKRIPTU =========="
exit 0
