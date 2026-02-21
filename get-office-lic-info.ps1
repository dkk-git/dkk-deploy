# --- 1. PŘÍPRAVA PROSTŘEDÍ ---
$tempDir = "C:\Temp"
if (!(Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

$logFile = "$tempDir\office_lic.log"
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
        Write-Log "Nalezen OSPP.VBS na: $path"

        # Spuštění aktivace - výstup přímo do logu
        cscript //nologo $path /act 2>&1 | ForEach-Object { Write-Log "  $_" }
        Write-Log "Aktivace provedena (/act)"

        # Výpis statusu - výstup přímo do logu
        cscript //nologo $path /dstatus 2>&1 | ForEach-Object { Write-Log "  $_" }
        Write-Log "Status zjištěn (/dstatus)"

        $found = $true
        break
    }
}

if (-not $found) {
    Write-Log "CHYBA: OSPP.vbs nebyl nalezen. Office 2016 pravděpodobně není nainstalován."
    Write-Log "========== KONEC SKRIPTU (chyba) =========="
    exit 1
}

Write-Log "Hotovo. Výsledek uložen v $logFile"
Write-Log "========== KONEC SKRIPTU =========="
exit 0
