# --- 1. KONFIGURACE STAHOVÁNÍ ---
$token = "github_pat_TVUJ_DLOUHY_TOKEN" # Sem vlož svůj vygenerovaný token
$url = "https://raw.githubusercontent.com/uzivatel/repo/main/tvuj_skript.bat"
$tempDir = "C:\Temp"
$downloadedFile = "$tempDir\office_lic_downloaded.bat"

# --- 2. PŘÍPRAVA PROSTŘEDÍ A STAHOVÁNÍ ---
if (!(Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/vnd.github.v3.raw"
}

try {
    Write-Host "Stahuji skript z GitHubu..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $url -Headers $headers -OutFile $downloadedFile -ErrorAction Stop
    Write-Host "Staženo úspěšně." -ForegroundColor Green
}
catch {
    Write-Error "Nepodařilo se stáhnout soubor z GitHubu. Prověř token a URL. Chyba: $_"
    exit 1
}

# --- 3. LOGIKA KONTROLY OFFICE (PŘEPSANÝ BAT SKRIPT) ---
$outFile = "$tempDir\office_lic.txt"
$osppPaths = @(
    "C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
    "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"
)

$found = $false

foreach ($path in $osppPaths) {
    if (Test-Path $path) {
        Write-Host "Nalezen OSPP.VBS na: $path" -ForegroundColor Yellow
        # Spuštění cscriptu a uložení výstupu do souboru
        cscript //nologo $path /dstatus > $outFile 2>&1
        $found = $true
        break # Pokud ho najdeme, končíme smyčku
    }
}

if (-not $found) {
    "OSPP.vbs nebyl nalezen" | Out-File -FilePath $outFile
    Write-Host "Chyba: OSPP.vbs nebyl nalezen." -ForegroundColor Red
    exit 1
}

Write-Host "Hotovo. Výsledek uložen v $outFile" -ForegroundColor Green
exit 0
