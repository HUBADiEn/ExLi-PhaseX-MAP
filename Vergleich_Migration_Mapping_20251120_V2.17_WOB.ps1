<#
.SYNOPSIS
    Vergleicht zwei grosse CSV-Dateien (PMS und PIM) anhand der EAN, fasst die Daten zusammen
    und führt Prüfungen auf den kombinierten Datensätzen durch.

.NOTES
    Version: Mapping_V2.17
    Änderungen gegenüber V2.16:
    - Memory-Optimierung: Alle Statistiken werden während der Check-Schleife berechnet (statt Where-Object)
    - Memory-Optimierung: Zusammenfassung verwendet nur vorberechnete Zähler (kein zusätzlicher RAM-Bedarf)
    - Memory-Optimierung: $All_Datasets wird nach Export freigegeben (reduziert Peak-RAM)
    - Bugfix: OutOfMemoryException in Pause-Ende Funktion behoben
#>

# --- Bootstrap: falls via Explorer gestartet, neue Konsole mit -NoExit oeffnen (CMD /K) ---
if (-not $env:PS_KEEP_NOEXIT) {
    try {
        $env:PS_KEEP_NOEXIT = '1'
        $scriptPath = $MyInvocation.MyCommand.Definition
        if (-not (Test-Path -LiteralPath $scriptPath)) { throw "Scriptpfad ungültig" }
        $quoted = '"' + $scriptPath.Replace('"','""') + '"'
        $arguments = '/k powershell.exe -NoLogo -NoExit -ExecutionPolicy Bypass -File ' + $quoted
        Start-Process -FilePath 'cmd.exe' -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $scriptPath)
    } catch { $global:__ForcePauseAtEnd = $true }
    return
}

# Variable initialisieren
$global:__ForcePauseAtEnd = $false
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Windows.Forms

#region ========================= BENUTZER-LOOKUP-TABELLE =========================
$UserLookupTable = @{
    'M0733302' = 'WOB'
    'M0779325' = 'AZG'
    'M0555315' = 'CPA'
}
#endregion

#region ========================= LIEFERANTEN-LOOKUP-TABELLE =========================
$SupplierLookupTable = @{
    '16409132'='AVA Verlagsauslieferung'
    '16801357'='Bremer Versandwerk GmbH'
    '16409120'='Buchzentrum'
    '16800790'='Carletto AG'
    '16517649'='Carlit + Ravensburger AG'
    '16803558'='ciando (Agency)'
    '16803554'='ciando GmbH'
    '15642908'='Ex Libris AG Dietikon 1'
    '16450805'='Grüezi Music AG'
    '16802683'='Libri (Agency)'
    '16776945'='Libri GmbH'
    '16803735'='Max Bersinger AG'
    '16801413'='MFP Tonträger'
    '16407363'='Musikvertrieb AG'
    '30000023'='Office World (Oridis)'
    '16409618'='OLF S.A.'
    '16411177'='Phonag Records AG'
    '16558172'='Phono-Vertrieb'
    '16212120'='Rainbow Home Entertainment'
    '16526960'='Sombo AG'
    '16699796'='Sony Music Entertainment'
    '16407336'='Starworld Enterprise GmbH'
    '16423780'='Thali AG'
    '16486030'='Universal Music GmbH'
    '30000223'='Vedes Grosshandel GmbH'
    '16706931'='Waldmeier AG'
    '16797703'='Warner Music Group'
    '16435880'='Zeitfracht Medien GmbH'
}
#endregion

# Flag für die finale Erfolgsprüfung
$scriptSuccessfullyCompleted = $false

# ========================= Excel-Formatierung (Freeze/Color/AutoFit) =========================
function ConvertTo-ExcelColumnName([int]$index) {
    $div = $index; $colName = ""
    while ($div -gt 0) {
        $mod = ($div - 1) % 26
        $colName = [char](65 + $mod) + $colName
        $div = [math]::Floor(($div - $mod) / 26)
    }
    return $colName
}

function Apply-WorksheetFormatting {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string[]]$SheetNames
    )

    try { [void][OfficeOpenXml.ExcelPackage] } catch { Add-Type -AssemblyName OfficeOpenXml -ErrorAction SilentlyContinue }
    try { [OfficeOpenXml.ExcelPackage]::set_LicenseContext([OfficeOpenXml.LicenseContext]::NonCommercial) } catch {}

    $fi  = [System.IO.FileInfo]::new($Path)
    $pkg = [OfficeOpenXml.ExcelPackage]::new($fi)
    try {
        foreach ($name in $SheetNames) {
            $ws = $pkg.Workbook.Worksheets[$name]
            if (-not $ws) { continue }
            if (-not $ws.Dimension) { continue }

            $ws.View.FreezePanes(3,2)
            $ws.Cells[$ws.Dimension.Address].AutoFitColumns()

            $lastCol    = $ws.Dimension.End.Column
            $lastLetter = ConvertTo-ExcelColumnName $lastCol

            $r1 = $ws.Cells["A1:$lastLetter`1"]
            $r1.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $r1.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(230,230,230))

            $r2 = $ws.Cells["A2:$lastLetter`2"]
            $r2.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $r2.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(242,242,242))
        }
        $pkg.Save()
    } finally {
        $pkg.Dispose()
    }
}

# ========================= STREAMING CSV EXPORT =========================
function Export-CsvStreaming {
    param(
        [Parameter(Mandatory=$true)]$Data,
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$ActivityName,
        [string[]]$ExcludeProperties = @(),
        [int]$BatchSize = 50000
    )
    
    $writer = $null
    try {
        $writer = [System.IO.StreamWriter]::new($Path, $false, [System.Text.Encoding]::UTF8)
        
        # Header erstellen
        $firstItem = $Data | Select-Object -First 1
        if (-not $firstItem) { 
            Write-Warning "Keine Daten zum Exportieren für $ActivityName"
            return $false
        }
        
        $properties = $firstItem.PSObject.Properties.Name | Where-Object { $ExcludeProperties -notcontains $_ }
        $header = ($properties -join ';')
        $writer.WriteLine($header)
        
        $counter = 0
        $totalCount = if ($Data -is [Array]) { $Data.Count } else { 1 }
        
        foreach ($item in $Data) {
            $counter++
            
            # Progress und GC alle BatchSize Zeilen
            if ($counter % $BatchSize -eq 0) {
                $percentage = [Math]::Min(99, [Math]::Floor(($counter / $totalCount) * 100))
                Write-Progress -Activity $ActivityName -Status "$counter von $totalCount Zeilen geschrieben" -PercentComplete $percentage
                $writer.Flush()
                [System.GC]::Collect()
            }
            
            # Zeile erstellen
            $values = @()
            foreach ($prop in $properties) {
                $value = $item.$prop
                if ($null -eq $value) {
                    $values += ''
                } else {
                    $valueStr = $value.ToString()
                    # CSV-Escaping: Wenn Semikolon, Anführungszeichen, Zeilenumbruch oder Komma enthalten
                    if ($valueStr -match '[";,\r\n]') {
                        $valueStr = '"' + $valueStr.Replace('"', '""') + '"'
                    }
                    $values += $valueStr
                }
            }
            $line = $values -join ';'
            $writer.WriteLine($line)
        }
        
        Write-Progress -Activity $ActivityName -Completed
        return $true
        
    } catch {
        Write-Warning "Fehler beim Streaming-Export: $($_.Exception.Message)"
        return $false
    } finally {
        if ($writer) {
            $writer.Flush()
            $writer.Close()
            $writer.Dispose()
        }
    }
}

try {
    # --- Konsolenfenster vergrössern ---
    try {
        if (-not $psISE) {
            $currentSize = $Host.UI.RawUI.WindowSize
            $newHeight = [int]($currentSize.Height * 1.5)
            if ($Host.UI.RawUI.BufferSize.Height -lt $newHeight) {
                $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($Host.UI.RawUI.BufferSize.Width, $newHeight)
            }
            $Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size($currentSize.Width, $newHeight)
        }
    } catch {}

    #region ========================= KONFIG =========================
    $ScriptVersion = "Mapping_V2.17"

    # OneDrive-Basis für Input
    try {
        $oneDrivePath = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive\Accounts\Business1").UserFolder
        if (-not $oneDrivePath) { throw "Der OneDrive-Pfad konnte nicht in der Registry gefunden werden (Wert 'UserFolder' ist leer)." }
        $InputDirectory = Join-Path -Path $oneDrivePath -ChildPath "PIM\PhaseX_Mapping"
    }
    catch {
        throw "KRITISCHER FEHLER: Der OneDrive-Pfad konnte nicht ermittelt werden. Ist OneDrive for Business ('Migros') korrekt eingerichtet? Details: $($_.Exception.Message)"
    }

    # Header
    $PMS_Header_Expected = @("SLLLFN","SLLEAN","SLLPAS","SLLCAT","SLLGNR","FMBIDX","MATURE","IMPDAT","CHGDAT","GNXGNR")
    $PIM_Header_Expected = @("Lieferant","EAN","Status","Kategorie","Genre","Formatcode","Mature","letzter Import","letzte Änderung","letzter Status")
    
    # Memory-Management Konstanten
    $GC_INTERVAL = 50000
    $INITIAL_CAPACITY = 100000
    $EXPORT_BATCH_SIZE = 50000
    #endregion

    #region ========================= VARIABLEN =========================
    $createdOutputFiles = [System.Collections.Generic.List[string]]::new()
    $pmsEanCount = 0; $pimEanCount = 0
    $script:supplierNameForSummary = ""; $script:foundPimDuplicates = $false
    
    # Statistik-Zähler (für Zusammenfassung ohne Where-Object)
    $script:matchedEanCount = 0
    $script:totalDatasetCount = 0
    $script:checkSummaryErrors = 0
    $script:check1Errors = 0
    $script:check2Errors = 0
    $script:check3Errors = 0
    $script:check4Errors = 0
    $script:check5Errors = 0
    #endregion

    #region ========================= FUNKTIONEN =========================
    function Get-FilePathDialog { param([string]$WindowTitle,[string]$InitialDirectory)
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = $WindowTitle; $dialog.Filter = "CSV-Dateien (*.csv)|*.csv"; $dialog.InitialDirectory = $InitialDirectory
        if ($dialog.ShowDialog() -eq 'OK') { return $dialog.FileName }; return $null
    }

    function Invoke-CalculateTimeDifference {
        param([Parameter(Mandatory=$true)][PSCustomObject]$Dataset)
        $pmsDateString = $Dataset.PMS_CHGDAT; $pimDateString = $Dataset.'PIM_letzte Änderung'
        if ([string]::IsNullOrWhiteSpace($pmsDateString) -or [string]::IsNullOrWhiteSpace($pimDateString)) { return "fehlende Daten" }
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        try { $pmsDateTime = [datetime]::ParseExact("$pmsDateString 12:00:00","dd.MM.yy HH:mm:ss",$culture) } catch { return "PMS-Datum ungültig: '$pmsDateString'" }
        $pimDateTime = $null; $trimmed = $pimDateString.Trim()
        try { $pimDateTime = [datetime]$trimmed } catch {
            try {
                $san = $trimmed -replace '[–—]','-' -replace '\s+',' '
                $pimDateTime = [datetime]::ParseExact($san, @("yyyy-MM-dd HH:mm:ss","dd.MM.yyyy HH:mm:ss"), $culture, [System.Globalization.DateTimeStyles]::None)
            } catch {
                $san2 = $pimDateString.Trim() -replace '[–—]','-' -replace '\s+',' '
                return "PIM-Datum unlesbar. Original: '$pimDateString', Bereinigt: '$san2'"
            }
        }
        [Math]::Round( ($pimDateTime - $pmsDateTime).TotalHours, 2)
    }

    function Invoke-Check1_Status { param([PSCustomObject]$Dataset) if ($Dataset.PMS_SLLPAS -eq $Dataset.PIM_Status) { "ok" } else { "nicht ok - PMS: '$($Dataset.PMS_SLLPAS)', PIM: '$($Dataset.PIM_Status)'" } }
    function Invoke-Check2_Kategorie { param([PSCustomObject]$Dataset)
        if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
        if (($Dataset.PMS_SLLCAT -eq "UKN") -and ([string]::IsNullOrEmpty($Dataset.PIM_Kategorie))) { return "ok" }
        if ($Dataset.PMS_SLLCAT -eq $Dataset.PIM_Kategorie) { return "ok" }
        "nicht ok"
    }
    function Invoke-Check3_Genre { param([PSCustomObject]$Dataset)
        $pmsStatus=$Dataset.PMS_SLLPAS; $pmsGenresRaw=$Dataset.PMS_SLLGNR; $pimGenre=$Dataset.PIM_Genre
        if ($pmsStatus -eq 'passive') { return "ok - Status = passive" }
        if ((-not [string]::IsNullOrEmpty($pmsGenresRaw)) -and ($pmsGenresRaw -notlike '*') -and ([string]::IsNullOrEmpty($pimGenre))) { return "nicht ok - Kein Genre im PIM vorhanden" }
        if ([string]::IsNullOrEmpty($pmsGenresRaw) -and [string]::IsNullOrEmpty($pimGenre)) { return "ok" }
        if (($pmsGenresRaw -like '*') -and ([string]::IsNullOrEmpty($pimGenre))) { return "nicht ok - Genre fehlt im PIM" }
        if ([string]::IsNullOrEmpty($pmsGenresRaw) -or [string]::IsNullOrEmpty($pimGenre)) { return "nicht ok" }
        $arr = $pmsGenresRaw.Trim('[]').Split(',') | ForEach-Object { $_.Trim() }
        if ($arr -contains $pimGenre) { "ok" } else { "nicht ok" }
    }
    function Invoke-Check4_Formatcode { param([PSCustomObject]$Dataset)
        $pms=$Dataset.PMS_FMBIDX; $pim=$Dataset.PIM_Formatcode
        if ($pms -eq $pim) { return "ok" }
        if ((-not [string]::IsNullOrEmpty($pms)) -and ([string]::IsNullOrEmpty($pim))) { return "nicht ok - kein Formatcode im PIM" }
        if ((-not [string]::IsNullOrEmpty($pms)) -and (-not [string]::IsNullOrEmpty($pim))) { return "nicht ok - unterschiedliche Formatcodes" }
        "nicht ok"
    }
    function Invoke-Check5_MatureContent { param([PSCustomObject]$Dataset)
        $pms=$Dataset.PMS_MATURE; $pim=$Dataset.PIM_Mature
        if (($pms -eq 'true') -and ($pim -eq 'Ja')) { return "ok" }
        if (($pms -eq 'false') -and ($pim -eq '0')) { return "ok" }
        if (($pms -eq 'true') -and ($pim -eq '1')) { return "ok" }
        if (([string]::IsNullOrEmpty($pms) -or $pms -eq 'false') -and ([string]::IsNullOrEmpty($pim) -or $pim -eq '0')) { return "ok" }
        $pmsIsSet = ($pms -eq 'true'); $pimIsSet = ($pim -eq 'Ja' -or $pim -eq '1')
        if     ($pmsIsSet -and $pimIsSet) { "nicht ok - Werte unterschiedlich (PMS: '$pms', PIM: '$pim')" }
        elseif ($pmsIsSet -and (-not $pimIsSet)) { if ([string]::IsNullOrEmpty($pim)) { "nicht ok - Mature nur im PMS gesetzt ('$pms')" } else { "nicht ok - Werte unterschiedlich (PMS: '$pms', PIM: '$pim')" } }
        elseif ((-not $pmsIsSet) -and $pimIsSet) { if ([string]::IsNullOrEmpty($pms)) { "nicht ok - Mature nur im PIM gesetzt ('$pim')" } else { "nicht ok - Werte unterschiedlich (PMS: '$pms', PIM: '$pim')" } }
        else { "nicht ok - Werte unterschiedlich (PMS: '$pms', PIM: '$pim')" }
    }
    
    function Pause-Ende {
        Write-Host ""
        $ColorOk = "Green"; $ColorNok = "Red"; $Line = "=" * 60
        
        # MEMORY-OPTIMIERT: Verwende vorberechnete Zähler statt Where-Object
        $matchedEanCount = $script:matchedEanCount
        $errorCount = $script:checkSummaryErrors
        $successCount = ($script:totalDatasetCount - $errorCount)
        
        if ($errorCount -gt 0) {
            $statusErrorCount = $script:check1Errors
            $categoryErrorCount = $script:check2Errors
            $genreErrorCount = $script:check3Errors
            $formatCodeErrorCount = $script:check4Errors
            $matureContentErrorCount = $script:check5Errors
        }
        
        $supplierDisplayString = if ($script:supplierNameForSummary -ne $pmsSupplier) { "$($script:supplierNameForSummary) ($pmsSupplier)" } else { $pmsSupplier }
        Write-Host $Line -ForegroundColor Cyan
        Write-Host "Vergleich Phase X - Mapping (Script Version $($ScriptVersion))" -ForegroundColor White
        Write-Host ""
        Write-Host "Input-Files:" -ForegroundColor Yellow
        Write-Host "  PMS: $(Split-Path $pmsFilePath -Leaf)"
        Write-Host "  PIM: $(Split-Path $pimFilePath -Leaf)"
        Write-Host ""
        Write-Host "Output-Files:" -ForegroundColor Yellow
        Write-Host "  Lokal (Original): '$LocalOutputDirectory'" -ForegroundColor Cyan
        Write-Host "  SharePoint (Kopie): '$SharePointOutputDirectory'" -ForegroundColor Cyan
        if ($createdOutputFiles.Count -gt 0) { foreach ($file in $createdOutputFiles) { Write-Host "    - $file" } } else { Write-Host "  Es wurden keine Output-Files erstellt." }
        Write-Host ""
        $runDateTime = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        Write-Host "Datum: $runDateTime" -ForegroundColor Yellow
        Write-Host "Dauer: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Zusammenfassung:" -ForegroundColor Yellow
        Write-Host "  Header-Überprüfung: OK" -ForegroundColor $ColorOk
        Write-Host "  Überprüfung der Lieferanten-Nummern: OK" -ForegroundColor $ColorOk
        Write-Host "  Anzahl EANs im PMS-File: $pmsEanCount"
        Write-Host "  Anzahl EANs im PIM-File: $pimEanCount"
        Write-Host "  Anzahl EANs in beiden Files: $matchedEanCount"
        Write-Host "  Anzahl fehlerfreie EANs: $successCount" -ForegroundColor $ColorOk
        if ($errorCount -gt 0) { Write-Host "  Anzahl EANs mit Fehlern: $errorCount" -ForegroundColor $ColorNok; $finalStatusText="Nicht ok - hat Fehler"; $finalStatusColor=$ColorNok } else { Write-Host "  Anzahl EANs mit Fehlern: 0" -ForegroundColor $ColorOk; $finalStatusText="OK - fehlerfrei"; $finalStatusColor=$ColorOk }
        if ($script:foundPimDuplicates) { Write-Host "  Doppelte EANs im PIM File gefunden" -ForegroundColor Red } else { Write-Host "  Keine doppelten EANs im PIM File gefunden" -ForegroundColor Green }
        Write-Host ""
        Write-Host "Mapping von Lieferant $supplierDisplayString ist $finalStatusText" -ForegroundColor $finalStatusColor
        if ($errorCount -gt 0) {
            Write-Host ""
            Write-Host "Fehler-Übersicht:" -ForegroundColor Yellow
            $statusColor = if ($statusErrorCount -gt 0) { "Red" } else { "Green" };  Write-Host "  - Anzahl Fehler bei Status: $statusErrorCount" -ForegroundColor $statusColor
            $catColor    = if ($categoryErrorCount -gt 0) { "Red" } else { "Green" };  Write-Host "  - Anzahl Fehler bei Kategorie: $categoryErrorCount" -ForegroundColor $catColor
            $genreColor  = if ($genreErrorCount -gt 0) { "Red" } else { "Green" };    Write-Host "  - Anzahl Fehler bei Genres: $genreErrorCount" -ForegroundColor $genreColor
            $formatColor = if ($formatCodeErrorCount -gt 0) { "Red" } else { "Green" };Write-Host "  - Anzahl Fehler bei Formatcode: $formatCodeErrorCount" -ForegroundColor $formatColor
            $matureColor = if ($matureContentErrorCount -gt 0) { "Red" } else { "Green" };Write-Host "  - Anzahl Fehler bei Mature Content: $matureContentErrorCount" -ForegroundColor $matureColor
        }
        Write-Host $Line -ForegroundColor Cyan
        if ($global:__ForcePauseAtEnd) { Write-Host "Hinweis: Relaunch mit eigenem Fenster war nicht möglich. Fenster bleibt hier offen." -ForegroundColor Yellow }
        [void](Read-Host "Drücke ENTER um das Fenster zu schliessen")
    }
    #endregion

    #region ========================= HAUPTTEIL =========================
    Write-Host "--- Skript-Version $($ScriptVersion) ---`n" -ForegroundColor Gray
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # 1. Verzeichnisse
    Write-Host "1. Prüfe Verzeichnisse..."
    if (-not (Test-Path -Path $InputDirectory -PathType Container)) { throw "Das angegebene Eingabeverzeichnis existiert nicht: `"$InputDirectory`"" }
    Write-Host "    Verzeichnis ist vorhanden."

    # 2. Dateien auswählen
    Write-Host "2. Bitte Dateien auswählen..."
    $absoluteInputDirectory = Convert-Path -Path $InputDirectory
    $pmsFilePath = Get-FilePathDialog -WindowTitle "Bitte die PMS-Datei auswählen" -InitialDirectory $absoluteInputDirectory; if (-not $pmsFilePath) { Write-Host "Aktion vom Benutzer abgebrochen."; exit }
    $pimFilePath = Get-FilePathDialog -WindowTitle "Bitte die PIM-Datei auswählen" -InitialDirectory $absoluteInputDirectory; if (-not $pimFilePath) { Write-Host "Aktion vom Benutzer abgebrochen."; exit }
    Write-Host "    PMS-Datei: $(Split-Path $pmsFilePath -Leaf)"
    Write-Host "    PIM-Datei: $(Split-Path $pimFilePath -Leaf)"

    # 3. Header prüfen
    Write-Host "3. Prüfe Header der CSV-Dateien..."
    Write-Host "    - Prüfe PMS-Datei..."
    $pmsHeaderLine = (Get-Content -Path $pmsFilePath -TotalCount 1).TrimEnd(';')
    if ([string]::IsNullOrWhiteSpace($pmsHeaderLine)) { throw "Die PMS-Datei '$pmsFilePath' ist leer oder die erste Zeile (Header) ist leer." }
    $pmsActualHeader = $pmsHeaderLine.Split(';')
    if ($null -ne (Compare-Object -ReferenceObject $PMS_Header_Expected -DifferenceObject $pmsActualHeader -CaseSensitive)) {
        throw "Der Header in der PMS-Datei ist nicht korrekt.`nErwartet: $($PMS_Header_Expected -join ';')`nGefunden:  $($pmsActualHeader -join ';')"
    }
    Write-Host "      -> Header in PMS-Datei ist korrekt." -ForegroundColor Green

    Write-Host "    - Prüfe PIM-Datei..."
    $pimHeaderLine = Get-Content -Path $pimFilePath -TotalCount 1 -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($pimHeaderLine)) { throw "Die PIM-Datei '$pmsFilePath' ist leer oder die erste Zeile (Header) ist leer." }
    $pimActualHeader = ($pimHeaderLine.Replace('"', '')).Split(';')
    if ($null -ne (Compare-Object -ReferenceObject $PIM_Header_Expected -DifferenceObject $pimActualHeader -CaseSensitive)) {
        throw "Der Header in der PIM-Datei ist nicht korrekt.`nErwartet: $($PIM_Header_Expected -join ';')`nGefunden:  $($pimActualHeader -join ';')"
    }
    Write-Host "      -> Header in PIM-Datei ist korrekt." -ForegroundColor Green

    # 4. Lieferanten-Check
    Write-Host "4. Führe Lieferanten-Check durch..."
    $pmsFirstDataRow = (Get-Content -Path $pmsFilePath -TotalCount 2 | Select-Object -Last 1).TrimEnd(';')
    $pmsFirstRecord = $pmsFirstDataRow | ConvertFrom-Csv -Header $PMS_Header_Expected -Delimiter ';'
    $pmsSupplier = $pmsFirstRecord.SLLLFN

    $pimFirstDataRow = (Get-Content -Path $pimFilePath -TotalCount 2 -Encoding UTF8 | Select-Object -Last 1)
    $pimFirstRecord = $pimFirstDataRow | ConvertFrom-Csv -Header $PIM_Header_Expected -Delimiter ';'
    $pimSupplier = $pimFirstRecord.Lieferant

    if ($pmsSupplier -eq $pimSupplier) {
        Write-Host "    Lieferantennummern stimmen überein: '$pmsSupplier'." -ForegroundColor Green
        $supplierName = if ($SupplierLookupTable.ContainsKey($pmsSupplier)) { $SupplierLookupTable[$pmsSupplier] } else { $pmsSupplier }
        $script:supplierNameForSummary = $supplierName
        $sanitizedSupplierName = $supplierName.Replace(' ', '-').Replace('+', '') -replace '[\\/:*?"<>|]', ''

        # Lokales Output-Verzeichnis (wo die Quell-Files liegen)
        $LocalOutputDirectory = Split-Path -Path $pmsFilePath -Parent
        
        # SharePoint Output-Verzeichnis vorbereiten
        $BaseOutputRoot = ".\VergleichsErgebnisse"
        if (-not (Test-Path -Path $BaseOutputRoot -PathType Container)) { New-Item -Path $BaseOutputRoot -ItemType Directory | Out-Null }
        $SharePointOutputDirectory = Join-Path -Path $BaseOutputRoot -ChildPath $sanitizedSupplierName
        if (-not (Test-Path -Path $SharePointOutputDirectory -PathType Container)) {
            try { New-Item -Path $SharePointOutputDirectory -ItemType Directory -ErrorAction Stop | Out-Null }
            catch { throw "FEHLER: Das SharePoint-Ausgabeverzeichnis '$SharePointOutputDirectory' konnte nicht erstellt werden. Details: $($_.Exception.Message)" }
        }

        $Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $SystemUserName = $env:USERNAME
        $FriendlyUserName = if ($UserLookupTable.ContainsKey($SystemUserName)) { $UserLookupTable[$SystemUserName] } else { $SystemUserName }

        $OutputFileName_All    = "PhaseX_Vergl_Mapping__$($sanitizedSupplierName)_$($pmsSupplier)__$($FriendlyUserName)__ALLE__$($Timestamp).csv"
        $OutputFileName_Errors = "PhaseX_Vergl_Mapping__$($sanitizedSupplierName)_$($pmsSupplier)__$($FriendlyUserName)__ERRORS__$($Timestamp).csv"
        
        # Lokale Pfade (für initiale Speicherung)
        $LocalOutputFilePath_All    = Join-Path -Path $LocalOutputDirectory -ChildPath $OutputFileName_All
        $LocalOutputFilePath_Errors = Join-Path -Path $LocalOutputDirectory -ChildPath $OutputFileName_Errors
        
        # SharePoint Pfade (für späteres Verschieben)
        $SharePointOutputFilePath_All    = Join-Path -Path $SharePointOutputDirectory -ChildPath $OutputFileName_All
        $SharePointOutputFilePath_Errors = Join-Path -Path $SharePointOutputDirectory -ChildPath $OutputFileName_Errors
    } else {
        throw "Lieferantennummern stimmen NICHT überein! `n - PMS-Datei: '$pmsSupplier' `n - PIM-Datei: '$pimSupplier'"
    }

    # 5. Dateien einlesen
    Write-Host "5. Lese und verarbeite Dateien... (Dies kann einige Minuten dauern)"
    
    $All_Datasets_Hashtable = New-Object 'System.Collections.Hashtable' -ArgumentList $INITIAL_CAPACITY
    $pmsSkippedCounter = 0; $pimSkippedCounter = 0
    $lineCounter = 0

    Write-Host "    - Verarbeite PMS-Datei (Streaming mit GC)..."
    
    [System.IO.File]::ReadLines($pmsFilePath, [System.Text.Encoding]::Default) | 
        Select-Object -Skip 1 | 
        ForEach-Object { 
            $lineCounter++
            
            if ($lineCounter % $GC_INTERVAL -eq 0) {
                Write-Progress -Activity "PMS-Datei einlesen" -Status "$lineCounter Zeilen verarbeitet" -PercentComplete -1
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            
            $_.TrimEnd(';') 
        } |
        ConvertFrom-Csv -Delimiter ';' -Header $PMS_Header_Expected | 
        ForEach-Object {
            $pmsRow = $_
            $rawEan = $pmsRow.SLLEAN
            if ([string]::IsNullOrWhiteSpace($rawEan)) { $pmsSkippedCounter++; return }
            $ean = $rawEan.Trim()
            if ($All_Datasets_Hashtable.ContainsKey($ean)) { 
                Write-Warning "Doppelte EAN '$ean' in PMS-Datei. Nur erster Eintrag wird berücksichtigt."
                return 
            }
            $newObject = [PSCustomObject]@{
                EAN = $ean
                'Gefunden ...' = "nur im PMS"
                'Check Summary' = ""
                'Check 1 Status' = ""
                'Check 2: Kategorie' = ""
                'Check 3: Genre' = ""
                'Check 4: Formatcode' = ""
                'Check 5: Mature Content' = ""
                PMS_SLLPAS = $pmsRow.SLLPAS
                PMS_SLLCAT = $pmsRow.SLLCAT
                PMS_SLLGNR = $pmsRow.SLLGNR
                PMS_FMBIDX = $pmsRow.FMBIDX
                PMS_MATURE = $pmsRow.MATURE
                PMS_IMPDAT = $pmsRow.IMPDAT
                PMS_CHGDAT = $pmsRow.CHGDAT
                PMS_GNXGNR = $pmsRow.GNXGNR
                PIM_Lieferant = $null
                PIM_Status = $null
                PIM_Kategorie = $null
                PIM_Genre = $null
                PIM_Formatcode = $null
                PIM_Mature = $null
                'PIM_letzter Import' = $null
                'PIM_letzte Änderung' = $null
                'PIM_letzter Status' = $null
                'ZeitDiff letzte Änderung' = ""
                'ZeitDiff Bewertung' = ""
            }
            $All_Datasets_Hashtable.Add($ean, $newObject)
        }
    
    Write-Progress -Activity "PMS-Datei einlesen" -Completed
    $pmsEanCount = $All_Datasets_Hashtable.Count
    Write-Host "    - PMS-Datei eingelesen. $($All_Datasets_Hashtable.Count) eindeutige Datensätze gefunden."
    if ($pmsSkippedCounter -gt 0) { Write-Warning "$pmsSkippedCounter Zeilen ohne EAN im PMS-File wurden übersprungen." }

    Write-Host "    - Speicher-Bereinigung..."
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()

    Write-Host "    - Verarbeite PIM-Datei (StreamReader mit GC)..."
    
    $pimSeenEans = New-Object 'System.Collections.Hashtable' -ArgumentList $INITIAL_CAPACITY
    $lineCounter = 0
    $reader = $null
    
    try {
        $reader = [System.IO.StreamReader]::new($pimFilePath, [System.Text.Encoding]::UTF8)
        $reader.ReadLine() | Out-Null
        
        while (-not $reader.EndOfStream) {
            $lineCounter++
            $pimEanCount++
            
            if ($lineCounter % $GC_INTERVAL -eq 0) {
                Write-Progress -Activity "PIM-Datei einlesen" -Status "$lineCounter Zeilen verarbeitet" -PercentComplete -1
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { $pimSkippedCounter++; continue }
            
            $line = $line.TrimEnd(';')
            $pimRow = $line | ConvertFrom-Csv -Delimiter ';' -Header $PIM_Header_Expected
            
            $rawEan = $pimRow.EAN
            if ([string]::IsNullOrWhiteSpace($rawEan)) { $pimSkippedCounter++; continue }
            $ean = $rawEan.Trim()

            if ($pimSeenEans.ContainsKey($ean)) {
                if ($All_Datasets_Hashtable.ContainsKey($ean)) {
                    $existingObject = $All_Datasets_Hashtable[$ean]
                    $existingObject.'Gefunden ...' = "mehrfach im PIM"
                    $existingObject.'Check Summary' = "nicht ok - EAN mehrfach im PIM"
                    $script:foundPimDuplicates = $true
                }
                continue
            } else { 
                $pimSeenEans.Add($ean, $true) 
            }

            if ($All_Datasets_Hashtable.ContainsKey($ean)) {
                $existingObject = $All_Datasets_Hashtable[$ean]
                $existingObject.'Gefunden ...' = "im PMS und im PIM"
                $existingObject.PIM_Lieferant = $pimRow.Lieferant
                $existingObject.PIM_Status = $pimRow.Status
                $existingObject.PIM_Kategorie = $pimRow.Kategorie
                $existingObject.PIM_Genre = $pimRow.Genre
                $existingObject.PIM_Formatcode = $pimRow.Formatcode
                $existingObject.PIM_Mature = $pimRow.Mature
                $existingObject.'PIM_letzter Import' = $pimRow.'letzter Import'
                $existingObject.'PIM_letzte Änderung' = $pimRow.'letzte Änderung'
                $existingObject.'PIM_letzter Status' = $pimRow.'letzter Status'
            } else {
                $newObject = [PSCustomObject]@{
                    EAN = $ean
                    'Gefunden ...' = "nur im PIM"
                    'Check Summary' = ""
                    'Check 1 Status' = ""
                    'Check 2: Kategorie' = ""
                    'Check 3: Genre' = ""
                    'Check 4: Formatcode' = ""
                    'Check 5: Mature Content' = ""
                    PMS_SLLPAS = $null
                    PMS_SLLCAT = $null
                    PMS_SLLGNR = $null
                    PMS_FMBIDX = $null
                    PMS_MATURE = $null
                    PMS_IMPDAT = $null
                    PMS_CHGDAT = $null
                    PMS_GNXGNR = $null
                    PIM_Lieferant = $pimRow.Lieferant
                    PIM_Status = $pimRow.Status
                    PIM_Kategorie = $pimRow.Kategorie
                    PIM_Genre = $pimRow.Genre
                    PIM_Formatcode = $pimRow.Formatcode
                    PIM_Mature = $pimRow.Mature
                    'PIM_letzter Import' = $pimRow.'letzter Import'
                    'PIM_letzte Änderung' = $pimRow.'letzte Änderung'
                    'PIM_letzter Status' = $pimRow.'letzter Status'
                    'ZeitDiff letzte Änderung' = ""
                    'ZeitDiff Bewertung' = ""
                }
                $All_Datasets_Hashtable.Add($ean, $newObject)
            }
        }
    } finally {
        if ($reader) { 
            $reader.Close()
            $reader.Dispose() 
        }
    }
    
    Write-Progress -Activity "PIM-Datei einlesen" -Completed
    Write-Host "    - PIM-Datei verarbeitet."
    if ($pimSkippedCounter -gt 0) { Write-Warning "$pimSkippedCounter Zeilen ohne EAN im PIM-File wurden übersprungen." }
    
    $pimSeenEans.Clear()
    $pimSeenEans = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Beide Files eingelesen. Dauer $($stopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Cyan
    Write-Host "Gesamtanzahl eindeutiger Datensätze: $($All_Datasets_Hashtable.Count)"

    $All_Datasets = $All_Datasets_Hashtable.Values
    $totalDatasets = $All_Datasets.Count
    $script:totalDatasetCount = $totalDatasets
    $i = 0
    
    $errorDatasetsList = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    Write-Host "6. Führe Checks durch..."
    foreach ($dataset in $All_Datasets) {
        $i++
        
        if ($i % 5000 -eq 0) {
            $percentage = [Math]::Floor(($i / $totalDatasets) * 100)
            Write-Progress -Activity "Schritt 6: Führe Checks durch" -Status "$percentage% abgeschlossen ($i von $totalDatasets EANs)" -PercentComplete $percentage
        }
        
        if ($dataset.'Check Summary' -eq 'nicht ok - EAN mehrfach im PIM') { 
            $errorDatasetsList.Add($dataset)
            $script:checkSummaryErrors++
            continue 
        }

        if ($dataset.'Gefunden ...' -eq "im PMS und im PIM") {
            # STATISTIK: Zähle matched EANs
            $script:matchedEanCount++
            
            $dataset.'ZeitDiff letzte Änderung' = Invoke-CalculateTimeDifference -Dataset $dataset
            $dataset.'Check 1 Status' = Invoke-Check1_Status -Dataset $dataset
            
            if ($dataset.'Check 1 Status' -eq 'ok') {
                $dataset.'Check 2: Kategorie' = Invoke-Check2_Kategorie -Dataset $dataset
                $dataset.'Check 3: Genre' = Invoke-Check3_Genre -Dataset $dataset
                if ($dataset.'Check 2: Kategorie' -eq 'ok - Status = passive') {
                    $dataset.'Check 4: Formatcode' = 'ok - Status = passive'
                    $dataset.'Check 5: Mature Content' = 'ok - Status = passive'
                } else {
                    $dataset.'Check 4: Formatcode' = Invoke-Check4_Formatcode -Dataset $dataset
                    $dataset.'Check 5: Mature Content' = Invoke-Check5_MatureContent -Dataset $dataset
                }
            } else {
                $dataset.'Check 2: Kategorie' = '---'
                $dataset.'Check 3: Genre' = '---'
                $dataset.'Check 4: Formatcode' = '---'
                $dataset.'Check 5: Mature Content' = '---'
            }
            
            $check1Ok = ($dataset.'Check 1 Status' -eq 'ok')
            $check2Ok = (($dataset.'Check 2: Kategorie' -eq 'ok') -or ($dataset.'Check 2: Kategorie' -eq 'ok - Status = passive'))
            $check3Ok = (($dataset.'Check 3: Genre' -eq 'ok') -or ($dataset.'Check 3: Genre' -eq 'ok - Status = passive'))
            $check4Ok = (($dataset.'Check 4: Formatcode' -eq 'ok') -or ($dataset.'Check 4: Formatcode' -eq 'ok - Status = passive'))
            $check5Ok = (($dataset.'Check 5: Mature Content' -eq 'ok') -or ($dataset.'Check 5: Mature Content' -eq 'ok - Status = passive'))
            
            if ($check1Ok -and $check2Ok -and $check3Ok -and $check4Ok -and $check5Ok) { 
                $dataset.'Check Summary' = 'ok' 
            } else { 
                $dataset.'Check Summary' = 'nicht ok'
                $errorDatasetsList.Add($dataset)
                $script:checkSummaryErrors++
                
                # Einzelne Fehler zählen
                if (-not $check1Ok) { $script:check1Errors++ }
                if (-not $check2Ok) { $script:check2Errors++ }
                if (-not $check3Ok) { $script:check3Errors++ }
                if (-not $check4Ok) { $script:check4Errors++ }
                if (-not $check5Ok) { $script:check5Errors++ }
            }
        } elseif ($dataset.'Gefunden ...' -eq "nur im PIM") {
            $dataset.'Check Summary' = 'ok - EAN nur im PIM'
        } else {
            if ($dataset.PMS_SLLPAS -eq 'passive') { 
                $dataset.'Check Summary' = 'ok - EAN fehlt im PIM - passive im PMS' 
            } else { 
                $dataset.'Check Summary' = 'nicht ok - EAN fehlt im PIM'
                $errorDatasetsList.Add($dataset)
                $script:checkSummaryErrors++
            }
        }
    }
    Write-Progress -Activity "Schritt 6: Führe Checks durch" -Completed
    Write-Host "    Checks abgeschlossen." -ForegroundColor Green

    # 7. Export vorbereiten
    Write-Host "7. Bereite Export vor..."
    
    $Error_Datasets = $errorDatasetsList.ToArray()
    $totalRowCount = $All_Datasets.Count
    $script:UseExcelExport = $false
    $fileExtension = ".csv"

    if ($totalRowCount -ge 1000000) {
        Write-Warning "Mehr als 1 Million Zeilen ($totalRowCount) gefunden."
        Write-Warning "Export erfolgt als CSV (Excel-Limit: 1.048.576 Zeilen)"
        Write-Warning "CSV-Export ist ~20x schneller (~1 Minute statt ~20 Minuten)"
    } else {
        try {
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel -ErrorAction Stop
                $script:UseExcelExport = $true
                $fileExtension = ".xlsx"
                Write-Host "    'ImportExcel'-Modul geladen. Erstelle .xlsx-Datei(en)." -ForegroundColor Green
            } else {
                Write-Warning "'ImportExcel'-Modul nicht gefunden."
                $choice = Read-Host "Möchtest du es für den Benutzer '$env:USERNAME' installieren (Internetverbindung nötig)? (j/n)"
                if ($choice -eq 'j') {
                    Write-Host "Installiere 'ImportExcel'..."
                    Install-Module ImportExcel -Scope CurrentUser -AllowClobber -Force -Confirm:$false
                    Import-Module ImportExcel -ErrorAction Stop
                    Write-Host "'ImportExcel' erfolgreich installiert und geladen." -ForegroundColor Green
                    $script:UseExcelExport = $true
                    $fileExtension = ".xlsx"
                } else {
                    Write-Warning "Installation übersprungen. Fallback auf CSV-Export."
                }
            }
        } catch {
            Write-Warning "Fehler bei 'ImportExcel': $($_.Exception.Message)"
            Write-Warning "Fallback auf CSV-Export."
            $script:UseExcelExport = $false
            $fileExtension = ".csv"
        }
    }

    $LocalOutputFilePath_All    = $LocalOutputFilePath_All.Replace(".csv", $fileExtension)
    $LocalOutputFilePath_Errors = $LocalOutputFilePath_Errors.Replace(".csv", $fileExtension)
    $SharePointOutputFilePath_All    = $SharePointOutputFilePath_All.Replace(".csv", $fileExtension)
    $SharePointOutputFilePath_Errors = $SharePointOutputFilePath_Errors.Replace(".csv", $fileExtension)
    $OutputFileName_All    = $OutputFileName_All.Replace(".csv", $fileExtension)
    $OutputFileName_Errors = $OutputFileName_Errors.Replace(".csv", $fileExtension)

    Write-Host ""
    Write-Host "8. Schreibe Ergebnisdateien (lokal)..."
    Write-Host "    Ausgabe-Datei (alle): '$OutputFileName_All'" -ForegroundColor Cyan
    Write-Host "    Ausgabe-Datei (Fehler): '$OutputFileName_Errors'" -ForegroundColor Cyan
    Write-Host ""

    # EXPORT MIT FEHLERBEHANDLUNG UND FALLBACK-STRATEGIE
    if ($script:UseExcelExport) {
        # Excel-Export
        $headerSummary = [PSCustomObject][ordered]@{
            'EAN' = $pmsSupplier
            'Gefunden ...' = "Script $($ScriptVersion.Replace('Mapping','M'))"
            'Check Summary' = $script:checkSummaryErrors
            'Check 1 Status' = $script:check1Errors
            'Check 2: Kategorie' = $script:check2Errors
            'Check 3: Genre' = $script:check3Errors
            'Check 4: Formatcode' = $script:check4Errors
            'Check 5: Mature Content' = $script:check5Errors
            'PMS_SLLPAS' = ' '
            'PMS_SLLCAT' = ' '
            'PMS_SLLGNR' = ' '
            'PMS_FMBIDX' = ' '
            'PMS_MATURE' = ' '
            'PMS_IMPDAT' = ' '
            'PMS_CHGDAT' = ' '
            'PMS_GNXGNR' = ' '
            'PIM_Lieferant' = ' '
            'PIM_Status' = ' '
            'PIM_Kategorie' = ' '
            'PIM_Genre' = ' '
            'PIM_Formatcode' = ' '
            'PIM_Mature' = ' '
            'PIM_letzter Import' = ' '
            'PIM_letzte Änderung' = ' '
            'PIM_letzter Status' = ' '
            'ZeitDiff letzte Änderung' = ' '
            'ZeitDiff Bewertung' = ' '
        }
        
        try {
            $headerSummary | Select-Object * -ExcludeProperty 'PIM_Lieferant' | Export-Excel -Path $LocalOutputFilePath_All -WorksheetName "Vergleich" -ClearSheet -StartRow 1 -NoHeader
            $All_Datasets | Export-Excel -Path $LocalOutputFilePath_All -WorksheetName "Vergleich" -StartRow 2 -AutoFilter -AutoSize -ExcludeProperty 'PIM_Lieferant'
            Apply-WorksheetFormatting -Path $LocalOutputFilePath_All -SheetNames @("Vergleich")
            $createdOutputFiles.Add($OutputFileName_All)
            Write-Host "    ✓ ALLE-Datei (Excel) erfolgreich erstellt." -ForegroundColor Green
        } catch {
            Write-Warning "Excel-Export fehlgeschlagen: $($_.Exception.Message)"
        }

        if ($Error_Datasets.Count -gt 0) {
            try {
                $headerSummary | Select-Object * -ExcludeProperty 'PIM_Lieferant' | Export-Excel -Path $LocalOutputFilePath_Errors -WorksheetName "Fehler" -ClearSheet -StartRow 1 -NoHeader
                $Error_Datasets | Export-Excel -Path $LocalOutputFilePath_Errors -WorksheetName "Fehler" -StartRow 2 -AutoFilter -AutoSize -ExcludeProperty 'PIM_Lieferant'
                Apply-WorksheetFormatting -Path $LocalOutputFilePath_Errors -SheetNames @("Fehler")
                $createdOutputFiles.Add($OutputFileName_Errors)
                Write-Host "    ✓ ERRORS-Datei (Excel) erfolgreich erstellt." -ForegroundColor Green
            } catch {
                Write-Warning "Excel-Export der ERRORS-Datei fehlgeschlagen: $($_.Exception.Message)"
            }
        }
    } else {
        # CSV-Export mit Fehlerbehandlung und Streaming-Fallback
        
        # ALLE-Datei exportieren
        try {
            Write-Host "    - Erstelle ALLE-Datei (Streaming-Export)..."
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            $success = Export-CsvStreaming -Data $All_Datasets -Path $LocalOutputFilePath_All -ActivityName "CSV-Export: ALLE-Datei" -ExcludeProperties @('PIM_Lieferant') -BatchSize $EXPORT_BATCH_SIZE
            
            if ($success) {
                $createdOutputFiles.Add($OutputFileName_All)
                Write-Host "    ✓ ALLE-Datei erfolgreich erstellt." -ForegroundColor Green
            } else {
                Write-Warning "Streaming-Export der ALLE-Datei fehlgeschlagen."
            }
        } catch {
            Write-Warning "ALLE-Datei konnte nicht erstellt werden: $($_.Exception.Message)"
            Write-Warning "Versuche nur ERROR-Datei zu erstellen..."
        }
        
        # ERROR-Datei exportieren (sollte immer funktionieren, da kleiner)
        if ($Error_Datasets.Count -gt 0) {
            try {
                Write-Host "    - Erstelle ERRORS-Datei (Streaming-Export)..."
                [System.GC]::Collect()
                
                $success = Export-CsvStreaming -Data $Error_Datasets -Path $LocalOutputFilePath_Errors -ActivityName "CSV-Export: ERRORS-Datei" -ExcludeProperties @('PIM_Lieferant') -BatchSize $EXPORT_BATCH_SIZE
                
                if ($success) {
                    $createdOutputFiles.Add($OutputFileName_Errors)
                    Write-Host "    ✓ ERRORS-Datei erfolgreich erstellt." -ForegroundColor Green
                } else {
                    Write-Warning "Streaming-Export der ERRORS-Datei fehlgeschlagen."
                }
            } catch {
                Write-Warning "ERRORS-Datei konnte nicht erstellt werden: $($_.Exception.Message)"
            }
        }
    }

    Write-Host ""
    if ($createdOutputFiles.Count -gt 0) {
        Write-Host "      Export abgeschlossen." -ForegroundColor Green
    } else {
        Write-Warning "      WARNUNG: Keine Dateien konnten exportiert werden!"
        Write-Warning "      Die Analyse ist jedoch vollständig - siehe Zusammenfassung unten."
    }

    # MEMORY-OPTIMIERUNG: Gebe $All_Datasets frei nach Export
    Write-Host "    - Speicher-Bereinigung nach Export..."
    $All_Datasets = $null
    $All_Datasets_Hashtable.Clear()
    $All_Datasets_Hashtable = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()

    # 9. Verschiebe Files zu SharePoint (nur wenn erfolgreich erstellt)
    if ($createdOutputFiles.Count -gt 0) {
        Write-Host ""
        Write-Host "9. Verschiebe Files zu SharePoint..."
        
        if ($createdOutputFiles.Contains($OutputFileName_All) -and (Test-Path $LocalOutputFilePath_All)) {
            try {
                Copy-Item -Path $LocalOutputFilePath_All -Destination $SharePointOutputFilePath_All -Force
                Write-Host "    - '$OutputFileName_All' nach SharePoint verschoben." -ForegroundColor Green
            } catch {
                Write-Warning "Fehler beim Verschieben der ALLE-Datei zu SharePoint: $($_.Exception.Message)"
            }
        }
        
        if ($createdOutputFiles.Contains($OutputFileName_Errors) -and (Test-Path $LocalOutputFilePath_Errors)) {
            try {
                Copy-Item -Path $LocalOutputFilePath_Errors -Destination $SharePointOutputFilePath_Errors -Force
                Write-Host "    - '$OutputFileName_Errors' nach SharePoint verschoben." -ForegroundColor Green
            } catch {
                Write-Warning "Fehler beim Verschieben der ERRORS-Datei zu SharePoint: $($_.Exception.Message)"
            }
        }
    }

    Write-Host ""
    Write-Host "--------------------------------------------------------" -ForegroundColor Green
    Write-Host "Verarbeitung abgeschlossen."

    $scriptSuccessfullyCompleted = $true
    #endregion
}
catch {
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    Write-Host "EIN FEHLER IST AUFGETRETEN:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    Write-Host ""
    Write-Host "HINWEIS: Die Datenanalyse wurde möglicherweise durchgeführt." -ForegroundColor Yellow
    Write-Host "         Prüfe die Zusammenfassung unten für Ergebnisse." -ForegroundColor Yellow
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Skript-Fehler", "OK", "Error")
}
finally {
    if ($stopwatch.IsRunning) { $stopwatch.Stop() }
    # ZUSAMMENFASSUNG WIRD IMMER ANGEZEIGT (verwendet vorberechnete Zähler - kein Memory-Overhead)
    if ($scriptSuccessfullyCompleted -or ($script:totalDatasetCount -gt 0)) { 
        Pause-Ende 
    } else { 
        Write-Host "`nDrücke ENTER um das Fenster zu schliessen."
        Read-Host 
    }
}
