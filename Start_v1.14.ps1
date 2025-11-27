<#
.SYNOPSIS
    Start-Script für PMS/PIM Vergleich - Prüft Modulversionen und startet Hauptlogik

.DESCRIPTION
    Dieses Script ist der Einstiegspunkt für den PMS/PIM Vergleich.
    Es prüft vor dem Start, ob alle benötigten Module in der korrekten Version vorhanden sind.

.NOTES
    File:           Start_v1.14.ps1
    Version:        1.14
    Änderungshistorie:
        1.14 - Erwartet main_v1.8 (Produktiv-Version ohne DEBUG)
        1.13 - Erwartet main_v1.7 (Fix: $PSScriptRoot statt $MyInvocation)
        1.10 - Erwartet main_v1.4 (Input-Pfad dynamisch: übergeordneter Ordner + PIM\PhaseX_Berechnung)
        1.9 - Erwartet functions-checks_v1.5 (Check 7: 2-stelliges Jahr)
        1.8 - Erwartet config_v1.1 (ScriptVersion-Zeile entfernt)
        1.7 - Fix: ScriptVersion als GLOBAL Variable (nicht script:)
            - Damit main.ps1 die Version sehen kann
            - Erwartet functions-checks_v1.4 (Check 14 Fix)
        1.6 - ScriptVersion wird von Start gesetzt und an main uebergeben
            - Format: "Berechnung_V<StartVersion>"
            - main_v1.1 erforderlich
        1.5 - Erwartet functions-checks_v1.3 (Check 14: L-Prio/PrioEP Korrelation)
        1.4 - Erwartet functions-checks_v1.2 (Checks 9,10,11,12: PMS 0 = PIM leer)
        1.3 - Erwartet functions-checks_v1.1 (Check 10 mit relativer Toleranz)
        1.2 - Fix: Module werden global geladen (nicht nur in Prueffunktion)
        1.1 - Bootstrap-Fix: Fenster bleibt offen und zeigt Output
        1.0 - Initiale Version
    
    Benötigte Module:
        - config_v1.1.ps1            (ModuleVersion_Config = 1.1)
        - functions-excel_v1.0.ps1   (ModuleVersion_Excel = 1.0)
        - functions-dialogs_v1.0.ps1 (ModuleVersion_Dialogs = 1.0)
        - functions-helpers_v1.0.ps1 (ModuleVersion_Helpers = 1.0)
        - functions-checks_v1.5.ps1  (ModuleVersion_Checks = 1.5)
        - main_v1.8.ps1              (ModuleVersion_Main = 1.8)
#>

# =====================================================================
# START-SCRIPT VERSION (wird fuer gesamtes Script verwendet)
# =====================================================================
$script:StartVersion = "1.14"

# Diese Variable wird von main.ps1 verwendet fuer Output und Fenster
# GLOBAL damit sie auch in main.ps1 sichtbar ist!
$global:ScriptVersion = "Berechnung_V$($script:StartVersion)"

# =====================================================================
# BENÖTIGTE MODUL-VERSIONEN (Reihenfolge wichtig: config zuerst, main zuletzt)
# =====================================================================
$script:RequiredModules = [ordered]@{
    'config'           = @{ File = 'config_v1.1.ps1';            Variable = 'ModuleVersion_Config';  Version = '1.1' }
    'functions-excel'  = @{ File = 'functions-excel_v1.0.ps1';   Variable = 'ModuleVersion_Excel';   Version = '1.0' }
    'functions-dialogs'= @{ File = 'functions-dialogs_v1.0.ps1'; Variable = 'ModuleVersion_Dialogs'; Version = '1.0' }
    'functions-helpers'= @{ File = 'functions-helpers_v1.0.ps1'; Variable = 'ModuleVersion_Helpers'; Version = '1.0' }
    'functions-checks' = @{ File = 'functions-checks_v1.5.ps1';  Variable = 'ModuleVersion_Checks';  Version = '1.5' }
    'main'             = @{ File = 'main_v1.8.ps1';              Variable = 'ModuleVersion_Main';    Version = '1.8' }
}

# =====================================================================
# BOOTSTRAP (wie Original V1.103)
# =====================================================================
if (-not $env:PS_KEEP_NOEXIT) {
    try {
        $env:PS_KEEP_NOEXIT = '1'
        $scriptPath = $MyInvocation.MyCommand.Definition
        if (-not (Test-Path -LiteralPath $scriptPath)) { throw "Scriptpfad ungueltig" }
        $quoted = '"' + $scriptPath.Replace('"', '""') + '"'
        # /c statt /k damit Fenster nach ENTER automatisch schliesst
        $arguments = '/c powershell.exe -NoLogo -ExecutionPolicy Bypass -File ' + $quoted + ' & pause'
        Start-Process -FilePath 'cmd.exe' -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $scriptPath)
    } catch {
        $global:__ForcePauseAtEnd = $true
    }
    return
}

$global:__ForcePauseAtEnd = $false
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Windows.Forms

# =====================================================================
# HAUPTTEIL
# =====================================================================
$scriptSuccessfullyCompleted = $false
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

try {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "PMS/PIM Vergleich - Start" -ForegroundColor Cyan
    Write-Host "Version: $($global:ScriptVersion)" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    Write-Host ""
    Write-Host "Pruefe Modul-Versionen..." -ForegroundColor Cyan
    Write-Host ""
    
    $versionErrors = @()
    
    # Module laden und Versionen pruefen
    foreach ($moduleName in $script:RequiredModules.Keys) {
        $moduleInfo = $script:RequiredModules[$moduleName]
        $filePath = Join-Path $ScriptDir $moduleInfo.File
        $requiredVersion = $moduleInfo.Version
        $versionVariable = $moduleInfo.Variable
        
        # Pruefe ob Datei existiert
        if (-not (Test-Path $filePath)) {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' nicht gefunden!"
            Write-Host "  [X] $($moduleInfo.File) - NICHT GEFUNDEN" -ForegroundColor Red
            continue
        }
        
        # Lade Modul (GLOBAL, nicht in einer Funktion!)
        try {
            . $filePath
        } catch {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' konnte nicht geladen werden: $($_.Exception.Message)"
            Write-Host "  [X] $($moduleInfo.File) - LADEFEHLER: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
        
        # Pruefe Version
        $actualVersion = Get-Variable -Name $versionVariable -ValueOnly -Scope Script -ErrorAction SilentlyContinue
        
        if (-not $actualVersion) {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' enthaelt keine Versionsvariable '$versionVariable'!"
            Write-Host "  [X] $($moduleInfo.File) - KEINE VERSION GEFUNDEN" -ForegroundColor Red
            continue
        }
        
        if ($actualVersion -ne $requiredVersion) {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' hat Version $actualVersion, benoetigt wird $requiredVersion!"
            Write-Host "  [X] $($moduleInfo.File) - Version $actualVersion (benoetigt: $requiredVersion)" -ForegroundColor Red
            continue
        }
        
        Write-Host "  [OK] $($moduleInfo.File) - Version $actualVersion" -ForegroundColor Green
    }
    
    Write-Host ""
    
    # Pruefe ob Fehler aufgetreten sind
    if ($versionErrors.Count -gt 0) {
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "VERSIONSPRUEFUNG FEHLGESCHLAGEN!" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host ""
        foreach ($err in $versionErrors) {
            Write-Host $err -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "Bitte stelle sicher, dass alle Module in der korrekten Version vorhanden sind." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Benoetigte Module (siehe .NOTES in diesem Script):" -ForegroundColor Cyan
        foreach ($moduleName in $script:RequiredModules.Keys) {
            $moduleInfo = $script:RequiredModules[$moduleName]
            Write-Host "  - $($moduleInfo.File) (Version $($moduleInfo.Version))" -ForegroundColor White
        }
        Write-Host ""
        Write-Host "Script wird beendet wegen Versionsfehlern." -ForegroundColor Red
        # Fenster bleibt offen durch finally-Block
    } else {
        Write-Host "Alle Module erfolgreich geladen und verifiziert." -ForegroundColor Green
        Write-Host ""
        
        # Hauptlogik starten
        Write-Host "Starte Hauptlogik..." -ForegroundColor Cyan
        Write-Host ""
        
        Invoke-MainLogic
        $scriptSuccessfullyCompleted = $true
    }
}
catch {
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    Write-Host "EIN KRITISCHER FEHLER IST AUFGETRETEN:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Kritischer Fehler", "OK", "Error") | Out-Null
}
finally {
    # Wenn Hauptlogik NICHT erfolgreich war (oder gar nicht gestartet), hier Pause machen
    # Bei erfolgreicher Hauptlogik macht main.ps1 selbst die Pause via Pause-Ende
    if (-not $scriptSuccessfullyCompleted) {
        Write-Host ""
        if ($global:__ForcePauseAtEnd) { 
            Write-Host "Hinweis: Relaunch mit eigenem Fenster war nicht moeglich." -ForegroundColor Yellow 
        }
        Write-Host "Druecke ENTER um das Fenster zu schliessen." -ForegroundColor White
        Read-Host
    }
}
