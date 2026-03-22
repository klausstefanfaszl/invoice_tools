@echo off
setlocal enabledelayedexpansion

:: ============================================================
::  Konfiguration — hier bei Bedarf anpassen
:: ============================================================
set BASE_DIR=F:\Dokumente\UHDE intern\Buchhaltung
set FORMAT=pdf
set CONFIG=%~dp0invoice_extractor_config_RA.xml

:: ── Ausführungsmodus: Python bevorzugt wenn aktueller als EXE ──────────────
set _PY=%~dp0invoice_tools.py
if not exist "%_PY%" set _PY=%~dp0src\invoice_tools.py
set _EXE=%~dp0invoice_tools.exe
set _USE_PY=0
where py >nul 2>&1
if !errorlevel!==0 (
    if exist "!_PY!" (
        if exist "!_EXE!" (
            for /f %%R in ('powershell -NoProfile -Command "if ((Get-Item '!_PY!').LastWriteTime -gt (Get-Item '!_EXE!').LastWriteTime) {'1'} else {'0'}"') do set _USE_PY=%%R
        ) else (
            set _USE_PY=1
        )
    )
)
if "!_USE_PY!"=="1" (set RUNNER=py "!_PY!") else (set RUNNER="!_EXE!")

:: ============================================================
::  Parameter prüfen
::  Aufruf:  rechnungsausgang_ex.bat JJJJ/MM
::      oder rechnungsausgang_ex.bat MM         (aktuelles Jahr)
:: ============================================================
if "%~1"=="" (
    echo.
    echo  Verwendung: %~nx0 JJJJ/MM
    echo          oder %~nx0 MM        ^(aktuelles Jahr wird verwendet^)
    echo.
    echo  Beispiele:  %~nx0 2026/01
    echo              %~nx0 02
    echo.
    exit /b 1
)

:: Schrägstrich zu Backslash normalisieren
set PARAM=%~1
set PARAM=%PARAM:/=\%

:: Format erkennen: enthält Backslash → JJJJ\MM, sonst nur MM
echo %PARAM%| findstr /c:"\" >nul
if %errorlevel%==0 (
    :: Format JJJJ\MM — Jahr und Monat trennen
    for /f "tokens=1,2 delims=\" %%a in ("%PARAM%") do (
        set YEAR=%%a
        set MONTH=%%b
    )
) else (
    :: Nur MM — aktuelles Jahr ermitteln
    set MONTH=%PARAM%
    for /f %%a in ('powershell -NoProfile -Command "Get-Date -Format yyyy"') do set YEAR=%%a
)

:: Basisvalidierung
if "!YEAR!"=="" (
    echo.
    echo  Fehler: Jahreszahl konnte nicht ermittelt werden.
    echo.
    exit /b 1
)
if "!MONTH!"=="" (
    echo.
    echo  Fehler: Monat konnte nicht ermittelt werden. Erwartet: MM oder JJJJ/MM
    echo.
    exit /b 1
)

:: ============================================================
::  Pfade aufbauen
:: ============================================================
set INPUT_DIR=%BASE_DIR%\%YEAR%\%MONTH%\Ausgangsrechnungen
set OUTPUT_DIR=%BASE_DIR%\%YEAR%\%MONTH%
set OUTPUT_FILE=%OUTPUT_DIR%\Rechnungsausgang_%MONTH%.%FORMAT%

:: Eingabeverzeichnis prüfen
if not exist "%INPUT_DIR%\" (
    echo.
    echo  Fehler: Verzeichnis nicht gefunden:
    echo  %INPUT_DIR%
    echo.
    exit /b 1
)

:: Ausgabeverzeichnis prüfen
if not exist "%OUTPUT_DIR%\" (
    echo.
    echo  Fehler: Ausgabeverzeichnis nicht gefunden:
    echo  %OUTPUT_DIR%
    echo.
    exit /b 1
)

:: PDF-Dateien zählen
set COUNT=0
for %%f in ("%INPUT_DIR%\*.pdf") do set /a COUNT+=1

if %COUNT%==0 (
    echo.
    echo  Keine PDF-Dateien gefunden in:
    echo  %INPUT_DIR%
    echo.
    exit /b 1
)

:: ============================================================
::  Verarbeitung starten
:: ============================================================
echo.
echo  Rechnungsausgang %MONTH%/%YEAR%
echo  -----------------------------------------------
echo  Eingabe : %INPUT_DIR%
echo  Dateien : %COUNT% PDF(s)
echo  Format  : %FORMAT%
echo  Ausgabe : %OUTPUT_FILE%
echo  Laufzeit: !RUNNER!
echo  -----------------------------------------------
echo.

!RUNNER! extractor -c "%CONFIG%" -f %FORMAT% -o "%OUTPUT_FILE%" "%INPUT_DIR%\*.pdf"
set EXITCODE=%errorlevel%

echo.
if %EXITCODE%==0 (
    echo  Erfolgreich abgeschlossen.
    echo  Ausgabedatei: %OUTPUT_FILE%
) else (
    echo  Fehler bei der Verarbeitung ^(Exit-Code: %EXITCODE%^).
)
echo.

exit /b %EXITCODE%
