@echo off
setlocal enabledelayedexpansion

:: ============================================================
::  Konfiguration — hier bei Bedarf anpassen
:: ============================================================
set CONFIG=%~dp0invoice_inbox_config.xml

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
::  Aufruf:  Rechnungseingang_inbox.bat [modus]
::  Modus:   dry    – Simulation, keine Dateien gespeichert
::           unread – nur ungelesene Mails (Standard)
::           all    – alle Mails verarbeiten
::           archiv – wie unread + erfolgreich verarbeitete Mails in Archiv-Ordner verschieben
:: ============================================================
set MODUS=archiv
if not "%~1"=="" set MODUS=%~1

if /i "%MODUS%"=="dry"    goto :start
if /i "%MODUS%"=="unread" goto :start
if /i "%MODUS%"=="all"    goto :start
if /i "%MODUS%"=="archiv" goto :start

echo.
echo  Unbekannter Modus: %MODUS%
echo.
echo  Verwendung: %~nx0 [modus]
echo  Modi:       dry    ^(Simulation^)
echo              unread ^(nur ungelesene Mails, Standard^)
echo              all    ^(alle Mails^)
echo              archiv ^(wie unread, verschiebt verarbeitete Mails in Archiv-Ordner^)
echo.
exit /b 1

:start
set STARTZEIT=%time%
echo.
echo  Eingangsrechnungen verarbeiten
echo  -----------------------------------------------
echo  Modus   : %MODUS%
echo  Config  : %CONFIG%
echo  Laufzeit: !RUNNER!
echo  Start   : %STARTZEIT%
echo  -----------------------------------------------
echo.

!RUNNER! inbox --modus %MODUS% --bzv export --export-excel --config "%CONFIG%" --bdir ..
set EXITCODE=%errorlevel%

echo.
if %EXITCODE%==0 (
    echo  Erfolgreich abgeschlossen.
) else (
    echo  Fehler bei der Verarbeitung ^(Exit-Code: %EXITCODE%^).
)
echo  Start : %STARTZEIT%
echo  Ende  : %time%
echo.

exit /b %EXITCODE%
