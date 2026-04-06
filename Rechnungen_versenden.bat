@echo off
setlocal enabledelayedexpansion

:: ============================================================
::  Konfiguration — hier bei Bedarf anpassen
:: ============================================================
set CONFIG=%~dp0invoice_tools_mailto_config.xml
set ADRESS_MODUS=2

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
::  Parameter einlesen
::  Aufruf:  Rechnungen_versenden.bat [min-rg-nr] [max-rg-nr] [Optionen]
::           Rechnungen_versenden.bat [min-rg-nr] --rg-datum TT.MM.JJJJ
::
::  Positionale Zahlen werden als min-rg-nr bzw. max-rg-nr interpretiert.
::  Alle anderen Argumente (z.B. --rg-datum) werden direkt durchgereicht.
::
::  Ohne min-rg-nr: interaktive Abfrage.
:: ============================================================

set MIN_RGNR=
set MAX_RGNR=
set EXTRA=

:parse
if "%~1"=="" goto :parse_done

:: Prüfen ob aktuelles Argument eine reine Zahl ist
echo %~1| findstr /r "^[0-9][0-9]*$" >nul 2>&1
if !errorlevel!==0 (
    :: Zahl: erst min-rg-nr, dann max-rg-nr belegen
    if "!MIN_RGNR!"=="" (
        set MIN_RGNR=%~1
    ) else if "!MAX_RGNR!"=="" (
        set MAX_RGNR=%~1
    ) else (
        echo.
        echo  Fehler: Zu viele numerische Argumente: %~1
        goto :usage
    )
) else (
    :: Kein Zahl -> als Extra-Option durchreichen
    set EXTRA=!EXTRA! %1
)
shift
goto :parse

:parse_done

:: Keine Nummer übergeben -> interaktiv abfragen
if "!MIN_RGNR!"=="" goto :abfrage
goto :start

:usage
echo.
echo  Verwendung: %~nx0 [min-rg-nr] [max-rg-nr] [--rg-datum TT.MM.JJJJ]
echo  Beispiele:  %~nx0 4700
echo              %~nx0 4700 4750
echo              %~nx0 4700 --rg-datum 15.04.2026
echo              %~nx0 --rg-datum 15.04
echo.
exit /b 1

:abfrage
echo.
set /p MIN_RGNR= Ab welcher Rechnungsnummer versenden?
echo.
if "!MIN_RGNR!"=="" (
    echo  Fehler: Keine Rechnungsnummer eingegeben.
    echo.
    exit /b 1
)
echo !MIN_RGNR!| findstr /r "^[0-9][0-9]*$" >nul 2>&1
if !errorlevel!==1 (
    echo  Fehler: Ungueltige Eingabe: !MIN_RGNR!
    echo.
    exit /b 1
)

:: ============================================================
::  Verarbeitung starten
:: ============================================================
:start
set MAX_ARG=
if not "!MAX_RGNR!"=="" set MAX_ARG=--max-rg-nr !MAX_RGNR!

set STARTZEIT=%time%
echo.
echo  Ausgangsrechnungen per E-Mail versenden
echo  -----------------------------------------------
echo  Adressmodus : %ADRESS_MODUS% ^(Datenbank^)
echo  Von RE-Nr   : !MIN_RGNR!
if not "!MAX_RGNR!"=="" echo  Bis RE-Nr   : !MAX_RGNR!
if not "!EXTRA!"==""    echo  Optionen    :!EXTRA!
echo  Config      : %CONFIG%
echo  Laufzeit    : !RUNNER!
echo  Start       : %STARTZEIT%
echo  -----------------------------------------------
echo.

!RUNNER! mailto --config "%CONFIG%" --adress-modus %ADRESS_MODUS% --min-rg-nr !MIN_RGNR! !MAX_ARG! !EXTRA!
set EXITCODE=%errorlevel%

echo.
if %EXITCODE%==0 (
    echo  Erfolgreich abgeschlossen.
) else (
    echo  Fehler beim Versand ^(Exit-Code: %EXITCODE%^).
)
echo  Start : %STARTZEIT%
echo  Ende  : %time%
echo.

exit /b %EXITCODE%
