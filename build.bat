@echo off
:: Baut invoice_tools.exe und kopiert sie ins src-Verzeichnis zum Testen.
:: Ausgabe: invoice_tools.exe + _internal\ im aktuellen Verzeichnis
:: Verwendung: build.bat [--force]
:: Mit --force wird immer gebaut, auch wenn die EXE aktuell ist.

cd /d "%~dp0"

set EXE=invoice_tools.exe
set FORCE=0
if /i "%~1"=="--force" set FORCE=1

if %FORCE%==1 goto :build

:: Prüfen ob EXE existiert
if not exist "%EXE%" (
    echo EXE nicht gefunden - Build wird gestartet.
    goto :build
)

:: Änderungsdaten der Quelldateien gegen EXE prüfen (xcopy /D kopiert nur neuere Dateien)
:: Trick: xcopy /D /L zeigt, was kopiert werden würde - wenn die Liste leer ist, ist alles aktuell.
set NEWER=0
for %%F in (invoice_tools.py invoice_extractor.py inbox_processor.py mailto_sender.py) do (
    xcopy /D /L /Y "%%F" "%EXE%" 2>nul | findstr /i "%%F" >nul && set NEWER=1
)

if %NEWER%==0 (
    echo Kein Build noetig - alle Quelldateien aelter als %EXE%.
    pause
    exit /b 0
)

:build
echo Baue invoice_tools.exe ...
py -m PyInstaller invoice_tools.spec --noconfirm
if errorlevel 1 (
    echo FEHLER beim Build.
    exit /b 1
)

echo.
echo Kopiere ins src-Verzeichnis ...
xcopy /E /Y /I "dist\invoice_tools\_internal" "_internal\" >nul
copy /Y "dist\invoice_tools\invoice_tools.exe" "invoice_tools.exe" >nul

echo Fertig: invoice_tools.exe + _internal\ im src-Verzeichnis.
