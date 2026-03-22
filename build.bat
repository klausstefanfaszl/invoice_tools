@echo off
:: Baut invoice_tools.exe und kopiert sie ins src-Verzeichnis zum Testen.
:: Ausgabe: invoice_tools.exe + _internal\ im aktuellen Verzeichnis
:: Verwendung: build.bat

cd /d "%~dp0"
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
