@echo off
:: Kopiert invoice_tools.exe, _internal\ und alle Konfigurationsdateien
:: ins Produktionsverzeichnis eine Ebene höher.
:: Verwendung: publish.bat

cd /d "%~dp0"
set PROD=..\

echo Kopiere ins Produktionsverzeichnis: %PROD%
echo.

:: _internal\ synchronisieren (nur geänderte Dateien, veraltete entfernen)
robocopy "_internal" "%PROD%_internal" /E /PURGE /NFL /NDL /NJH /NJS
if %errorlevel% geq 8 echo FEHLER beim Kopieren von _internal & exit /b 1

:: EXE
copy /Y "invoice_tools.exe"               "%PROD%invoice_tools.exe" >nul

:: XML-Konfigurationen
copy /Y "invoice_extractor_config.xml"    "%PROD%invoice_extractor_config.xml" >nul
copy /Y "invoice_extractor_config_RE.xml" "%PROD%invoice_extractor_config_RE.xml" >nul
copy /Y "invoice_extractor_config_RA.xml" "%PROD%invoice_extractor_config_RA.xml" >nul
copy /Y "invoice_inbox_config.xml"        "%PROD%invoice_inbox_config.xml" >nul
copy /Y "invoice_tools_api_config.xml"    "%PROD%invoice_tools_api_config.xml" >nul
if exist "invoice_tools_mailto_config.xml" copy /Y "invoice_tools_mailto_config.xml" "%PROD%invoice_tools_mailto_config.xml" >nul

:: Batch-Skripte
copy /Y "rechnungseingang_in.bat"         "%PROD%rechnungseingang_in.bat" >nul
copy /Y "rechnungseingang_ex.bat"         "%PROD%rechnungseingang_ex.bat" >nul
copy /Y "rechnungsausgang_ex.bat"         "%PROD%rechnungsausgang_ex.bat" >nul
copy /Y "Rechnungen_versenden.bat"        "%PROD%Rechnungen_versenden.bat" >nul

:: Dokumentation
copy /Y "invoice_tools_doku.pdf"          "%PROD%invoice_tools_doku.pdf" >nul
if exist "invoice_tools_doku.html"           copy /Y "invoice_tools_doku.html"           "%PROD%invoice_tools_doku.html" >nul
copy /Y "rechnungseingang_in_doku.pdf"       "%PROD%rechnungseingang_in_doku.pdf" >nul
copy /Y "rechnungseingang_ex_doku.pdf"       "%PROD%rechnungseingang_ex_doku.pdf" >nul
copy /Y "rechnungsausgang_ex_doku.pdf"       "%PROD%rechnungsausgang_ex_doku.pdf" >nul
if exist "Rechnungen_versenden_doku.pdf"     copy /Y "Rechnungen_versenden_doku.pdf"   "%PROD%Rechnungen_versenden_doku.pdf" >nul

echo Fertig.
