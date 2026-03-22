@echo off
:: Kopiert invoice_tools.exe, _internal\ und alle Konfigurationsdateien
:: ins Produktionsverzeichnis eine Ebene höher.
:: Verwendung: publish.bat

cd /d "%~dp0"
set PROD=..\

echo Kopiere ins Produktionsverzeichnis: %PROD%

xcopy /E /Y /I "_internal" "%PROD%_internal\" >nul
copy /Y "invoice_tools.exe"               "%PROD%invoice_tools.exe" >nul

copy /Y "invoice_extractor_config.xml"    "%PROD%invoice_extractor_config.xml" >nul
copy /Y "invoice_extractor_config_RE.xml" "%PROD%invoice_extractor_config_RE.xml" >nul
copy /Y "invoice_extractor_config_RA.xml" "%PROD%invoice_extractor_config_RA.xml" >nul
copy /Y "invoice_inbox_config.xml"        "%PROD%invoice_inbox_config.xml" >nul
copy /Y "invoice_tools_api_config.xml"    "%PROD%invoice_tools_api_config.xml" >nul

copy /Y "rechnungseingang_in.bat"         "%PROD%rechnungseingang_in.bat" >nul
copy /Y "rechnungseingang_ex.bat"         "%PROD%rechnungseingang_ex.bat" >nul
copy /Y "rechnungsausgang_ex.bat"         "%PROD%rechnungsausgang_ex.bat" >nul

echo Fertig.
