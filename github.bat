@echo off
:: Speichert den aktuellen Stand in GitHub (Standard) oder stellt ihn wieder her.
:: Verwendung: github.bat          -> speichern (push)
::             github.bat restore  -> wiederherstellen (pull)

cd /d "%~dp0"

if /i "%1"=="restore" goto restore

:store
echo Speichere in GitHub ...
git add -A
git diff --cached --quiet
if not errorlevel 1 (
    echo Keine Änderungen vorhanden - nichts zu speichern.
    goto end
)
git commit -m "Aktualisierung %date% %time%"
git pull --rebase origin main
if errorlevel 1 (
    echo FEHLER beim Pull/Rebase vor dem Push.
    exit /b 1
)
git push -u origin main
if errorlevel 1 (
    echo FEHLER beim Push.
    exit /b 1
)
echo Fertig gespeichert.
goto end

:restore
echo Wiederherstellen von GitHub ...
git pull origin main
if errorlevel 1 (
    echo FEHLER beim Pull.
    exit /b 1
)
echo Fertig wiederhergestellt.

:end
