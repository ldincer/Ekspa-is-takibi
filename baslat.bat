@echo off
REM Ekspa Is Takibi - Windows baslatici
cd /d "%~dp0"
where py >nul 2>nul && (py -3 app.py %*) || (python app.py %*)
pause
