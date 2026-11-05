@echo off
REM EC ISL - Generador del estado de cuenta (v1.1)
REM Doble clic para abrir la ventana, o arrastrar el Excel de la consulta.
if not defined EC_ISL_MIN (
  set EC_ISL_MIN=1
  start "EC ISL" /min "%~f0" %*
  exit /b
)
cd /d "%~dp0"
if "%~1"=="" (
  python "%~dp0app\iniciar.py"
) else (
  python "%~dp0app\iniciar.py" "%~1"
)
echo.
pause
