@echo off
REM EC ISL - Generador del estado de cuenta (v1.1)
REM Doble clic para abrir la ventana, o arrastrar el Excel de la consulta.
cd /d "%~dp0"
if "%~1"=="" (
  python "%~dp0ec_isl_v1.1.py"
) else (
  python "%~dp0ec_isl_v1.1.py" "%~1"
)
echo.
pause
