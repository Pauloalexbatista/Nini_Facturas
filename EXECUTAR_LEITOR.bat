@echo off
echo.
echo ==========================================
echo    A INICIAR LEITOR DE FATURAS...
echo ==========================================
echo.

cd /d "%~dp0"

set VENV_PATH=9_Arquivos_do_Sistema\.venv\Scripts\python.exe
set SCRIPT_PATH=9_Arquivos_do_Sistema\leitor_faturas.py

if exist "%VENV_PATH%" (
    "%VENV_PATH%" "%SCRIPT_PATH%"
) else (
    echo [AVISO] Ambiente virtual nao encontrado em: %VENV_PATH%
    echo A tentar correr com o python do sistema...
    python "%SCRIPT_PATH%"
)

pause
