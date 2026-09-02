@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo ================================================
    echo  Python nao encontrado!
    echo  Instale em: https://www.python.org/downloads/
    echo  Durante a instalacao, MARQUE a opcao:
    echo  "Add python.exe to PATH"
    echo  Depois rode este arquivo novamente.
    echo ================================================
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Criando ambiente virtual...
    python -m venv .venv
)

echo Instalando dependencias (primeira vez demora um pouco)...
".venv\Scripts\python.exe" -m pip install -r requirements.txt

echo Iniciando o Almoxarifado...
start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py

pause