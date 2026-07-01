@echo off
REM Agendador mensal — coleta métricas do mês anterior automaticamente
REM Configure no Agendador de Tarefas do Windows para rodar no dia 2 de cada mês

cd /d "%~dp0"

REM Verifica se o Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado. Instale em python.org
    pause
    exit /b 1
)

REM Instala dependencias se necessario
pip install requests --quiet

REM Executa a coleta
python coletar_metricas.py

echo.
echo Coleta concluida.
pause
