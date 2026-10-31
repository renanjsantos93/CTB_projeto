@echo off
rem Instala a atualizacao automatica da OLODUM. Basta dar dois cliques neste arquivo.
cd /d "%~dp0"

echo.
echo [1/3] Instalando as bibliotecas do Python...
python -m pip install --upgrade pywin32 openpyxl
if errorlevel 1 goto erro

echo.
echo [2/3] Testando: atualizando a OLODUM e gerando o arquivo do painel (pode levar alguns minutos)...
python atualizar_olodum.py --uma-vez
if errorlevel 1 goto erro

echo.
echo [3/3] Criando a tarefa "Atualizar OLODUM" no Agendador de Tarefas...
powershell -ExecutionPolicy Bypass -File "%~dp0instalar_tarefa_olodum.ps1"
if errorlevel 1 goto erro

echo.
echo Tudo pronto. A partir de agora a OLODUM se atualiza sozinha.
pause
exit /b 0

:erro
echo.
echo Algo deu errado. Veja a mensagem acima e o arquivo atualizar_olodum.log nesta pasta.
pause
exit /b 1
