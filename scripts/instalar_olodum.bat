@echo off
rem Instala a atualizacao automatica da OLODUM. Basta dar dois cliques neste arquivo.
setlocal enabledelayedexpansion
rem UTF-8 no console, para caminhos com acento (ex.: "Area de Trabalho" do OneDrive) nao se perderem.
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"

echo.
echo [0/3] Procurando o Python instalado...
set "PYEXE="
rem Pasta do painel (a que contem a pasta "sistema", onde este arquivo fica)
for %%I in ("%~dp0..") do set "RAIZ=%%~fI\"
rem 1) Python portatil dentro da pasta do painel: python\, sistema\python\ ou qualquer subpasta
if exist "%RAIZ%python\python.exe" set "PYEXE=%RAIZ%python\python.exe"
if not defined PYEXE if exist "%RAIZ%sistema\python\python.exe" set "PYEXE=%RAIZ%sistema\python\python.exe"
if not defined PYEXE for /r "%RAIZ%." %%F in (python.ex?) do if not defined PYEXE if /i "%%~nxF"=="python.exe" set "PYEXE=%%~fF"
rem 2) Lancador "py" do instalador oficial do python.org
if not defined PYEXE for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
rem 3) Comando "python" (ignora o atalho da Microsoft Store, que nao imprime nada)
if not defined PYEXE for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
rem 4) Python portatil dentro de uma pasta Painel* no OneDrive ou na Area de Trabalho (ex.: Painel_CONTROLTUB\python)
for %%R in ("%OneDriveCommercial%" "%OneDrive%" "%USERPROFILE%") do (
  if not defined PYEXE for /d %%A in ("%%~R\*") do (
    if not defined PYEXE for /d %%B in ("%%~A\Painel*") do if exist "%%~B\python\python.exe" set "PYEXE=%%~B\python\python.exe"
  )
)
rem 5) Pastas onde o Python e o Anaconda costumam ficar
if not defined PYEXE for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
if not defined PYEXE for /d %%D in ("%ProgramFiles%\Python3*") do if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
if not defined PYEXE for %%D in ("%USERPROFILE%\anaconda3" "%USERPROFILE%\miniconda3" "%LOCALAPPDATA%\anaconda3" "%ProgramData%\anaconda3") do if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
if not defined PYEXE goto sempython
echo Python encontrado: !PYEXE!

echo.
echo [1/3] Instalando as bibliotecas do Python...
rem Com as bibliotecas ja presentes (ex.: o Python portatil da pasta do painel), nao precisa de internet.
"!PYEXE!" -c "import win32com.client, openpyxl" >nul 2>&1
if errorlevel 1 (
  "!PYEXE!" -m pip install --upgrade pywin32 openpyxl
  if errorlevel 1 goto erro
) else (
  echo As bibliotecas ja estao instaladas.
)

echo.
echo [2/3] Testando: atualizando a OLODUM e gerando o arquivo do painel (pode levar alguns minutos)...
"!PYEXE!" atualizar_olodum.py --uma-vez
rem 2 = atualizou, mas a conferencia achou base que nao entrou: segue e cria a tarefa assim mesmo.
if errorlevel 2 (
  echo Atencao: o painel foi gerado, mas veja os itens PROBLEMA acima.
) else if errorlevel 1 goto erro

echo.
echo [3/3] Criando a tarefa "Atualizar OLODUM" no Agendador de Tarefas...
powershell -ExecutionPolicy Bypass -File "%~dp0instalar_tarefa_olodum.ps1" -Python "!PYEXE!"
if errorlevel 1 goto erro

echo.
echo Tudo pronto. A partir de agora a OLODUM se atualiza sozinha.
pause
exit /b 0

:sempython
echo.
echo Nao encontrei o Python neste computador.
echo Instale pelo site https://www.python.org/downloads/ e, na primeira tela do instalador,
echo marque a opcao "Add python.exe to PATH". Depois rode este arquivo de novo.
pause
exit /b 1

:erro
echo.
echo Algo deu errado. Veja a mensagem acima e o arquivo atualizar_olodum.log nesta pasta.
pause
exit /b 1
