@echo off
rem Botao "Atualizar Painel": atualiza a OLODUM, gera o Painel_OLODUM.html e abre o painel.
rem Fica na pasta compartilhada "Painel OLODUM", ao lado da pasta "sistema".
setlocal enabledelayedexpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"

echo.
echo Atualizando o Painel OLODUM. Nao feche esta janela (leva alguns minutos)...
echo.

set "PYEXE="
rem 1) Python portatil guardado junto do painel
if exist "%~dp0sistema\python\python.exe" set "PYEXE=%~dp0sistema\python\python.exe"
rem 2) Lancador "py" ou comando "python" do computador
if not defined PYEXE for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
if not defined PYEXE for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
rem 3) Python portatil numa pasta Painel_* do OneDrive ou da Area de Trabalho
for %%R in ("%OneDriveCommercial%" "%OneDrive%" "%USERPROFILE%") do (
  if not defined PYEXE for /d %%A in ("%%~R\*") do (
    if not defined PYEXE for /d %%B in ("%%~A\Painel_*") do if exist "%%~B\python\python.exe" set "PYEXE=%%~B\python\python.exe"
  )
)
if not defined PYEXE for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
if not defined PYEXE goto sempython

"!PYEXE!" "%~dp0sistema\atualizar_olodum.py" --uma-vez
if errorlevel 1 goto erro

echo.
echo Painel atualizado. Abrindo...
start "" "%~dp0Painel_OLODUM.html"
timeout /t 5 >nul
exit /b 0

:sempython
echo Nao encontrei o Python. Peca para o Renan copiar a pasta "python" para dentro de "sistema".
pause
exit /b 1

:erro
echo.
echo Nao foi possivel atualizar. Veja a mensagem acima e o arquivo sistema\atualizar_olodum.log.
echo Se a OLODUM estiver aberta por alguem, feche a planilha e clique de novo.
pause
exit /b 1
