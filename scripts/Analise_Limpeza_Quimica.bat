@echo off
rem Botao "Analise de Limpeza Quimica": cruza as pastas com X em Limpeza Quimica da 06-Lista de Pastas_Milplan
rem com as atividades do HMSWEB (PCP_ATIVIDADE da OLODUM) e abre o Analise_Limpeza_Quimica.xlsx.
setlocal enabledelayedexpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"

echo.
echo Analisando as pastas de limpeza quimica. Nao feche esta janela...
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

"!PYEXE!" "%~dp0sistema\analise_limpeza_quimica.py"
if errorlevel 1 goto erro

echo.
echo Abrindo a analise...
for /f "delims=" %%F in ('dir /b /o-d "%~dp0Analise_Limpeza_Quimica*.xlsx" 2^>nul') do (start "" "%~dp0%%F" & goto fim)
:fim
timeout /t 5 >nul
exit /b 0

:sempython
echo Nao encontrei o Python. Peca para o Renan copiar a pasta "python" para dentro de "sistema".
pause
exit /b 1

:erro
echo.
echo Nao foi possivel fazer a analise. Veja a mensagem acima.
pause
exit /b 1
