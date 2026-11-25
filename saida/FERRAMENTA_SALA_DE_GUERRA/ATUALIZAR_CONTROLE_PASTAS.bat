@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
rem ==========================================================================
rem  ATUALIZA O CONTROLE DE PASTAS - SALA DE GUERRA
rem  Esta pasta (PAINEL DE ATUALIZACAO) fica dentro da pasta compartilhada "Painel OLODUM"
rem  do OneDrive. Le o Painel OLODUM (.html) e a OLODUM_BASEDADOS (.xlsx) mais recentes
rem  dessa pasta e atualiza a planilha que fica em "Painel OLODUM\CONTROLE SALA DE GUERRA".
rem  Funciona para qualquer pessoa com acesso a pasta compartilhada.
rem  Uso: dois cliques (atualiza agora) | "auto" (tarefa agendada: so atualiza se a base for nova)
rem ==========================================================================
set "PASTA_ATUALIZACAO=%~dp0"
if "%PASTA_ATUALIZACAO:~-1%"=="\" set "PASTA_ATUALIZACAO=%PASTA_ATUALIZACAO:~0,-1%"
for %%I in ("%PASTA_ATUALIZACAO%\..") do set "PASTA_OLODUM=%%~fI"
set "CONTROLE=%PASTA_OLODUM%\CONTROLE SALA DE GUERRA\CONTROLE_PASTAS_SALA_DE_GUERRA.xlsx"
set "HIST=%PASTA_ATUALIZACAO%\HISTORICO"
if not exist "%HIST%" mkdir "%HIST%"

if not exist "%CONTROLE%" (
  echo ERRO: planilha nao encontrada:
  echo   %CONTROLE%
  echo Coloque a CONTROLE_PASTAS_SALA_DE_GUERRA.xlsx na pasta "CONTROLE SALA DE GUERRA" dentro de "Painel OLODUM".
  if /i not "%~1"=="auto" pause
  exit /b 1
)

set "PY="
call :tentar "py"
call :tentar "python"
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do call :tentar "%%~D\python.exe"
for %%D in ("%USERPROFILE%\anaconda3" "%USERPROFILE%\miniconda3" "%LOCALAPPDATA%\anaconda3" "%ProgramData%\anaconda3" "%ProgramData%\miniconda3") do call :tentar "%%~D\python.exe"
for /d %%D in ("%ProgramFiles%\Python3*") do call :tentar "%%~D\python.exe"
if not defined PY for /f "delims=" %%P in ('where /r "%PASTA_OLODUM%" python.exe 2^>nul') do call :tentar "%%P"
if not defined PY (
  echo ERRO: nenhum Python que funcione foi encontrado neste computador.
  if /i not "%~1"=="auto" pause
  exit /b 1
)
"%PY%" -c "import openpyxl" >nul 2>nul || "%PY%" -m pip install --user openpyxl >nul 2>nul

if /i "%~1"=="auto" goto :automatico

echo Python: %PY%
echo Atualizando %CONTROLE% ...
"%PY%" "%~dp0sala_guerra_pastas.py" "%PASTA_OLODUM%" "%CONTROLE%" --atualizar --pcp "%PASTA_OLODUM%" --backup "%HIST%" > "%HIST%\ultima_atualizacao.log" 2>&1
set ERRO=%ERRORLEVEL%
type "%HIST%\ultima_atualizacao.log"
echo.
if not "%ERRO%"=="0" (
  echo *** ATUALIZACAO NAO CONCLUIDA - veja a mensagem acima ***
) else (
  echo Concluido. Abra a planilha: o Excel recalcula tudo ao abrir.
)
pause
exit /b %ERRO%

:automatico
"%PY%" "%~dp0sala_guerra_pastas.py" "%PASTA_OLODUM%" "%CONTROLE%" --atualizar --automatico --pcp "%PASTA_OLODUM%" --backup "%HIST%" >> "%HIST%\atualizacao_automatica.log" 2>&1
exit /b

rem --- testa se o Python indicado executa de verdade (ignora o atalho da Microsoft Store)
:tentar
if defined PY exit /b 0
"%~1" -c "import sys" >nul 2>nul && set "PY=%~1"
exit /b 0
