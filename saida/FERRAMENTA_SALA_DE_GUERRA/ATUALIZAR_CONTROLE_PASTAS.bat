@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
rem ==========================================================================
rem  ATUALIZA O CONTROLE DE PASTAS - SALA DE GUERRA
rem  Le o Painel OLODUM (.html mais recente) e a OLODUM_BASEDADOS (mais recente)
rem  da pasta 07.OLODUM e atualiza a planilha de controle que esta NESTA pasta
rem  (a mesma deste .bat, ex.: 10-CONTROLE DE PASTA\PAINEL DE ATUALIZACAO).
rem  Se a pasta do OLODUM mudar, altere so a linha PASTA_OLODUM abaixo.
rem ==========================================================================
set "BASE=J:\PARNAIBA - UGNL 2\04. QAQC\09. Matriz de rastreabilidade"
set "PASTA_OLODUM=%BASE%\07.OLODUM"
set "PASTA_CONTROLE=%~dp0"
if "%PASTA_CONTROLE:~-1%"=="\" set "PASTA_CONTROLE=%PASTA_CONTROLE:~0,-1%"
set "CONTROLE=%PASTA_CONTROLE%\CONTROLE_PASTAS_SALA_DE_GUERRA.xlsx"
set "LOG=%PASTA_CONTROLE%\HISTORICO\ultima_atualizacao.log"

if not exist "%PASTA_OLODUM%" (
  echo ERRO: pasta do OLODUM nao encontrada: %PASTA_OLODUM%
  echo Verifique se a unidade J: esta conectada.
  if /i not "%~1"=="auto" pause
  exit /b 1
)
if not exist "%CONTROLE%" (
  echo ERRO: CONTROLE_PASTAS_SALA_DE_GUERRA.xlsx nao esta na mesma pasta deste .bat.
  if /i not "%~1"=="auto" pause
  exit /b 1
)
if not exist "%PASTA_CONTROLE%\HISTORICO" mkdir "%PASTA_CONTROLE%\HISTORICO"

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY ( where python >nul 2>nul && set "PY=python" )
if not defined PY (
  echo ERRO: Python nao encontrado. Veja o LEIA-ME.txt ^(passo 1^).
  if /i not "%~1"=="auto" pause
  exit /b 1
)
%PY% -c "import openpyxl" >nul 2>nul || %PY% -m pip install --user openpyxl

echo Atualizando o controle de pastas...
%PY% "%~dp0sala_guerra_pastas.py" "%PASTA_OLODUM%" "%CONTROLE%" "%CONTROLE%" --pcp "%PASTA_OLODUM%" --backup "%PASTA_CONTROLE%\HISTORICO" > "%LOG%" 2>&1
set ERRO=%ERRORLEVEL%
type "%LOG%"
echo.
if not "%ERRO%"=="0" (
  echo *** ATUALIZACAO NAO CONCLUIDA - veja a mensagem acima ***
) else (
  echo Concluido. Abra a planilha: o Excel recalcula tudo ao abrir.
)
if /i not "%~1"=="auto" pause
exit /b %ERRO%
