@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
rem ==========================================================================
rem  ATUALIZA O CONTROLE DE PASTAS - SALA DE GUERRA
rem  Le o Painel OLODUM (.html) e a OLODUM_BASEDADOS (.xlsx) mais recentes que
rem  estiverem NESTA pasta (PAINEL DE ATUALIZACAO) e atualiza SOMENTE as colunas
rem  automaticas (ITEM ate STATUS CALCULADO, AVANCO e PROGRAMACAO DA SEMANA) e a
rem  aba BASE_OLODUM da planilha de controle que fica na pasta de cima.
rem  Para trocar a planilha (outro nome de arquivo), altere so a linha NOME_CONTROLE.
rem ==========================================================================
set "NOME_CONTROLE=CONTROLE_PASTAS_SALA_DE_GUERRA.xlsx"

set "PASTA_DADOS=%~dp0"
if "%PASTA_DADOS:~-1%"=="\" set "PASTA_DADOS=%PASTA_DADOS:~0,-1%"
for %%I in ("%PASTA_DADOS%\..") do set "PASTA_CONTROLE=%%~fI"
set "CONTROLE=%PASTA_CONTROLE%\%NOME_CONTROLE%"
set "LOG=%PASTA_DADOS%\HISTORICO\ultima_atualizacao.log"

if not exist "%CONTROLE%" (
  echo ERRO: planilha nao encontrada: %CONTROLE%
  echo Confira a linha NOME_CONTROLE neste .bat.
  if /i not "%~1"=="auto" pause
  exit /b 1
)
if not exist "%PASTA_DADOS%\HISTORICO" mkdir "%PASTA_DADOS%\HISTORICO"

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY ( where python >nul 2>nul && set "PY=python" )
if not defined PY (
  echo ERRO: Python nao encontrado. Veja o LEIA-ME.txt ^(passo 1^).
  if /i not "%~1"=="auto" pause
  exit /b 1
)
%PY% -c "import openpyxl" >nul 2>nul || %PY% -m pip install --user openpyxl

echo Atualizando %NOME_CONTROLE% ...
%PY% "%~dp0sala_guerra_pastas.py" "%PASTA_DADOS%" "%CONTROLE%" --atualizar --pcp "%PASTA_DADOS%" --backup "%PASTA_DADOS%\HISTORICO" > "%LOG%" 2>&1
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
