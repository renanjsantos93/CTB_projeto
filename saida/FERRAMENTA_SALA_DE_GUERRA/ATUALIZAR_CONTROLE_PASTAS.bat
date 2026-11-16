@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
rem ==========================================================================
rem  ATUALIZA O CONTROLE DE PASTAS - SALA DE GUERRA
rem  Le o Painel OLODUM (.html mais recente) e a OLODUM_BASEDADOS (mais recente)
rem  da pasta 07.OLODUM e atualiza a planilha de controle na pasta 10.
rem  Se mudar algum caminho, altere so as linhas "set" abaixo.
rem ==========================================================================
set "BASE=J:\PARNAIBA - UGNL 2\04. QAQC\09. Matriz de rastreabilidade"
set "PASTA_OLODUM=%BASE%\07.OLODUM"
set "PASTA_CONTROLE=%BASE%\10-CONTROLE DE PASTA"
set "CONTROLE=%PASTA_CONTROLE%\CONTROLE_PASTAS_SALA_DE_GUERRA.xlsx"
set "LOG=%PASTA_CONTROLE%\HISTORICO\ultima_atualizacao.log"

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
