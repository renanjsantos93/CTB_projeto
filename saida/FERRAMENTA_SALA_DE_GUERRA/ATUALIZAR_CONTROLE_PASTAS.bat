@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
rem ==========================================================================
rem  ATUALIZA O CONTROLE DE PASTAS - SALA DE GUERRA
rem  Esta pasta (PAINEL DE ATUALIZACAO) fica dentro da pasta "Painel OLODUM" do OneDrive.
rem  Le o Painel OLODUM (.html) e a OLODUM_BASEDADOS (.xlsx) mais recentes da pasta
rem  "Painel OLODUM" (e subpastas) e atualiza SOMENTE as colunas automaticas e a aba
rem  BASE_OLODUM da planilha de controle na rede (CONTROLE abaixo).
rem ==========================================================================
set "CONTROLE=J:\PARNAIBA - UGNL 2\04. QAQC\09. Matriz de rastreabilidade\10-CONTROLE DE PASTA\CONTROLE_PASTAS_SALA_DE_GUERRA.xlsx"

set "PASTA_ATUALIZACAO=%~dp0"
if "%PASTA_ATUALIZACAO:~-1%"=="\" set "PASTA_ATUALIZACAO=%PASTA_ATUALIZACAO:~0,-1%"
for %%I in ("%PASTA_ATUALIZACAO%\..") do set "PASTA_OLODUM=%%~fI"
set "LOG=%PASTA_ATUALIZACAO%\HISTORICO\ultima_atualizacao.log"
if not exist "%PASTA_ATUALIZACAO%\HISTORICO" mkdir "%PASTA_ATUALIZACAO%\HISTORICO"

if not exist "%CONTROLE%" (
  echo ERRO: planilha nao encontrada:
  echo   %CONTROLE%
  echo Verifique se a unidade J: esta conectada e o nome do arquivo ^(linha CONTROLE deste .bat^).
  if /i not "%~1"=="auto" pause
  exit /b 1
)

echo Procurando o Python...
set "PY="
call :tentar "py"
call :tentar "python"
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do call :tentar "%%~D\python.exe"
for %%D in ("%USERPROFILE%\anaconda3" "%USERPROFILE%\miniconda3" "%LOCALAPPDATA%\anaconda3" "%ProgramData%\anaconda3" "%ProgramData%\miniconda3") do call :tentar "%%~D\python.exe"
for /d %%D in ("%ProgramFiles%\Python3*") do call :tentar "%%~D\python.exe"
if not defined PY for /f "delims=" %%P in ('where /r "%PASTA_OLODUM%" python.exe 2^>nul') do call :tentar "%%P"
if not defined PY (
  echo.
  echo ERRO: nenhum Python que funcione foi encontrado neste computador.
  echo Procurei no PATH, em AppData\Local\Programs\Python, Anaconda e em:
  echo   %PASTA_OLODUM%
  echo Me envie o caminho do python.exe usado para gerar o Painel OLODUM.
  if /i not "%~1"=="auto" pause
  exit /b 1
)
echo Python: %PY%
"%PY%" -c "import openpyxl" >nul 2>nul || "%PY%" -m pip install --user openpyxl

echo Atualizando %CONTROLE% ...
"%PY%" "%~dp0sala_guerra_pastas.py" "%PASTA_OLODUM%" "%CONTROLE%" --atualizar --pcp "%PASTA_OLODUM%" --backup "%PASTA_ATUALIZACAO%\HISTORICO" > "%LOG%" 2>&1
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

rem --- testa se o Python indicado executa de verdade (ignora o atalho da Microsoft Store)
:tentar
if defined PY exit /b 0
"%~1" -c "import sys" >nul 2>nul && set "PY=%~1"
exit /b 0
