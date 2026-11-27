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
set "RAIZ=%~dp0"
rem 1) Python portatil dentro da pasta do painel: python\, sistema\python\ ou qualquer subpasta
if exist "%RAIZ%python\python.exe" set "PYEXE=%RAIZ%python\python.exe"
if not defined PYEXE if exist "%RAIZ%sistema\python\python.exe" set "PYEXE=%RAIZ%sistema\python\python.exe"
if not defined PYEXE for /r "%RAIZ%." %%F in (python.ex?) do if not defined PYEXE if /i "%%~nxF"=="python.exe" set "PYEXE=%%~fF"
rem 2) Lancador "py" ou comando "python" do computador
if not defined PYEXE for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
if not defined PYEXE for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
rem 3) Python portatil numa pasta Painel* do OneDrive ou da Area de Trabalho (ex.: Painel_CONTROLTUB\python)
for %%R in ("%OneDriveCommercial%" "%OneDrive%" "%USERPROFILE%") do (
  if not defined PYEXE for /d %%A in ("%%~R\*") do (
    if not defined PYEXE for /d %%B in ("%%~A\Painel*") do if exist "%%~B\python\python.exe" set "PYEXE=%%~B\python\python.exe"
  )
)
if not defined PYEXE for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
if not defined PYEXE goto sempython
echo Python: !PYEXE!
"!PYEXE!" -c "import win32com.client, openpyxl" >nul 2>&1
if errorlevel 1 goto sembibliotecas

"!PYEXE!" "%~dp0sistema\atualizar_olodum.py" --uma-vez
if errorlevel 2 goto aviso
if errorlevel 1 goto erro

echo.
echo Painel atualizado. Abrindo...
start "" "%~dp0Painel_OLODUM.html"
timeout /t 5 >nul
exit /b 0

:aviso
echo.
echo ATENCAO: o painel foi atualizado, mas algumas bases nao entraram (veja os itens PROBLEMA acima).
echo Corrija e clique de novo. Aperte uma tecla para abrir o painel.
pause >nul
start "" "%~dp0Painel_OLODUM.html"
exit /b 0

:sembibliotecas
echo.
echo O Python !PYEXE! nao tem as bibliotecas pywin32 e openpyxl.
echo Use a pasta "python" que esta dentro da pasta Painel OLODUM (copiada do computador do Renan)
echo ou instale com: "!PYEXE!" -m pip install pywin32 openpyxl
pause
exit /b 1

:sempython
echo Nao encontrei o Python. A pasta "python" (Python portatil) precisa estar dentro da pasta Painel OLODUM,
echo com os arquivos baixados neste computador (OneDrive: botao direito > Sempre manter neste dispositivo).
pause
exit /b 1

:erro
echo.
echo Nao foi possivel atualizar. Veja a mensagem acima e o arquivo sistema\atualizar_olodum.log.
echo Se a OLODUM estiver aberta por alguem, feche a planilha e clique de novo.
pause
exit /b 1
