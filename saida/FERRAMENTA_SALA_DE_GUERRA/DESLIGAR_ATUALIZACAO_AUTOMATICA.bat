@echo off
rem Desliga a atualizacao automatica neste computador.
schtasks /Delete /TN "Controle Pastas Sala de Guerra" /F
pause
