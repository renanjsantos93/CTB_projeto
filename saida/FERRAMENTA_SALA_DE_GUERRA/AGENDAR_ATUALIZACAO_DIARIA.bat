@echo off
rem Agenda a atualizacao automatica de segunda a sexta as 07:40 (no seu usuario, com a J: mapeada).
rem Para mudar o horario, troque 07:40. Para remover: schtasks /Delete /TN "Controle Pastas Sala de Guerra" /F
schtasks /Create /F /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 07:40 /TN "Controle Pastas Sala de Guerra" /TR "\"%~dp0ATUALIZAR_CONTROLE_PASTAS.bat\" auto"
if errorlevel 1 (echo Nao foi possivel agendar. Rode este arquivo de novo ou fale com a TI.) else (echo Agendado: seg a sex, 07:40.)
pause
