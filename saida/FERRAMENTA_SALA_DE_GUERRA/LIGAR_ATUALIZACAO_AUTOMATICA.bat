@echo off
rem Liga a ATUALIZACAO AUTOMATICA neste computador: a cada 15 minutos verifica se chegou
rem Painel OLODUM / OLODUM_BASEDADOS nova na pasta "Painel OLODUM" e, se chegou, atualiza o controle.
rem Ligue em UM computador so (o que fica ligado durante o dia).
schtasks /Create /F /SC MINUTE /MO 15 /TN "Controle Pastas Sala de Guerra" /TR "wscript.exe \"%~dp0ATUALIZAR_OCULTO.vbs\""
if errorlevel 1 (
  echo Nao foi possivel agendar. Fale com a TI ou crie a tarefa pelo Agendador de Tarefas.
) else (
  echo Atualizacao automatica LIGADA: verifica a cada 15 minutos.
  echo Registro em PAINEL DE ATUALIZACAO\HISTORICO\atualizacao_automatica.log
)
pause
