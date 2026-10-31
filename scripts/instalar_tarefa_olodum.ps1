# Cria a tarefa do Agendador que deixa o atualizar_olodum.py rodando sempre que você entra no Windows.
# Rode uma vez, no PowerShell, de dentro da pasta scripts:
#   powershell -ExecutionPolicy Bypass -File .\instalar_tarefa_olodum.ps1

$pasta = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $pasta "atualizar_olodum.py"

# pythonw roda sem abrir janela preta.
$pythonw = (Get-Command pythonw.exe -ErrorAction Stop).Source

$acao = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$script`"" -WorkingDirectory $pasta
$gatilho = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
# Se o script cair, o Windows sobe de novo em 5 minutos; sem limite de tempo de execução.
$config = New-ScheduledTaskSettingsSet -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 5) `
    -ExecutionTimeLimit (New-TimeSpan -Seconds 0) -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries

Register-ScheduledTask -TaskName "Atualizar OLODUM" -Action $acao -Trigger $gatilho -Settings $config `
    -Description "Atualiza a OLODUM_BASEDADOS e a base do painel quando as planilhas mudam e a cada 30 min" -Force

Start-ScheduledTask -TaskName "Atualizar OLODUM"
Write-Host "Tarefa 'Atualizar OLODUM' criada e iniciada. O registro fica em $pasta\atualizar_olodum.log"
