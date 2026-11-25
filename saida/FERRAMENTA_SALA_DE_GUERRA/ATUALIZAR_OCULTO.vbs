' Executa a atualizacao automatica sem abrir janela (usado pela tarefa agendada).
Set fso = CreateObject("Scripting.FileSystemObject")
pasta = fso.GetParentFolderName(WScript.ScriptFullName)
CreateObject("WScript.Shell").Run """" & pasta & "\ATUALIZAR_CONTROLE_PASTAS.bat"" auto", 0, False
