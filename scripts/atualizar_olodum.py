"""Atualiza a OLODUM_BASEDADOS e a base do painel sem ninguém precisar abrir o Excel.

Fica rodando em segundo plano (o Agendador de Tarefas inicia ao ligar o computador) e atualiza:
  - sempre que uma das planilhas da pasta 01-Base de dados for salva;
  - e, mesmo sem mudança, a cada INTERVALO_MINUTOS.

Cada atualização:
  1. abre a OLODUM_BASEDADOS escondida, faz o "Atualizar Tudo", salva e fecha;
  2. gera o arquivo de dados do painel (painel_sistema_pcp.py);
  3. copia esse arquivo para a pasta do OneDrive/SharePoint de onde o painel lê.

Uso:
  python atualizar_olodum.py            fica observando (é o que o Agendador roda)
  python atualizar_olodum.py --uma-vez  atualiza uma vez e sai (para testar)

Precisa de: pip install pywin32 openpyxl
"""
import datetime
import logging
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

import painel_sistema_pcp

# ---- Ajuste estes caminhos -------------------------------------------------------------
PASTA_OLODUM = Path(r"J:\PARNAIBA - UGNL 2\04. QAQC\09. Matriz de rastreabilidade\07.OLODUM")
PASTA_BASES = PASTA_OLODUM / "01-Base de dados"
PLANILHA_OLODUM = PASTA_OLODUM / "OLODUM_BASEDADOS_Rev1.xlsx"
# Pasta sincronizada do OneDrive/SharePoint onde fica o arquivo que o painel lê.
PASTA_PAINEL = Path(os.environ.get("OneDriveCommercial", Path.home() / "OneDrive")) / "Painel OLODUM"
ARQUIVO_PAINEL = "pcp_painel.json"

INTERVALO_MINUTOS = 30       # atualização periódica, mesmo sem arquivo novo
VERIFICAR_SEGUNDOS = 30      # de quanto em quanto tempo olha se algum arquivo foi salvo
ESPERA_ESTAVEL_SEGUNDOS = 60  # espera o arquivo parar de mudar antes de atualizar
# -----------------------------------------------------------------------------------------

LOG = Path(__file__).with_name("atualizar_olodum.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(LOG, encoding="utf-8"), logging.StreamHandler()],
)


def retrato_bases():
    """Data de modificação de cada planilha da pasta de bases (sem arquivos temporários ~$)."""
    retrato = {}
    for arq in PASTA_BASES.glob("*.xls*"):
        if not arq.name.startswith("~$"):
            retrato[arq.name] = arq.stat().st_mtime
    return retrato


def atualizar_excel():
    """Abre a OLODUM escondida, roda o Atualizar Tudo até o fim, salva e fecha."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    try:
        livro = excel.Workbooks.Open(str(PLANILHA_OLODUM), UpdateLinks=0, ReadOnly=False)
        if livro.ReadOnly:
            livro.Close(SaveChanges=False)
            raise RuntimeError("a OLODUM está aberta por outra pessoa; tento de novo na próxima rodada")
        # Sem atualização em segundo plano, o RefreshAll só termina quando todas as consultas acabam.
        for conexao in livro.Connections:
            try:
                conexao.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass
        livro.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        excel.CalculateFull()
        livro.Save()
        livro.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


def publicar_dados():
    """Gera a base do painel a partir da OLODUM salva e copia para a pasta do OneDrive."""
    PASTA_PAINEL.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        saida = Path(tmp) / ARQUIVO_PAINEL
        painel_sistema_pcp.main(str(PLANILHA_OLODUM), str(saida))
        destino = PASTA_PAINEL / ARQUIVO_PAINEL
        # Copia para um nome provisório e troca de uma vez, para o OneDrive nunca subir um arquivo pela metade.
        provisorio = destino.with_suffix(".tmp")
        shutil.copyfile(saida, provisorio)
        os.replace(provisorio, destino)
    return destino


def atualizar(motivo):
    logging.info("Atualizando (%s)", motivo)
    inicio = time.time()
    try:
        atualizar_excel()
        destino = publicar_dados()
        logging.info("Pronto em %.0f s: %s", time.time() - inicio, destino)
        return True
    except Exception:
        logging.exception("Falha ao atualizar")
        return False


def observar():
    logging.info("Observando %s (e atualizando a cada %d min)", PASTA_BASES, INTERVALO_MINUTOS)
    visto = retrato_bases()
    ultima = 0.0
    mudou_em = None
    while True:
        time.sleep(VERIFICAR_SEGUNDOS)
        try:
            agora = retrato_bases()
        except OSError:
            logging.warning("Sem acesso à pasta %s; tento de novo", PASTA_BASES)
            continue
        if agora != visto:
            visto = agora
            mudou_em = time.time()  # recomeça a espera enquanto o arquivo ainda está sendo salvo
            continue
        if mudou_em and time.time() - mudou_em >= ESPERA_ESTAVEL_SEGUNDOS:
            if atualizar("planilha salva na pasta de bases"):
                ultima = time.time()
            mudou_em = None
        elif time.time() - ultima >= INTERVALO_MINUTOS * 60:
            if atualizar("rodada de %s" % datetime.datetime.now().strftime("%H:%M")):
                ultima = time.time()
            else:
                ultima = time.time() - INTERVALO_MINUTOS * 60 + 5 * 60  # nova tentativa em 5 min


if __name__ == "__main__":
    if "--uma-vez" in sys.argv:
        sys.exit(0 if atualizar("execução manual") else 1)
    observar()
