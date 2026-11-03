r"""Atualiza a OLODUM_BASEDADOS e a base do painel sem ninguém precisar abrir o Excel.

Fica rodando em segundo plano (o Agendador de Tarefas inicia ao ligar o computador) e atualiza:
  - sempre que uma das planilhas da pasta 01-Base de dados for salva;
  - e, mesmo sem mudança, a cada INTERVALO_MINUTOS.

Cada atualização:
  1. abre a OLODUM_BASEDADOS escondida, faz o "Atualizar Tudo", salva e fecha;
  2. gera a base do painel (painel_sistema_pcp.py) e o Painel_OLODUM.html (gerar_painel_html.py);
  3. grava os dois na pasta compartilhada do OneDrive.

Pasta compartilhada (o "Painel OLODUM" do OneDrive):
  Painel_OLODUM.html     o painel; abre em qualquer navegador
  Atualizar_Painel.bat   o botão: atualiza tudo na hora e abre o painel
  pcp_painel.json        a base, para trocar no painel do claude.ai se quiser
  sistema\              estes scripts (e, opcionalmente, um Python portátil em sistema\python)

Uso:
  python atualizar_olodum.py            fica observando (é o que o Agendador roda)
  python atualizar_olodum.py --uma-vez  atualiza uma vez e sai (é o que o Atualizar_Painel.bat roda)

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

# Python portátil (embeddable) não põe a pasta do script no caminho de importação; garante que acha o vizinho.
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import gerar_painel_html  # noqa: E402
import painel_sistema_pcp  # noqa: E402

# ---- Ajuste estes caminhos -------------------------------------------------------------
PASTA_OLODUM = Path(r"J:\PARNAIBA - UGNL 2\04. QAQC\09. Matriz de rastreabilidade\07.OLODUM")
PASTA_BASES = PASTA_OLODUM / "01-Base de dados"
PLANILHA_OLODUM = PASTA_OLODUM / "OLODUM_BASEDADOS_Rev1.xlsx"
# Pasta compartilhada do painel: a que contém a pasta "sistema" destes scripts;
# fora dela, a pasta "Painel OLODUM" do OneDrive de quem roda.
if AQUI.name.lower() == "sistema":
    PASTA_PAINEL = AQUI.parent
else:
    PASTA_PAINEL = Path(os.environ.get("OneDriveCommercial", Path.home() / "OneDrive")) / "Painel OLODUM"
ARQUIVO_PAINEL = "pcp_painel.json"
ARQUIVO_HTML = "Painel_OLODUM.html"
# Evita duas atualizações ao mesmo tempo (o botão de alguém e a rodada automática, por exemplo).
TRAVA = PASTA_PAINEL / "atualizando.lock"
TRAVA_VALIDA_MINUTOS = 20

INTERVALO_MINUTOS = 30       # atualização periódica, mesmo sem arquivo novo
VERIFICAR_SEGUNDOS = 30      # de quanto em quanto tempo olha se algum arquivo foi salvo
ESPERA_ESTAVEL_SEGUNDOS = 60  # espera o arquivo parar de mudar antes de atualizar
# -----------------------------------------------------------------------------------------

LOG = AQUI / "atualizar_olodum.log"
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


def quem_esta_com_a_olodum():
    """Nome gravado pelo Excel no arquivo de trava ~$ ao lado da planilha (1º byte = tamanho do nome)."""
    trava = PLANILHA_OLODUM.with_name("~$" + PLANILHA_OLODUM.name[2:])
    try:
        dados = trava.read_bytes()
        nome = dados[1:1 + dados[0]].decode("latin-1").strip()
        return "%s (trava %s)" % (nome or "alguém", trava.name)
    except (OSError, IndexError):
        return "outra pessoa"


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
            raise RuntimeError("a OLODUM está aberta por %s; feche a planilha (ou o Excel escondido no "
                               "Gerenciador de Tarefas) e tente de novo" % quem_esta_com_a_olodum())
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


def gravar(origem, destino):
    """Copia para um nome provisório e troca de uma vez, para o OneDrive nunca subir um arquivo pela metade."""
    provisorio = destino.with_name(destino.name + ".tmp")
    shutil.copyfile(origem, provisorio)
    os.replace(provisorio, destino)


def publicar_dados():
    """Gera a base e o Painel_OLODUM.html a partir da OLODUM salva e grava na pasta compartilhada."""
    PASTA_PAINEL.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        base = Path(tmp) / ARQUIVO_PAINEL
        html = Path(tmp) / ARQUIVO_HTML
        painel_sistema_pcp.main(str(PLANILHA_OLODUM), str(base))
        gerar_painel_html.gerar(base, html)
        gravar(base, PASTA_PAINEL / ARQUIVO_PAINEL)
        gravar(html, PASTA_PAINEL / ARQUIVO_HTML)
    return PASTA_PAINEL / ARQUIVO_HTML


def pegar_trava():
    """True se ninguém mais está atualizando; uma trava esquecida há mais de TRAVA_VALIDA_MINUTOS é ignorada."""
    try:
        if TRAVA.exists() and time.time() - TRAVA.stat().st_mtime < TRAVA_VALIDA_MINUTOS * 60:
            return False
        TRAVA.write_text("%s %s" % (os.environ.get("USERNAME", "?"), datetime.datetime.now().isoformat()), encoding="utf-8")
        return True
    except OSError:
        return True  # sem conseguir criar a trava, atualiza assim mesmo


def soltar_trava():
    try:
        TRAVA.unlink()
    except OSError:
        pass


def atualizar(motivo):
    if not pegar_trava():
        logging.info("Outra pessoa já está atualizando o painel; pulando (%s)", motivo)
        return True
    logging.info("Atualizando (%s) por %s", motivo, os.environ.get("USERNAME", "?"))
    inicio = time.time()
    try:
        atualizar_excel()
        destino = publicar_dados()
        logging.info("Pronto em %.0f s: %s", time.time() - inicio, destino)
        return True
    except Exception:
        logging.exception("Falha ao atualizar")
        return False
    finally:
        soltar_trava()


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
