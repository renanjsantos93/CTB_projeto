r"""Confere se a OLODUM e o painel ficaram iguais às bases da pasta 01-Base de dados.

Roda no fim de cada atualização (atualizar_olodum.py) e também sozinho:
  python verificar_bases.py [<OLODUM_BASEDADOS_Rev1.xlsx> [<pcp_painel.json>]]

O que confere:
  1. As consultas da OLODUM (Power Query): lê no próprio arquivo de onde cada uma busca os dados
     (ex.: 01-Base de dados\04-Dataplanejamento.xlsx, aba DATAPLANEJADA) e avisa se o arquivo não existe
     (nome trocado, extensão diferente, arquivo movido para "Obsoleto").
  2. Se a tabela da OLODUM tem as mesmas linhas da sua base (quantidade de linhas preenchidas).
     Diferença = a consulta não atualizou e a OLODUM ficou com os dados antigos. Na atualização só
     confere a DATAPLANEJADA (as outras são grandes e a leitura pela rede é lenta); rodando sozinho, todas.
  3. A data planejada de cada subsistema no painel contra a 04-Dataplanejamento (coluna D = subsistema,
     coluna E = data, a primeira linha de cada subsistema, como o PROCV da coluna DATA PLANEJADA).
Devolve a lista de problemas (vazia = tudo confere).
"""
import base64
import datetime
import io
import json
import re
import struct
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl

# Consulta da OLODUM -> aba da OLODUM onde ela é carregada (quando o nome da aba não é o da consulta).
ABA_DA_CONSULTA = {"PCPHMWEB": "PCP_ATIVIDADE"}
CONSULTA_DATAS = "DATAPLANEJADA"
COL_SUBSISTEMA, COL_DATA = 3, 4  # colunas D e E da aba DATAPLANEJADA (0 = coluna A)


def consultas_da_olodum(planilha):
    """{consulta: (arquivo, aba, promove_cabecalho)} lidos do código Power Query salvo dentro da OLODUM."""
    with zipfile.ZipFile(planilha) as z:
        for nome in z.namelist():
            if not nome.startswith("customXml/item") or "Props" in nome:
                continue
            bruto = z.read(nome)
            texto = bruto.decode("utf-16") if bruto[:2] in (b"\xff\xfe", b"\xfe\xff") else bruto.decode("utf-8", "ignore")
            m = re.search(r"<DataMashup[^>]*>\s*([A-Za-z0-9+/=\s]+?)\s*</DataMashup>", texto)
            if not m:
                continue
            pacote = base64.b64decode(m.group(1))
            tamanho = struct.unpack("<I", pacote[4:8])[0]
            codigo = zipfile.ZipFile(io.BytesIO(pacote[8:8 + tamanho])).read("Formulas/Section1.m").decode("utf-8", "replace")
            consultas = {}
            for bloco in re.split(r"\n\s*shared\s+", codigo)[1:]:
                nome_c = re.match(r'(#"[^"]+"|[^\s=]+)', bloco).group(1).strip('#"')
                arq = re.search(r'File\.Contents\("([^"]+)"\)', bloco)
                aba = re.search(r'\{\[Item="([^"]+)",Kind="Sheet"\]\}', bloco)
                if arq:
                    consultas[nome_c] = (arq.group(1), aba.group(1) if aba else None, "PromoteHeaders" in bloco)
            return consultas
    return {}


def linhas_preenchidas(planilha, aba):
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    try:
        ws = wb[aba]
    except KeyError:
        return None
    n = sum(1 for r in ws.iter_rows(values_only=True) if any(v not in (None, "") for v in r))
    wb.close()
    return n


def data_iso(v):
    if hasattr(v, "date"):
        return v.date().isoformat()
    if isinstance(v, (int, float)) and 30000 < v < 80000:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", str(v or "").strip())
    if m:
        return "%s-%02d-%02d" % (m.group(3), int(m.group(2)), int(m.group(1)))
    return str(v or "").strip()


def datas_da_base(planilha, aba):
    """Subsistema -> data, a primeira linha de cada subsistema (como o PROCV)."""
    datas = {}
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    for r in wb[aba].iter_rows(values_only=True):
        if len(r) > COL_DATA and r[COL_SUBSISTEMA]:
            datas.setdefault(str(r[COL_SUBSISTEMA]).strip(), r[COL_DATA])
    wb.close()
    return datas


def fmt(iso):
    return "%s/%s/%s" % (iso[8:10], iso[5:7], iso[:4]) if re.match(r"^\d{4}-\d{2}-\d{2}$", iso) else iso or "(vazio)"


def verificar(planilha_olodum, painel_json=None, log=print, completo=False):
    problemas = []

    def aviso(msg):
        problemas.append(msg)
        log("PROBLEMA: " + msg)

    try:
        consultas = consultas_da_olodum(planilha_olodum)
    except Exception as e:  # noqa: BLE001
        aviso("não consegui ler as consultas da OLODUM (%s)" % e)
        consultas = {}
    fonte_datas = None
    for nome, (arquivo, aba, promove) in sorted(consultas.items()):
        caminho = Path(arquivo)
        if not caminho.exists():
            pasta, _, nome_arq = arquivo.rpartition("\\")
            aviso('a consulta "%s" da OLODUM lê o arquivo "%s" (pasta %s), que não existe com esse nome; '
                  "a OLODUM fica com os dados antigos dessa consulta. Acerte o nome do arquivo ou a fonte da "
                  "consulta (Dados > Consultas e Conexões)" % (nome, nome_arq, pasta.rpartition("\\")[2] or pasta))
            continue
        quando = datetime.datetime.fromtimestamp(caminho.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
        if nome == CONSULTA_DATAS:
            fonte_datas = (caminho, aba)
        aba_olodum = ABA_DA_CONSULTA.get(nome, nome)
        # Contar as linhas das bases grandes (01, 02) pela rede leva minutos; na atualização, só a das datas.
        # Que cada consulta atualizou, o próprio Excel já conferiu (atualizar_olodum.conferir_consultas).
        if not completo and nome != CONSULTA_DATAS:
            log('OK  consulta "%s": %s existe (salva em %s)' % (nome, caminho.name, quando))
            continue
        try:
            n_base = linhas_preenchidas(caminho, aba)
            n_olodum = linhas_preenchidas(planilha_olodum, aba_olodum)
        except Exception as e:  # noqa: BLE001
            aviso('não consegui comparar a consulta "%s" com %s (%s)' % (nome, caminho.name, e))
            continue
        if n_base is None:
            aviso('a consulta "%s" busca a aba "%s" em %s, mas a aba não existe lá' % (nome, aba, caminho.name))
            continue
        if n_olodum is None:
            log('Consulta "%s": %s (salvo em %s), %d linhas; aba "%s" não encontrada na OLODUM para comparar'
                % (nome, caminho.name, quando, n_base, aba_olodum))
            continue
        # Sem "promover cabeçalho", a OLODUM ganha uma linha de cabeçalho a mais (Column1, Column2...).
        esperado = n_base if promove else n_base + 1
        if n_olodum != esperado:
            aviso('a aba "%s" da OLODUM tem %d linhas e a base %s (salva em %s) tem %d: a consulta "%s" não '
                  "atualizou" % (aba_olodum, n_olodum - 1, caminho.name, quando, esperado - 1, nome))
        else:
            log('OK  consulta "%s": %s (salva em %s), %d linhas iguais na OLODUM' % (nome, caminho.name, quando, n_olodum - 1))

    # Data planejada: painel x 04-Dataplanejamento.
    if fonte_datas and painel_json and Path(painel_json).exists():
        try:
            base = datas_da_base(*fonte_datas)
            painel = defaultdict(set)
            for r in json.loads(Path(painel_json).read_text(encoding="utf-8")):
                painel[r["subsistema"]].add(r["data_planejada"])
            diferentes = []
            for sub, datas in sorted(painel.items()):
                if sub not in base:
                    continue
                esperado = data_iso(base[sub])
                if re.match(r"^\d{4}-\d{2}-\d{2}$", esperado) and datas != {esperado}:
                    diferentes.append("%s: painel %s, base %s" % (sub, ", ".join(sorted(fmt(d) for d in datas)), fmt(esperado)))
            if diferentes:
                aviso("%d subsistema(s) com data planejada diferente da %s, por exemplo: %s"
                      % (len(diferentes), fonte_datas[0].name, "; ".join(diferentes[:5])))
            else:
                log("OK  data planejada: os %d subsistemas do painel que estão na %s têm a mesma data"
                    % (sum(1 for s in painel if s in base), fonte_datas[0].name))
        except Exception as e:  # noqa: BLE001
            aviso("não consegui comparar as datas planejadas (%s)" % e)
    return problemas


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import atualizar_olodum as ao
        args = [ao.PLANILHA_OLODUM, ao.PASTA_PAINEL / ao.ARQUIVO_PAINEL]
    sys.exit(1 if verificar(*args, completo=True) else 0)
