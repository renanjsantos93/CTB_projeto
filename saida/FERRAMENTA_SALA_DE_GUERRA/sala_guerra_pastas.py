"""Gera o CONTROLE DE PASTAS - SALA DE GUERRA (tubulação) a partir do Painel OLODUM.

Uso:
    python scripts/sala_guerra_pastas.py PAINEL_OLODUM.html CONTROLE.xlsx SAIDA.xlsx \
        --pcp OLODUM_BASEDADOS.xlsx [--calendario 03.CALENDARIO_SEMANAL.xlsx]

PAINEL_OLODUM.html   painel exportado (ou a pasta onde ele fica: usa o .html de OLODUM mais recente) (PCP/HMSWEB, juntas do ControlTub, pendências e escopo).
CONTROLE.xlsx        controle anterior (gerado por este script, inclusive revisões editadas à mão
                     como a REV1). São preservados: Nº da pasta, SUBSOP, categoria de entrega, as
                     semanas programadas, data de conclusão, responsável, observação, restrições,
                     as semanas de avanço já gravadas e a aba RESTRIÇÕES.
--calendario         calendário da obra (DATA x SEMANA). Se omitido, usa a aba CALENDÁRIO do controle.
--pcp                planilha OLODUM_BASEDADOS (a mesma que gera o painel), aba PCP_ATIVIDADE, ou a pasta
                     onde ela fica (usa a OLODUM_BASEDADOS*.xlsx mais recente). Recomendado.
--backup PASTA       guarda uma cópia datada do controle antes de sobrescrevê-lo.
--painel-anterior    só sem --pcp: painel da atualização anterior (o que já estava aprovado nele recebe
                     "ATÉ S-xxx").

AVANÇO (MONTAGEM, TESTE, REMONTAGEM, PRÉ WALK) = semana do calendário da obra da DATA DA EXECUÇÃO na
PCP_ATIVIDADE:
    MONTAGEM   -> Lista de Verificação de Montagem - Tubulação
    TESTE      -> Teste Hidrostático em Tubulação
    REMONTAGEM -> Remontagem
    PRÉ WALK   -> só FURUI - INTRA SKID sem juntas no ControlTub: a última data de execução entre
                  mapa de juntas, teste hidrostático e lavagem (as que existirem na PCP), quando todas
                  estiverem executadas.
Datas fora do calendário seguem a mesma numeração (semanas de segunda a domingo).
Sem --pcp (método de contingência): a semana é gravada quando o script vê a atividade aprovada pela
primeira vez no painel, e as semanas já gravadas no controle são mantidas.
"""
import argparse
import collections
import datetime as dt
import glob
import json
import os
import re
import shutil
import sys
import zipfile
from xml.sax.saxutils import escape as _esc

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------------------- leitura
ATV_LVM = "Lista de Verificação de Montagem - Tubulação"
ATV_MJ = "Mapeamento de Juntas - Report ControlTub"
ATV_TH = "Teste Hidrostático em Tubulação"
ATV_LAV = "Lavagem"
ATV_REM = "Remontagem"
PREFIXO_PASTA = "421-M-TU"
ESC_RADIX, ESC_FURUI, ESC_INTRA, ESC_LOOSE = "RADIX", "FURUI", "FURUI - INTRA SKID", "FURUI - LOOSE ITEM"
ATV_STATUS = ("MONTAGEM (LVM)", "MAPA JUNTAS", "TESTE HIDROSTÁTICO", "LAVAGEM", "REMONTAGEM")


def regra_intra(b):
    """True quando a pasta segue o critério INTRA SKID (pré walkdown / documentação HMSWEB):
    INTRA SKID sem juntas no ControlTub, ou LOOSE ITEM sem juntas no ControlTub e com algum certificado
    executado ou aprovado no HMSWEB."""
    if (b["JUNTAS PREVISTAS"] or 0) > 0:
        return False
    if b["ESCOPO"] == ESC_INTRA:
        return True
    return b["ESCOPO"] == ESC_LOOSE and any(b[k] in ("APROVADA", "EXECUTADA") for k in ATV_STATUS)


def ler_painel(caminho):
    dados = {}
    for s in open(caminho, encoding="utf-8").read().split("\n"):
        m = re.match(r"\s*window\.(__\w+__)\s*=\s*(.*?);?\s*$", s)
        if m and m.group(2).startswith(("[", "{")):
            dados[m.group(1)] = json.loads(m.group(2))
    for k in ("__DADOS__", "__JUNTAS__", "__PENDLISTA__", "__ESCOPOS__"):
        if k not in dados:
            raise SystemExit("Painel sem a base %s" % k)
    data = None
    for r in dados["__DADOS__"]:
        if r.get("atualizado_em"):
            data = dt.datetime.fromisoformat(r["atualizado_em"].replace("Z", "+00:00"))
            break
    dados["_data"] = data
    return dados


def situacao(t):
    """t = [previsto, executado, aprovado, reprovado] somados da atividade na pasta."""
    if t is None:
        return "N/A"
    prev, exe, apr, rep = t
    if prev and apr >= prev:
        return "APROVADA"
    if rep:
        return "REPROVADA"
    if exe:
        return "EXECUTADA"
    return "PENDENTE"


def base_por_pasta(d):
    atv = collections.defaultdict(lambda: collections.defaultdict(lambda: [0, 0, 0, 0]))
    info = collections.defaultdict(lambda: {"sis": collections.Counter(), "sub": collections.Counter(),
                                            "forn": collections.Counter(), "data_th": None})
    for r in d["__DADOS__"]:
        p = r["pasta"]
        if not p.startswith(PREFIXO_PASTA):
            continue
        t = atv[p][r["atividade"]]
        for i, k in enumerate(("previsto", "executado", "aprovado", "reprovado")):
            t[i] += r[k] or 0
        inf = info[p]
        if "não cadastrado" not in r["sistema"]:
            inf["sis"][r["sistema"]] += 1
        if "não cadastrado" not in r["subsistema"]:
            inf["sub"][r["subsistema"]] += 1
        if "não cadastrado" not in r["fornecedor"]:
            inf["forn"][r["fornecedor"]] += 1
        if r["atividade"] == ATV_TH and re.match(r"\d{4}-\d{2}-\d{2}", r["data_planejada"] or ""):
            inf["data_th"] = r["data_planejada"]

    juntas = collections.defaultdict(lambda: [0, 0, None, None])
    for j in d["__JUNTAS__"]:
        if j["tipo"] != "pasta" or not str(j["pasta"]).startswith(PREFIXO_PASTA):
            continue
        x = juntas[j["pasta"]]
        x[0] += j["previstas"] or 0
        x[1] += j["executadas"] or 0
        x[2] = x[2] or j["sistema"]
        x[3] = x[3] or j["subsistema"]

    escopo = {e["pasta"]: e["escopo"] for e in d["__ESCOPOS__"]}
    pend_a = collections.Counter()
    for p in d["__PENDLISTA__"]:
        if p["disciplina"] == "Tubulação" and str(p["categoria"]).startswith("A"):
            pend_a[p["subsistema"]] += 1

    base = []
    for p in sorted(set(atv) | set(juntas)):
        inf, jt, a = info[p], juntas.get(p), atv.get(p, {})
        sis = inf["sis"].most_common(1)[0][0] if inf["sis"] else (jt[2] if jt else "")
        sub = inf["sub"].most_common(1)[0][0] if inf["sub"] else (jt[3] if jt else "")
        base.append({
            "PASTA": p, "SISTEMA": sis, "SUBSISTEMA": sub, "ESCOPO": escopo.get(p, ""),
            "FORNECEDOR": inf["forn"].most_common(1)[0][0] if inf["forn"] else "",
            "JUNTAS PREVISTAS": jt[0] if jt else None, "JUNTAS LIBERADAS": jt[1] if jt else None,
            "MONTAGEM (LVM)": situacao(a.get(ATV_LVM)), "MAPA JUNTAS": situacao(a.get(ATV_MJ)),
            "TESTE HIDROSTÁTICO": situacao(a.get(ATV_TH)), "LAVAGEM": situacao(a.get(ATV_LAV)),
            "REMONTAGEM": situacao(a.get(ATV_REM)),
            "DATA PLANEJADA TH": dt.date.fromisoformat(inf["data_th"]) if inf["data_th"] else None,
            "PEND. A TUB. SUBSISTEMA": pend_a.get(sub, 0),
        })
    return base


def etapas_aprovadas(b):
    """Etapas com avanço aprovado na PCP para uma linha da base (mesma regra do status calculado)."""
    ok = lambda s: s in ("APROVADA", "N/A")
    intra_sem_ctb = regra_intra(b)
    return {
        "AVANÇO MONTAGEM": b["MONTAGEM (LVM)"] == "APROVADA",
        "AVANÇO TESTE": b["TESTE HIDROSTÁTICO"] == "APROVADA",
        "AVANÇO REMONTAGEM": b["REMONTAGEM"] == "APROVADA",
        "AVANÇO PRÉ WALK": intra_sem_ctb and all(b[k] in ("APROVADA", "EXECUTADA", "N/A")
                                                 for k in ("MAPA JUNTAS", "TESTE HIDROSTÁTICO", "LAVAGEM")),
    }


def _atv_base(nome):
    """'Remontagem (449.COM.FP.045)' -> 'REMONTAGEM'."""
    return re.sub(r"\s*\(.*$", "", str(nome or "")).strip().upper()


def ler_pcp(caminho):
    """{pasta: {atividade normalizada: data da execução ou None}} da aba PCP_ATIVIDADE."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb["PCP_ATIVIDADE"]
    linhas = ws.iter_rows(values_only=True)
    hdr = {_norm(h): i for i, h in enumerate(next(linhas))}
    for k in ("ATIVIDADE", "DATA DA EXECUÇÃO", "Nº DA PASTA", "COD IDENTIFICAÇÃO"):
        if k not in hdr:
            raise SystemExit("PCP_ATIVIDADE sem a coluna %s" % k)
    pcp = collections.defaultdict(dict)
    for r in linhas:
        p = r[hdr["Nº DA PASTA"]]
        if not str(p or "").startswith(PREFIXO_PASTA):
            p = r[hdr["COD IDENTIFICAÇÃO"]]
        if not str(p or "").startswith(PREFIXO_PASTA):
            continue
        data = r[hdr["DATA DA EXECUÇÃO"]]
        if isinstance(data, str):
            try:
                data = dt.datetime.strptime(data.strip()[:10], "%d/%m/%Y")
            except ValueError:
                data = None
        atv = _atv_base(r[hdr["ATIVIDADE"]])
        anterior = pcp[str(p).strip()].get(atv)
        if atv not in pcp[str(p).strip()] or (data and (anterior is None or data > anterior)):
            pcp[str(p).strip()][atv] = data if isinstance(data, dt.datetime) else None
    wb.close()
    return pcp


def avancos_pcp(pcp, base, cal):
    """{pasta: {coluna AVANÇO: 'S-xxx' ou None}} pela DATA DA EXECUÇÃO."""
    info = {b["PASTA"]: b for b in base}
    sem = lambda d: txt_semana(semana_da_data(d, cal)) if d else None
    mapa = {"AVANÇO MONTAGEM": _atv_base(ATV_LVM), "AVANÇO TESTE": _atv_base(ATV_TH),
            "AVANÇO REMONTAGEM": _atv_base(ATV_REM)}
    exig_pre = [_atv_base(a) for a in (ATV_MJ, ATV_TH, ATV_LAV)]
    out = {}
    for p, atvs in pcp.items():
        o = {col: sem(atvs.get(a)) for col, a in mapa.items()}
        b = info.get(p)
        o["AVANÇO PRÉ WALK"] = None
        if b and regra_intra(b):
            datas = [atvs[a] for a in exig_pre if a in atvs]
            if datas and all(datas):
                o["AVANÇO PRÉ WALK"] = sem(max(datas))
        out[p] = o
    return out


def ler_calendario(caminho_ou_wb):
    wb = caminho_ou_wb if not isinstance(caminho_ou_wb, str) else openpyxl.load_workbook(caminho_ou_wb, data_only=True)
    ws = wb["CALENDÁRIO"] if "CALENDÁRIO" in wb.sheetnames else wb.worksheets[0]
    cal = {}
    for data, sem in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        if isinstance(data, dt.datetime) and sem:
            cal[data.date()] = sem
    if not cal:
        raise SystemExit("Calendário vazio")
    return cal


def semana_da_data(data, cal):
    """Número da semana do calendário da obra; fora dele, continua a numeração semanal (segunda a domingo)."""
    if isinstance(data, dt.datetime):
        data = data.date()
    if data in cal:
        return num_semana(cal[data])
    ref = min(cal)
    while ref.weekday() != 0:  # primeira segunda-feira do calendário
        ref += dt.timedelta(days=1)
    return num_semana(cal[ref]) + (data - ref).days // 7


def num_semana(v):
    """'SEM 62', 's-062', 'S-62', 62 -> 62."""
    if v is None:
        return None
    m = re.search(r"(\d+)", str(v))
    return int(m.group(1)) if m else None


def txt_semana(n):
    return "S-%03d" % n


def _norm(h):
    t = re.sub(r"\s+", " ", str(h or "")).strip().upper()
    # variações de "Nº": N°, Nº, N.º, Nro, NO., N. -> Nº
    t = re.sub(r"^N\s*(?:[°º˚O]|\.\s*[°º]?|RO\.?)?\s*(?=DA |DE |DO )", "Nº ", t)
    return t


# Colunas do CONTROLE (layout da REV1). tipo: inp = manual, auto = fórmula, pcp = semana de avanço
# gravada pelo script, aux = auxiliar do painel. grupo define a cor do cabeçalho.
CTRL = [
    ("ITEM", "auto", 6), ("Nº DA PASTA", "inp", 16), ("SUBSOP HMS", "inp", 20), ("SISTEMA", "auto", 16),
    ("ESCOPO", "auto", 17), ("FORNECEDOR", "auto", 11), ("JUNTAS PREVISTAS CTB", "auto", 9),
    ("JUNTAS LIBERADAS CTB", "auto", 9), ("SALDO DE JUNTAS", "auto", 8), ("AVANÇO CTB", "auto", 10),
    ("MONTAGEM HMS (LVM)", "auto", 11), ("MAPA JUNTAS HMS", "auto", 11), ("CERTIFICADO TESTE HMSWEB", "auto", 10),
    ("LAVAGEM HMS", "auto", 11), ("REMONTAGEM HMS", "auto", 11), ("DATA PLANEJADA TH", "auto", 11),
    ("PEND. A TUB. SUBSISTEMA", "auto", 9), ("STATUS CALCULADO", "auto", 22),
    ("CATEGORIA DE ENTREGA", "prog", 15), ("SEMANA PROG MONTAGEM", "prog", 11), ("SEMANA PROG TESTE", "prog", 10),
    ("SEMANA PROG REMONTAGEM", "prog", 12), ("SEMANA PROG PRÉ WALK", "prog", 11), ("STATUS PROGRAMAÇÃO", "progauto", 16),
    ("DATA CONCLUSÃO INFORMADO RESPONSAVEL", "prog", 12), ("RESPONSÁVEL", "prog", 14), ("OBSERVAÇÃO", "prog", 30),
    ("RESTRIÇÕES ABERTAS", "prog", 18),
    ("AVANÇO MONTAGEM", "pcp", 11), ("AVANÇO TESTE", "pcp", 10), ("AVANÇO REMONTAGEM", "pcp", 12),
    ("AVANÇO PRÉ WALK", "pcp", 11), ("PROGRAMAÇÃO DA SEMANA", "pcpauto", 13),
    ("PONTOS TESTE", "aux", 9), ("PONTOS REMONT.", "aux", 9), ("PONTOS PRÉ WALK", "aux", 9), ("ORDEM PROG.", "aux", 9),
]
C = {nome: get_column_letter(i) for i, (nome, _, _) in enumerate(CTRL, 1)}
MANUAIS = [n for n, t, _ in CTRL if t in ("inp", "prog")]
AVANCOS = ["AVANÇO MONTAGEM", "AVANÇO TESTE", "AVANÇO REMONTAGEM", "AVANÇO PRÉ WALK"]
LEGADO = {"PROGRAMAÇÃO": "CATEGORIA DE ENTREGA", "PRIORIDADE": "CATEGORIA DE ENTREGA",
          "Nº DA PASTA - PARA": "Nº DA PASTA", "DATA ENVIO": None}

LIN_CAB = 4
MAX_LIN = 600
MAX_RESTR = 500


def ler_controle(caminho):
    """Devolve (linhas por pasta com os campos manuais e de avanço, restrições, workbook)."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    linhas, restr = [], []
    if "CONTROLE" in wb.sheetnames:
        ws, lin_cab = wb["CONTROLE"], LIN_CAB
    else:  # formato original 01-CONTROLE_PASTA.xlsx
        ws = wb["CONTROLE DE PASTA"] if "CONTROLE DE PASTA" in wb.sheetnames else wb.worksheets[0]
        lin_cab = 1
    hdr = {}
    for c in ws[lin_cab]:
        h = _norm(c.value)
        h = LEGADO.get(h, h)
        if h and h not in hdr:
            hdr[h] = c.column
    campos = [n for n in MANUAIS + AVANCOS if _norm(n) in hdr]
    for r in range(lin_cab + 1, ws.max_row + 1):
        p = ws.cell(r, hdr["Nº DA PASTA"]).value
        if not p or "NÃO CADASTRADA" in str(p).upper():
            continue
        d = {k: ws.cell(r, hdr[_norm(k)]).value for k in campos}
        d["Nº DA PASTA"] = str(p).strip()
        linhas.append(d)
    if "RESTRIÇÕES" in wb.sheetnames:
        wr = wb["RESTRIÇÕES"]
        for r in range(LIN_CAB + 1, wr.max_row + 1):
            vals = [wr.cell(r, c).value for c in range(1, len(COL_RESTR) + 1)]
            if vals[1]:
                restr.append(vals)
    return linhas, restr, wb


# ----------------------------------------------------------------------------- estilo
FONTE = "Arial"
AZUL = "1F3864"
F_TIT = Font(name=FONTE, size=16, bold=True, color=AZUL)
F_SUB = Font(name=FONTE, size=9, italic=True, color="595959")
F_CAB = Font(name=FONTE, size=9, bold=True, color="FFFFFF")
F_TXT = Font(name=FONTE, size=9)
F_INP = Font(name=FONTE, size=9, color="0000FF")
F_KPI = Font(name=FONTE, size=20, bold=True, color=AZUL)
F_KPI_L = Font(name=FONTE, size=8, bold=True, color="595959")
F_H2 = Font(name=FONTE, size=11, bold=True, color=AZUL)
FILL_CAB = PatternFill("solid", fgColor=AZUL)
FILL_CAB_AUTO = PatternFill("solid", fgColor="2E75B6")
FILL_CAB_INP = PatternFill("solid", fgColor="BF8F00")
FILL_CAB_PROG = PatternFill("solid", fgColor="C0504D")
FILL_CAB_AV = PatternFill("solid", fgColor="92D050")
FILL_CAB_AUX = PatternFill("solid", fgColor="808080")
FILL_INP = PatternFill("solid", fgColor="FFF2CC")
FILL_KPI = PatternFill("solid", fgColor="EEF3FA")
FILL_ZEBRA = PatternFill("solid", fgColor="F7F9FC")
FINO = Side(style="thin", color="D0D7E2")
BORDA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
ESQ = Alignment(horizontal="left", vertical="center")

STATUS_LISTA = ["EM MONTAGEM", "P/TESTE", "P/REMONTAGEM", "QUALIDADE-REMONTAGEM", "PRÉ WALKDOWN", "PENDENTE DOCUMENTAÇÃO HMSWEB",
                "CONCLUÍDA", "SEM CTB", "SEM ESCOPO", "NÃO ENCONTRADA"]
COR_STATUS = {"EM MONTAGEM": "FFEB9C", "P/TESTE": "C6EFCE", "P/REMONTAGEM": "DDEBF7", "QUALIDADE-REMONTAGEM": "D9C3E9", "PRÉ WALKDOWN": "B4C6E7",
              "PENDENTE DOCUMENTAÇÃO HMSWEB": "F4B183", "CONCLUÍDA": "A9D08E", "SEM CTB": "EDEDED",
              "SEM ESCOPO": "D9D9D9", "NÃO ENCONTRADA": "F8CBAD"}
# etapa da programação -> (coluna da semana programada, status, coluna de avanço)
ETAPAS = [("MONTAGEM", "SEMANA PROG MONTAGEM", "EM MONTAGEM", "AVANÇO MONTAGEM"),
          ("TESTE", "SEMANA PROG TESTE", "P/TESTE", "AVANÇO TESTE"),
          ("REMONTAGEM", "SEMANA PROG REMONTAGEM", "P/REMONTAGEM", "AVANÇO REMONTAGEM"),
          ("PRÉ WALK", "SEMANA PROG PRÉ WALK", "PRÉ WALKDOWN", "AVANÇO PRÉ WALK")]
CATEG_LISTA = ["PRIORIDADE 1", "PRIORIDADE 2", "OPORTUNIDADE"]
TIPOS_RESTR = ["Material", "Suporte", "Junta reprovada / END", "Pendência A (punch)", "Documentação / Isométrico",
               "Liberação de área / PT", "Andaime / Acesso", "Equipamento de teste", "Spool / Pré-fabricação",
               "Projeto / Engenharia", "Outros"]
STATUS_RESTR = ["ABERTA", "EM TRATAMENTO", "ENCERRADA"]
COL_RESTR = ["ID", "Nº DA PASTA", "SUBSOP (auto)", "TIPO", "DESCRIÇÃO", "RESPONSÁVEL", "DATA ABERTURA",
             "DATA PREVISTA", "STATUS", "DATA ENCERRAMENTO", "DIAS EM ABERTO", "SITUAÇÃO PRAZO"]


def cab(ws, linha, colunas, fill=FILL_CAB, altura=36):
    for i, nome in enumerate(colunas, 1):
        c = ws.cell(linha, i, nome)
        c.font, c.fill, c.alignment, c.border = F_CAB, fill, CENTRO, BORDA
    ws.row_dimensions[linha].height = altura


def titulo(ws, texto, sub):
    ws["A1"] = texto
    ws["A1"].font = F_TIT
    ws["A2"] = sub
    ws["A2"].font = F_SUB
    ws.sheet_view.showGridLines = False


def wk(ref):
    """Número da semana digitada ('s-062', 'S-62', 'SEM 62', 62); -1 se vazio ou inválido."""
    return ('IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(UPPER(TRIM(%s)),"SEM",""),"S",""),"-","")),-1)' % ref)


# ----------------------------------------------------------------------------- abas
def aba_calendario(wb, cal):
    ws = wb.create_sheet("CALENDÁRIO")
    cab(ws, 1, ["DATA", "SEMANA"], altura=20)
    for i, (d, s) in enumerate(sorted(cal.items()), 2):
        ws.cell(i, 1, d).number_format = "dd/mm/yyyy"
        ws.cell(i, 2, s)
        ws.cell(i, 1).font = ws.cell(i, 2).font = F_TXT
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    ws.freeze_panes = "A2"
    return len(cal) + 1


def aba_parametros(wb, ult_cal):
    ws = wb.create_sheet("PARÂMETROS")
    titulo(ws, "PARÂMETROS E LISTAS", "Células amarelas podem ser alteradas; todas as abas usam estes valores.")
    itens = [
        (4, "Avanço mínimo do ControlTub para liberar teste", 0.8, "0%",
         "Pasta RADIX/FURUI com avanço do CTB igual ou superior vai para P/TESTE."),
        (5, "Data de referência (hoje)", "=TODAY()", "dd/mm/yyyy", ""),
        (6, "Semana da data de referência (calendário da obra)",
         "=IFERROR(VALUE(SUBSTITUTE(INDEX('CALENDÁRIO'!$B$2:$B$%d,MATCH($B$5,'CALENDÁRIO'!$A$2:$A$%d,1)),\"SEM\",\"\")),\"\")"
         % (ult_cal, ult_cal), "0", "Aba CALENDÁRIO (03.CALENDÁRIO SEMANAL)."),
        (7, "Semana vigente informada (opcional)", None, "General",
         "Preencha (ex.: S-063) para analisar outra semana; vazio = semana da data de hoje."),
        (8, "SEMANA VIGENTE (usada no controle)", '=IF(B7<>"",%s,B6)' % wk("B7"), "0", ""),
        (9, "Semana vigente (texto)", '="S-"&TEXT(B8,"000")', "General", ""),
    ]
    for r, rot, val, fmt, nota in itens:
        ws.cell(r, 1, rot).font = Font(name=FONTE, size=9, bold=True)
        c = ws.cell(r, 2, val)
        c.number_format, c.font = fmt, F_TXT
        ws.cell(r, 3, nota).font = F_SUB
    for r in (4, 7):
        ws.cell(r, 2).fill, ws.cell(r, 2).font = FILL_INP, F_INP
    ws["B8"].font = ws["B9"].font = Font(name=FONTE, size=11, bold=True, color="C00000")

    listas = [("STATUS", STATUS_LISTA), ("CATEGORIA", CATEG_LISTA), ("TIPO RESTRIÇÃO", TIPOS_RESTR),
              ("STATUS RESTRIÇÃO", STATUS_RESTR)]
    for j, (nome, itens_l) in enumerate(listas):
        col = 5 + j
        c = ws.cell(12, col, nome)
        c.font, c.fill, c.alignment = F_CAB, FILL_CAB, CENTRO
        for i, v in enumerate(itens_l):
            ws.cell(13 + i, col, v).font = F_TXT
    ws["A12"] = "REGRAS"
    ws["A12"].font = F_H2
    regras = [
        "STATUS CALCULADO - escopo RADIX e FURUI (e INTRA SKID / LOOSE ITEM com juntas no ControlTub):",
        "  EM MONTAGEM: avanço ControlTub < avanço mínimo (B4), mesmo com teste aprovado no HMSWEB.",
        "  P/TESTE: avanço ControlTub >= avanço mínimo e sem certificado de teste aprovado no HMSWEB.",
        "  P/REMONTAGEM: teste aprovado no HMSWEB e remontagem pendente (ou reprovada).",
        "  QUALIDADE-REMONTAGEM: teste aprovado e remontagem executada (aguardando aprovação) ou aprovada",
        "    com outro certificado ainda pendente.",
        "  CONCLUÍDA: todas as atividades aprovadas (LV, mapa de juntas, teste, lavagem e remontagem; N/A não conta).",
        "  SEM CTB: sem juntas no ControlTub e sem teste aprovado.",
        "STATUS CALCULADO - escopo FURUI - INTRA SKID (sem juntas no ControlTub):",
        "  PENDENTE DOCUMENTAÇÃO HMSWEB: mapa de juntas, teste ou lavagem PENDENTE (ou reprovado).",
        "  PRÉ WALKDOWN: mapa de juntas, teste e lavagem EXECUTADOS ou APROVADOS; LV e/ou remontagem não aprovadas.",
        "  CONCLUÍDA: todas as atividades aprovadas.",
        "  Atividade inexistente na PCP para a pasta (N/A) não é exigida.",
        "STATUS CALCULADO - escopo FURUI - LOOSE ITEM sem juntas no ControlTub:",
        "  sem nenhum certificado executado/aprovado no HMSWEB -> EM MONTAGEM;",
        "  com algum certificado executado/aprovado -> mesmo critério da INTRA SKID.",
        "OUTROS: SEM ESCOPO (fora da lista de escopo Milplan) e NÃO ENCONTRADA (fora do Painel OLODUM).",
        "",
        "STATUS PROGRAMAÇÃO: etapa cuja SEMANA PROG é a semana vigente (B8).",
        "  MONTAGEM -> EM MONTAGEM | TESTE -> P/TESTE | REMONTAGEM -> P/REMONTAGEM | PRÉ WALK -> PRÉ WALKDOWN.",
        "  Se mais de uma etapa estiver na semana vigente, vale a mais adiantada.",
        "",
        "AVANÇO (MONTAGEM/TESTE/REMONTAGEM/PRÉ WALK): semana do calendário da DATA DA EXECUÇÃO na PCP_ATIVIDADE",
        "  (OLODUM_BASEDADOS): LV de montagem / teste hidrostático / remontagem. PRÉ WALK (INTRA SKID): última",
        "  execução entre mapa de juntas, teste hidrostático e lavagem, quando todas estiverem executadas.",
        "PROGRAMAÇÃO DA SEMANA: REALIZADA quando o avanço da etapa programada caiu na semana vigente.",
    ]
    for i, t in enumerate(regras):
        ws.cell(13 + i, 1, t).font = F_TXT
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 4
    for col in "EFGH":
        ws.column_dimensions[col].width = 26
    n = {k: len(v) for k, v in listas}
    return {
        "min": "'PARÂMETROS'!$B$4", "hoje": "'PARÂMETROS'!$B$5", "sv": "'PARÂMETROS'!$B$8",
        "svt": "'PARÂMETROS'!$B$9",
        "status": "'PARÂMETROS'!$E$13:$E$%d" % (12 + n["STATUS"]),
        "categ": "'PARÂMETROS'!$F$13:$F$%d" % (12 + n["CATEGORIA"]),
        "tipo": "'PARÂMETROS'!$G$13:$G$%d" % (12 + n["TIPO RESTRIÇÃO"]),
        "sres": "'PARÂMETROS'!$H$13:$H$%d" % (12 + n["STATUS RESTRIÇÃO"]),
    }


BASE_COLS = ["PASTA", "SISTEMA", "SUBSISTEMA", "ESCOPO", "FORNECEDOR", "JUNTAS PREVISTAS", "JUNTAS LIBERADAS",
             "MONTAGEM (LVM)", "MAPA JUNTAS", "TESTE HIDROSTÁTICO", "LAVAGEM", "REMONTAGEM", "DATA PLANEJADA TH",
             "PEND. A TUB. SUBSISTEMA"]


def aba_base(wb, base, gerado_em):
    ws = wb.create_sheet("BASE_OLODUM")
    titulo(ws, "BASE PAINEL OLODUM (por pasta)",
           "Importada do Painel OLODUM em %s. Juntas = ControlTub (liberadas/previstas); atividades = PCP/HMSWEB "
           "(APROVADA, EXECUTADA, REPROVADA, PENDENTE, N/A). Não editar: é sobrescrita a cada atualização." % gerado_em)
    cab(ws, LIN_CAB, BASE_COLS)
    for i, b in enumerate(base, LIN_CAB + 1):
        for j, k in enumerate(BASE_COLS, 1):
            c = ws.cell(i, j, b[k])
            c.font, c.border = F_TXT, BORDA
            if k == "DATA PLANEJADA TH":
                c.number_format = "dd/mm/yyyy"
    for j, w in enumerate([17, 18, 22, 18, 14, 11, 11, 13, 13, 13, 13, 13, 12, 12], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(LIN_CAB + 1, 2)
    ws.auto_filter.ref = "A%d:%s%d" % (LIN_CAB, get_column_letter(len(BASE_COLS)), LIN_CAB + len(base))
    return LIN_CAB + len(base)


def _ok_exec(ref):
    return 'OR(%s="APROVADA",%s="EXECUTADA",%s="N/A")' % (ref, ref, ref)


def f_controle(r, P, nb, lin_cab=LIN_CAB):
    b = lambda col: "'BASE_OLODUM'!$%s:$%s" % (col, col)  # coluna inteira: a base pode crescer
    pasta = "$%s%d" % (C["Nº DA PASTA"], r)
    m = "MATCH(%s,%s,0)" % (pasta, b("A"))
    x = lambda nome: "%s%d" % (C[nome], r)

    def look(col, texto=True):
        return '=IF(%s="","",IFERROR(INDEX(%s,%s)%s,""))' % (pasta, b(col), m, '&""' if texto else "")

    def look_num(col):
        return '=IF(%s="","",IFERROR(IF(INDEX(%s,%s)="","",INDEX(%s,%s)),""))' % (pasta, b(col), m, b(col), m)

    sv = P["sv"]
    status_prog = '""'
    for _, col_sem, st, _ in ETAPAS:  # da menos para a mais adiantada: a última que casar prevalece
        status_prog = 'IF(%s=%s,"%s",%s)' % (wk(x(col_sem)), sv, st, status_prog)
    realizada = '""'
    for _, _, st, col_av in reversed(ETAPAS):
        realizada = 'IF(%s="%s",%s,%s)' % (x("STATUS PROGRAMAÇÃO"), st, x(col_av), realizada)
    rest = 'IF(%s="",5000,0)' % x("RESTRIÇÕES ABERTAS")
    cat = ('IF(OR(%(c)s="PRIORIDADE 1",%(c)s="PRIORIDADE"),3000,IF(%(c)s="PRIORIDADE 2",2000,'
           'IF(%(c)s="OPORTUNIDADE",1000,0)))' % {"c": x("CATEGORIA DE ENTREGA")})
    return {
        "ITEM": '=IF(%s="","",ROW()-%d)' % (pasta, lin_cab),
        "SISTEMA": look("B"), "ESCOPO": look("D"), "FORNECEDOR": look("E"),
        "JUNTAS PREVISTAS CTB": look_num("F"), "JUNTAS LIBERADAS CTB": look_num("G"),
        "SALDO DE JUNTAS": '=IF(ISNUMBER(%s),%s-%s,"")' % (x("JUNTAS PREVISTAS CTB"), x("JUNTAS PREVISTAS CTB"),
                                                           x("JUNTAS LIBERADAS CTB")),
        "AVANÇO CTB": '=IF(N(%s)>0,%s/%s,"")' % (x("JUNTAS PREVISTAS CTB"), x("JUNTAS LIBERADAS CTB"),
                                               x("JUNTAS PREVISTAS CTB")),
        "MONTAGEM HMS (LVM)": look("H"), "MAPA JUNTAS HMS": look("I"),
        "CERTIFICADO TESTE HMSWEB": look("J"),  # mesmo critério das demais: PENDENTE/EXECUTADA/APROVADA/...
        "LAVAGEM HMS": look("K"), "REMONTAGEM HMS": look("L"), "DATA PLANEJADA TH": look_num("M"),
        "PEND. A TUB. SUBSISTEMA": look("N", texto=False),
        "STATUS CALCULADO": (
            '=IF(%(p)s="","",IF(ISNA(%(m)s),"NÃO ENCONTRADA",'
            # critério RADIX/FURUI: RADIX, FURUI (exceto INTRA SKID e LOOSE ITEM) e INTRA/LOOSE com juntas no CTB
            'IF(OR(%(e)s="%(radix)s",AND(LEFT(%(e)s,5)="%(furui)s",%(e)s<>"%(intra)s",%(e)s<>"%(loose)s"),'
            'AND(OR(%(e)s="%(intra)s",%(e)s="%(loose)s"),N(%(prev)s)>0)),%(padrao)s,'
            # INTRA SKID sem juntas no CTB
            'IF(%(e)s="%(intra)s",%(intrab)s,'
            # LOOSE ITEM sem juntas no CTB: sem avanço de certificado = EM MONTAGEM; com avanço = critério INTRA SKID
            'IF(%(e)s="%(loose)s",IF(COUNTIF(%(k)s:%(o)s,"APROVADA")+COUNTIF(%(k)s:%(o)s,"EXECUTADA")>0,'
            '%(intrab)s,"EM MONTAGEM"),'
            '"SEM ESCOPO")))))'
            % {"p": pasta, "m": m, "e": x("ESCOPO"), "radix": ESC_RADIX, "furui": ESC_FURUI, "intra": ESC_INTRA,
               "loose": ESC_LOOSE, "prev": x("JUNTAS PREVISTAS CTB"),
               "k": x("MONTAGEM HMS (LVM)"), "o": x("REMONTAGEM HMS"),
               "padrao": ('IF(AND(N(%(prev)s)>0,N(%(av)s)<%(min)s),"EM MONTAGEM",'
                          # teste aprovado: todas aprovadas (ou N/A) = CONCLUÍDA; remontagem executada ou
                          # aprovada com outro certificado pendente = QUALIDADE-REMONTAGEM; senão P/REMONTAGEM
                          'IF(%(th)s="APROVADA",IF(COUNTIF(%(k)s:%(o)s,"APROVADA")+COUNTIF(%(k)s:%(o)s,"N/A")=5,'
                          '"CONCLUÍDA",IF(OR(%(rem)s="EXECUTADA",%(rem)s="APROVADA"),"QUALIDADE-REMONTAGEM",'
                          '"P/REMONTAGEM")),'
                          'IF(%(av)s="","SEM CTB","P/TESTE")))'
                          % {"prev": x("JUNTAS PREVISTAS CTB"), "av": x("AVANÇO CTB"), "min": P["min"],
                             "th": x("CERTIFICADO TESTE HMSWEB"), "rem": x("REMONTAGEM HMS"),
                             "k": x("MONTAGEM HMS (LVM)"), "o": x("REMONTAGEM HMS")}),
               # mapa de juntas, teste e lavagem executados ou aprovados (ou N/A) = PRÉ WALKDOWN;
               # tudo aprovado = CONCLUÍDA; algum dos três pendente/reprovado = PENDENTE DOCUMENTAÇÃO HMSWEB
               "intrab": ('IF(AND(%(ok_mj)s,%(ok_th)s,%(ok_lav)s),IF(COUNTIF(%(k)s:%(o)s,"APROVADA")'
                          '+COUNTIF(%(k)s:%(o)s,"N/A")=5,"CONCLUÍDA","PRÉ WALKDOWN"),"PENDENTE DOCUMENTAÇÃO HMSWEB")'
                          % {"ok_mj": _ok_exec(x("MAPA JUNTAS HMS")), "ok_th": _ok_exec(x("CERTIFICADO TESTE HMSWEB")),
                             "ok_lav": _ok_exec(x("LAVAGEM HMS")), "k": x("MONTAGEM HMS (LVM)"),
                             "o": x("REMONTAGEM HMS")})}),
        "STATUS PROGRAMAÇÃO": '=IF(%s="","",%s)' % (pasta, status_prog),
        "PROGRAMAÇÃO DA SEMANA": '=IF(%s="","",IF(%s=%s,"REALIZADA","PENDENTE"))'
                                 % (x("STATUS PROGRAMAÇÃO"), realizada, P["svt"]),
        "PONTOS TESTE": '=IF(%s<>"P/TESTE",0,%s+%s+N(%s)*100+ROW()/1000000)'
                        % (x("STATUS CALCULADO"), rest, cat, x("AVANÇO CTB")),
        "PONTOS REMONT.": '=IF(%s<>"P/REMONTAGEM",0,%s+%s+ROW()/1000000)' % (x("STATUS CALCULADO"), rest, cat),
        "PONTOS PRÉ WALK": '=IF(%s<>"PRÉ WALKDOWN",0,%s+%s+ROW()/1000000)' % (x("STATUS CALCULADO"), rest, cat),
        "ORDEM PROG.": '=IF(AND(%s<>"",%s<>""),ROW(),"")' % (pasta, x("STATUS PROGRAMAÇÃO")),
    }


def aba_controle(wb, linhas, P, nb):
    ws = wb.create_sheet("CONTROLE")
    titulo(ws, "CONTROLE DE PASTAS DE TUBULAÇÃO - SALA DE GUERRA",
           "Fundo amarelo = preenchimento manual. Azul = Painel OLODUM / ControlTub / HMSWEB. Verde = semana em que a "
           "etapa foi executada na PCP (DATA DA EXECUÇÃO). Para incluir uma pasta, digite o número na coluna B.")
    ws["A3"] = ('="Semana vigente: "&%s&"   |   Exemplo: CATEGORIA = PRIORIDADE | SEMANA PROG TESTE = S-062 '
                '| DATA CONCLUSÃO = 20/11/2026 | RESTRIÇÕES ABERTAS = falta suporte S-12"' % P["svt"])
    ws["A3"].font = Font(name=FONTE, size=9, bold=True, color="C00000")
    fills = {"inp": FILL_CAB_INP, "auto": FILL_CAB_AUTO, "prog": FILL_CAB_PROG, "progauto": FILL_CAB_PROG,
             "pcp": FILL_CAB_AV, "pcpauto": FILL_CAB_AV, "aux": FILL_CAB_AUX}
    for i, (nome, tipo, larg) in enumerate(CTRL, 1):
        c = ws.cell(LIN_CAB, i, nome)
        c.alignment, c.border, c.fill = CENTRO, BORDA, fills[tipo]
        c.font = Font(name=FONTE, size=9, bold=True, color="FF0000") if tipo.startswith("pcp") else F_CAB
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[LIN_CAB].height = 52

    ultima = LIN_CAB + max(MAX_LIN, len(linhas) + 50)
    for idx in range(LIN_CAB + 1, ultima + 1):
        dados = linhas[idx - LIN_CAB - 1] if idx - LIN_CAB - 1 < len(linhas) else {}
        f = f_controle(idx, P, nb)
        for i, (nome, tipo, _) in enumerate(CTRL, 1):
            c = ws.cell(idx, i)
            c.border = BORDA
            if tipo in ("inp", "prog"):
                c.value = dados.get(nome)
                c.font, c.fill = F_INP, FILL_INP
            elif tipo == "pcp":
                c.value = dados.get(nome)
                c.font = Font(name=FONTE, size=9, bold=True, color="375623")
            else:
                c.value = f[nome]
                c.font = F_TXT if tipo != "aux" else Font(name=FONTE, size=8, color="808080")
            if nome == "AVANÇO CTB":
                c.number_format = "0%"
            elif nome in ("DATA CONCLUSÃO INFORMADO RESPONSAVEL", "DATA PLANEJADA TH"):
                c.number_format = "dd/mm/yyyy"
            elif tipo == "aux":
                c.number_format = "0.00"
            if nome not in ("OBSERVAÇÃO", "Nº DA PASTA", "SUBSOP HMS", "SISTEMA", "RESPONSÁVEL", "RESTRIÇÕES ABERTAS"):
                c.alignment = Alignment(horizontal="center", vertical="center")
    rng = lambda nome: "%s%d:%s%d" % (C[nome], LIN_CAB + 1, C[nome], ultima)

    dv = DataValidation(type="list", formula1="=" + P["categ"], allow_blank=True)
    dv.error, dv.errorTitle = "Escolha PRIORIDADE ou OPORTUNIDADE.", "Valor inválido"
    ws.add_data_validation(dv)
    dv.add(rng("CATEGORIA DE ENTREGA"))
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
    dv.error = "Informe uma data (dd/mm/aaaa)."
    ws.add_data_validation(dv)
    dv.add(rng("DATA CONCLUSÃO INFORMADO RESPONSAVEL"))

    for nome in ("STATUS CALCULADO", "STATUS PROGRAMAÇÃO"):
        for st, cor in COR_STATUS.items():
            ws.conditional_formatting.add(rng(nome), CellIsRule(operator="equal", formula=['"%s"' % st],
                                          fill=PatternFill("solid", fgColor=cor), font=Font(name=FONTE, size=9, bold=True)))
    # semana programada = semana vigente: destaca
    for _, col_sem, _, _ in ETAPAS:
        ws.conditional_formatting.add(rng(col_sem), FormulaRule(
            formula=["%s=%s" % (wk("%s%d" % (C[col_sem], LIN_CAB + 1)), P["sv"])],
            fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONTE, size=9, bold=True, color="C00000")))
    for col_av in AVANCOS:
        ws.conditional_formatting.add(rng(col_av), FormulaRule(
            formula=["%s%d=%s" % (C[col_av], LIN_CAB + 1, P["svt"])],
            fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(rng("PROGRAMAÇÃO DA SEMANA"), CellIsRule(operator="equal", formula=['"REALIZADA"'],
                                  fill=PatternFill("solid", fgColor="00B050"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(rng("PROGRAMAÇÃO DA SEMANA"), CellIsRule(operator="equal", formula=['"PENDENTE"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng("RESTRIÇÕES ABERTAS"), FormulaRule(
        formula=['%s%d<>""' % (C["RESTRIÇÕES ABERTAS"], LIN_CAB + 1)],
        fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    for nome in ("MONTAGEM HMS (LVM)", "MAPA JUNTAS HMS", "CERTIFICADO TESTE HMSWEB", "LAVAGEM HMS", "REMONTAGEM HMS"):
        ws.conditional_formatting.add(rng(nome), CellIsRule(operator="equal", formula=['"APROVADA"'],
                                      font=Font(name=FONTE, size=9, color="006100")))
        ws.conditional_formatting.add(rng(nome), CellIsRule(operator="equal", formula=['"REPROVADA"'],
                                      fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng("AVANÇO CTB"), DataBarRule(start_type="num", start_value=0, end_type="num",
                                  end_value=1, color="5B9BD5"))
    ws.freeze_panes = ws.cell(LIN_CAB + 1, 3)
    ws.auto_filter.ref = "A%d:%s%d" % (LIN_CAB, get_column_letter(len(CTRL)), ultima)
    return ultima


def aba_restricoes(wb, restr, P, ult_ctrl):
    ws = wb.create_sheet("RESTRIÇÕES")
    titulo(ws, "CONTROLE DE RESTRIÇÕES DAS PASTAS",
           "Registro detalhado (opcional) das restrições, uma linha por restrição. Colunas amarelas = manual. "
           "O resumo por tipo aparece no PAINEL.")
    tipos_col = ["auto", "inp", "auto", "inp", "inp", "inp", "inp", "inp", "inp", "inp", "auto", "auto"]
    for i, nome in enumerate(COL_RESTR, 1):
        c = ws.cell(LIN_CAB, i, nome)
        c.font, c.alignment, c.border = F_CAB, CENTRO, BORDA
        c.fill = FILL_CAB_INP if tipos_col[i - 1] == "inp" else FILL_CAB_AUTO
    ws.row_dimensions[LIN_CAB].height = 36
    ultima = LIN_CAB + max(MAX_RESTR, len(restr) + 50)
    ctrl_p = "'CONTROLE'!$%s$%d:$%s$%d" % (C["Nº DA PASTA"], LIN_CAB + 1, C["Nº DA PASTA"], ult_ctrl)
    ctrl_s = "'CONTROLE'!$%s$%d:$%s$%d" % (C["SUBSOP HMS"], LIN_CAB + 1, C["SUBSOP HMS"], ult_ctrl)
    for r in range(LIN_CAB + 1, ultima + 1):
        vals = restr[r - LIN_CAB - 1] if r - LIN_CAB - 1 < len(restr) else [None] * len(COL_RESTR)
        f = {
            1: '=IF(B%d="","","R-"&TEXT(ROW()-%d,"000"))' % (r, LIN_CAB),
            3: '=IF(B%d="","",IFERROR(INDEX(%s,MATCH(B%d,%s,0)),"PASTA FORA DO CONTROLE"))' % (r, ctrl_s, r, ctrl_p),
            11: '=IF(OR(B%d="",G%d=""),"",IF(J%d<>"",J%d,%s)-G%d)' % (r, r, r, r, P["hoje"], r),
            12: ('=IF(B%d="","",IF(I%d="ENCERRADA","ENCERRADA",IF(H%d="","SEM PRAZO",IF(H%d<%s,"ATRASADA",'
                 '"NO PRAZO"))))' % (r, r, r, r, P["hoje"])),
        }
        for i in range(1, len(COL_RESTR) + 1):
            c = ws.cell(r, i)
            c.border = BORDA
            if i in f:
                c.value, c.font = f[i], F_TXT
            else:
                c.value, c.font, c.fill = vals[i - 1], F_INP, FILL_INP
            if i in (7, 8, 10):
                c.number_format = "dd/mm/yyyy"
            c.alignment = Alignment(horizontal="left" if i == 5 else "center", vertical="center", wrap_text=(i == 5))
    for i, w in enumerate([8, 16, 22, 22, 50, 16, 12, 12, 14, 13, 10, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    rng = lambda col: "%s%d:%s%d" % (col, LIN_CAB + 1, col, ultima)
    for col, lista in (("D", P["tipo"]), ("I", P["sres"])):
        dv = DataValidation(type="list", formula1="=" + lista, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(rng(col))
    for col in "GHJ":
        dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(rng(col))
    for st, cor, fc in (("ATRASADA", "FFC7CE", "9C0006"), ("NO PRAZO", "C6EFCE", "006100"), ("ENCERRADA", "EDEDED", "595959")):
        ws.conditional_formatting.add(rng("L"), CellIsRule(operator="equal", formula=['"%s"' % st],
                                      fill=PatternFill("solid", fgColor=cor), font=Font(name=FONTE, size=9, bold=True, color=fc)))
    ws.freeze_panes = ws.cell(LIN_CAB + 1, 3)
    ws.auto_filter.ref = "A%d:L%d" % (LIN_CAB, ultima)
    return ultima


def aba_programacao(wb, ult_ctrl, P):
    ws = wb.create_sheet("PROGRAMAÇÃO")
    titulo(ws, "PROGRAMAÇÃO DA SEMANA VIGENTE",
           "Pastas com alguma SEMANA PROG igual à semana vigente (aba PARÂMETROS). Lista automática na ordem do "
           "CONTROLE: não editar aqui.")
    ws["F1"] = '="SEMANA "&%s' % P["svt"]
    ws["F1"].font = Font(name=FONTE, size=16, bold=True, color="C00000")
    cols = [("#", None, 5), ("Nº DA PASTA", "Nº DA PASTA", 17), ("SUBSOP HMS", "SUBSOP HMS", 21),
            ("ESCOPO", "ESCOPO", 17), ("CATEGORIA", "CATEGORIA DE ENTREGA", 13),
            ("STATUS PROGRAMAÇÃO", "STATUS PROGRAMAÇÃO", 15), ("STATUS CALCULADO", "STATUS CALCULADO", 22),
            ("PROG. x CALC.", None, 12), ("AVANÇO CTB", "AVANÇO CTB", 9),
            ("PROGRAMAÇÃO DA SEMANA", "PROGRAMAÇÃO DA SEMANA", 12),
            ("AVANÇO MONTAGEM", "AVANÇO MONTAGEM", 11), ("AVANÇO TESTE", "AVANÇO TESTE", 10),
            ("AVANÇO REMONTAGEM", "AVANÇO REMONTAGEM", 12), ("AVANÇO PRÉ WALK", "AVANÇO PRÉ WALK", 11),
            ("DATA CONCLUSÃO", "DATA CONCLUSÃO INFORMADO RESPONSAVEL", 11), ("RESPONSÁVEL", "RESPONSÁVEL", 14),
            ("RESTRIÇÕES", "RESTRIÇÕES ABERTAS", 22), ("OBSERVAÇÃO", "OBSERVAÇÃO", 30)]
    cab(ws, LIN_CAB, [c[0] for c in cols], altura=40)
    ordem = "'CONTROLE'!$%s$%d:$%s$%d" % (C["ORDEM PROG."], LIN_CAB + 1, C["ORDEM PROG."], ult_ctrl)
    n = 300
    for k in range(1, n + 1):
        r = LIN_CAB + k
        ws.cell(r, 1, '=IF(COUNT(%s)>=%d,%d,"")' % (ordem, k, k))
        for j, (rot, s, _) in enumerate(cols[1:], 2):
            if s is None:  # PROG. x CALC.
                ws.cell(r, j, '=IF($A%d="","",IF(F%d=G%d,"CONFERE","DIFERENTE"))' % (r, r, r))
                continue
            idx = "INDEX('CONTROLE'!$%s:$%s,SMALL(%s,$A%d))" % (C[s], C[s], ordem, r)
            ws.cell(r, j, '=IF($A%d="","",IF(%s="","",%s))' % (r, idx, idx))
        for j, (rot, _, _) in enumerate(cols, 1):
            c = ws.cell(r, j)
            c.font, c.border = F_TXT, BORDA
            c.alignment = Alignment(horizontal="left" if rot in ("OBSERVAÇÃO", "RESTRIÇÕES") else "center",
                                    vertical="center")
            if rot == "AVANÇO CTB":
                c.number_format = "0%"
            if rot == "DATA CONCLUSÃO":
                c.number_format = "dd/mm/yyyy"
    rng = lambda col: "%s%d:%s%d" % (col, LIN_CAB + 1, col, LIN_CAB + n)
    for col in "FG":
        for st, cor in COR_STATUS.items():
            ws.conditional_formatting.add(rng(col), CellIsRule(operator="equal", formula=['"%s"' % st],
                                          fill=PatternFill("solid", fgColor=cor), font=Font(name=FONTE, size=9, bold=True)))
    ws.conditional_formatting.add(rng("H"), CellIsRule(operator="equal", formula=['"DIFERENTE"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng("J"), CellIsRule(operator="equal", formula=['"REALIZADA"'],
                                  fill=PatternFill("solid", fgColor="00B050"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(rng("J"), CellIsRule(operator="equal", formula=['"PENDENTE"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng("Q"), FormulaRule(formula=['Q%d<>""' % (LIN_CAB + 1)],
                                  fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.conditional_formatting.add(rng("I"), DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                        color="5B9BD5"))
    for j, (_, _, w) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = "A%d:%s%d" % (LIN_CAB, get_column_letter(len(cols)), LIN_CAB + n)


def aba_painel(wb, ult_ctrl, ult_restr, gerado_em, P):
    ws = wb.create_sheet("PAINEL", 0)
    ws.sheet_view.showGridLines = False
    ws["B1"] = "PAINEL GERENCIAL - SALA DE GUERRA DE PASTAS DE TUBULAÇÃO"
    ws["B1"].font = Font(name=FONTE, size=18, bold=True, color=AZUL)
    ws["B2"] = ('="Base Painel OLODUM: %s   |   Semana vigente: "&%s&"   |   Avanço mínimo p/ teste: "&'
                'TEXT(%s,"0%%")&"   |   Atualizado em "&%s' % (gerado_em, P["svt"], P["min"], _data_txt(P["hoje"])))
    ws["B2"].font = F_SUB
    ws.column_dimensions["A"].width = 2
    for col, w in zip("BCDEFGHIJKLMNOPQR", [5, 17, 20, 13, 10, 14, 16, 11, 3, 5, 17, 20, 13, 10, 14, 16, 3]):
        ws.column_dimensions[col].width = w

    R = lambda nome: "'CONTROLE'!$%s$%d:$%s$%d" % (C[nome], LIN_CAB + 1, C[nome], ult_ctrl)
    st, cat, pasta = R("STATUS CALCULADO"), R("CATEGORIA DE ENTREGA"), R("Nº DA PASTA")
    sp, ps = R("STATUS PROGRAMAÇÃO"), R("PROGRAMAÇÃO DA SEMANA")

    # ---------------- KPIs
    kpis = [
        ("PASTAS CONTROLADAS", '=COUNTA(%s)' % pasta, "0"),
        ("SEMANA VIGENTE", "=" + P["svt"], "General"),
        ("PROGRAMADAS NA SEMANA", '=COUNTIF(%s,"?*")' % sp, "0"),
        ("REALIZADAS NA SEMANA", '=COUNTIF(%s,"REALIZADA")' % ps, "0"),
        ("ADERÊNCIA DA PROGRAMAÇÃO", '=IF(F5=0,0,H5/F5)', "0%"),
        ("P/ TESTE (CALCULADO)", '=COUNTIF(%s,"P/TESTE")' % st, "0"),
        ("P/ REMONTAGEM (CALCULADO)", '=COUNTIF(%s,"P/REMONTAGEM")' % st, "0"),
        ("PRÉ WALKDOWN (INTRA SKID)", '=COUNTIF(%s,"PRÉ WALKDOWN")' % st, "0"),
    ]
    pos = ["B", "D", "F", "H", "K", "M", "O", "Q"]
    for (rot, f, fmt), col in zip(kpis, pos):
        c2 = get_column_letter(openpyxl.utils.column_index_from_string(col) + 1)
        ws.merge_cells("%s4:%s4" % (col, c2))
        ws.merge_cells("%s5:%s5" % (col, c2))
        ws["%s4" % col] = rot
        ws["%s5" % col] = f
        ws["%s5" % col].number_format = fmt
        ws["%s4" % col].font, ws["%s5" % col].font = F_KPI_L, F_KPI
        for rr in (4, 5):
            for cc in (col, c2):
                ws["%s%d" % (cc, rr)].fill = FILL_KPI
            ws["%s%d" % (col, rr)].alignment = CENTRO
    ws["D5"].font = Font(name=FONTE, size=20, bold=True, color="C00000")
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 34

    def cabecalho(lin, col, nomes):
        for j, h in enumerate(nomes):
            c = ws.cell(lin, col + j, h)
            c.font, c.fill, c.alignment, c.border = F_CAB, FILL_CAB, CENTRO, BORDA
        ws.row_dimensions[lin].height = 30

    def corpo(r, c1, c2, bold_cols=(), fmt=None):
        for j in range(c1, c2 + 1):
            c = ws.cell(r, j)
            c.border, c.alignment = BORDA, CENTRO
            c.font = Font(name=FONTE, size=9, bold=(j in bold_cols))
            if fmt and j in fmt:
                c.number_format = fmt[j]

    # ---------------- Programação da semana
    ws["B7"] = '="PROGRAMAÇÃO DA SEMANA "&%s' % P["svt"]
    ws["B7"].font = F_H2
    cabecalho(8, 3, ["ETAPA PROGRAMADA", "", "PROGRAMADAS", "REALIZADAS", "PENDENTES", "ADERÊNCIA",
                     "STATUS CALC. DIFERENTE"])
    ws.merge_cells("C8:D8")
    lin = 9
    for nome, _, stp, _ in ETAPAS:
        ws.merge_cells("C%d:D%d" % (lin, lin))
        ws.cell(lin, 3, nome + " (" + stp + ")")
        ws.cell(lin, 5, '=COUNTIF(%s,"%s")' % (sp, stp))
        ws.cell(lin, 6, '=COUNTIFS(%s,"%s",%s,"REALIZADA")' % (sp, stp, ps))
        ws.cell(lin, 7, "=E%d-F%d" % (lin, lin))
        ws.cell(lin, 8, '=IF(E%d=0,"-",F%d/E%d)' % (lin, lin, lin))
        ws.cell(lin, 9, '=E%d-COUNTIFS(%s,"%s",%s,"%s")' % (lin, sp, stp, st, stp))
        ws.cell(lin, 3).fill = PatternFill("solid", fgColor=COR_STATUS[stp])
        corpo(lin, 3, 9, bold_cols=(3,), fmt={8: "0%"})
        lin += 1
    ws.merge_cells("C%d:D%d" % (lin, lin))
    ws.cell(lin, 3, "TOTAL")
    for j in (5, 6, 7, 9):
        L = get_column_letter(j)
        ws.cell(lin, j, "=SUM(%s9:%s%d)" % (L, L, lin - 1))
    ws.cell(lin, 8, '=IF(E%d=0,"-",F%d/E%d)' % (lin, lin, lin))
    corpo(lin, 3, 9, bold_cols=range(3, 10), fmt={8: "0%"})
    for j in range(3, 10):
        ws.cell(lin, j).fill = FILL_KPI
    ws.conditional_formatting.add("I9:I%d" % lin, CellIsRule(operator="greaterThan", formula=["0"],
                                  fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    ws.cell(lin + 1, 3, "STATUS CALC. DIFERENTE = programada numa etapa que não bate com o status calculado "
                        "(ver aba PROGRAMAÇÃO, coluna PROG. x CALC.).").font = F_SUB
    ws.cell(lin + 2, 3, "Categoria das programadas:").font = Font(name=FONTE, size=9, bold=True)
    ws.cell(lin + 2, 6, "PRIORIDADE").font = F_TXT
    ws.cell(lin + 2, 7, '=COUNTIFS(%s,"?*",%s,"PRIORIDADE")' % (sp, cat))
    ws.cell(lin + 2, 8, "OPORTUNIDADE").font = F_TXT
    ws.cell(lin + 2, 9, '=COUNTIFS(%s,"?*",%s,"OPORTUNIDADE")' % (sp, cat))
    for j in (7, 9):
        ws.cell(lin + 2, j).font = Font(name=FONTE, size=10, bold=True, color=AZUL)
    fim_prog = lin + 2

    # gráfico programadas x realizadas
    ch = BarChart()
    ch.type, ch.grouping = "col", "clustered"
    ch.title = "Programadas × realizadas na semana"
    ch.add_data(Reference(ws, min_col=5, max_col=6, min_row=8, max_row=8 + len(ETAPAS)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=3, min_row=9, max_row=8 + len(ETAPAS)))
    for s, cor in zip(ch.series, ["1F3864", "00B050"]):
        s.graphicalProperties.solidFill = cor
        s.graphicalProperties.line.solidFill = cor
    ch.height, ch.width = 6.4, 15.5
    ch.roundedCorners = False
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ws.add_chart(ch, "K7")

    # ---------------- Status calculado x categoria
    r0 = max(fim_prog, 20) + 2
    ws.cell(r0, 2, "SITUAÇÃO DAS PASTAS (STATUS CALCULADO) × CATEGORIA DE ENTREGA").font = F_H2
    cabecalho(r0 + 1, 3, ["STATUS CALCULADO", "", "PRIORIDADE", "OPORTUNIDADE", "SEM CATEGORIA", "TOTAL", "% DO TOTAL"])
    ws.merge_cells("C%d:D%d" % (r0 + 1, r0 + 1))
    rt = r0 + 2 + len(STATUS_LISTA)
    for i, s in enumerate(STATUS_LISTA):
        r = r0 + 2 + i
        ws.merge_cells("C%d:D%d" % (r, r))
        ws.cell(r, 3, s).fill = PatternFill("solid", fgColor=COR_STATUS[s])
        ws.cell(r, 5, '=COUNTIFS(%s,$C%d,%s,"PRIORIDADE")' % (st, r, cat))
        ws.cell(r, 6, '=COUNTIFS(%s,$C%d,%s,"OPORTUNIDADE")' % (st, r, cat))
        ws.cell(r, 8, '=COUNTIF(%s,$C%d)' % (st, r))
        ws.cell(r, 7, "=H%d-E%d-F%d" % (r, r, r))
        ws.cell(r, 9, '=IF($H$%d=0,0,H%d/$H$%d)' % (rt, r, rt))
        corpo(r, 3, 9, bold_cols=(3, 8), fmt={9: "0%"})
    ws.merge_cells("C%d:D%d" % (rt, rt))
    ws.cell(rt, 3, "TOTAL")
    for j in range(5, 10):
        L = get_column_letter(j)
        ws.cell(rt, j, "=SUM(%s%d:%s%d)" % (L, r0 + 2, L, rt - 1))
    corpo(rt, 3, 9, bold_cols=range(3, 10), fmt={9: "0%"})
    for j in range(3, 10):
        ws.cell(rt, j).fill = FILL_KPI

    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "bar", "stacked", 100
    ch.title = "Pastas por status e categoria"
    ch.add_data(Reference(ws, min_col=5, max_col=7, min_row=r0 + 1, max_row=rt - 1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=3, min_row=r0 + 2, max_row=rt - 1))
    for s, cor in zip(ch.series, ["1F3864", "5B9BD5", "BFBFBF"]):
        s.graphicalProperties.solidFill = cor
        s.graphicalProperties.line.solidFill = cor
    ch.height, ch.width = 7.0, 15.5
    ch.roundedCorners = False
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(ch, "K%d" % r0)

    # ---------------- Rankings
    def ranking(lin, col, titulo_txt, nota, pontos, campos, n=20):
        ws.cell(lin, col, titulo_txt).font = F_H2
        ws.cell(lin + 1, col, nota).font = F_SUB
        cabecalho(lin + 2, col, ["#"] + [c[0] for c in campos])
        pt = R(pontos)
        Lk = get_column_letter(col)
        for k in range(1, n + 1):
            r = lin + 2 + k
            ws.cell(r, col, '=IF(COUNTIF(%s,">0")>=%d,%d,"")' % (pt, k, k))
            ref = "MATCH(LARGE(%s,$%s%d),%s,0)" % (pt, Lk, r, pt)
            for j, (_, s, fmt) in enumerate(campos, 1):
                idx = "INDEX(%s,%s)" % (R(s), ref)
                c = ws.cell(r, col + j, '=IF($%s%d="","",IF(%s="","",%s))' % (Lk, r, idx, idx))
                if fmt:
                    c.number_format = fmt
            for j in range(len(campos) + 1):
                c = ws.cell(r, col + j)
                c.border, c.font = BORDA, F_TXT
                c.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
                if k % 2 == 0:
                    c.fill = FILL_ZEBRA
        # restrição em vermelho
        Lr = get_column_letter(col + len(campos))
        ws.conditional_formatting.add("%s%d:%s%d" % (Lr, lin + 3, Lr, lin + 2 + n), FormulaRule(
            formula=['%s%d<>""' % (Lr, lin + 3)], fill=PatternFill("solid", fgColor="FFC7CE"),
            font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
        return lin + 2 + n

    lin = rt + 3
    base_campos = [("Nº DA PASTA", "Nº DA PASTA", None), ("SUBSOP HMS", "SUBSOP HMS", None),
                   ("CATEGORIA", "CATEGORIA DE ENTREGA", None)]
    fim = ranking(lin, 2, "MELHORES PASTAS PARA TESTE",
                  "Status calculado P/TESTE. Ordem: sem restrição > prioridade > oportunidade > maior avanço CTB.",
                  "PONTOS TESTE", base_campos + [("AVANÇO CTB", "AVANÇO CTB", "0%"),
                                                 ("PROGRAM. SEMANA", "STATUS PROGRAMAÇÃO", None),
                                                 ("RESTRIÇÃO", "RESTRIÇÕES ABERTAS", None)])
    ranking(lin, 11, "MELHORES PASTAS PARA REMONTAGEM",
            "Teste aprovado no HMSWEB, LV de montagem/remontagem pendente. Ordem: sem restrição > categoria.",
            "PONTOS REMONT.", base_campos + [("AVANÇO TESTE", "AVANÇO TESTE", None),
                                             ("PROGRAM. SEMANA", "STATUS PROGRAMAÇÃO", None),
                                             ("RESTRIÇÃO", "RESTRIÇÕES ABERTAS", None)])
    ws.conditional_formatting.add("F%d:F%d" % (lin + 3, fim), DataBarRule(start_type="num", start_value=0,
                                  end_type="num", end_value=1, color="5B9BD5"))
    lin = fim + 2
    fim2 = ranking(lin, 2, "PASTAS INTRA SKID EM PRÉ WALKDOWN",
                   "Certificados aprovados no HMSWEB, pendente LV de montagem e/ou remontagem. Ordem: sem restrição > categoria.",
                   "PONTOS PRÉ WALK", base_campos + [("AVANÇO PRÉ WALK", "AVANÇO PRÉ WALK", None),
                                                     ("PROGRAM. SEMANA", "STATUS PROGRAMAÇÃO", None),
                                                     ("RESTRIÇÃO", "RESTRIÇÕES ABERTAS", None)])

    # restrições por tipo (aba RESTRIÇÕES) + pastas com restrição no CONTROLE
    ws.cell(lin, 11, "RESTRIÇÕES").font = F_H2
    ws.cell(lin + 1, 11, "Pastas com RESTRIÇÕES ABERTAS preenchida no CONTROLE:").font = F_SUB
    ws.cell(lin + 1, 17, '=COUNTIF(%s,"?*")' % R("RESTRIÇÕES ABERTAS")).font = Font(name=FONTE, size=11, bold=True, color="C00000")
    cabecalho(lin + 2, 11, ["TIPO (aba RESTRIÇÕES)", "", "", "ABERTAS", "ATRASADAS", "ENCERRADAS"])
    ws.merge_cells("K%d:M%d" % (lin + 2, lin + 2))
    rr = lambda col: "'RESTRIÇÕES'!$%s$%d:$%s$%d" % (col, LIN_CAB + 1, col, ult_restr)
    for i, t in enumerate(TIPOS_RESTR):
        r = lin + 3 + i
        ws.merge_cells("K%d:M%d" % (r, r))
        ws.cell(r, 11, t)
        ws.cell(r, 14, '=COUNTIFS(%s,$K%d,%s,"<>ENCERRADA")' % (rr("D"), r, rr("I")))
        ws.cell(r, 15, '=COUNTIFS(%s,$K%d,%s,"ATRASADA")' % (rr("D"), r, rr("L")))
        ws.cell(r, 16, '=COUNTIFS(%s,$K%d,%s,"ENCERRADA")' % (rr("D"), r, rr("I")))
        corpo(r, 11, 16)
        ws.cell(r, 11).alignment = ESQ
    ws.conditional_formatting.add("O%d:O%d" % (lin + 3, lin + 2 + len(TIPOS_RESTR)), CellIsRule(
        operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(name=FONTE, size=9, bold=True, color="9C0006")))

    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = AZUL
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


# ----------------------------------------------------------------------------- avanço por semana
def gravar_avancos(linhas, base, d, cal, anterior=None):
    """Preenche as colunas AVANÇO das linhas. Mantém a semana já gravada; grava a semana do painel para o que
    passou a aprovado; com o painel anterior, o que já estava aprovado nele recebe 'ATÉ S-xxx'."""
    sem_atual = num_semana(cal.get(d["_data"].date()))
    if sem_atual is None:
        raise SystemExit("Data do painel %s fora do calendário" % d["_data"].date())
    ap_agora = {b["PASTA"]: etapas_aprovadas(b) for b in base}
    ap_antes, sem_antes = {}, None
    if anterior:
        ap_antes = {b["PASTA"]: etapas_aprovadas(b) for b in base_por_pasta(anterior)}
        sem_antes = num_semana(cal.get(anterior["_data"].date()))
    cont = collections.Counter()
    for l in linhas:
        agora = ap_agora.get(l["Nº DA PASTA"])
        for col in AVANCOS:
            atual = l.get(col)
            if not agora or not agora[col]:
                l[col] = None  # não aprovada (ou estornada) na PCP
                continue
            if atual not in (None, ""):
                cont["mantida"] += 1
                continue
            if anterior and ap_antes.get(l["Nº DA PASTA"], {}).get(col):
                l[col] = "ATÉ " + txt_semana(sem_antes)
                cont["anterior"] += 1
            else:
                l[col] = txt_semana(sem_atual)
                cont["nova"] += 1
    return sem_atual, cont


# ----------------------------------------------------------------------------- principal
def mais_recente(caminho, padroes):
    """Se 'caminho' é uma pasta, devolve o arquivo mais recente (com subpastas) que casa com os padrões."""
    if not os.path.isdir(caminho):
        return caminho
    achados = []
    for pad in padroes:
        achados += glob.glob(os.path.join(caminho, "**", pad), recursive=True)
    achados = [f for f in set(achados) if not os.path.basename(f).startswith("~$")]
    if not achados:
        raise SystemExit("Nenhum arquivo %s encontrado em %s" % (" / ".join(padroes), caminho))
    return max(achados, key=os.path.getmtime)


def _quando(f):
    return dt.datetime.fromtimestamp(os.path.getmtime(f)).strftime("%d/%m/%Y %H:%M")


def gerar(painel, controle, saida, calendario=None, painel_anterior=None, pcp=None, backup=None):
    painel = mais_recente(painel, ["*OLODUM*.html", "*olodum*.html"])
    print("Painel OLODUM: %s (salvo em %s)" % (painel, _quando(painel)))
    if pcp:
        pcp = mais_recente(pcp, ["OLODUM_BASEDADOS*.xlsx", "OLODUM_BASEDADOS*.xlsm"])
        print("Base PCP:      %s (salva em %s)" % (pcp, _quando(pcp)))
    d = ler_painel(painel)
    base = base_por_pasta(d)
    gerado_em = d["_data"].strftime("%d/%m/%Y %H:%M UTC") if d["_data"] else ""
    linhas, restr, wb_ant = ler_controle(controle)
    cal = ler_calendario(calendario if calendario else wb_ant)

    vistos = {l["Nº DA PASTA"] for l in linhas}
    sub = {b["PASTA"]: b["SUBSISTEMA"] for b in base}
    for l in linhas:
        if not l.get("SUBSOP HMS"):
            l["SUBSOP HMS"] = sub.get(l["Nº DA PASTA"], "")
    novas = [b for b in base if b["PASTA"] not in vistos]
    for b in novas:
        linhas.append({"Nº DA PASTA": b["PASTA"], "SUBSOP HMS": b["SUBSISTEMA"]})
    if pcp:
        av = avancos_pcp(ler_pcp(pcp), base, cal)
        cont = collections.Counter()
        for l in linhas:
            a = av.get(l["Nº DA PASTA"], {})
            for col in AVANCOS:
                l[col] = a.get(col)
                cont[col] += bool(l[col])
        sem = semana_da_data(d["_data"], cal)
        print("Avanço pela DATA DA EXECUÇÃO da PCP: " + ", ".join("%s %d" % (k, cont[k]) for k in AVANCOS))
    else:
        anterior = ler_painel(painel_anterior) if painel_anterior else None
        sem, cont = gravar_avancos(linhas, base, d, cal, anterior)
        print("Semanas de avanço: %d mantidas, %d 'ATÉ' (painel anterior), %d novas nesta semana"
              % (cont["mantida"], cont["anterior"], cont["nova"]))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ult_cal = aba_calendario(wb, cal)
    P = aba_parametros(wb, ult_cal)
    nb = aba_base(wb, base, gerado_em)
    ult_ctrl = aba_controle(wb, linhas, P, nb)
    ult_restr = aba_restricoes(wb, restr, P, ult_ctrl)
    aba_programacao(wb, ult_ctrl, P)
    aba_painel(wb, ult_ctrl, ult_restr, gerado_em, P)
    ordem = ["PAINEL", "CONTROLE", "PROGRAMAÇÃO", "RESTRIÇÕES", "BASE_OLODUM", "PARÂMETROS", "CALENDÁRIO"]
    wb._sheets = [wb[n] for n in ordem]
    wb["CONTROLE"].sheet_properties.tabColor = "2E75B6"
    wb["RESTRIÇÕES"].sheet_properties.tabColor = "C00000"
    wb["PROGRAMAÇÃO"].sheet_properties.tabColor = "00B050"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    if backup and os.path.exists(saida):
        os.makedirs(backup, exist_ok=True)
        nome, ext = os.path.splitext(os.path.basename(saida))
        copia = os.path.join(backup, "%s_%s%s" % (nome, dt.datetime.now().strftime("%Y%m%d_%H%M"), ext))
        shutil.copy2(saida, copia)
        print("Cópia de segurança: %s" % copia)
    temp = saida + ".tmp.xlsx"
    wb.save(temp)
    subs_g = sorted({_subsistema_principal(l.get("SUBSOP HMS")) for l in linhas if l.get("SUBSOP HMS")} - {""})
    _painel_no_arquivo(temp, LIN_CAB, P, subs_g, gerado_em, ult_ctrl)
    try:
        os.replace(temp, saida)
    except PermissionError:
        os.remove(temp)
        raise SystemExit("ERRO: não consegui gravar %s. A planilha está aberta no Excel (sua ou de outra pessoa)? "
                         "Feche e rode de novo." % saida)
    print("Painel de %s (semana %s) | pastas no controle: %d (%d novas)" % (gerado_em, txt_semana(sem), len(linhas), len(novas)))
    print("Gravado:", saida)



# ----------------------------------------------------------------------------- modo ATUALIZAR (no próprio arquivo)
# Altera só as células das colunas automáticas do CONTROLE e a aba BASE_OLODUM, editando o XML do .xlsx.
# Gráficos, validações, formatação, abas e colunas manuais ficam exatamente como estão.
_RE_ROW = re.compile(r'<row\b[^>]*?(?:/>|>.*?</row>)', re.S)
_RE_C = re.compile(r'<c\b[^>]*?(?:/>|>.*?</c>)', re.S)
_RE_ATTR = lambda nome: re.compile(r'\b%s="([^"]*)"' % nome)


def _col_idx(letras):
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def _ref(cel):
    m = re.match(r"([A-Z]+)(\d+)", cel)
    return m.group(1), int(m.group(2))


def _cel_formula(ref, s, formula):
    f = formula[1:] if formula.startswith("=") else formula
    return '<c r="%s"%s><f>%s</f></c>' % (ref, s, _esc(f))


def _cel_valor(ref, s, v):
    if v is None or v == "":
        return '<c r="%s"%s/>' % (ref, s)
    if isinstance(v, (dt.date, dt.datetime)):
        d = v.date() if isinstance(v, dt.datetime) else v
        v = (d - dt.date(1899, 12, 30)).days
    if isinstance(v, bool):
        v = int(v)
    if isinstance(v, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(v) if isinstance(v, float) else v)
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, _esc(str(v)))


def _data_txt(ref):
    """Data dd/mm/aaaa em texto sem depender do idioma do Excel (TEXT com "yyyy" falha no Excel em português)."""
    return 'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&"/"&YEAR(%s)' % (ref, ref, ref)


def _bloco(st, tag):
    m = re.search(r"<%s\b[^>]*>(.*?)</%s>" % (tag, tag), st, re.S)
    return m


def _add_item(st, tag, item_tag, item):
    """Acrescenta 'item' ao bloco <tag> do styles.xml (se ainda não existir) e devolve (styles, índice)."""
    m = _bloco(st, tag)
    itens = re.findall(r"<%s\b[^>]*?(?:/>|>.*?</%s>)" % (item_tag, item_tag), m.group(1), re.S)
    if item in itens:
        return st, itens.index(item)
    ini, fim = m.start(1), m.end(1)
    st = st[:ini] + m.group(1) + item + st[fim:]
    st = re.sub(r'(<%s\b[^>]*\bcount=")\d+"' % tag, r'\g<1>%d"' % (len(itens) + 1), st, count=1)
    return st, len(itens)


def _xf(st, i):
    m = _bloco(st, "cellXfs")
    return re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", m.group(1), re.S)[int(i)]


def _attr_set(xml_tag, nome, valor):
    if re.search(r'\b%s="' % nome, xml_tag):
        return re.sub(r'\b%s="[^"]*"' % nome, '%s="%s"' % (nome, valor), xml_tag, count=1)
    return xml_tag.replace("<xf", '<xf %s="%s"' % (nome, valor), 1)


def _xf_pct(st, s_base):
    xf = _attr_set(_attr_set(_xf(st, s_base), "numFmtId", "9"), "applyNumberFormat", "1")
    return _add_item(st, "cellXfs", "xf", xf)


def _xf_cor_fonte(st, s_base, rgb):
    xf = _xf(st, s_base)
    fid = int(re.search(r'fontId="(\d+)"', xf).group(1))
    fonte = re.findall(r"<font\b[^>]*?(?:/>|>.*?</font>)", _bloco(st, "fonts").group(1), re.S)[fid]
    if "<color" in fonte:
        fonte = re.sub(r"<color\b[^>]*/>", '<color rgb="%s"/>' % rgb, fonte)
    else:
        fonte = fonte.replace("</font>", '<color rgb="%s"/></font>' % rgb) if "</font>" in fonte \
            else '<font><color rgb="%s"/></font>' % rgb
    st, nf = _add_item(st, "fonts", "font", fonte)
    xf = _attr_set(_attr_set(xf, "fontId", str(nf)), "applyFont", "1")
    return _add_item(st, "cellXfs", "xf", xf)


def _topo_controle(x_ctrl, st, lin_hdr, ult_lin, P):
    """Linha da semana vigente, títulos e SUBTOTAIS (respeitam o filtro) acima do cabeçalho do CONTROLE.
    Usa o estilo amarelo da célula de subtotal da coluna B (linha acima do cabeçalho)."""
    sub, lab, lsem = lin_hdr - 1, lin_hdr - 2, lin_hdr - 3
    if lsem < 1:
        return x_ctrl, st, False
    s_sub = _estilos_linha(x_ctrl, sub).get("B")
    s_hdr = _estilos_linha(x_ctrl, lin_hdr).get("D") or _estilos_linha(x_ctrl, lin_hdr).get("A")
    s_tit = _estilos_linha(x_ctrl, 1).get("A")
    s_pct = s_red = None
    if s_sub is not None:
        st, s_pct = _xf_pct(st, s_sub)
    if s_tit is not None:
        st, s_red = _xf_cor_fonte(st, s_tit, "FFC00000")
    r1, r2 = lin_hdr + 1, max(1000, ult_lin + 200)
    vis = "SUBTOTAL(3,OFFSET($B$%d,ROW($B$%d:$B$%d)-ROW($B$%d),0))" % (r1, r1, r2, r1)
    rg = lambda L: "$%s$%d:$%s$%d" % (L, r1, L, r2)
    conta = lambda L, crit: "=SUMPRODUCT(%s*(%s=%s))" % (vis, rg(L), crit)
    svt, hoje = P["svt"], P["hoje"]
    seg = "(%s-WEEKDAY(%s,2)+1)" % (hoje, hoje)
    sub_f, lab_f = {}, {}
    sub_f["B"], lab_f["B"] = "=SUBTOTAL(3,%s)" % rg("B"), "PASTAS (FILTRO)"
    for nome, rot in (("JUNTAS PREVISTAS CTB", "JUNTAS PREV."), ("JUNTAS LIBERADAS CTB", "JUNTAS LIB."),
                      ("SALDO DE JUNTAS", "SALDO JUNTAS")):
        sub_f[C[nome]], lab_f[C[nome]] = "=SUBTOTAL(9,%s)" % rg(C[nome]), rot
    G, H, J = C["JUNTAS PREVISTAS CTB"], C["JUNTAS LIBERADAS CTB"], C["AVANÇO CTB"]
    sub_f[J], lab_f[J] = '=IF(N(%s%d)>0,%s%d/%s%d,"")' % (G, sub, H, sub, G, sub), "AVANÇO CTB"
    for nome in ("MONTAGEM HMS (LVM)", "MAPA JUNTAS HMS", "CERTIFICADO TESTE HMSWEB", "LAVAGEM HMS", "REMONTAGEM HMS"):
        sub_f[C[nome]], lab_f[C[nome]] = conta(C[nome], '"APROVADA"'), "APROVADAS"
    sub_f[C["CATEGORIA DE ENTREGA"]], lab_f[C["CATEGORIA DE ENTREGA"]] = \
        "=SUBTOTAL(3,%s)" % rg(C["CATEGORIA DE ENTREGA"]), "COM CATEGORIA"
    for _, col_sem, stp, col_av in ETAPAS:
        sub_f[C[col_sem]] = conta(C["STATUS PROGRAMAÇÃO"], '"%s"' % stp)
        lab_f[C[col_sem]] = '="PROG. "&%s' % svt
        sub_f[C[col_av]] = conta(C[col_av], svt)
        lab_f[C[col_av]] = '="EXEC. "&%s' % svt
    sub_f[C["RESTRIÇÕES ABERTAS"]], lab_f[C["RESTRIÇÕES ABERTAS"]] = \
        "=SUBTOTAL(3,%s)" % rg(C["RESTRIÇÕES ABERTAS"]), "COM RESTRIÇÃO"
    L = C["PROGRAMAÇÃO DA SEMANA"]
    sub_f[L], lab_f[L] = conta(L, '"REALIZADA"'), '="REALIZADAS "&%s' % svt
    tit = ('=IF(\'PARÂMETROS\'!$B$7<>"","SEMANA ANALISADA: "&%s,"SEMANA VIGENTE: "&%s&"   ("&TEXT(DAY(%s),"00")'
           '&"/"&TEXT(MONTH(%s),"00")&" A "&TEXT(DAY(%s+6),"00")&"/"&TEXT(MONTH(%s+6),"00")&"/"&YEAR(%s+6)&")")'
           % (svt, svt, seg, seg, seg, seg, seg))
    est = lambda s: ' s="%s"' % s if s is not None else ""
    ed = {sub: {}, lab: {}, lsem: {"A": (lambda ref, s: _cel_formula(ref, est(s_red), tit))}}
    for col, f in sub_f.items():
        sty = s_pct if col == J else s_sub
        ed[sub][col] = (lambda f, sty: lambda ref, s: _cel_formula(ref, est(sty), f))(f, sty)
    for col, t in lab_f.items():
        if t.startswith("="):
            ed[lab][col] = (lambda t: lambda ref, s: _cel_formula(ref, est(s_hdr), t))(t)
        else:
            ed[lab][col] = (lambda t: lambda ref, s: _cel_valor(ref, est(s_hdr), t))(t)
    x_ctrl = _patch_linhas(x_ctrl, ed, {}, alturas={lab: 30, lsem: 22})
    return x_ctrl, st, True


def _patch_linhas(xml, novas, estilo_col, remover_apos=None, alturas=None):
    """novas = {linha: {coluna: função(ref, s) -> xml da célula}}. remover_apos: apaga linhas > N que não estejam
    em 'novas' (usado na BASE)."""
    ini = xml.index("<sheetData")
    abre = xml.index(">", ini) + 1
    if xml[ini:abre].endswith("/>"):  # <sheetData/> vazio
        corpo, depois = "", xml[abre:]
    else:
        fim = xml.index("</sheetData>")
        corpo, depois = xml[abre:fim], xml[fim + len("</sheetData>"):]
    linhas = {}
    for m in _RE_ROW.finditer(corpo):
        raw = m.group(0)
        r = int(_RE_ATTR("r").search(raw).group(1))
        linhas[r] = raw
    for r in sorted(set(linhas) | set(novas)):
        if remover_apos is not None and r > remover_apos and r not in novas:
            del linhas[r]
            continue
        if r not in novas:
            continue
        raw = linhas.get(r, '<row r="%d"/>' % r)
        tag = re.match(r"<row\b[^>]*?(?=/?>)", raw).group(0)
        tag = re.sub(r'\sspans="[^"]*"', "", tag)
        if alturas and r in alturas:
            tag = re.sub(r'\s(ht|customHeight)="[^"]*"', "", tag) + ' ht="%s" customHeight="1"' % alturas[r]
        celulas = {}
        for c in _RE_C.findall(raw[len(tag):]):
            celulas[_ref(_RE_ATTR("r").search(c).group(1))[0]] = c
        for col, fn in novas[r].items():
            atual = celulas.get(col)
            sm = _RE_ATTR("s").search(atual) if atual else None
            s = ' s="%s"' % sm.group(1) if sm else (' s="%s"' % estilo_col[col] if estilo_col.get(col) else "")
            celulas[col] = fn("%s%d" % (col, r), s)
        corpo_row = "".join(celulas[c] for c in sorted(celulas, key=_col_idx))
        linhas[r] = "%s>%s</row>" % (tag, corpo_row)
    novo = "".join(linhas[r] for r in sorted(linhas))
    return xml[:ini] + "<sheetData>" + novo + "</sheetData>" + depois


def _estilos_linha(xml, r):
    m = re.search(r'<row\b[^>]*\br="%d"[^>]*?(?:/>|>.*?</row>)' % r, xml, re.S)
    est = {}
    if m:
        for c in _RE_C.findall(m.group(0)):
            col = _ref(_RE_ATTR("r").search(c).group(1))[0]
            sm = _RE_ATTR("s").search(c)
            if sm:
                est[col] = sm.group(1)
    return est


def _grupos_compartilhados(xml):
    """{si: (colunas do ref)} das fórmulas compartilhadas (Excel grava assim ao arrastar fórmulas)."""
    g = {}
    for m in re.finditer(r'<f\b[^>]*\bt="shared"[^>]*\bref="([A-Z]+)\d+(?::([A-Z]+)\d+)?"[^>]*\bsi="(\d+)"', xml):
        a, b = m.group(1), m.group(2) or m.group(1)
        g[m.group(3)] = set(get_column_letter(i) for i in range(_col_idx(a), _col_idx(b) + 1))
    for m in re.finditer(r'<f\b[^>]*\bsi="(\d+)"[^>]*\bref="([A-Z]+)\d+(?::([A-Z]+)\d+)?"', xml):
        a, b = m.group(2), m.group(3) or m.group(2)
        g[m.group(1)] = set(get_column_letter(i) for i in range(_col_idx(a), _col_idx(b) + 1))
    return g


# ----------------------------------------------------------------------------- PAINEL (refeito no pacote .xlsx)
ABA_RESUMO = "RESUMO GERENCIAL"  # aba antiga: removida a cada atualização
_NS_WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_BORDA = ('<border><left style="thin"><color rgb="FFBFC9D9"/></left><right style="thin"><color rgb="FFBFC9D9"/>'
          '</right><top style="thin"><color rgb="FFBFC9D9"/></top><bottom style="thin"><color rgb="FFBFC9D9"/>'
          '</bottom><diagonal/></border>')
CAT_PAINEL = ["PRIORIDADE 1", "PRIORIDADE 2", "OPORTUNIDADE"]
STATUS_SUB = ["EM MONTAGEM", "P/TESTE", "P/REMONTAGEM", "QUALIDADE-REMONTAGEM", "PRÉ WALKDOWN",
              "PENDENTE DOCUMENTAÇÃO HMSWEB", "CONCLUÍDA"]


def _estilo(st, cor="FF000000", tam=10, negrito=False, italico=False, fundo=None, horiz="left", quebra=False,
            numfmt=0, borda=False):
    fonte = '<font>%s%s<sz val="%s"/><color rgb="%s"/><name val="Arial"/><family val="2"/></font>' % (
        "<b/>" if negrito else "", "<i/>" if italico else "", tam, cor)
    st, fid = _add_item(st, "fonts", "font", fonte)
    fill_id = 0
    if fundo:
        st, fill_id = _add_item(st, "fills", "fill", '<fill><patternFill patternType="solid"><fgColor rgb="%s"/>'
                                                     '<bgColor indexed="64"/></patternFill></fill>' % fundo)
    bid = 0
    if borda:
        st, bid = _add_item(st, "borders", "border", _BORDA)
    xf = ('<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyNumberFormat="1" applyFont="1" '
          'applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="%s" vertical="center"%s/></xf>'
          % (numfmt, fid, fill_id, bid, horiz, ' wrapText="1"' if quebra else ""))
    return _add_item(st, "cellXfs", "xf", xf)


def _subsistema_principal(v):
    """Primeiro subsistema do SUBSOP ('A /B' -> 'A')."""
    return re.split(r"\s*/\s*", str(v or "").strip())[0].strip()


def _painel_xml(st, lin_hdr, P, subs, base_txt, ult=1000):
    r1, r2 = lin_hdr + 1, 3000
    q2 = max(ult + 300, r1 + 600)  # faixa menor para as fórmulas SUMPRODUCT por subsistema
    rq = lambda nome: "'CONTROLE'!$%s$%d:$%s$%d" % (C[nome], r1, C[nome], q2)
    rg = lambda nome: "'CONTROLE'!$%s$%d:$%s$%d" % (C[nome], r1, C[nome], r2)
    R, S, Y, G, H, B, SUB, X, AG = (rg("STATUS CALCULADO"), rg("CATEGORIA DE ENTREGA"),
                                   rg("DATA CONCLUSÃO INFORMADO RESPONSAVEL"), rg("JUNTAS PREVISTAS CTB"),
                                   rg("JUNTAS LIBERADAS CTB"), rg("Nº DA PASTA"), rg("SUBSOP HMS"),
                                   rg("STATUS PROGRAMAÇÃO"), rg("PROGRAMAÇÃO DA SEMANA"))
    AZ, CL = "FF1F3864", "FFEEF3FA"
    E = {}
    st, E["tit"] = _estilo(st, AZ, 18, negrito=True)
    st, E["sub"] = _estilo(st, "FF595959", 9, italico=True)
    st, E["kl"] = _estilo(st, "FF595959", 8, negrito=True, fundo=CL, horiz="center", quebra=True)
    st, E["kv"] = _estilo(st, AZ, 22, negrito=True, fundo=CL, horiz="center")
    st, E["kr"] = _estilo(st, "FFC00000", 22, negrito=True, fundo=CL, horiz="center")
    st, E["kp"] = _estilo(st, AZ, 22, negrito=True, fundo=CL, horiz="center", numfmt=9)
    st, E["h2"] = _estilo(st, AZ, 11, negrito=True)
    st, E["cab"] = _estilo(st, "FFFFFFFF", 9, negrito=True, fundo=AZ, horiz="center", quebra=True, borda=True)
    st, E["n"] = _estilo(st, "FF000000", 10, horiz="center", borda=True)
    st, E["nb"] = _estilo(st, "FF000000", 10, negrito=True, horiz="center", borda=True)
    st, E["p"] = _estilo(st, "FF000000", 10, horiz="center", numfmt=9, borda=True)
    st, E["lab"] = _estilo(st, "FF000000", 9, borda=True)
    st, E["tl"] = _estilo(st, AZ, 10, negrito=True, fundo=CL, horiz="center", borda=True)
    st, E["tn"] = _estilo(st, AZ, 10, negrito=True, fundo=CL, horiz="center", borda=True)
    st, E["tp"] = _estilo(st, AZ, 10, negrito=True, fundo=CL, horiz="center", numfmt=9, borda=True)
    st, E["nota"] = _estilo(st, "FF595959", 8, italico=True)
    for s_ in STATUS_LISTA:
        st, E["st:" + s_] = _estilo(st, "FF000000", 9, negrito=True, fundo="FF" + COR_STATUS[s_], horiz="center",
                                    quebra=True, borda=True)

    cel, merges, alt, cf = [], [], {}, []

    def put(ref, val, est, f=False):
        cel.append((ref, val, E[est], f))

    def caixa(r, c1, c2, val, est, f=False):  # célula mesclada com estilo em todas as células
        L1, L2 = get_column_letter(c1), get_column_letter(c2)
        put("%s%d" % (L1, r), val, est, f)
        for c in range(c1 + 1, c2 + 1):
            put("%s%d" % (get_column_letter(c), r), None, est)
        if c2 > c1:
            merges.append("%s%d:%s%d" % (L1, r, L2, r))

    hoje = P["hoje"]
    put("B1", "PAINEL GERENCIAL - SALA DE GUERRA DE PASTAS DE TUBULAÇÃO", "tit")
    alt[1] = 28
    put("B2", '="Base Painel OLODUM: %s   |   Semana vigente: "&%s&"   |   Atualizado em "&%s'
        % (base_txt, P["svt"], _data_txt(hoje)), "sub", True)

    # indicadores
    kpis = [("PASTAS CONTROLADAS", '=COUNTIF(%s,"?*")' % B, "kv"),
            ("SEMANA VIGENTE", "=" + P["svt"], "kr"),
            ("PROGRAMADAS NA SEMANA", '=COUNTIF(%s,"?*")' % X, "kv"),
            ("REALIZADO NA SEMANA", '=COUNTIF(%s,"REALIZADA")' % AG, "kv"),
            ("% AVANÇO CONCLUÍDO", '=IF(B5=0,0,COUNTIF(%s,"CONCLUÍDA")/B5)' % R, "kp")]
    for i, (rot, f, est) in enumerate(kpis):
        c = 2 + 2 * i
        caixa(4, c, c + 1, rot, "kl")
        caixa(5, c, c + 1, f, est, True)
    alt[4], alt[5] = 24, 40

    # situação x categoria (com / sem data de conclusão)
    put("B7", "SITUAÇÃO DAS PASTAS (STATUS CALCULADO) × CATEGORIA DE ENTREGA", "h2")
    alt[7] = 20
    caixa(8, 2, 2, "STATUS CALCULADO", "cab")
    caixa(9, 2, 2, None, "cab")
    merges.append("B8:B9")
    for i, cat in enumerate(CAT_PAINEL):
        c = 3 + 2 * i
        caixa(8, c, c + 1, cat, "cab")
        put("%s9" % get_column_letter(c), "COM DATA", "cab")
        put("%s9" % get_column_letter(c + 1), "SEM DATA", "cab")
    for c, rot in ((9, "SEM CATEGORIA"), (10, "TOTAL"), (11, "% DO TOTAL")):
        L = get_column_letter(c)
        put("%s8" % L, rot, "cab")
        put("%s9" % L, None, "cab")
        merges.append("%s8:%s9" % (L, L))
    alt[8], alt[9] = 22, 22
    r0 = 10
    rt = r0 + len(STATUS_LISTA)
    for i, s_ in enumerate(STATUS_LISTA):
        r = r0 + i
        put("B%d" % r, s_, "st:" + s_)
        base = '%s,$B%d,%s,"<>"' % (R, r, B)
        for j, cat in enumerate(CAT_PAINEL):
            Lc, Ls = get_column_letter(3 + 2 * j), get_column_letter(4 + 2 * j)
            put("%s%d" % (Lc, r), '=COUNTIFS(%s,%s,"%s",%s,">0")' % (base, S, cat, Y), "n", True)
            put("%s%d" % (Ls, r), '=COUNTIFS(%s,%s,"%s")-%s%d' % (base, S, cat, Lc, r), "n", True)
        put("I%d" % r, '=J%d-SUM(C%d:H%d)' % (r, r, r), "n", True)
        put("J%d" % r, '=COUNTIFS(%s)' % base, "nb", True)
        put("K%d" % r, '=IF($J$%d=0,0,J%d/$J$%d)' % (rt, r, rt), "p", True)
        alt[r] = 18
    put("B%d" % rt, "TOTAL", "tl")
    for c in range(3, 11):
        L = get_column_letter(c)
        put("%s%d" % (L, rt), "=SUM(%s%d:%s%d)" % (L, r0, L, rt - 1), "tn", True)
    put("K%d" % rt, "=SUM(K%d:K%d)" % (r0, rt - 1), "tp", True)
    put("B%d" % (rt + 1), "COM DATA = data de conclusão informada pela produção (coluna DATA CONCLUSÃO, só datas; "
                           "HOLD e textos contam como SEM DATA).", "nota")
    cf.append(("K%d:K%d" % (r0, rt - 1), "FF5B9BD5"))

    # programação da semana (ao lado)
    put("N7", '="PROGRAMAÇÃO DA SEMANA "&%s' % P["svt"], "h2", True)
    for c, rot in zip(range(14, 19), ["ETAPA PROGRAMADA", "PROGRAMADAS", "REALIZADAS", "PENDENTES", "ADERÊNCIA"]):
        put("%s8" % get_column_letter(c), rot, "cab")
    for i, (nome, _, stp, _) in enumerate(ETAPAS):
        r = 9 + i
        put("N%d" % r, stp, "st:" + stp)
        put("O%d" % r, '=COUNTIF(%s,"%s")' % (X, stp), "n", True)
        put("P%d" % r, '=COUNTIFS(%s,"%s",%s,"REALIZADA")' % (X, stp, AG), "n", True)
        put("Q%d" % r, "=O%d-P%d" % (r, r), "n", True)
        put("R%d" % r, "=IF(O%d=0,0,P%d/O%d)" % (r, r, r), "p", True)
    rp = 9 + len(ETAPAS)
    put("N%d" % rp, "TOTAL", "tl")
    for L in "OPQ":
        put("%s%d" % (L, rp), "=SUM(%s9:%s%d)" % (L, L, rp - 1), "tn", True)
    put("R%d" % rp, "=IF(O%d=0,0,P%d/O%d)" % (rp, rp, rp), "tp", True)
    put("N%d" % (rp + 2), "PROGRAMADAS NA SEMANA POR CATEGORIA", "h2")
    for c, rot in zip(range(14, 16), ["CATEGORIA", "PROGRAMADAS"]):
        put("%s%d" % (get_column_letter(c), rp + 3), rot, "cab")
    for i, cat in enumerate(CAT_PAINEL + ["SEM CATEGORIA"]):
        r = rp + 4 + i
        put("N%d" % r, cat, "lab")
        crit = '""' if cat == "SEM CATEGORIA" else '"%s"' % cat
        put("O%d" % r, '=COUNTIFS(%s,"?*",%s,%s)' % (X, S, crit), "n", True)

    # avanço do status calculado por subsistema
    rs = max(rt + 3, rp + 10)
    put("B%d" % rs, "AVANÇO DO STATUS CALCULADO POR SUBSISTEMA", "h2")
    put("F%d" % rs, "Pasta contada no 1º subsistema do SUBSOP HMS. OUTROS = SEM CTB, SEM ESCOPO e NÃO ENCONTRADA.", "nota")
    cab_s = ["SUBSISTEMA", "PASTAS"] + STATUS_SUB + ["OUTROS", "% CONCLUÍDA", "AVANÇO CTB"]
    for c, rot in enumerate(cab_s, 2):
        put("%s%d" % (get_column_letter(c), rs + 1), rot, "cab")
    alt[rs + 1] = 48
    rr = rs + 2
    SUBq, Rq, Gq, Hq = rq("SUBSOP HMS"), rq("STATUS CALCULADO"), rq("JUNTAS PREVISTAS CTB"), rq("JUNTAS LIBERADAS CTB")
    for sub in subs:
        sub = sub.replace('"', "")
        # 1º subsistema do SUBSOP, ignorando espaços antes/depois
        eh = '--(LEFT(TRIM(%s),%d)="%s")' % (SUBq, len(sub), sub)
        put("B%d" % rr, sub, "lab")
        put("C%d" % rr, "=SUMPRODUCT(%s)" % eh, "nb", True)
        for k, s_ in enumerate(STATUS_SUB):
            put("%s%d" % (get_column_letter(4 + k), rr), '=SUMPRODUCT(%s,--(%s="%s"))' % (eh, Rq, s_), "n", True)
        Lo, Lc = get_column_letter(4 + len(STATUS_SUB)), get_column_letter(3 + len(STATUS_SUB))
        put("%s%d" % (Lo, rr), "=C%d-SUM(D%d:%s%d)" % (rr, rr, Lc, rr), "n", True)
        Lp, La = get_column_letter(5 + len(STATUS_SUB)), get_column_letter(6 + len(STATUS_SUB))
        put("%s%d" % (Lp, rr), '=IF(C%d=0,"",%s%d/C%d)' % (rr, Lc, rr, rr), "p", True)
        put("%s%d" % (La, rr), '=IFERROR(SUMPRODUCT(%s,%s)/SUMPRODUCT(%s,%s),"")' % (eh, Hq, eh, Gq), "p", True)
        rr += 1
    put("B%d" % rr, "TOTAL", "tl")
    for c in range(3, 5 + len(STATUS_SUB)):
        L = get_column_letter(c)
        put("%s%d" % (L, rr), "=SUM(%s%d:%s%d)" % (L, rs + 2, L, rr - 1), "tn", True)
    Lp, La, Lc = (get_column_letter(5 + len(STATUS_SUB)), get_column_letter(6 + len(STATUS_SUB)),
                  get_column_letter(3 + len(STATUS_SUB)))
    put("%s%d" % (Lp, rr), '=IF(C%d=0,"",%s%d/C%d)' % (rr, Lc, rr, rr), "tp", True)
    put("%s%d" % (La, rr), '=IFERROR(SUM(%s)/SUM(%s),"")' % (H, G), "tp", True)
    if rr > rs + 2:
        cf.append(("%s%d:%s%d" % (Lp, rs + 2, Lp, rr - 1), "FF00B050"))
        cf.append(("%s%d:%s%d" % (La, rs + 2, La, rr - 1), "FF5B9BD5"))

    # XML
    linhas = collections.defaultdict(list)
    for ref, val, s_, f in cel:
        linhas[_ref(ref)[1]].append((ref, val, s_, f))
    corpo = []
    for r in sorted(linhas):
        cs = []
        for ref, val, s_, f in sorted(linhas[r], key=lambda c: _col_idx(_ref(c[0])[0])):
            cs.append(_cel_formula(ref, ' s="%d"' % s_, val) if f else _cel_valor(ref, ' s="%d"' % s_, val))
        ht = ' ht="%s" customHeight="1"' % alt[r] if r in alt else ""
        corpo.append('<row r="%d"%s>%s</row>' % (r, ht, "".join(cs)))
    larg = {1: 2, 2: 34}
    for c in range(3, 14):
        larg[c] = 11.5
    larg[14], larg[15], larg[16], larg[17], larg[18] = 30, 14, 13, 13, 13
    cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (c, c, w) for c, w in sorted(larg.items()))
    cfx = "".join('<conditionalFormatting sqref="%s"><cfRule type="dataBar" priority="%d"><dataBar>'
                  '<cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="%s"/></dataBar></cfRule>'
                  '</conditionalFormatting>' % (ref, i + 1, cor) for i, (ref, cor) in enumerate(cf))
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<sheetPr><tabColor rgb="FF1F3864"/><pageSetUpPr fitToPage="1"/></sheetPr>'
           '<dimension ref="B1:R%d"/>'
           '<sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0">'
           '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>'
           '<selection pane="bottomLeft" activeCell="B6" sqref="B6"/></sheetView></sheetViews>'
           '<sheetFormatPr defaultRowHeight="15"/><cols>%s</cols>'
           '<sheetData>%s</sheetData><mergeCells count="%d">%s</mergeCells>%s'
           '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
           '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
           '</worksheet>') % (rr, cols, "".join(corpo), len(merges),
                              "".join('<mergeCell ref="%s"/>' % m for m in merges), cfx)
    return xml, st


def _mapa_abas(arquivos):
    wbx = arquivos["xl/workbook.xml"].decode("utf-8")
    rels = arquivos["xl/_rels/workbook.xml.rels"].decode("utf-8")
    abas = []
    for m in re.finditer(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*/>', wbx):
        t = re.search(r'<Relationship\b[^>]*Id="%s"[^>]*Target="([^"]+)"' % m.group(2), rels) or \
            re.search(r'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="%s"' % m.group(2), rels)
        abas.append((m.group(1).replace("&amp;", "&"), m.group(2), "xl/" + t.group(1).lstrip("/").replace("xl/", ""),
                     m.group(0)))
    return abas


def _rels_da_aba(caminho):
    pasta, nome = caminho.rsplit("/", 1)
    return "%s/_rels/%s.rels" % (pasta, nome)


def _apagar_partes(arquivos, partes):
    ct = arquivos["[Content_Types].xml"].decode("utf-8")
    for p in partes:
        arquivos.pop(p, None)
        ct = re.sub(r'<Override\b[^>]*PartName="/%s"[^>]*/>' % re.escape(p), "", ct)
    arquivos["[Content_Types].xml"] = ct.encode("utf-8")


def _desenhos_da_aba(arquivos, caminho):
    """Partes ligadas à aba (desenhos, gráficos, impressora) para apagar junto com o conteúdo antigo."""
    rel = _rels_da_aba(caminho)
    partes = [rel] if rel in arquivos else []
    if rel not in arquivos:
        return partes
    base = caminho.rsplit("/", 1)[0]
    for alvo in re.findall(r'Target="([^"]+)"', arquivos[rel].decode("utf-8")):
        p = os.path.normpath(os.path.join(base, alvo)).replace("\\", "/")
        partes.append(p)
        r2 = _rels_da_aba(p)
        if r2 in arquivos:
            partes.append(r2)
            b2 = p.rsplit("/", 1)[0]
            for a2 in re.findall(r'Target="([^"]+)"', arquivos[r2].decode("utf-8")):
                partes.append(os.path.normpath(os.path.join(b2, a2)).replace("\\", "/"))
    return partes


def remover_aba(arquivos, nome):
    abas = _mapa_abas(arquivos)
    idx = next((i for i, a in enumerate(abas) if a[0] == nome), None)
    if idx is None:
        return arquivos, False
    _, rid, caminho, tag = abas[idx]
    _apagar_partes(arquivos, _desenhos_da_aba(arquivos, caminho) + [caminho])
    wbx = arquivos["xl/workbook.xml"].decode("utf-8").replace(tag, "")
    wbx = re.sub(r'<definedName\b[^>]*localSheetId="%d"[^>]*>.*?</definedName>' % idx, "", wbx, flags=re.S)
    wbx = re.sub(r'localSheetId="(\d+)"', lambda m: 'localSheetId="%d"' % (int(m.group(1)) - (int(m.group(1)) > idx)), wbx)
    wbx = re.sub(r"<definedNames>\s*</definedNames>", "", wbx)
    arquivos["xl/workbook.xml"] = wbx.encode("utf-8")
    rels = arquivos["xl/_rels/workbook.xml.rels"].decode("utf-8")
    arquivos["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship\b[^>]*Id="%s"[^>]*/>' % rid, "", rels).encode("utf-8")
    if "docProps/app.xml" in arquivos:
        app = arquivos["docProps/app.xml"].decode("utf-8")
        if "<vt:lpstr>%s</vt:lpstr>" % nome in app:
            app = app.replace("<vt:lpstr>%s</vt:lpstr>" % nome, "", 1)
            app = re.sub(r"(<HeadingPairs>.*?<vt:i4>)(\d+)(</vt:i4>)", lambda m: "%s%d%s" % (
                m.group(1), int(m.group(2)) - 1, m.group(3)), app, count=1, flags=re.S)
            app = re.sub(r'(<TitlesOfParts>\s*<vt:vector size=")(\d+)"', lambda m: '%s%d"' % (
                m.group(1), int(m.group(2)) - 1), app, count=1)
            arquivos["docProps/app.xml"] = app.encode("utf-8")
    return arquivos, True


def instalar_painel(arquivos, lin_hdr, P, subs, base_txt, ult=1000):
    """Remove a aba RESUMO GERENCIAL e refaz a aba PAINEL (sem os gráficos antigos), que passa a abrir primeiro."""
    arquivos, _ = remover_aba(arquivos, ABA_RESUMO)
    abas = _mapa_abas(arquivos)
    painel = next((a for a in abas if a[0] == "PAINEL"), None)
    if painel is None:
        return arquivos
    st = arquivos["xl/styles.xml"].decode("utf-8")
    xml, st = _painel_xml(st, lin_hdr, P, subs, base_txt, ult)
    arquivos["xl/styles.xml"] = st.encode("utf-8")
    _apagar_partes(arquivos, _desenhos_da_aba(arquivos, painel[2]))
    for nome in [n for n in arquivos if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]:
        x = arquivos[nome].decode("utf-8")
        x2 = re.sub(r'\stabSelected="1"', "", x)
        if x2 != x:
            arquivos[nome] = x2.encode("utf-8")
    arquivos[painel[2]] = xml.encode("utf-8")
    wbx = arquivos["xl/workbook.xml"].decode("utf-8")
    idx = [a[0] for a in _mapa_abas(arquivos)].index("PAINEL")
    wbx = re.sub(r'\sactiveTab="\d+"', "", wbx)
    wbx = re.sub(r'\sfirstSheet="\d+"', "", wbx)
    wbx = re.sub(r"<workbookView\b", '<workbookView activeTab="%d"' % idx, wbx, count=1)
    arquivos["xl/workbook.xml"] = wbx.encode("utf-8")
    return arquivos


def categorias_parametros(arquivos):
    """Lista de CATEGORIA na aba PARÂMETROS (F13...) e validação da coluna CATEGORIA no CONTROLE."""
    abas = {a[0]: a[2] for a in _mapa_abas(arquivos)}
    if "PARÂMETROS" not in abas:
        return arquivos
    x = arquivos[abas["PARÂMETROS"]].decode("utf-8")
    est = _estilos_linha(x, 13)
    s13 = est.get("F") or est.get("E")
    ed = {13 + i: {"F": (lambda v: lambda ref, s: _cel_valor(ref, ' s="%s"' % s13 if s13 else s, v))(cat)}
          for i, cat in enumerate(CATEG_LISTA)}
    arquivos[abas["PARÂMETROS"]] = _patch_linhas(x, ed, {}).encode("utf-8")
    fim = 12 + len(CATEG_LISTA)
    for nome in ("CONTROLE",):
        if nome in abas:
            xc = arquivos[abas[nome]].decode("utf-8")
            xc2 = re.sub(r"('?PARÂMETROS'?!\$F\$13:\$F\$)\d+", r"\g<1>%d" % fim, xc)
            xc2 = re.sub(r"Escolha PRIORIDADE ou OPORTUNIDADE\.", "Escolha PRIORIDADE 1, PRIORIDADE 2 ou OPORTUNIDADE.", xc2)
            arquivos[abas[nome]] = xc2.encode("utf-8")
    return arquivos


def _painel_no_arquivo(caminho, lin_hdr, P, subs, base_txt, ult=1000):
    z = zipfile.ZipFile(caminho)
    arquivos = {n: z.read(n) for n in z.namelist()}
    infos = {i.filename: i for i in z.infolist()}
    z.close()
    arquivos = instalar_painel(arquivos, lin_hdr, P, subs, base_txt, ult)
    temp = caminho + ".pn.xlsx"
    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, data in arquivos.items():
            zi = infos.get(n) or zipfile.ZipInfo(n)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zo.writestr(zi, data)
    os.replace(temp, caminho)


def atualizar(painel, controle, pcp=None, backup=None, calendario=None, so_formulas=False):
    """so_formulas: reescreve só as fórmulas e o topo (semana/subtotais); não mexe na BASE_OLODUM nem nos AVANÇOS."""
    if so_formulas:
        d, base, pcp = None, [], None
        print("Modo só fórmulas: BASE_OLODUM e AVANÇOS ficam como estão.")
    else:
        painel = mais_recente(painel, ["*OLODUM*.html", "*olodum*.html"])
        print("Painel OLODUM: %s (salvo em %s)" % (painel, _quando(painel)))
        if pcp:
            pcp = mais_recente(pcp, ["OLODUM_BASEDADOS*.xlsx", "OLODUM_BASEDADOS*.xlsm"])
            print("Base PCP:      %s (salva em %s)" % (pcp, _quando(pcp)))
        d = ler_painel(painel)
        base = base_por_pasta(d)
    print("Controle:      %s" % controle)

    # leitura (somente leitura: nada é gravado pelo openpyxl)
    import warnings
    warnings.filterwarnings("ignore", module="openpyxl")
    wbr = openpyxl.load_workbook(controle, data_only=True)  # só leitura: este objeto nunca é gravado
    abas = {_norm(n): n for n in wbr.sheetnames}
    for aba in ("CONTROLE", "BASE_OLODUM", "PARÂMETROS"):
        if _norm(aba) not in abas:
            raise SystemExit("A planilha não tem a aba %s (abas: %s)" % (aba, ", ".join(wbr.sheetnames)))
    cal = ler_calendario(calendario if calendario else wbr)
    ws = wbr[abas["CONTROLE"]]
    # linha do cabeçalho: a primeira (1 a 15) que tenha "Nº DA PASTA"
    lin_hdr = None
    for r in range(1, 16):
        if any(_norm(ws.cell(r, c).value) == "Nº DA PASTA" for c in range(1, min(ws.max_column, 60) + 1)):
            lin_hdr = r
            break
    if lin_hdr is None:
        print("Não achei o cabeçalho 'Nº DA PASTA' no CONTROLE. Conteúdo das primeiras linhas:")
        for r in range(1, 8):
            vals = [str(ws.cell(r, c).value) for c in range(1, 8) if ws.cell(r, c).value is not None]
            print("  linha %d: %s" % (r, " | ".join(v[:40] for v in vals)))
        raise SystemExit("Cabeçalho 'Nº DA PASTA' não encontrado no CONTROLE")
    if lin_hdr != LIN_CAB:
        print("Cabeçalho do CONTROLE na linha %d." % lin_hdr)
    hdr = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(lin_hdr, c).value
        if v is not None and _norm(v) not in hdr:
            hdr[_norm(v)] = get_column_letter(c)
    col_pasta = hdr["Nº DA PASTA"]
    ult_lin = ws.max_row
    pastas = {}
    subs = set()
    col_sub = hdr.get("SUBSOP HMS")
    for r in range(lin_hdr + 1, ws.max_row + 1):
        v = ws["%s%d" % (col_pasta, r)].value
        if v and "NÃO CADASTRADA" not in str(v).upper():
            pastas[r] = str(v).strip()
            if col_sub:
                subs.add(_subsistema_principal(ws["%s%d" % (col_sub, r)].value))
    avancos_ant = {}
    if not pcp and not so_formulas:
        for a in AVANCOS:
            L = hdr.get(_norm(a))
            if L:
                for r in pastas:
                    avancos_ant.setdefault(r, {})[a] = ws["%s%d" % (L, r)].value
    wp = wbr[abas["PARÂMETROS"]]
    param_ok = "SEMANA VIGENTE" in _norm(wp["A8"].value) and "MÍNIMO" in _norm(wp["A4"].value)
    alvo_nomes = {"CONTROLE": abas["CONTROLE"], "BASE_OLODUM": abas["BASE_OLODUM"]}
    wbr.close()
    # cabeçalho da BASE_OLODUM (a primeira linha com PASTA na coluna A)
    wsb = wbr[abas["BASE_OLODUM"]]
    lin_b = next((r for r in range(1, 16) if _norm(wsb.cell(r, 1).value) == "PASTA"), None)
    if lin_b is None:
        raise SystemExit("Cabeçalho 'PASTA' não encontrado na coluna A da aba BASE_OLODUM")

    layout_ok = all(hdr.get(_norm(nome)) == C[nome] for nome, _, _ in CTRL)
    if not layout_ok:
        dif = [nome for nome, _, _ in CTRL if hdr.get(_norm(nome)) != C[nome]]
        print("AVISO: colunas fora da posição original (%s). As fórmulas não serão reescritas; "
              "só a BASE_OLODUM e as colunas de AVANÇO." % ", ".join(dif[:6]))
    if not param_ok:
        print("AVISO: aba PARÂMETROS diferente da original; as fórmulas não serão reescritas.")

    # pastas novas do painel entram no fim
    existentes = set(pastas.values())
    prox = max(pastas) + 1 if pastas else lin_hdr + 1
    novas = []
    for b in base:
        if b["PASTA"] not in existentes:
            subs.add(_subsistema_principal(b["SUBSISTEMA"]))
            pastas[prox] = b["PASTA"]
            novas.append((prox, b))
            prox += 1
    ult_lin = max(ult_lin, prox - 1)

    # avanços
    linhas = [dict({"Nº DA PASTA": p}, **avancos_ant.get(r, {})) for r, p in sorted(pastas.items())]
    if so_formulas:
        pass
    elif pcp:
        av = avancos_pcp(ler_pcp(pcp), base, cal)
        for l in linhas:
            for col in AVANCOS:
                l[col] = av.get(l["Nº DA PASTA"], {}).get(col)
    else:
        gravar_avancos(linhas, base, d, cal)
    av_por_linha = {} if so_formulas else dict(zip(sorted(pastas), linhas))

    # ---- edição do XML
    z = zipfile.ZipFile(controle)
    arquivos = {n: z.read(n) for n in z.namelist()}
    infos = {i.filename: i for i in z.infolist()}
    z.close()
    wbxml = arquivos["xl/workbook.xml"].decode("utf-8")
    rels = arquivos["xl/_rels/workbook.xml.rels"].decode("utf-8")
    alvo = {}
    for m in re.finditer(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wbxml):
        t = re.search(r'<Relationship\b[^>]*Id="%s"[^>]*Target="([^"]+)"' % m.group(2), rels) or \
            re.search(r'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="%s"' % m.group(2), rels)
        alvo[m.group(1).replace("&amp;", "&")] = "xl/" + t.group(1).lstrip("/").replace("xl/", "")
    alvo["CONTROLE"], alvo["BASE_OLODUM"] = alvo[alvo_nomes["CONTROLE"]], alvo[alvo_nomes["BASE_OLODUM"]]
    x_ctrl = arquivos[alvo["CONTROLE"]].decode("utf-8")
    x_base = arquivos[alvo["BASE_OLODUM"]].decode("utf-8")

    # BASE_OLODUM: substitui as linhas de dados
    if so_formulas:
        pass
    else:
        est_base = _estilos_linha(x_base, lin_b + 1) or _estilos_linha(x_base, lin_b)
        letras_b = [get_column_letter(i) for i in range(1, len(BASE_COLS) + 1)]
        novas_b = {}
        for i, b in enumerate(base, lin_b + 1):
            novas_b[i] = {L: (lambda v: lambda ref, s: _cel_valor(ref, s, v))(b[k]) for L, k in zip(letras_b, BASE_COLS)}
        x_base = _patch_linhas(x_base, novas_b, est_base, remover_apos=lin_b)
        ult_b = lin_b + len(base)
        x_base = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:%s%d"/>' % (letras_b[-1], ult_b), x_base)
        x_base = re.sub(r'(<autoFilter\b[^>]*\bref=")[^"]*"', r'\g<1>A%d:%s%d"' % (lin_b, letras_b[-1], ult_b), x_base)
        x_base = re.sub(r"Importada do Painel OLODUM em [^.<]*", "Importada do Painel OLODUM em %s"
                        % d["_data"].strftime("%d/%m/%Y %H:%M UTC"), x_base)

    # CONTROLE: fórmulas das colunas automáticas + avanços + pastas novas
    auto = [C[n] for n, t, _ in CTRL if t in ("auto", "progauto", "pcpauto", "aux")] if (layout_ok and param_ok) else []
    grupos = _grupos_compartilhados(x_ctrl)
    intocaveis = {si for si, cols in grupos.items() if not cols <= set(auto)}
    if intocaveis:
        # células de grupos compartilhados que atravessam colunas manuais: não mexer nessas fórmulas
        cel_intoc = set()
        for raw in _RE_C.findall(x_ctrl):
            m = re.search(r'<f\b[^>]*\bsi="(\d+)"', raw)
            if m and m.group(1) in intocaveis:
                cel_intoc.add(_RE_ATTR("r").search(raw).group(1))
    else:
        cel_intoc = set()
    est_ctrl = _estilos_linha(x_ctrl, lin_hdr + 1)
    P = {"min": "'PARÂMETROS'!$B$4", "hoje": "'PARÂMETROS'!$B$5", "sv": "'PARÂMETROS'!$B$8",
         "svt": "'PARÂMETROS'!$B$9"}
    nome_por_col = {C[n]: n for n, _, _ in CTRL}
    edits, n_form = {}, 0
    for r in range(lin_hdr + 1, ult_lin + 1):
        e = {}
        if auto:
            f = f_controle(r, P, 0, lin_hdr)
            for L in auto:
                if "%s%d" % (L, r) in cel_intoc:
                    continue
                e[L] = (lambda fx: lambda ref, s: _cel_formula(ref, s, fx))(f[nome_por_col[L]])
                n_form += 1
        if r in av_por_linha:
            for a in AVANCOS:
                L = hdr.get(_norm(a))
                if L:
                    e[L] = (lambda v: lambda ref, s: _cel_valor(ref, s, v))(av_por_linha[r].get(a))
        if e:
            edits[r] = e
    for r, b in novas:
        edits.setdefault(r, {})
        edits[r][col_pasta] = (lambda v: lambda ref, s: _cel_valor(ref, s, v))(b["PASTA"])
        if hdr.get("SUBSOP HMS"):
            edits[r][hdr["SUBSOP HMS"]] = (lambda v: lambda ref, s: _cel_valor(ref, s, v))(b["SUBSISTEMA"])
    x_ctrl = _patch_linhas(x_ctrl, edits, est_ctrl)
    topo = False
    if auto:  # semana vigente + subtotais acima do cabeçalho (quando há linhas livres acima dele)
        st = arquivos["xl/styles.xml"].decode("utf-8")
        x_ctrl, st, topo = _topo_controle(x_ctrl, st, lin_hdr, ult_lin, P)
        arquivos["xl/styles.xml"] = st.encode("utf-8")
    if auto:  # a coluna do certificado deixou de mostrar SIM/NÃO: a cor verde passa para APROVADA
        x_ctrl = re.sub(r'(<conditionalFormatting\b[^>]*sqref="%s\d+:%s\d+"[^>]*>(?:(?!</conditionalFormatting>).)*?<formula>)(?:"SIM"|&quot;SIM&quot;)(</formula>)'
                        % (C["CERTIFICADO TESTE HMSWEB"], C["CERTIFICADO TESTE HMSWEB"]),
                        r'\1"APROVADA"\2', x_ctrl, flags=re.S)

    arquivos[alvo["CONTROLE"]] = x_ctrl.encode("utf-8")
    arquivos[alvo["BASE_OLODUM"]] = x_base.encode("utf-8")
    # fórmulas que apontavam para um trecho fixo da BASE passam a usar a coluna inteira (todas as abas)
    for nome, caminho in alvo.items():
        x = arquivos[caminho].decode("utf-8")
        x2 = re.sub(r"('?BASE_OLODUM'?!)\$([A-Z]+)\$\d+:\$\2\$\d+", r"\1$\2:$\2", x)
        if nome == "PAINEL":
            if d:
                x2 = re.sub(r"Base Painel OLODUM: \d\d/\d\d/\d{4} \d\d:\d\d UTC",
                            "Base Painel OLODUM: %s" % d["_data"].strftime("%d/%m/%Y %H:%M UTC"), x2)
            # TEXT(data;"dd/mm/yyyy") mostra "yyyy" no Excel em português
            x2 = re.sub(r"TEXT\(('?PARÂMETROS'?!\$B\$5),(?:\"|&quot;)dd/mm/yyyy(?:\"|&quot;)\)",
                        lambda m: _esc(_data_txt(m.group(1))).replace("&quot;", '"'), x2)
        if x2 != x:
            arquivos[caminho] = x2.encode("utf-8")
    # recálculo completo ao abrir; descarta a cadeia de cálculo antiga
    if "<calcPr" in wbxml:
        wbxml = re.sub(r"<calcPr\b([^>]*?)/?>", lambda m: "<calcPr%s fullCalcOnLoad=\"1\"/>"
                       % re.sub(r'\sfullCalcOnLoad="[^"]*"', "", m.group(1)).rstrip("/"), wbxml, count=1)
    else:
        pos = wbxml.index("</definedNames>") + len("</definedNames>") if "</definedNames>" in wbxml \
            else wbxml.index("</sheets>") + len("</sheets>")
        wbxml = wbxml[:pos] + '<calcPr fullCalcOnLoad="1"/>' + wbxml[pos:]
    arquivos["xl/workbook.xml"] = wbxml.encode("utf-8")
    if auto:
        base_txt = d["_data"].strftime("%d/%m/%Y %H:%M UTC") if d else None
        if not base_txt:  # modo só fórmulas: mantém a data do painel que já estava no PAINEL
            for a in _mapa_abas(arquivos):
                if a[0] == "PAINEL":
                    m = re.search(r"Base Painel OLODUM: (\d\d/\d\d/\d{4} \d\d:\d\d UTC)",
                                  arquivos[a[2]].decode("utf-8"))
                    base_txt = m.group(1) if m else "-"
        arquivos = categorias_parametros(arquivos)
        arquivos = instalar_painel(arquivos, lin_hdr, P, sorted(subs - {""}), base_txt, ult_lin)
    if "xl/calcChain.xml" in arquivos:
        del arquivos["xl/calcChain.xml"]
        arquivos["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', "",
                                                        arquivos["xl/_rels/workbook.xml.rels"].decode("utf-8")).encode("utf-8")
        ct = arquivos["[Content_Types].xml"].decode("utf-8")
        arquivos["[Content_Types].xml"] = re.sub(r'<Override\b[^>]*calcChain[^>]*/>', "", ct).encode("utf-8")

    if backup:
        os.makedirs(backup, exist_ok=True)
        nome, ext = os.path.splitext(os.path.basename(controle))
        copia = os.path.join(backup, "%s_%s%s" % (nome, dt.datetime.now().strftime("%Y%m%d_%H%M"), ext))
        shutil.copy2(controle, copia)
        print("Cópia de segurança: %s" % copia)
    temp = controle + ".tmp.xlsx"
    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, data in arquivos.items():
            zi = infos.get(n) or zipfile.ZipInfo(n)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zo.writestr(zi, data)
    try:
        os.replace(temp, controle)
    except PermissionError:
        os.remove(temp)
        raise SystemExit("ERRO: não consegui gravar %s. A planilha está aberta no Excel (sua ou de outra pessoa)? "
                         "Feche e rode de novo." % controle)
    if d:
        sem = semana_da_data(d["_data"], cal)
        cont = collections.Counter(a for l in linhas for a in AVANCOS if l.get(a))
        print("Painel de %s (semana %s) | pastas: %d (%d novas: %s)" % (
            d["_data"].strftime("%d/%m/%Y %H:%M UTC"), txt_semana(sem), len(pastas), len(novas),
            ", ".join(b["PASTA"] for _, b in novas) or "-"))
        print("Avanços: " + ", ".join("%s %d" % (a, cont[a]) for a in AVANCOS))
    print("Fórmulas atualizadas: %d | BASE_OLODUM: %s | topo (semana e subtotais): %s"
          % (n_form, "%d pastas" % len(base) if d else "mantida", "sim" if topo else "não"))
    print("Atualizado:", controle)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("painel")
    ap.add_argument("controle")
    ap.add_argument("saida", nargs="?", help="omitir com --atualizar")
    ap.add_argument("--so-formulas", action="store_true",
                    help="com --atualizar: reescreve só fórmulas e topo, mantendo BASE_OLODUM e AVANÇOS")
    ap.add_argument("--atualizar", action="store_true",
                    help="atualiza só as colunas automáticas e a BASE_OLODUM no próprio CONTROLE (mantém o resto)")
    ap.add_argument("--calendario")
    ap.add_argument("--painel-anterior")
    ap.add_argument("--pcp")
    ap.add_argument("--backup", help="pasta onde guardar uma cópia do controle antes de sobrescrever")
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if a.atualizar:
        atualizar(a.painel, a.controle, a.pcp, a.backup, a.calendario, a.so_formulas)
    else:
        if not a.saida:
            ap.error("informe SAIDA ou use --atualizar")
        gerar(a.painel, a.controle, a.saida, a.calendario, a.painel_anterior, a.pcp, a.backup)
