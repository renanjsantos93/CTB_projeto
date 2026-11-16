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
ESC_RADIX, ESC_FURUI, ESC_INTRA = "RADIX", "FURUI", "FURUI - INTRA SKID"


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
    intra_sem_ctb = b["ESCOPO"] == ESC_INTRA and not (b["JUNTAS PREVISTAS"] or 0) > 0
    return {
        "AVANÇO MONTAGEM": b["MONTAGEM (LVM)"] == "APROVADA",
        "AVANÇO TESTE": b["TESTE HIDROSTÁTICO"] == "APROVADA",
        "AVANÇO REMONTAGEM": b["REMONTAGEM"] == "APROVADA",
        "AVANÇO PRÉ WALK": intra_sem_ctb and ok(b["MAPA JUNTAS"]) and b["TESTE HIDROSTÁTICO"] == "APROVADA"
                           and ok(b["LAVAGEM"]),
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
        if b and b["ESCOPO"] == ESC_INTRA and not (b["JUNTAS PREVISTAS"] or 0) > 0:
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
    return re.sub(r"\s+", " ", str(h or "")).strip().upper()


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
          "N° DA PASTA - PARA": "Nº DA PASTA", "DATA ENVIO": None}

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

STATUS_LISTA = ["EM MONTAGEM", "P/TESTE", "P/REMONTAGEM", "PRÉ WALKDOWN", "PENDENTE DOCUMENTAÇÃO HMSWEB",
                "CONCLUÍDA", "SEM CTB", "SEM ESCOPO", "NÃO ENCONTRADA"]
COR_STATUS = {"EM MONTAGEM": "FFEB9C", "P/TESTE": "C6EFCE", "P/REMONTAGEM": "DDEBF7", "PRÉ WALKDOWN": "B4C6E7",
              "PENDENTE DOCUMENTAÇÃO HMSWEB": "F4B183", "CONCLUÍDA": "A9D08E", "SEM CTB": "EDEDED",
              "SEM ESCOPO": "D9D9D9", "NÃO ENCONTRADA": "F8CBAD"}
# etapa da programação -> (coluna da semana programada, status, coluna de avanço)
ETAPAS = [("MONTAGEM", "SEMANA PROG MONTAGEM", "EM MONTAGEM", "AVANÇO MONTAGEM"),
          ("TESTE", "SEMANA PROG TESTE", "P/TESTE", "AVANÇO TESTE"),
          ("REMONTAGEM", "SEMANA PROG REMONTAGEM", "P/REMONTAGEM", "AVANÇO REMONTAGEM"),
          ("PRÉ WALK", "SEMANA PROG PRÉ WALK", "PRÉ WALKDOWN", "AVANÇO PRÉ WALK")]
CATEG_LISTA = ["PRIORIDADE", "OPORTUNIDADE"]
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
        "STATUS CALCULADO - escopo RADIX e FURUI (e INTRA SKID com juntas cadastradas no ControlTub):",
        "  EM MONTAGEM: avanço ControlTub < avanço mínimo (B4), mesmo com teste aprovado no HMSWEB.",
        "  P/TESTE: avanço ControlTub >= avanço mínimo e sem certificado de teste aprovado no HMSWEB.",
        "  P/REMONTAGEM: teste aprovado no HMSWEB e LV de montagem e/ou remontagem pendente.",
        "  CONCLUÍDA: teste, LV de montagem e remontagem aprovados.",
        "  SEM CTB: sem juntas no ControlTub e sem teste aprovado.",
        "STATUS CALCULADO - escopo FURUI - INTRA SKID (sem juntas no ControlTub):",
        "  PENDENTE DOCUMENTAÇÃO HMSWEB: algum certificado (exceto LV de montagem e remontagem) não aprovado.",
        "  PRÉ WALKDOWN: demais certificados aprovados; pendente LV de montagem e/ou remontagem.",
        "  CONCLUÍDA: todos os certificados aprovados.",
        "  Atividade inexistente na PCP para a pasta (N/A) não é exigida.",
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


def f_controle(r, P, nb):
    b = lambda col: "'BASE_OLODUM'!$%s$%d:$%s$%d" % (col, LIN_CAB + 1, col, nb)
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
    cat = 'IF(%(c)s="PRIORIDADE",2000,IF(%(c)s="OPORTUNIDADE",1000,0))' % {"c": x("CATEGORIA DE ENTREGA")}
    return {
        "ITEM": '=IF(%s="","",ROW()-%d)' % (pasta, LIN_CAB),
        "SISTEMA": look("B"), "ESCOPO": look("D"), "FORNECEDOR": look("E"),
        "JUNTAS PREVISTAS CTB": look_num("F"), "JUNTAS LIBERADAS CTB": look_num("G"),
        "SALDO DE JUNTAS": '=IF(ISNUMBER(%s),%s-%s,"")' % (x("JUNTAS PREVISTAS CTB"), x("JUNTAS PREVISTAS CTB"),
                                                           x("JUNTAS LIBERADAS CTB")),
        "AVANÇO CTB": '=IF(N(%s)>0,%s/%s,"")' % (x("JUNTAS PREVISTAS CTB"), x("JUNTAS LIBERADAS CTB"),
                                               x("JUNTAS PREVISTAS CTB")),
        "MONTAGEM HMS (LVM)": look("H"), "MAPA JUNTAS HMS": look("I"),
        "CERTIFICADO TESTE HMSWEB": '=IF(%s="","",IFERROR(IF(INDEX(%s,%s)="APROVADA","SIM","NÃO"),""))'
                                    % (pasta, b("J"), m),
        "LAVAGEM HMS": look("K"), "REMONTAGEM HMS": look("L"), "DATA PLANEJADA TH": look_num("M"),
        "PEND. A TUB. SUBSISTEMA": look("N", texto=False),
        "STATUS CALCULADO": (
            '=IF(%(p)s="","",IF(ISNA(%(m)s),"NÃO ENCONTRADA",'
            'IF(OR(%(e)s="%(radix)s",%(e)s="%(furui)s",AND(%(e)s="%(intra)s",N(%(prev)s)>0)),'
            'IF(AND(N(%(prev)s)>0,N(%(av)s)<%(min)s),"EM MONTAGEM",'
            'IF(%(th)s="SIM",IF(AND(%(lvm)s="APROVADA",%(rem)s="APROVADA"),"CONCLUÍDA","P/REMONTAGEM"),'
            'IF(%(av)s="","SEM CTB","P/TESTE"))),'
            'IF(%(e)s="%(intra)s",IF(AND(OR(%(mj)s="APROVADA",%(mj)s="N/A"),%(th)s="SIM",'
            'OR(%(lav)s="APROVADA",%(lav)s="N/A")),IF(AND(OR(%(lvm)s="APROVADA",%(lvm)s="N/A"),'
            'OR(%(rem)s="APROVADA",%(rem)s="N/A")),"CONCLUÍDA","PRÉ WALKDOWN"),"PENDENTE DOCUMENTAÇÃO HMSWEB"),'
            '"SEM ESCOPO"))))'
            % {"p": pasta, "m": m, "e": x("ESCOPO"), "radix": ESC_RADIX, "furui": ESC_FURUI, "intra": ESC_INTRA,
               "prev": x("JUNTAS PREVISTAS CTB"), "th": x("CERTIFICADO TESTE HMSWEB"), "rem": x("REMONTAGEM HMS"),
               "lvm": x("MONTAGEM HMS (LVM)"), "mj": x("MAPA JUNTAS HMS"), "lav": x("LAVAGEM HMS"),
               "av": x("AVANÇO CTB"), "min": P["min"]}),
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
    for nome in ("MONTAGEM HMS (LVM)", "MAPA JUNTAS HMS", "LAVAGEM HMS", "REMONTAGEM HMS"):
        ws.conditional_formatting.add(rng(nome), CellIsRule(operator="equal", formula=['"APROVADA"'],
                                      font=Font(name=FONTE, size=9, color="006100")))
        ws.conditional_formatting.add(rng(nome), CellIsRule(operator="equal", formula=['"REPROVADA"'],
                                      fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng("CERTIFICADO TESTE HMSWEB"), CellIsRule(operator="equal", formula=['"SIM"'],
                                  fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONTE, size=9, bold=True, color="006100")))
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
                'TEXT(%s,"0%%")&"   |   Atualizado em "&TEXT(%s,"dd/mm/yyyy")' % (gerado_em, P["svt"], P["min"], P["hoje"]))
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
    try:
        os.replace(temp, saida)
    except PermissionError:
        os.remove(temp)
        raise SystemExit("ERRO: não consegui gravar %s. A planilha está aberta no Excel (sua ou de outra pessoa)? "
                         "Feche e rode de novo." % saida)
    print("Painel de %s (semana %s) | pastas no controle: %d (%d novas)" % (gerado_em, txt_semana(sem), len(linhas), len(novas)))
    print("Gravado:", saida)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("painel")
    ap.add_argument("controle")
    ap.add_argument("saida")
    ap.add_argument("--calendario")
    ap.add_argument("--painel-anterior")
    ap.add_argument("--pcp")
    ap.add_argument("--backup", help="pasta onde guardar uma cópia do controle antes de sobrescrever")
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    gerar(a.painel, a.controle, a.saida, a.calendario, a.painel_anterior, a.pcp, a.backup)
