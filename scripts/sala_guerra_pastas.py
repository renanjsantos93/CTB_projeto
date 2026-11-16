"""Gera o CONTROLE DE PASTAS - SALA DE GUERRA (tubulação) a partir do Painel OLODUM.

Uso:
    python scripts/sala_guerra_pastas.py PAINEL_OLODUM.html CONTROLE.xlsx SAIDA.xlsx

PAINEL_OLODUM.html  painel exportado (traz embutidas a PCP/HMSWEB, as juntas do
                    ControlTub, as pendências e o escopo das pastas).
CONTROLE.xlsx       controle anterior. Pode ser o "01-CONTROLE_PASTA.xlsx" original
                    (aba CONTROLE DE PASTA) ou uma planilha gerada por este script:
                    as colunas de preenchimento manual (programação, status da sala
                    de guerra, avanço Milplan, data de envio, responsável, observação)
                    e a aba RESTRIÇÕES são preservadas.

A cada atualização do painel basta rodar de novo apontando o controle mais recente.
"""
import argparse
import collections
import datetime as dt
import json
import re
import sys

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
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


def ler_painel(caminho):
    linhas = open(caminho, encoding="utf-8").read().split("\n")
    dados = {}
    for s in linhas:
        m = re.match(r"\s*window\.(__\w+__)\s*=\s*(.*?);?\s*$", s)
        if m and m.group(2).startswith(("[", "{")):
            dados[m.group(1)] = json.loads(m.group(2))
    for k in ("__DADOS__", "__JUNTAS__", "__PENDLISTA__", "__ESCOPOS__"):
        if k not in dados:
            raise SystemExit("Painel sem a base %s" % k)
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

    pastas = sorted(set(atv) | set(juntas))
    base = []
    for p in pastas:
        inf = info[p]
        jt = juntas.get(p)
        sis = inf["sis"].most_common(1)[0][0] if inf["sis"] else (jt[2] if jt else "")
        sub = inf["sub"].most_common(1)[0][0] if inf["sub"] else (jt[3] if jt else "")
        a = atv.get(p, {})
        base.append({
            "PASTA": p,
            "SISTEMA": sis,
            "SUBSISTEMA": sub,
            "ESCOPO": escopo.get(p, ""),
            "FORNECEDOR": inf["forn"].most_common(1)[0][0] if inf["forn"] else "",
            "JUNTAS PREVISTAS": jt[0] if jt else None,
            "JUNTAS LIBERADAS": jt[1] if jt else None,
            "MONTAGEM (LVM)": situacao(a.get(ATV_LVM)),
            "MAPA JUNTAS": situacao(a.get(ATV_MJ)),
            "TESTE HIDROSTÁTICO": situacao(a.get(ATV_TH)),
            "LAVAGEM": situacao(a.get(ATV_LAV)),
            "REMONTAGEM": situacao(a.get(ATV_REM)),
            "DATA PLANEJADA TH": dt.date.fromisoformat(inf["data_th"]) if inf["data_th"] else None,
            "PEND. A TUB. SUBSISTEMA": pend_a.get(sub, 0),
        })
    return base


def _norm(h):
    return re.sub(r"\s+", " ", str(h or "")).strip().upper()


def ler_controle(caminho):
    """Lê o controle anterior. Devolve (lista de dicts por pasta, lista de restrições)."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    linhas, restr = [], []
    if "CONTROLE" in wb.sheetnames:  # gerado por este script
        ws = wb["CONTROLE"]
        hdr = {_norm(c.value): c.column for c in ws[LIN_CAB]}
        for r in range(LIN_CAB + 1, ws.max_row + 1):
            p = ws.cell(r, hdr["Nº DA PASTA"]).value
            if not p:
                continue
            linhas.append({k: ws.cell(r, hdr[_norm(k)]).value for k in
                           ("Nº DA PASTA", "SUBSOP HMS", "PROGRAMAÇÃO", "SEMANA PROG.", "STATUS SALA DE GUERRA",
                            "AVANÇO FÍSICO MILPLAN", "DATA ENVIO", "RESPONSÁVEL", "OBSERVAÇÃO") if _norm(k) in hdr})
        if "RESTRIÇÕES" in wb.sheetnames:
            wr = wb["RESTRIÇÕES"]
            for r in range(LIN_CAB + 1, wr.max_row + 1):
                vals = [wr.cell(r, c).value for c in range(1, len(COL_RESTR) + 1)]
                if vals[1]:
                    restr.append(vals)
    else:  # formato original 01-CONTROLE_PASTA.xlsx
        ws = wb["CONTROLE DE PASTA"] if "CONTROLE DE PASTA" in wb.sheetnames else wb.worksheets[0]
        hdr = {_norm(c.value): c.column for c in ws[1]}
        col = lambda nome: hdr.get(_norm(nome))
        for r in range(2, ws.max_row + 1):
            p = ws.cell(r, col("N° DA PASTA - PARA")).value
            if not p:
                continue
            g = lambda nome: ws.cell(r, col(nome)).value if col(nome) else None
            linhas.append({"Nº DA PASTA": str(p).strip(), "SUBSOP HMS": g("SUBSOP HMS"),
                           "PROGRAMAÇÃO": g("PRIORIDADE"), "STATUS SALA DE GUERRA": g("STATUS"),
                           "DATA ENVIO": g("DATA ENVIO"), "OBSERVAÇÃO": g("OBSERVAÇÃO")})
    return linhas, restr


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
FILL_INP = PatternFill("solid", fgColor="FFF2CC")
FILL_KPI = PatternFill("solid", fgColor="EEF3FA")
FILL_ZEBRA = PatternFill("solid", fgColor="F7F9FC")
FINO = Side(style="thin", color="D0D7E2")
BORDA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
ESQ = Alignment(horizontal="left", vertical="center")

LIN_CAB = 4          # linha do cabeçalho das abas de dados
MAX_LIN = 600        # linhas com fórmula prontas no CONTROLE (cabem pastas novas)
MAX_RESTR = 500

STATUS_LISTA = ["EM MONTAGEM", "P/TESTE", "P/REMONTAGEM", "PRÉ WALKDOWN", "PENDENTE DOCUMENTAÇÃO HMSWEB",
                "CONCLUÍDA", "SEM CTB", "SEM ESCOPO", "NÃO ENCONTRADA"]
COR_STATUS = {"EM MONTAGEM": "FFEB9C", "P/TESTE": "C6EFCE", "P/REMONTAGEM": "DDEBF7", "PRÉ WALKDOWN": "B4C6E7",
              "PENDENTE DOCUMENTAÇÃO HMSWEB": "F4B183", "CONCLUÍDA": "A9D08E", "SEM CTB": "EDEDED",
              "SEM ESCOPO": "D9D9D9", "NÃO ENCONTRADA": "F8CBAD"}
ESC_RADIX, ESC_FURUI, ESC_INTRA = "RADIX", "FURUI", "FURUI - INTRA SKID"
PROG_LISTA = ["PRIORIDADE", "OPORTUNIDADE"]
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


# ----------------------------------------------------------------------------- abas
def aba_parametros(wb):
    ws = wb.create_sheet("PARÂMETROS")
    titulo(ws, "PARÂMETROS E LISTAS", "Células amarelas podem ser alteradas; todas as abas usam estes valores.")
    ws["A4"], ws["B4"] = "Avanço mínimo do ControlTub para liberar teste", 0.8
    ws["A5"], ws["B5"] = "Tolerância de desvio Milplan × ControlTub (pontos percentuais)", 0.05
    ws["A6"], ws["B6"] = "Data de referência (hoje)", "=TODAY()"
    for r in (4, 5):
        ws.cell(r, 2).fill, ws.cell(r, 2).font = FILL_INP, F_INP
        ws.cell(r, 2).number_format = "0%"
    ws["B6"].number_format = "dd/mm/yyyy"
    ws["C4"] = "Critério do e-mail da sala de guerra: pasta com avanço do CTB igual ou superior a 80% vai para teste."
    ws["C5"] = "Acima desta diferença entre o avanço informado pela Milplan e o ControlTub a pasta é marcada como DIVERGENTE."
    for r in range(4, 7):
        ws.cell(r, 1).font = Font(name=FONTE, size=9, bold=True)
        ws.cell(r, 3).font = F_SUB
    listas = [("STATUS", STATUS_LISTA), ("PROGRAMAÇÃO", PROG_LISTA), ("TIPO RESTRIÇÃO", TIPOS_RESTR),
              ("STATUS RESTRIÇÃO", STATUS_RESTR)]
    for j, (nome, itens) in enumerate(listas):
        col = 5 + j
        c = ws.cell(9, col, nome)
        c.font, c.fill, c.alignment = F_CAB, FILL_CAB, CENTRO
        for i, v in enumerate(itens):
            ws.cell(10 + i, col, v).font = F_TXT
    ws["A9"] = "REGRA DO STATUS CALCULADO (coluna STATUS CALCULADO do CONTROLE)"
    ws["A9"].font = F_H2
    regras = [
        "ESCOPO RADIX e FURUI (e INTRA SKID com junta liberada no ControlTub):",
        "  EM MONTAGEM: avanço ControlTub < avanço mínimo (B4).",
        "  P/TESTE: avanço ControlTub >= avanço mínimo e sem certificado de teste aprovado no HMSWEB.",
        "  P/REMONTAGEM: teste aprovado no HMSWEB e LV de montagem e/ou remontagem pendente.",
        "  CONCLUÍDA: teste, LV de montagem e remontagem aprovados.",
        "  SEM CTB: sem juntas no ControlTub e sem teste aprovado.",
        "ESCOPO FURUI - INTRA SKID (sem junta liberada no ControlTub):",
        "  PENDENTE DOCUMENTAÇÃO HMSWEB: algum certificado (exceto LV de montagem e remontagem) não aprovado.",
        "  PRÉ WALKDOWN: demais certificados aprovados; pendente LV de montagem e/ou remontagem.",
        "  CONCLUÍDA: todos os certificados aprovados.",
        "OUTROS: SEM ESCOPO (pasta sem escopo na lista Milplan) e NÃO ENCONTRADA (fora do Painel OLODUM).",
        "Atividade que não existe na PCP para a pasta (N/A) não é exigida.",
    ]
    for i, t in enumerate(regras):
        ws.cell(10 + i, 1, t).font = F_TXT
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 4
    for col in "EFGH":
        ws.column_dimensions[col].width = 24
    n = {k: len(v) for k, v in listas}
    return {
        "min": "'PARÂMETROS'!$B$4", "tol": "'PARÂMETROS'!$B$5", "hoje": "'PARÂMETROS'!$B$6",
        "status": "'PARÂMETROS'!$E$10:$E$%d" % (9 + n["STATUS"]),
        "prog": "'PARÂMETROS'!$F$10:$F$%d" % (9 + n["PROGRAMAÇÃO"]),
        "tipo": "'PARÂMETROS'!$G$10:$G$%d" % (9 + n["TIPO RESTRIÇÃO"]),
        "sres": "'PARÂMETROS'!$H$10:$H$%d" % (9 + n["STATUS RESTRIÇÃO"]),
    }


BASE_COLS = ["PASTA", "SISTEMA", "SUBSISTEMA", "ESCOPO", "FORNECEDOR", "JUNTAS PREVISTAS", "JUNTAS LIBERADAS",
             "MONTAGEM (LVM)", "MAPA JUNTAS", "TESTE HIDROSTÁTICO", "LAVAGEM", "REMONTAGEM", "DATA PLANEJADA TH",
             "PEND. A TUB. SUBSISTEMA"]


def aba_base(wb, base, gerado_em):
    ws = wb.create_sheet("BASE_OLODUM")
    titulo(ws, "BASE PAINEL OLODUM (por pasta)",
           "Importada do Painel OLODUM em %s. Juntas = ControlTub (liberadas/previstas); atividades = PCP/HMSWEB "
           "(APROVADA, EXECUTADA, REPROVADA, PENDENTE). Não editar: é sobrescrita a cada atualização." % gerado_em)
    cab(ws, LIN_CAB, BASE_COLS)
    for i, b in enumerate(base, LIN_CAB + 1):
        for j, k in enumerate(BASE_COLS, 1):
            c = ws.cell(i, j, b[k])
            c.font, c.border = F_TXT, BORDA
            if k == "DATA PLANEJADA TH":
                c.number_format = "dd/mm/yyyy"
    larg = [17, 18, 22, 18, 14, 11, 11, 13, 13, 13, 13, 13, 12, 12]
    for j, w in enumerate(larg, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(LIN_CAB + 1, 2)
    ws.auto_filter.ref = "A%d:%s%d" % (LIN_CAB, get_column_letter(len(BASE_COLS)), LIN_CAB + len(base))
    return LIN_CAB + len(base)


# Colunas do CONTROLE: (nome, tipo, largura). tipo: inp = preenchimento manual, auto = fórmula
CTRL = [
    ("ITEM", "auto", 6), ("Nº DA PASTA", "inp", 16), ("SUBSOP HMS", "inp", 20), ("SISTEMA", "auto", 16),
    ("ESCOPO", "auto", 15), ("FORNECEDOR", "auto", 11), ("JUNTAS PREVISTAS CTB", "auto", 9),
    ("JUNTAS LIBERADAS CTB", "auto", 9), ("SALDO DE JUNTAS", "auto", 8), ("AVANÇO CTB", "auto", 10),
    ("MONTAGEM HMS (LVM)", "auto", 11), ("MAPA JUNTAS HMS", "auto", 11), ("CERTIFICADO TESTE HMSWEB", "auto", 10),
    ("LAVAGEM HMS", "auto", 11), ("REMONTAGEM HMS", "auto", 11), ("DATA PLANEJADA TH", "auto", 11),
    ("PEND. A TUB. SUBSISTEMA", "auto", 9), ("STATUS CALCULADO", "auto", 22),
    ("PROGRAMAÇÃO", "inp", 13), ("SEMANA PROG.", "inp", 9), ("STATUS SALA DE GUERRA", "inp", 14),
    ("CONFERÊNCIA STATUS", "auto", 22), ("AVANÇO FÍSICO MILPLAN", "inp", 10), ("DESVIO MILPLAN − CTB", "auto", 10),
    ("ANÁLISE AVANÇO", "auto", 14), ("RESTRIÇÕES ABERTAS", "auto", 10), ("LIBERAR P/ TESTE?", "auto", 11),
    ("DATA ENVIO", "inp", 11), ("RESPONSÁVEL", "inp", 14), ("OBSERVAÇÃO", "inp", 30),
    ("PONTOS TESTE", "aux", 9), ("PONTOS REMONT.", "aux", 9), ("PONTOS DESVIO", "aux", 9), ("ORDEM PROG.", "aux", 9),
]
C = {nome: get_column_letter(i) for i, (nome, _, _) in enumerate(CTRL, 1)}


def f_controle(r, P, nb):
    """Fórmulas da linha r do CONTROLE. nb = última linha da BASE."""
    b = lambda col: "'BASE_OLODUM'!$%s$%d:$%s$%d" % (col, LIN_CAB + 1, col, nb)
    pasta = "$%s%d" % (C["Nº DA PASTA"], r)
    m = "MATCH(%s,%s,0)" % (pasta, b("A"))

    def look(col, texto=True):
        return '=IF(%s="","",IFERROR(INDEX(%s,%s)%s,""))' % (pasta, b(col), m, '&""' if texto else "")

    x = lambda nome: "%s%d" % (C[nome], r)
    f = {
        "ITEM": '=IF(%s="","",ROW()-%d)' % (pasta, LIN_CAB),
        "SISTEMA": look("B"),
        "ESCOPO": look("D"),
        "FORNECEDOR": look("E"),
        "JUNTAS PREVISTAS CTB": '=IF(%s="","",IFERROR(IF(INDEX(%s,%s)="","",INDEX(%s,%s)),""))'
                                % (pasta, b("F"), m, b("F"), m),
        "JUNTAS LIBERADAS CTB": '=IF(%s="","",IFERROR(IF(INDEX(%s,%s)="","",INDEX(%s,%s)),""))'
                                % (pasta, b("G"), m, b("G"), m),
        "SALDO DE JUNTAS": '=IF(ISNUMBER(%s),%s-%s,"")' % (x("JUNTAS PREVISTAS CTB"), x("JUNTAS PREVISTAS CTB"),
                                                           x("JUNTAS LIBERADAS CTB")),
        "AVANÇO CTB": '=IF(N(%s)>0,%s/%s,"")' % (x("JUNTAS PREVISTAS CTB"), x("JUNTAS LIBERADAS CTB"),
                                               x("JUNTAS PREVISTAS CTB")),
        "MONTAGEM HMS (LVM)": look("H"),
        "MAPA JUNTAS HMS": look("I"),
        "CERTIFICADO TESTE HMSWEB": '=IF(%s="","",IFERROR(IF(INDEX(%s,%s)="APROVADA","SIM","NÃO"),""))'
                                    % (pasta, b("J"), m),
        "LAVAGEM HMS": look("K"),
        "REMONTAGEM HMS": look("L"),
        "DATA PLANEJADA TH": '=IF(%s="","",IFERROR(IF(INDEX(%s,%s)="","",INDEX(%s,%s)),""))'
                             % (pasta, b("M"), m, b("M"), m),
        "PEND. A TUB. SUBSISTEMA": look("N", texto=False),
        "STATUS CALCULADO": (
            '=IF(%(p)s="","",IF(ISNA(%(m)s),"NÃO ENCONTRADA",'
            'IF(OR(%(e)s="%(radix)s",%(e)s="%(furui)s",AND(%(e)s="%(intra)s",N(%(lib)s)>0)),'
            # critério RADIX / FURUI
            'IF(%(th)s="SIM",IF(AND(%(lvm)s="APROVADA",%(rem)s="APROVADA"),"CONCLUÍDA","P/REMONTAGEM"),'
            'IF(%(av)s="","SEM CTB",IF(%(av)s>=%(min)s,"P/TESTE","EM MONTAGEM"))),'
            # critério FURUI - INTRA SKID
            'IF(%(e)s="%(intra)s",IF(AND(OR(%(mj)s="APROVADA",%(mj)s="N/A"),%(th)s="SIM",'
            'OR(%(lav)s="APROVADA",%(lav)s="N/A")),IF(AND(OR(%(lvm)s="APROVADA",%(lvm)s="N/A"),'
            'OR(%(rem)s="APROVADA",%(rem)s="N/A")),"CONCLUÍDA","PRÉ WALKDOWN"),"PENDENTE DOCUMENTAÇÃO HMSWEB"),'
            '"SEM ESCOPO"))))'
            % {"p": pasta, "m": m, "e": x("ESCOPO"), "radix": ESC_RADIX, "furui": ESC_FURUI, "intra": ESC_INTRA,
               "lib": x("JUNTAS LIBERADAS CTB"), "th": x("CERTIFICADO TESTE HMSWEB"), "rem": x("REMONTAGEM HMS"),
               "lvm": x("MONTAGEM HMS (LVM)"), "mj": x("MAPA JUNTAS HMS"), "lav": x("LAVAGEM HMS"),
               "av": x("AVANÇO CTB"), "min": P["min"]}),
        "CONFERÊNCIA STATUS": '=IF(OR(%s="",%s=""),"",IF(%s=%s,"OK","DIVERGE: CALC. "&%s))'
                              % (pasta, x("STATUS SALA DE GUERRA"), x("STATUS SALA DE GUERRA"),
                                 x("STATUS CALCULADO"), x("STATUS CALCULADO")),
        "DESVIO MILPLAN − CTB": '=IF(AND(ISNUMBER(%s),ISNUMBER(%s)),%s-%s,"")'
                                % (x("AVANÇO FÍSICO MILPLAN"), x("AVANÇO CTB"), x("AVANÇO FÍSICO MILPLAN"),
                                   x("AVANÇO CTB")),
        "ANÁLISE AVANÇO": ('=IF(%(p)s="","",IF(NOT(ISNUMBER(%(mil)s)),"SEM INFO MILPLAN",IF(NOT(ISNUMBER(%(ctb)s)),'
                           '"SEM CTB",IF(ABS(%(d)s)<=%(tol)s,"ALINHADO",IF(%(d)s>0,"MILPLAN ACIMA","CTB ACIMA")))))'
                           % {"p": pasta, "mil": x("AVANÇO FÍSICO MILPLAN"), "ctb": x("AVANÇO CTB"),
                              "d": x("DESVIO MILPLAN − CTB"), "tol": P["tol"]}),
        "RESTRIÇÕES ABERTAS": '=IF(%s="","",COUNTIFS(\'RESTRIÇÕES\'!$B:$B,%s,\'RESTRIÇÕES\'!$I:$I,"<>ENCERRADA"))'
                              % (pasta, pasta),
        "LIBERAR P/ TESTE?": ('=IF(%s<>"P/TESTE","",IF(AND(%s=0,%s<>"REPROVADA",%s<>"REPROVADA"),"SIM","RESTRIÇÃO"))'
                              % (x("STATUS CALCULADO"), x("RESTRIÇÕES ABERTAS"), x("MONTAGEM HMS (LVM)"),
                                 x("MAPA JUNTAS HMS"))),
        # Pontos para ordenar o painel (maior = melhor). ROW()/1e6 desempata e mantém únicos.
        "PONTOS TESTE": ('=IF(%(s)s<>"P/TESTE",0,IF(%(l)s="SIM",5000,0)+IF(%(pr)s="PRIORIDADE",2000,'
                         'IF(%(pr)s="OPORTUNIDADE",1000,0))+N(%(av)s)*100+ROW()/1000000)'
                         % {"s": x("STATUS CALCULADO"), "l": x("LIBERAR P/ TESTE?"), "pr": x("PROGRAMAÇÃO"),
                            "av": x("AVANÇO CTB")}),
        "PONTOS REMONT.": ('=IF(%(s)s<>"P/REMONTAGEM",0,IF(N(%(r)s)=0,5000,0)+IF(%(pr)s="PRIORIDADE",2000,'
                           'IF(%(pr)s="OPORTUNIDADE",1000,0))+ROW()/1000000)'
                           % {"s": x("STATUS CALCULADO"), "r": x("RESTRIÇÕES ABERTAS"), "pr": x("PROGRAMAÇÃO")}),
        "PONTOS DESVIO": '=IF(ISNUMBER(%s),ABS(%s)+ROW()/10000000,"")' % (x("DESVIO MILPLAN − CTB"),
                                                                          x("DESVIO MILPLAN − CTB")),
        "ORDEM PROG.": '=IF(AND(%s<>"",%s<>""),ROW(),"")' % (pasta, x("PROGRAMAÇÃO")),
    }
    return f


def aba_controle(wb, linhas, P, nb):
    ws = wb.create_sheet("CONTROLE", 0)
    titulo(ws, "CONTROLE DE PASTAS DE TUBULAÇÃO - SALA DE GUERRA",
           "Colunas com cabeçalho DOURADO e fundo amarelo = preenchimento manual. Cabeçalho AZUL = automático "
           "(Painel OLODUM / ControlTub / HMSWEB). Para incluir uma pasta, digite o número na coluna B de uma linha vazia.")
    ws["A3"] = ("Exemplo de preenchimento: PROGRAMAÇÃO = PRIORIDADE | SEMANA PROG. = S46 | STATUS SALA DE GUERRA = P/TESTE "
                "| AVANÇO FÍSICO MILPLAN = 85% | DATA ENVIO = 13/11/2026")
    ws["A3"].font = F_SUB
    for i, (nome, tipo, larg) in enumerate(CTRL, 1):
        c = ws.cell(LIN_CAB, i, nome)
        c.font, c.alignment, c.border = F_CAB, CENTRO, BORDA
        c.fill = FILL_CAB_INP if tipo == "inp" else (FILL_CAB_AUTO if tipo == "auto" else PatternFill("solid", fgColor="808080"))
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[LIN_CAB].height = 48
    ws.cell(LIN_CAB, CTRL.index(("AVANÇO FÍSICO MILPLAN", "inp", 10)) + 1).comment = Comment(
        "Avanço físico informado pelo time da Milplan (0% a 100%). Comparado com o avanço de juntas do ControlTub.", "Sala de guerra")
    ws.cell(LIN_CAB, len(CTRL) - 3).comment = Comment(
        "Colunas cinza: auxiliares para ordenar o PAINEL. Não editar.", "Sala de guerra")

    ultima = LIN_CAB + max(MAX_LIN, len(linhas) + 50)
    for idx in range(LIN_CAB + 1, ultima + 1):
        dados = linhas[idx - LIN_CAB - 1] if idx - LIN_CAB - 1 < len(linhas) else {}
        f = f_controle(idx, P, nb)
        for i, (nome, tipo, _) in enumerate(CTRL, 1):
            c = ws.cell(idx, i)
            c.border = BORDA
            if tipo == "inp":
                c.value = dados.get(nome)
                c.font, c.fill = F_INP, FILL_INP
            else:
                c.value = f[nome]
                c.font = F_TXT if tipo == "auto" else Font(name=FONTE, size=8, color="808080")
            if nome in ("AVANÇO CTB", "AVANÇO FÍSICO MILPLAN"):
                c.number_format = "0%"
            elif nome == "DESVIO MILPLAN − CTB":
                c.number_format = '+0%;-0%;0%'
            elif nome in ("DATA ENVIO", "DATA PLANEJADA TH"):
                c.number_format = "dd/mm/yyyy"
            elif tipo == "aux":
                c.number_format = "0.00"
            if nome not in ("OBSERVAÇÃO", "Nº DA PASTA", "SUBSOP HMS", "SISTEMA", "RESPONSÁVEL"):
                c.alignment = Alignment(horizontal="center", vertical="center")
    rng = lambda nome: "%s%d:%s%d" % (C[nome], LIN_CAB + 1, C[nome], ultima)

    # validações
    for nome, lista in (("PROGRAMAÇÃO", P["prog"]), ("STATUS SALA DE GUERRA", P["status"])):
        dv = DataValidation(type="list", formula1="=" + lista, allow_blank=True)
        dv.error, dv.errorTitle = "Escolha um valor da lista (aba PARÂMETROS).", "Valor inválido"
        ws.add_data_validation(dv)
        dv.add(rng(nome))
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True)
    dv.error = "Informe um percentual entre 0% e 100%."
    ws.add_data_validation(dv)
    dv.add(rng("AVANÇO FÍSICO MILPLAN"))

    # formatação condicional
    cores = COR_STATUS
    for nome in ("STATUS CALCULADO",):
        for st, cor in cores.items():
            ws.conditional_formatting.add(rng(nome), CellIsRule(operator="equal", formula=['"%s"' % st],
                                          fill=PatternFill("solid", fgColor=cor), font=Font(name=FONTE, size=9, bold=True)))
    ws.conditional_formatting.add(rng("CONFERÊNCIA STATUS"), FormulaRule(
        formula=['LEFT(%s%d,6)="DIVERG"' % (C["CONFERÊNCIA STATUS"], LIN_CAB + 1)],
        fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    for st, cor in (("ALINHADO", "C6EFCE"), ("MILPLAN ACIMA", "F8CBAD"), ("CTB ACIMA", "FFEB9C")):
        ws.conditional_formatting.add(rng("ANÁLISE AVANÇO"), CellIsRule(operator="equal", formula=['"%s"' % st],
                                      fill=PatternFill("solid", fgColor=cor)))
    ws.conditional_formatting.add(rng("LIBERAR P/ TESTE?"), CellIsRule(operator="equal", formula=['"SIM"'],
                                  fill=PatternFill("solid", fgColor="00B050"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(rng("LIBERAR P/ TESTE?"), CellIsRule(operator="equal", formula=['"RESTRIÇÃO"'],
                                  fill=PatternFill("solid", fgColor="FF0000"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(rng("RESTRIÇÕES ABERTAS"), CellIsRule(operator="greaterThan", formula=["0"],
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
    ws.conditional_formatting.add(rng("AVANÇO FÍSICO MILPLAN"), DataBarRule(start_type="num", start_value=0,
                                  end_type="num", end_value=1, color="BF8F00"))

    ws.freeze_panes = ws.cell(LIN_CAB + 1, 3)
    ws.auto_filter.ref = "A%d:%s%d" % (LIN_CAB, get_column_letter(len(CTRL)), ultima)
    return ultima


def aba_restricoes(wb, restr, P, ult_ctrl):
    ws = wb.create_sheet("RESTRIÇÕES")
    titulo(ws, "CONTROLE DE RESTRIÇÕES DAS PASTAS",
           "Uma linha por restrição. Enquanto o STATUS não for ENCERRADA a pasta aparece com restrição no CONTROLE "
           "e sai da lista de 'melhores para teste/remontagem' do PAINEL. Colunas amarelas = preenchimento manual.")
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
        dv.error = "Informe uma data (dd/mm/aaaa)."
        ws.add_data_validation(dv)
        dv.add(rng(col))
    for st, cor, fc in (("ATRASADA", "FFC7CE", "9C0006"), ("NO PRAZO", "C6EFCE", "006100"), ("ENCERRADA", "EDEDED", "595959")):
        ws.conditional_formatting.add(rng("L"), CellIsRule(operator="equal", formula=['"%s"' % st],
                                      fill=PatternFill("solid", fgColor=cor), font=Font(name=FONTE, size=9, bold=True, color=fc)))
    ws.conditional_formatting.add(rng("C"), CellIsRule(operator="equal", formula=['"PASTA FORA DO CONTROLE"'],
                                  font=Font(name=FONTE, size=9, bold=True, color="C00000")))
    ws.freeze_panes = ws.cell(LIN_CAB + 1, 3)
    ws.auto_filter.ref = "A%d:L%d" % (LIN_CAB, ultima)
    return ultima


def aba_programacao(wb, ult_ctrl):
    ws = wb.create_sheet("PROGRAMAÇÃO")
    titulo(ws, "PASTAS DA PROGRAMAÇÃO SEMANAL - SITUAÇÃO",
           "Lista automática de todas as pastas com PROGRAMAÇÃO preenchida no CONTROLE (prioridade e oportunidade), "
           "na ordem do CONTROLE. Não editar: altere no CONTROLE.")
    cols = ["#", "Nº DA PASTA", "SUBSOP HMS", "PROGRAMAÇÃO", "SEMANA", "STATUS CALCULADO", "AVANÇO CTB",
            "SALDO JUNTAS", "CERT. TESTE", "REMONTAGEM", "RESTRIÇÕES", "LIBERAR P/ TESTE?", "AVANÇO MILPLAN",
            "ANÁLISE AVANÇO", "CONFERÊNCIA", "DATA PLAN. TH"]
    src = ["", "Nº DA PASTA", "SUBSOP HMS", "PROGRAMAÇÃO", "SEMANA PROG.", "STATUS CALCULADO", "AVANÇO CTB",
           "SALDO DE JUNTAS", "CERTIFICADO TESTE HMSWEB", "REMONTAGEM HMS", "RESTRIÇÕES ABERTAS", "LIBERAR P/ TESTE?",
           "AVANÇO FÍSICO MILPLAN", "ANÁLISE AVANÇO", "CONFERÊNCIA STATUS", "DATA PLANEJADA TH"]
    cab(ws, LIN_CAB, cols)
    ordem = "'CONTROLE'!$%s$%d:$%s$%d" % (C["ORDEM PROG."], LIN_CAB + 1, C["ORDEM PROG."], ult_ctrl)
    n = 250
    for k in range(1, n + 1):
        r = LIN_CAB + k
        ws.cell(r, 1, '=IF(COUNT(%s)>=%d,%d,"")' % (ordem, k, k))
        for j, s in enumerate(src[1:], 2):
            ws.cell(r, j, '=IF($A%d="","",INDEX(\'CONTROLE\'!$%s:$%s,SMALL(%s,$A%d))&"")'
                    % (r, C[s], C[s], ordem, r) if s not in ("AVANÇO CTB", "AVANÇO FÍSICO MILPLAN", "SALDO DE JUNTAS",
                                                              "RESTRIÇÕES ABERTAS", "DATA PLANEJADA TH")
                    else '=IF($A%d="","",IF(INDEX(\'CONTROLE\'!$%s:$%s,SMALL(%s,$A%d))="","",INDEX(\'CONTROLE\'!$%s:$%s,SMALL(%s,$A%d))))'
                    % (r, C[s], C[s], ordem, r, C[s], C[s], ordem, r))
        for j in range(1, len(cols) + 1):
            c = ws.cell(r, j)
            c.font, c.border = F_TXT, BORDA
            c.alignment = Alignment(horizontal="center", vertical="center")
            if cols[j - 1] in ("AVANÇO CTB", "AVANÇO MILPLAN"):
                c.number_format = "0%"
            if cols[j - 1] == "DATA PLAN. TH":
                c.number_format = "dd/mm/yyyy"
    rng = lambda col: "%s%d:%s%d" % (col, LIN_CAB + 1, col, LIN_CAB + n)
    cores = COR_STATUS
    for st, cor in cores.items():
        ws.conditional_formatting.add(rng("F"), CellIsRule(operator="equal", formula=['"%s"' % st],
                                      fill=PatternFill("solid", fgColor=cor), font=Font(name=FONTE, size=9, bold=True)))
    ws.conditional_formatting.add(rng("L"), CellIsRule(operator="equal", formula=['"RESTRIÇÃO"'],
                                  fill=PatternFill("solid", fgColor="FF0000"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(rng("L"), CellIsRule(operator="equal", formula=['"SIM"'],
                                  fill=PatternFill("solid", fgColor="00B050"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(rng("O"), FormulaRule(formula=['LEFT(O%d,6)="DIVERG"' % (LIN_CAB + 1)],
                                  fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.conditional_formatting.add(rng("G"), DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                        color="5B9BD5"))
    for j, w in enumerate([5, 17, 22, 13, 8, 15, 10, 9, 9, 12, 10, 11, 10, 15, 26, 11], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C5"


def aba_painel(wb, ult_ctrl, ult_restr, gerado_em):
    ws = wb.create_sheet("PAINEL", 0)
    ws.sheet_view.showGridLines = False
    ws["B1"] = "PAINEL GERENCIAL - SALA DE GUERRA DE PASTAS DE TUBULAÇÃO"
    ws["B1"].font = Font(name=FONTE, size=18, bold=True, color=AZUL)
    ws["B2"] = ('="Base Painel OLODUM: %s   |   Avanço mínimo p/ teste: "&TEXT(\'PARÂMETROS\'!$B$4,"0%%")&'
                '"   |   Atualizado em "&TEXT(\'PARÂMETROS\'!$B$6,"dd/mm/yyyy")' % gerado_em)
    ws["B2"].font = F_SUB
    ws.column_dimensions["A"].width = 2
    for col, w in zip("BCDEFGHIJKLMNOPQRS", [5, 17, 20, 12, 10, 10, 11, 11, 3, 5, 17, 20, 12, 10, 10, 11, 11, 3]):
        ws.column_dimensions[col].width = w

    R = lambda nome: "'CONTROLE'!$%s$%d:$%s$%d" % (C[nome], LIN_CAB + 1, C[nome], ult_ctrl)
    st, pr, pasta = R("STATUS CALCULADO"), R("PROGRAMAÇÃO"), R("Nº DA PASTA")

    # ---------------- KPIs
    kpis = [
        ("PASTAS CONTROLADAS", '=COUNTA(%s)' % pasta),
        ("NA PROGRAMAÇÃO", '=COUNTIF(%s,"PRIORIDADE")+COUNTIF(%s,"OPORTUNIDADE")' % (pr, pr)),
        ("P/ TESTE (CTB ≥ MÍN.)", '=COUNTIF(%s,"P/TESTE")' % st),
        ("LIBERAR P/ TESTE (SEM RESTRIÇÃO)", '=COUNTIF(%s,"SIM")' % R("LIBERAR P/ TESTE?")),
        ("P/ REMONTAGEM", '=COUNTIF(%s,"P/REMONTAGEM")' % st),
        ("PRÉ WALKDOWN (INTRA SKID)", '=COUNTIF(%s,"PRÉ WALKDOWN")' % st),
        ("RESTRIÇÕES ABERTAS", '=COUNTIFS(\'RESTRIÇÕES\'!$B$%d:$B$%d,"<>",\'RESTRIÇÕES\'!$I$%d:$I$%d,"<>ENCERRADA")'
         % (LIN_CAB + 1, ult_restr, LIN_CAB + 1, ult_restr)),
        ("STATUS DIVERGENTE", '=COUNTIF(%s,"DIVERG*")' % R("CONFERÊNCIA STATUS")),
    ]
    pos = ["B", "D", "F", "H", "K", "M", "O", "Q"]
    for (rot, f), col in zip(kpis, pos):
        c2 = get_column_letter(openpyxl.utils.column_index_from_string(col) + 1)
        ws.merge_cells("%s4:%s4" % (col, c2))
        ws.merge_cells("%s5:%s5" % (col, c2))
        ws["%s4" % col] = rot
        ws["%s5" % col] = f
        ws["%s4" % col].font, ws["%s5" % col].font = F_KPI_L, F_KPI
        for rr in (4, 5):
            for cc in (col, c2):
                ws["%s%d" % (cc, rr)].fill = FILL_KPI
            ws["%s%d" % (col, rr)].alignment = CENTRO
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 34

    # ---------------- Status x programação
    ws["B7"] = "SITUAÇÃO DAS PASTAS × PROGRAMAÇÃO SEMANAL"
    ws["B7"].font = F_H2
    hdr = ["STATUS CALCULADO", "", "PRIORIDADE", "OPORTUNIDADE", "FORA DA PROG.", "TOTAL", "% DO TOTAL"]
    for j, h in enumerate(hdr):
        c = ws.cell(8, 3 + j, h)
        c.font, c.fill, c.alignment, c.border = F_CAB, FILL_CAB, CENTRO, BORDA
    ws.merge_cells("C8:D8")
    for i, s in enumerate(STATUS_LISTA):
        r = 9 + i
        ws.merge_cells("C%d:D%d" % (r, r))
        ws.cell(r, 3, s)
        ws.cell(r, 5, '=COUNTIFS(%s,$C%d,%s,"PRIORIDADE")' % (st, r, pr))
        ws.cell(r, 6, '=COUNTIFS(%s,$C%d,%s,"OPORTUNIDADE")' % (st, r, pr))
        ws.cell(r, 8, '=COUNTIF(%s,$C%d)' % (st, r))
        ws.cell(r, 7, "=H%d-E%d-F%d" % (r, r, r))
        ws.cell(r, 9, '=IF($H$%d=0,0,H%d/$H$%d)' % (9 + len(STATUS_LISTA), r, 9 + len(STATUS_LISTA)))
    rt = 9 + len(STATUS_LISTA)
    ws.merge_cells("C%d:D%d" % (rt, rt))
    ws.cell(rt, 3, "TOTAL")
    for j in range(5, 10):
        L = get_column_letter(j)
        ws.cell(rt, j, "=SUM(%s9:%s%d)" % (L, L, rt - 1))
    for r in range(9, rt + 1):
        for j in range(3, 10):
            c = ws.cell(r, j)
            c.border, c.alignment = BORDA, CENTRO
            c.font = Font(name=FONTE, size=9, bold=(r == rt or j in (3, 8)))
            if j == 9:
                c.number_format = "0%"
        if r == rt:
            for j in range(3, 10):
                ws.cell(r, j).fill = FILL_KPI
    cores = COR_STATUS
    for i, s in enumerate(STATUS_LISTA):
        ws.cell(9 + i, 3).fill = PatternFill("solid", fgColor=cores[s])

    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "bar", "stacked", 100
    ch.title = "Pastas por status e programação"
    ch.y_axis.title = "Qtd. de pastas"
    data = Reference(ws, min_col=5, max_col=7, min_row=8, max_row=rt - 1)
    cats = Reference(ws, min_col=3, min_row=9, max_row=rt - 1)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s, cor in zip(ch.series, ["1F3864", "5B9BD5", "BFBFBF"]):
        s.graphicalProperties.solidFill = cor
        s.graphicalProperties.line.solidFill = cor
    ch.height, ch.width = 8.4, 15.5
    ch.roundedCorners = False
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.y_axis.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(ch, "K7")

    # ---------------- Avanço Milplan x ControlTub
    r0 = rt + 7  # abaixo do gráfico
    ws.cell(r0, 2, "AVANÇO FÍSICO MILPLAN × CONTROLTUB").font = F_H2
    ws.cell(r0, 6, '="Tolerância: ±"&TEXT(\'PARÂMETROS\'!$B$5,"0%")').font = F_SUB
    for j, h in enumerate(["ANÁLISE", "", "TODAS", "NA PROGRAMAÇÃO"]):
        c = ws.cell(r0 + 1, 3 + j, h)
        c.font, c.fill, c.alignment, c.border = F_CAB, FILL_CAB, CENTRO, BORDA
    ws.merge_cells("C%d:D%d" % (r0 + 1, r0 + 1))
    an = R("ANÁLISE AVANÇO")
    analises = [("ALINHADO", "C6EFCE"), ("MILPLAN ACIMA", "F8CBAD"), ("CTB ACIMA", "FFEB9C"),
                ("SEM INFO MILPLAN", "EDEDED"), ("SEM CTB", "EDEDED")]
    for i, (a, cor) in enumerate(analises):
        r = r0 + 2 + i
        ws.merge_cells("C%d:D%d" % (r, r))
        ws.cell(r, 3, a).fill = PatternFill("solid", fgColor=cor)
        ws.cell(r, 5, '=COUNTIF(%s,$C%d)' % (an, r))
        ws.cell(r, 6, '=COUNTIFS(%s,$C%d,%s,"PRIORIDADE")+COUNTIFS(%s,$C%d,%s,"OPORTUNIDADE")' % (an, r, pr, an, r, pr))
        for j in range(3, 7):
            c = ws.cell(r, j)
            c.border, c.alignment, c.font = BORDA, CENTRO, Font(name=FONTE, size=9, bold=(j == 3))
    ra = r0 + 2 + len(analises)
    ws.cell(ra, 3, "Avanço médio CTB (pastas com Milplan informado)").font = F_SUB
    ws.cell(ra, 8, '=IFERROR(AVERAGEIFS(%s,%s,">=0"),"-")' % (R("AVANÇO CTB"), R("AVANÇO FÍSICO MILPLAN")))
    ws.cell(ra + 1, 3, "Avanço médio Milplan informado").font = F_SUB
    ws.cell(ra + 1, 8, '=IFERROR(AVERAGE(%s),"-")' % R("AVANÇO FÍSICO MILPLAN"))
    for rr in (ra, ra + 1):
        ws.cell(rr, 8).number_format = "0%"
        ws.cell(rr, 8).font = Font(name=FONTE, size=10, bold=True, color=AZUL)
        ws.cell(rr, 8).alignment = CENTRO

    # maiores desvios
    ws.cell(r0, 11, "MAIORES DESVIOS MILPLAN × CTB").font = F_H2
    hdr = ["#", "Nº DA PASTA", "SUBSOP HMS", "PROGRAMAÇÃO", "CTB", "MILPLAN", "DESVIO"]
    for j, h in enumerate(hdr):
        c = ws.cell(r0 + 1, 11 + j, h)
        c.font, c.fill, c.alignment, c.border = F_CAB, FILL_CAB, CENTRO, BORDA
    pdv = R("PONTOS DESVIO")
    for k in range(1, 8):
        r = r0 + 1 + k
        ws.cell(r, 11, '=IF(COUNT(%s)>=%d,%d,"")' % (pdv, k, k))
        ref = "MATCH(LARGE(%s,$K%d),%s,0)" % (pdv, r, pdv)
        for j, s in enumerate(["Nº DA PASTA", "SUBSOP HMS", "PROGRAMAÇÃO", "AVANÇO CTB", "AVANÇO FÍSICO MILPLAN",
                               "DESVIO MILPLAN − CTB"], 12):
            ws.cell(r, j, '=IF($K%d="","",INDEX(%s,%s))' % (r, R(s), ref))
        for j in range(11, 18):
            c = ws.cell(r, j)
            c.border, c.alignment, c.font = BORDA, CENTRO, F_TXT
            if j in (15, 16):
                c.number_format = "0%"
            if j == 17:
                c.number_format = "+0%;-0%;0%"
    ws.cell(r0 + 9, 11, "Preencha AVANÇO FÍSICO MILPLAN no CONTROLE para alimentar este quadro.").font = F_SUB

    # ---------------- Rankings
    def ranking(lin, col, titulo_txt, nota, pontos, campos, n=25):
        ws.cell(lin, col, titulo_txt).font = F_H2
        ws.cell(lin + 1, col, nota).font = F_SUB
        nomes = ["#"] + [c[0] for c in campos]
        for j, h in enumerate(nomes):
            c = ws.cell(lin + 2, col + j, h)
            c.font, c.fill, c.alignment, c.border = F_CAB, FILL_CAB, CENTRO, BORDA
        ws.row_dimensions[lin + 2].height = 30
        pt = R(pontos)
        Lk = get_column_letter(col)
        for k in range(1, n + 1):
            r = lin + 2 + k
            ws.cell(r, col, '=IF(COUNTIF(%s,">0")>=%d,%d,"")' % (pt, k, k))
            ref = "MATCH(LARGE(%s,$%s%d),%s,0)" % (pt, Lk, r, pt)
            for j, (_, s, fmt) in enumerate(campos, 1):
                c = ws.cell(r, col + j, '=IF($%s%d="","",INDEX(%s,%s))' % (Lk, r, R(s), ref))
                if fmt:
                    c.number_format = fmt
            for j in range(len(nomes)):
                c = ws.cell(r, col + j)
                c.border, c.alignment, c.font = BORDA, CENTRO, F_TXT
                if k % 2 == 0:
                    c.fill = FILL_ZEBRA
        return lin + 2 + n

    lin = ra + 4
    campos_t = [("Nº DA PASTA", "Nº DA PASTA", None), ("SUBSOP HMS", "SUBSOP HMS", None),
                ("PROGRAMAÇÃO", "PROGRAMAÇÃO", None), ("AVANÇO CTB", "AVANÇO CTB", "0%"),
                ("SALDO JUNTAS", "SALDO DE JUNTAS", "0"), ("LIBERAR?", "LIBERAR P/ TESTE?", None)]
    fim_t = ranking(lin, 2, "MELHORES PASTAS PARA TESTE",
                    "Ordem: sem restrição > prioridade > oportunidade > maior avanço CTB.", "PONTOS TESTE", campos_t)
    campos_r = [("Nº DA PASTA", "Nº DA PASTA", None), ("SUBSOP HMS", "SUBSOP HMS", None),
                ("PROGRAMAÇÃO", "PROGRAMAÇÃO", None), ("RESTRIÇÕES", "RESTRIÇÕES ABERTAS", "0"),
                ("LAVAGEM HMS", "LAVAGEM HMS", None), ("DATA PLAN. TH", "DATA PLANEJADA TH", "dd/mm/yy")]
    ranking(lin, 11, "MELHORES PASTAS PARA REMONTAGEM",
            "Teste aprovado no HMSWEB e remontagem pendente. Ordem: sem restrição > prioridade > oportunidade.",
            "PONTOS REMONT.", campos_r)
    a1, a2 = lin + 3, lin + 27
    ws.conditional_formatting.add("H%d:H%d" % (a1, a2), CellIsRule(operator="equal", formula=['"SIM"'],
                                  fill=PatternFill("solid", fgColor="00B050"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add("H%d:H%d" % (a1, a2), CellIsRule(operator="equal", formula=['"RESTRIÇÃO"'],
                                  fill=PatternFill("solid", fgColor="FF0000"), font=Font(name=FONTE, size=9, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add("O%d:O%d" % (a1, a2), CellIsRule(operator="greaterThan", formula=["0"],
                                  fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONTE, size=9, bold=True, color="9C0006")))
    ws.conditional_formatting.add("F%d:F%d" % (a1, a2), DataBarRule(start_type="num", start_value=0, end_type="num",
                                  end_value=1, color="5B9BD5"))

    # ---------------- Restrições por tipo
    lin = fim_t + 2
    ws.cell(lin, 2, "RESTRIÇÕES POR TIPO").font = F_H2
    for j, h in enumerate(["TIPO", "", "ABERTAS", "ATRASADAS", "ENCERRADAS"]):
        c = ws.cell(lin + 1, 3 + j, h)
        c.font, c.fill, c.alignment, c.border = F_CAB, FILL_CAB, CENTRO, BORDA
    ws.merge_cells("C%d:D%d" % (lin + 1, lin + 1))
    rr = lambda col: "'RESTRIÇÕES'!$%s$%d:$%s$%d" % (col, LIN_CAB + 1, col, ult_restr)
    for i, t in enumerate(TIPOS_RESTR):
        r = lin + 2 + i
        ws.merge_cells("C%d:D%d" % (r, r))
        ws.cell(r, 3, t)
        ws.cell(r, 5, '=COUNTIFS(%s,$C%d,%s,"<>ENCERRADA")' % (rr("D"), r, rr("I")))
        ws.cell(r, 6, '=COUNTIFS(%s,$C%d,%s,"ATRASADA")' % (rr("D"), r, rr("L")))
        ws.cell(r, 7, '=COUNTIFS(%s,$C%d,%s,"ENCERRADA")' % (rr("D"), r, rr("I")))
        for j in range(3, 8):
            c = ws.cell(r, j)
            c.border, c.font = BORDA, F_TXT
            c.alignment = ESQ if j == 3 else CENTRO
    ws.conditional_formatting.add("F%d:F%d" % (lin + 2, lin + 1 + len(TIPOS_RESTR)), CellIsRule(
        operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(name=FONTE, size=9, bold=True, color="9C0006")))

    # pastas com restrição aberta (lista)
    ws.cell(lin, 11, "PASTAS DA PROGRAMAÇÃO COM PENDÊNCIA DE STATUS").font = F_H2
    ws.cell(lin + 1, 11, "Status informado na sala de guerra diferente do calculado pela base.").font = F_SUB
    ws.cell(lin + 2, 11, "Total divergentes:").font = Font(name=FONTE, size=9, bold=True)
    ws.cell(lin + 2, 14, '=COUNTIF(%s,"DIVERG*")' % R("CONFERÊNCIA STATUS")).font = Font(name=FONTE, size=11, bold=True, color="C00000")
    ws.cell(lin + 3, 11, "Filtre a coluna CONFERÊNCIA STATUS do CONTROLE ou veja a aba PROGRAMAÇÃO.").font = F_SUB
    for j, (s, rot) in enumerate([("P/TESTE", "Divergentes informados P/TESTE"),
                                  ("P/REMONTAGEM", "Divergentes informados P/REMONTAGEM"),
                                  ("EM MONTAGEM", "Divergentes informados EM MONTAGEM")]):
        r = lin + 5 + j
        ws.cell(r, 11, rot).font = F_TXT
        ws.cell(r, 14, '=COUNTIFS(%s,"DIVERG*",%s,"%s")' % (R("CONFERÊNCIA STATUS"), R("STATUS SALA DE GUERRA"), s))
        ws.cell(r, 14).font = Font(name=FONTE, size=10, bold=True)
        ws.cell(r, 14).alignment = CENTRO

    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = AZUL
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0


# ----------------------------------------------------------------------------- principal
def gerar(painel, controle, saida):
    d = ler_painel(painel)
    base = base_por_pasta(d)
    gerado_em = ""
    for r in d["__DADOS__"]:
        if r.get("atualizado_em"):
            gerado_em = r["atualizado_em"][:16].replace("T", " ") + " UTC"
            break
    linhas, restr = ler_controle(controle)
    vistos = {l["Nº DA PASTA"] for l in linhas}
    sub = {b["PASTA"]: b["SUBSISTEMA"] for b in base}
    for l in linhas:
        if not l.get("SUBSOP HMS"):
            l["SUBSOP HMS"] = sub.get(l["Nº DA PASTA"], "")
    novas = [b for b in base if b["PASTA"] not in vistos]
    for b in novas:
        linhas.append({"Nº DA PASTA": b["PASTA"], "SUBSOP HMS": b["SUBSISTEMA"]})

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    P = aba_parametros(wb)
    nb = aba_base(wb, base, gerado_em)
    ult_ctrl = aba_controle(wb, linhas, P, nb)
    if not restr:
        # Restrições já visíveis na base: protocolos reprovados no HMSWEB.
        hoje = dt.date.today()
        for r in d.get("__PCPPEND__", []):
            if r["pasta"].startswith(PREFIXO_PASTA) and r["situacao"] == "Reprovada":
                restr.append([None, r["pasta"], None, "Junta reprovada / END" if "Juntas" in r["atividade"]
                              else "Documentação / Isométrico",
                              "HMSWEB: %s reprovada" % r["atividade"], "", hoje, None, "ABERTA", None, None, None])
    ult_restr = aba_restricoes(wb, restr, P, ult_ctrl)
    aba_programacao(wb, ult_ctrl)
    aba_painel(wb, ult_ctrl, ult_restr, gerado_em)
    ordem = ["PAINEL", "CONTROLE", "PROGRAMAÇÃO", "RESTRIÇÕES", "BASE_OLODUM", "PARÂMETROS"]
    wb._sheets = [wb[n] for n in ordem]
    wb["CONTROLE"].sheet_properties.tabColor = "2E75B6"
    wb["RESTRIÇÕES"].sheet_properties.tabColor = "C00000"
    wb["PROGRAMAÇÃO"].sheet_properties.tabColor = "00B050"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(saida)
    print("Pastas no controle: %d (%d novas vindas do Painel OLODUM)" % (len(linhas), len(novas)))
    print("Pastas na base OLODUM: %d | restrições: %d" % (len(base), len(restr)))
    print("Gravado:", saida)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("painel")
    ap.add_argument("controle")
    ap.add_argument("saida")
    a = ap.parse_args()
    gerar(a.painel, a.controle, a.saida)
