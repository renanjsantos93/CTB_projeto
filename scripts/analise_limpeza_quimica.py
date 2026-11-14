r"""Análise: as pastas marcadas para limpeza química na 06-Lista de Pastas_Milplan têm atividade no HMSWEB?

Cruza:
  - 06-Lista de Pastas_Milplan, aba MILPLAN: pastas com "X" na coluna I (Limpeza Química), pela coluna N
    (Nº DA PASTA - PARA), com o escopo (coluna G), a linha (coluna R, Line Code) e o subsistema (coluna U);
  - aba PCP_ATIVIDADE da OLODUM_BASEDADOS (as atividades do HMSWEB, com o Nº DA PASTA).

Para cada pasta com X, diz se o HMSWEB tem a atividade "Limpeza Química" na pasta (e quantas foram executadas e
aprovadas), se a pasta existe no HMSWEB só com outras atividades, ou se não existe. Também procura a atividade
pelas linhas da pasta (COD IDENTIFICAÇÃO), para achar uma limpeza química lançada sem o Nº da pasta, e lista as
pastas que têm a atividade no HMSWEB mas não têm X na lista.

Uso:
  python analise_limpeza_quimica.py                          usa as planilhas da pasta 07.OLODUM (J:)
  python analise_limpeza_quimica.py <olodum.xlsx> <06-lista.xlsx> <saida.xlsx>
Saída: Analise_Limpeza_Quimica.xlsx na pasta do painel (ou o caminho dado).
"""
import datetime
import re
import sys
import tempfile
import unicodedata
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))

ABA_LISTA = "MILPLAN"
LINHA_INICIAL = 5
COL_ESCOPO, COL_LQ, COL_PASTA, COL_LINHA, COL_SUBSISTEMA = 6, 8, 13, 17, 20  # G, I, N, R, U (0 = coluna A)
IGNORAR = {"", "CANCELADO", "N/A", "#N/A", "#VALUE!", "0"}

OK = "Com atividade de Limpeza Química"
SEM_ATV = "Pasta no HMSWEB sem atividade de Limpeza Química"
SEM_PASTA = "Pasta não encontrada no HMSWEB"


def chave(t):
    """Texto sem acento, espaço e pontuação, em maiúsculas (para comparar nomes e códigos)."""
    t = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]", "", t.upper())


def pasta_base(p):
    """'421-M-TU-0005.1' -> '421-M-TU-0005' (a subdivisão conta como a pasta)."""
    return re.sub(r"\.\d+$", "", str(p or "").strip())


def eh_limpeza_quimica(atividade):
    return "LIMPEZAQUIMICA" in chave(atividade)


def preenchido(v):
    return v not in (None, "", 0, "-")


def ler_lista(planilha):
    pastas = {}
    for r in openpyxl.load_workbook(planilha, read_only=True, data_only=True)[ABA_LISTA].iter_rows(
            min_row=LINHA_INICIAL, values_only=True):
        if not r or len(r) <= COL_SUBSISTEMA:
            continue
        pasta = str(r[COL_PASTA] or "").strip()
        if pasta.upper() in IGNORAR:
            continue
        p = pastas.setdefault(pasta, dict(lq=False, escopos=set(), linhas=set(), linhas_lq=set(), subsistemas=set()))
        escopo = str(r[COL_ESCOPO] or "").strip()
        linha = str(r[COL_LINHA] or "").strip()
        if escopo.upper() not in IGNORAR:
            p["escopos"].add(escopo)
        if linha.upper() not in IGNORAR:
            p["linhas"].add(linha)
        if str(r[COL_LQ] or "").strip().upper() == "X":
            p["lq"] = True
            if linha.upper() not in IGNORAR:
                p["linhas_lq"].add(linha)
            sub = str(r[COL_SUBSISTEMA] or "").strip()
            if sub.upper() not in IGNORAR:
                p["subsistemas"].add(sub)
    return pastas


def ler_hmsweb(planilha):
    linhas = openpyxl.load_workbook(planilha, read_only=True, data_only=True)["PCP_ATIVIDADE"].iter_rows(values_only=True)
    cab = [str(c or "").strip() for c in next(linhas)]
    col = {chave(c): i for i, c in enumerate(cab)}

    def v(r, nome):
        i = col.get(chave(nome))
        return r[i] if i is not None and i < len(r) else None

    atividades = []
    for r in linhas:
        if not any(r):
            continue
        pasta = str(v(r, "Nº DA PASTA") or "").strip()
        atividades.append(dict(
            atividade=re.sub(r"\s*\([^)]*\)\s*$", "", str(v(r, "ATIVIDADE") or "")).strip(),
            fase=str(v(r, "FASE") or "").strip(), disciplina=str(v(r, "DISCIPLINA") or "").strip(),
            tag=str(v(r, "COD IDENTIFICAÇÃO") or "").strip(), objeto=str(v(r, "DESCRIÇÃO DO OBJETO") or "").strip(),
            subsistema=str(v(r, "SUBSISTEMA") or "").strip(),
            pasta="" if pasta.upper() in IGNORAR else pasta,
            executada=preenchido(v(r, "DATA DA EXECUÇÃO")), aprovada=preenchido(v(r, "DATA DA APROVAÇÃO")),
            reprovada=preenchido(v(r, "DATA DA REPROVAÇÃO")),
            data_planejada=v(r, "DATA PLANEJADA"),
        ))
    return atividades


def analisar(planilha_olodum, planilha_lista):
    lista = ler_lista(planilha_lista)
    atividades = ler_hmsweb(planilha_olodum)
    por_pasta = defaultdict(list)
    lq_por_tag = defaultdict(list)
    for a in atividades:
        if a["pasta"]:
            por_pasta[pasta_base(a["pasta"])].append(a)
        if eh_limpeza_quimica(a["atividade"]):
            lq_por_tag[chave(a["tag"])].append(a)

    resultado = []
    for pasta, p in sorted(lista.items()):
        if not p["lq"]:
            continue
        da_pasta = por_pasta.get(pasta_base(pasta), [])
        lq = [a for a in da_pasta if eh_limpeza_quimica(a["atividade"])]
        situacao = OK if lq else SEM_ATV if da_pasta else SEM_PASTA
        # Limpeza química lançada numa linha da pasta, mas com outro Nº da pasta (ou sem).
        pelas_linhas = [a for linha in p["linhas"] for a in lq_por_tag.get(chave(linha), [])
                        if pasta_base(a["pasta"]) != pasta_base(pasta)]
        obs = ""
        if pelas_linhas:
            outras = sorted({a["pasta"] or "sem Nº da pasta" for a in pelas_linhas})
            obs = "Limpeza Química lançada nas linhas da pasta em: " + ", ".join(outras)
        resultado.append(dict(
            pasta=pasta, escopo=" / ".join(sorted(p["escopos"])),
            subsistemas_lista=", ".join(sorted(p["subsistemas"])),
            linhas_lq=len(p["linhas_lq"]), situacao=situacao,
            lq_previstas=len(lq), lq_executadas=sum(a["executada"] for a in lq),
            lq_aprovadas=sum(a["aprovada"] for a in lq), lq_reprovadas=sum(a["reprovada"] for a in lq),
            outras_atividades=len(da_pasta) - len(lq),
            subsistemas_hmsweb=", ".join(sorted({a["subsistema"] for a in da_pasta if a["subsistema"]})),
            obs=obs, linhas=", ".join(sorted(p["linhas_lq"])),
        ))

    # Pastas com a atividade no HMSWEB e sem X na lista.
    com_x = {pasta_base(p) for p, x in lista.items() if x["lq"]}
    sem_x = defaultdict(list)
    for a in atividades:
        if eh_limpeza_quimica(a["atividade"]) and pasta_base(a["pasta"]) not in com_x:
            sem_x[a["pasta"] or "(sem Nº da pasta)"].append(a)
    detalhe = [a for a in atividades if eh_limpeza_quimica(a["atividade"])]
    return resultado, sem_x, detalhe, lista


# ---- Excel ----------------------------------------------------------------------------------------
CAB = Font(bold=True, color="FFFFFF")
CAB_FUNDO = PatternFill("solid", fgColor="1F3A5F")
COR = {OK: "C6EFCE", SEM_ATV: "FFEB9C", SEM_PASTA: "FFC7CE"}


def aba(wb, titulo, cab, linhas, larguras):
    ws = wb.create_sheet(titulo)
    ws.append(cab)
    for c in ws[1]:
        c.font, c.fill = CAB, CAB_FUNDO
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for ln in linhas:
        ws.append(ln)
    for i, w in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"
    if linhas:
        ws.auto_filter.ref = ws.dimensions
    return ws


def gravar(resultado, sem_x, detalhe, lista, planilha_olodum, planilha_lista, saida):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo"
    contagem = {s: sum(r["situacao"] == s for r in resultado) for s in (OK, SEM_ATV, SEM_PASTA)}
    aprovadas = sum(1 for r in resultado if r["lq_previstas"] and r["lq_aprovadas"] >= r["lq_previstas"])
    ws.append(["Análise de limpeza química: 06-Lista de Pastas_Milplan x HMSWEB"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"Gerada em {datetime.datetime.now():%d/%m/%Y %H:%M}"])
    ws.append([f"HMSWEB (PCP_ATIVIDADE): {Path(planilha_olodum).name}, salva em "
               f"{datetime.datetime.fromtimestamp(Path(planilha_olodum).stat().st_mtime):%d/%m/%Y %H:%M}"])
    ws.append([f"Lista de pastas: {Path(planilha_lista).name}"])
    ws.append([])
    ws.append(["Pastas com X em Limpeza Química (coluna I)", len(resultado)])
    for s in (OK, SEM_ATV, SEM_PASTA):
        ws.append(["   " + s, contagem[s]])
        ws.cell(ws.max_row, 1).fill = PatternFill("solid", fgColor=COR[s])
    ws.append(["   ...com todas as limpezas químicas aprovadas", aprovadas])
    ws.append(["Pastas com Limpeza Química no HMSWEB e sem X na lista", len(sem_x)])
    ws.append(["Atividades de Limpeza Química no HMSWEB (todas)", len(detalhe)])
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 10

    w = aba(wb, "Pastas com X", [
        "Nº da pasta", "Escopo", "Situação", "Limpeza Química prevista (HMSWEB)", "Executadas", "Aprovadas",
        "Reprovadas", "Outras atividades da pasta (HMSWEB)", "Subsistema (lista de pastas)",
        "Subsistema (HMSWEB)", "Linhas com X", "Observação", "Linhas"],
        [[r["pasta"], r["escopo"], r["situacao"], r["lq_previstas"], r["lq_executadas"], r["lq_aprovadas"],
          r["lq_reprovadas"], r["outras_atividades"], r["subsistemas_lista"], r["subsistemas_hmsweb"],
          r["linhas_lq"], r["obs"], r["linhas"]] for r in resultado],
        [16, 18, 44, 14, 11, 11, 11, 16, 30, 30, 10, 40, 60])
    for i, r in enumerate(resultado, 2):
        w.cell(i, 3).fill = PatternFill("solid", fgColor=COR[r["situacao"]])

    aba(wb, "Limpeza Química sem X", [
        "Nº da pasta", "Na lista de pastas?", "Atividades de Limpeza Química", "Executadas", "Aprovadas",
        "Subsistema (HMSWEB)"],
        [[p, "sim, sem X" if pasta_base(p) in {pasta_base(x) for x in lista} else "não", len(a),
          sum(x["executada"] for x in a), sum(x["aprovada"] for x in a),
          ", ".join(sorted({x["subsistema"] for x in a if x["subsistema"]}))] for p, a in sorted(sem_x.items())],
        [16, 18, 16, 11, 11, 40])

    def situacao(a):
        return "Aprovada" if a["aprovada"] else "Reprovada" if a["reprovada"] else "Executada" if a["executada"] else "Não executada"

    def data(v):
        return v.date() if hasattr(v, "date") else v

    wd = aba(wb, "Atividades LQ no HMSWEB", [
        "Nº da pasta", "Tag (COD IDENTIFICAÇÃO)", "Descrição do objeto", "Subsistema", "Atividade", "Fase",
        "Situação", "Data planejada"],
        [[a["pasta"], a["tag"], a["objeto"], a["subsistema"], a["atividade"], a["fase"], situacao(a),
          data(a["data_planejada"])] for a in sorted(detalhe, key=lambda a: (a["pasta"], a["tag"]))],
        [16, 28, 30, 22, 24, 30, 14, 14])
    for c in wd["H"][1:]:
        c.number_format = "DD/MM/YYYY"
    wb.save(saida)


def main(planilha_olodum=None, planilha_lista=None, saida=None):
    if not planilha_olodum:
        import atualizar_olodum as ao
        planilha_olodum = ao.PLANILHA_OLODUM
        planilha_lista = ao.planilha_de_pastas()
        if not planilha_lista:
            raise SystemExit("Não achei a 06-Lista de Pastas_Milplan em " + str(ao.PASTA_BASES))
        saida = ao.PASTA_PAINEL / "Analise_Limpeza_Quimica.xlsx"
        with tempfile.TemporaryDirectory() as tmp:
            planilha_lista_xlsx = ao.em_xlsx(planilha_lista, tmp)
            resultado = analisar(planilha_olodum, planilha_lista_xlsx)
    else:
        resultado = analisar(planilha_olodum, planilha_lista)
    saida = Path(saida)
    try:
        gravar(*resultado, planilha_olodum, planilha_lista, saida)
    except PermissionError:  # a análise anterior está aberta no Excel: grava com outro nome
        saida = saida.with_name(f"{saida.stem}_{datetime.datetime.now():%Y%m%d_%H%M}{saida.suffix}")
        gravar(*resultado, planilha_olodum, planilha_lista, saida)
    res = resultado[0]
    for s in (OK, SEM_ATV, SEM_PASTA):
        print(f"{sum(r['situacao'] == s for r in res):4d}  {s}")
    print(f"{len(resultado[1]):4d}  pastas com Limpeza Química no HMSWEB e sem X na lista")
    print("Análise gravada em", saida)
    return saida


if __name__ == "__main__":
    main(*sys.argv[1:4])
