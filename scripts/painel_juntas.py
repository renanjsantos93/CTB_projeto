"""Gera a base de juntas do ControlTub para o painel, a partir da 01.Situacao_Geral_de_Juntas.

Uso: python painel_juntas.py <01.Situacao_Geral_de_Juntas.xlsx> <OLODUM_BASEDADOS.xlsx> <saida.json>

Cada linha da saída é um par sistema x subsistema x pasta, com:
  previstas   = juntas da planilha
  executadas  = juntas com Situação "Junta Liberada"
Subsistema: a coluna SUBSOP da planilha de juntas; sem ela, o da linha na aba BASE_DADOS da OLODUM.
Pasta: a malha da linha na aba Item de Malha da OLODUM (é o Nº DA PASTA, 421-M-TU-...).
A linha é comparada pela TAG padronizada DN-CLASSE-SERVIÇO-NÚMERO (sem ", sem a zona -Z01 e sem sufixos).
"""
import json
import re
import sys
from collections import Counter, defaultdict

import openpyxl

SEM_SISTEMA = "Sistema não cadastrado"
SEM_SUBSISTEMA = "Subsistema não cadastrado"
SEM_PASTA = "Pasta não cadastrada"
COL_LINHA, COL_SOP, COL_SUBSOP, COL_SITUACAO = 2, 4, 6, 93  # colunas da aba SGJ (0 = coluna A)
LINHA_INICIAL = 11  # dados começam depois das duas linhas de cabeçalho


def chave_linha(tag):
    """'1"-A1E-CF-4310A-Z02' e '1-A1E-CF-4310A-C-3' -> '1-A1E-CF-4310A'."""
    t = str(tag or "").replace('"', "").replace("''", "").strip().upper()
    t = re.sub(r"-Z\d+(/\d+)?$", "", t)
    return "-".join(t.split("-")[:4])


def sistema_do_subsistema(sub):
    """'421-PRO-MR-252-001' -> '421-PRO-MR-252'."""
    return sub.rsplit("-", 1)[0] if "-" in sub else sub


def ler_olodum(olodum):
    """Linha -> (sistema, subsistema) pela BASE_DADOS e linha -> pasta pela Item de Malha."""
    wb = openpyxl.load_workbook(olodum, read_only=True, data_only=True)
    linhas = wb["BASE_DADOS"].iter_rows(values_only=True)
    cab = next(linhas)
    sub_da_linha = {}
    for valores in linhas:
        r = dict(zip(cab, valores))
        if r.get("Cod Familia") == "LL" and r.get("Tag do Item") and r.get("Cod SubSistema"):
            sub_da_linha.setdefault(chave_linha(r["Tag do Item"]), (r["Cod Sistema"], r["Cod SubSistema"]))
    malhas = defaultdict(Counter)
    for valores in wb["Item de Malha"].iter_rows(min_row=2, values_only=True):
        if valores and len(valores) > 2 and valores[1] and valores[2]:
            malhas[chave_linha(valores[1])][str(valores[2]).strip()] += 1
    # Linha em mais de uma malha: fica na mais frequente (empate: a de menor código), para não contar a junta duas vezes.
    pasta_da_linha = {k: sorted(c.items(), key=lambda kv: (-kv[1], kv[0]))[0][0] for k, c in malhas.items()}
    return sub_da_linha, pasta_da_linha


def main(sgj, olodum, saida):
    sub_da_linha, pasta_da_linha = ler_olodum(olodum)
    wb = openpyxl.load_workbook(sgj, read_only=True, data_only=True)
    agg = defaultdict(lambda: dict(previstas=0, executadas=0))
    sem_ligacao = 0
    for r in wb["SGJ"].iter_rows(min_row=LINHA_INICIAL, values_only=True):
        if not r or len(r) <= COL_SITUACAO or not r[COL_LINHA]:
            continue
        linha = chave_linha(r[COL_LINHA])
        if r[COL_SUBSOP]:
            subsistema = str(r[COL_SUBSOP]).strip()
            sistema = str(r[COL_SOP]).strip() if r[COL_SOP] else sistema_do_subsistema(subsistema)
        elif linha in sub_da_linha:
            sistema, subsistema = sub_da_linha[linha]
        else:
            sistema, subsistema = SEM_SISTEMA, SEM_SUBSISTEMA
            sem_ligacao += 1
        pasta = pasta_da_linha.get(linha, SEM_PASTA)
        a = agg[(sistema, subsistema, pasta)]
        a["previstas"] += 1
        a["executadas"] += str(r[COL_SITUACAO] or "").strip().lower() == "junta liberada"
    rows = [dict(sistema=s, subsistema=sub, pasta=pa, **agg[(s, sub, pa)]) for (s, sub, pa) in sorted(agg)]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    total = sum(r["previstas"] for r in rows)
    print(f"{len(rows)} linhas, {total} juntas ({sem_ligacao} sem subsistema)")


if __name__ == "__main__":
    main(*sys.argv[1:4])
