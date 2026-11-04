"""Gera a base de juntas do ControlTub para o painel, a partir da 01.Situacao_Geral_de_Juntas.

Uso: python painel_juntas.py <01.Situacao_Geral_de_Juntas.xlsx> <saida.json>

Duas somas na mesma saída, cada linha com:
  previstas   = juntas da planilha
  executadas  = juntas com Situação "Junta Liberada"
tipo "sub":   por sistema x subsistema, para o avanço por subsistema. Subsistema: a coluna SUBSOP
              (sistema: a SOP); junta sem SUBSOP vai para "Subsistema não cadastrado".
tipo "pasta": por pasta, para o avanço por pasta. Pasta: a coluna STH (onde fica o Nº da pasta).
              Junta sem STH vai para "Pasta não cadastrada" com "Sistema não cadastrado".
"""
import json
import sys
from collections import defaultdict

import openpyxl

SEM_SISTEMA = "Sistema não cadastrado"
SEM_SUBSISTEMA = "Subsistema não cadastrado"
SEM_PASTA = "Pasta não cadastrada"
COL_LINHA, COL_SOP, COL_SUBSOP, COL_STH, COL_SITUACAO = 2, 4, 6, 7, 93  # colunas da aba SGJ (0 = coluna A)
LINHA_INICIAL = 11  # dados começam depois das duas linhas de cabeçalho


def sistema_do_subsistema(sub):
    """'421-PRO-MR-252-001' -> '421-PRO-MR-252'."""
    return sub.rsplit("-", 1)[0] if "-" in sub else sub


def main(sgj, saida):
    wb = openpyxl.load_workbook(sgj, read_only=True, data_only=True)
    agg = defaultdict(lambda: dict(previstas=0, executadas=0))
    sem_ligacao = 0
    for r in wb["SGJ"].iter_rows(min_row=LINHA_INICIAL, values_only=True):
        if not r or len(r) <= COL_SITUACAO or not r[COL_LINHA]:
            continue
        if r[COL_SUBSOP]:
            subsistema = str(r[COL_SUBSOP]).strip()
            sistema = str(r[COL_SOP]).strip() if r[COL_SOP] else sistema_do_subsistema(subsistema)
        else:
            sistema, subsistema = SEM_SISTEMA, SEM_SUBSISTEMA
            sem_ligacao += 1
        liberada = str(r[COL_SITUACAO] or "").strip().lower() == "junta liberada"
        sth = str(r[COL_STH]).strip() if r[COL_STH] else ""
        if sth:
            chave_pasta = ("pasta", sistema, subsistema, sth)
        else:
            chave_pasta = ("pasta", SEM_SISTEMA, SEM_SUBSISTEMA, SEM_PASTA)
        for k in (("sub", sistema, subsistema, ""), chave_pasta):
            agg[k]["previstas"] += 1
            agg[k]["executadas"] += liberada
    rows = [dict(tipo=t, sistema=s, subsistema=sub, pasta=pa, **agg[(t, s, sub, pa)]) for (t, s, sub, pa) in sorted(agg)]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    total = sum(r["previstas"] for r in rows if r["tipo"] == "sub")
    print(f"{len(rows)} linhas, {total} juntas ({sem_ligacao} sem subsistema)")


if __name__ == "__main__":
    main(*sys.argv[1:3])
