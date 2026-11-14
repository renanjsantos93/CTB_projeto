"""Gera a relação pasta -> escopo e limpeza química para o painel, a partir da 06-Lista de Pastas_Milplan.

Uso: python painel_escopos.py <06-Lista de Pastas_Milplan.xlsx> <saida.json>

Aba MILPLAN, pela coluna N (Nº DA PASTA - PARA):
  coluna G (ESCOPO)           -> escopo da pasta (uma pasta pode ter mais de um escopo)
  coluna I (Limpeza Química)  -> "X" em alguma linha da pasta = a pasta precisa de limpeza química
Cada linha da saída é um par pasta x escopo, com limpeza_quimica (true/false) da pasta; uma pasta com limpeza
química e sem escopo sai com escopo "". Pastas "CANCELADO", "N/A" ou vazias ficam de fora.
"""
import json
import sys
from collections import defaultdict

import openpyxl

ABA = "MILPLAN"
LINHA_INICIAL = 5  # os dados começam depois do cabeçalho da linha 4
COL_ESCOPO, COL_LQ, COL_PASTA = 6, 8, 13  # colunas G, I e N (0 = coluna A)
IGNORAR = {"", "CANCELADO", "N/A", "#N/A", "#VALUE!"}


def main(planilha, saida):
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    escopos = defaultdict(set)
    com_lq = set()
    for r in wb[ABA].iter_rows(min_row=LINHA_INICIAL, values_only=True):
        if not r or len(r) <= COL_PASTA:
            continue
        pasta = str(r[COL_PASTA] or "").strip()
        if pasta.upper() in IGNORAR:
            continue
        escopo = str(r[COL_ESCOPO] or "").strip()
        if escopo.upper() not in IGNORAR:
            escopos[pasta].add(escopo)
        if str(r[COL_LQ] or "").strip().upper() == "X":
            com_lq.add(pasta)
    rows = [dict(pasta=p, escopo=e, limpeza_quimica=p in com_lq)
            for p in sorted(set(escopos) | com_lq) for e in (sorted(escopos[p]) or [""])]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"{len(escopos)} pastas com escopo ({sum(map(len, escopos.values()))} pares pasta x escopo), "
          f"{len(com_lq)} pastas com limpeza química")


if __name__ == "__main__":
    main(*sys.argv[1:3])
