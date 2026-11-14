"""Gera a relação pasta -> escopo para o painel, a partir da 06-Lista de Pastas_Milplan.

Uso: python painel_escopos.py <06-Lista de Pastas_Milplan.xlsx> <saida.json>

Aba MILPLAN: coluna G (ESCOPO) e coluna N (Nº DA PASTA - PARA). Cada linha da saída é um par pasta x escopo
(uma pasta pode ter mais de um escopo). Pastas "CANCELADO", "N/A" ou sem escopo ficam de fora.
"""
import json
import sys

import openpyxl

ABA = "MILPLAN"
LINHA_INICIAL = 5  # os dados começam depois do cabeçalho da linha 4
COL_ESCOPO, COL_PASTA = 6, 13  # colunas G e N (0 = coluna A)
IGNORAR = {"", "CANCELADO", "N/A", "#N/A", "#VALUE!"}


def main(planilha, saida):
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    pares = set()
    for r in wb[ABA].iter_rows(min_row=LINHA_INICIAL, values_only=True):
        if not r or len(r) <= COL_PASTA:
            continue
        pasta = str(r[COL_PASTA] or "").strip()
        escopo = str(r[COL_ESCOPO] or "").strip()
        if pasta.upper() in IGNORAR or escopo.upper() in IGNORAR:
            continue
        pares.add((pasta, escopo))
    rows = [dict(pasta=p, escopo=e) for p, e in sorted(pares)]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"{len({p for p, _ in pares})} pastas com escopo ({len(rows)} pares pasta x escopo)")


if __name__ == "__main__":
    main(*sys.argv[1:3])
