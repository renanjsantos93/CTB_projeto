"""Gera a base de pendências do painel a partir da 05-Planilha_pendencia.

Uso: python painel_pendencias.py <05-Planilha_pendencia.xlsx> <saida.json> [<pcp_painel.json>]

Cada linha da saída é um par subsistema x categoria (A, B, RNC) x disciplina x etapa x skid, com:
  total       = pendências cadastradas
  encerradas  = pendências com STATUS "FECHADA" (o avanço); as demais estão abertas
O pcp_painel.json, quando dado, só serve para escrever o skid com o mesmo nome usado na base de atividades.
"""
import json
import sys
from collections import defaultdict

import openpyxl

from painel_sistema_pcp import etapa_da_fase

SEM_SISTEMA = "Sistema não cadastrado"
SEM_SUBSISTEMA = "Subsistema não cadastrado"
SEM_SKID = "Skid não cadastrado"
CATEGORIAS = {"A": "A", "B": "B", "RNC": "RNC"}  # o código entre parênteses no início da CATEGORIA
DISCIPLINAS = {
    "TUBULAÇÃO": "Tubulação", "INSTRUMENTAÇÃO": "Instrumentação", "ELÉTRICA": "Elétrica",
    "EQUIPAMENTOS ESTÁTICOS": "Equipamento estático", "EQUIPAMENTOS DINÂMICOS": "Equipamento dinâmico",
    "AUTOMAÇÃO": "Automação", "ESTRUTURA METÁLICA": "Estrutura", "CIVIL": "Civil", "GERAL": "Geral",
}


def codigo(v):
    """'421-PRO-MR-252 - MR COMPRESSOR' -> '421-PRO-MR-252'."""
    return str(v).split(" - ")[0].strip()


def categoria(v):
    texto = str(v or "").strip()
    if texto.startswith("(") and ")" in texto:
        return CATEGORIAS.get(texto[1:texto.index(")")].strip().upper())
    return None


def nomes_de_skid(pcp_json):
    """Código do skid -> nome usado na base de atividades (ex.: 'M07' -> 'M07 - MR High Pressure Compressor A')."""
    if not pcp_json:
        return {}
    with open(pcp_json, encoding="utf-8") as f:
        return {codigo(r["skid"]): r["skid"] for r in json.load(f) if r["skid"] != SEM_SKID}


def main(planilha, saida, pcp_json=None):
    skids = nomes_de_skid(pcp_json)
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    linhas = wb.worksheets[0].iter_rows(values_only=True)
    cab = next(linhas)
    agg = defaultdict(lambda: dict(total=0, encerradas=0))
    nomes = {}
    ignoradas = 0
    for valores in linhas:
        if not any(valores):
            continue
        r = dict(zip(cab, valores))
        cat = categoria(r.get("CATEGORIA"))
        if not cat:
            ignoradas += 1
            continue
        sistema = codigo(r["SISTEMA"]) if r.get("SISTEMA") else SEM_SISTEMA
        subsistema = codigo(r["SUBSISTEMA"]) if r.get("SUBSISTEMA") else SEM_SUBSISTEMA
        nomes.setdefault((sistema, subsistema), r.get("NOME DO SUBSISTEMA") or "")
        disciplina = DISCIPLINAS.get(str(r.get("DISCIPLINA") or "").strip().upper(), "Sem disciplina")
        etapa = etapa_da_fase(r.get("FASE")) if r.get("FASE") else "Sem etapa"
        skid_cod = str(r.get("SKID") or "").strip()
        skid = skids.get(skid_cod, skid_cod) if skid_cod and skid_cod.upper() != "N/A" else SEM_SKID
        a = agg[(sistema, subsistema, cat, disciplina, etapa, skid)]
        a["total"] += 1
        a["encerradas"] += str(r.get("STATUS") or "").strip().upper() == "FECHADA"

    rows = [
        dict(sistema=s, subsistema=sub, desc_subsistema=nomes[(s, sub)], categoria=c,
             disciplina=d, etapa=e, skid=k, **agg[(s, sub, c, d, e, k)])
        for (s, sub, c, d, e, k) in sorted(agg)
    ]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    total = sum(r["total"] for r in rows)
    print(f"{len(rows)} linhas, {total} pendências" + (f" ({ignoradas} sem categoria A/B/RNC ignoradas)" if ignoradas else ""))


if __name__ == "__main__":
    main(*sys.argv[1:4])
