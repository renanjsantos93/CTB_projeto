"""Gera a base de pendências do painel a partir da 05-Planilha_pendencia.

Uso: python painel_pendencias.py <05-Planilha_pendencia.xlsx> <saida.json> [<pcp_painel.json>] [<lista.json>]

Cada linha da saída é um par subsistema x categoria (A, B, RNC) x disciplina x etapa x skid, com:
  total       = pendências cadastradas
  encerradas  = pendências com STATUS "FECHADA" (o avanço); as demais estão abertas
O pcp_painel.json, quando dado, só serve para escrever o skid com o mesmo nome usado na base de atividades.

Com <lista.json>, grava também a lista das pendências ABERTAS, uma por linha, com item, descrição,
categoria, solução proposta e tipo (o que o painel mostra na "Lista de pendências abertas").
"""
import re
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


NOMES_CATEGORIA = {"A": "A - Impeditiva", "B": "B - Não impeditiva", "RNC": "RNC - Não conformidade"}


def texto(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def sem_codigo(v):
    """'(8) Identificação' -> 'Identificação'."""
    return re.sub(r"^\(\s*\d+\s*\)\s*", "", texto(v))


def nomes_de_skid(pcp_json):
    """Código do skid -> nome usado na base de atividades (ex.: 'M07' -> 'M07 - MR High Pressure Compressor A')."""
    if not pcp_json:
        return {}
    with open(pcp_json, encoding="utf-8") as f:
        return {codigo(r["skid"]): r["skid"] for r in json.load(f) if r["skid"] != SEM_SKID}


def main(planilha, saida, pcp_json=None, saida_lista=None):
    skids = nomes_de_skid(pcp_json)
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    linhas = wb.worksheets[0].iter_rows(values_only=True)
    cab = next(linhas)
    agg = defaultdict(lambda: dict(total=0, encerradas=0))
    nomes = {}
    abertas = []
    ignoradas = 0
    for valores in linhas:
        if not any(valores):
            continue
        r = dict(zip(cab, valores))
        cat = categoria(r.get("CATEGORIA"))
        if not cat:
            ignoradas += 1
            continue
        # Sem subsistema, a pendência vai para a linha única "Subsistema não cadastrado", sem ligar ao sistema.
        if r.get("SUBSISTEMA"):
            sistema = codigo(r["SISTEMA"]) if r.get("SISTEMA") else SEM_SISTEMA
            subsistema = codigo(r["SUBSISTEMA"])
        else:
            sistema, subsistema = SEM_SISTEMA, SEM_SUBSISTEMA
        nomes.setdefault((sistema, subsistema), r.get("NOME DO SUBSISTEMA") or "")
        disciplina = DISCIPLINAS.get(str(r.get("DISCIPLINA") or "").strip().upper(), "Sem disciplina")
        etapa = etapa_da_fase(r.get("FASE")) if r.get("FASE") else "Sem etapa"
        skid_cod = str(r.get("SKID") or "").strip()
        skid = skids.get(skid_cod, skid_cod) if skid_cod and skid_cod.upper() != "N/A" else SEM_SKID
        a = agg[(sistema, subsistema, cat, disciplina, etapa, skid)]
        a["total"] += 1
        fechada = str(r.get("STATUS") or "").strip().upper() == "FECHADA"
        a["encerradas"] += fechada
        if not fechada:
            abertas.append(dict(
                numero=r.get("NÚMERO PENDÊNCIA"), sistema=sistema, subsistema=subsistema,
                item=texto(r.get("ITEM")), descricao=texto(r.get("DESCRIÇÃO")), categoria=NOMES_CATEGORIA[cat],
                solucao=texto(r.get("SOLUÇÃO PROPOSTA")), tipo=sem_codigo(r.get("TIPO")),
                disciplina=disciplina, etapa=etapa, skid=skid))

    rows = [
        dict(sistema=s, subsistema=sub, desc_subsistema=nomes[(s, sub)], categoria=c,
             disciplina=d, etapa=e, skid=k, **agg[(s, sub, c, d, e, k)])
        for (s, sub, c, d, e, k) in sorted(agg)
    ]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    if saida_lista:
        for a in abertas:
            a["desc_subsistema"] = nomes.get((a["sistema"], a["subsistema"]), "")
        with open(saida_lista, "w", encoding="utf-8") as f:
            json.dump(abertas, f, ensure_ascii=False)
    total = sum(r["total"] for r in rows)
    print(f"{len(rows)} linhas, {total} pendências" + (f" ({ignoradas} sem categoria A/B/RNC ignoradas)" if ignoradas else ""))


if __name__ == "__main__":
    main(*sys.argv[1:5])
