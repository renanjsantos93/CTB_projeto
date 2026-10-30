"""Gera a base agregada do Painel por Sistema a partir da aba PCP_ATIVIDADE.

Uso: python scripts/painel_sistema_pcp.py <planilha.xlsx> <saida.json>

Cada linha da saída é um par sistema x atividade, com as contagens:
  previsto   = registros cadastrados na PCP_ATIVIDADE
  executado  = registros com DATA DA EXECUÇÃO
  verificado = registros com DATA DA VERIFICAÇÃO
  aprovado   = registros com DATA DA APROVAÇÃO (registro emitido)
  reprovado  = registros com DATA DA REPROVAÇÃO
"""
import json
import re
import sys
from collections import defaultdict

import openpyxl

ETAPAS = [
    ("Recebimento", "Recebimento"),
    ("Construção e Montagem", "Montagem"),
    ("Condicionamento", "Condicionamento"),
    ("Comissionamento", "Comissionamento"),
    ("COMPLETAÇÃO MECÂNICA", "Completação mecânica"),
    ("TTAS", "TTAS"),
    ("Preservação", "Preservação"),
]
DISCIPLINAS = {
    "TUB": "Tubulação", "INS": "Instrumentação", "ELE": "Elétrica",
    "EQE": "Equipamento estático", "EQD": "Equipamento dinâmico",
    "AUT": "Automação", "CIV": "Civil", "ESTR": "Estrutura", "GER": "Geral",
}


def etapa_da_fase(fase):
    for prefixo, etapa in ETAPAS:
        if (fase or "").startswith(prefixo):
            return etapa
    return fase or "Sem fase"


def preenchido(v):
    return v not in (None, "", 0, "-")


def main(planilha, saida):
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    linhas = wb["PCP_ATIVIDADE"].iter_rows(values_only=True)
    cab = next(linhas)
    agg = defaultdict(lambda: dict(previsto=0, executado=0, verificado=0, aprovado=0, reprovado=0))
    descricoes = {}
    for valores in linhas:
        if not any(valores):
            continue
        r = dict(zip(cab, valores))
        sistema = r["SISTEMA"] or "Sem sistema"
        descricoes.setdefault(sistema, r["DESCRIÇÃO DO SISTEMA"] or "")
        disciplina = DISCIPLINAS.get(r["DISCIPLINA"], r["DISCIPLINA"] or "Sem disciplina")
        etapa = etapa_da_fase(r["FASE"])
        atividade = re.sub(r"\s*\(\s*449\.[^)]*\)\s*$", "", r["ATIVIDADE"] or "").strip()
        a = agg[(sistema, disciplina, etapa, atividade)]
        a["previsto"] += 1
        a["executado"] += preenchido(r["DATA DA EXECUÇÃO"])
        a["verificado"] += preenchido(r["DATA DA VERIFICAÇÃO"])
        a["aprovado"] += preenchido(r["DATA DA APROVAÇÃO"])
        a["reprovado"] += preenchido(r["DATA DA REPROVAÇÃO"])

    ordem_etapa = {e: i for i, (_, e) in enumerate(ETAPAS)}
    chaves = sorted(agg, key=lambda k: (k[0], k[1], ordem_etapa.get(k[2], 99), k[3]))
    rows = [
        dict(id=f"{i + 1:04d}", sistema=s, descricao=descricoes[s], disciplina=d,
             etapa=e, atividade=at, **agg[(s, d, e, at)])
        for i, (s, d, e, at) in enumerate(chaves)
    ]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"{len(rows)} linhas, {sum(r['previsto'] for r in rows)} registros")


if __name__ == "__main__":
    main(*sys.argv[1:3])
