"""Gera a base agregada do Painel por Sistema a partir da aba PCP_ATIVIDADE.

Uso: python scripts/painel_sistema_pcp.py <planilha.xlsx> <saida.json>

Cada linha da saída é um par subsistema x atividade x data planejada x skid x área, com as contagens:
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


def data_planejada(v):
    """Data ISO (aaaa-mm-dd), ou o texto da célula quando não é data."""
    if hasattr(v, "date"):
        return v.date().isoformat()
    texto = str(v or "").strip()
    return "Sem data planejada" if texto in ("", "#N/A") else texto


def depara(v, vazio):
    """Valor das colunas DEPARA DE SKID / DEPARA DE AREA, sem o " - " solto no fim."""
    texto = re.sub(r"\s*-\s*$", "", str(v or "").strip())
    return vazio if texto in ("", "#N/A", "0", "Pendente Cadastro area") else texto


def preenchido(v):
    return v not in (None, "", 0, "-")


def main(planilha, saida):
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    linhas = wb["PCP_ATIVIDADE"].iter_rows(values_only=True)
    cab = next(linhas)
    agg = defaultdict(lambda: dict(previsto=0, executado=0, verificado=0, aprovado=0, reprovado=0))
    descricoes = {}
    desc_sub = {}
    for valores in linhas:
        if not any(valores):
            continue
        r = dict(zip(cab, valores))
        sistema = r["SISTEMA"] or "Sistema não cadastrado"
        descricoes.setdefault(sistema, r["DESCRIÇÃO DO SISTEMA"] or "")
        subsistema = r["SUBSISTEMA"] or "Subsistema não cadastrado"
        desc_sub.setdefault(subsistema, r["DESCRIÇÃO DO SUBSISTEMA"] or "")
        sigla = r["DISCIPLINA"] or ""
        disciplina = DISCIPLINAS.get(r["DISCIPLINA"], r["DISCIPLINA"] or "Sem disciplina")
        etapa = etapa_da_fase(r["FASE"])
        atividade = re.sub(r"\s*\(\s*449\.[^)]*\)\s*$", "", r["ATIVIDADE"] or "").strip()
        atividade = re.sub(r"\s*-\s*$", "", atividade)
        planejada = data_planejada(r.get("DATA PLANEJADA"))
        skid = depara(r.get("DEPARA DE SKID"), "Skid não cadastrado")
        area = depara(r.get("DEPARA DE AREA"), "Área não cadastrada")
        a = agg[(sistema, subsistema, sigla, disciplina, etapa, atividade, planejada, skid, area)]
        a["previsto"] += 1
        a["executado"] += preenchido(r["DATA DA EXECUÇÃO"])
        a["verificado"] += preenchido(r["DATA DA VERIFICAÇÃO"])
        a["aprovado"] += preenchido(r["DATA DA APROVAÇÃO"])
        a["reprovado"] += preenchido(r["DATA DA REPROVAÇÃO"])

    ordem_etapa = {e: i for i, (_, e) in enumerate(ETAPAS)}
    chaves = sorted(agg, key=lambda k: (k[0], k[1], k[3], ordem_etapa.get(k[4], 99), k[5], k[6], k[7], k[8]))
    rows = [
        dict(id=f"{i + 1:05d}", sistema=s, descricao=descricoes[s], subsistema=sub,
             desc_subsistema=desc_sub[sub], sigla=sg, disciplina=d, etapa=e, atividade=at,
             data_planejada=dp, skid=sk, area=ar, **agg[(s, sub, sg, d, e, at, dp, sk, ar)])
        for i, (s, sub, sg, d, e, at, dp, sk, ar) in enumerate(chaves)
    ]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"{len(rows)} linhas, {sum(r['previsto'] for r in rows)} registros")


if __name__ == "__main__":
    main(*sys.argv[1:3])
