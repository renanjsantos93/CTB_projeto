"""Gera a base agregada do Painel por Sistema a partir da aba PCP_ATIVIDADE.

Uso: python scripts/painel_sistema_pcp.py <planilha.xlsx> <saida.json>

Cada linha da saída é um par subsistema x atividade x data planejada x skid x área x fornecedor x pasta x grupo
(GRUPO DE SUBSISTEMA), com as contagens:
  previsto   = registros cadastrados na PCP_ATIVIDADE
  executado  = registros com DATA DA EXECUÇÃO
  verificado = registros com DATA DA VERIFICAÇÃO
  aprovado   = registros com DATA DA APROVAÇÃO (registro emitido)
  reprovado  = registros com DATA DA REPROVAÇÃO
"""
import datetime
import json
import os
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
SEM_PASTA = "Pasta não cadastrada"
SEM_GRUPO = "Grupo não cadastrado"
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
    """Data ISO (aaaa-mm-dd), ou o texto da célula quando não é data. Aceita data do Excel, número de série
    do Excel (ex.: 46371) e texto dd/mm/aaaa ou aaaa-mm-dd."""
    if hasattr(v, "date"):
        return v.date().isoformat()
    if isinstance(v, (int, float)) and 30000 < v < 80000:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()
    texto = str(v or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", texto)
    if m:
        return "%s-%02d-%02d" % (m.group(3), int(m.group(2)), int(m.group(1)))
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", texto)
    if m:
        return m.group(0)
    return "Sem data planejada" if texto.upper() in ("", "#N/A", "#N/D", "#VALUE!", "#VALOR!", "#REF!", "0") else texto


def coluna(cab, nome):
    """Nome da coluna do cabeçalho igual a `nome`, ignorando espaços, acentos e maiúsculas
    (ex.: "DATA PLANEJADA ", "Data Planejada", "DATAPLANEJADA")."""
    import unicodedata
    def chave(t):
        t = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode()
        return re.sub(r"[^A-Z0-9]", "", t.upper())
    alvo = chave(nome)
    return next((c for c in cab if chave(c) == alvo), nome)


def depara(v, vazio):
    """Valor das colunas DEPARA DE SKID / DEPARA DE AREA, sem o " - " solto no fim."""
    texto = re.sub(r"\s*-\s*$", "", str(v or "").strip())
    return vazio if texto in ("", "#N/A", "#VALUE!", "0", "Pendente Cadastro area") else texto


def preenchido(v):
    return v not in (None, "", 0, "-")


def data_da_base(planilha, wb):
    """Data (UTC) mostrada no "Atualizado em" do painel: quando o 01-PCPHMWEB foi salvo na pasta
    01-Base de dados, ao lado da OLODUM. Sem esse arquivo, quando a própria OLODUM foi salva."""
    pasta_bases = os.path.join(os.path.dirname(os.path.abspath(planilha)), "01-Base de dados")
    try:
        pcp = [os.path.join(pasta_bases, n) for n in os.listdir(pasta_bases)
               if n.upper().startswith("01-PCPHMWEB") and ".xls" in n.lower() and not n.startswith("~$")]
    except OSError:
        pcp = []
    if pcp:
        ts = max(os.path.getmtime(p) for p in pcp)
        return datetime.datetime.fromtimestamp(ts, datetime.timezone.utc).replace(tzinfo=None)
    return wb.properties.modified or datetime.datetime.fromtimestamp(
        os.path.getmtime(planilha), datetime.timezone.utc).replace(tzinfo=None)


def main(planilha, saida, saida_pendentes=None):
    """Com <saida_pendentes>, grava também cada protocolo ainda não aprovado (um registro da PCP_ATIVIDADE
    sem DATA DA APROVAÇÃO), com a tag, o objeto e a situação, para o quadro "Atividades pendentes"."""
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    atualizado_em = data_da_base(planilha, wb).replace(microsecond=0).isoformat() + "Z"
    linhas = wb["PCP_ATIVIDADE"].iter_rows(values_only=True)
    cab = next(linhas)
    col_planejada = coluna(cab, "DATA PLANEJADA")
    col_grupo = coluna(cab, "GRUPO DE SUBSISTEMA")
    if col_planejada not in cab:
        print("AVISO: a coluna DATA PLANEJADA não existe na PCP_ATIVIDADE; o painel fica sem datas planejadas")
    agg = defaultdict(lambda: dict(previsto=0, executado=0, verificado=0, aprovado=0, reprovado=0))
    descricoes = {}
    desc_sub = {}
    pendentes = []
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
        planejada = data_planejada(r.get(col_planejada))
        skid = depara(r.get("DEPARA DE SKID"), "Skid não cadastrado")
        area = depara(r.get("DEPARA DE AREA"), "Área não cadastrada")
        fornecedor = depara(r.get("FORNECEDOR"), "Fornecedor não cadastrado")
        pasta = depara(r.get("Nº DA PASTA"), SEM_PASTA)
        grupo = depara(r.get(col_grupo), SEM_GRUPO)
        a = agg[(sistema, subsistema, sigla, disciplina, etapa, atividade, planejada, skid, area, fornecedor, pasta, grupo)]
        a["previsto"] += 1
        a["executado"] += preenchido(r["DATA DA EXECUÇÃO"])
        a["verificado"] += preenchido(r["DATA DA VERIFICAÇÃO"])
        a["aprovado"] += preenchido(r["DATA DA APROVAÇÃO"])
        a["reprovado"] += preenchido(r["DATA DA REPROVAÇÃO"])
        if saida_pendentes and not preenchido(r["DATA DA APROVAÇÃO"]):
            situacao = ("Reprovada" if preenchido(r["DATA DA REPROVAÇÃO"])
                        else "Executada, aguardando aprovação" if preenchido(r["DATA DA EXECUÇÃO"])
                        else "Não executada")
            pendentes.append(dict(
                sistema=sistema, subsistema=subsistema, pasta=pasta, disciplina=disciplina, etapa=etapa,
                atividade=atividade, data_planejada=planejada, skid=skid, area=area, fornecedor=fornecedor, grupo=grupo,
                tag=str(r.get("COD IDENTIFICAÇÃO") or "").strip(), objeto=str(r.get("DESCRIÇÃO DO OBJETO") or "").strip(),
                situacao=situacao))

    # Nº DA PASTA preenchido com o código de um subsistema ou sistema não é pasta: fica com pasta vazia ("")
    # e o quadro de avanço por pasta não conta essas atividades (no resto do painel elas contam normalmente).
    codigos = {k[0] for k in agg} | {k[1] for k in agg}
    if any(k[10] in codigos for k in agg):
        juntos = defaultdict(lambda: dict(previsto=0, executado=0, verificado=0, aprovado=0, reprovado=0))
        for k, v in agg.items():
            k2 = k[:10] + ("" if k[10] in codigos else k[10],) + k[11:]
            for campo, n in v.items():
                juntos[k2][campo] += n
        agg = juntos

    ordem_etapa = {e: i for i, (_, e) in enumerate(ETAPAS)}
    chaves = sorted(agg, key=lambda k: (k[0], k[1], k[3], ordem_etapa.get(k[4], 99), k[5], k[6], k[7], k[8], k[9], k[10], k[11]))
    rows = [
        dict(id=f"{i + 1:05d}", sistema=s, descricao=descricoes[s], subsistema=sub,
             desc_subsistema=desc_sub[sub], sigla=sg, disciplina=d, etapa=e, atividade=at,
             data_planejada=dp, skid=sk, area=ar, fornecedor=fo, pasta=pa, grupo=gr, atualizado_em=atualizado_em,
             **agg[(s, sub, sg, d, e, at, dp, sk, ar, fo, pa, gr)])
        for i, (s, sub, sg, d, e, at, dp, sk, ar, fo, pa, gr) in enumerate(chaves)
    ]
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    if saida_pendentes:
        for r in pendentes:
            if r["pasta"] in codigos:
                r["pasta"] = ""
        with open(saida_pendentes, "w", encoding="utf-8") as f:
            json.dump(pendentes, f, ensure_ascii=False, separators=(",", ":"))
    com_data = sum(r["previsto"] for r in rows if re.match(r"^\d{4}-\d{2}-\d{2}$", r["data_planejada"]))
    print(f"{len(rows)} linhas, {sum(r['previsto'] for r in rows)} registros, {com_data} com data planejada")
    if not com_data:
        print("AVISO: nenhum registro com data planejada; confira a coluna DATA PLANEJADA da PCP_ATIVIDADE "
              "(se está com #N/D, a consulta da 04-Dataplanejamento não atualizou)")


if __name__ == "__main__":
    main(*sys.argv[1:4])
