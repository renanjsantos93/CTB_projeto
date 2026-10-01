"""Relação de relatórios do databook 395-421 (versão escaneada, 27 relatórios).

O PDF é escaneado (sem camada de texto): nº, data e desenho de cada relatório
foram transcritos da imagem na tabela RELATORIOS abaixo. Os croquis são os
mesmos dos relatórios 001-027 do databook BRASNAVAL, então spools e
observações vêm de relacao_databook.OBS_CROQUI.

Uso:
    python scripts/relacao_databook_395_421.py SAIDA.xlsx
"""
import sys
from datetime import date

from openpyxl import Workbook

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from relacao_databook import (CROQUI_IGUAL, OBS_CROQUI, escrever_aba,  # noqa: E402
                              spools, tag_do_desenho)

# (nº do relatório, data, desenho de referência) — página = nº do relatório
RELATORIOS = [
    ("001", date(2026, 8, 25), "24``-B1E-MR-4012-Z02"),
    ("002", date(2026, 8, 25), "24``-B1E-MR-5003A-Z03"),
    ("003", date(2026, 8, 25), "I-IS-025-421-510-46-9069"),
    ("004", date(2026, 9, 1), "MO7-4´´-B3B-LO-5325A-PP-2"),
    ("005", date(2026, 8, 26), "6´´-A1E-CF-4004A-Z02"),
    ("006", date(2026, 8, 26), "8´´-B3B-MR-4019A-Z02"),
    ("007", date(2026, 8, 26), "4´´-A3B-HF-4006A-Z03"),
    ("008", date(2026, 8, 26), '4"-A3B-HF-5001A-Z03'),
    ("009", date(2026, 8, 26), "2´´-A1E-IA-911-Z03"),
    ("010", date(2026, 8, 27), "I-IS-025-421-510-46-9059 -FOLHA 1/3"),
    ("011", date(2026, 8, 26), "6´´-D1E-NG-4301A-C-3´´-Z02"),
    ("012", date(2026, 9, 1), "MO8-6´´-B3B-MR-5405B-PP-2 / FOLHA 2/2"),
    ("013", date(2026, 9, 1), "MO5-4´´-B3B-LO-5119A-PP-2"),
    ("014", date(2026, 8, 31), "I-IS-025-421-510-46-9147"),
    ("015", date(2026, 8, 31), "I-IS-025-421-510-46-9148"),
    ("016", date(2026, 8, 26), "I-IS-025-421-510-46-9150"),
    ("017", date(2026, 8, 26), "I-IS-025-421-510-46-9151"),
    ("018", date(2026, 8, 25), "I-IS-025-421-510-46-9160"),
    ("019", date(2026, 8, 26), "I-IS-025-421-510-9011"),
    ("020", date(2026, 8, 26), "I-IS-025-421-510-46-9165"),
    ("021", date(2026, 8, 25), "I-IS-025-421-510-46-9138"),
    ("022", date(2026, 8, 31), "MO7-4´´-B3B-LO-5309A-PP-2"),
    ("023", date(2026, 9, 1), "MO6-4´´-B3B-LO-5113B-PP-2"),
    ("024", date(2026, 9, 1), "MO7-6´´-B3B-MR-5405A-PP-2"),
    ("025", date(2026, 9, 1), "MO5-4´´-B3B-MR-5113A-PP-2"),
    ("026", date(2026, 9, 1), "MO8-6´´-B3B-MR-5405B-PP-2"),
    ("027", date(2026, 9, 1), "MO5-10´´-B3B-LO-5103A-2"),
]


def observacao(r, rels):
    pag = r["pagina"]
    obs = []
    outros = [o["numero"] for o in rels if o["desenho"] == r["desenho"] and o is not r]
    if outros:
        obs.append("Mesmo desenho dos relatórios " + ", ".join(outros))
    if r["obs_tag"]:
        obs.append(r["obs_tag"])
    if not OBS_CROQUI[pag][0]:
        obs.append("Spool não identificado no croqui – considerado spool único 001")
    if CROQUI_IGUAL.get(pag, 99) <= len(rels):
        obs.append(f"Croqui idêntico ao do relatório {CROQUI_IGUAL[pag]:03d}")
    if OBS_CROQUI[pag][1]:
        obs.append("Croqui: " + OBS_CROQUI[pag][1])
    return "; ".join(obs)


def main(saida):
    rels = []
    for num, data, desenho in RELATORIOS:
        tag, obs_tag = tag_do_desenho(desenho)
        rels.append({"pagina": int(num), "numero": num, "tipo": "Relatório Dimensional",
                     "data": data, "desenho": desenho, "tag": tag, "obs_tag": obs_tag,
                     "laudo": "Aprovado"})
    wb = Workbook()

    linhas, item = [], 0
    for r in rels:
        for n in spools(r):
            item += 1
            linhas.append([item, r["numero"], r["tipo"], r["data"], r["desenho"], n,
                           r["tag"], r["laudo"], r["pagina"], observacao(r, rels)])
    ws = wb.active
    ws.title = "Relação por spool"
    escrever_aba(ws, ["Item", "Nº do relatório", "Tipo de relatório", "Data do relatório",
                      "Documento (desenho de referência)", "Spool", "TAG PADRONIZADA",
                      "Laudo", "Página no databook", "Observação"],
                 linhas, [6, 11, 20, 12, 40, 9, 22, 10, 10, 70], alerta_col=9)

    linhas = []
    for r in rels:
        sp = spools(r)
        linhas.append([r["numero"], r["tipo"], r["data"], r["desenho"], r["tag"],
                       ", ".join(sp), len(sp), r["laudo"], r["pagina"],
                       observacao(r, rels)])
    ws = wb.create_sheet("Resumo por relatório")
    escrever_aba(ws, ["Nº do relatório", "Tipo de relatório", "Data do relatório",
                      "Documento (desenho de referência)", "TAG PADRONIZADA", "Spools",
                      "Qtd. spools", "Laudo", "Página no databook", "Observação"],
                 linhas, [11, 20, 12, 40, 22, 18, 8, 10, 10, 70], alerta_col=9)

    wb.save(saida)
    print(f"{len(rels)} relatórios -> {saida}")


if __name__ == "__main__":
    main(sys.argv[1])
