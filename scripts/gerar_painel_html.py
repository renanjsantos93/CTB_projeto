"""Gera o Painel_OLODUM.html: o painel completo num arquivo só, com a base embutida.

Abre em qualquer navegador, sem internet e sem conector; quem tem acesso à pasta do OneDrive vê o painel.

Uso: python gerar_painel_html.py <pcp_painel.json> <Painel_OLODUM.html>
"""
import json
import sys
from pathlib import Path

TEMPLATE = Path(__file__).resolve().with_name("painel_template.html")
MARCA = "/*__DADOS__*/[]"


def gerar(json_dados, destino_html):
    rows = json.loads(Path(json_dados).read_text(encoding="utf-8"))
    # "</" escapado para o texto da base nunca fechar a tag <script> por engano.
    dados = json.dumps(rows, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    modelo = TEMPLATE.read_text(encoding="utf-8")
    if modelo.count(MARCA) != 1:
        raise RuntimeError("painel_template.html sem o marcador da base")
    Path(destino_html).write_text(modelo.replace(MARCA, dados), encoding="utf-8")


if __name__ == "__main__":
    gerar(*sys.argv[1:3])
