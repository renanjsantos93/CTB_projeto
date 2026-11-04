"""Gera o Painel_OLODUM.html: o painel completo num arquivo só, com as bases embutidas.

Abre em qualquer navegador, sem internet e sem conector; quem tem acesso à pasta do OneDrive vê o painel.

Uso: python gerar_painel_html.py <pcp_painel.json> <Painel_OLODUM.html> [<pendencias.json>] [<lista_pendencias.json>]
"""
import json
import sys
from pathlib import Path

TEMPLATE = Path(__file__).resolve().with_name("painel_template.html")


def embutir(modelo, marca, json_path):
    rows = json.loads(Path(json_path).read_text(encoding="utf-8")) if json_path else []
    # "</" escapado para o texto da base nunca fechar a tag <script> por engano.
    dados = json.dumps(rows, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    if modelo.count(marca) != 1:
        raise RuntimeError("painel_template.html sem o marcador %s" % marca)
    return modelo.replace(marca, dados)


def gerar(json_dados, destino_html, json_pendencias=None, json_lista=None):
    modelo = TEMPLATE.read_text(encoding="utf-8")
    modelo = embutir(modelo, "/*__DADOS__*/[]", json_dados)
    modelo = embutir(modelo, "/*__PEND__*/[]", json_pendencias)
    modelo = embutir(modelo, "/*__PENDLISTA__*/[]", json_lista)
    Path(destino_html).write_text(modelo, encoding="utf-8")


if __name__ == "__main__":
    gerar(*sys.argv[1:5])
