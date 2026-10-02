"""Converte o número da junta para texto com 3 dígitos (1 -> "001").

Uso:
    python scripts/junta_texto.py ENTRADA.xlsx SAIDA.xlsx --aba "MAPA DE JUNTAS" \
        --colunas "Nº JUNTA" "Nº JUNTA2" [--digitos 3]

As colunas são localizadas pelo cabeçalho. Só as células numéricas inteiras
dessas colunas são alteradas (gravadas como texto); o restante do arquivo é
preservado byte a byte.
"""
import argparse
import sys
import zipfile

import openpyxl

from padronizar_tag import _caminho_aba, _col_letra, _gravar_celula, _norm_header


def processar(entrada, saida, aba, colunas, digitos=3):
    wb = openpyxl.load_workbook(entrada)
    ws = wb[aba] if aba else wb.worksheets[0]
    alvos = {_norm_header(c) for c in colunas}
    encontrados, linha_hdr = {}, None
    for row in ws.iter_rows(min_row=1, max_row=30):
        for c in row:
            h = _norm_header(c.value)
            if h in alvos and h not in encontrados:
                encontrados[h] = c.column
                linha_hdr = c.row
    faltando = alvos - set(encontrados)
    if faltando:
        raise SystemExit("Colunas não encontradas: %s" % ", ".join(sorted(faltando)))

    novos, ignorados = {}, []
    for col in encontrados.values():
        for r in range(linha_hdr + 1, ws.max_row + 1):
            v = ws.cell(r, col).value
            if v is None or v == "":
                continue
            ref = "%s%d" % (_col_letra(col), r)
            if isinstance(v, bool) or not isinstance(v, (int, float)) or v != int(v) or v < 0:
                if not (isinstance(v, str) and v.isdigit() and len(v) >= digitos):
                    ignorados.append((ref, v))
                if isinstance(v, str) and v.strip().isdigit():
                    novos[ref] = v.strip().zfill(digitos)
                continue
            novos[ref] = str(int(v)).zfill(digitos)

    with zipfile.ZipFile(entrada) as zin:
        caminho = _caminho_aba(zin, ws.title)
        xml = zin.read(caminho).decode("utf-8")
        for ref, texto in novos.items():
            xml = _gravar_celula(xml, ref, texto)
        with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                dados = xml.encode("utf-8") if item.filename == caminho else zin.read(item.filename)
                zout.writestr(item, dados)
    return encontrados, novos, ignorados


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("entrada")
    ap.add_argument("saida")
    ap.add_argument("--aba")
    ap.add_argument("--colunas", nargs="+", default=["Nº JUNTA"])
    ap.add_argument("--digitos", type=int, default=3)
    a = ap.parse_args(argv)
    cols, novos, ignorados = processar(a.entrada, a.saida, a.aba, a.colunas, a.digitos)
    print("Colunas: %s | %d células convertidas para texto" % (
        ", ".join("%s (%s)" % (h, _col_letra(c)) for h, c in cols.items()), len(novos)))
    for ref, v in ignorados:
        print("  não convertida %s: %r" % (ref, v))
    return 0


if __name__ == "__main__":
    sys.exit(main())
