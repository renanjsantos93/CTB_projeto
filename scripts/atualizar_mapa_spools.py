"""Atualiza as etapas de fabricação do Mapa de Spools a partir do Mapa de Juntas.

Regra (procedimento "Atualização Automática do Mapa de Spools"):
 1. Só entram na análise as juntas Pipe Shop (P/C = P / PIPE / PIPE SHOP).
    Juntas de Campo (P/C = C / CAMPO) e sem P/C são desconsideradas.
 2. Uma etapa do spool só é preenchida quando TODAS as juntas Pipe Shop do
    spool têm data na etapa correspondente do Mapa de Juntas.
 3. A data gravada é a maior (mais recente) entre essas juntas.
 4. Se pelo menos uma junta Pipe Shop está sem data, o campo do spool fica
    em branco (valores antigos são apagados, salvo com --nao-limpar).

Correspondência de etapas (Mapa de Juntas -> Mapa de Spools do ControlTub):
    Visual de Ajuste (VA)          -> Data Corte
    Ajuste (Visual de Ajuste)      -> Data VA Fab
    Soldagem                       -> Data Solda Fab
    EVS (Ensaio Visual)            -> Data EV Fab
    END (LP/PM, RX/US)             -> Data END Fab
    Rastreabilidade - dimensional  -> Data DF Fab
O Mapa de Juntas do ControlTub tem só a etapa "Visual Ajuste" (sem "Ajuste"
separado), por isso Data Corte e Data VA Fab vêm da mesma data. As colunas de
Montagem (VA Mon, Solda Mon...) nunca são alteradas. Etapas sem coluna em um
dos mapas são ignoradas e avisadas no relatório.

END: a data da junta é a mais recente entre os ensaios (LP, PM, RX, US...).
Ensaios com status "N" (não aplicável) são ignorados; um ensaio com status
"P" (pendente) e sem data deixa a junta sem data de END.

A associação junta -> spool é feita por Linha (com zona, ex.: -Z02) + nº do
spool. Se a linha do Mapa de Spools não tiver zona, usa-se a TAG padronizada
(4 blocos) quando ela identifica uma única linha do Mapa de Juntas.

Espessura: a coluna "Espessura" do spool recebe a maior espessura encontrada
entre todas as juntas da mesma linha no Mapa de Juntas (Pipe e Campo).

Uso:
    python scripts/atualizar_mapa_spools.py MAPA_JUNTAS.xlsx MAPA_SPOOLS.xlsx SAIDA.xlsx \\
        [--aba-juntas "SGJ (2)"] [--aba-spools "Mapa de Spools"] \\
        [--criar-colunas] [--nao-limpar] [--relatorio RELATORIO.csv]

Só as células das colunas de data do Mapa de Spools são alteradas; o restante
do arquivo (logos, fórmulas, formatação) é preservado.
"""
import argparse
import csv
import datetime as dt
import re
import sys
import unicodedata
import zipfile
from collections import OrderedDict, defaultdict

import openpyxl

from padronizar_tag import (_ampliar_aba, _caminho_aba, _col_letra, _estilo, _gravar_celula,
                            _limpar, padronizar_tag)


def _norm(v):
    s = unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode()
    s = s.upper().replace("º", "O").replace("°", "O")
    return re.sub(r"\s+", " ", s).strip()


# --------------------------------------------------------------------------
# Etapas
# --------------------------------------------------------------------------

# Etapas no Mapa de Juntas: (chave, nome, cabeçalhos de grupo)
ETAPAS = [
    ("VA", "Visual de Ajuste",
     {"VISUAL AJUSTE", "VISUAL DE AJUSTE", "VA"}),
    ("SOLDAGEM", "Soldagem",
     {"SOLDAGEM", "SOLDA"}),
    ("EVS", "EVS",
     {"ENSAIO VISUAL", "EVS", "ENSAIO VISUAL DE SOLDA", "INSPECAO VISUAL DE SOLDA"}),
    ("END", "END",
     {"LIQUIDO PENETRANTE / PM", "LIQUIDO PENETRANTE", "LP", "PM", "LP/PM", "PARTICULA MAGNETICA",
      "RX/US", "RX", "US", "ULTRASSOM", "RADIOGRAFIA", "END"}),
    ("DIMENSIONAL", "Dimensional",
     {"RASTREABILIDADE - DIMENSIONAL", "RASTREABILIDADE DIMENSIONAL", "RASTREABILIDADE / DIMENSIONAL",
      "DIMENSIONAL", "CONTROLE DIMENSIONAL", "INSPECAO DIMENSIONAL", "DIMENSIONAL FINAL"}),
]

# Colunas do Mapa de Spools: (chave, título ao criar, etapa de origem, cabeçalhos aceitos)
COLUNAS = [
    ("CORTE", "Data Corte", "VA",
     ["DATA CORTE", "DATA DE CORTE"]),
    ("VA", "Data VA Fab", "VA",
     ["DATA VA FAB", "DATA DE AJUSTE", "DATA AJUSTE", "DATA VA", "DATA DE VA",
      "DATA VISUAL DE AJUSTE", "DATA VISUAL AJUSTE"]),
    ("SOLDAGEM", "Data Solda Fab", "SOLDAGEM",
     ["DATA SOLDA FAB", "DATA DE SOLDAGEM", "DATA SOLDAGEM", "DATA SOLDA"]),
    ("EVS", "Data EV Fab", "EVS",
     ["DATA EV FAB", "DATA EVS FAB", "DATA DE EVS", "DATA EVS", "DATA EV"]),
    ("END", "Data END Fab", "END",
     ["DATA END FAB", "DATA DE END", "DATA END"]),
    ("DIMENSIONAL", "Data DF Fab", "DIMENSIONAL",
     ["DATA DF FAB", "DATA DIMENSIONAL", "DATA DE DIMENSIONAL", "DATA DF"]),
]

PIPE = {"P", "PIPE", "PIPE SHOP", "PIPESHOP", "OFICINA"}
CAMPO = {"C", "CAMPO", "CAMP", "FIELD"}

CAB_LINHA = ["LINHA", "NO DA LINHA", "N DA LINHA", "N. DA LINHA", "LINE", "LINE NUMBER", "TAG LINHA",
             "TAG DA LINHA", "TAG", "ISOMETRICO"]
CAB_SPOOL = ["SPOOL", "NO SPOOL", "NO DO SPOOL", "N SPOOL", "N DO SPOOL", "SPOOL NO"]


# --------------------------------------------------------------------------
# Datas
# --------------------------------------------------------------------------

def ler_data(v):
    """Retorna (date|None, preenchido). preenchido=True com date=None indica valor inválido."""
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None, False
    if isinstance(v, dt.datetime):
        return v.date(), True
    if isinstance(v, dt.date):
        return v, True
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 < v < 80000:  # nº serial do Excel
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date(), True
        return None, True
    m = re.fullmatch(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", str(v).strip())
    if not m:
        return None, True
    d, mes, a = (int(x) for x in m.groups())
    if a < 1000:  # 26 ou 026 (erro de digitação) -> 2026
        a = 2000 + a % 100
    try:
        return dt.date(a, mes, d), True
    except ValueError:
        return None, True


def ler_numero(v):
    """Número de célula que pode vir como texto com vírgula (' 4,55')."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    m = re.fullmatch(r"\s*(\d+(?:[.,]\d+)?)\s*(?:MM)?\s*", str(v or "").upper())
    return float(m.group(1).replace(",", ".")) if m else None


def _serial(d):
    return (d - dt.date(1899, 12, 30)).days


# --------------------------------------------------------------------------
# Leitura do Mapa de Juntas
# --------------------------------------------------------------------------

def _chave_linha(v):
    return _limpar(v).replace('"', "") if v not in (None, "") else ""


def _chave_tag(v):
    return padronizar_tag(v)[0].replace('"', "") if v not in (None, "") else ""


def _spool(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    s = str(v).strip().upper()
    return s.zfill(3) if s.isdigit() else s


def _junta(v):
    return _spool(v)


def localizar_juntas(ws, max_linhas=30):
    """Linha de cabeçalho (Linha/Spool/Junta/P/C) e colunas de data por etapa."""
    linhas = list(ws.iter_rows(min_row=1, max_row=max_linhas, values_only=True))
    for i, row in enumerate(linhas):
        hs = [_norm(c) for c in row]
        if "SPOOL" in hs and "P/C" in hs and "JUNTA" in hs:
            break
    else:
        raise SystemExit("Cabeçalho do Mapa de Juntas (Linha/Spool/Junta/P/C) não encontrado em %r" % ws.title)
    hdr = i + 1
    hs = [_norm(c) for c in linhas[i]]
    sub = [_norm(c) for c in linhas[i + 1]] if i + 1 < len(linhas) else []
    col = {}
    col["linha"] = hs.index("LINHA") + 1 if "LINHA" in hs else None
    col["iso"] = hs.index("ISOMETRICO") + 1 if "ISOMETRICO" in hs else None
    col["doc"] = hs.index("DOCUMENTO") + 1 if "DOCUMENTO" in hs else None
    col["esp"] = hs.index("ESPESSURA") + 1 if "ESPESSURA" in hs else None
    if not col["linha"] and not col["iso"]:
        raise SystemExit("Coluna 'Linha' não encontrada no Mapa de Juntas")
    col["spool"] = hs.index("SPOOL") + 1
    col["junta"] = hs.index("JUNTA") + 1
    col["pc"] = hs.index("P/C") + 1

    # grupos: cabeçalho na linha hdr, subcolunas (Status/Data...) na linha seguinte
    inicios = [(j, h) for j, h in enumerate(hs) if h]
    etapas = defaultdict(list)  # chave -> [(nome_grupo, col_status|None, [cols_data])]
    for n, (j, h) in enumerate(inicios):
        fim = inicios[n + 1][0] if n + 1 < len(inicios) else max(len(hs), len(sub))
        datas = [k + 1 for k in range(j, fim) if k < len(sub) and sub[k].startswith("DATA")]
        status = next((k + 1 for k in range(j, fim) if k < len(sub) and sub[k] == "STATUS"), None)
        if not datas:
            continue
        for chave, _nome, grupos in ETAPAS:
            if h in grupos:
                etapas[chave].append((h, status, datas))
    tem_sub = any(s for s in sub)
    return hdr + (2 if tem_sub else 1), col, dict(etapas)


def ler_juntas(caminho, aba=None):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    abas = [wb[aba]] if aba else [ws for ws in wb.worksheets if _tem_juntas(ws)]
    juntas, etapas_encontradas, avisos = [], set(), []
    vistas, duplicadas = set(), 0
    for ws in abas:
        primeira, col, etapas = localizar_juntas(ws)
        etapas_encontradas |= set(etapas)
        for r, row in enumerate(ws.iter_rows(min_row=primeira, values_only=True), primeira):
            def v(c):
                return row[c - 1] if c and c - 1 < len(row) else None
            linha_bruta = v(col["linha"]) or v(col["iso"])
            spool = _spool(v(col["spool"]))
            # linha de numeração das colunas (1, 2, 3...) logo abaixo do cabeçalho
            if not isinstance(linha_bruta, str) or not spool:
                continue
            doc = _chave_linha(v(col["doc"]))
            ident = (_chave_linha(linha_bruta), doc, spool, _junta(v(col["junta"])))
            if ident in vistas:  # junta repetida no mapa (ou em outra aba)
                duplicadas += 1
                continue
            vistas.add(ident)
            pc = _norm(v(col["pc"]))
            tipo = "PIPE" if pc in PIPE else "CAMPO" if pc in CAMPO else ""
            if not tipo and pc:
                avisos.append("%s!%d: P/C desconhecido %r (junta desconsiderada)" % (ws.title, r, v(col["pc"])))
            datas = {}
            for chave, grupos in etapas.items():
                datas[chave] = _data_etapa(chave, grupos, v)
            juntas.append({
                "aba": ws.title, "linha_xlsx": r, "linha": str(linha_bruta).strip(),
                "chave": _chave_linha(linha_bruta), "tag": _chave_tag(linha_bruta), "doc": doc,
                "spool": spool, "junta": _junta(v(col["junta"])), "tipo": tipo, "datas": datas,
                "espessura": ler_numero(v(col["esp"])),
            })
    if duplicadas:
        avisos.append("%d juntas repetidas no Mapa de Juntas (contadas uma vez)" % duplicadas)
    return juntas, etapas_encontradas, avisos


def _tem_juntas(ws):
    for row in ws.iter_rows(min_row=1, max_row=30, values_only=True):
        hs = [_norm(c) for c in row]
        if "SPOOL" in hs and "P/C" in hs and "JUNTA" in hs:
            return True
    return False


def _data_etapa(chave, grupos, v):
    """(date|None, motivo) da junta na etapa. motivo explica a ausência de data."""
    if chave != "END":
        datas, invalidas = [], []
        for _h, _st, cols in grupos:
            for c in cols:
                d, preenchido = ler_data(v(c))
                if d:
                    datas.append(d)
                elif preenchido:
                    invalidas.append(str(v(c)))
        if datas:
            return max(datas), ""
        return None, ("data inválida %r" % invalidas[0]) if invalidas else "sem data"
    # END: vários ensaios; "N" = não aplicável, "P" sem data = pendente
    datas, pendentes = [], []
    for h, st, cols in grupos:
        status = _norm(v(st)) if st else ""
        if status == "N":
            continue
        ds = [d for d in (ler_data(v(c))[0] for c in cols) if d]
        if ds:
            datas.extend(ds)
        elif status == "P":
            pendentes.append(h)
    if pendentes:
        return None, "END pendente (%s)" % ", ".join(pendentes)
    if datas:
        return max(datas), ""
    return None, "sem data"


# --------------------------------------------------------------------------
# Consolidação por spool
# --------------------------------------------------------------------------

def consolidar(juntas, etapas):
    """{(chave_linha, documento, spool): {"juntas": [...], "etapas": {chave: (date|None, motivo)}}}"""
    spools = OrderedDict()
    for j in juntas:
        s = spools.setdefault((j["chave"], j["doc"], j["spool"]), {"linha": j["linha"], "tag": j["tag"],
                                                                   "chave": j["chave"],
                                                         "juntas": [], "campo": []})
        (s["juntas"] if j["tipo"] == "PIPE" else s["campo"]).append(j)
    for s in spools.values():
        s["etapas"] = {}
        for chave in etapas:
            if not s["juntas"]:
                s["etapas"][chave] = (None, "spool sem juntas Pipe Shop")
                continue
            faltando = [j for j in s["juntas"] if not j["datas"][chave][0]]
            if faltando:
                s["etapas"][chave] = (None, "%d de %d juntas Pipe sem data: %s" % (
                    len(faltando), len(s["juntas"]),
                    ", ".join("%s (%s)" % (j["junta"], j["datas"][chave][1]) for j in faltando[:10])
                    + (" ..." if len(faltando) > 10 else "")))
            else:
                s["etapas"][chave] = (max(j["datas"][chave][0] for j in s["juntas"]), "")
    return spools


# --------------------------------------------------------------------------
# Mapa de Spools
# --------------------------------------------------------------------------

def localizar_spools(ws, max_linhas=30):
    for row in ws.iter_rows(min_row=1, max_row=max_linhas):
        hs = {c.column: _norm(c.value) for c in row if c.value not in (None, "")}
        c_spool = next((c for c, h in hs.items() if h in CAB_SPOOL), None)
        c_linha = None
        for alvo in CAB_LINHA:  # ordem de preferência
            c_linha = next((c for c, h in hs.items() if h == alvo), None)
            if c_linha:
                break
        if c_spool and c_linha:
            destinos = {}
            for chave, _titulo, _origem, cabs in COLUNAS:
                for alvo in cabs:
                    c = next((c for c, h in hs.items() if h == alvo and c not in destinos.values()), None)
                    if c:
                        destinos[chave] = c
                        break
            extras = {nome: next((c for c, h in hs.items() if h == alvo), None) for nome, alvo in
                      (("doc", "DOCUMENTO"), ("total", "TOTAL_JUNTAS"), ("esp", "ESPESSURA"))}
            return row[0].row, c_linha, c_spool, destinos, max(hs), extras
    raise SystemExit("Cabeçalho do Mapa de Spools (Linha + Spool) não encontrado em %r" % ws.title)


def _cellxfs(styles):
    m = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S)
    return m, re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", m.group(1), re.S)


def _numfmt_data(styles, codigo="dd/mm/yyyy"):
    """Id do formato de número dd/mm/yyyy (criado em <numFmts> se não existir)."""
    m = re.search(r'<numFmt numFmtId="(\d+)" formatCode="%s"' % re.escape(codigo), styles)
    if m:
        return int(m.group(1)), styles
    ids = [int(i) for i in re.findall(r'<numFmt numFmtId="(\d+)"', styles)]
    novo_id = max(ids + [163]) + 1
    tag = '<numFmt numFmtId="%d" formatCode="%s"/>' % (novo_id, codigo)
    novo = '<numFmts count="%d">' % (len(ids) + 1)
    if re.search(r"<numFmts\b[^>]*/>", styles):
        styles = re.sub(r"<numFmts\b[^>]*/>", novo + tag + "</numFmts>", styles, count=1)
    elif "<numFmts" in styles:
        styles = re.sub(r"<numFmts\b[^>]*>", novo, styles, count=1)
        styles = styles.replace("</numFmts>", tag + "</numFmts>", 1)
    else:
        styles = re.sub(r"(<styleSheet\b[^>]*>)", r'\1<numFmts count="1">%s</numFmts>' % tag, styles, count=1)
    return novo_id, styles


def _estilo_data(styles, base, cache):
    """Índice de um estilo igual a `base` com formato de data (dd/mm/aaaa)."""
    base = int(base or 0)
    if base in cache:
        return cache[base], styles
    m, xfs = _cellxfs(styles)
    xf = xfs[base] if base < len(xfs) else xfs[0]
    fmt = re.search(r'numFmtId="(\d+)"', xf)
    fmt_id = int(fmt.group(1)) if fmt else 0
    nfmt = re.search(r'<numFmt numFmtId="%d" formatCode="([^"]*)"' % fmt_id, styles)
    if 14 <= fmt_id <= 22 or (nfmt and re.search(r"[dy]", nfmt.group(1).lower()) and "h" not in nfmt.group(1).lower()):
        cache[base] = base
        return base, styles
    id_data, styles = _numfmt_data(styles)
    m, xfs = _cellxfs(styles)
    novo = re.sub(r'\snumFmtId="\d+"', "", xf)
    novo = re.sub(r'\sapplyNumberFormat="\d+"', "", novo)
    novo = re.sub(r"^<xf\b", '<xf numFmtId="%d" applyNumberFormat="1"' % id_data, novo)
    if novo in xfs:
        cache[base] = xfs.index(novo)
        return cache[base], styles
    corpo = m.group(1) + novo
    ini = styles[:m.start()] + re.sub(r'count="\d+"', 'count="%d"' % (len(xfs) + 1),
                                      styles[m.start():m.start(1)], count=1)
    styles = ini + corpo + styles[m.end(1):]
    cache[base] = len(xfs)
    return len(xfs), styles


def atualizar_spools(juntas_xlsx, spools_xlsx, saida, aba_juntas=None, aba_spools=None,
                     criar_colunas=False, limpar=True):
    juntas, etapas_juntas, avisos = ler_juntas(juntas_xlsx, aba_juntas)
    spools = consolidar(juntas, etapas_juntas)
    por_linha, por_tag = defaultdict(list), defaultdict(list)
    for k, s in spools.items():
        por_linha[(k[0], k[2])].append(k)
        por_tag[(s["tag"], k[2])].append(k)

    wb = openpyxl.load_workbook(spools_xlsx)
    ws = wb[aba_spools] if aba_spools else wb.worksheets[0]
    hdr, c_linha, c_spool, destinos, ultima, extras = localizar_spools(ws)

    nomes = {k: n for k, n, _g in ETAPAS}
    for chave, nome, _g in ETAPAS:
        if chave not in etapas_juntas:
            avisos.append("Etapa %s: coluna de data não encontrada no Mapa de Juntas (ignorada)" % nome)
    novas = []
    for chave, titulo, origem, _cabs in COLUNAS:
        if origem not in etapas_juntas:
            destinos.pop(chave, None)
        elif chave not in destinos:
            if criar_colunas:
                ultima += 1
                destinos[chave] = ultima
                novas.append((ultima, titulo))
            else:
                avisos.append("Etapa %s: coluna %r não encontrada no Mapa de Spools (use --criar-colunas)"
                              % (nomes[origem], titulo))
    origem_col = {k: o for k, _t, o, _c in COLUNAS}

    esp_linha = {}
    for j in juntas:
        if j["espessura"] is not None:
            esp_linha[j["chave"]] = max(j["espessura"], esp_linha.get(j["chave"], 0))

    gravar, gravar_esp, resultado, sem_juntas, divergencias = {}, {}, [], [], []
    for r in range(hdr + 1, ws.max_row + 1):
        linha, sp = ws.cell(r, c_linha).value, _spool(ws.cell(r, c_spool).value)
        if not linha or not sp:
            continue
        # Linha + Documento + Spool; sem Documento correspondente, Linha + Spool
        # (ou TAG padronizada + Spool) desde que aponte para um único spool
        doc = _chave_linha(ws.cell(r, extras["doc"]).value) if extras["doc"] else None
        s = spools.get((_chave_linha(linha), doc, sp))
        motivo_sem = "spool não encontrado no Mapa de Juntas"
        if s is None:
            for cand in (por_linha.get((_chave_linha(linha), sp), []), por_tag.get((_chave_tag(linha), sp), [])):
                if len(cand) == 1:
                    s = spools[cand[0]]
                    break
                if len(cand) > 1:
                    motivo_sem = "spool repetido em %d documentos do Mapa de Juntas; Documento não confere" % len(cand)
                    break
        if s is None:
            sem_juntas.append((r, linha, sp, motivo_sem))
            continue
        # conferência: total de juntas do spool nos dois mapas
        c = extras["total"]
        esperado = len(s["juntas"]) + len(s["campo"])
        if c and isinstance(ws.cell(r, c).value, (int, float)) and ws.cell(r, c).value != esperado:
            divergencias.append((r, linha, sp, ws.cell(hdr, c).value, ws.cell(r, c).value, esperado))
        for chave, c in destinos.items():
            data, _ = s["etapas"][origem_col[chave]]
            atual, _ = ler_data(ws.cell(r, c).value)
            ref = "%s%d" % (_col_letra(c), r)
            if data and data != atual:
                gravar[ref] = data
            elif not data and ws.cell(r, c).value not in (None, "") and limpar:
                gravar[ref] = None
        # espessura do spool = maior espessura das juntas da linha
        c = extras["esp"]
        esp = esp_linha.get(s["chave"])
        if c and esp is not None and ler_numero(ws.cell(r, c).value) != esp:
            gravar_esp["%s%d" % (_col_letra(c), r)] = esp
        for chave in (k for k, *_ in ETAPAS if k in etapas_juntas):
            data, motivo = s["etapas"][chave]
            resultado.append((r, linha, sp, chave, data, motivo, len(s["juntas"]), len(s["campo"])))

    with zipfile.ZipFile(spools_xlsx) as zin:
        caminho = _caminho_aba(zin, ws.title)
        xml = zin.read(caminho).decode("utf-8")
        styles = zin.read("xl/styles.xml").decode("utf-8")
        cache = {}
        for c, titulo in novas:
            ant = _col_letra(c - 1)
            xml = _gravar_celula(xml, "%s%d" % (_col_letra(c), hdr), titulo, _estilo(xml, "%s%d" % (ant, hdr)))
            xml = _ampliar_aba(xml, c)
        for ref, data in gravar.items():
            if data is None:
                xml = _gravar_celula(xml, ref, None)
                continue
            base = _estilo(xml, ref)
            if base is None:  # célula nova: herda o estilo da coluna anterior na mesma linha
                lin = re.match(r"[A-Z]+(\d+)", ref).group(1)
                col = re.match(r"[A-Z]+", ref).group(0)
                n = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(col)))
                base = _estilo(xml, "%s%s" % (_col_letra(n - 1), lin)) if n > 1 else None
            estilo, styles = _estilo_data(styles, base, cache)
            xml = _gravar_celula(xml, ref, None, estilo, numero=_serial(data), forcar_estilo=True)
        for ref, esp in gravar_esp.items():
            xml = _gravar_celula(xml, ref, None, numero=esp)
        alterados = {caminho: xml, "xl/styles.xml": styles}
        with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                dados = alterados[item.filename].encode("utf-8") if item.filename in alterados \
                    else zin.read(item.filename)
                zout.writestr(item, dados)

    return {"aba": ws.title, "destinos": destinos, "gravados": gravar, "espessuras": gravar_esp,
            "col_espessura": extras["esp"], "resultado": resultado,
            "sem_juntas": sem_juntas, "divergencias": divergencias, "avisos": avisos, "spools": spools, "juntas": juntas}


def salvar_relatorio(caminho, res):
    nomes = {k: n for k, n, _g in ETAPAS}
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Linha planilha", "Linha", "Spool", "Etapa", "Data consolidada", "Juntas Pipe",
                    "Juntas Campo (desconsideradas)", "Pendência"])
        for r, linha, sp, chave, data, motivo, n_pipe, n_campo in res["resultado"]:
            w.writerow([r, linha, sp, nomes[chave], data.strftime("%d/%m/%Y") if data else "",
                        n_pipe, n_campo, motivo])
        for r, linha, sp, motivo in res["sem_juntas"]:
            w.writerow([r, linha, sp, "", "", 0, 0, motivo])
        for r, linha, sp, coluna, valor, esperado in res["divergencias"]:
            w.writerow([r, linha, sp, "", "", "", "", "%s = %s no Mapa de Spools, %s no Mapa de Juntas"
                        % (coluna, valor, esperado)])


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mapa_juntas")
    ap.add_argument("mapa_spools")
    ap.add_argument("saida")
    ap.add_argument("--aba-juntas", help="aba do Mapa de Juntas (padrão: todas as abas com juntas)")
    ap.add_argument("--aba-spools", help="aba do Mapa de Spools (padrão: primeira)")
    ap.add_argument("--criar-colunas", action="store_true",
                    help="cria no Mapa de Spools as colunas de data que não existirem")
    ap.add_argument("--nao-limpar", action="store_true",
                    help="não apaga datas já existentes de etapas que ficaram incompletas")
    ap.add_argument("--relatorio", help="CSV com a situação de cada spool/etapa")
    a = ap.parse_args(argv)

    res = atualizar_spools(a.mapa_juntas, a.mapa_spools, a.saida, a.aba_juntas, a.aba_spools,
                           a.criar_colunas, not a.nao_limpar)
    nomes = {k: n for k, n, _g in ETAPAS}
    titulos = {k: (t, o) for k, t, o, _c in COLUNAS}
    n_pipe = sum(1 for j in res["juntas"] if j["tipo"] == "PIPE")
    print("Mapa de Juntas: %d juntas (%d Pipe Shop), %d spools" % (len(res["juntas"]), n_pipe, len(res["spools"])))
    preench = defaultdict(int)
    for _r, _l, _s, chave, data, *_ in res["resultado"]:
        if data:
            preench[chave] += 1
    print("Aba %s:" % res["aba"])
    for k, c in res["destinos"].items():
        t, o = titulos[k]
        print("  %-15s (col. %-2s) <- %-16s %d spools concluídos" % (t, _col_letra(c), nomes[o], preench[o]))
    if not res["destinos"]:
        print("  nenhuma coluna atualizada")
    if res["col_espessura"]:
        print("  %-15s (col. %-2s) <- maior espessura da linha: %d spools atualizados"
              % ("Espessura", _col_letra(res["col_espessura"]), len(res["espessuras"])))
    print("%d células alteradas" % (len(res["gravados"]) + len(res["espessuras"])))
    for r, linha, sp, motivo in res["sem_juntas"]:
        print("  linha %d: spool %s / %s: %s" % (r, sp, linha, motivo))
    if res["divergencias"]:
        print("  AVISO: %d divergências de quantidade de juntas entre os mapas (ver relatório)"
              % len(res["divergencias"]))
    for av in res["avisos"]:
        print("  AVISO:", av)
    if a.relatorio:
        salvar_relatorio(a.relatorio, res)
        print("Relatório: %s" % a.relatorio)
    return 0


if __name__ == "__main__":
    sys.exit(main())
