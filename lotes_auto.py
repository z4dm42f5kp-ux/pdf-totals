"""lotes_auto.py - piezas para destruccion por numero de lote.

Lee todos los PDFs de una carpeta (devoluciones, acuses de entrega, recepciones de
producto, inventarios y reportes de retiro) y calcula por lote:

  1. Cantidad de producto retirado    piezas devueltas o recuperadas del mercado
  2. Cantidad no distribuida          piezas en inventario que nunca se distribuyeron
  3. Cantidad total para destruccion  1 + 2

Las cantidades no se buscan solo por palabras clave. Cada tabla se analiza por su
estructura:

  - de que tipo es cada celda: lote, fecha, importe, entero, codigo o texto;
  - que columna cumple  piezas x precio unitario = importe;
  - que columna suma exactamente el total impreso en el documento;
  - en los reportes de retiro, que columnas cumplen
        fabricado - distribuido = no distribuido   y
        no distribuido + retirado = total para destruccion.

Las palabras de los encabezados (PIEZAS, RECIBIDA, FACTURADA, INVENTARIO...) son
una senal mas, no la unica. Cada cantidad queda marcada como verificada o no, y al
final se comprueban los topes por lote y el total esperado. Si algo no cuadra, los
PDFs escaneados dudosos se leen otra vez con mas resolucion, y un PDF solo se marca
como MALO cuando la prueba es concluyente.
"""
import argparse
import os
import pathlib
import re
import sys
import unicodedata
from collections import Counter, defaultdict, namedtuple
from dataclasses import dataclass, field

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

VERSION = "2.0.0"

# ============================== CONFIGURACION ===============================
# Total esperado de la tabla "Cantidad total para destruccion" (None = sin comprobar).
TOTAL_ESPERADO = 344451
# Topes por lote. Un lote que los supera indica un error de lectura o un duplicado.
MAX_RETIRADO_POR_LOTE = 48042
MAX_NO_DISTRIBUIDA_POR_LOTE = 82090
# Los reportes de retiro (fabricado / distribuido / no distribuido / retirado / total)
# dan la cantidad no distribuida de cada lote y sirven de referencia para comprobar lo
# retirado segun las devoluciones. Con True, su retirado tambien se usa para los lotes
# que no tienen ninguna devolucion en la carpeta.
USAR_REPORTE_PARA_RETIRADO = False
# Lote repetido en varias paginas del mismo PDF:
#   "iguales": se cuenta una vez si la cantidad es la misma (copia del mismo documento);
#              con cantidades distintas se suman (documentos distintos) y se avisa.
#   "todos":   cualquier repeticion se cuenta una sola vez (se conserva la mayor).
DUPLICADOS_EN_ARCHIVO = "iguales"
# PDFs cuyo nombre contiene esto se omiten (no cuentan en ningun total).
OMITIR_NOMBRE = re.compile(r"falta[\s_\-]*revisi[oó]n", re.I)
# Numero de lote valido: 8 digitos que empiezan con 24, 25 o 26.
LOTE_ID = re.compile(r"(?<![\d.,])(2[456]\d{6})(?![\d]|[.,]\d)")
# OCR: resolucion de la primera lectura y de la segunda (solo PDFs dudosos).
OCR_DPI = 300
OCR_DPI_REINTENTO = 400
# Confianza minima del OCR para no considerar dudosa una cantidad sin verificar.
CONFIANZA_MINIMA = 0.90
# Formularios "de lado" (un producto por columna): renglon de cantidad preferido.
PRIORIDAD_RENGLON = ["devuel", "recibid", "retir", "recuper", "pieza", "cantidad", "factur"]

# Significado de los encabezados de columna (texto sin acentos y en minusculas).
SEMANTICA = {
    "ret": r"retir|recuper|devuel|devoluc|recib|recolect|regres",
    "nod": r"no\s*distribu|sin\s*distribu|inventari|existenci|stock|almacen|retenid|cuarentena|bloquead|saldo",
    "orig": r"factur|fabric|producid|produccion|liberad|vendid|surtid|pedid|solicitad|enviad|remitid"
            r"|(?<!no )(?<!sin )distribuid",
    "caja": r"caja",
    "cant": r"cantidad|\bcant\b|pieza|\bpzas?\b|\bpzs\b|unidades|\buds\b|frasco|ampollet|ampula|vial",
    "total": r"total|destruc",
}
SEMANTICA = {k: re.compile(v) for k, v in SEMANTICA.items()}

# Tipo de documento por pagina: (expresion, peso).
TIPO_DOCUMENTO = {
    "ret": [(r"devoluci", 3), (r"devuelt", 2), (r"retiro", 3), (r"retirad", 3), (r"recall", 3),
            (r"recuperad", 2), (r"recolecci", 2), (r"acuse de entrega", 2),
            (r"recepcion de producto", 2), (r"anomalia", 2), (r"motivo", 1), (r"faltante", 1),
            (r"maltratad", 1)],
    "nod": [(r"inventario", 3), (r"existencia", 3), (r"stock", 2), (r"no distribuid", 4),
            (r"sin distribuir", 4), (r"cuarentena", 2), (r"retenid", 2), (r"bloquead", 2)],
}
# ============================================================================

TIPOS = {"ret": "retirado", "nod": "no distribuida"}
MESES = r"(?:ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]*\.?"
FECHA_RE = re.compile(
    r"^(?:\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2}"
    r"|\d{1,2}[/\-. ]?" + MESES + r"[/\-. ]?\d{2,4}|" + MESES + r"[/\-. ]?\d{2,4}|\d{1,2}/\d{4})$", re.I)
DINERO_RE = re.compile(r"^(?:[$sS]\s*)?(\d{1,3}(?:,\d{3})+|\d+)\.\d{2}$|^[$sS]\s*(\d{1,3}(?:,\d{3})+|\d+)$")
UNIDAD_RE = re.compile(r"^(\d[\d.,]*)\s*(pzas?|piezas?|pzs|pz|uds?|unidades|frascos?|amp\w*|viales?|cajas?)\.?$", re.I)
MARCAS = "✓✔√/\\|*·~'\"`´¨°º"
LOTE_PREFIJO_OK = re.compile(r"(?:l[o0]te|lot|serie|no|num|n|l)$", re.I)
LOTE_PALABRA = re.compile(r"l[o0]te|serie|serial|n/s|s/n", re.I)

# Etiquetas en lineas de texto ("LOTE: 25240310" / "CANTIDAD EN PIEZAS: 1,250").
CANT_PALABRAS = r"cantidad|cant|piezas|pieza|pzas|pza|pzs|pz|unidades|unid|uds"
Q = r"(?:" + CANT_PALABRAS + r")(?![a-z])\.?"
CANT_ETIQUETA = re.compile(
    Q + r"(?:\s*(?:/|en|de)?\s*" + Q + r")*\s*(?:\(\w*\))?\s*[:=\-]?\s*(?P<n>\d[\d.,]*)", re.I)
LOTE_ETIQUETA = re.compile(r"(?:l[o0]te|serial|serie|n/s|s/n)\.?\s*(?:n[º°o]\.?|n[uú]m\.?|n[uú]mero)?"
                           r"\s*[:#\-]?\s*(?P<id>\d{8})", re.I)


# ------------------------------- utilidades --------------------------------
def sin_acentos(texto):
    t = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in t if not unicodedata.combining(c)).lower()


def lotes_en(texto):
    """Todos los lotes validos del texto, en orden."""
    res = []
    for m in LOTE_ID.finditer(texto or ""):
        previo = re.search(r"[A-Za-z]+$", texto[:m.start()])
        if previo and not LOTE_PREFIJO_OK.search(previo.group()):
            continue                      # p. ej. una factura FE25012345
        res.append(m.group(1))
    return res


def valid_lote(texto):
    """Primer lote valido del texto o None (se conserva el nombre por compatibilidad)."""
    lotes = lotes_en(texto)
    return lotes[0] if lotes else None


def a_entero(s):
    """Entero en notacion mexicana: 1,250 / 1.250 (coma leida como punto) / 1250."""
    s = (s or "").replace(" ", "")
    if re.fullmatch(r"\d{1,3}(?:[,.]\d{3})+", s):
        return int(re.sub(r"[,.]", "", s))
    if re.fullmatch(r"\d+", s):
        return int(s)
    return None


def a_dinero(s):
    s = re.sub(r"^[$sS]\s*", "", (s or "").strip()).replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def fmt(v):
    return f"{int(v):,}" if float(v).is_integer() else f"{v:,.2f}"


def orden_lote(s):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", str(s))]


def semantica(texto):
    t = sin_acentos(texto)
    return {k: bool(rx.search(t)) for k, rx in SEMANTICA.items()}


def tipo_documento(texto):
    """'ret', 'nod' o None segun el vocabulario de la pagina."""
    t = sin_acentos(texto)
    puntos = {k: sum(p for rx, p in reglas if re.search(rx, t)) for k, reglas in TIPO_DOCUMENTO.items()}
    if puntos["ret"] >= puntos["nod"] + 2:
        return "ret"
    if puntos["nod"] >= puntos["ret"] + 2:
        return "nod"
    return None


# -------------------------------- modelo -----------------------------------
Tok = namedtuple("Tok", "x0 x1 y h text conf caja")
Celda = namedtuple("Celda", "x0 x1 text conf tipo valor")
Item = namedtuple("Item", "fila lote celda_lote celdas")


class Renglon:
    def __init__(self, y, h):
        self.y, self.h, self.toks, self.texto = y, h, [], ""


@dataclass
class Entrada:
    archivo: str
    pagina: int
    lote: str
    piezas: int
    tipo: str                 # "ret" o "nod"
    forma: str                # "tabla", "de lado", "etiquetas"
    conf: float = 1.0
    verificado: list = field(default_factory=list)
    fallas: list = field(default_factory=list)
    notas: list = field(default_factory=list)
    excluido: str = ""

    def clave(self):
        return (self.tipo, self.lote, self.piezas)


@dataclass
class Referencia:
    archivo: str
    pagina: int
    lote: str
    ret: object = None
    nod: object = None
    total: object = None
    cuadra: bool = True
    ret_opciones: tuple = ()      # valores aceptables si la fila del reporte no cuadra

    def acepta_ret(self, valor):
        return valor == self.ret or valor in self.ret_opciones


@dataclass
class Lectura:
    archivo: str
    entradas: list
    referencias: list
    lineas: list
    fuente: str
    traza: list
    pasada: int = 1
    error: str = ""
    avisos: list = field(default_factory=list)


class OCRNoDisponible(Exception):
    """El PDF es escaneado y el OCR no esta instalado o no carga en este sistema."""


# ------------------------------- lectura -----------------------------------
_ocr = None


def load_ocr():
    """rapidocr_onnxruntime (Python <= 3.12) o rapidocr 3.x + onnxruntime. Devuelve
    una funcion imagen -> [(caja, texto, confianza)]."""
    errores = []
    try:
        from rapidocr_onnxruntime import RapidOCR
        motor = RapidOCR()
        return lambda arr: (motor(arr)[0] or [])
    except Exception as e:              # ImportError, o "DLL load failed" en Windows
        errores.append(f"rapidocr_onnxruntime: {e}")
    try:
        import logging
        from rapidocr import RapidOCR
        logging.getLogger("RapidOCR").setLevel(logging.WARNING)
        motor = RapidOCR()

        def leer(arr):
            r = motor(arr)
            return [] if r.boxes is None else list(zip(r.boxes.tolist(), r.txts, r.scores))
        return leer
    except Exception as e:
        errores.append(f"rapidocr: {e}")
    raise OCRNoDisponible("; ".join(errores))


def page_tokens(page, pasada=1):
    """Palabras de la pagina con posicion y confianza; fuente 'texto' u 'OCR'."""
    if len(page.chars) >= 20 or (page.chars and not page.images):
        words = page.extract_words(x_tolerance=1.5, y_tolerance=2)
        return [Tok(w["x0"], w["x1"], (w["top"] + w["bottom"]) / 2, w["bottom"] - w["top"],
                    w["text"], 1.0, i) for i, w in enumerate(words)], "texto"
    global _ocr
    if _ocr is None:
        _ocr = load_ocr()
    import numpy as np
    try:
        img = page.to_image(resolution=OCR_DPI if pasada == 1 else OCR_DPI_REINTENTO).original
    except ImportError as e:            # sin pypdfium2 (p. ej. BSD) no se puede rasterizar
        raise OCRNoDisponible(str(e)) from e
    img = img.convert("RGB")
    if pasada > 1:                      # segunda lectura: mas contraste y nitidez
        from PIL import ImageFilter, ImageOps
        img = ImageOps.autocontrast(img.convert("L"), cutoff=1).filter(ImageFilter.SHARPEN).convert("RGB")
    resultado = list(_ocr(np.array(img)))
    # Inclinacion del escaneo: angulo mediano del borde superior de las cajas de texto.
    import math
    angulos = []
    for box, text, _ in resultado:
        (ax, ay), (bx, by) = box[0], box[1]
        if bx - ax > 3 * abs(box[3][1] - box[0][1]) and len(text) >= 4:
            angulos.append(math.atan2(by - ay, bx - ax))
    ang = sorted(angulos)[len(angulos) // 2] if len(angulos) >= 3 else 0.0
    cos, sen = math.cos(-ang), math.sin(-ang)
    toks = []
    for n, (box, text, score) in enumerate(resultado):
        box = [(x * cos - y * sen, x * sen + y * cos) for x, y in box]   # enderezar
        xs, ys = [p[0] for p in box], [p[1] for p in box]
        x0, x1, h = min(xs), max(xs), max(ys) - min(ys)
        y = (min(ys) + max(ys)) / 2
        ancho = max(len(text), 1)
        for m in re.finditer(r"\S+", text):   # una caja de OCR -> palabras
            toks.append(Tok(x0 + (x1 - x0) * m.start() / ancho, x0 + (x1 - x0) * m.end() / ancho,
                            y, h, m.group(), float(score), n))
    return toks, "OCR"


def agrupar_renglones(toks):
    if not toks:
        return []
    hs = sorted(t.h for t in toks)
    hmed = hs[len(hs) // 2]
    renglones = []
    for t in sorted(toks, key=lambda t: t.y):
        r = renglones[-1] if renglones else None
        if r and abs(t.y - r.y) < 0.6 * max(hmed, min(t.h, r.h)):
            r.toks.append(t)
            ys = sorted(x.y for x in r.toks)
            r.y = ys[len(ys) // 2]
            r.h = sorted(x.h for x in r.toks)[len(r.toks) // 2]
        else:
            r = Renglon(t.y, t.h)
            r.toks.append(t)
            renglones.append(r)
    for r in renglones:
        r.toks.sort(key=lambda t: t.x0)
        r.texto = " ".join(t.text for t in r.toks)
    return renglones


def limpiar_celda(texto):
    s = texto.strip().strip(MARCAS + " ")
    s = re.sub(r"^(\d[\d.,]*)\s+[vVyYrR✓✔√/]$", r"\1", s)     # palomita escrita a mano
    if re.fullmatch(r"\d[\d.,]*[.,]", s):
        s = s.rstrip(".,")                                            # "2,481." -> "2,481"
    return s


def tipo_celda(texto):
    """('lote'|'fecha'|'dinero'|'entero'|'cajas'|'codigo'|'texto'|'vacio', valor)."""
    s = limpiar_celda(texto)
    if not s:
        return "vacio", None
    lotes = lotes_en(s)
    if lotes and re.fullmatch(r"[A-Za-z]{0,4}[\s:#.\-]*" + lotes[0] + r"[\s.]*", s):
        return "lote", lotes[0]
    if FECHA_RE.match(s):
        return "fecha", s
    if DINERO_RE.match(s):
        return "dinero", a_dinero(s)
    v = a_entero(s)
    if v is not None:
        return "entero", v
    m = UNIDAD_RE.match(s)
    if m and a_entero(m.group(1)) is not None:
        return ("cajas" if m.group(2).lower().startswith("caj") else "entero"), a_entero(m.group(1))
    if re.search(r"[A-Za-z]", s) and re.search(r"\d", s) and " " not in s:
        return "codigo", s
    return "texto", s


def es_valor(texto):
    return tipo_celda(texto)[0] in ("entero", "dinero", "fecha", "lote", "cajas")


def es_marca(texto):
    return bool(re.fullmatch(r"[✓✔√/\\|*vVyY]", texto.strip()))


def armar_celdas(renglon):
    """Une palabras cercanas en celdas, sin juntar nunca dos valores numericos."""
    grupos = []
    for t in renglon.toks:
        if grupos:
            g = grupos[-1]
            ult = g[-1]
            gtexto = " ".join(x.text for x in g)
            hueco = t.x0 - ult.x1
            h = max(t.h, ult.h)
            misma_caja = t.caja == ult.caja and t.conf < 1.0
            if gtexto.strip() == "$" and re.match(r"\d", t.text) and hueco < 12 * h:
                g.append(t)
                continue
            if es_marca(t.text) and hueco < 1.5 * h and re.search(r"\d", gtexto):
                g.append(t)
                continue
            juntar = not (es_valor(gtexto) and es_valor(t.text)) and not lotes_en(t.text) \
                and not lotes_en(gtexto) and (misma_caja or hueco < 0.5 * h)
            if juntar:
                g.append(t)
                continue
        grupos.append([t])
    celdas = []
    for g in grupos:
        texto = " ".join(x.text for x in g)
        tipo, valor = tipo_celda(texto)
        celdas.append(Celda(g[0].x0, g[-1].x1, texto, min(x.conf for x in g), tipo, valor))
    return celdas


def solape(a0, a1, b0, b1):
    return min(a1, b1) - max(a0, b0)


# ----------------------- formularios de lado -------------------------------
def analizar_de_lado(renglones, celdas, pnum, tipo_pag, usados, traza):
    """Un renglon 'Lote' con un lote por columna y renglones de cantidad alineados."""
    entradas = []
    for i, r in enumerate(renglones):
        cs = celdas[i]
        etiqueta = " ".join(c.text for c in cs if c.tipo in ("texto", "codigo") and c.x1 <= min(
            (x.x0 for x in cs if x.tipo == "lote"), default=0) + 1)
        cols = [((c.x0 + c.x1) / 2, c.valor) for c in cs if c.tipo == "lote"]
        if len(cols) < 2 or not LOTE_PALABRA.search(etiqueta):
            continue
        xs = sorted(c[0] for c in cols)
        tol = min(b - a for a, b in zip(xs, xs[1:])) * 0.5
        mejor = None
        for j in range(max(0, i - 15), min(len(renglones), i + 25)):
            if j == i:
                continue
            cj = celdas[j]
            primero = min((c.x0 for c in cj if c.tipo == "entero"), default=None)
            if primero is None:
                continue
            etq = " ".join(c.text for c in cj if c.x1 <= primero and c.tipo in ("texto", "codigo"))
            sem = semantica(etq)
            if not (sem["cant"] or sem["ret"] or sem["nod"]):
                continue
            prio = next((k for k, w in enumerate(PRIORIDAD_RENGLON) if w in sin_acentos(etq)),
                        len(PRIORIDAD_RENGLON))
            pares = []
            for c in cj:
                if c.tipo != "entero":
                    continue
                cx = (c.x0 + c.x1) / 2
                col = min(cols, key=lambda k: abs(k[0] - cx))
                if abs(col[0] - cx) <= tol:
                    pares.append((col[1], c.valor, c.conf))
            if pares and (mejor is None or prio < mejor[0]):
                mejor = (prio, pares, etq, j, sem)
        traza.append(f"pagina {pnum}: formulario de lado, renglon '{etiqueta}' con lotes "
                     f"{[c[1] for c in cols]}" + (f"; cantidades del renglon '{mejor[2]}'" if mejor
                                                   else "; ningun renglon de cantidad con numeros"))
        usados.add(i)
        if mejor:
            usados.add(mejor[3])
            tipo = "nod" if mejor[4]["nod"] else ("ret" if mejor[4]["ret"] else tipo_pag)
            for lote, v, conf in mejor[1]:
                entradas.append(Entrada("", pnum, lote, v, tipo, "de lado", conf,
                                        notas=[f"renglon '{mejor[2]}'"]))
    return entradas


# -------------------------------- tablas -----------------------------------
def filas_con_items(renglones, celdas, usados):
    """Renglones con un lote y al menos un entero. Tambien acepta el lote en el renglon
    siguiente, debajo de la descripcion ("TORNILLO ... / LOTE: 25240310")."""
    items, saltar = [], set()
    for i, cs in enumerate(celdas):
        if i in usados or i in saltar:
            continue
        lotes = [c for c in cs if c.tipo == "lote"] + [c for c in cs if c.tipo != "lote" and lotes_en(c.text)]
        enteros = [c for c in cs if c.tipo == "entero"]
        if lotes and enteros:
            lc = lotes[0]
            items.append(Item(i, lc.valor if lc.tipo == "lote" else lotes_en(lc.text)[0], lc, cs))
        elif enteros and not lotes and len(cs) >= 3:
            for j in (i + 1, i + 2):
                if j >= len(celdas) or j in usados:
                    break
                cj = celdas[j]
                if any(c.tipo == "entero" for c in cj):
                    break
                lc = next((c for c in cj if lotes_en(c.text)), None)
                if lc and len(cj) <= 2 and any(c.tipo in ("texto", "codigo") and
                                               solape(c.x0, c.x1, lc.x0, lc.x1) > 0 for c in cs):
                    items.append(Item(i, lotes_en(lc.text)[0], lc, cs))
                    saltar.add(j)
                    break
    return items, saltar


def agrupar_tablas(items, renglones):
    if not items:
        return []
    hmed = sorted(r.h for r in renglones)[len(renglones) // 2]
    tablas, actual = [], [items[0]]
    for it in items[1:]:
        if renglones[it.fila].y - renglones[actual[-1].fila].y > 5 * hmed:
            tablas.append(actual)
            actual = []
        actual.append(it)
    tablas.append(actual)
    return tablas


def columnas_de(tabla):
    """Agrupa las celdas de los renglones de la tabla en columnas por posicion."""
    cols = []
    for k, it in enumerate(tabla):
        for c in it.celdas:
            w = max(c.x1 - c.x0, 1e-6)
            mejor, ov_mejor = None, 0
            for col in cols:
                ov = solape(c.x0, c.x1, col["x0"], col["x1"])
                if ov > ov_mejor:
                    mejor, ov_mejor = col, ov
            if mejor is not None and ov_mejor >= 0.4 * min(w, mejor["x1"] - mejor["x0"]):
                mejor["x0"], mejor["x1"] = min(mejor["x0"], c.x0), max(mejor["x1"], c.x1)
                mejor["celdas"].setdefault(k, []).append(c)
            else:
                cols.append({"x0": c.x0, "x1": c.x1, "celdas": {k: [c]}})
    cols.sort(key=lambda c: c["x0"])
    for col in cols:                    # varias celdas del mismo renglon -> una
        for k, cs in list(col["celdas"].items()):
            if len(cs) > 1:
                texto = " ".join(c.text for c in sorted(cs, key=lambda c: c.x0))
                tipo, valor = tipo_celda(texto)
                col["celdas"][k] = [Celda(min(c.x0 for c in cs), max(c.x1 for c in cs), texto,
                                          min(c.conf for c in cs), tipo, valor)]
        col["celda"] = {k: cs[0] for k, cs in col["celdas"].items()}
    return cols


def encabezados(cols, tabla, renglones, celdas):
    """Texto de encabezado de cada columna: renglones sin numeros justo encima de la
    tabla (hasta 3, para encabezados en dos lineas) y la etiqueta a la izquierda."""
    primero = tabla[0].fila
    hmed = renglones[primero].h
    zona, y_prev = [], renglones[primero].y
    for j in range(primero - 1, max(-1, primero - 4), -1):
        r = renglones[j]
        if lotes_en(r.texto) or y_prev - r.y > 3.5 * max(hmed, r.h):
            break
        if sum(1 for c in celdas[j] if c.tipo in ("entero", "dinero")) > 1:
            break
        zona.append(r)
        y_prev = r.y
    textos = [[] for _ in cols]
    for r in zona:
        for t in r.toks:
            cx = (t.x0 + t.x1) / 2
            dentro = [k for k, c in enumerate(cols) if c["x0"] - 0.5 * t.h <= cx <= c["x1"] + 0.5 * t.h]
            if dentro:
                k = min(dentro, key=lambda k: abs((cols[k]["x0"] + cols[k]["x1"]) / 2 - cx))
            else:
                k = min(range(len(cols)), key=lambda k: abs((cols[k]["x0"] + cols[k]["x1"]) / 2 - cx))
                if abs((cols[k]["x0"] + cols[k]["x1"]) / 2 - cx) > 3 * t.h + (cols[k]["x1"] - cols[k]["x0"]):
                    continue
            textos[k].append(t.text)
    for k, col in enumerate(cols):      # etiqueta inmediatamente a la izquierda ("Piezas: 30")
        for i, c in col["celda"].items():
            fila = tabla[i].celdas
            izq = [x for x in fila if x.x1 <= c.x0 + 1 and x.tipo in ("texto", "codigo")]
            if izq:
                etq = max(izq, key=lambda x: x.x1)
                if etq.text.rstrip().endswith(":") or semantica(etq.text)["cant"]:
                    textos[k].append(etq.text)
    return [" ".join(t) for t in textos]


def es_secuencia(vals):
    return len(vals) >= 2 and all(b - a == 1 for a, b in zip(vals, vals[1:]))


def relaciones_suma(M, cand):
    """Ternas (a, b, s) con v[a] + v[b] == v[s] en todos los renglones (o en >= 80%)."""
    res = []
    n = len(M)
    for s in cand:
        for a in cand:
            for b in cand:
                if len({a, b, s}) < 3 or a > b:
                    continue
                ok = sum(1 for fila in M if None not in (fila[a], fila[b], fila[s])
                         and fila[a] + fila[b] == fila[s])
                completos = sum(1 for fila in M if None not in (fila[a], fila[b], fila[s]))
                trivial = all(fila[a] in (0, None) for fila in M) or all(fila[b] in (0, None) for fila in M)
                if completos >= 2 and not trivial and ok >= max(2, completos - (completos // 5)):
                    res.append((a, b, s))
    return res


def roles_de_reporte(M, cand, sems, sumas):
    """Reporte de retiro: fabricado = distribuido + no distribuido, y
    total = no distribuido + retirado. El sumando comun es lo no distribuido; el otro
    sumando menor es lo retirado (no se puede retirar mas de lo distribuido)."""
    for i, (a1, b1, s1) in enumerate(sumas):
        for a2, b2, s2 in sumas[i + 1:]:
            comun = {a1, b1} & {a2, b2}
            if len(comun) != 1 or s1 == s2:
                continue
            x = comun.pop()
            p = b1 if a1 == x else a1
            q = b2 if a2 == x else a2
            if p == q:
                continue
            filas = [f for f in M if f[p] is not None and f[q] is not None]
            if filas and all(f[p] <= f[q] for f in filas):
                return {"nod": x, "ret": p, "total": s1, "dist": q, "fab": s2}
            if filas and all(f[q] <= f[p] for f in filas):
                return {"nod": x, "ret": q, "total": s2, "dist": p, "fab": s1}
    for a, b, s in sumas:               # una sola suma: hace falta el encabezado
        if sems[s]["total"]:
            if sems[a]["nod"] or sems[b]["ret"]:
                return {"nod": a, "ret": b, "total": s}
            if sems[b]["nod"] or sems[a]["ret"]:
                return {"nod": b, "ret": a, "total": s}
    return None


def fila_de_reporte(f, roles, pnum, lote, traza):
    """Una fila del reporte. Si un valor falta o esta mal leido, las dos relaciones
    (fabricado - distribuido = no distribuido; no distribuido + retirado = total)
    permiten recuperarlo."""
    g = lambda k: f[roles[k]] if k in roles else None
    fab, dist, nod, ret, tot = g("fab"), g("dist"), g("nod"), g("ret"), g("total")
    opciones_nod = [x for x in (nod, None if None in (fab, dist) else fab - dist,
                                None if None in (tot, ret) else tot - ret) if x is not None and x >= 0]
    nod_ok = Counter(opciones_nod).most_common(1)[0][0] if opciones_nod else None
    cuadra = None not in (nod_ok, ret, tot) and nod_ok + ret == tot and nod_ok == nod
    opciones = ()
    if not cuadra:
        if ret is None and None not in (tot, nod_ok):
            ret = tot - nod_ok
        elif None not in (tot, nod_ok, ret):
            opciones = (tot - nod_ok,)
        traza.append(f"pagina {pnum}: fila del reporte del lote {lote} no cuadra "
                     f"(no distribuido {nod}, retirado {ret}, total {tot}); se acepta tambien "
                     f"retirado = {fmt(tot - nod_ok) if None not in (tot, nod_ok) else '-'}")
    return Referencia("", pnum, lote, ret, nod_ok, tot, bool(cuadra), opciones)


def analizar_tablas(renglones, celdas, pnum, tipo_pag, usados, traza):
    entradas, referencias = [], []
    items, saltados = filas_con_items(renglones, celdas, usados)
    tablas = agrupar_tablas(items, renglones)
    filas_items = {it.fila for it in items} | saltados
    # Numeros fuera de las tablas: posibles totales impresos del documento.
    enteros_doc, dinero_doc, totales_rotulados = [], [], []
    for i, cs in enumerate(celdas):
        if i in filas_items:
            continue
        ents = [c.valor for c in cs if c.tipo == "entero"]
        enteros_doc += ents
        dinero_doc += [c.valor for c in cs if c.tipo == "dinero"]
        t = sin_acentos(renglones[i].texto)
        if ents and "total" in t and re.search(r"pieza|cantidad|pzas|unidades", t):
            totales_rotulados.append(max(ents))

    for tabla in tablas:
        cols = columnas_de(tabla)
        heads = encabezados(cols, tabla, renglones, celdas)
        sems = [semantica(h) for h in heads]
        n = len(tabla)
        col_lote = next((k for k, c in enumerate(cols)
                         if any(x is it.celda_lote for i, it in enumerate(tabla) for x in [c["celda"].get(i)])), None)
        # Columnas de enteros candidatas a cantidad
        cand = []
        for k, col in enumerate(cols):
            if k == col_lote:
                continue
            tipos = Counter(c.tipo for c in col["celda"].values() if c.tipo != "vacio")
            if not tipos or tipos["entero"] < 0.6 * sum(tipos.values()) or len(col["celda"]) < 0.5 * n:
                continue
            vals = [col["celda"][i].valor if i in col["celda"] and col["celda"][i].tipo == "entero" else None
                    for i in range(n)]
            conocidos = [v for v in vals if v is not None]
            if n >= 2 and es_secuencia(conocidos):
                continue                  # numero de renglon
            if n >= 2 and len(set(conocidos)) == 1 and conocidos[0] >= 10000:
                continue                  # codigo repetido (p. ej. CODIGO SAP)
            if all(1990 <= v <= 2100 for v in conocidos) and not sems[k]["cant"]:
                continue                  # anios
            cand.append(k)
        M = [[(cols[k]["celda"][i].valor if i in cols[k]["celda"] and cols[k]["celda"][i].tipo == "entero"
               else None) if k in cand else None for k in range(len(cols))] for i in range(n)]

        # 1) Reporte de retiro (relaciones aritmeticas entre columnas)
        roles = None
        if len(cand) >= 3 and n >= 2:
            sumas = relaciones_suma(M, cand)
            if sumas:
                roles = roles_de_reporte(M, cand, sems, sumas)
        if roles:
            traza.append(f"pagina {pnum}: reporte de retiro; columnas: no distribuido '{heads[roles['nod']]}', "
                         f"retirado '{heads[roles['ret']]}', total '{heads[roles['total']]}'")
            for i, it in enumerate(tabla):
                referencias.append(fila_de_reporte(M[i], roles, pnum, it.lote, traza))
            continue

        # 2) Tabla normal: elegir la columna de cantidad
        puntos = {}
        detalles = {}
        for k in cand:
            p, por = 0.0, []
            s = sems[k]
            if s["cant"] or s["ret"] or s["nod"]:
                p += 2
                por.append("encabezado de cantidad")
            if s["ret"] or s["nod"]:
                p += 1
            if s["orig"] and not (s["ret"] or s["nod"]):
                p -= 3
                por.append("cantidad original (facturada/distribuida)")
            if s["caja"] and not re.search(r"pieza", sin_acentos(heads[k])):
                p -= 4
                por.append("cuenta cajas")
            vals = [M[i][k] for i in range(n)]
            rel = 0
            for i, it in enumerate(tabla):
                din = sorted(c.valor for c in it.celdas if c.tipo == "dinero")
                if vals[i] is not None and any(abs(vals[i] * a - b) <= 0.011 for a in din for b in din if a < b):
                    rel += 1
            if rel:
                p += 4
                por.append(f"piezas x precio = importe en {rel} renglon(es)")
            suma = sum(v for v in vals if v is not None)
            if suma >= 10 and suma in enteros_doc:
                p += 3
                por.append(f"suma {fmt(suma)} = numero impreso en el documento")
            if len(cand) == 1 and n >= 2:
                p += 1
                por.append("unica columna de enteros")
            if all(v is not None and v >= 1_000_000 for v in vals):
                p -= 5
            puntos[k], detalles[k] = p, por
        if not puntos:
            traza.append(f"pagina {pnum}: tabla de {n} renglon(es) sin columna de enteros utilizable")
            continue
        k = max(puntos, key=lambda c: (puntos[c], -cols[c]["x0"]))
        minimo = 2 if n == 1 else 1
        if puntos[k] < minimo:
            traza.append(f"pagina {pnum}: tabla de {n} renglon(es) descartada: ninguna columna de cantidad "
                         f"clara ({', '.join(repr(heads[c]) for c in cand) or 'sin columnas de enteros'})")
            continue
        empate = [c for c in puntos if c != k and puntos[c] == puntos[k]]
        s = sems[k]
        tipo = "nod" if (s["nod"] and not s["ret"]) else ("ret" if s["ret"] else tipo_pag)
        traza.append(f"pagina {pnum}: tabla de {n} renglon(es); columna de cantidad '{heads[k]}' "
                     f"({'; '.join(detalles[k]) or 'por estructura'})"
                     + (f"; EMPATE con {[heads[c] for c in empate]}" if empate else ""))
        nuevas = []
        for i, it in enumerate(tabla):
            c = cols[k]["celda"].get(i)
            if c is None or c.tipo != "entero":
                continue
            e = Entrada("", pnum, it.lote, c.valor, tipo, "tabla", c.conf)
            din = sorted(x.valor for x in it.celdas if x.tipo == "dinero")
            par = next(((a, b) for a in din for b in din if a < b and abs(c.valor * a - b) <= 0.011), None)
            if par:
                e.verificado.append(f"piezas x precio = importe ({fmt(c.valor)} x {par[0]:,.2f} = {par[1]:,.2f})")
            e.precio = par[0] if par else (din[0] if len(din) == 1 else None)
            e.dinero = din
            if empate:
                e.notas.append("columna de cantidad dudosa (empate)")
            nuevas.append(e)
        verificar_tabla(nuevas, enteros_doc, dinero_doc, totales_rotulados)
        entradas += nuevas
        usados.update(it.fila for it in tabla)
    usados.update(saltados)
    return entradas, referencias


def verificar_tabla(nuevas, enteros_doc, dinero_doc, totales_rotulados):
    """Compara la suma de la tabla con los totales impresos. Si el total rotulado de
    piezas no coincide y solo un renglon no esta verificado, lo corrige cuando el
    precio confirma el valor corregido."""
    if not nuevas:
        return
    suma = sum(e.piezas for e in nuevas)
    if suma >= 10 and suma in enteros_doc:
        for e in nuevas:
            e.verificado.append(f"suma de la tabla = total impreso ({fmt(suma)})")
    precios = [getattr(e, "precio", None) for e in nuevas]
    if all(p is not None for p in precios):
        importe = round(sum(e.piezas * p for e, p in zip(nuevas, precios)), 2)
        if any(abs(importe - d) <= 0.011 for d in dinero_doc):
            for e in nuevas:
                e.verificado.append(f"suma de piezas x precio = importe total impreso ({importe:,.2f})")
    for total in totales_rotulados:
        if total == suma:
            continue
        sin_verif = [e for e in nuevas if not e.verificado]
        if len(sin_verif) == 1:
            e = sin_verif[0]
            correcto = total - (suma - e.piezas)
            p = getattr(e, "precio", None)
            din = getattr(e, "dinero", [])
            con_renglon = any(abs(correcto * a - b) <= 0.011 for a in din for b in din if a < b)
            con_total = bool(p) and any(abs(correcto * p - d) <= 0.011 for d in dinero_doc)
            if correcto > 0 and (con_renglon or con_total):
                e.notas.append(f"CORREGIDO: se leyo {fmt(e.piezas)}; el total impreso ({fmt(total)}) "
                               f"y el importe confirman {fmt(correcto)}")
                e.piezas = correcto
                e.verificado.append("corregido con el total y el importe impresos")
                continue
        for e in nuevas:
            e.fallas.append(f"la suma de la tabla ({fmt(suma)}) no coincide con el total de piezas "
                            f"impreso ({fmt(total)})")


# --------------------------- lineas con etiquetas --------------------------
def analizar_etiquetas(renglones, pnum, tipo_pag, usados, traza):
    ev = []
    for i, r in enumerate(renglones):
        if i in usados:
            continue
        e = [(m.start(), "L", lotes_en(m.group("id"))[0] if lotes_en(m.group("id")) else None, 1.0)
             for m in LOTE_ETIQUETA.finditer(r.texto)]
        e += [(m.start(), "C", a_entero(m.group("n").rstrip(".,")), min(t.conf for t in r.toks))
              for m in CANT_ETIQUETA.finditer(r.texto)]
        ev += [(k, v, c) for _, k, v, c in sorted(e, key=lambda x: x[0]) if v is not None]
    entradas = []
    if not ev:
        return entradas
    primero, pend = ev[0][0], None
    for kind, val, conf in ev:
        if kind == primero:
            pend = (val, conf)
        elif pend is not None:
            lote, v, c = (pend[0], val, conf) if primero == "L" else (val, pend[0], pend[1])
            entradas.append(Entrada("", pnum, lote, v, tipo_pag, "etiquetas", c))
            pend = None
    if entradas:
        traza.append(f"pagina {pnum}: lineas con etiquetas: "
                     + ", ".join(f"lote {e.lote} = {fmt(e.piezas)}" for e in entradas))
    return entradas


# ------------------------------- un PDF ------------------------------------
def analizar_pagina(renglones, pnum, traza):
    celdas = [armar_celdas(r) for r in renglones]
    tipo_pag = tipo_documento(" ".join(r.texto for r in renglones))
    usados = set()
    entradas = analizar_de_lado(renglones, celdas, pnum, tipo_pag, usados, traza)
    e2, refs = analizar_tablas(renglones, celdas, pnum, tipo_pag, usados, traza)
    entradas += e2
    entradas += analizar_etiquetas(renglones, pnum, tipo_pag, usados, traza)
    return entradas, refs, tipo_pag


def deduplicar_en_archivo(entradas, avisos, traza):
    """El mismo lote repetido en el mismo PDF no se suma dos veces."""
    por = defaultdict(list)
    for e in entradas:
        por[(e.tipo, e.lote)].append(e)
    for (tipo, lote), es in por.items():
        if len(es) < 2:
            continue
        paginas = defaultdict(list)
        for e in es:
            paginas[e.pagina].append(e)
        for p, pes in paginas.items():          # dentro de una pagina
            vistos = set()
            for e in pes:
                if e.piezas in vistos:
                    e.excluido = f"renglon repetido (lote {lote}, {fmt(e.piezas)}) en la misma pagina"
                vistos.add(e.piezas)
            activos = [e for e in pes if not e.excluido]
            if len(activos) >= 3:
                for e in activos:
                    if e.piezas == sum(x.piezas for x in activos if x is not e):
                        e.excluido = "renglon de total del lote"
                        break
        sumas = {p: sum(e.piezas for e in pes if not e.excluido) for p, pes in paginas.items()}
        if len(sumas) < 2:
            continue
        orden = sorted(sumas)
        valores = [sumas[p] for p in orden]
        if len(set(valores)) == 1 or DUPLICADOS_EN_ARCHIVO == "todos":
            if len(set(valores)) == 1:
                quedar = orden[0]
            else:
                quedar = max(orden, key=lambda p: (sumas[p], -p))
            for p in orden:
                if p != quedar:
                    for e in paginas[p]:
                        if not e.excluido:
                            e.excluido = f"lote repetido en la pagina {quedar} del mismo archivo"
            traza.append(f"lote {lote}: aparece en las paginas {orden}; se cuenta una vez (pagina {quedar})")
            continue
        total_pag = next((p for p in orden if sumas[p] == sum(sumas[q] for q in orden if q != p)), None)
        if total_pag is not None:
            for e in paginas[total_pag]:
                if not e.excluido:
                    e.excluido = "pagina con el total del lote (suma de las otras paginas)"
            traza.append(f"lote {lote}: la pagina {total_pag} es la suma de las otras; no se cuenta")
            continue
        avisos.append(f"lote {lote} aparece en varias paginas con cantidades distintas "
                      f"({', '.join(f'p{p}: {fmt(sumas[p])}' for p in orden)}); se suman")


def leer_pdf(ruta, nombre, pasada=1, traza=None):
    traza = [] if traza is None else traza
    entradas, refs, lineas, fuentes, tipos, avisos = [], [], [], set(), [], []
    with pdfplumber.open(ruta) as doc:
        for pnum, page in enumerate(doc.pages, 1):
            toks, fuente = page_tokens(page, pasada)
            fuentes.add(fuente)
            renglones = agrupar_renglones(toks)
            lineas += [r.texto for r in renglones]
            traza.append(f"pagina {pnum}: {len(renglones)} renglones ({fuente})")
            e, r, tipo_pag = analizar_pagina(renglones, pnum, traza)
            tipos.append(tipo_pag)
            entradas += e
            refs += r
    # paginas sin tipo reconocido: tipo mayoritario del archivo
    conocido = Counter(t for t in tipos if t)
    tipo_archivo = conocido.most_common(1)[0][0] if conocido else None
    for e in entradas:
        e.archivo = nombre
        if e.tipo is None:
            e.tipo = tipo_archivo or "ret"
            if tipo_archivo is None:
                e.notas.append("tipo de documento no reconocido: se cuenta como retirado")
    for r in refs:
        r.archivo = nombre
    deduplicar_en_archivo(entradas, avisos, traza)
    fuente = "mixto" if len(fuentes) > 1 else (fuentes.pop() if fuentes else "")
    return Lectura(nombre, entradas, refs, lineas, fuente, traza, pasada, avisos=avisos)


def calidad(L):
    ents = [e for e in L.entradas if not e.excluido]
    confs = [e.conf for e in ents] or [0]
    return (-sum(len(e.fallas) for e in ents), sum(1 for e in ents if e.verificado),
            len(ents) + len(L.referencias), sum(confs) / len(confs))


def dudosa(L):
    ents = [e for e in L.entradas if not e.excluido]
    if L.fuente in ("OCR", "mixto") and not ents and not L.referencias:
        return True
    return any(e.fallas or (not e.verificado and e.conf < CONFIANZA_MINIMA) for e in ents)


# ---------------------------- conciliacion ---------------------------------
@dataclass
class Resultado:
    ret: dict
    nod: dict
    refs: dict
    nod_fuente: dict
    problemas: list
    total: int


def conciliar(lecturas):
    ret, nod_doc = defaultdict(int), defaultdict(int)
    for L in lecturas.values():
        for e in L.entradas:
            if not e.excluido:
                (ret if e.tipo == "ret" else nod_doc)[e.lote] += e.piezas
    refs = {}
    for L in lecturas.values():
        for r in L.referencias:
            if r.lote not in refs or (r.cuadra and not refs[r.lote].cuadra):
                refs[r.lote] = r
    nod, nod_fuente = dict(nod_doc), {l: "documentos" for l in nod_doc}
    for lote, r in refs.items():
        if r.nod is not None and lote not in nod_doc:
            nod[lote], nod_fuente[lote] = r.nod, "reporte"
        if USAR_REPORTE_PARA_RETIRADO and r.ret is not None and lote not in ret:
            ret[lote] = r.ret
    problemas = []
    for lote in sorted(set(ret) | set(nod) | set(refs), key=orden_lote):
        rv, nv, r = ret.get(lote, 0), nod.get(lote, 0), refs.get(lote)
        if rv > MAX_RETIRADO_POR_LOTE:
            problemas.append(("tope", "ret", lote, rv, MAX_RETIRADO_POR_LOTE))
        if nv > MAX_NO_DISTRIBUIDA_POR_LOTE:
            problemas.append(("tope", "nod", lote, nv, MAX_NO_DISTRIBUIDA_POR_LOTE))
        if r and r.ret is not None and not r.acepta_ret(rv):
            problemas.append(("reporte", "ret", lote, rv, r.ret))
        if r and r.nod is not None and lote in nod_doc and nod_doc[lote] != r.nod:
            problemas.append(("reporte", "nod", lote, nod_doc[lote], r.nod))
    total = sum(ret.values()) + sum(nod.values())
    return Resultado(dict(ret), nod, refs, nod_fuente, problemas, total)


def lote_cuadra(lecturas, tipo, lote, objetivo, tope, ref=None):
    s = sum(e.piezas for L in lecturas.values() for e in L.entradas
            if not e.excluido and e.tipo == tipo and e.lote == lote)
    if ref is not None and tipo == "ret":
        return ref.acepta_ret(s)
    return s == objetivo if objetivo is not None else s <= tope


COPIA_RE = re.compile(r"copia|copy|duplicad|\(\d+\)|_\d+$|-\s*\d+$", re.I)


def ordenar_copias(entradas):
    """El original primero: los nombres con 'copia', '(1)', '_2'... van despues."""
    return sorted(entradas, key=lambda e: (bool(COPIA_RE.search(pathlib.Path(e.archivo).stem)),
                                           len(e.archivo), e.archivo))


def cambiar_lectura(lecturas, alternativas, f):
    lecturas[f], alternativas[f] = alternativas[f], lecturas[f]


def intentar_arreglos(lecturas, alternativas, res, total_esperado, malos, correcciones):
    """Busca una explicacion exacta y unica para cada problema: otra lectura del mismo
    PDF, o un renglon que es copia exacta de otro PDF. Solo entonces se corrige."""
    for clase, tipo, lote, valor, ref in res.problemas:
        objetivo = ref if clase == "reporte" else None
        tope = ref if clase == "tope" else None
        refobj = res.refs.get(lote) if clase == "reporte" else None
        cuadra = lambda lect: lote_cuadra(lect, tipo, lote, objetivo, tope, refobj)
        if cuadra(lecturas):
            continue
        archivos = sorted({e.archivo for L in lecturas.values() for e in L.entradas
                           if not e.excluido and e.tipo == tipo and e.lote == lote})
        # a) otra lectura de un PDF
        arreglos = []
        for f in archivos:
            alt = alternativas.get(f)
            if alt is None or alt is lecturas[f]:
                continue
            prueba = dict(lecturas)
            prueba[f] = alt
            if cuadra(prueba):
                arreglos.append(f)
        if len(arreglos) == 1:
            f = arreglos[0]
            cambiar_lectura(lecturas, alternativas, f)
            correcciones.append(f"lote {lote}: se usa la lectura {lecturas[f].pasada} de {f} "
                                f"(con ella el lote cuadra)")
            continue
        # b) renglon copiado exactamente de otro PDF
        copias = []
        por_clave = defaultdict(list)
        for L in lecturas.values():
            for e in L.entradas:
                if not e.excluido and e.tipo == tipo and e.lote == lote:
                    por_clave[e.clave()].append(e)
        for clave, es in por_clave.items():
            if len({e.archivo for e in es}) < 2:
                continue
            for e in ordenar_copias(es)[1:]:
                e.excluido = "prueba"
                if cuadra(lecturas):
                    copias.append(e)
                e.excluido = ""
        if len({c.archivo for c in copias}) == 1:
            c = copias[0]
            orig = ordenar_copias(por_clave[c.clave()])[0]
            c.excluido = f"copia de {orig.archivo} (mismo lote y cantidad)"
            malos.setdefault(c.archivo, []).append(
                f"lote {lote}: {fmt(c.piezas)} piezas repetidas de {orig.archivo}; sin esta copia el lote "
                + ("cuadra con el reporte" if objetivo is not None else "queda dentro del tope"))

    if total_esperado is None:
        return
    res2 = conciliar(lecturas)
    dif = res2.total - total_esperado
    if dif == 0:
        return
    # c) total global: otra lectura de un solo PDF que haga coincidir el total exacto
    arreglos = []
    for f, alt in alternativas.items():
        if alt is lecturas.get(f):
            continue
        prueba = dict(lecturas)
        prueba[f] = alt
        if conciliar(prueba).total == total_esperado:
            arreglos.append(f)
    if len(arreglos) == 1:
        f = arreglos[0]
        cambiar_lectura(lecturas, alternativas, f)
        correcciones.append(f"se usa la lectura {lecturas[f].pasada} de {f}: con ella el total coincide "
                            f"exactamente con el esperado ({fmt(total_esperado)})")
        return
    # d) total global: una copia exacta entre PDFs que explique toda la diferencia
    por_clave = defaultdict(list)
    for L in lecturas.values():
        for e in L.entradas:
            if not e.excluido:
                por_clave[e.clave()].append(e)
    candidatos = [ordenar_copias(es)[1:] for es in por_clave.values() if len({e.archivo for e in es}) > 1]
    candidatos = [e for grupo in candidatos for e in grupo if e.piezas == dif and e.lote not in res2.refs]
    if len(candidatos) == 1:
        c = candidatos[0]
        orig = ordenar_copias(por_clave[c.clave()])[0]
        c.excluido = f"copia de {orig.archivo} (mismo lote y cantidad)"
        malos.setdefault(c.archivo, []).append(
            f"lote {c.lote}: {fmt(c.piezas)} piezas repetidas de {orig.archivo}; sin esta copia el total "
            f"coincide exactamente con el esperado ({fmt(total_esperado)})")


# ------------------------------- salidas -----------------------------------
ROJO = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")


def escribir_excel(ruta, res, total_esperado):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lotes"
    negrita = Font(bold=True)
    lotes = sorted(set(res.ret) | set(res.nod) | set(res.refs), key=orden_lote)
    malos = {(t, l) for _, t, l, _, _ in res.problemas}
    n = len(lotes)
    fin = 3 + n                        # fila del TOTAL
    titulos = ["Cantidad de producto retirado", "Cantidad no distribuida", "Cantidad total para destrucción"]
    for t, titulo in enumerate(titulos):
        c0 = 1 + 3 * t
        ws.cell(1, c0, titulo).font = negrita
        ws.cell(2, c0, "Lote").font = negrita
        ws.cell(2, c0 + 1, "Cantidad en Piezas").font = negrita
        for i, lote in enumerate(lotes):
            fila = 3 + i
            ws.cell(fila, c0, int(lote) if lote.isdigit() else lote)
            if t == 0:
                celda = ws.cell(fila, c0 + 1, res.ret.get(lote, 0))
                if ("ret", lote) in malos:
                    celda.fill = ROJO
            elif t == 1:
                celda = ws.cell(fila, c0 + 1, res.nod.get(lote, 0))
                if ("nod", lote) in malos:
                    celda.fill = ROJO
            else:
                celda = ws.cell(fila, c0 + 1, f"=B{fila}+E{fila}")
            celda.number_format = "#,##0"
        ws.cell(fin, c0, "TOTAL").font = negrita
        col = "BEH"[t]
        celda = ws.cell(fin, c0 + 1, f"=SUM({col}3:{col}{fin - 1})" if n else 0)
        celda.number_format = "#,##0"
        celda.font = negrita
        ws.column_dimensions["ADG"[t]].width = 14
        ws.column_dimensions["BEH"[t]].width = 20
    ws.column_dimensions["C"].width = 3
    ws.column_dimensions["F"].width = 3
    if total_esperado is not None:
        ws.cell(fin + 2, 7, "Total esperado")
        c = ws.cell(fin + 2, 8, total_esperado)
        c.number_format = "#,##0"
        ws.cell(fin + 3, 7, "Diferencia")
        c = ws.cell(fin + 3, 8, f"=H{fin}-H{fin + 2}")
        c.number_format = "#,##0"
        if res.total != total_esperado:
            c.fill = ROJO
    ws.freeze_panes = "A3"
    wb.save(ruta)


def escribir_detalle(ruta, lecturas, res):
    porlote = defaultdict(list)
    for L in lecturas.values():
        for e in L.entradas:
            porlote[e.lote].append(e)
    lineas = []
    for lote in sorted(set(porlote) | set(res.refs), key=orden_lote):
        rv, nv = res.ret.get(lote, 0), res.nod.get(lote, 0)
        lineas.append(f"{lote}   retirado {fmt(rv)}   no distribuida {fmt(nv)}   total {fmt(rv + nv)}")
        for e in sorted(porlote[lote], key=lambda e: (e.tipo, e.archivo, e.pagina)):
            estado = f"NO SUMADO: {e.excluido}" if e.excluido else (
                "verificado" if e.verificado else f"sin verificar (confianza {e.conf:.2f})")
            lineas.append(f"      {TIPOS[e.tipo]:<15} {fmt(e.piezas):>9}   {e.archivo} p{e.pagina}  [{e.forma}; {estado}]")
            for x in e.verificado + e.fallas + e.notas:
                lineas.append(f"            - {x}")
        r = res.refs.get(lote)
        if r:
            lineas.append(f"      reporte ({r.archivo}): retirado {fmt(r.ret) if r.ret is not None else '-'}, "
                          f"no distribuida {fmt(r.nod) if r.nod is not None else '-'}, "
                          f"total {fmt(r.total) if r.total is not None else '-'}"
                          + ("" if r.cuadra else "   (sus columnas no cuadran entre si)"))
            if res.nod_fuente.get(lote) == "reporte":
                lineas.append("      la cantidad no distribuida se toma del reporte")
        lineas.append("")
    pathlib.Path(ruta).write_text("\n".join(lineas), encoding="utf-8")


def escribir_revision(ruta, lecturas, res, total_esperado, omitidos, sin_resultado, malos,
                      correcciones):
    t = []
    t.append("RESUMEN")
    t.append(f"   Cantidad de producto retirado:   {fmt(sum(res.ret.values()))}")
    t.append(f"   Cantidad no distribuida:         {fmt(sum(res.nod.values()))}")
    t.append(f"   Cantidad total para destruccion: {fmt(res.total)}")
    if total_esperado is not None:
        dif = res.total - total_esperado
        t.append(f"   Total esperado:                  {fmt(total_esperado)}   "
                 + ("COINCIDE" if dif == 0 else f"DIFERENCIA {dif:+,}"))
    t.append("")
    t.append(f"PDFs MARCADOS COMO MALOS (prueba concluyente): {len(malos)}")
    for f, motivos in sorted(malos.items()):
        t.append(f"   {f}")
        t += [f"      - {m}" for m in motivos]
    t.append("")
    t.append(f"CORRECCIONES APLICADAS: {len(correcciones)}")
    t += [f"   {c}" for c in correcciones]
    for L in lecturas.values():
        for e in L.entradas:
            t += [f"   {e.archivo} p{e.pagina}: {n}" for n in e.notas if n.startswith("CORREGIDO")]
    t.append("")
    t.append(f"LOTES QUE NO CUADRAN: {len(res.problemas)}")
    for clase, tipo, lote, valor, ref in res.problemas:
        fuentes = sorted({f"{e.archivo} ({fmt(e.piezas)})" for L in lecturas.values() for e in L.entradas
                          if not e.excluido and e.tipo == tipo and e.lote == lote})
        if clase == "tope":
            t.append(f"   lote {lote}: {TIPOS[tipo]} {fmt(valor)} supera el tope de {fmt(ref)}")
        else:
            t.append(f"   lote {lote}: {TIPOS[tipo]} segun los documentos {fmt(valor)}, segun el reporte {fmt(ref)} "
                     f"(diferencia {valor - ref:+,})")
        t.append(f"      documentos: {', '.join(fuentes) or 'ninguno'}")
    t.append("")
    sospechosos = []
    for L in lecturas.values():
        motivos = []
        for e in L.entradas:
            if e.excluido:
                continue
            motivos += [f"lote {e.lote}: {x}" for x in e.fallas]
            if not e.verificado and e.conf < CONFIANZA_MINIMA:
                motivos.append(f"lote {e.lote}: {fmt(e.piezas)} leido con confianza baja ({e.conf:.2f})")
            motivos += [f"lote {e.lote}: {x}" for x in e.notas if "dudosa" in x]
        motivos += L.avisos
        if motivos and L.archivo not in malos:
            sospechosos.append((L.archivo, motivos))
    t.append(f"PDFs SOSPECHOSOS (revisar): {len(sospechosos)}")
    for f, motivos in sospechosos:
        t.append(f"   {f}")
        t += [f"      - {m}" for m in dict.fromkeys(motivos)]
    t.append("")
    rep = []
    for L in lecturas.values():
        for e in L.entradas:
            if e.excluido and not e.excluido.startswith("copia de"):
                rep.append(f"   {e.archivo} p{e.pagina}: lote {e.lote} {fmt(e.piezas)} -> {e.excluido}")
    t.append(f"REPETICIONES DENTRO DEL MISMO ARCHIVO (no sumadas): {len(rep)}")
    t += rep
    t.append("")
    por_clave = defaultdict(set)
    for L in lecturas.values():
        for e in L.entradas:
            if not e.excluido:
                por_clave[e.clave()].add(e.archivo)
    dup = [(k, fs) for k, fs in por_clave.items() if len(fs) > 1]
    t.append(f"MISMO LOTE Y CANTIDAD EN VARIOS ARCHIVOS (posibles duplicados, sumados): {len(dup)}")
    for (tipo, lote, v), fs in dup:
        t.append(f"   lote {lote} {TIPOS[tipo]} {fmt(v)}: {', '.join(sorted(fs))}")
    t.append("")
    t.append(f"PDFs SIN CANTIDADES LEGIBLES: {len(sin_resultado)}")
    t += [f"   {nm}   [{motivo}]" for nm, motivo in sin_resultado]
    t.append("")
    t.append(f"PDFs OMITIDOS POR 'FALTA REVISION' (no cuentan en ningun total): {len(omitidos)}")
    t += [f"   {nm}" for nm in omitidos]
    pathlib.Path(ruta).write_text("\n".join(t) + "\n", encoding="utf-8")


def ocr_install_hint():
    pkgs = "rapidocr onnxruntime" if sys.version_info >= (3, 13) else "rapidocr_onnxruntime"
    if os.name == "nt":
        pkgs += " msvc-runtime"
    return f'"{sys.executable}" -m pip install {pkgs}'


def resumen_archivo(L):
    ret = sum(1 for e in L.entradas if not e.excluido and e.tipo == "ret")
    nod = sum(1 for e in L.entradas if not e.excluido and e.tipo == "nod")
    partes = []
    if ret:
        partes.append(f"{ret} retirado")
    if nod:
        partes.append(f"{nod} no distribuida")
    if L.referencias:
        partes.append(f"reporte con {len(L.referencias)} lotes")
    return (", ".join(partes) or "0 cantidades") + f" ({L.fuente})"


# ---------------------------- programa principal ---------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="lotes_auto.py",
        description="Calcula por lote la cantidad retirada, la no distribuida y el total para "
                    "destruccion a partir de los PDFs de una carpeta (incluye subcarpetas).")
    ap.add_argument("carpeta", help="carpeta con los PDFs")
    ap.add_argument("--total-esperado", type=int, default=TOTAL_ESPERADO, metavar="N",
                    help=f"total esperado para destruccion (por omision: {TOTAL_ESPERADO}; 0 = no comprobar)")
    ap.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    args = ap.parse_args(argv)
    esperado = args.total_esperado or None

    folder = pathlib.Path(args.carpeta).expanduser()
    if not folder.is_dir():
        print(f"*** No existe la carpeta: {folder}")
        return 2
    pdfs = sorted(p for p in folder.rglob("*") if p.suffix.lower() == ".pdf")
    print(f"{len(pdfs)} PDFs encontrados en {folder}")

    lecturas, rutas, omitidos, sin_resultado, log = {}, {}, [], [], []
    aviso_ocr = False
    for n, pdf in enumerate(pdfs, 1):
        name = str(pdf.relative_to(folder))
        print(f"[{n}/{len(pdfs)}] {name}", end="", flush=True)
        if OMITIR_NOMBRE.search(pdf.name):
            omitidos.append(name)
            print("  -> omitido (FALTA REVISION)")
            continue
        try:
            L = leer_pdf(pdf, name)
        except OCRNoDisponible as e:
            print("  -> escaneado, OCR no disponible")
            sin_resultado.append((name, "escaneado; OCR no disponible en este sistema"))
            log += ["=" * 70, name, "Motivo: escaneado; OCR no disponible", f"Detalle: {e}", ""]
            if not aviso_ocr:
                print("     Para leer PDFs escaneados instala el OCR:  " + ocr_install_hint())
                aviso_ocr = True
            continue
        except Exception as e:
            print(f"  -> ERROR: {e}")
            sin_resultado.append((name, f"error al leer: {e}"))
            log += ["=" * 70, name, f"Motivo: error al leer: {e}", ""]
            continue
        lecturas[name], rutas[name] = L, pdf
        print(f"  -> {resumen_archivo(L)}")

    # ----- comprobaciones y segunda lectura -----
    res = conciliar(lecturas)
    alternativas = {}
    problema_lotes = {l for _, _, l, _, _ in res.problemas}
    hay_que_reintentar = bool(res.problemas) or (esperado is not None and res.total != esperado)
    dudosas = [f for f, L in lecturas.items() if L.fuente in ("OCR", "mixto") and (
        dudosa(L) or any(e.lote in problema_lotes for e in L.entradas)
        or (hay_que_reintentar and any(not e.verificado for e in L.entradas)))]
    if dudosas and (hay_que_reintentar or any(dudosa(lecturas[f]) for f in dudosas)):
        print(f"\nSegunda lectura ({OCR_DPI_REINTENTO} ppp, con mas contraste) de {len(dudosas)} PDF(s) dudoso(s)...")
        for f in dudosas:
            try:
                L2 = leer_pdf(rutas[f], f, pasada=2)
            except Exception as e:
                print(f"   {f}: {e}")
                continue
            alternativas[f] = L2 if calidad(L2) <= calidad(lecturas[f]) else lecturas[f]
            if calidad(L2) > calidad(lecturas[f]):
                alternativas[f] = lecturas[f]
                lecturas[f] = L2
                print(f"   {f}: se usa la segunda lectura -> {resumen_archivo(L2)}")
    malos, correcciones = {}, []
    res = conciliar(lecturas)
    if res.problemas or (esperado is not None and res.total != esperado):
        intentar_arreglos(lecturas, alternativas, res, esperado, malos, correcciones)
        res = conciliar(lecturas)
    for f, alt in alternativas.items():     # las dos lecturas no coinciden
        otra = defaultdict(int)
        for e in alt.entradas:
            if not e.excluido:
                otra[(e.tipo, e.lote)] += e.piezas
        for e in lecturas[f].entradas:
            if not e.excluido and (e.tipo, e.lote) in otra and otra[(e.tipo, e.lote)] != e.piezas:
                e.notas.append(f"dudosa: la otra lectura dio {fmt(otra[(e.tipo, e.lote)])}")
    for L in lecturas.values():             # lo que coincide con el reporte queda verificado
        for e in L.entradas:
            r = res.refs.get(e.lote)
            ref = None if r is None else (r.ret if e.tipo == "ret" else r.nod)
            total = res.ret.get(e.lote) if e.tipo == "ret" else res.nod.get(e.lote)
            if not e.excluido and ref is not None and (ref == total or (e.tipo == "ret" and r.acepta_ret(total))):
                e.verificado.append(f"el total del lote coincide con el reporte ({r.archivo})")
    for f, L in lecturas.items():           # cifras que contradicen el propio documento
        fallas = sorted({x for e in L.entradas if not e.excluido for x in e.fallas})
        if fallas:
            alt = alternativas.get(f)
            if alt is None or any(e.fallas for e in alt.entradas if not e.excluido):
                malos.setdefault(f, []).extend(f"sus propias cifras no cuadran: {x}" for x in fallas)

    for f, L in lecturas.items():
        if not L.entradas and not L.referencias:
            sin_resultado.append((f, "sin cantidades legibles"))
            log += ["=" * 70, f, f"Motivo: sin cantidades legibles (fuente: {L.fuente})", "Texto leido:"]
            log += ["  | " + x for x in L.lineas] or ["  | (nada: no se pudo leer texto)"]
            log += ["Analisis:"] + ["  - " + x for x in L.traza] + [""]

    # ----- salidas -----
    out = folder / "totales_lotes.xlsx"
    try:
        escribir_excel(out, res, esperado)
    except PermissionError:
        print(f"\n*** No se pudo guardar {out.name}: esta abierto en Excel. Cierralo y vuelve a correr el script.")
        return 1
    escribir_detalle(folder / "detalle_lotes.txt", lecturas, res)
    rev = folder / "revisar_a_mano.txt"
    escribir_revision(rev, lecturas, res, esperado, omitidos, sin_resultado, malos, correcciones)
    logf = folder / "sin_resultado.log"
    if log:
        logf.write_text("\n".join(log) + "\n", encoding="utf-8")
    elif logf.exists():
        logf.unlink()

    ret_t, nod_t = sum(res.ret.values()), sum(res.nod.values())
    print(f"\nCantidad de producto retirado:   {fmt(ret_t):>12}")
    print(f"Cantidad no distribuida:         {fmt(nod_t):>12}")
    print(f"Cantidad total para destruccion: {fmt(res.total):>12}")
    if esperado is not None:
        dif = res.total - esperado
        print(f"Total esperado:                  {fmt(esperado):>12}   "
              + ("COINCIDE" if dif == 0 else f"*** DIFERENCIA {dif:+,}"))
    if res.problemas:
        print(f"*** {len(res.problemas)} lote(s) no cuadran (tope o reporte); ver revisar_a_mano.txt")
    if malos:
        print(f"*** {len(malos)} PDF(s) marcados como MALOS:")
        for f, motivos in sorted(malos.items()):
            print(f"      {f}: {motivos[0]}")
    if correcciones:
        print(f"{len(correcciones)} correccion(es) aplicadas; ver revisar_a_mano.txt")
    print(f"{len(omitidos)} omitidos por FALTA REVISION | {len(sin_resultado)} sin cantidades legibles")
    print(f"Revision: {rev}")
    print(f"Detalle por lote: {folder / 'detalle_lotes.txt'}")
    print(f"Excel: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
