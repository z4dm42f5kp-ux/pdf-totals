import argparse, os, re, sys, pathlib
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font

# Suma la cantidad de piezas por numero de lote en todos los PDFs de una carpeta.
# Entiende tres formatos comunes:
#   1) Tabla con columnas  LOTE | CANTIDAD   (encabezado arriba, numeros abajo)
#   2) Tabla con columna CANTIDAD y el lote dentro de la descripcion ("... LOTE: 4567 ...")
#   3) Texto con etiquetas:  LOTE: 4567  /  CANTIDAD EN PIEZAS: 1,250
# Los PDFs escaneados se leen con OCR (si rapidocr_onnxruntime esta instalado).

VERSION = "1.0.0"

# ===== Etiquetas (puedes agregar palabras aqui) =====
LOTE_WORDS = r"l[o0]te|serial|serie|n/s|s/n"
CANT_WORDS = r"cantidad|cant|piezas|pieza|pzas|pza|pzs|pz|unidades|unid|uds"
# En formatos "de lado" (cada columna es un producto) puede haber varios renglones de
# cantidad (Recibida, Facturada...). Se usa el primero de esta lista que tenga numeros:
PRIORIDAD_CANTIDAD = ["pieza", "devuel", "frasco", "ampollet", "ampula", "vial", "recibid", "cantidad", "factur"]
# Una sola cantidad mayor a esto se avisa en pantalla como sospechosa (igual se suma):
AVISO_CANTIDAD_GRANDE = 50000
# PDFs cuyo nombre contiene esto se omiten (no cuentan en el total):
OMITIR_NOMBRE = re.compile(r"falta[\s_\-]*revisi[oó]n", re.I)
# =====================================================

# El OCR a veces quita espacios ("CANTIDADENPIEZAS:1,250"), por eso no se exigen.
LOTE_RE = re.compile(
    r"(?:" + LOTE_WORDS + r")\.?\s*(?:n[º°o]\.?|n[uú]m\.?|n[uú]mero)?\s*[:#\-]?\s*"
    r"(?P<id>(?=[\w\-/.]*\d)[A-Za-z0-9](?:(?!" + CANT_WORDS + r")[A-Za-z0-9\-/.])*)", re.I)
Q = r"(?:" + CANT_WORDS + r")\.?"
CANT_RE = re.compile(
    Q + r"(?:\s*(?:/|en|de)?\s*" + Q + r")*\s*(?:\(\w*\))?\s*[:=\-]?\s*(?P<n>\d[\d.,]*)", re.I)
# Encabezados de columna (sin "unidad", que es la unidad de medida, no la cantidad)
CANT_HDR = re.compile(r"cantidad|cant(?![a-z])|piezas|pzas|pzs|unidades|(?<![a-z])uds|(?<![a-z])pz(?![a-z])"
                      r"|devuelt[oa]s?|frascos|ampolletas|[aá]mpulas|viales", re.I)
LOTE_HDR = re.compile(LOTE_WORDS, re.I)
LEAD_NUM = re.compile(r"\d[\d.,]*")

# ===== Formato del lote: exactamente 8 digitos que empiezan con 24, 25 o 26 =====
LOTE_ID = re.compile(r"(?<![\d.,])(2[456]\d{6})(?![\d]|[.,]\d)")

def valid_lote(text):
    """Devuelve el primer lote valido (8 digitos, empieza con 24/25/26) o None."""
    m = LOTE_ID.search(text or "")
    return m.group(1) if m else None


def to_num(s):
    # Formato mexicano: 1,250 = 1250 ; 1,250.50 = 1250.5
    s = s.strip().rstrip(".,")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def fmt(v):
    return int(v) if float(v).is_integer() else round(v, 2)


def norm(s):
    return s.strip().strip(".,;:-/").upper()


def sort_key(s):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", s)]


# ---------- Lectura de palabras con su posicion ----------
_ocr = None


class OCRNoDisponible(Exception):
    """El PDF es escaneado y el OCR no esta instalado o no carga en este sistema."""

def load_ocr():
    """Carga el OCR. Usa rapidocr_onnxruntime (Python <= 3.12) o, si no esta,
    rapidocr 3.x + onnxruntime (Python 3.13 y posteriores). Devuelve una funcion
    imagen -> [(caja, texto, confianza)]."""
    errores = []
    try:
        from rapidocr_onnxruntime import RapidOCR
        eng = RapidOCR()
        return lambda arr: (eng(arr)[0] or [])
    except Exception as e:              # ImportError, o "DLL load failed" en Windows
        errores.append(f"rapidocr_onnxruntime: {e}")
    try:
        import logging
        from rapidocr import RapidOCR
        logging.getLogger("RapidOCR").setLevel(logging.WARNING)
        eng = RapidOCR()

        def run(arr):
            r = eng(arr)
            if r.boxes is None:
                return []
            return list(zip(r.boxes.tolist(), r.txts, r.scores))
        return run
    except Exception as e:
        errores.append(f"rapidocr: {e}")
    raise OCRNoDisponible("; ".join(errores))


def page_tokens(page):
    """Devuelve [(x0, x1, y_centro, alto, texto)] y la fuente ('texto' u 'OCR')."""
    if page.chars:
        words = page.extract_words(x_tolerance=1.5, y_tolerance=2)
        return [(w["x0"], w["x1"], (w["top"] + w["bottom"]) / 2, w["bottom"] - w["top"], w["text"]) for w in words], "texto"
    global _ocr
    if _ocr is None:
        _ocr = load_ocr()
    import numpy as np
    try:
        img = page.to_image(resolution=300).original.convert("RGB")
    except ImportError as e:            # sin pypdfium2 (p. ej. en BSD) no se puede rasterizar
        raise OCRNoDisponible(str(e)) from e
    toks = []
    for box, text, score in _ocr(np.array(img)):
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        toks.append((min(xs), max(xs), (min(ys) + max(ys)) / 2, max(ys) - min(ys), text))
    return toks, "OCR"


def group_rows(tokens):
    """Agrupa palabras en renglones segun su altura en la pagina."""
    rows = []
    for t in sorted(tokens, key=lambda t: t[2]):
        if rows and abs(t[2] - rows[-1]["y"]) < max(t[3], rows[-1]["h"]) * 0.6:
            rows[-1]["toks"].append(t)
        else:
            rows.append({"y": t[2], "h": t[3], "toks": [t]})
    for r in rows:
        r["toks"].sort(key=lambda t: t[0])
        r["text"] = " ".join(t[4] for t in r["toks"])
    return rows


def in_column(row, col, need_digit=True):
    """Palabra del renglon que cae debajo de la columna 'col' (la mas cercana)."""
    x0, x1 = col[0], col[1]
    pad = max((x1 - x0) * 0.75, 15)
    cx = (x0 + x1) / 2
    best = None
    for t in row["toks"]:
        tc = (t[0] + t[1]) / 2
        if x0 - pad <= tc <= x1 + pad and (not need_digit or re.search(r"\d", t[4])):
            d = abs(tc - cx)
            if best is None or d < best[0]:
                best = (d, t)
    return best[1] if best else None


def explode(tok):
    """Separa una caja de texto con varias palabras, estimando la posicion de cada una."""
    x0, x1, y, h, text = tok
    out, width = [], max(len(text), 1)
    for m in re.finditer(r"\S+", text):
        a = x0 + (x1 - x0) * m.start() / width
        b = x0 + (x1 - x0) * m.end() / width
        out.append((a, b, m.group()))
    return out


def label_and_values(row):
    """Etiqueta del renglon (palabras iniciales sin numeros) y el resto de palabras."""
    words = [w for t in row["toks"] for w in explode(t)]
    i = 0
    while i < len(words) and not re.search(r"\d", words[i][2]):
        i += 1
    return " ".join(w[2] for w in words[:i]), words[i:]


def parse_sideways(rows, trace=None):
    """Formato de lado: renglon 'Lote' con un lote por columna y renglones de cantidad
    con el valor de cada columna debajo/alineado. 'trace' (lista) recibe notas para
    diagnostico.py."""
    found = []
    for i, row in enumerate(rows):
        label, vals = label_and_values(row)
        if not LOTE_HDR.search(label):
            continue
        cols = [((a + b) / 2, valid_lote(w)) for a, b, w in vals if valid_lote(w)]
        if len(cols) < 2:
            continue
        xs = sorted(c[0] for c in cols)
        gaps = [b - a for a, b in zip(xs, xs[1:])]
        tol = sorted(gaps)[len(gaps) // 2] * 0.5
        best = None
        for r2 in rows[max(0, i - 15):i + 25]:
            if r2 is row:
                continue
            lab2, vals2 = label_and_values(r2)
            if not CANT_HDR.search(lab2):
                continue
            prio = next((k for k, w in enumerate(PRIORIDAD_CANTIDAD) if w in lab2.lower()), len(PRIORIDAD_CANTIDAD))
            pairs = []
            for a, b, w in vals2:
                m = re.fullmatch(r"\(?(\d[\d.,]*)\)?", w)
                v = to_num(m.group(1)) if m else None
                if v is None:
                    continue
                cx = (a + b) / 2
                col = min(cols, key=lambda c: abs(c[0] - cx))
                if abs(col[0] - cx) <= tol:
                    pairs.append((col[1], v))
            if pairs and (best is None or prio < best[0]):
                best = (prio, pairs, lab2)
        if trace is not None:
            trace.append(f"de lado: renglon '{label}' con lotes {[c[1] for c in cols]}"
                         + (f"; cantidades del renglon '{best[2]}'" if best else "; sin renglon de cantidad con numeros"))
        if best:
            found += best[1]
    return found


def header_cells(row):
    """Une palabras muy juntas del encabezado en una sola celda: 'No.' 'de' 'piezas'
    -> 'No. de piezas', para que la columna cubra todo el titulo."""
    cells = []
    for t in row["toks"]:
        if cells and t[0] - cells[-1][1] < max(t[3], cells[-1][3]) * 0.6:
            c = cells[-1]
            cells[-1] = (c[0], t[1], c[2], max(c[3], t[3]), c[4] + " " + t[4])
        else:
            cells.append(t)
    return cells


def header_prio(rows, i, tok, toks=None):
    """Prioridad de una columna de cantidad segun su encabezado completo (incluye
    palabras justo arriba/abajo, p. ej. 'Cantidad' / 'Recibida' en dos renglones,
    y la palabra siguiente en el mismo renglon, p. ej. 'Piezas Devueltas')."""
    x0, x1 = tok[0], tok[1]
    pad = max((x1 - x0) * 0.3, 5)
    ctx = [tok[4]]
    toks = toks or rows[i]["toks"]
    k = toks.index(tok)
    if k + 1 < len(toks) and not CANT_HDR.search(toks[k + 1][4]):
        ctx.append(toks[k + 1][4])
    for j in (i - 1, i + 1):
        if 0 <= j < len(rows) and not re.search(r"\d", rows[j]["text"]):
            ctx += [t[4] for t in rows[j]["toks"] if x0 - pad <= (t[0] + t[1]) / 2 <= x1 + pad]
    ctx = " ".join(ctx).lower()
    return next((n for n, w in enumerate(PRIORIDAD_CANTIDAD) if w in ctx), len(PRIORIDAD_CANTIDAD))


def parse_table(rows, trace=None):
    """Formatos 1 y 2: encabezado CANTIDAD (y quiza LOTE) con valores debajo.
    Si hay varias columnas de cantidad se usa la de mayor prioridad (PRIORIDAD_CANTIDAD)."""
    found = []
    qcol = lcol = None
    qprio, last_hdr = 99, -9
    for i, row in enumerate(rows):
        # Renglon de encabezado: tiene una palabra de cantidad (p. ej. "No. de piezas"),
        # no es "PIEZAS: 30" (etiqueta con valor) y no contiene lotes. Puede tener otros
        # numeros sueltos como "Hoja 1 de 2" o "Folio: 2".
        if (not CANT_RE.search(row["text"]) and not valid_lote(row["text"])
                and sum(bool(LEAD_NUM.fullmatch(t[4])) for t in row["toks"]) <= 2):
            cells = header_cells(row)
            qs = [t for t in cells if CANT_HDR.search(t[4])]
            q = min(qs, key=lambda t: (header_prio(rows, i, t, cells), t[0])) if qs else None
            if q:
                p = header_prio(rows, i, q, cells)
                # Segundo renglon del mismo encabezado (p. ej. "PIEZAS" / "DEVUELTAS"):
                # solo cambia la columna si apunta a una claramente mejor.
                if qcol is not None and i == last_hdr + 1 and p >= qprio:
                    last_hdr = i
                    continue
                qcol, qprio, last_hdr = q, p, i
                lcol = next((t for t in cells if LOTE_HDR.search(t[4]) and t is not q), None)
                if trace is not None:
                    trace.append(f"tabla: columna de cantidad '{q[4]}' (prioridad {p})"
                                 + (f", columna de lote '{lcol[4]}'" if lcol else ", lote buscado en el renglon"))
                continue
        if not qcol:
            continue
        qt = in_column(row, qcol)
        if not qt:
            continue
        m = LEAD_NUM.search(qt[4])
        v = to_num(m.group()) if m else None
        if v is None:
            continue
        if valid_lote(qt[4]):
            continue                      # el "numero" de cantidad es en realidad un lote
        lote = None
        if lcol:
            lt = in_column(row, lcol)
            if lt and lt is not qt:
                lote = valid_lote(lt[4])
        if not lote:
            # lote en otra parte del renglon (p. ej. en la descripcion) o en los 2 siguientes
            for r2 in [row] + rows[i + 1:i + 3]:
                if r2 is not row and in_column(r2, qcol):
                    break
                lote = valid_lote(r2["text"])
                if lote:
                    break
        if lote:
            found.append((lote, v))
            if trace is not None:
                trace.append(f"tabla: renglon -> lote {lote}, cantidad {fmt(v)}")
    return found


def inline_events(rows):
    """Formato 3: etiquetas en el mismo renglon (LOTE: x / CANTIDAD: y)."""
    ev = []
    for row in rows:
        line = row["text"]
        e = [(m.start(), "L", valid_lote(m.group("id"))) for m in LOTE_RE.finditer(line)]
        e += [(m.start(), "C", to_num(m.group("n"))) for m in CANT_RE.finditer(line)]
        ev += [(k, v) for _, k, v in sorted(e) if v is not None]
    return ev


def pair_events(ev):
    """Empareja lotes y cantidades en el orden en que aparecen en el documento."""
    out = []
    if not ev:
        return out
    first = ev[0][0]                      # 'L' = lote antes de cantidad, 'C' = al reves
    pend = None
    for kind, val in ev:
        if kind == first:
            pend = val
        elif pend is not None:
            lote, v = (pend, val) if first == "L" else (val, pend)
            out.append((lote, v))
            pend = None
    return out


def read_pdf(path, trace=None):
    """Lee un PDF y devuelve (encontrados, lineas_leidas, fuente).
    encontrados = [(lote, cantidad), ...] ya filtrados (lote valido, cantidad entera > 0).
    'trace' (lista) recibe notas de lo detectado; lo usa diagnostico.py."""
    found, ev, leidas, fuente = [], [], [], ""
    with pdfplumber.open(path) as doc:
        for pnum, page in enumerate(doc.pages, 1):
            toks, fuente = page_tokens(page)
            rows = group_rows(toks)
            leidas += [r["text"] for r in rows]
            if trace is not None:
                trace.append(f"pagina {pnum}: {len(rows)} renglones leidos ({fuente})")
            t = parse_sideways(rows, trace) or parse_table(rows, trace)
            if t:
                found += t
            else:
                e = inline_events(rows)
                if trace is not None and e:
                    trace.append(f"pagina {pnum}: etiquetas en linea: "
                                 + ", ".join(f"{'lote' if k == 'L' else 'cantidad'} {v if k == 'L' else fmt(v)}"
                                             for k, v in e))
                ev += e
    found += pair_events(ev)
    # Ultimo filtro: lote valido y cantidad entera (las piezas no llevan decimales)
    checked = [(valid_lote(l), l, v) for l, v in found]
    if trace is not None:
        for ok, l, v in checked:
            if not ok:
                trace.append(f"descartado: '{l}' no es un lote valido (8 digitos, empieza con 24/25/26)")
            elif not float(v).is_integer() or v <= 0:
                trace.append(f"descartado: lote {ok} con cantidad {v} (no es entera o no es positiva)")
    found = [(ok, v) for ok, l, v in checked if ok and float(v).is_integer() and v > 0]
    return found, leidas, fuente


def ocr_install_hint():
    if sys.version_info >= (3, 13):
        pkgs = "rapidocr onnxruntime"
    else:
        pkgs = "rapidocr_onnxruntime"
    if os.name == "nt":
        pkgs += " msvc-runtime"
    return f'"{sys.executable}" -m pip install {pkgs}'


# ---------- Programa principal ----------
def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="lotes_auto.py",
        description="Suma las piezas por numero de lote en todos los PDFs de una carpeta "
                    "(incluye subcarpetas) y guarda el resultado en totales_lotes.xlsx.")
    ap.add_argument("carpeta", help="carpeta con los PDFs")
    ap.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    args = ap.parse_args(argv)

    folder = pathlib.Path(args.carpeta).expanduser()
    if not folder.is_dir():
        print(f"*** No existe la carpeta: {folder}")
        return 2
    pdfs = sorted(p for p in folder.rglob("*") if p.suffix.lower() == ".pdf")
    print(f"{len(pdfs)} PDFs encontrados en {folder}")

    totals = {}
    n_total = 0
    sin_resultado = []  # [(archivo, motivo)]
    omitidos = []
    detalle = {}        # lote -> [(archivo, cantidad)]
    grandes = []
    log = []            # texto leido de los PDFs sin resultado
    aviso_ocr = False
    for n, pdf in enumerate(pdfs, 1):
        name = str(pdf.relative_to(folder))
        print(f"[{n}/{len(pdfs)}] {name}", end="", flush=True)
        if OMITIR_NOMBRE.search(pdf.name):
            omitidos.append(name)
            print("  -> omitido (FALTA REVISION)")
            continue
        try:
            found, leidas, fuente = read_pdf(pdf)
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

        for lote, v in found:
            totals[lote] = totals.get(lote, 0) + v
            detalle.setdefault(lote, []).append((name, v))
            if v > AVISO_CANTIDAD_GRANDE:
                grandes.append((name, lote, v))
        n_total += len(found)
        print(f"  -> {len(found)} cantidades ({fuente})")
        if not found:
            sin_resultado.append((name, "sin cantidades legibles"))
            log += ["=" * 70, name, f"Motivo: sin cantidades legibles (fuente: {fuente})",
                    "Texto leido:"]
            log += ["  | " + l for l in leidas] or ["  | (nada: no se pudo leer texto)"]
            log.append("")
            print("     (texto leido guardado en sin_resultado.log)")

    # ----- Excel -----
    wb = Workbook()
    ws = wb.active
    ws.title = "Lotes"
    ws.append(["Lote", "Cantidad en Piezas"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for lote in sorted(totals, key=sort_key):
        ws.append([lote, fmt(totals[lote])])
        ws.cell(row=ws.max_row, column=2).number_format = "#,##0"   # 1,406 con coma de miles
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 20

    out = folder / "totales_lotes.xlsx"
    try:
        wb.save(out)
    except PermissionError:
        print(f"\n*** No se pudo guardar {out.name}: esta abierto en Excel. Cierralo y vuelve a correr el script.")
        return 1

    print(f"\n{len(totals)} lotes | {n_total} cantidades leidas")

    # ----- Lista para revisar a mano -----
    rev = folder / "revisar_a_mano.txt"
    if omitidos or sin_resultado:
        txt = [f"PDFs YA MARCADOS 'FALTA REVISION' (omitidos, no cuentan en el total): {len(omitidos)}"]
        txt += ["   " + nm for nm in omitidos]
        txt += ["", f"PDFs SIN CANTIDADES LEGIBLES (revisar a mano): {len(sin_resultado)}"]
        txt += [f"   {nm}   [{motivo}]" for nm, motivo in sin_resultado]
        rev.write_text("\n".join(txt) + "\n", encoding="utf-8")
        print(f"{len(omitidos)} omitidos por FALTA REVISION | {len(sin_resultado)} sin cantidades legibles")
        print(f"Lista: {rev}")
    elif rev.exists():
        rev.unlink()

    # ----- Log con el texto leido de los PDFs sin resultado -----
    logf = folder / "sin_resultado.log"
    if log:
        logf.write_text("\n".join(log) + "\n", encoding="utf-8")
        print(f"Texto leido de esos PDFs: {logf}")
    elif logf.exists():
        logf.unlink()

    # ----- Detalle: de que PDF salio cada cantidad de cada lote -----
    det = folder / "detalle_lotes.txt"
    lines = []
    for lote in sorted(detalle, key=sort_key):
        lines.append(f"{lote}   TOTAL = {fmt(totals[lote]):,}")
        for nm, v in detalle[lote]:
            lines.append(f"      {fmt(v):>10,}   {nm}")
        lines.append("")
    det.write_text("\n".join(lines), encoding="utf-8")
    if grandes:
        print(f"\n*** {len(grandes)} cantidades sospechosamente grandes (> {AVISO_CANTIDAD_GRANDE:,}), revisalas:")
        for nm, lote, v in grandes:
            print(f"   {fmt(v):>12,}  lote {lote}  en  {nm}")
    print(f"Detalle por lote: {det}")
    print(f"Excel: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
