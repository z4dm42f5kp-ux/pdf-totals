import re, sys, pathlib
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font

# Suma la cantidad de piezas por numero de lote en todos los PDFs de una carpeta.
# Entiende tres formatos comunes:
#   1) Tabla con columnas  LOTE | CANTIDAD   (encabezado arriba, numeros abajo)
#   2) Tabla con columna CANTIDAD y el lote dentro de la descripcion ("... LOTE: 4567 ...")
#   3) Texto con etiquetas:  LOTE: 4567  /  CANTIDAD EN PIEZAS: 1,250
# Los PDFs escaneados se leen con OCR.

# ===== Etiquetas (puedes agregar palabras aqui) =====
LOTE_WORDS = r"l[o0]te|serial|serie|n/s|s/n"
CANT_WORDS = r"cantidad|cant|piezas|pieza|pzas|pza|pzs|pz|unidades|unid|uds"
# En formatos "de lado" (cada columna es un producto) puede haber varios renglones de
# cantidad (Recibida, Facturada...). Se usa el primero de esta lista que tenga numeros:
PRIORIDAD_CANTIDAD = ["pieza", "devuel", "recibid", "cantidad", "factur"]
# =====================================================

# El OCR a veces quita espacios ("CANTIDADENPIEZAS:1,250"), por eso no se exigen.
LOTE_RE = re.compile(
    r"(?:" + LOTE_WORDS + r")\.?\s*(?:n[º°o]\.?|n[uú]m\.?|n[uú]mero)?\s*[:#\-]?\s*"
    r"(?P<id>(?=[\w\-/.]*\d)[A-Za-z0-9](?:(?!" + CANT_WORDS + r")[A-Za-z0-9\-/.])*)", re.I)
Q = r"(?:" + CANT_WORDS + r")\.?"
CANT_RE = re.compile(
    Q + r"(?:\s*(?:/|en|de)?\s*" + Q + r")*\s*(?:\(\w*\))?\s*[:=\-]?\s*(?P<n>\d[\d.,]*)", re.I)
# Encabezados de columna (sin "unidad", que es la unidad de medida, no la cantidad)
CANT_HDR = re.compile(r"cantidad|cant(?![a-z])|piezas|pzas|pzs|unidades|(?<![a-z])uds|(?<![a-z])pz(?![a-z])", re.I)
LOTE_HDR = re.compile(LOTE_WORDS, re.I)
LEAD_NUM = re.compile(r"\d[\d.,]*")


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

def page_tokens(page):
    """Devuelve [(x0, x1, y_centro, alto, texto)] y la fuente ('texto' u 'OCR')."""
    if page.chars:
        words = page.extract_words(x_tolerance=1.5, y_tolerance=2)
        return [(w["x0"], w["x1"], (w["top"] + w["bottom"]) / 2, w["bottom"] - w["top"], w["text"]) for w in words], "texto"
    global _ocr
    import numpy as np
    if _ocr is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr = RapidOCR()
    img = page.to_image(resolution=300).original.convert("RGB")
    res, _ = _ocr(np.array(img))
    toks = []
    for box, text, score in (res or []):
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


def parse_sideways(rows):
    """Formato de lado: renglon 'Lote' con un lote por columna y renglones de cantidad
    con el valor de cada columna debajo/alineado."""
    found = []
    for i, row in enumerate(rows):
        label, vals = label_and_values(row)
        if not LOTE_HDR.search(label):
            continue
        cols = [((a + b) / 2, norm(w)) for a, b, w in vals if re.search(r"\d{3,}", w)]
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
                best = (prio, pairs)
        if best:
            found += best[1]
    return found


def parse_table(rows):
    """Formatos 1 y 2: encabezado CANTIDAD (y quiza LOTE) con valores debajo."""
    found = []
    qcol = lcol = None
    for i, row in enumerate(rows):
        if not re.search(r"\d", row["text"]):
            q = next((t for t in row["toks"] if CANT_HDR.search(t[4])), None)
            if q:
                qcol = q
                lcol = next((t for t in row["toks"] if LOTE_HDR.search(t[4]) and t is not q), None)
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
        lote = None
        if lcol:
            lt = in_column(row, lcol)
            if lt and lt is not qt:
                lm = LOTE_RE.search(lt[4])
                lote = lm.group("id") if lm else lt[4]
        if not lote:
            # lote dentro de la descripcion: este renglon o los 2 siguientes
            for r2 in [row] + rows[i + 1:i + 3]:
                if r2 is not row and in_column(r2, qcol):
                    break
                lm = LOTE_RE.search(r2["text"])
                if lm:
                    lote = lm.group("id")
                    break
        if lote and re.search(r"\d", lote):
            found.append((norm(lote), v))
    return found


def inline_events(rows):
    """Formato 3: etiquetas en el mismo renglon (LOTE: x / CANTIDAD: y)."""
    ev = []
    for row in rows:
        line = row["text"]
        e = [(m.start(), "L", norm(m.group("id"))) for m in LOTE_RE.finditer(line)]
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


# ---------- Programa principal ----------
folder = pathlib.Path(sys.argv[1]).expanduser()
pdfs = sorted(p for p in folder.rglob("*") if p.suffix.lower() == ".pdf")
print(f"{len(pdfs)} PDFs encontrados en {folder}")

totals = {}
n_total = 0
sin_resultado = []
for n, pdf in enumerate(pdfs, 1):
    name = str(pdf.relative_to(folder))
    print(f"[{n}/{len(pdfs)}] {name}", end="", flush=True)
    try:
        found, ev, leidas, fuente = [], [], [], ""
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                toks, fuente = page_tokens(page)
                rows = group_rows(toks)
                leidas += [r["text"] for r in rows]
                t = parse_sideways(rows) or parse_table(rows)
                if t:
                    found += t
                else:
                    ev += inline_events(rows)
        found += pair_events(ev)
        for lote, v in found:
            totals[lote] = totals.get(lote, 0) + v
        n_total += len(found)
        print(f"  -> {len(found)} cantidades ({fuente})")
        if not found:
            sin_resultado.append(name)
            print("     Texto leido de este PDF (primeras 15 lineas):")
            for l in leidas[:15]:
                print("     | " + l)
            if not leidas:
                print("     | (nada: no se pudo leer texto)")
    except Exception as e:
        print(f"  -> ERROR: {e}")
        if isinstance(e, ImportError) or "dll" in str(e).lower():
            print("\n*** El OCR no pudo cargarse. Ejecuta:")
            print("    & $py -m pip install rapidocr_onnxruntime msvc-runtime")
            sys.exit(1)

wb = Workbook()
ws = wb.active
ws.title = "Lotes"
ws.append(["Lote", "Cantidad en Piezas"])
for c in ws[1]:
    c.font = Font(bold=True)
for lote in sorted(totals, key=sort_key):
    ws.append([lote, fmt(totals[lote])])
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 20

out = folder / "totales_lotes.xlsx"
try:
    wb.save(out)
except PermissionError:
    print(f"\n*** No se pudo guardar {out.name}: esta abierto en Excel. Cierralo y vuelve a correr el script.")
    sys.exit(1)

print(f"\n{len(totals)} lotes | {n_total} cantidades leidas")
if sin_resultado:
    rev = folder / "revisar_a_mano.txt"
    rev.write_text("\n".join(sin_resultado), encoding="utf-8")
    print(f"{len(sin_resultado)} PDFs sin cantidades legibles (lista en {rev.name}):")
    for nm in sin_resultado:
        print("   - " + nm)
print(f"Excel: {out}")
