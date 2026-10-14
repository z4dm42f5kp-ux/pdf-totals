import re, sys, pathlib
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font

# ===== Etiquetas del numero de lote (lote y serie se tratan igual) =====
LOTE_WORDS = r"l[o0]te|serial|serie|n/s|s/n"
# ===== Etiquetas de la cantidad =====
CANT_WORDS = r"cantidad|cant|piezas|pieza|pzas|pza|pzs|pz|unidades|unid|uds"

# Nota: el OCR a veces quita los espacios ("CANTIDADENPIEZAS:1,250", "NODELOTE:4567"),
# por eso las etiquetas se reconocen aunque esten pegadas a otras palabras.
LOTE_RE = re.compile(
    r"(?:" + LOTE_WORDS + r")\.?\s*(?:n[º°o]\.?|n[uú]m\.?|n[uú]mero)?\s*[:#\-]?\s*"
    r"(?P<id>(?=[\w\-/.]*\d)[A-Za-z0-9](?:(?!" + CANT_WORDS + r")[A-Za-z0-9\-/.])*)", re.I)
Q = r"(?:" + CANT_WORDS + r")\.?"
CANT_RE = re.compile(
    Q + r"(?:\s*(?:/|en|de)?\s*" + Q + r")*\s*(?:\(\w*\))?\s*[:=\-]?\s*(?P<n>\d[\d.,]*)", re.I)
LOTE_WORD_RE = re.compile(r"(?:" + LOTE_WORDS + r")", re.I)

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

_ocr = None

def ocr_lines(page):
    # Lee una pagina escaneada (imagen) con OCR y devuelve sus lineas de texto
    global _ocr
    import numpy as np
    if _ocr is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr = RapidOCR()
    img = page.to_image(resolution=300).original.convert("RGB")
    res, _ = _ocr(np.array(img))
    if not res:
        return []
    items = []
    for box, text, score in res:
        ys = [p[1] for p in box]
        xs = [p[0] for p in box]
        items.append(((min(ys) + max(ys)) / 2, max(ys) - min(ys), min(xs), text))
    items.sort()
    lines = []
    for yc, h, x, t in items:
        if lines and abs(yc - lines[-1]["y"]) < max(h, lines[-1]["h"]) * 0.6:
            lines[-1]["parts"].append((x, t))
        else:
            lines.append({"y": yc, "h": h, "parts": [(x, t)]})
    return [" ".join(t for x, t in sorted(l["parts"])) for l in lines]

def page_lines(page):
    # Usa el texto del PDF si lo tiene; si es escaneado, usa OCR
    if page.chars:
        return (page.extract_text() or "").splitlines(), "texto"
    return ocr_lines(page), "OCR"

def sort_key(s):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", s)]

totals, detail, avisos = {}, [], []

folder = pathlib.Path(sys.argv[1]).expanduser()
pdfs = sorted(p for p in folder.rglob("*") if p.suffix.lower() == ".pdf")
print(f"{len(pdfs)} PDFs encontrados en {folder}")

for n, pdf in enumerate(pdfs, 1):
    name = str(pdf.relative_to(folder))
    print(f"[{n}/{len(pdfs)}] {name}", end="", flush=True)
    antes = len(detail)
    try:
        with pdfplumber.open(pdf) as doc:
            current, cur_page = None, None
            for pnum, page in enumerate(doc.pages, 1):
                lines, fuente = page_lines(page)
                if not lines:
                    avisos.append([name, pnum, "", "Pagina sin texto legible (ni con OCR)"])
                for line in lines:
                    events = [(m.start(), "L", m) for m in LOTE_RE.finditer(line)]
                    events += [(m.start(), "C", m) for m in CANT_RE.finditer(line)]
                    if not any(e[1] == "L" for e in events) and LOTE_WORD_RE.search(line):
                        avisos.append([name, pnum, "", "Revisar, lote no reconocido: " + line.strip()[:120]])
                    for _, kind, m in sorted(events, key=lambda e: e[0]):
                        if kind == "L":
                            if current:
                                avisos.append([name, cur_page, current, "Lote sin cantidad"])
                            current, cur_page = norm(m.group("id")), pnum
                        elif current:
                            v = to_num(m.group("n"))
                            if v is not None:
                                totals[current] = totals.get(current, 0) + v
                                detail.append([name, pnum, current, fmt(v), fuente])
                                current = None
                        else:
                            avisos.append([name, pnum, "", "Cantidad sin lote: " + line.strip()[:120]])
            if current:
                avisos.append([name, cur_page, current, "Lote sin cantidad"])
        print(f"  -> {len(detail) - antes} cantidades")
    except Exception as e:
        print(f"   ERROR en {name}: {e}")
        avisos.append([name, "", "", f"Error: {e}"])
        if "ocr" in str(e).lower() or "dll" in str(e).lower() or isinstance(e, ImportError):
            print("\n*** El OCR no pudo cargarse. Ejecuta:")
            print("    & $py -m pip install rapidocr_onnxruntime msvc-runtime")
            print("    y vuelve a correr el script.")
            sys.exit(1)

bold = Font(bold=True)
wb = Workbook()
ws = wb.active
ws.title = "Lotes"
ws.append(["Lote", "Cantidad en Piezas"])
for lote in sorted(totals, key=sort_key):
    ws.append([lote, fmt(totals[lote])])
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 20

ws2 = wb.create_sheet("Detalle")
ws2.append(["Archivo", "Pagina", "Lote", "Cantidad", "Fuente"])
for r in detail:
    ws2.append(r)

if avisos:
    ws3 = wb.create_sheet("Avisos")
    ws3.append(["Archivo", "Pagina", "Lote", "Aviso"])
    for r in avisos:
        ws3.append(r)

for sh in wb.worksheets:
    for c in sh[1]:
        c.font = bold
    sh.freeze_panes = "A2"

out = folder / "totales_lotes.xlsx"
try:
    wb.save(out)
except PermissionError:
    out = folder / "totales_lotes_nuevo.xlsx"
    wb.save(out)

md = ["| Lote | Cantidad en Piezas |", "|---|---|"]
md += [f"| {l} | {fmt(totals[l])} |" for l in sorted(totals, key=sort_key)]
(folder / "totales_lotes.md").write_text("\n".join(md), encoding="utf-8")

print(f"\n{len(totals)} lotes | {len(detail)} cantidades leidas | {len(avisos)} avisos")
print(f"Excel: {out}")
