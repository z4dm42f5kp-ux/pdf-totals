import re, sys, pathlib
import pdfplumber

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


# Muestra lo que el script "ve" dentro de los PDFs, para encontrar por que no detecta lotes.
folder = pathlib.Path(sys.argv[1]).expanduser()
pdfs = sorted(p for p in folder.rglob("*") if p.suffix.lower() == ".pdf")
print(f"PDFs encontrados: {len(pdfs)}")
if not pdfs:
    print("No hay PDFs en esa carpeta. Revisa la ruta.")
    sys.exit()

for pdf in pdfs[:2]:
    print("\n" + "=" * 60)
    print(f"ARCHIVO: {pdf.name}")
    with pdfplumber.open(pdf) as doc:
        print(f"Paginas: {len(doc.pages)}")
        page = doc.pages[0]
        text = page.extract_text() or ""
        print(f"Caracteres de texto en pagina 1: {len(page.chars)}")
        if not page.chars:
            print(">> PDF escaneado (imagen): leyendo con OCR...")
            text = "\n".join(ocr_lines(page))
        print("-" * 60)
        print("PRIMERAS 40 LINEAS DE LA PAGINA 1:")
        for line in text.splitlines()[:40]:
            print("  | " + line)
        print("-" * 60)
        print("LINEAS CON LOTE / SERIE / CANTIDAD / PIEZAS (todas las paginas, max 20):")
        hits = 0
        for pnum, pg in enumerate(doc.pages, 1):
            lines = (pg.extract_text() or "").splitlines() if pg.chars else ocr_lines(pg)
            for line in lines:
                if re.search(r"lote|serie|cant|pieza|pza|l\s*o\s*t\s*e", line, re.I):
                    print(f"  p{pnum} | {line}")
                    hits += 1
                    if hits >= 20:
                        break
            if hits >= 20:
                break
        if hits == 0:
            print("  (ninguna)")
