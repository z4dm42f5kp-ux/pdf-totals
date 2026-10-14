import re, sys, pathlib
import pdfplumber

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
            print(">> SIN TEXTO: probablemente es un PDF escaneado (imagen).")
            continue
        print(f"Tablas detectadas en pagina 1: {len(page.extract_tables())}")
        print("-" * 60)
        print("PRIMERAS 40 LINEAS DE LA PAGINA 1:")
        for line in text.splitlines()[:40]:
            print("  | " + line)
        print("-" * 60)
        print("LINEAS CON LOTE / SERIE / CANTIDAD / PIEZAS (todas las paginas, max 20):")
        hits = 0
        for pnum, pg in enumerate(doc.pages, 1):
            for line in (pg.extract_text() or "").splitlines():
                if re.search(r"lote|serie|cant|pieza|pza|l\s*o\s*t\s*e", line, re.I):
                    print(f"  p{pnum} | {line}")
                    hits += 1
                    if hits >= 20:
                        break
            if hits >= 20:
                break
        if hits == 0:
            print("  (ninguna)")
