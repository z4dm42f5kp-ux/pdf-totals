"""diagnostico.py - muestra lo que lotes_auto.py lee y detecta en uno o varios PDFs.

Usa exactamente el mismo codigo de lectura que lotes_auto.py (mismo OCR, mismos
filtros), asi que lo que muestra es lo que el programa principal ve. No modifica
ningun archivo.

Uso:
    python diagnostico.py RUTA [--lineas N]

RUTA puede ser un PDF o una carpeta (se revisan todos los PDFs, con subcarpetas).
"""
import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import lotes_auto as la  # noqa: E402


def diagnose(pdf, max_lines):
    print("=" * 70)
    print(f"ARCHIVO: {pdf}")
    if la.OMITIR_NOMBRE.search(pdf.name):
        print("Nota: el nombre contiene 'FALTA REVISION'; lotes_auto.py lo omite.")
    trace = []
    try:
        found, lines, source = la.read_pdf(pdf, trace)
    except la.OCRNoDisponible as e:
        print("Resultado: PDF escaneado y el OCR no esta disponible en este sistema.")
        print(f"Detalle: {e}")
        print("Instalar OCR: " + la.ocr_install_hint())
        return
    except Exception as e:
        print(f"Resultado: error al leer el PDF: {e}")
        return

    print(f"Fuente del texto: {source}")
    print("-" * 70)
    shown = lines[:max_lines]
    print(f"TEXTO LEIDO ({len(shown)} de {len(lines)} renglones):")
    for line in shown:
        print("  | " + line)
    if not lines:
        print("  | (nada: no se pudo leer texto)")
    print("-" * 70)
    print("DETECTADO:")
    for note in trace:
        print("  - " + note)
    print("-" * 70)
    if found:
        print(f"RESULTADO: {len(found)} cantidades")
        for lote, v in found:
            print(f"  lote {lote}   cantidad {la.fmt(v):,}")
    else:
        print("RESULTADO: 0 cantidades (este PDF iria a revisar_a_mano.txt)")


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="diagnostico.py",
        description="Muestra lo que lotes_auto.py lee y detecta en uno o varios PDFs.")
    ap.add_argument("ruta", help="un PDF o una carpeta con PDFs")
    ap.add_argument("--lineas", type=int, default=40, metavar="N",
                    help="maximo de renglones de texto a mostrar por PDF (por omision: 40)")
    ap.add_argument("--version", action="version", version=f"%(prog)s {la.VERSION}")
    args = ap.parse_args(argv)

    path = pathlib.Path(args.ruta).expanduser()
    if path.is_file():
        pdfs = [path]
    elif path.is_dir():
        pdfs = sorted(p for p in path.rglob("*") if p.suffix.lower() == ".pdf")
    else:
        print(f"*** No existe: {path}")
        return 2
    print(f"PDFs a revisar: {len(pdfs)}")
    for pdf in pdfs:
        diagnose(pdf, args.lineas)
    return 0


if __name__ == "__main__":
    sys.exit(main())
