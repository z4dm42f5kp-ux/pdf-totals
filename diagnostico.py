"""diagnostico.py - muestra lo que lotes_auto.py lee y detecta en uno o varios PDFs.

Usa exactamente el mismo codigo de lectura que lotes_auto.py (mismo OCR, mismas
reglas), asi que lo que muestra es lo que el programa principal ve: el texto leido,
que tabla y que columna eligio y por que, el tipo de cada cantidad (retirado o no
distribuida) y las comprobaciones que supero. No modifica ningun archivo.

Uso:
    python diagnostico.py RUTA [--lineas N] [--lectura 1|2|3]

RUTA puede ser un PDF o una carpeta (se revisan todos los PDFs, con subcarpetas).
"""
import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import lotes_auto as la  # noqa: E402


def diagnose(pdf, max_lines, pasada):
    print("=" * 70)
    print(f"ARCHIVO: {pdf}")
    if la.OMITIR_NOMBRE.search(pdf.name):
        print("Nota: el nombre contiene 'FALTA REVISION'; lotes_auto.py lo omite.")
    try:
        L = la.leer_pdf(pdf, pdf.name, pasada=pasada)
    except la.OCRNoDisponible as e:
        print("Resultado: PDF escaneado y el OCR no esta disponible en este sistema.")
        print(f"Detalle: {e}")
        print("Instalar OCR: " + la.ocr_install_hint())
        return
    except Exception as e:
        print(f"Resultado: error al leer el PDF: {e}")
        return

    print(f"Fuente del texto: {L.fuente}   (lectura {pasada})")
    print("-" * 70)
    shown = L.lineas[:max_lines]
    print(f"TEXTO LEIDO ({len(shown)} de {len(L.lineas)} renglones):")
    for line in shown:
        print("  | " + line)
    if not L.lineas:
        print("  | (nada: no se pudo leer texto)")
    print("-" * 70)
    print("DETECTADO:")
    for note in L.traza:
        print("  - " + note)
    for aviso in L.avisos:
        print("  ! " + aviso)
    print("-" * 70)
    activas = [e for e in L.entradas if not e.excluido]
    print(f"RESULTADO: {len(activas)} cantidades" + (f", reporte con {len(L.referencias)} lotes"
                                                     if L.referencias else ""))
    for e in L.entradas:
        estado = (f"NO SE SUMA: {e.excluido}" if e.excluido else
                  "verificado" if e.verificado else f"sin verificar, confianza {e.conf:.2f}")
        print(f"  lote {e.lote}   {la.TIPOS[e.tipo]:<15} {la.fmt(e.piezas):>9}   p{e.pagina} [{e.forma}; {estado}]")
        for x in e.verificado + e.fallas + e.notas:
            print(f"        - {x}")
    for r in L.referencias:
        val = lambda v: la.fmt(v) if v is not None else "-"
        print(f"  reporte: lote {r.lote}   no distribuida {val(r.nod)}   retirado {val(r.ret)}   "
              f"total {val(r.total)}" + ("" if r.cuadra else "   (la fila no cuadra)"))
    if not L.entradas and not L.referencias:
        print("  (este PDF iria a revisar_a_mano.txt)")


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="diagnostico.py",
        description="Muestra lo que lotes_auto.py lee y detecta en uno o varios PDFs.")
    ap.add_argument("ruta", help="un PDF o una carpeta con PDFs")
    ap.add_argument("--lineas", type=int, default=40, metavar="N",
                    help="maximo de renglones de texto a mostrar por PDF (por omision: 40)")
    ap.add_argument("--lectura", type=int, choices=(1, 2, 3), default=1,
                    help="lectura del OCR: 1 normal, 2 mas resolucion y contraste, 3 ampliada y enfocada")
    ap.add_argument("--segunda-lectura", action="store_true", help=argparse.SUPPRESS)
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
        diagnose(pdf, args.lineas, 2 if args.segunda_lectura else args.lectura)
    return 0


if __name__ == "__main__":
    sys.exit(main())
