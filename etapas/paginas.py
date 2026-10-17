"""paginas.py - read the pages of PDFs and print one JSON line per page.

    python etapas/paginas.py RUTA... [--lectura 1|2|3] [--procesos N] [--cache DIR]

RUTA is a PDF or a folder (all PDFs in it and in its subfolders, except the folder
resultados_lotes). Each output line is a record "pagina" with the words of the page and
their positions, read from the text of the PDF or with OCR. A page that cannot be read
gives a record "error". The pages are read in parallel, like lotes_auto.py does, and
the same cache is used: by default resultados_lotes/cache in the first folder given.

This stage only reads. It does not decide anything about batches or quantities. Pipe
its output into etapas/cantidades.py, or into etapas/texto.py to see the text.
"""
import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from _comun import escribir, la  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(prog="paginas.py", description=__doc__.split("\n\n")[0])
    ap.add_argument("rutas", nargs="+", metavar="RUTA", help="PDF o carpeta")
    ap.add_argument("--lectura", type=int, choices=(1, 2, 3), default=1,
                    help="lectura del OCR: 1 normal, 2 mas resolucion y contraste, 3 ampliada y enfocada")
    ap.add_argument("--procesos", type=int, default=None, metavar="N",
                    help="procesos que leen paginas a la vez (por omision: automatico)")
    ap.add_argument("--cache", metavar="DIR", default=None,
                    help="carpeta de la cache (por omision: resultados_lotes/cache de la primera carpeta)")
    ap.add_argument("--sin-cache", action="store_true", help="no usar ninguna cache")
    ap.add_argument("--version", action="version", version=f"%(prog)s {la.VERSION}")
    args = ap.parse_args(argv)

    pdfs, cache = [], None
    for r in args.rutas:
        p = pathlib.Path(r).expanduser()
        if p.is_dir():
            salida = p / la.CARPETA_RESULTADOS
            pdfs += [(str(x.relative_to(p)), x) for x in sorted(p.rglob("*"))
                     if x.suffix.lower() == ".pdf" and salida not in x.parents]
            cache = cache or salida / "cache"
        elif p.is_file():
            pdfs.append((p.name, p))
        else:
            sys.exit(f"*** No existe: {p}")
    if args.cache:
        cache = pathlib.Path(args.cache).expanduser()
    if args.sin_cache:
        cache = None
    lector = la.LectorPaginas(cache, args.procesos or la.PROCESOS)
    trabajos, npags = [], {}
    for nombre, pdf in pdfs:
        try:
            huella = la.huella_pdf(pdf)
            with la.pdfplumber.open(pdf) as doc:
                npags[nombre] = len(doc.pages)
        except Exception as e:
            escribir({"tipo": "error", "archivo": nombre, "pagina": None, "error": str(e)})
            continue
        trabajos += [(nombre, pdf, huella, n, args.lectura) for n in range(1, npags[nombre] + 1)]
    lector.npags = npags
    leidas = {}
    pendientes = [n for n, _ in pdfs if n in npags]

    def terminado(nombre, _seg):         # escribir los PDFs en orden, en cuanto se pueda
        while pendientes and all((pendientes[0], n, args.lectura) in leidas
                                 for n in range(1, npags[pendientes[0]] + 1)):
            f = pendientes.pop(0)
            ruta = str(dict(pdfs)[f])
            for n in range(1, npags[f] + 1):
                valor = leidas.pop((f, n, args.lectura))
                if isinstance(valor, Exception):
                    escribir({"tipo": "error", "archivo": f, "pagina": n, "error": str(valor)})
                    continue
                toks, fuente = valor
                escribir({"tipo": "pagina", "archivo": f, "ruta": ruta, "pagina": n, "paginas": npags[f],
                          "lectura": args.lectura, "fuente": fuente, "toks": [list(t) for t in toks]})
            sys.stdout.flush()

    try:
        lector.leer(trabajos, terminado, leidas)
    finally:
        lector.cerrar()
    terminado(None, 0)
    print(f"{len(trabajos)} pagina(s): {lector.leidas} leidas, {lector.de_cache} de la cache", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
