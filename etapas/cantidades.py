"""cantidades.py - find the batches and quantities in page records.

    python etapas/paginas.py FOLDER | python etapas/cantidades.py [--perfil P] [--sector S]

Reads the records "pagina" of etapas/paginas.py (the pages of one PDF must be together,
as paginas.py writes them) and analyses each PDF exactly like lotes_auto.py: tables,
withdrawal reports, labelled lines, forms with one product per column, the checks of
each table and the rules for batches repeated inside the same PDF. It writes one record
"archivo" for each PDF, then its records "cantidad" and "referencia".

Files whose name contains FALTA REVISION are marked with "omitido": true; the next
stage does not count them. Records "error" are passed on unchanged.
"""
import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from _comun import aplicar_perfil, escribir, la, leer, opciones_perfil  # noqa: E402


def salida_documento(nombre, paginas, errores):
    omitido = bool(la.OMITIR_NOMBRE.search(pathlib.Path(nombre).name))
    if errores:
        escribir({"tipo": "archivo", "archivo": nombre, "fuente": "", "paginas": len(paginas) + len(errores),
                  "omitido": omitido, "avisos": [], "error": errores[0]})
        return
    L = la.analizar_documento(nombre, [(p[1], p[2]) for p in sorted(paginas)], 1)
    escribir({"tipo": "archivo", "archivo": nombre, "fuente": L.fuente, "paginas": L.paginas,
              "omitido": omitido, "avisos": L.avisos})
    for e in L.entradas:
        escribir({"tipo": "cantidad", "archivo": nombre, "pagina": e.pagina, "lote": e.lote, "piezas": e.piezas,
                  "clase": e.tipo, "forma": e.forma, "conf": round(e.conf, 4), "verificado": e.verificado,
                  "fallas": e.fallas, "notas": e.notas, "excluido": e.excluido})
    for r in L.referencias:
        escribir({"tipo": "referencia", "archivo": nombre, "pagina": r.pagina, "lote": r.lote, "ret": r.ret,
                  "nod": r.nod, "total": r.total, "cuadra": r.cuadra, "ret_opciones": list(r.ret_opciones)})


def main(argv=None):
    ap = argparse.ArgumentParser(prog="cantidades.py", description=__doc__.split("\n\n")[0])
    opciones_perfil(ap, excel=True)
    args = ap.parse_args(argv)
    aplicar_perfil(args)
    actual, paginas, errores = None, [], []
    for reg in leer():
        if reg["tipo"] not in ("pagina", "error"):
            escribir(reg)                     # registros de otras etapas: pasan sin cambios
            continue
        if reg["archivo"] != actual:
            if actual is not None:
                salida_documento(actual, paginas, errores)
            actual, paginas, errores = reg["archivo"], [], []
        if reg["tipo"] == "error":
            if reg.get("pagina") is None:     # el PDF no se pudo abrir
                escribir(reg)
                actual = None
                continue
            errores.append(f"pagina {reg['pagina']}: {reg['error']}")
        else:
            paginas.append((reg["pagina"], [la.Tok(*t) for t in reg["toks"]], reg["fuente"]))
    if actual is not None:
        salida_documento(actual, paginas, errores)
    return 0


if __name__ == "__main__":
    sys.exit(main())
