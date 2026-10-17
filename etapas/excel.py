"""excel.py - write the records "lote" into totales_lotes.xlsx.

    ... | python etapas/conciliar.py --carpeta DIR | python etapas/excel.py -o totales_lotes.xlsx

Writes the same workbook as lotes_auto.py: one sheet with the tables of withdrawn, not
distributed and total quantities, side by side, with the expected totals and the
batches that do not match in red. The expected totals come from the record "resumen".
"""
import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from _comun import la, leer, opciones_perfil  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(prog="excel.py", description=__doc__.split("\n\n")[0])
    ap.add_argument("-o", "--salida", metavar="XLSX", default="totales_lotes.xlsx", help="archivo que se escribe")
    opciones_perfil(ap)
    args = ap.parse_args(argv)
    try:                                   # solo para el idioma de los titulos del Excel
        la.aplicar_perfil(args.perfil, args.sector,
                          la.leer_config(args.config) if args.config else None)
    except la.ConfigNoValida as e:
        sys.exit(f"*** {e}")
    ret, nod, refs, problemas, resumen = {}, {}, {}, [], None
    for reg in leer():
        if reg["tipo"] == "lote":
            lote = reg["lote"]
            if reg["retirado"]:
                ret[lote] = reg["retirado"]
            if reg["no_distribuida"]:
                nod[lote] = reg["no_distribuida"]
            if not reg["retirado"] and not reg["no_distribuida"]:
                refs[lote] = None          # lote sin cantidades: tambien va en la tabla
            problemas += [(p["clase"], p["tipo"], lote, p["valor"], p["referencia"]) for p in reg["problemas"]]
        elif reg["tipo"] == "resumen":
            resumen = reg
    esp = (resumen or {}).get("esperado", {})
    la.ESPERADOS = {"ret": esp.get("retirado"), "nod": esp.get("no_distribuida"), "total": esp.get("total")}
    res = la.Resultado(ret, nod, refs, {}, problemas, sum(ret.values()) + sum(nod.values()))
    salida = pathlib.Path(args.salida).expanduser()
    try:
        la.escribir_excel(salida, res, la.ESPERADOS["total"])
    except PermissionError:
        sys.exit(f"*** No se pudo guardar {salida}: esta abierto en Excel.")
    print(f"Excel: {salida}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
