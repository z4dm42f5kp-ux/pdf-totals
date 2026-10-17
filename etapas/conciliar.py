"""conciliar.py - add up the quantities of each batch and check them.

    ... | python etapas/conciliar.py [--carpeta DIR | --excel RUTA] [--perfil P] [--sector S]

Reads the records "cantidad" and "referencia" of etapas/cantidades.py and applies the
same rules as lotes_auto.py: quantities of skipped files (FALTA REVISION) do not count,
batches that are not in the reference Excel are left out, the not distributed quantity
comes from the Excel, and every batch is checked against the limits for one batch and
against the withdrawal report. When a batch or the expected total does not match, it
applies the same exact corrections as lotes_auto.py: a row copied exactly from another
PDF, a batch number misread by one digit, a check mark read as a digit, the other copy
of a repeated page (--sin-correcciones turns this off). It writes a record "correccion"
for each correction, "malo" for each PDF marked as bad, one record "lote" for each batch
and a record "resumen" with the three totals and the expected totals.

This stage does not read pages again. lotes_auto.py does the whole job, including the
second and third readings of doubtful pages and the file revisar_a_mano.txt.
"""
import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from _comun import aplicar_perfil, escribir, la, leer, opciones_perfil  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(prog="conciliar.py", description=__doc__.split("\n\n")[0])
    opciones_perfil(ap, excel=True)
    ap.add_argument("--sin-correcciones", action="store_true", help="no aplicar las correcciones exactas")
    args = ap.parse_args(argv)
    aplicar_perfil(args, excel_necesario=True)
    lecturas = {}
    omitidos = set()
    for reg in leer():
        t, f = reg["tipo"], reg.get("archivo")
        if t == "archivo" and reg.get("omitido"):
            omitidos.add(f)
        if t not in ("cantidad", "referencia") or f in omitidos:
            continue
        L = lecturas.setdefault(f, la.Lectura(f, [], [], [], "", []))
        if t == "cantidad":
            L.entradas.append(la.Entrada(f, reg["pagina"], reg["lote"], reg["piezas"], reg["clase"], reg["forma"],
                                         reg.get("conf", 1.0), list(reg.get("verificado", [])),
                                         list(reg.get("fallas", [])), list(reg.get("notas", [])),
                                         reg.get("excluido", "")))
        else:
            L.referencias.append(la.Referencia(f, reg["pagina"], reg["lote"], reg["ret"], reg["nod"],
                                               reg["total"], reg["cuadra"], tuple(reg.get("ret_opciones", []))))
    res = la.conciliar(lecturas)
    malos, correcciones = {}, []
    medir = la.valor_global()
    if not args.sin_correcciones and (res.problemas or (medir is not None and medir[0](res) != medir[1])):
        la.intentar_arreglos(lecturas, {}, res, la.ESPERADOS["total"], malos, correcciones)
        res = la.conciliar(lecturas)
    for f, L in lecturas.items():          # cifras que contradicen el propio documento
        fallas = sorted({x for e in L.entradas if not e.excluido for x in e.fallas})
        if fallas:
            malos.setdefault(f, []).extend(f"sus propias cifras no cuadran: {x}" for x in fallas)
    for c in correcciones:
        escribir({"tipo": "correccion", "texto": c})
    for f, motivos in sorted(malos.items()):
        escribir({"tipo": "malo", "archivo": f, "motivos": motivos})
    problemas = {}
    for clase, tipo, lote, valor, ref in res.problemas:
        problemas.setdefault(lote, []).append({"clase": clase, "tipo": tipo, "valor": valor,
                                               "referencia": None if ref == float("inf") else ref})
    for lote in sorted(set(res.ret) | set(res.nod) | set(res.refs), key=la.orden_lote):
        r, n = res.ret.get(lote, 0), res.nod.get(lote, 0)
        escribir({"tipo": "lote", "lote": lote, "retirado": r, "no_distribuida": n, "total": r + n,
                  "no_distribuida_de": res.nod_fuente.get(lote), "problemas": problemas.get(lote, [])})
    ret, nod = sum(res.ret.values()), sum(res.nod.values())
    esperado = {"retirado": la.TOTAL_RETIRADO_ESPERADO, "no_distribuida": la.TOTAL_NO_DISTRIBUIDA_ESPERADO,
                "total": la.TOTAL_ESPERADO}
    escribir({"tipo": "resumen", "retirado": ret, "no_distribuida": nod, "total": res.total,
              "esperado": esperado, "omitidos": sorted(omitidos),
              "coincide": {k: (v == esperado[k]) if esperado[k] is not None else None
                           for k, v in (("retirado", ret), ("no_distribuida", nod), ("total", res.total))}})
    return 0


if __name__ == "__main__":
    sys.exit(main())
