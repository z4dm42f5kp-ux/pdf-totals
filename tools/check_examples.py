"""Run lotes_auto.py on the example folder and compare the result with expected.json.

Usage:
    python tools/check_examples.py [FOLDER] [-- OPTIONS FOR lotes_auto.py]

The default folder is examples/ in the project folder. The folder is copied to a
temporary folder first, so the examples stay unchanged. The program prints every batch
whose withdrawn or not distributed quantity differs from the expected value.

Then it runs the pipeline of small programs in etapas/ (paginas.py | cantidades.py |
conciliar.py | excel.py) on the same folder, with the cache of the first run, and checks
that every batch has the same quantities as with lotes_auto.py.

expected.json can list known differences (known_differences): differences that
lotes_auto.py is known to produce on these examples, for example a cell that the OCR
cannot read. They are reported but do not count as failures.

Return values: 0 when every batch matches apart from the known differences, 1 when at
least one other batch differs, 2 when lotes_auto.py fails.
"""
import json
import pathlib
import shutil
import subprocess
import sys
import tempfile

from openpyxl import load_workbook

ROOT = pathlib.Path(__file__).resolve().parent.parent


def leer_tabla(ws, col_lote, col_cant):
    valores = {}
    fila = 3
    while True:
        lote = ws.cell(fila, col_lote).value
        if lote is None or lote == "TOTAL":
            return valores
        valores[str(lote)] = ws.cell(fila, col_cant).value or 0
        fila += 1


def main(argv):
    extra = []
    if "--" in argv:
        i = argv.index("--")
        argv, extra = argv[:i], argv[i + 1:]
    origen = pathlib.Path(argv[1]) if len(argv) > 1 else ROOT / "examples"
    esperado = json.loads((origen / "expected.json").read_text(encoding="utf-8"))
    tot = esperado["totals"]
    with tempfile.TemporaryDirectory() as tmp:
        carpeta = pathlib.Path(tmp) / "examples"
        shutil.copytree(origen, carpeta, ignore=shutil.ignore_patterns("resultados_lotes"))
        cmd = [sys.executable, str(ROOT / "lotes_auto.py"), str(carpeta),
               "--retirado-esperado", str(tot["retirado"]),
               "--no-distribuida-esperada", str(tot["no_distribuida"]),
               "--total-esperado", str(tot["total"])] + extra
        r = subprocess.run(cmd)
        if r.returncode != 0:
            print(f"lotes_auto.py failed with return code {r.returncode}")
            return 2
        res = carpeta / "resultados_lotes"
        ws = load_workbook(res / "totales_lotes.xlsx").active
        ret, nod = leer_tabla(ws, 1, 2), leer_tabla(ws, 4, 5)
        conocidas = {}
        for k in esperado.get("known_differences", []):
            clave = (k["kind"], k["lote"])
            conocidas[clave] = conocidas.get(clave, 0) + k["difference"]
        errores, aceptadas = [], []
        for tipo, obtenido, correcto in (("retirado", ret, esperado["retirado"]),
                                         ("no_distribuida", nod, esperado["no_distribuida"])):
            for lote in sorted(set(obtenido) | set(correcto)):
                a, b = obtenido.get(lote, 0), correcto.get(lote, 0)
                if a == b:
                    continue
                linea = f"{tipo} {lote}: lotes_auto.py {a:,}, expected {b:,} ({a - b:+,})"
                if conocidas.get((tipo, lote)) == a - b:
                    aceptadas.append(linea)
                else:
                    errores.append(linea)
        for k in esperado.get("known_differences", []):
            r = ret if k["kind"] == "retirado" else nod
            correcto = esperado[k["kind"]]
            if r.get(k["lote"], 0) == correcto.get(k["lote"], 0):
                print(f"Note: the known difference in {k['kind']} {k['lote']} did not occur. "
                      "Remove it from KNOWN_DIFFERENCES in tools/make_examples.py.")
        print()
        print(f"retirado {sum(ret.values()):,} (expected {tot['retirado']:,}), "
              f"no distribuida {sum(nod.values()):,} (expected {tot['no_distribuida']:,})")
        if aceptadas:
            print(f"{len(aceptadas)} known differences (see known_differences in expected.json):")
            for e in aceptadas:
                print("  " + e)
        if errores:
            print(f"{len(errores)} differences:")
            for e in errores:
                print("  " + e)
            guardar = ROOT / "check_examples_output"
            shutil.rmtree(guardar, ignore_errors=True)
            shutil.copytree(res, guardar, ignore=shutil.ignore_patterns("cache"))
            print(f"The output files of lotes_auto.py are in {guardar}")
            return 1
        print("All batches match, apart from the known differences.")
        return comparar_tuberia(carpeta, ret, nod, extra)


def comparar_tuberia(carpeta, ret, nod, extra):
    """etapas/: same batches and quantities as lotes_auto.py on the same folder."""
    py, et = sys.executable, ROOT / "etapas"
    xlsx = carpeta / "tuberia.xlsx"
    perfil = [x for i, x in enumerate(extra) if x in ("--perfil", "--sector") or
              (i and extra[i - 1] in ("--perfil", "--sector"))]
    pasos = [[py, str(et / "paginas.py"), str(carpeta)],
             [py, str(et / "cantidades.py"), "--carpeta", str(carpeta)] + perfil,
             [py, str(et / "conciliar.py"), "--carpeta", str(carpeta)] + perfil,
             [py, str(et / "excel.py"), "-o", str(xlsx)] + perfil]
    procesos, entrada = [], None
    for paso in pasos:
        p = subprocess.Popen(paso, stdin=entrada, stdout=subprocess.PIPE if paso is not pasos[-1] else None)
        if entrada is not None:
            entrada.close()
        entrada = p.stdout
        procesos.append(p)
    codigos = [p.wait() for p in procesos]
    if any(codigos):
        print(f"The pipeline in etapas/ failed with return codes {codigos}")
        return 2
    ws = load_workbook(xlsx).active
    t_ret, t_nod = leer_tabla(ws, 1, 2), leer_tabla(ws, 4, 5)
    difs = [f"{tipo} {lote}: lotes_auto.py {a.get(lote, 0):,}, etapas/ {b.get(lote, 0):,}"
            for tipo, a, b in (("retirado", ret, t_ret), ("no_distribuida", nod, t_nod))
            for lote in sorted(set(a) | set(b)) if a.get(lote, 0) != b.get(lote, 0)]
    if difs:
        print(f"The pipeline in etapas/ gives {len(difs)} other quantities:")
        for x in difs:
            print("  " + x)
        return 1
    print("The pipeline in etapas/ gives the same quantities as lotes_auto.py.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
