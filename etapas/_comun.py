"""Shared helpers for the pipeline stages in etapas/ (issue #8).

Every stage reads and writes JSON Lines: one JSON object per line, with a "tipo" key.
The output is pure ASCII (non-ASCII letters are written as \\u escapes), so the lines
pass unchanged through any pipe, including Windows PowerShell 5.1, which re-encodes
the text that flows between programs.

Record types:
    pagina      one page:      archivo, ruta, pagina, paginas, fuente, toks
    error       a page or file that could not be read: archivo, pagina, error
    archivo     one PDF:       archivo, fuente, paginas, omitido, avisos
    cantidad    one quantity:  archivo, pagina, lote, piezas, clase (ret|nod), forma,
                               conf, verificado, fallas, notas, excluido
    referencia  one row of a withdrawal report: archivo, pagina, lote, ret, nod, total,
                               cuadra, ret_opciones
    correccion  an exact correction that conciliar.py applied: texto
    malo        a PDF marked as bad on exact proof: archivo, motivos
    lote        one batch:     lote, retirado, no_distribuida, total, problemas
    resumen     the totals:    retirado, no_distribuida, total, esperado
"""
import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import lotes_auto as la  # noqa: E402


def escribir(registro, salida=None):
    (salida or sys.stdout).write(json.dumps(registro, ensure_ascii=True, separators=(",", ":")) + "\n")


def leer(entrada=None):
    """Records from stdin (or a file object). Empty lines are skipped. A line that is not
    JSON stops the stage with a clear message instead of a traceback."""
    for n, linea in enumerate(entrada or sys.stdin, 1):
        linea = linea.strip().lstrip("﻿")
        if not linea:
            continue
        try:
            reg = json.loads(linea)
        except ValueError:
            sys.exit(f"*** linea {n} de la entrada no es JSON: {linea[:80]!r}")
        if not isinstance(reg, dict) or "tipo" not in reg:
            sys.exit(f"*** linea {n} de la entrada no es un registro de tute-lote")
        yield reg


def opciones_perfil(ap, excel=False):
    ap.add_argument("--perfil", choices=list(la.PERFILES), default=None,
                    help=f"pais o idioma de los documentos (por omision: {la.PERFIL_PREDETERMINADO})")
    ap.add_argument("--sector", choices=list(la.SECTORES), default=None, help="sector de los productos")
    ap.add_argument("--config", metavar="RUTA", default=None,
                    help=f"archivo de configuracion ({la.CONFIG_NOMBRE})")
    if excel:
        ap.add_argument("--excel", metavar="RUTA", default=None, help="Excel de referencia")
        ap.add_argument("--carpeta", metavar="DIR", default=None,
                        help="carpeta donde buscar el Excel de referencia y tute-lote.json")


def aplicar_perfil(args, excel_necesario=False):
    """Applies --perfil, --sector, --config (or tute-lote.json in --carpeta) and loads the
    reference Excel when it is given or when the batch format comes from it. Messages go
    to stderr, so that stdout carries only records."""
    base = pathlib.Path(args.carpeta).expanduser() if getattr(args, "carpeta", None) else None
    ruta_cfg = (pathlib.Path(args.config).expanduser() if args.config
                else (base / la.CONFIG_NOMBRE if base else None))
    try:
        config = la.leer_config(ruta_cfg) if ruta_cfg and (args.config or ruta_cfg.is_file()) else None
        for d in la.aplicar_perfil(args.perfil, args.sector, config):
            print(f"Perfil: {d}", file=sys.stderr)
        excel = getattr(args, "excel", None)
        if excel:
            la._EXCEL = la.leer_excel_referencia(pathlib.Path(excel).expanduser())
            if la._EXCEL is None:
                sys.exit(f"*** {excel} no tiene una columna de lotes con cantidades")
        elif base and (excel_necesario or la.LOTE_ID is None):
            la._EXCEL = la.buscar_excel(base)
    except (la.ConfigNoValida, la.ExcelNoLegible) as e:
        sys.exit(f"*** {e}")
    la.ESPERADOS = {"ret": la.TOTAL_RETIRADO_ESPERADO, "nod": la.TOTAL_NO_DISTRIBUIDA_ESPERADO,
                    "total": la.TOTAL_ESPERADO}
    if la.LOTE_ID is None:
        sys.exit("*** Falta el formato del numero de lote: usa --excel o --carpeta con el Excel de "
                 f"referencia, o \"batch_pattern\" en {la.CONFIG_NOMBRE}")
