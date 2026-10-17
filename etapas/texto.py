"""texto.py - print the text of the page records, one line of the page per line.

    python etapas/paginas.py FOLDER | python etapas/texto.py [--archivo]

Groups the words of each page into lines, the same way lotes_auto.py does, so that the
text can be searched with grep, findstr or Select-String. With --archivo, every line
starts with the PDF name and the page number.
"""
import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from _comun import la, leer  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(prog="texto.py", description=__doc__.split("\n\n")[0])
    ap.add_argument("--archivo", action="store_true", help="anteponer el PDF y la pagina a cada linea")
    args = ap.parse_args(argv)
    for reg in leer():
        if reg["tipo"] != "pagina":
            continue
        for r in la.agrupar_renglones([la.Tok(*t) for t in reg["toks"]]):
            print(f"{reg['archivo']}:{reg['pagina']}: {r.texto}" if args.archivo else r.texto)
    return 0


if __name__ == "__main__":
    sys.exit(main())
