# Example documents

This folder contains 120 PDF documents of an invented product withdrawal, a reference Excel file and the expected result. Everything in this folder is invented: the batch numbers, the companies, the people and the quantities. The documents contain no real data.

The folder is used to test lotes_auto.py automatically and to try the program without real documents.

## Contents

- 55 delivery receipts with text, with one to five batches each.
- 20 return forms with the unit price, the total price and printed totals.
- 12 letters with labelled lines, such as `LOTE: 25240139` and `CANTIDAD EN PIEZAS: 1,250`.
- 20 scanned documents. They contain only an image of the page, so they are read with OCR.
- 4 PDFs with three pages, in which the first page is repeated as the third page. The repeated page counts once.
- 6 files whose name contains `FALTA REVISION`. They are read but never counted.
- 1 inventory list of the not distributed pieces.
- 1 exact copy of another PDF, with `(copia)` in its name. It must not count twice.
- 1 withdrawal report with manufactured, distributed, not distributed, withdrawn and total quantities.
- `Retirado_retiro_producto_Traxicoll.xlsx`, the reference Excel file.
- `expected.json`, the correct result for every batch, and a description of every document.

## Use

To run the program on the examples and compare the result with `expected.json`:

```
python tools/check_examples.py
```

The examples are copied to a temporary folder first, so this folder stays unchanged. The first run takes some minutes, because the scanned documents are read with OCR.

`expected.json` lists one known difference. In one scanned return form, the OCR does not read the cell with the number of pieces. lotes_auto.py lists this PDF in revisar_a_mano.txt as a PDF without readable quantities and names its batch. tools/check_examples.py reports this difference but does not count it as a failure.

To run the program yourself:

```
python lotes_auto.py examples --retirado-esperado 357499 --no-distribuida-esperada 129884 --total-esperado 487383
```

The results are written to `examples/resultados_lotes`, which is not part of the project.

## Regenerating the examples

The files are created by `tools/make_examples.py`. The output is the same on every run. After a change to the generator, run:

```
python tools/make_examples.py
```

To check that this folder matches the generator:

```
python tools/make_examples.py --check
```

The generator needs the packages in `requirements-dev.txt`. Use exactly these versions: another version of reportlab, Pillow or pypdfium2 can write different bytes, and then `--check` reports differences.
