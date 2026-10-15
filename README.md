# pdf-totals

Languages: [English](README.md), [Deutsch](README.de.md), [Español](README.es.md)

pdf-totals reads a folder of PDF documents about the withdrawal of a product from the market, such as return forms, delivery receipts, goods receipts, inventory lists and withdrawal reports. For each batch number it calculates three quantities and writes them to an Excel file:

- the withdrawn quantity: pieces that were returned or recovered from the market;
- the not distributed quantity: pieces that are still in inventory, taken from a reference Excel file;
- the total quantity for destruction: the sum of both.

The program does not rely only on the words in the documents. It examines the structure of each table: which column contains whole numbers, which column multiplied by the unit price gives the printed total price, and which column adds up to a printed total. In a withdrawal report it finds the columns by arithmetic, because manufactured minus distributed equals not distributed, and not distributed plus withdrawn equals the total.

The withdrawn quantity of each batch is compared with the reference Excel file. No single batch may exceed `48,042` withdrawn pieces or `82,090` pieces that were not distributed. The withdrawn pieces must add up to exactly `209,908`, the not distributed pieces to exactly `134,543`, and the total for destruction to exactly `344,451`. If something does not match, doubtful scanned PDFs are read a second time, differences are corrected only when an exact proof exists, and PDFs are marked as bad only on conclusive evidence.

The program reads the text of PDFs that contain text and reads scanned PDFs with optical character recognition (OCR). It was written for documents about returned pharmaceutical products in Mexico. Its messages and output files are in Spanish. A batch number must have exactly eight digits and start with `25` or `26`. These rules are easy to change.

## Requirements

- Python 3.9 or a newer version. `Python 3.12` is recommended, because the preferred OCR engine supports Python versions up to `Python 3.12`.
- The Python packages pdfplumber and openpyxl.
- For scanned PDFs: the package rapidocr_onnxruntime with `Python 3.12` or an older version, or the packages rapidocr and onnxruntime with `Python 3.13` or a newer version.

On Windows and macOS you do not need administrator rights. On Linux and on the BSD systems, you need the root user to install Python from the packages of the system. On Linux you can avoid this with the program uv, as the manual describes.

OCR is available on Windows, on macOS and on Linux distributions that use glibc. On FreeBSD, OpenBSD, NetBSD and Alpine Linux, the program reads only PDFs that contain text. It lists scanned PDFs for manual review.

## Quick start

Windows 10 and Windows 11, in PowerShell, without administrator rights:

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

macOS, when `Python 3.12` is installed:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Linux, when `Python 3.12` is installed:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

The manual describes the installation on other Linux distributions and on FreeBSD, OpenBSD and NetBSD.

## How to use the program

Windows:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

macOS, Linux, FreeBSD, OpenBSD and NetBSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

The program also searches all subfolders. It skips every file whose name contains `FALTA REVISION`. The reference Excel file is the file `Retirado_retiro_producto_Traxicoll.xlsx` in the folder or in a subfolder. When the program finishes, the folder `resultados_lotes` inside the folder contains these files:

- totales_lotes.xlsx contains three tables side by side on one sheet: withdrawn, not distributed, and total for destruction, each with the batch number and the number of pieces.
- detalle_lotes.txt lists every quantity of each batch, the PDF that contains it and whether it is verified.
- revisar_a_mano.txt lists the PDFs that are marked as bad, the corrections, the batches that do not match, the suspicious PDFs and the PDFs from which no quantity could be read.
- sin_resultado.log contains the text that was read from each PDF that produced no quantity.
- the folder cache contains what was read from each page. The next run uses it, so it takes seconds instead of minutes.

The program reads several pages at the same time, one for each processor core. The option `--procesos 1` reads one page after the other. The option `--sin-cache` reads every page again.

To see what the program reads and detects in one PDF, run diagnostico.py with the path of the PDF. diagnostico.py does not write any files.

## Documentation

The manual describes all options, the supported layouts, the output files, the settings, and the installation on each system:

- English: [docs/en/MANUAL.md](docs/en/MANUAL.md)
- Deutsch: [docs/de/MANUAL.md](docs/de/MANUAL.md)
- Español: [docs/es/MANUAL.md](docs/es/MANUAL.md)

The German and Spanish documents are translated automatically from the English documents with Argos Translate. If a translation differs from the English document, the English document applies.

## License

MIT License. See the file LICENSE.
