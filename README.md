# pdf-totals

Languages: [English](README.md), [Deutsch](README.de.md), [Español](README.es.md)

pdf-totals reads a folder of PDF documents, such as invoices, delivery notes and forms for returned products. It finds every batch number and the number of pieces that is recorded for each batch. It then writes the total number of pieces for each batch to an Excel file.

The program reads the text of PDFs that contain text. It reads scanned PDFs with optical character recognition (OCR). It recognizes several common layouts: tables with a batch column and a quantity column, tables in which the batch number is part of the description, forms in which each column describes one product, and lines with labels such as `LOTE: 25240310` and `CANTIDAD: 1,250`.

The program was written for documents about returned pharmaceutical products in Mexico. Its messages and output files are in Spanish. A batch number must have exactly eight digits and start with `24`, `25` or `26`. This rule is easy to change.

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

The program also searches all subfolders. It skips every file whose name contains `FALTA REVISION`. When the program finishes, the folder contains these files:

- totales_lotes.xlsx contains one row for each batch with the total number of pieces.
- detalle_lotes.txt lists every amount that was added to each batch and the name of the PDF that contains it.
- revisar_a_mano.txt lists the skipped PDFs and the PDFs from which no quantity could be read.
- sin_resultado.log contains the text that was read from each PDF that produced no quantity.

To see what the program reads and detects in one PDF, run diagnostico.py with the path of the PDF. diagnostico.py does not write any files.

## Documentation

The manual describes all options, the supported layouts, the output files, the settings, and the installation on each system:

- English: [docs/en/MANUAL.md](docs/en/MANUAL.md)
- Deutsch: [docs/de/MANUAL.md](docs/de/MANUAL.md)
- Español: [docs/es/MANUAL.md](docs/es/MANUAL.md)

The German and Spanish documents are translated automatically from the English documents with Argos Translate. If a translation differs from the English document, the English document applies.

## License

MIT License. See the file LICENSE.
