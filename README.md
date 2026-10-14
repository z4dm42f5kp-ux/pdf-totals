# pdf-totals

Other languages: [Deutsch](README.de.md), [Español](README.es.md)

pdf-totals reads a folder of PDF documents, such as invoices, delivery notes and product return forms, finds every lot number and the number of pieces recorded for it, and writes the total number of pieces per lot to an Excel file.

It reads PDFs that contain text directly and reads scanned PDFs with optical character recognition (OCR). It understands several common layouts: tables with a lot column and a quantity column, tables where the lot number appears inside the description, forms where each column is one product, and labelled lines such as `LOTE: 25240310` followed by `CANTIDAD: 1,250`.

The program was written for Mexican pharmaceutical return documents. Its messages and output files are in Spanish. Lot numbers are expected to have exactly eight digits and start with 24, 25 or 26; this rule is easy to change.

## Requirements

- Python 3.9 or later. Python 3.12 is recommended because the preferred OCR engine supports Python up to 3.12.
- The Python packages pdfplumber and openpyxl.
- For scanned PDFs: rapidocr_onnxruntime (Python 3.12 or earlier) or rapidocr together with onnxruntime (Python 3.13 or later).

Administrator rights are not needed on Windows or macOS. On Linux and the BSDs, installing Python from the system packages requires root; on Linux this can be avoided with the uv tool, as described in the manual. OCR is available on Windows, macOS and glibc-based Linux. On FreeBSD, OpenBSD, NetBSD and Alpine Linux only PDFs that contain text can be read; scanned PDFs are listed for manual review.

## Quick start

Windows 10 and 11, in PowerShell, without administrator rights:

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

macOS, with Python 3.12 installed:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Linux, with Python 3.12 installed:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Installation on other Linux distributions and on the BSDs is described in the manual.

## Usage

Windows:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

macOS, Linux and BSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

The folder is searched recursively. Files whose name contains `FALTA REVISION` are skipped. When the program finishes, the folder contains:

- `totales_lotes.xlsx`: one row per lot with the total number of pieces.
- `detalle_lotes.txt`: every amount that was added to each lot and the PDF it came from.
- `revisar_a_mano.txt`: PDFs that were skipped and PDFs from which no quantity could be read.
- `sin_resultado.log`: the text that was read from each PDF that produced no quantity.

To see what the program reads and detects in a single PDF without writing any files, run `diagnostico.py` with the path of the PDF.

## Documentation

The manual describes every option, the supported layouts, the output files, the configuration settings and the installation on each platform:

- English: [docs/en/MANUAL.md](docs/en/MANUAL.md)
- Deutsch: [docs/de/MANUAL.md](docs/de/MANUAL.md)
- Español: [docs/es/MANUAL.md](docs/es/MANUAL.md)

The German and Spanish documents are translated automatically from the English originals with Argos Translate. If they disagree, the English documents are authoritative.

## License

MIT. See the file LICENSE.
