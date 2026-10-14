# pdf-totals manual

## Name

lotes_auto.py - add up the number of pieces per lot number found in a folder of PDF documents

diagnostico.py - show what lotes_auto.py reads and detects in one or more PDF documents

## Synopsis

```
python lotes_auto.py [-h] [--version] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] PATH
```

## Description

lotes_auto.py searches FOLDER and all of its subfolders for files ending in .pdf. From each PDF it reads every lot number and the number of pieces recorded for that lot. It then adds up the pieces of each lot over all documents and writes one row per lot to an Excel workbook. Lots that appear in several documents, or several times in one document, are added together.

The text of a PDF is read directly when the PDF contains text. When a page contains no text, it is treated as a scanned image and read with optical character recognition (OCR). If no OCR engine is installed, scanned PDFs are not read and are listed for manual review.

The program was written for Mexican pharmaceutical return documents. All messages and output files are in Spanish; their meaning is explained in this manual.

diagnostico.py uses exactly the same reading and detection code as lotes_auto.py. It prints the text that was read from each PDF, what was detected in it and the result, so that a PDF that is not read correctly can be examined. It does not write any files.

## Supported layouts

The program recognizes four layouts. Each page is examined in the order given here, and the first layout that produces results is used for that page.

1. Forms in which each column is one product. A row labelled with a lot word, such as `Lote`, contains one lot number per column. Other rows, labelled for example `Cantidad Recibida` or `Cantidad Facturada`, contain the quantity of each product in the same column. If several quantity rows exist, the row is chosen by the priority described under Configuration.

2. Tables with a heading row. A heading such as `CANTIDAD`, `PIEZAS` or `No. de piezas` marks the quantity column, and the values below it are read. The lot number is taken from a heading such as `LOTE` when present, otherwise from anywhere in the same row, for example inside the description, or from the next two rows. Headings that span two lines, such as `PIEZAS` above `DEVUELTAS`, are read as one heading.

3. Labelled lines in which the lot comes first, for example a line `LOTE: 25240310` followed by a line `CANTIDAD EN PIEZAS: 1,250`. Each lot is paired with the first quantity that follows it.

4. Labelled lines in which the quantity comes first. Each quantity is paired with the first lot that follows it. Whether a document uses layout 3 or layout 4 is decided by its first label.

Words recognized as lot labels: `lote`, `serie`, `serial`, `n/s`, `s/n`, optionally preceded or followed by `No.`, `Num.` or `Numero`. Words recognized as quantity labels: `cantidad`, `cant.`, `piezas`, `pieza`, `pzas`, `pza`, `pzs`, `pz`, `unidades`, `unid.`, `uds`. Words recognized as quantity headings in tables also include `devueltas`, `devuelto`, `frascos`, `ampolletas`, `ampulas` and `viales`. The word `unidad` alone is not a quantity, because it names the unit of measure. `Cajas` is never used, because it counts boxes, not pieces.

## Rules applied to every value

- A lot number must consist of exactly eight digits and start with 24, 25 or 26. Anything else is ignored, including prices, amounts, dates and product codes. Characters attached to a lot number by the OCR, such as a vertical bar, are removed.
- Numbers are read in Mexican notation: 1,250 is one thousand two hundred and fifty, and 1,250.50 is one thousand two hundred and fifty and a half.
- Quantities with decimals, and quantities that are zero or negative, are ignored, because pieces are whole numbers. This also prevents prices from being counted as pieces.
- Files whose name contains `FALTA REVISION`, in upper or lower case and with or without an accent on the O, are skipped and do not count towards any total.
- A single quantity larger than 50,000 is still added, but is reported at the end of the run as suspicious.

## Options

lotes_auto.py:

```
FOLDER       folder that contains the PDF files; subfolders are included
--version    print the version number and exit
-h, --help   print a short usage message and exit
```

diagnostico.py:

```
PATH         a single PDF file, or a folder whose PDF files are all examined
--lineas N   show at most N lines of text per PDF; the default is 40
--version    print the version number and exit
-h, --help   print a short usage message and exit
```

## Output files

lotes_auto.py writes the following files into FOLDER. Existing files with the same names are replaced. A file that would be empty is removed instead.

`totales_lotes.xlsx` is the result. It has one sheet named `Lotes` with two columns, `Lote` and `Cantidad en Piezas`, one row per lot, sorted by lot number. The quantities are stored as numbers and displayed with thousands separators.

`detalle_lotes.txt` lists, for every lot, each amount that was added to it and the PDF that the amount came from. Use it to find the source of a total that looks wrong.

`revisar_a_mano.txt` has two sections. The first lists the files that were skipped because their name contains `FALTA REVISION`. The second lists the PDFs from which no quantity could be read, each with the reason: no readable quantities, scanned with no OCR available, or an error while reading the file.

`sin_resultado.log` contains, for every PDF that produced no quantity, the reason and the complete text that was read from it. It is the first place to look when a PDF is not read correctly.

## Console messages

While it runs, lotes_auto.py prints one line per PDF. The message after the arrow has the following meaning:

```
-> N cantidades (texto)                 N quantities were read from a PDF that contains text
-> N cantidades (OCR)                   N quantities were read from a scanned PDF
-> 0 cantidades (...)                   nothing was read; see sin_resultado.log
-> omitido (FALTA REVISION)             the file was skipped because of its name
-> escaneado, OCR no disponible         a scanned PDF could not be read because OCR is missing
-> ERROR: ...                           the file could not be opened or read
```

At the end it prints the number of lots and quantities, the names of the files it wrote, and, if any single quantity is larger than 50,000, a list of those quantities headed by `cantidades sospechosamente grandes`.

## Exit status

```
0   success
1   the Excel file could not be written, usually because it is open in Excel
2   invalid arguments, or FOLDER does not exist
```

## Configuration

The settings are at the top of lotes_auto.py. Edit them with any text editor and save the file in UTF-8 encoding.

`LOTE_WORDS` and `CANT_WORDS` are regular expressions with the words that label lot numbers and quantities.

`PRIORIDAD_CANTIDAD` decides which quantity column or row is used when a document has several. The first word in the list that appears in a heading wins. The default order is `pieza`, `devuel`, `frasco`, `ampollet`, `ampula`, `vial`, `recibid`, `cantidad`, `factur`; that is, a column about pieces or returned goods is preferred, and an invoiced quantity is used only as a last resort.

`AVISO_CANTIDAD_GRANDE` is the size above which a single quantity is reported as suspicious. The default is `50000`.

`OMITIR_NOMBRE` is the pattern of file names that are skipped. The default matches `FALTA REVISION`.

`LOTE_ID` is the pattern of a valid lot number. The default accepts exactly eight digits starting with 24, 25 or 26. To accept, for example, lots starting with 27 as well, change `2[456]` to `2[4567]`.

## Installation

### Overview

The program needs Python and the Python packages pdfplumber and openpyxl. To read scanned PDFs it also needs an OCR engine:

- With Python 3.12 or earlier: the package rapidocr_onnxruntime. This is the preferred engine.
- With Python 3.13 or later: the packages rapidocr and onnxruntime. In testing this engine misread a digit that the preferred engine read correctly, so Python 3.12 is recommended when many scanned PDFs are processed.

The OCR engine is not available for FreeBSD, OpenBSD, NetBSD or Alpine Linux. On those systems only PDFs that contain text are read.

The instructions below place everything in a folder named `pdf-totals` in your home folder, with a private Python environment in `pdf-totals/venv` (`pdf-totals\python` on Windows). Nothing is installed system-wide except, where noted, Python itself.

### Windows 10 and Windows 11

No administrator rights are needed. Open PowerShell and run the following command. It downloads the embeddable Python 3.12 distribution from python.org into your user folder, installs pip and the required packages, and downloads the two scripts. Windows PowerShell 5.1 does not support the `&&` operator, so the steps are separated by semicolons.

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

The package msvc-runtime places the Microsoft Visual C++ runtime files next to Python, which the OCR engine needs and which some systems lack. Warnings from pip that scripts are installed in a folder that is not on PATH can be ignored.

Type `curl.exe`, not `curl`. In Windows PowerShell 5.1, `curl` is an alias for `Invoke-WebRequest`, which takes different options. `curl.exe` is included in Windows 10 version 1803 and later.

To run the program:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

If downloads are blocked on the machine, download the two files `python-3.12.10-embed-amd64.zip` and `get-pip.py` on another computer, copy them to the Downloads folder, and replace the two `Invoke-WebRequest` steps with the corresponding paths.

These instructions are for 64-bit Intel and AMD processors. On Windows on ARM, use `python-3.12.10-embed-arm64.zip` instead; this has not been tested.

### macOS

No administrator rights are needed beyond installing Python. Install Python 3.12 either with the installer from python.org (https://www.python.org/downloads/macos/) or with Homebrew (`brew install python@3.12`). Then run in Terminal:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

If Homebrew is used and python3.12 is not found, use `$(brew --prefix)/bin/python3.12` instead. Both Intel and Apple silicon Macs are supported.

To run the program:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

### Linux

On Linux the OCR engine is installed together with opencv-python-headless instead of opencv-python. The headless variant does not need the system graphics libraries, so no extra system packages are required.

All distributions use the same command to create the environment once Python 3.12 is available as `python3.12` (on Debian 12 use `python3`, which is version 3.11):

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

To run the program:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

How to obtain Python 3.12 on common distributions:

Debian 12 and Ubuntu 24.04 ship a suitable Python. Install the venv module and curl as root:

```
sudo apt install python3-venv curl
```

On Ubuntu 24.04 the command is `python3.12`. On Debian 12 the command is `python3` (version 3.11); replace `python3.12` with `python3` in the command above. Debian 13 and newer Ubuntu releases ship Python 3.13 or later; use the uv method below.

Fedora ships an additional Python 3.12 package:

```
sudo dnf install python3.12
```

openSUSE Leap 15.6 and Tumbleweed ship an additional Python 3.12 package:

```
sudo zypper install python312
```

Arch Linux ships only the newest Python. Use the uv method below.

Any glibc-based distribution, without root: uv is a single program that installs into your home folder and can download a private copy of Python 3.12.

```
curl -LsSf https://astral.sh/uv/install.sh | sh && mkdir -p ~/pdf-totals && cd ~/pdf-totals && ~/.local/bin/uv venv -p 3.12 venv && ~/.local/bin/uv pip install -p venv/bin/python pdfplumber openpyxl rapidocr_onnxruntime && ~/.local/bin/uv pip uninstall -p venv/bin/python opencv-python && ~/.local/bin/uv pip install -p venv/bin/python opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Alpine Linux uses the musl C library, for which no OCR engine is available. Only PDFs that contain text are read:

```
doas apk add python3 curl && mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv venv && venv/bin/pip install pdfplumber openpyxl && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Use `sudo` instead of `doas` if that is how your Alpine system is configured.

### FreeBSD, OpenBSD and NetBSD

Only PDFs that contain text are read. The OCR engine and the PDF rendering library used by pdfplumber are not available for these systems, and pdfplumber reads text without them. The cryptography package needed by pdfminer.six is taken from the system packages, because building it requires a Rust compiler. The remaining packages are written in pure Python and are installed without their optional dependencies.

FreeBSD, with the commands for root run through `su` or `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

The package names carry the Python version, 311 in this example. If your system uses another version, find the right names with `pkg search cryptography` and change `python3.11` accordingly.

OpenBSD, with the first command run as root through `doas`:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

NetBSD, with the first command run as root:

```
pkgin install py312-pip py312-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

As on FreeBSD, the package names carry the Python version, 312 in this example; use `pkgin search cryptography` to find the names on your system.

To run the program on any BSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

## Updating

Download the two scripts again. On Windows:

```
curl.exe -L -o "$env:USERPROFILE\pdf-totals\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$env:USERPROFILE\pdf-totals\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

On macOS and Linux:

```
cd ~/pdf-totals && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

On the BSDs, use `fetch` or `ftp` in the same way. GitHub may serve the previous version for a few minutes after a change. To obtain a specific version immediately, replace `main` in the address with the commit identifier.

## Troubleshooting

The Excel file cannot be saved. The program stops with exit status 1 when totales_lotes.xlsx is open in Excel. Close it and run the program again.

All or many PDFs show `0 cantidades`. Run diagnostico.py on one of them. If the text it shows is empty, the PDF is a scan and the OCR is not reading it. If the text is present but nothing is detected, the labels or the layout are not recognized; the text shown, or the corresponding entry in sin_resultado.log, shows which words the document uses, and they can be added to the configuration.

A total is much too large. Search for the lot in detalle_lotes.txt to see every amount that was added and its source. Common causes are a quantity column chosen wrongly because the document uses an unknown heading, two neighbouring numbers merged by the OCR, and the same PDF present twice in the folder.

The OCR cannot be loaded on Windows (`DLL load failed`). Install msvc-runtime as shown in the Windows instructions.

The OCR cannot be loaded on Linux (`libGL.so.1: cannot open shared object file`). Replace opencv-python with opencv-python-headless as shown in the Linux instructions.

pip cannot install rapidocr_onnxruntime. This package supports Python up to 3.12. Use Python 3.12, or install rapidocr and onnxruntime instead.

## Limitations

Handwritten numbers are read by the OCR when they are written clearly. Untidy handwriting, check marks and handwritten abbreviations such as `N/A` are not read reliably. PDFs whose quantities cannot be read appear in revisar_a_mano.txt and must be checked by hand.

The OCR can occasionally misread a digit, in particular on blurred or skewed scans. Totals obtained from scanned PDFs should be spot-checked against detalle_lotes.txt.

The lot number format is fixed by `LOTE_ID`. Documents that use another format are ignored until the pattern is changed.

Quantities are expected to be whole numbers of pieces. Documents that count boxes or other units are not converted.

## Files

```
lotes_auto.py      the program
diagnostico.py     the diagnostic tool; must be in the same folder as lotes_auto.py
README.md          short description (also README.de.md and README.es.md)
docs/              this manual in English, German and Spanish
tools/             the script that translates the documentation
LICENSE            license text
```

## License

MIT License. See the file LICENSE.
