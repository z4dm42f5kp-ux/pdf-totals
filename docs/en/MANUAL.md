# pdf-totals manual

## Name

lotes_auto.py - calculate the total number of pieces for each batch number in a folder of PDF documents

diagnostico.py - show what lotes_auto.py reads and detects in one or more PDF documents

## Synopsis

```
python lotes_auto.py [-h] [--version] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] PATH
```

## Description

lotes_auto.py searches the folder FOLDER and all of its subfolders for files whose name ends in .pdf. In each PDF it finds every batch number and the number of pieces that the document records for that batch. It then calculates the total number of pieces of each batch over all documents and writes one row per batch to an Excel workbook. When a batch appears in several documents, or several times in one document, its quantities are added together.

When a PDF contains text, the program reads the text directly. When a page contains no text, the program treats the page as a scanned image and reads it with optical character recognition (OCR). If no OCR engine is installed, the program does not read scanned PDFs and lists them for manual review.

The program was written for documents about returned pharmaceutical products in Mexico. All messages and output files are in Spanish. This manual explains their meaning.

diagnostico.py uses exactly the same code as lotes_auto.py to read and analyse documents. It shows the text that it read from each PDF, what it detected in that text, and the result. Use it to examine a PDF that is not read correctly. It does not write any files.

## Supported layouts

The program recognizes four layouts. It examines each page in the order given here and uses the first layout that produces results for that page.

1. Forms in which each column describes one product. One row has a batch label, such as `Lote`, and contains one batch number per column. Other rows have a quantity label, such as `Cantidad Recibida` or `Cantidad Facturada`, and contain the quantity of each product in the same column. If there are several quantity rows, the program chooses one by the priority that is described in the section Configuration.

2. Tables with column titles. A column title such as `CANTIDAD`, `PIEZAS` or `No. de piezas` marks the quantity column, and the program reads the values below it. The program takes the batch number from a column with a title such as `LOTE`. If there is no such column, it takes the batch number from any other place in the same row, for example from the description, or from the next two rows. A column title that is written on two lines, such as `PIEZAS` above `DEVUELTAS`, is read as one title.

3. Lines with labels, in which the batch comes first. An example is a line `LOTE: 25240310` followed by a line `CANTIDAD EN PIEZAS: 1,250`. The program pairs each batch with the first quantity that follows it.

4. Lines with labels, in which the quantity comes first. The program pairs each quantity with the first batch that follows it. The first label in the document decides whether the document uses layout 3 or layout 4.

These words are recognized as batch labels: `lote`, `serie`, `serial`, `n/s` and `s/n`. They may be combined with `No.`, `Num.` or `Numero`.

These words are recognized as quantity labels: `cantidad`, `cant.`, `piezas`, `pieza`, `pzas`, `pza`, `pzs`, `pz`, `unidades`, `unid.` and `uds`.

These words are also recognized as quantity column titles: `devueltas`, `devuelto`, `frascos`, `ampolletas`, `ampulas` and `viales`.

The word `unidad` alone does not mark a quantity, because it names the unit of measure. The word `cajas` is never used, because it counts boxes and not pieces.

## Rules for every value

- A batch number has exactly eight digits and starts with `24`, `25` or `26`. The program ignores all other numbers, for example prices, amounts, dates and product codes. If the OCR attaches other characters to a batch number, such as a vertical bar, the program removes them.
- The program reads numbers in Mexican notation. A comma separates the thousands and a point separates the decimals. For example, `1,250` is one thousand two hundred and fifty, and `1,250.50` is one thousand two hundred and fifty and a half.
- The program ignores quantities that have decimals, and quantities that are zero or negative, because a number of pieces is always a whole number. This rule also prevents prices from being counted as pieces.
- The program skips every file whose name contains `FALTA REVISION`. Capital and small letters are treated alike, and the O may have an accent. Skipped files do not count towards any total.
- If a single quantity is greater than `50,000`, the program still adds it to the total. When the program finishes, it reports such quantities as suspicious.

## Options

Options of lotes_auto.py:

- `FOLDER` is the folder that contains the PDF files. Subfolders are included.
- `--version` shows the version number. The program then stops.
- `-h` or `--help` shows a short help text. The program then stops.

Options of diagnostico.py:

- `PATH` is a single PDF file, or a folder. If it is a folder, the program examines all PDF files in it.
- `--lineas N` shows no more than N lines of text for each PDF. The default value is 40.
- `--version` shows the version number. The program then stops.
- `-h` or `--help` shows a short help text. The program then stops.

## Output files

lotes_auto.py writes the following files into the folder FOLDER. It replaces existing files with the same names. If a file would be empty, the program deletes it instead.

`totales_lotes.xlsx` contains the result. It has one sheet with the name `Lotes`. The sheet has two columns, `Lote` and `Cantidad en Piezas`, and one row for each batch. The rows are sorted by batch number. The quantities are stored as numbers and are shown with digit grouping, for example `1,406`.

`detalle_lotes.txt` lists every amount that was added to each batch, together with the name of the PDF that contains the amount. Use this file to find the source of a total that looks wrong.

`revisar_a_mano.txt` has two sections. The first section lists the files that were skipped because their name contains `FALTA REVISION`. The second section lists the PDFs from which no quantity could be read, each with the reason. The possible reasons are: no readable quantities; a scanned PDF and no OCR engine; an error while reading the file.

`sin_resultado.log` contains, for every PDF that produced no quantity, the reason and the complete text that the program read from it. Look at this file first when a PDF is not read correctly.

## Console messages

While it runs, lotes_auto.py shows one line for each PDF. The text after the arrow has the following meaning:

- `-> N cantidades (texto)` means that the program read N quantities from a PDF that contains text.
- `-> N cantidades (OCR)` means that the program read N quantities from a scanned PDF.
- `-> 0 cantidades` means that the program read no quantity. The file sin_resultado.log shows the text that it read.
- `-> omitido (FALTA REVISION)` means that the program skipped the file because of its name.
- `-> escaneado, OCR no disponible` means that the PDF is scanned and that no OCR engine is installed.
- `-> ERROR` means that the program could not open or read the file.

At the end, the program shows the number of batches and quantities and the names of the files that it wrote. If a single quantity is greater than `50,000`, it also shows a list of these quantities under the title `cantidades sospechosamente grandes`.

## Return codes

- `0` means that the program finished successfully.
- `1` means that the program could not write the Excel file. Usually the file is open in Excel.
- `2` means that the arguments are not valid, or that the folder does not exist.

## Configuration

The settings are at the beginning of the file lotes_auto.py. You can change them with any text editor. Save the file in the UTF-8 encoding.

`LOTE_WORDS` and `CANT_WORDS` are regular expressions. They contain the words that mark batch numbers and quantities.

`PRIORIDAD_CANTIDAD` decides which quantity column or row is used when a document has several. The program uses the first word in this list that appears in a column title. The default order is `pieza`, `devuel`, `frasco`, `ampollet`, `ampula`, `vial`, `recibid`, `cantidad`, `factur`. This means that a column about pieces or about returned goods is preferred, and that an invoiced quantity is used only if no other quantity exists.

`AVISO_CANTIDAD_GRANDE` is the limit above which a single quantity is reported as suspicious. The default value is `50000`.

`OMITIR_NOMBRE` is the pattern for the names of files that are skipped. The default value matches `FALTA REVISION`.

`LOTE_ID` is the pattern for a valid batch number. The default value accepts exactly eight digits that start with `24`, `25` or `26`. For example, to also accept batch numbers that start with `27`, change `2[456]` to `2[4567]`.

## Installation

### General information

The program needs Python and the Python packages pdfplumber and openpyxl. To read scanned PDFs, it also needs an OCR engine:

- With `Python 3.12` or an older version: the package rapidocr_onnxruntime. This is the preferred engine.
- With `Python 3.13` or a newer version: the packages rapidocr and onnxruntime. In a test, this engine read one digit wrongly that the preferred engine read correctly. For this reason, `Python 3.12` is recommended if you process many scanned PDFs.

No OCR engine is available for FreeBSD, OpenBSD, NetBSD or Alpine Linux. On these systems, the program reads only PDFs that contain text.

The following instructions put everything into a folder with the name `pdf-totals` in your home folder. The Python packages are installed into a private Python environment in `pdf-totals/venv`, or in `pdf-totals\python` on Windows. Nothing is installed for the whole system, except Python itself where this is stated.

### Windows 10 and Windows 11

You do not need administrator rights. Open PowerShell and run the following command. The command downloads the embeddable distribution of `Python 3.12` from python.org into your user folder. It then installs pip and the required packages, and downloads the two scripts. Windows PowerShell 5.1 does not support the `&&` operator, so the steps are separated with the `;` character.

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

The package msvc-runtime copies the Microsoft Visual C++ runtime files into the Python folder. The OCR engine needs these files, and some systems do not have them. You can ignore the pip warnings about scripts that are installed in a folder that is not on PATH.

Use `curl.exe` and not `curl`. In Windows PowerShell 5.1, `curl` is another name for `Invoke-WebRequest`, which has different options. Windows 10 includes `curl.exe` since version 1803.

To run the program:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

If downloads are blocked on the computer, download the files `python-3.12.10-embed-amd64.zip` and `get-pip.py` on another computer. Copy them into the Downloads folder. Then replace the two `Invoke-WebRequest` steps in the command with the paths of these files.

These instructions are for 64-bit processors from Intel and AMD. For Windows on ARM processors, use `python-3.12.10-embed-arm64.zip` instead. This has not been tested.

### macOS

You need administrator rights only to install Python. Install `Python 3.12` with the installer from python.org (https://www.python.org/downloads/macos/) or with Homebrew (`brew install python@3.12`). Then run the following command in the Terminal application:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

If you use Homebrew and the command `python3.12` is not found, use `$(brew --prefix)/bin/python3.12` instead. Macs with Intel processors and Macs with Apple processors are both supported.

To run the program:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

### Linux

On Linux, the OCR engine is installed with the package opencv-python-headless instead of opencv-python. The package opencv-python-headless does not need the graphics libraries of the system, so you do not need to install additional system packages.

When `Python 3.12` is available as the command `python3.12`, the following command creates the environment on every distribution. On Debian 12, use the command `python3` instead, which is `Python 3.11`.

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

To run the program:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

The following paragraphs explain how to get `Python 3.12` on common distributions.

Debian 12 and Ubuntu 24.04 include a suitable version of Python. Install the venv module and the curl program. This needs the root user:

```
sudo apt install python3-venv curl
```

On Ubuntu 24.04, the command is `python3.12`. On Debian 12, the command is `python3`, which is `Python 3.11`. In that case, replace `python3.12` with `python3` in the command above. Debian 13 and newer versions of Ubuntu include `Python 3.13` or a newer version. On these systems, use the method with uv that is described below.

Fedora offers an additional package for `Python 3.12`:

```
sudo dnf install python3.12
```

openSUSE Leap 15.6 and openSUSE Tumbleweed offer an additional package for `Python 3.12`:

```
sudo zypper install python312
```

Arch Linux offers only the newest version of Python. Use the method with uv that is described below.

On every Linux distribution that uses glibc, you can use uv and do not need the root user. uv is a single program. It is installed into your home folder and can download a private copy of `Python 3.12`.

```
curl -LsSf https://astral.sh/uv/install.sh | sh && mkdir -p ~/pdf-totals && cd ~/pdf-totals && ~/.local/bin/uv venv -p 3.12 venv && ~/.local/bin/uv pip install -p venv/bin/python pdfplumber openpyxl rapidocr_onnxruntime && ~/.local/bin/uv pip uninstall -p venv/bin/python opencv-python && ~/.local/bin/uv pip install -p venv/bin/python opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Alpine Linux uses the C library musl. No OCR engine is available for musl, so the program reads only PDFs that contain text:

```
doas apk add python3 curl && mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv venv && venv/bin/pip install pdfplumber openpyxl && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

If your Alpine Linux system uses `sudo` instead of `doas`, replace `doas` with `sudo`.

### FreeBSD, OpenBSD and NetBSD

On these systems, the program reads only PDFs that contain text. The OCR engine is not available for these systems. The PDF rendering library that pdfplumber uses is not available either, but pdfplumber can read text without it.

The package pdfminer.six needs the package cryptography. A Rust compiler is needed to build cryptography from source code, so cryptography is installed from the packages of the system. All other packages are written in pure Python. They are installed without their optional dependencies.

FreeBSD. Run the first command as the root user, for example with `su` or `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

The package names contain the version of Python, which is `311` in this example. If your system uses another version, find the correct names with the command `pkg search cryptography`. Then change `python3.11` to match.

OpenBSD. Run the first command as the root user with `doas`:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

NetBSD. Run the first command as the root user:

```
pkgin install py312-pip py312-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

As on FreeBSD, the package names contain the version of Python, which is `312` in this example. Find the correct names on your system with the command `pkgin search cryptography`.

To run the program on FreeBSD, OpenBSD or NetBSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

## Updating

To update the program, download the two scripts again. On Windows:

```
curl.exe -L -o "$env:USERPROFILE\pdf-totals\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$env:USERPROFILE\pdf-totals\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

On macOS and Linux:

```
cd ~/pdf-totals && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

On FreeBSD, use `fetch` in the same way. On OpenBSD and NetBSD, use `ftp`.

After a change, GitHub can deliver the previous version for a few minutes. To get a specific version immediately, replace `main` in the address with the identifier of the commit.

## Troubleshooting

The Excel file cannot be saved. The program stops with return code 1 when totales_lotes.xlsx is open in Excel. Close the file and run the program again.

All or many PDFs show `0 cantidades`. Run diagnostico.py with one of these PDFs. If the text that it shows is empty, the PDF is scanned and the OCR does not read it. If the text is there but nothing is detected, the program does not know the labels or the layout of the document. The text shows which words the document uses. The same text is in sin_resultado.log. You can add these words to the configuration.

A total is much too large. Search for the batch number in detalle_lotes.txt. This file shows every amount that was added and the PDF that it came from. These are common causes:

- The program chose the wrong quantity column, because the document uses a column title that the program does not know.
- The OCR read two numbers that are close together as one number.
- The same PDF is in the folder twice.

The OCR cannot be loaded on Windows, and the message contains `DLL load failed`. Install the package msvc-runtime as shown in the instructions for Windows.

The OCR cannot be loaded on Linux, and the message contains `libGL.so.1`. Replace opencv-python with opencv-python-headless as shown in the instructions for Linux.

pip cannot install rapidocr_onnxruntime. This package supports Python versions up to `Python 3.12`. Use `Python 3.12`, or install the packages rapidocr and onnxruntime instead.

## Limitations

The OCR reads handwritten numbers when they are written clearly. It does not reliably read handwriting that is hard to read, check marks, or handwritten abbreviations such as `N/A`. PDFs with quantities that cannot be read are listed in revisar_a_mano.txt and must be checked by a person.

The OCR can read a digit wrongly, especially when a scan is blurred or tilted. Compare some totals from scanned PDFs with detalle_lotes.txt.

The format of a batch number is set by `LOTE_ID`. The program ignores batch numbers in other formats until the pattern is changed.

The program expects quantities to be whole numbers of pieces. It does not convert documents that count boxes or other units.

## Files in the project

- lotes_auto.py is the program.
- diagnostico.py is the diagnostic tool. It must be in the same folder as lotes_auto.py.
- README.md is a short description. It is also available as README.de.md and README.es.md.
- The folder `docs` contains this manual in English, German and Spanish.
- `tools/translate_docs.py` is the script that translates the documentation.
- `LICENSE` contains the text of the license.

## License

MIT License. See the file LICENSE.
