# tute-lote manual

## Name

lotes_auto.py - calculate, for each batch number, the pieces that were withdrawn, the pieces that were not distributed and the total number of pieces for destruction

diagnostico.py - show what lotes_auto.py reads and detects in one or more PDF documents

## Synopsis

```
python lotes_auto.py [-h] [--version] [--excel PATH] [--retirado-esperado N] [--no-distribuida-esperada N] [--total-esperado N] [--procesos N] [--sin-cache] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] [--lectura N] PATH
```

## Description

lotes_auto.py reads all PDF files in the folder FOLDER and in all of its subfolders. The documents are about the withdrawal of a product from the market: return forms, delivery receipts, goods receipts, inventory lists and withdrawal reports. For each batch number the program calculates three quantities:

- The withdrawn quantity (`Cantidad de producto retirado`). These are the pieces that were returned or recovered from the market. They come from return forms, delivery receipts and goods receipts.
- The not distributed quantity (`Cantidad no distribuida`). These are the pieces that are still in inventory and were never distributed. They are taken from the reference Excel file, which is described below. Without the reference Excel file, they come from inventory documents and from the withdrawal report.
- The total quantity for destruction (`Cantidad total para destrucción`). This is the sum of the two other quantities.

The program then checks the result. The limits apply to each batch on its own, not to the sum of a table: no single batch may have more than `48,042` withdrawn pieces or more than `82,090` pieces that were not distributed. The withdrawn quantities of all batches must add up to exactly `209,908`, the not distributed quantities to exactly `134,543`, and the total for destruction to exactly `344,451`. These values can be changed in the configuration or on the command line.

## The reference Excel file

The program looks for the Excel file with the name `Retirado_retiro_producto_Traxicoll` (.xlsx or .xlsm) in the folder FOLDER and in all of its subfolders. If there is no file with this name, it uses the first Excel file that has a column whose title contains `no distribuida`. A file in the old .xls format cannot be read; open it in Excel and save it as an .xlsx workbook. A file can also be given with the option `--excel`. In this file the program finds three columns: the column with the batch numbers, the column with the not distributed quantity, and the column whose title contains `retirado`, which is the withdrawn quantity according to the customers. Other columns, such as the total for destruction, are ignored.

The not distributed quantity of each batch is taken from this file, not from PDF documents. The withdrawn quantity is still calculated from the PDF documents, and the result for each batch is compared with the withdrawn quantity in the Excel file. This comparison shows exactly which batch does not match and which PDF documents contribute to it.

A batch number that appears in a PDF document but not in the Excel file is not counted, because it is almost always a batch number that was read wrongly. These quantities are listed in the review file. If the Excel file contains formulas, it must have been saved by Excel, so that the calculated values are stored in the file.

When a PDF contains text, the program reads the text directly. When a page contains no text, the program treats the page as a scanned image and reads it with optical character recognition (OCR). If no OCR engine is installed, the program does not read scanned PDFs and lists them for manual review.

The program was written for documents about returned pharmaceutical products in Mexico. All messages and output files are in Spanish. This manual explains their meaning.

diagnostico.py uses exactly the same code as lotes_auto.py to read and analyse documents. It shows the text that it read from each PDF, the decisions that it made and the result. Use it to examine a PDF that is not read correctly. It does not write any files.

## How the program finds the quantities

The program does not rely only on the words in the documents. It examines the structure of each table and uses arithmetic to decide which number is which.

First, the program finds the rows that contain a batch number. The columns of the table are formed from the positions of the values in these rows. Every value is classified as a batch number, a date, a price, a whole number, a code or text.

Then the program decides which column contains the number of pieces. Only columns with whole numbers are considered. Row numbers, codes that repeat in every row and years are excluded. The following evidence is used, from strongest to weakest:

1. Pieces multiplied by the unit price equal the total price in the same row. For example, `4,408` multiplied by `$27.00` is `$119,016.00`.
2. The sum of the column equals a number that is printed elsewhere in the document, for example `TOTAL DE PIEZAS 6,889` or `No CAJAS / PIEZAS 1978`.
3. The title of the column. Titles such as `PIEZAS`, `CANTIDAD EN PIEZAS`, `No. de piezas`, `RECIBIDA` or `DEVUELTAS` count in favour. Titles such as `FACTURADA` or `DISTRIBUIDA` count against, because they describe the original quantity and not the returned quantity. Titles with `CAJAS` count against, because they count boxes and not pieces.

A withdrawal report is recognized by arithmetic alone. In a withdrawal report, each row contains five numbers for a batch, and two relations hold in every row:

- manufactured minus distributed equals not distributed;
- not distributed plus withdrawn equals the total for destruction.

The program finds the columns that fulfil these relations, so the column titles are not needed. The withdrawn column is the one that is never larger than the distributed column. If a row of the report does not fulfil the relations, for example because the OCR read a digit wrongly, the program uses the relations to recover the correct value.

The type of each quantity is decided as follows. A quantity is not distributed if its column title or row title says so, for example `EXISTENCIA`, `INVENTARIO` or `NO DISTRIBUIDO`. Otherwise the words of the whole page decide: pages about returns, withdrawals or receipts give withdrawn quantities, and pages about inventory give not distributed quantities. If a page has no clear type, the type of the other pages of the same PDF is used. If the PDF has no clear type, the quantity counts as withdrawn and a note is added.

The program also reads two other layouts:

- Forms in which each column describes one product. One row contains the batch numbers, and other rows contain quantities such as `Cantidad Recibida` or `Cantidad Facturada`. The row is chosen by the priority that is described in the section Configuration.
- Text lines with labels, such as a line `LOTE: 25240310` and a line `CANTIDAD EN PIEZAS: 1,250`.

## Checks and corrections

Every quantity is marked as verified or not verified. A quantity is verified when pieces multiplied by the unit price equal the printed total price, when the sum of its table equals a printed total, or when the total of its batch equals the withdrawal report.

If a table contradicts its own printed total of pieces, and exactly one row is not verified, the program calculates the correct value of that row from the printed total. It uses the corrected value only if the printed total price confirms it. Otherwise, the PDF is marked as bad.

Two pages of the same PDF that contain exactly the same batch numbers are treated as the same document, for example a return form and its detailed list, or the same receipt printed twice. Only one of the two pages is counted. The program prefers the page whose sum matches its own printed total. If the two pages do not agree for a batch, the program uses the copy that matches the reference Excel file.

A batch that appears on several other pages of the same PDF is counted once when the quantity is the same. A page whose quantity is the sum of the other pages is not counted. If the same batch appears on several pages with different quantities, the quantities are added, and the PDF is listed for review. Several rows of the same batch on the same page are always added, because they are separate lines of the document, for example different invoices.

If a batch exceeds its limit, does not match the withdrawal report, or the total does not match the expected total, the program reads the doubtful pages of the doubtful scanned PDFs two more times. A doubtful page is a scanned page with a quantity that is not verified, that has a failed check, that belongs to a batch that does not match, or whose batch number is not in the Excel file; a page of a withdrawal report whose row does not match; and a page without quantities that mentions batches or pieces, that was read with a low confidence, or from which almost nothing was read. Pages whose quantities are all verified are not read again, because another reading cannot improve them. If a PDF produced no quantity at all, all its scanned pages are read again. The second reading uses a higher resolution and more contrast. The third reading enlarges the image to twice its size and sharpens it, which helps with small or blurred text. The program keeps the reading with the most verified quantities, and keeps the other readings for the next step. Then it tries to explain each difference exactly and without ambiguity:

- If the other reading of exactly one PDF makes the batch match, that reading is used.
- If a row in one PDF is an exact copy of a row in another PDF, and the batch matches when the copy is removed, the copy is not counted.
- If a quantity has a batch number that is not in the reference Excel file, differs from a batch in the Excel file by one digit or by two swapped digits, and makes that batch match exactly, the batch number is corrected.
- If a pen stroke or check mark next to a number was read as an extra digit 1, 7 or 4, and removing it makes the batch match exactly, the number is corrected.
- If the other reading of exactly one PDF, or the removal of exactly one copied row, makes the withdrawn total exactly equal to its expected total, this change is applied.

A PDF is marked as bad only when the proof is conclusive: when its own numbers contradict its own printed totals in every reading, or when it contains an exact copy that is proven by a batch or by the expected total. All other doubts are listed as suspicious and are not changed.

## Rules for every value

- A batch number has exactly eight digits and starts with `25` or `26`. The program ignores all other numbers, for example prices, dates, invoice numbers and product codes. If the OCR attaches other characters to a batch number, such as a vertical bar, the program removes them. If the OCR splits a handwritten batch number in two, the parts are joined again.
- The program reads numbers in Mexican notation. A comma separates the thousands and a point separates the decimals. For example, `1,250` is one thousand two hundred and fifty, and `1,250.50` is one thousand two hundred and fifty and a half. A number such as `4.408`, where the OCR read the comma as a point, is read as `4,408`, because pieces are always whole numbers.
- Handwritten check marks and strokes next to a number, such as `695 ✓` or `/ 630`, are ignored.
- A column whose title names a reason, a code, a folio or an invoice is not used as the quantity, and neither is a column with the same value in every row.
- The program skips every file whose name contains `FALTA REVISION`. The program ignores the difference between upper case and lower case, and the O may have an accent. Skipped files do not count towards any total.

## Options

Options of lotes_auto.py:

- `FOLDER` is the folder that contains the PDF files. Subfolders are included.
- `--excel PATH` uses this reference Excel file instead of searching the folder.
- `--retirado-esperado N` sets the expected total of withdrawn pieces. The default value is `209,908`. The value 0 switches the check off.
- `--no-distribuida-esperada N` sets the expected total of pieces that were not distributed. The default value is `134,543`. The value 0 switches the check off.
- `--total-esperado N` sets the expected total for destruction. The default value is `344,451`. The value 0 switches the check off.
- `--procesos N` reads N pages at the same time, in N separate processes. The default is automatic: one process for each processor core, at most 8, and not more than the free memory allows (about 1.2 GB for each process). The value 1 reads one page after the other in a single process.
- `--sin-cache` reads every page again, even when it is in the cache (see below).
- `--version` shows the version number. The program then stops.
- `-h` or `--help` shows a short help text. The program then stops.

Options of diagnostico.py:

- `PATH` is a single PDF file, or a folder. If it is a folder, the program examines all PDF files in it.
- `--lineas N` shows no more than N lines of text for each PDF. The default value is 40.
- `--lectura N` uses OCR reading N: 1 is the normal reading, 2 uses a higher resolution and more contrast, and 3 enlarges and sharpens the image.
- `--version` shows the version number. The program then stops.
- `-h` or `--help` shows a short help text. The program then stops.

## Output files

lotes_auto.py creates the folder `resultados_lotes` inside the folder FOLDER and writes the following files into it. It replaces existing files with the same names. The folder `resultados_lotes` is ignored when the program searches for PDF documents and Excel files.

`totales_lotes.xlsx` contains the result on one sheet with the name `Lotes`. The sheet contains three tables side by side: `Cantidad de producto retirado`, `Cantidad no distribuida` and `Cantidad total para destrucción`. Each table has the columns `Lote` and `Cantidad en Piezas` and a row `TOTAL`. All three tables list the same batches in the same order, so a row can be read across the tables. The third table and the totals are formulas, so they are updated when you change a number by hand. Below each table, the sheet shows the expected total and the difference. Cells of batches that exceed a limit or do not match the withdrawal report are red.

`detalle_lotes.txt` lists, for every batch, each quantity that was found, the PDF and page that contain it, whether it is verified and why, and whether it was counted. It also shows the values of the withdrawal report. Use this file to find the source of a total that looks wrong.

`revisar_a_mano.txt` lists everything that needs a person, in numbered sections: a summary with the three totals and their expected values; a list of actions in the order in which they should be done; the batches that do not match, with every quantity that was found for them, the PDFs without readable quantities that mention them, the skipped `FALTA REVISION` files that mention them, and a note when such a file contains exactly the missing quantity; the batches of the Excel file without any PDF; the quantities whose batch number is not in the Excel file, with the most similar batch of the Excel file; for every PDF without readable quantities, the number of pages, the OCR confidence, the reason, the batches that were seen, the numbers that might be misread batch numbers, the first lines of text and what to do; the PDFs that are marked as bad; the corrections; the suspicious PDFs; the repetitions within a PDF; the same batch and quantity in several PDFs; the skipped files with their quantities for information; and a list of all PDFs with their result.

The files whose name contains `FALTA REVISION` are read only to report what they contain. They never count towards a total.

The folder `paginas` contains an image (PNG) of every page that needs a person: pages without readable quantities, pages with quantities of batches that do not match, suspicious quantities, corrected quantities, quantities with a batch number that is not in the Excel file, and pages of PDFs that are marked as bad. revisar_a_mano.txt names the image next to each problem, so that a page can be checked without opening the PDF. The images cannot be created on FreeBSD, OpenBSD and NetBSD.

The folder `cache` contains what was read from each page, one small file for each page and reading. The next run uses these files instead of reading the pages again, so a second run takes seconds instead of minutes. A file in the cache belongs to the content of a PDF, not to its name: when a PDF is renamed, for example to add `FALTA REVISION`, its pages are not read again; when a PDF is replaced by a new version, its pages are read again. Files of PDFs that are no longer in the folder are deleted at the end of a run. You can delete the folder `cache` at any time.

`sin_resultado.log` contains, for every PDF that produced no quantity, the reason, the complete text that the program read from it and the decisions that the program made. Look at this file first when a PDF is not read correctly.

## Console messages

While it runs, lotes_auto.py shows one line for each PDF when all of its pages have been read. Because several pages are read at the same time, the PDFs are not always shown in alphabetical order. In brackets, the line shows the number of pages and the seconds that were needed to read them, or `cache` when the pages came from the cache. The text after the arrow says how many withdrawn and not distributed quantities were read, whether the PDF is a withdrawal report, and whether the text was read directly (`texto`) or with OCR (`OCR`). Other messages:

- `-> omitido (FALTA REVISION)` means that the program skipped the file because of its name.
- `-> escaneado, OCR no disponible` means that the PDF is scanned and that no OCR engine is installed.
- `-> ERROR` means that the program could not open or read the file.
- `Nuevas lecturas` means that the program reads the doubtful pages of the doubtful scanned PDFs again. Each line names the pages that were read again.
- `... sigue leyendo` means that one page has taken a long time. It is shown every 45 seconds while the page is read, with the PDF and the page number, so that a slow page can be found.
- `*** ... se deja sin leer` means that one page took more than 600 seconds. The program stops reading this page, continues with the other pages, and lists the PDF in revisar_a_mano.txt with the reason. The page is read again in the next run.
- `Tiempo` at the end shows the duration of the run, the number of pages that were read, and the number of pages that came from the cache.

At the end, the program shows the three totals and compares the total for destruction with the expected total. The word `COINCIDE` means that the totals are equal. Otherwise the difference is shown. Then it lists the PDFs that are marked as bad.

## Return values of the program

- `0` means that the program finished. Check the summary to see whether the totals match.
- `1` means that the program could not write the Excel file. Usually the file is open in Excel.
- `2` means that the arguments are not valid, or that the folder does not exist.

## Configuration

The settings are at the beginning of the file lotes_auto.py. You can change them with any text editor. Save the file in the UTF-8 encoding.

`TOTAL_ESPERADO`, `TOTAL_RETIRADO_ESPERADO` and `TOTAL_NO_DISTRIBUIDA_ESPERADO` are the expected totals for destruction, for withdrawn pieces and for pieces that were not distributed. The default values are `344451`, `209908` and `134543`. The value `None` switches a check off.

`MAX_RETIRADO_POR_LOTE` and `MAX_NO_DISTRIBUIDA_POR_LOTE` are the limits for a single batch. They are not limits for the sum of a table. The default values are `48042` and `82090`.

`EXCEL_NOMBRE` is the name of the reference Excel file that is searched first. The default value is `Retirado_retiro_producto_Traxicoll`. `CARPETA_RESULTADOS` is the name of the folder for the output files. The default value is `resultados_lotes`.

`EXCEL_REFERENCIA` is the path of the reference Excel file. The default value `None` means that the file is searched in the folder and its subfolders. `SOLO_LOTES_DEL_EXCEL` decides whether quantities with a batch number that is not in the Excel file are left out. The default value is `True`.

`USAR_REPORTE_PARA_RETIRADO` decides how the withdrawn quantities of a withdrawal report are used. With the default value `False`, they are only used to check the return documents. With `True`, they are also used for batches that have no return document in the folder.

`DUPLICADOS_EN_ARCHIVO` decides how a batch that is repeated in the same PDF is counted. With the default value `iguales`, it is counted once only when the quantities are equal. With `todos`, it is always counted once, and the largest quantity is kept.

`OMITIR_NOMBRE` is the pattern for the names of files that are skipped. The default value matches `FALTA REVISION`.

`LOTE_ID` is the pattern for a valid batch number. The default value accepts exactly eight digits that start with `25` or `26`. For example, to also accept batch numbers that start with `24`, change `2[56]` to `2[456]`.

`OCR_DPI` is the resolution of the first and the third OCR reading, and `OCR_DPI_REINTENTO` is the resolution of the second OCR reading. The pages are not converted to black and white, because in tests the OCR engine read such pages much worse. `CONFIANZA_MINIMA` is the OCR confidence below which a quantity that is not verified is listed as suspicious.

`PROCESOS` is the number of processes that read pages at the same time. The default value `0` means automatic. `LADO_MAX_PX` is the largest size, in pixels, of the longer side of a page image for the OCR. The default value is `6000`. The OCR engine reduces every image to 2000 pixels anyway, so this limit does not change the result for normal pages. It prevents huge images, for example from scans that were saved with 72 dots per inch, whose pages are about one meter high. `LADO_MAX_PX_AMPLIADA` is the same limit for the third reading before the image is enlarged. `USAR_CACHE` decides whether the cache is used. `CACHE_VERSION` is part of the name of every cache file; change it to make the program ignore the cache, for example after a change to the OCR settings. `AVISO_LENTO_S` is the number of seconds after which a slow page is reported. `LIMITE_PAGINA_S` is the number of seconds after which a page is given up. The default value is `600`.

`PRIORIDAD_RENGLON` decides which quantity row is used in forms in which each column describes one product.

`SEMANTICA` and `TIPO_DOCUMENTO` contain the words that are used as one signal among others to understand column titles and to recognize the type of a page.

## Installation

### General information

The program needs Python and the Python packages pdfplumber and openpyxl. To read scanned PDFs, it also needs an OCR engine:

- With `Python 3.12` or an older version: the package rapidocr_onnxruntime. This is the preferred OCR engine.
- With `Python 3.13` or a newer version: the packages rapidocr and onnxruntime. In a test, this OCR engine read one digit wrongly. The preferred OCR engine read the same digit correctly. For this reason, `Python 3.12` is recommended if you process many scanned PDFs.

No OCR engine is available for FreeBSD, OpenBSD, NetBSD or Alpine Linux. On these systems, the program reads only PDFs that contain text.

The following instructions put everything into a folder with the name `tute-lote` in your home folder. The Python packages are installed into a private Python environment in `tute-lote/venv`, or in `tute-lote\python` on Windows. Nothing is installed for the whole system, except Python itself where this is stated.

### Windows 10 and Windows 11

You do not need administrator rights. Open PowerShell and run the following command. The command downloads the embeddable distribution of `Python 3.12` from python.org into your user folder. It then installs pip and the required packages, and downloads the two scripts. Windows PowerShell 5.1 does not support the `&&` operator, so the steps are separated with the `;` character.

```
$d="$env:USERPROFILE\tute-lote"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

The package msvc-runtime copies the Microsoft Visual C++ runtime files into the Python folder. The OCR engine needs these files, and some systems do not have them. You can ignore the pip warnings about scripts that are installed in a folder that is not on PATH.

Use `curl.exe` and not `curl`. In Windows PowerShell 5.1, `curl` is another name for `Invoke-WebRequest`, which has different options. Windows 10 includes `curl.exe` since version 1803.

To run the program:

```
& "$env:USERPROFILE\tute-lote\python\python.exe" "$env:USERPROFILE\tute-lote\lotes_auto.py" "C:\path\to\pdfs"
```

If downloads are blocked on the computer, download the files `python-3.12.10-embed-amd64.zip` and `get-pip.py` on another computer. Copy them into the Downloads folder. Then replace the two `Invoke-WebRequest` steps in the command with the paths of these files.

These instructions are for 64-bit processors from Intel and AMD. For Windows on ARM processors, use `python-3.12.10-embed-arm64.zip` instead. This has not been tested.

### macOS

You need administrator rights only to install Python. Install `Python 3.12` with the installer from python.org (https://www.python.org/downloads/macos/) or with Homebrew (`brew install python@3.12`). Then run the following command in the Terminal application:

```
mkdir -p ~/tute-lote && cd ~/tute-lote && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

If you use Homebrew and the command `python3.12` is not found, use `$(brew --prefix)/bin/python3.12` instead. Macs with Intel processors and Macs with Apple processors are both supported.

To run the program:

```
~/tute-lote/venv/bin/python ~/tute-lote/lotes_auto.py /path/to/pdfs
```

### Linux

On Linux, the OCR engine is installed with the package opencv-python-headless instead of opencv-python. The package opencv-python-headless does not need the graphics libraries of the system, so you do not need to install additional system packages.

When `Python 3.12` is available as the command `python3.12`, the following command creates the environment on every distribution. On Debian 12, use the command `python3` instead, which is `Python 3.11`.

```
mkdir -p ~/tute-lote && cd ~/tute-lote && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

To run the program:

```
~/tute-lote/venv/bin/python ~/tute-lote/lotes_auto.py /path/to/pdfs
```

The following paragraphs explain how to get `Python 3.12` on common Linux distributions.

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
curl -LsSf https://astral.sh/uv/install.sh | sh && mkdir -p ~/tute-lote && cd ~/tute-lote && ~/.local/bin/uv venv -p 3.12 venv && ~/.local/bin/uv pip install -p venv/bin/python pdfplumber openpyxl rapidocr_onnxruntime && ~/.local/bin/uv pip uninstall -p venv/bin/python opencv-python && ~/.local/bin/uv pip install -p venv/bin/python opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

Alpine Linux uses the C library musl. No OCR engine is available for musl, so the program reads only PDFs that contain text:

```
doas apk add python3 curl && mkdir -p ~/tute-lote && cd ~/tute-lote && python3 -m venv venv && venv/bin/pip install pdfplumber openpyxl && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

If your Alpine Linux system uses `sudo` instead of `doas`, replace `doas` with `sudo`.

### FreeBSD, OpenBSD and NetBSD

On these systems, the program reads only PDFs that contain text. The OCR engine is not available for these systems. The PDF rendering library that pdfplumber uses is not available either, but pdfplumber can read text without it.

The package pdfminer.six needs the package cryptography. A Rust compiler is needed to build cryptography from source code, so cryptography is installed from the packages of the system. All other packages are written in pure Python. They are installed without their optional dependencies.

On FreeBSD, run the first command as the root user, for example with `su` or `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/tute-lote && cd ~/tute-lote && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

The package names contain the version of Python, which is `311` in this example. If your system uses another version, find the correct names with the command `pkg search cryptography`. Then use the matching command instead of `python3.11`, for example `python3.12`.

On OpenBSD, run the first command as the root user with `doas`:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/tute-lote && cd ~/tute-lote && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

On NetBSD, run the first command as the root user:

```
pkgin install py312-pip py312-cryptography
```

```
mkdir -p ~/tute-lote && cd ~/tute-lote && python3.12 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

As on FreeBSD, the package names contain the version of Python, which is `312` in this example. Find the correct names on your system with the command `pkgin search cryptography`.

To run the program on FreeBSD, OpenBSD or NetBSD:

```
~/tute-lote/venv/bin/python ~/tute-lote/lotes_auto.py /path/to/pdfs
```

## Updating

To update the program, download the two scripts again. On Windows:

```
curl.exe -L -o "$env:USERPROFILE\tute-lote\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py; curl.exe -L -o "$env:USERPROFILE\tute-lote\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

On macOS and Linux:

```
cd ~/tute-lote && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

On FreeBSD, use `fetch` in the same way. On OpenBSD and NetBSD, use `ftp`.

The project was called pdf-totals before. Installations that were made before the new name use the folder `pdf-totals` instead of `tute-lote`. They keep working: replace `tute-lote` with `pdf-totals` in the folder paths of the commands above, but not in the addresses.

After a change, GitHub can deliver the previous version for a few minutes. To get a specific version immediately, replace `main` in the address with the identifier of the commit.

## Troubleshooting

The Excel file cannot be saved. The program stops with return code 1 when totales_lotes.xlsx is open in Excel. Close the file and run the program again.

The total does not match the expected total. Open revisar_a_mano.txt. The section about batches that do not match shows which batches exceed a limit or differ from the withdrawal report, and which PDFs contribute to them. Then search for these batches in detalle_lotes.txt. Common causes are:

- A PDF that belongs to the withdrawal is missing from the folder, or is named `FALTA REVISION`.
- The OCR read a digit wrongly. The two readings of a PDF are compared, and a difference between them is noted as `dudosa`.
- The same document is in the folder twice under different names.
- The program used the wrong column for the quantity, because the document uses a layout that the program does not know. Run diagnostico.py with this PDF to see which column was chosen and why.

All or many PDFs show no quantities. Run diagnostico.py with one of these PDFs. If the text that it shows is empty, the PDF is scanned and the OCR does not read it. If the text is there but nothing is detected, the program does not know the layout of the document.

The OCR cannot be loaded on Windows, and the message contains `DLL load failed`. Install the package msvc-runtime as shown in the instructions for Windows.

The OCR cannot be loaded on Linux, and the message contains `libGL.so.1`. Replace opencv-python with opencv-python-headless as shown in the instructions for Linux.

pip cannot install rapidocr_onnxruntime. This package supports Python versions up to `Python 3.12`. Use `Python 3.12`, or install the packages rapidocr and onnxruntime instead.

## Limitations

The OCR reads handwritten numbers when they are written clearly. It does not reliably read handwriting that is hard to read, or handwritten abbreviations such as `N/A`. Quantities that are written by hand into a form are therefore often not read. These PDFs are listed in revisar_a_mano.txt and must be checked by a person.

The OCR can read a digit wrongly, especially when a scan is blurred. Quantities that are verified by a price, a printed total or the withdrawal report are reliable. Quantities that are not verified should be compared with the PDF.

If a process stops unexpectedly, for example because the memory is full, the program reads the remaining pages one after the other in a single process. The OCR takes about one to five seconds for each page and process. The first run over a large folder therefore takes some minutes. The next runs use the cache and take seconds. Only the doubtful pages are read a second and a third time.

The program expects quantities to be whole numbers of pieces. It does not convert documents that count boxes or other units.

## Files in the project

- lotes_auto.py is the program.
- diagnostico.py is the diagnostic tool. It must be in the same folder as lotes_auto.py.
- requirements.txt lists the tested versions of the required Python packages. requirements-dev.txt adds the packages for the tests and for the example documents.
- README.md is a short description. It is also available as README.de.md and README.es.md.
- The folder `docs` contains this manual in English, German and Spanish, and the file `security-feeds.opml`. This file is a list of feeds with the releases of every Python package that the program uses. Import it into a feed reader to follow the updates.
- The folder `examples` contains 120 invented PDF documents, a reference Excel file and the expected result in `expected.json`. See `examples/README.md`.
- `tools/make_examples.py` creates the example documents. `tools/check_examples.py` runs the program on them and compares the result with the expected result.
- `tools/translate_docs.py` translates the documentation.
- `AGENTS.md` contains the rules for AI agents that run, test or change the program.
- `SECURITY.md` describes how to report a vulnerability, and the security model.
- `RELEASES.txt` describes the stable and rolling releases, how they are made and how to download them.
- `Dockerfile` describes the container image that is published with every release.
- `bandit.yaml` is the configuration of the security analyser bandit.
- The folder `.github` contains the automatic workflows: translation, tests, security checks, releases and dependency updates.
- `LICENSE` contains the text of the license.

## License

MIT License. See the file LICENSE.
