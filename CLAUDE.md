# Handoff notes for Claude Code: tute-lote (formerly pdf-totals)

This file is for a future Claude Code session continuing work on this project. It
records the user's situation, the business rules (as clarified by the user over a long
conversation), the code architecture, the test setup, the lessons learned and the open
items. Last updated: 2026-10-17 (lotes_auto.py v2.3.0).

---

## 1. The user and their environment

- The user is in Mexico and works on a pharmaceutical product recall (retiro de mercado)
  for the product TRAXICOLL / CEFTRIAXONA (Productos Farmacéuticos Collins). They write
  to Claude in English. The program's messages and output files are in Spanish.
- They run the script on an **unprivileged Windows 10 VM with no admin rights**. They
  must not need to sign into any account on the VM.
- Python: **embeddable Python 3.12** at `%USERPROFILE%\pdf-totals\python\python.exe`
  (pip installed through get-pip.py, `import site` enabled in `python312._pth`).
  Packages: pdfplumber, openpyxl, rapidocr_onnxruntime (1.4.4), msvc-runtime. pypdfium2
  comes with pdfplumber.
- Scripts live in `%USERPROFILE%\pdf-totals\` (lotes_auto.py, diagnostico.py).
- The real PDFs are in `%USERPROFILE%\pdf-totals\Franz` (59 PDFs, many of them scanned or
  handwritten). The reference Excel `Retirado_retiro_producto_Traxicoll(.xlsx)` is in
  that folder or a subfolder.
- They update the scripts with `curl.exe -L -o` from raw.githubusercontent.com. Always
  pin the URL to a commit hash, so that they can confirm which version they run
  (`--version`).
- Their shell is **PowerShell 5.1**, which has no `&&`; chain with `;`.
- **User preference (always follow):** "When providing commands always use proper && and
  always give one single massive command even if its huge config files which would need
  echo | tee." In PowerShell, this means a single line joined with `;`.
- They cannot upload the PDFs (confidential). They send phone photos of pages, of the
  Excel and of console output. Results come back as console text or as the generated
  `revisar_a_mano.txt` / `detalle_lotes.txt`.
- They are busy and became frustrated when debugging dragged on. They want a clear fix,
  tests run **before** pushing, and a short answer with one command to run.

Standard update-and-run command (replace HASH with the new commit):

```
$d="$env:USERPROFILE\pdf-totals"; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/HASH/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/HASH/diagnostico.py; & "$d\python\python.exe" "$d\lotes_auto.py" --version; & "$d\python\python.exe" "$d\lotes_auto.py" "$d\Franz"
```

---

## 2. Business rules (final, as clarified by the user)

Output: `resultados_lotes\totales_lotes.xlsx`, with **one sheet** (`Lotes`). The user
explicitly did not want extra tabs. It holds three tables side by side, each with the
columns `Lote | Cantidad en Piezas`, comma thousands separators and a `TOTAL` row:

1. **Cantidad de producto retirado**: the pieces returned or recovered from the market,
   read from the PDFs.
2. **Cantidad no distribuida**: taken **from the reference Excel**, column J
   ("Cantidad no distribuida (piezas)( inventarios"), not from the PDFs.
3. **Cantidad total para destrucción** = 1 + 2 (Excel formulas).

Rules:

- **Expected totals:** retirado **209,908**; no distribuida **134,543**; total
  **344,451**.
- **Per-lote caps (they apply to one lote, NOT to the sum of a table):**
  - retirado ≤ **48,042** for any single lote;
  - no distribuida ≤ **82,090** for any single lote.
- **Lote numbers** have exactly **8 digits**, start **only with 25 or 26**, and have no
  decimals or pipes. The regex is
  `LOTE_ID = (?<![\d.,])(2[56]\d{6})(?![\d]|[.,]\d)`. The user removed 24.
- **Reference Excel:** the file named `Retirado_retiro_producto_Traxicoll` (.xlsx/.xlsm)
  anywhere under the folder. Its columns:
  - column J: no distribuida;
  - column K: "Cantidad de producto retirado (clientes)", the per-lote retirado
    according to the customers, used for **cross-checking** the PDF totals;
  - column L (yellow, total para destrucción): **ignored**.

  The old `.xls` format is refused with a message telling the user to save it as .xlsx.
  If there is no file with that name, the first .xlsx with a "no distribuida" column is
  used. Lotes not in the Excel are excluded from retirado (`SOLO_LOTES_DEL_EXCEL`) and
  listed in revisar_a_mano.txt with the most similar Excel lote.
- **Skip files** whose name contains "FALTA REVISION" (regex
  `falta[\s_\-]*revisi[oó]n`). They never count. They are still read so that their
  contents can be listed in revisar_a_mano.txt for information.
- **Within one PDF** (merged multi-page), duplicated lotes or totals must not be double
  counted:
  - two pages with the same set of lotes → count one (prefer the page that matches its
    printed total);
  - the same lote with the same quantity on several pages → count once;
  - a page that is the sum of the other pages → excluded;
  - rows on the same page are always added, except a total row.
- **Mark a PDF as bad only when more than 99% sure** (exact proof). Otherwise read it
  again (2nd and 3rd OCR readings).
- **revisar_a_mano.txt** must be explicit about everything that could not be read:
  every unreadable PDF, why, the lotes it mentions, the first lines, what to do, and
  page images.
- All outputs go into the created folder `resultados_lotes` inside the PDF folder:
  - totales_lotes.xlsx;
  - detalle_lotes.txt;
  - revisar_a_mano.txt;
  - sin_resultado.log;
  - `paginas/*.png` (images of the pages to check);
  - `cache/` (new in v2.2).

  The `resultados_lotes` folder is ignored when the program searches for PDFs and Excel
  files.

Document types seen (from user photos): GFC return forms, Levic acuses (Culiacán,
Monterrey, Tijuana), Reinvex, Difarmer Mexicali (2 pages), LUIN ("PÁGINA 13 DE 50"),
Tenorio, Pharma Amigo, sideways forms (one product per column), a recall report
(fabricado / distribuido / no distribuido / retirado / total), and inventory lists.
Handwriting, check marks (✓, |, —) next to numbers, stamps, signatures and skewed scans
are common.

---

## 3. Repository and git workflow

- Repo: **github.com/z4dm42f5kp-ux/tute-lote** (MIT; renamed from `pdf-totals` on 2026-10-15; pushing to the old URL still works through GitHub's redirect). On 2026-10-15 unauthenticated raw.githubusercontent.com URLs returned 404, so the repository was probably made private. The user's VM downloads the scripts with unauthenticated `curl.exe`, so this only works while the repo is public. Check before giving them a command. Local clone used so far:
  `/home/claude/pdf-totals`.
- Files:
  - `lotes_auto.py`: main program, about 2,400 lines, Spanish identifiers and messages.
  - `diagnostico.py`: shows what lotes_auto reads for one PDF or a folder. Options:
    `--lineas N`, `--lectura 1|2|3` (the hidden `--segunda-lectura` equals 2) and
    `--version`. It writes no files.
  - `README.md`, `docs/en/MANUAL.md`: English, the source of truth. The user wanted no
    emojis and no odd formatting, "BSD-quality" docs, and install instructions for
    Windows 10/11, macOS, Linux (several distros) and the BSDs.
  - `README.de.md`, `README.es.md`, `docs/de/`, `docs/es/`: **generated** by the GitHub
    Action `.github/workflows/translate-docs.yml`, which runs
    `tools/translate_docs.py` with Argos Translate (protect regex + German glossary).
    Never edit them by hand. The bot pushes commits after every change to README.md or
    docs/en.
  - `CLAUDE.md`: this file (not translated).
  - `etapas/`: the pipeline of small programs (issue #8), see 4.5.
  - `tools/check_profiles.py`: invented en/de/es documents for the profiles (issue #5).
  - `AGENTS.md`: rules for AI agents (running, testing, PDF prompt injection, fetching).
  - `SECURITY.md`: vulnerability reporting, threat model, agent security model.
  - `RELEASES.txt`: stable (vN.0.x) and rolling (vN.M.x, M > 0) releases; how to tag.
  - `requirements.txt` (tested runtime versions), `requirements-dev.txt` (pinned
    reportlab/Pillow/pypdfium2/pyflakes for the example generator), `bandit.yaml`
    (reviewed skips with reasons), `Dockerfile` + `.dockerignore` (ghcr.io image).
  - `examples/`: 120 invented PDFs + reference Excel + `expected.json`, generated by
    `tools/make_examples.py` (deterministic: fixed seed, reportlab `invariant=1`, fixed
    PIL/openpyxl dates; `--check` regenerates and compares byte by byte).
    `tools/check_examples.py` runs lotes_auto on a temp copy and compares every lote.
  - `docs/security-feeds.opml`: RSS/Atom feeds of the dependencies' releases.
  - Workflows: `translate-docs.yml`, `tests.yml` (py_compile 3.9, pyflakes, examples
    `--check`, check_examples with OCR), `security.yml` (pip-audit, bandit, pyflakes;
    weekly), `release.yml` (on tag `v*` = VERSION: GitHub release with SHA256SUMS and
    a ghcr.io image), `dependabot.yml` (pip, actions, docker; weekly).
- Commit as:
  ```
  git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "...

  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01Ttwea2D8xmakuXbzKicts4"
  ```
  (Use the trailer lines given by the current session's system reminder if they differ.)
- Push with `git fetch -q origin main && git rebase -q origin/main && git push origin main`.
  The translation bot creates commits, so a plain push is often rejected.
- Bump `VERSION` in lotes_auto.py for user-visible changes. The header of
  revisar_a_mano.txt shows it.

Recent history:

| Commit | Change |
|---|---|
| (2026-10-17) | v2.3.0: issue #5 profiles (`--perfil`, `--sector`, `tute-lote.json`), issue #8 pipeline `etapas/`, issue #16 answered with measurements. Default profile byte-identical to 2.2.0 |
| (2026-10-15) | issues #1 #2 #3 #4 #6 #7 #9: rename docs to tute-lote, AGENTS.md, SECURITY.md, examples + CI tests, security CI + feeds, release process, logo (removed 2026-10-16 at the user's request). lotes_auto.py unchanged (2.2.0) |
| 897e1a5 | v2.2: speed (parallel processes, cache, targeted re-reads, image cap, watchdog) |
| a022f0e | third OCR reading, page images, safer choice of reading |
| 7c615f3 | Excel name lookup, resultados_lotes folder |
| 22a64e4 | v2.1: reference Excel, separate expected totals, lotes 25/26, duplicate pages, explicit review file |
| 81cc15c | v2: three quantities, structure-based detection |

---

## 4. How lotes_auto.py works

### 4.1 Reading pages

- `page_tokens(page, pasada)` decides between text and OCR:
  - If the page has text (`len(page.chars) >= 20`, or chars and no images), it uses
    pdfplumber words and returns fuente `texto`.
  - Otherwise it renders the page (`renderizar`, size-capped), preprocesses it, runs
    the OCR, deskews by the median box angle, and splits each OCR box into word tokens
    `Tok(x0 x1 y h text conf caja)`. It returns fuente `OCR`.
- The OCR readings (`preprocesar`):
  - reading 1: 300 dpi, raw image;
  - reading 2: 400 dpi, autocontrast + SHARPEN;
  - reading 3: 300 dpi (capped at 4400 px), enlarged 2× with Lanczos, then
    UnsharpMask(2, 160, 3).

  Binarization was tested and rejected: it got 10 of 63 values right, against 41 of 63
  for grey images.
- OCR engine:
  - `rapidocr_onnxruntime` 1.4.4 on Python ≤ 3.12;
  - `rapidocr` 3.x + onnxruntime on 3.13+.

  `load_ocr()` takes the thread count from `_hilos_ocr`. rapidocr 1.4.4 removes spaces,
  so the labels are matched tolerantly.

### 4.2 Analysis (pure Python, cheap)

`analizar_documento(nombre, [(toks, fuente)...], pasada)` → `Lectura`. Per page,
`analizar_pagina` runs the steps below, then `deduplicar_en_archivo` runs on the whole
file.

1. **Rows and cells:** robust row grouping (median), joining of split handwritten
   lotes, and stripping of marks.
2. **Cell types:** lote, fecha, dinero, entero, cajas, codigo, texto or vacio.
3. **Sideways forms:** `analizar_de_lado`.
4. **Tables** (`analizar_tablas`):
   - rows with lotes → item rows → columns clustered by x-overlap;
   - candidate integer columns are scored by:
     - pieces × price = amount;
     - column sum = a number printed on the page;
     - header words;
     - penalties for orig/facturada, cajas, motivo/folio/clave/codigo/sap, constant
       values, values above the per-lote cap, and values ≥ 1M.
   - **Recall reports** are detected by arithmetic (`relaciones_suma`,
     `roles_de_reporte`): fabricado − distribuido = no distribuido, and
     no distribuido + retirado = total. The roles come from two sum triples that share
     an addend; the smaller partner is retirado.
5. **Labelled lines** (`analizar_etiquetas`): `LOTE: ... CANTIDAD: ...`.
6. **Verification** (`verificar_tabla`):
   - checks: row relation, table sum = printed total, money totals;
   - corrections: through the printed total + price, or with `sin_marca` (a check mark
     read as 1, 7 or 4).

### 4.3 Reconciliation (`main`)

1. Read all pages (reading 1), then `conciliar` against the Excel. The checks are:
   per-lote caps, Excel K via `Referencia.acepta_ret` with `ret_opciones`, and the
   expected totals. `valor_global` uses the retirado total if it is known.
2. If something is wrong, re-read the **doubtful pages** of the doubtful OCR files with
   readings 2 and 3.
   - Doubtful files (unchanged since v2.1): OCR files where `dudosa(L)`, or files with
     an entry in a problem lote, or (a re-read is needed and the file has any unverified
     entry).
   - Doubtful pages (`paginas_dudosas`, new in v2.2):
     - OCR pages with unverified or failed entries, entries in problem lotes, or lotes
       outside the Excel;
     - pages with report rows that don't match;
     - pages without entries that are relevant, have low confidence or are almost
       empty;
     - all OCR pages if the file produced nothing.
   - The other pages keep their reading-1 tokens. The best reading is chosen per file by
     `calidad` = (verified count, −fallas, entries + refs, mean conf). The others go to
     `alternativas[f]`.
3. `intentar_arreglos` applies only exact, unique fixes, in this order:
   - a) another reading of one PDF makes the lote match;
   - b) a row copied exactly from another PDF (`ordenar_copias` prefers names without
     "copia");
   - d0) the duplicate-page copy swap;
   - e) a lote misread by one digit or a transposition (`lote_parecido`) that fills the
     deficit exactly;
   - f) a mark read as a digit;
   - c) global: another reading that makes the total exact;
   - d) global: one copy that explains the whole difference.
4. Notes are added: "dudosa" when the readings disagree, and verified when the lote
   matches the report. A file is marked bad only if all its readings contradict their
   own printed figures.
5. Outputs are written:
   - escribir_excel;
   - escribir_detalle;
   - paginas_a_revisar + exportar_paginas (110 dpi, capped at 2200 px);
   - escribir_revision;
   - sin_resultado.log;
   - a console summary with COINCIDE or DIFERENCIA, plus a `Tiempo:` line.

CLI options:
- `carpeta`;
- `--total-esperado`, `--retirado-esperado`, `--no-distribuida-esperada` (0 = no check);
- `--excel RUTA`;
- `--procesos N` (0/omitted = automatic; 1 = no pool);
- `--sin-cache`;
- `--perfil traxicoll|mx|us|de|eu`, `--sector pharma|devices|food|cosmetics|automotive|consumer|chemicals|general`,
  `--config RUTA` (default: `tute-lote.json` in the folder);
- `--version`.

Return codes: 0 = ok, 1 = the Excel output is open/locked, 2 = bad arguments, folder or
Excel.

---

### 4.4 Profiles, sectors and tute-lote.json (v2.3, issue #5)

- `aplicar_perfil(nombre, sector, config)` is called once in `main` (and in diagnostico.py
  and etapas/). With the default profile `traxicoll`, its sector `pharma` and no config it
  changes **nothing** (the outputs of the private kit stayed byte-identical). It mutates
  module globals, so call it once per process.
- Other profiles (`mx`, `us`, `de`, `eu`) extend the Spanish vocabulary with `VOCABULARIO`
  (en, de, fr, it): SEMANTICA, TIPO_DOCUMENTO, months, units, labels, Excel headers,
  skip-name words. `NUMEROS = "auto"` accepts `1.234,56 €` money. They set expected totals
  and Excel name to None, caps to infinity, and `LOTE_MAYUSCULAS = True` (batch numbers
  compared in upper case).
- Batch format for other profiles: config `batch_pattern` > sector `lote` (automotive =
  VIN `[A-HJ-NPR-Z0-9]{17}`) > **list mode**: `LOTE_ID is None` until
  `leer_excel_referencia` finds a column titled Lote/Lot/Batch/Charge/Los/Lotto and builds
  an alternation of the exact batch numbers (`poner_patron_lote`). If none: exit 2.
- `SECTORES` add count units and labels (GS1 `(10)`, `Ch.-B.`, `L`, VIN...). Excel titles
  in English/German via `TEXTOS_EXCEL`.
- Hook globals that keep the old behaviour by default: `FILA_TOTAL_RE`,
  `TOTAL_PALABRA_RE`, `PIEZA_RE`, `CAJA_PREFIJOS`, `EXCEL_NOD_RE`, `EXCEL_RET_RE`,
  `EXCEL_RET_EXCLUYE`, `CASI_LOTE_RE`, `RELEVANTE_RE`.
- Research sources used for the vocabulary are listed in the comment above `PERFILES`.

### 4.5 Pipeline in etapas/ (v2.3, issue #8)

`paginas.py` (pages -> JSONL, parallel + same cache) | `texto.py` (text lines) |
`cantidades.py` (analizar_documento per PDF) | `conciliar.py` (conciliar +
intentar_arreglos without alternative readings) | `excel.py` (escribir_excel). JSON is
`ensure_ascii` so it survives PowerShell 5.1 pipes. Messages go to stderr. lotes_auto.py
stays the complete program; it is unchanged by etapas/. `tools/check_examples.py` also
checks that the pipeline gives the same quantities as lotes_auto.py on the examples.

## 5. Performance work in v2.2

User report: "very very slow, basically hanged from 36/59 pdfs onward". Constraint: do
**not** remove any checks.

### Findings

- rapidocr shrinks every image to **max_side_len 2000 px** before detection, so higher
  DPI mostly costs rendering and PIL time. Recognition dominates: about 2–4 s per page
  for about 90 boxes. Detection takes about 1–2 s, the angle classifier about 0.1 s.
- **Likely cause of the "hang":** a scan saved at 72 dpi has a huge page size (about a
  meter). At 300 dpi, and ×2 again for reading 3, this makes a gigantic image. A test
  PDF like that **ran out of 7 GB of memory and was killed** in the old version; on
  Windows that would swap or freeze. Now capped: `LADO_MAX_PX = 6000` and
  `LADO_MAX_PX_AMPLIADA = 4400`.
  - Do not lower these caps. With 4400, a letter page at 400 dpi (4402 px) got resized
    and one OCR result changed.
- Changing `rec_batch_num` (6 → 12/24/64) is **slower and changes the OCR text**. Keep
  the default.
- 1 onnx thread per process × N processes gives more throughput than N threads in one
  process (about +45% on 2 cores). Results are identical across thread counts.
- `fork` workers deadlocked after the main process had used onnxruntime. Always use the
  **spawn** context (as on Windows anyway).

### What was built

- `LectorPaginas`:
  - Runs a ProcessPoolExecutor (spawn) with per-page jobs. A worker keeps one PDF open
    (`_leer_pagina`).
  - Workers are started by `_iniciar_proceso(hilos)`, which sets OMP and cv2 threads.
  - The number of processes is automatic: `procesos_automaticos()` = min(cores, 8), and
    no more than free memory / 1.2 GB.
  - A worker crash (BrokenProcessPool) makes the remaining pages run in-process, and the
    rest of the run stays in one process.
  - Watchdog: progress lines every 45 s (`AVISO_LENTO_S`). After 600 s
    (`LIMITE_PAGINA_S`) a page is abandoned with a TimeoutError, which is reported in
    revisar_a_mano.txt as "error al leer". The pool is killed and restarted for the
    remaining pages.
- **Disk cache** in `resultados_lotes/cache/`:
  - One JSON file per page and reading, named `{huella}_p{NNNN}_l{k}.json`; text pages
    use `l0`.
  - `huella_pdf` = sha1 of the file **content** + CACHE_VERSION + the DPI and cap
    settings. Renaming a PDF keeps the cache; replacing it invalidates it.
  - Stale entries are pruned at the end of a run.
  - Bump `CACHE_VERSION` when changing anything that affects the tokens (OCR, preprocess,
    DPI, deskew, the token format).
- Results are re-sorted into the original file order after the parallel reading, so the
  outputs are deterministic.
- `relaciones_suma` was rewritten to be faster, with identical output and order.

### Benchmarks

On a 2-core, 7 GB container:

| Run | Time |
|---|---|
| Old version, 13 test folders, cold | 587 s |
| v2.2, cold | about 510–560 s (small single-page fixtures; spawn start-up costs about 3 s per pool) |
| v2.2, warm cache | **10.8 s** |
| 48-page scanned PDF, old | 394 s |
| 48-page scanned PDF, v2.2 | 325 s |

All 13 test folders produce **byte-identical xlsx, detalle and revisar files** (except
the version and date header) compared with v2.1.0.

### Not yet confirmed on the real data

The user has not yet run v2.2 on the 59 real PDFs. Ask for the console lines,
especially the `... sigue leyendo` lines and the per-PDF `[N pag, X s]` timings, and the
final `Tiempo:` line.

---

## 6. Testing

Rule from the user: **test before pushing**. The previous scratch environment is
ephemeral. The test kit (`pdf-totals-pruebas.zip`) was sent to the user; ask them to
attach it if needed. It contains:

- `fixtures.py`: generates GFC, Levic, Reinvex, the recall report, the inventory and the
  merged PDF, plus scan helpers (`escanear`, `sello`, `mano`, `palomita`). It uses
  `fonts/Caveat.ttf` and DejaVuSans. Usage: `python fixtures.py OUTDIR`.
- `fixtures2.py`: generates Difarmer (2 pages), Levic Monterrey (2 pages), Tenorio, LUIN
  and the reference Excel `Recall Collins.xlsx`. It imports fixtures.py.
- `malos_escaneos.py`: faint, noisy and blurred variants (the `malos/` folder).
- `bench.sh SCRIPT OUTDIR` runs every folder, saves outputs and times (`KEEP=1` keeps
  the cache). `compara.sh A B` compares two runs.
- The test folders themselves (w2–w5, s1–s6, malos, fx, fx2) and `base/` (the outputs of
  v2.1.0).

The fixture world is small, so the real expected totals do not apply (DIFERENCIA vs
209,908 is normal). Expected results with v2.1/v2.2 on Python 3.12:

| Folder | retirado | no distrib. | total | Notes |
|---|---|---|---|---|
| w2 | 45,435 | 25,086 | 70,521 | matches the fixture Excel (sum J+K = 70,521); 1 correction |
| w4 | 45,435 | 25,086 | 70,521 | same as w2, plus a decoy `aaa_otro.xlsx` (name priority) |
| w3 | 43,735 | 25,086 | 68,821 | Tenorio renamed FALTA REVISION (−1,800, flagged with an exact-difference hint); unreadable `31-AGO ILEGIBLE`; lote misread 25240391→25240398 corrected; 2 corrections, 1 lote not matching |
| w5 | — | — | — | old `.xls` Excel → message, exit 2 |
| s1, s3, s6 | 36,066 | 24,269 | 60,335 | no Excel; nod from the report and inventory |
| s2 | 36,066 | 24,269 | 60,335 | "(copia)" duplicate marked MALO |
| s4 | 36,066 | 179 | 36,245 | no report |
| s5 | 86,066 | 179 | 86,245 | `extra 50000.pdf` → cap exceeded, flagged |
| malos | 29,654 | 0 | 29,654 | 5 unreadable; poor-scan experiment |
| fx | 112,090 | 24,269 | 136,359 | includes bad variants; 4 lotes not matching, 2 MALOS |
| fx2 | 19,121 | 25,086 | 44,207 | set 2 only; 12 lotes not matching, 1 MALO |

With Python 3.13 (rapidocr 3.x), w4 gives 45,432 (the same in the old version).

Environments used, recreate them as needed:
- a Python 3.12 venv with OCR, reportlab and pypdfium2 (the main one);
- Python 3.13 with rapidocr 3.10;
- a "BSD-like" venv without OCR or Pillow (scanned pages must give
  "escaneado, OCR no disponible");
- Python 3.9 for a syntax check (no backslashes in f-string expressions; `cancel_futures`
  needs 3.9+).

Run `python3 -m pyflakes lotes_auto.py diagnostico.py tools etapas` before committing.
`python tools/check_profiles.py` (about 30 s, one OCR page) checks the profiles.

**In-repo test (no private kit needed):** `python tools/make_examples.py --check` and
`python tools/check_examples.py` (about 6 minutes, 162 page reads, OCR on 20 scanned
PDFs). Expected: every lote matches except the one entry in `known_differences`
(25240123 retirado −954: the OCR loses the pieces cell of the one-row scanned return
form `098 ... DEVOLUCION ESCANEADA TIJUANA.pdf`). If a change to lotes_auto.py fixes it,
check_examples prints a note; then remove it from `KNOWN_DIFFERENCES` in
make_examples.py. Totals of the example world: retirado 357,499, no distribuida
129,884, total 487,383. The private kit (below) is still needed for changes to reading
or analysis, because its outputs must stay byte-identical.

Useful stress tests:
- a letter scan saved with `resolution=72` (huge page);
- 48 pages merged from the w2 PDFs;
- killing a worker (find the child with `pgrep -P <main pid>`, **not** `pgrep -f`, which
  matches your own shell and kills it);
- the watchdog test (import lotes_auto, set `LIMITE_PAGINA_S` small, call `main([...])`).

---

## 7. Lessons and things the user disliked

- Too much debugging chatter without a fix. Do the work, test it, then give one short
  answer with one command.
- Multiple Excel tabs: they want one sheet with the three tables side by side.
- Caps are **per lote**, not per table. Lotes are 25/26 only. No-distribuida comes from
  the Excel.
- Early failures:
  - an empty spreadsheet, because scanned PDFs need OCR;
  - "2,481." (trailing dot) broke a column;
  - skewed scans split rows;
  - leftover numbers in data rows counted as printed totals (rows with dates are now
    excluded);
  - quantities over 1,000,000 picked as pieces (penalty);
  - an empty reading winning the choice (verified count now comes first in `calidad`).
- Windows setup pitfalls the user hit:
  - a typo in the python.org embeddable URL (404);
  - heredoc pastes in the terminal (use Notepad);
  - `curl -O` vs `-o` (the "port number" error);
  - pip warnings about Scripts not being on PATH (harmless).
- The user asked whether the program is "self-healing" from its txt outputs: a future
  Claude can read the user's revisar_a_mano.txt, detalle_lotes.txt and
  sin_resultado.log to diagnose problems and improve the rules.

---

## 8. Open items / next steps

1. **Get feedback from the real run of v2.2** (timings, `sigue leyendo` lines, totals vs
   209,908 / 134,543 / 344,451, revisar_a_mano.txt). Fix the layouts or misreads it
   reveals, add a fixture that reproduces each one, and re-run the whole bench.
2. **Offered, no reply yet:** `correcciones_manuales.xlsx`, a file with PDF name, lote
   and pieces that the user fills in for unreadable or handwritten PDFs, applied on
   every run (and reported in revisar_a_mano.txt). This is likely needed to reach the
   exact totals, because handwriting is often unreadable.
3. Ideas for more cold-run speed without losing checks:
   - skip reading 3 when reading 2 already verifies everything on the page (this changes
     the selection semantics, so ask the user first);
   - extract embedded JPEGs directly instead of rendering (changes the OCR input; would
     need re-validation).
4. Keep README.md and docs/en/MANUAL.md in sync with any new option or config. The
   translation bot updates de/es.
5. GitHub issues: **#5** done as profiles/sectors/config (v2.3); more languages or
   sectors can be added to `VOCABULARIO`/`SECTORES` with a case in
   tools/check_profiles.py. **#8** done additively as etapas/ (lotes_auto.py was not
   split, so the user's two-file update still works). **#16** (switch time-critical
   scripts to another language): measured on the 120 examples, cold run 215 s wall,
   374 s worker CPU (OCR/render in onnxruntime, pdfium, Pillow, OpenCV: C/C++), 4.5 s
   Python in the main process; cached run 3.3 s. A rewrite cannot gain more than about
   2 %. Answered on the issue. Dependabot PRs #10-#15 were left for the user: #11-#15
   (action bumps) passed CI; #10 (Docker python 3.14) would switch the OCR engine to
   rapidocr 3.x (other results than the user's 3.12 on Windows), so keep 3.12.
6. Single-row OCR recovery: when the pieces cell of a one-row table is lost but the
   printed TOTAL DE PIEZAS and price x pieces = amount agree, the quantity could be
   recovered exactly (see the known difference in examples/). Needs a VERSION bump and
   the full bench.
7. Bandit B324 (SHA-1 in `huella_pdf`) is skipped in bandit.yaml; at the next change of
   lotes_auto.py, `hashlib.sha1(usedforsecurity=False)` (3.9+) would allow removing the
   skip. The hash value is the same, so the cache stays valid.
