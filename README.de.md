# tute-lote

Diese Datei wurde automatisch aus dem englischen Original mit Argos Translate übersetzt. Wenn es sich vom englischen Original unterscheidet, gilt das englische Original.

Sprachen: [English](README.md), [Deutsch](README.de.md), [Español](README.es.md)

tute-lote liest einen Ordner mit PDF-Dokumenten über die Rücknahme eines Produkts vom Markt, wie Rückgabeformulare, Lieferbelege, Warenbelege, Inventarlisten und Rücknahmeberichte. Für jede Chargennummer berechnet es drei Mengen und schreibt sie in eine Excel-Datei:

- die aus dem Markt genommene Menge: Stücke, die zurückgegeben oder vom Markt zurückgewonnen wurden;
- die nicht verteilte Menge: Stücke, die sich noch im Bestand befinden, entnommen aus einer Excel-Referenzdatei;
- die Gesamtmenge für die Zerstörung: die Summe beider.

Das Programm verlässt sich nicht nur auf die Wörter in den Dokumenten. Es untersucht die Struktur jeder Tabelle: welche Spalte ganze Zahlen enthält, welche Spalte multipliziert mit dem Stückpreis den gedruckten Gesamtpreis ergibt und welche Spalte sich zu einer gedruckten Gesamtsumme addiert. In einem Rückzugsbericht findet es die Spalten nach Arithmetik, weil hergestellt minus verteilt gleich nicht verteilt ist, und nicht verteilt plus zurückgezogen gleich der Summe.

Die zurückgezogene Menge jeder Charge wird mit der Excel-Referenzdatei verglichen. Keine einzelne Charge darf entfernte `48,042`-Stücke oder nicht verteilte `82,090`-Stücke überschreiten. Die zurückgezogenen Stücke müssen sich zu genau `209,908` addieren, die nicht verteilten Stücke zu genau `134,543` und die Summe für die Zerstörung zu genau `344,451`. Wenn etwas nicht übereinstimmt, werden zweifelhafte gescannte PDFs ein zweites Mal gelesen, Differenzen werden nur dann korrigiert, wenn ein exakter Beweis vorliegt, und PDFs werden nur nach schlüssigen Beweisen als schlecht markiert.

Das Programm liest den Text von PDFs, die Text enthalten, und liest gescannte PDFs mit optischer Zeichenerkennung (OCR). Es wurde für Dokumente über zurückgegebene pharmazeutische Produkte in Mexiko geschrieben. Seine Nachrichten und Ausgabedateien sind in Spanisch. Eine Chargennummer muss genau acht Ziffern haben und mit `25` oder `26` beginnen. Diese Regeln sind leicht zu ändern.

## Anforderungen

- Python 3.9 oder eine neuere Version. `Python 3.12` wird empfohlen, da die bevorzugte OCR-Engine Python-Versionen bis `Python 3.12` unterstützt.
- Die Python-Pakete pdfplumber und openpyxl.
- Für gescannte PDFs: das Paket rapidocr_onnxruntime mit `Python 3.12` oder einer älteren Version oder die Pakete rapidocr und onnxruntime mit `Python 3.13` oder einer neueren Version.

Unter Windows und macOS benötigen Sie keine Administratorrechte. Unter Linux und auf den BSD-Systemen muss der root-Benutzer Python aus den Paketen des Systems installieren. Unter Linux können Sie dies mit dem Programm uv vermeiden, wie das Handbuch beschreibt.

OCR ist unter Windows, macOS und Linux-Distributionen verfügbar, die glibc verwenden. Auf FreeBSD, OpenBSD, NetBSD und Alpine Linux liest das Programm nur PDFs, die Text enthalten. Es listet gescannte PDFs zur manuellen Überprüfung auf.

## Schnellstart

Windows 10 und Windows 11 in PowerShell, ohne Administratorrechte:

```
$d="$env:USERPROFILE\tute-lote"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

macOS, wenn `Python 3.12` installiert ist:

```
mkdir -p ~/tute-lote && cd ~/tute-lote && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

Linux, wenn `Python 3.12` installiert ist:

```
mkdir -p ~/tute-lote && cd ~/tute-lote && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/tute-lote/main/diagnostico.py
```

Das Handbuch beschreibt die Installation auf anderen Linux-Distributionen und auf FreeBSD, OpenBSD und NetBSD.

## Wie man das Programm benutzt

Windows:

```
& "$env:USERPROFILE\tute-lote\python\python.exe" "$env:USERPROFILE\tute-lote\lotes_auto.py" "C:\path\to\pdfs"
```

macOS, Linux, FreeBSD, OpenBSD und NetBSD:

```
~/tute-lote/venv/bin/python ~/tute-lote/lotes_auto.py /path/to/pdfs
```

Das Programm durchsucht auch alle Unterordner. Es überspringt jede Datei, deren Name `FALTA REVISION` enthält. Die Excel-Referenzdatei ist die Datei `Retirado_retiro_producto_Traxicoll.xlsx` im Ordner oder in einem Unterordner. Wenn das Programm beendet ist, enthält der Ordner `resultados_lotes` im Ordner diese Dateien:

- totales_lotes.xlsx enthält drei Tabellen nebeneinander auf einem Blatt: zurückgezogen, nicht verteilt und insgesamt zur Zerstörung, jeweils mit der Chargennummer und der Stückzahl.
- detalle_lotes.txt listet jede Menge jeder Charge, das PDF, das sie enthält und ob sie verifiziert ist.
- revisar_a_mano.txt listet die als schlecht markierten PDFs, die Korrekturen, die Chargen, die nicht übereinstimmen, die verdächtigen PDFs und die PDFs, aus denen keine Menge gelesen werden konnte.
- sin_resultado.log enthält den Text, der aus jedem PDF gelesen wurde, das keine Menge produziert hat.
- Der Ordner-Cache enthält, was von jeder Seite gelesen wurde. Der nächste Lauf verwendet es, also dauert es Sekunden statt Minuten.

Das Programm liest mehrere Seiten gleichzeitig, eine für jeden Prozessorkern. Die Option `--procesos 1` liest eine Seite nach der anderen. Die Option `--sin-cache` liest jede Seite erneut.

Um zu sehen, was das Programm in einem PDF liest und erkennt, führen Sie diagnostico.py mit dem Pfad des PDF aus. diagnostico.py schreibt keine Dateien.

## Dokumentation

Das Handbuch beschreibt alle Optionen, die unterstützten Layouts, die Ausgabedateien, die Einstellungen und die Installation auf jedem System:

- Englisch: [docs/en/MANUAL.md](docs/en/MANUAL.md)
- Deutsch: [docs/de/MANUAL.md](docs/de/MANUAL.md)
- Español: [docs/es/MANUAL.md](docs/es/MANUAL.md)

Die deutschen und spanischen Dokumente werden automatisch aus den englischen Dokumenten mit Argos Translate übersetzt. Wenn eine Übersetzung vom englischen Dokument abweicht, gilt das englische Dokument.

## Beispieldokumente und Tests

Der Ordner `examples` enthält 120 erfundene PDF-Dokumente, eine Excel-Referenzdatei und das erwartete Ergebnis. Verwenden Sie sie, um das Programm ohne echte Dokumente auszuprobieren. Der Befehl `python tools/check_examples.py` führt das Programm auf diesen Dokumenten aus und vergleicht das Ergebnis mit dem erwarteten Ergebnis. Die automatischen Tests auf GitHub führen nach jeder Änderung den gleichen Befehl aus.

## Freigaben und Sicherheit

`RELEASES.txt` beschreibt die stabilen und rollenden Releases und wie man eine bestimmte Version herunterlädt. `SECURITY.md` beschreibt, wie man eine Schwachstelle meldet und wie das Programm vertrauliche Dokumente schützt. `AGENTS.md` enthält die Regeln für KI-Agenten, die das Programm ausführen oder ändern, einschließlich der Regeln für Anweisungen, die in Dokumenten verborgen sind.

## Lizenz

MIT-Lizenz. Siehe die Datei LICENSE.
