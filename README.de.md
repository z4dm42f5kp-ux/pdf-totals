# pdf-totals

Diese Datei wurde automatisch aus dem englischen Original mit Argos Translate übersetzt. Wenn es sich vom englischen Original unterscheidet, gilt das englische Original.

Andere Sprachen: [Deutsch](README.de.md), [Español](README.es.md)]

pdf-totals liest einen Ordner mit PDF-Dokumenten wie Rechnungen, Lieferscheinen und Produktrückgabeformularen, findet jede Chargennummer und die Anzahl der dafür aufgezeichneten Teile und schreibt die Gesamtzahl der Teile pro Charge in eine Excel-Datei.

Es liest PDFs, die Text direkt enthalten, und liest gescannte PDFs mit optischer Zeichenerkennung (OCR). Es versteht mehrere gängige Layouts: Tabellen mit einer Chargenspalte und einer Mengenspalte, Tabellen, in denen die Chargennummer in der Beschreibung erscheint, Formulare, in denen jede Spalte ein Produkt ist, und beschriftete Zeilen wie `LOTE: 25240310` gefolgt von `CANTIDAD: 1,250`.

Das Programm wurde für mexikanische pharmazeutische Rückgabedokumente geschrieben. Seine Nachrichten und Ausgabedateien sind in Spanisch. Es wird erwartet, dass die Losnummern genau acht Ziffern haben und mit 24, 25 oder 26 beginnen; diese Regel ist leicht zu ändern.

## Anforderungen

- Python 3.9 oder höher. Python 3.12 wird empfohlen, da die bevorzugte OCR-Engine Python bis zu 3.12. unterstützt.
- Die Python Pakete pdfplumber und openpyxl.
- Für gescannte PDFs: rapidocr_onnxruntime (Python 3.12 oder früher) oder rapidocr zusammen mit onnxruntime (Python 3.13 oder später).

Administratorrechte sind unter Windows oder macOS nicht erforderlich. Unter Linux und den BSDs erfordert die Installation von Python aus den Systempaketen root; unter Linux kann dies mit dem uv-Tool vermieden werden, wie im Handbuch beschrieben. OCR ist unter Windows, macOS und glibc-basiertem Linux verfügbar. Auf FreeBSD, OpenBSD, NetBSD und Alpine Linux können nur PDFs gelesen werden, die Text enthalten; gescannte PDFs werden zur manuellen Überprüfung aufgelistet.

## Schnellstart

Windows 10 und 11 in PowerShell, ohne Administratorrechte:

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

macOS, mit Python 3.12 installiert:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Linux, mit Python 3.12 installiert:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Die Installation auf anderen Linux-Distributionen und auf den BSDs wird im Handbuch beschrieben.

## Nutzung

Windows:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

macOS, Linux und BSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

Der Ordner wird rekursiv durchsucht. Dateien, deren Name `FALTA REVISION` enthält, werden übersprungen. Wenn das Programm beendet ist, enthält der Ordner:

- `totales_lotes.xlsx`: eine Reihe pro Los mit der Gesamtzahl der Stücke.
- `detalle_lotes.txt`: jeder Betrag, der zu jedem Lot hinzugefügt wurde, und das PDF, aus dem es stammt.
- `revisar_a_mano.txt`: PDFs, die übersprungen wurden und PDFs, aus denen keine Menge gelesen werden konnte.
- `sin_resultado.log`: Der Text, der aus jedem PDF gelesen wurde, das keine Menge produzierte.

Um zu sehen, was das Programm in einem einzelnen PDF liest und erkennt, ohne Dateien zu schreiben, führen Sie `diagnostico.py` mit dem Pfad des PDF aus.

## Dokumentation

Das Handbuch beschreibt jede Option, die unterstützten Layouts, die Ausgabedateien, die Konfigurationseinstellungen und die Installation auf jeder Plattform:

- Englisch: [docs/en/MANUAL.md](docs/en/MANUAL.md)]
- Deutsch: [docs/de/MANUAL.md](docs/de/MANUAL.md)]
- Español: [docs/es/MANUAL.md](docs/es/MANUAL.md)]

Die deutschen und spanischen Dokumente werden automatisch aus den englischen Originalen mit Argos Translate übersetzt. Wenn sie nicht einverstanden sind, sind die englischen Dokumente autoritativ.

## Lizenz

MIT. Siehe die Datei LICENSE.
