# pdf-totals-Handbuch

Diese Datei wurde automatisch aus dem englischen Original mit Argos Translate übersetzt. Wenn es sich vom englischen Original unterscheidet, gilt das englische Original.

## Name

lotes_auto.py - Berechnen Sie die Gesamtzahl der Teile für jede Chargennummer in einem Ordner mit PDF-Dokumenten

diagnostico.py - zeigen, was lotes_auto.py in einem oder mehreren PDF-Dokumenten liest und erkennt

## Synopsis

```
python lotes_auto.py [-h] [--version] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] PATH
```

## Beschreibung

lotes_auto.py sucht den Ordner FOLDER und alle seine Unterordner nach Dateien, deren Name in .pdf endet. In jedem PDF findet es jede Chargennummer und die Anzahl der Stücke, die das Dokument für diese Charge aufzeichnet. Es berechnet dann die Gesamtzahl der Teile jeder Charge über alle Dokumente und schreibt eine Zeile pro Charge in eine Excel-Arbeitsmappe. Wenn eine Charge in mehreren Dokumenten oder mehrmals in einem Dokument erscheint, werden ihre Mengen addiert.

Wenn ein PDF Text enthält, liest das Programm den Text direkt. Wenn eine Seite keinen Text enthält, behandelt das Programm die Seite als gescanntes Bild und liest es mit optischer Zeichenerkennung (OCR). Wenn keine OCR-Engine installiert ist, liest das Programm keine gescannten PDFs und listet sie zur manuellen Überprüfung auf.

Das Programm wurde für Dokumente über zurückgegebene pharmazeutische Produkte in Mexiko geschrieben. Alle Nachrichten und Ausgabedateien sind in Spanisch. Dieses Handbuch erklärt ihre Bedeutung.

diagnostico.py verwendet genau den gleichen Code wie lotes_auto.py, um Dokumente zu lesen und zu analysieren. Es zeigt den Text, den es aus jedem PDF gelesen hat, was es in diesem Text erkannt hat, und das Ergebnis. Verwenden Sie es, um ein PDF zu untersuchen, das nicht richtig gelesen wird. Es schreibt keine Dateien.

## Unterstützte Layouts

Das Programm erkennt vier Layouts. Es untersucht jede Seite in der hier angegebenen Reihenfolge und verwendet das erste Layout, das Ergebnisse für diese Seite erzeugt.

1. Formulare, in denen jede Spalte ein Erzeugnis beschreibt. Eine Zeile hat eine Chargenbezeichnung wie `Lote` und enthält eine Chargennummer pro Spalte. Andere Zeilen haben eine Mengenbezeichnung wie `Cantidad Recibida` oder `Cantidad Facturada` und enthalten die Menge jedes Produkts in derselben Spalte. Wenn es mehrere Mengenreihen gibt, wählt das Programm eine nach der Priorität, die im Abschnitt Konfiguration beschrieben ist.

2. Tabellen mit Spaltentiteln. Ein Spaltentitel wie `CANTIDAD`, `PIEZAS` oder `No. de piezas` markiert die Zahlenspalte, und das Programm liest die Werte darunter. Das Programm nimmt die Chargennummer aus einer Spalte mit einem Titel wie `LOTE`. Wenn es keine solche Spalte gibt, nimmt sie die Chargennummer von einer anderen Stelle in derselben Zeile, beispielsweise aus der Beschreibung oder aus den nächsten beiden Zeilen. Ein Spaltentitel, der auf zwei Zeilen geschrieben ist, z. B. `PIEZAS` über `DEVUELTAS`, wird als ein Titel gelesen.

3. Textzeilen mit Beschriftungen, in denen die Charge an erster Stelle steht. Ein Beispiel ist eine Zeile `LOTE: 25240310` gefolgt von einer Zeile `CANTIDAD EN PIEZAS: 1,250`. Das Programm koppelt jede Charge mit der ersten Menge, die darauf folgt.

4. Textzeilen mit Etiketten, in denen die Menge an erster Stelle steht. Das Programm koppelt jede Menge mit der ersten Charge, die darauf folgt. Das erste Label im Dokument entscheidet, ob das Dokument Layout 3 oder Layout 4 verwendet.

Diese Wörter werden als Chargenbezeichnungen erkannt: `lote`, `serie`, `serial`, `n/s` und `s/n`. Sie können mit `No.`, `Num.` oder `Numero` kombiniert werden.

Diese Wörter werden als Mengenetiketten erkannt: `cantidad`, `cant.`, `piezas`, `pieza`, `pzas`, `pza`, `pzs`, `pz`, `unidades`, `unid.` und `uds`.

Diese Wörter werden auch als Zahlenspaltentitel erkannt: `devueltas`, `devuelto`, `frascos`, `ampolletas`, `ampulas` und `viales`.

Das Wort `unidad` allein markiert keine Menge, weil es die Maßeinheit benennt. Das Wort `cajas` wird nie verwendet, weil es Boxen und nicht Stücke zählt.

## Regeln für jeden Wert

- Eine Chargennummer hat genau acht Stellen und beginnt mit `24`, `25` oder `26`. Das Programm ignoriert alle anderen Zahlen, zum Beispiel Preise, Beträge, Daten und Produktcodes. Wenn die OCR andere Zeichen an eine Chargennummer anhängt, z. B. einen vertikalen Balken, entfernt das Programm sie.
- Das Programm liest Zahlen in mexikanischer Notation. Ein Komma trennt die Tausende und ein Punkt trennt die Dezimalzahlen. Zum Beispiel ist `1,250` eintausendzweihundertfünfzig, und `1,250.50` ist eintausendzweihundertfünfzigeinhalb.
- Das Programm ignoriert Mengen, die Dezimalzahlen haben, und Mengen, die Null oder negativ sind, weil eine Anzahl von Stücken immer eine ganze Zahl ist. Diese Regel verhindert auch, dass Preise als Stücke gezählt werden.
- Das Programm überspringt jede Datei, deren Name `FALTA REVISION` enthält. Das Programm ignoriert den Unterschied zwischen Großbuchstaben und Kleinbuchstaben, und das O kann einen Akzent haben. Übersprungene Dateien zählen nicht zu einer Gesamtsumme.
- Wenn eine einzelne Menge größer als `50,000` ist, fügt das Programm sie immer noch zur Summe hinzu. Wenn das Programm beendet ist, meldet es solche Mengen als verdächtig.

## Optionen

Optionen von lotes_auto.py:

- `FOLDER` ist der Ordner, der die PDF-Dateien enthält. Unterordner sind enthalten.
- `--version` zeigt die Versionsnummer an. Das Programm stoppt dann.
- `-h` oder `--help` zeigt einen kurzen Hilfetext. Das Programm stoppt dann.

Optionen von diagnostico.py:

- `PATH` ist eine einzelne PDF-Datei oder ein Ordner. Wenn es sich um einen Ordner handelt, untersucht das Programm alle darin enthaltenen PDF-Dateien.
- `--lineas N` zeigt nicht mehr als N Textzeilen für jedes PDF. Der Standardwert ist 40.
- `--version` zeigt die Versionsnummer an. Das Programm stoppt dann.
- `-h` oder `--help` zeigt einen kurzen Hilfetext. Das Programm stoppt dann.

## Ausgabedateien

lotes_auto.py schreibt die folgenden Dateien in den Ordner FOLDER. Es ersetzt bestehende Dateien mit den gleichen Namen. Wenn eine Datei leer wäre, löscht das Programm sie stattdessen.

`totales_lotes.xlsx` enthält das Ergebnis. Es hat ein Blatt mit dem Namen `Lotes`. Das Blatt hat zwei Spalten, `Lote` und `Cantidad en Piezas`, und eine Zeile für jede Charge. Die Zeilen werden nach Chargennummer sortiert. Die Mengen werden als Zahlen gespeichert und mit Zifferngruppierung, beispielsweise `1,406`, dargestellt.

`detalle_lotes.txt` listet jeden Betrag auf, der zu jeder Charge hinzugefügt wurde, zusammen mit dem Namen des PDF, das den Betrag enthält. Verwenden Sie diese Datei, um die Quelle einer Summe zu finden, die falsch aussieht.

`revisar_a_mano.txt` hat zwei Abschnitte. Der erste Abschnitt listet die Dateien auf, die übersprungen wurden, weil ihr Name `FALTA REVISION` enthält. Der zweite Abschnitt listet die PDFs auf, aus denen keine Menge gelesen werden konnte, jeweils mit dem Grund. Die möglichen Gründe sind: keine lesbaren Mengen; ein gescanntes PDF und keine OCR-Engine; ein Fehler beim Lesen der Datei.

`sin_resultado.log` enthält für jedes PDF, das keine Menge produziert hat, den Grund und den vollständigen Text, den das Programm daraus gelesen hat. Schauen Sie sich diese Datei zuerst an, wenn ein PDF nicht richtig gelesen wird.

## Konsolennachrichten

Während es läuft, zeigt lotes_auto.py eine Zeile für jedes PDF an. Der Text nach dem Pfeil hat folgende Bedeutung:

- `-> N cantidades (texto)` bedeutet, dass das Programm N Mengen aus einem PDF liest, das Text enthält.
- `-> N cantidades (OCR)` bedeutet, dass das Programm N Mengen aus einem gescannten PDF gelesen hat.
- `-> 0 cantidades` bedeutet, dass das Programm keine Menge gelesen hat. Die Datei sin_resultado.log zeigt den Text, den sie gelesen hat.
- `-> omitido (FALTA REVISION)` bedeutet, dass das Programm die Datei aufgrund ihres Namens übersprungen hat.
- `-> escaneado, OCR no disponible` bedeutet, dass das PDF gescannt wird und keine OCR-Engine installiert ist.
- `-> ERROR` bedeutet, dass das Programm die Datei nicht öffnen oder lesen konnte.

Am Ende zeigt das Programm die Anzahl der Chargen und Mengen und die Namen der Dateien, die es geschrieben hat. Wenn eine einzelne Menge größer als `50,000` ist, zeigt sie auch eine Liste dieser Mengen unter dem Titel `cantidades sospechosamente grandes` an.

## Rückgabewerte des Programms

- `0` bedeutet, dass das Programm erfolgreich abgeschlossen wurde.
- `1` bedeutet, dass das Programm die Excel-Datei nicht schreiben konnte. Normalerweise ist die Datei in Excel geöffnet.
- `2` bedeutet, dass die Argumente nicht gültig sind oder dass der Ordner nicht existiert.

## Konfiguration

Die Einstellungen befinden sich am Anfang der Datei lotes_auto.py. Sie können sie mit jedem Texteditor ändern. Speichern Sie die Datei in der UTF-8-Codierung.

`LOTE_WORDS` und `CANT_WORDS` sind reguläre Ausdrücke. Sie enthalten die Wörter, die Chargennummern und -mengen markieren.

`PRIORIDAD_CANTIDAD` entscheidet, welche Menge Spalte oder Zeile verwendet wird, wenn ein Dokument mehrere hat. Das Programm verwendet das erste Wort in dieser Liste, das in einem Spaltentitel erscheint. Die Standardliste lautet `pieza`, `devuel`, `frasco`, `ampollet`, `ampula`, `vial`, `recibid`, `cantidad`, `factur`. Dies bedeutet, dass eine Spalte über Stücke oder über zurückgegebene Ware bevorzugt wird und dass eine fakturierte Menge nur verwendet wird, wenn keine andere Menge existiert.

`AVISO_CANTIDAD_GRANDE` ist der Grenzwert, oberhalb dessen eine einzelne Menge als verdächtig gemeldet wird. Der Standardwert ist `50000`.

`OMITIR_NOMBRE` ist das Muster für die Namen von Dateien, die übersprungen werden. Der Standardwert stimmt mit `FALTA REVISION` überein.

`LOTE_ID` ist das Muster für eine gültige Chargennummer. Der Standardwert akzeptiert genau acht Ziffern, die mit `24`, `25` oder `26` beginnen. Um beispielsweise auch Chargennummern zu akzeptieren, die mit `27` beginnen, ändern Sie `2[456]` in `2[4567]`.

## Installation

### Allgemeine Informationen

Das Programm benötigt Python und die Python-Pakete pdfplumber und openpyxl. Um gescannte PDFs zu lesen, benötigt es auch eine OCR-Engine:

- Mit `Python 3.12` oder einer älteren Version: das Paket rapidocr_onnxruntime. Dies ist die bevorzugte OCR-Engine.
- Mit `Python 3.13` oder einer neueren Version: die Pakete rapidocr und onnxruntime. In einem Test hat diese OCR-Engine eine Ziffer falsch gelesen. Die bevorzugte OCR-Engine liest die gleiche Ziffer korrekt. Aus diesem Grund wird `Python 3.12` empfohlen, wenn Sie viele gescannte PDFs verarbeiten.

Für FreeBSD, OpenBSD, NetBSD oder Alpine Linux ist keine OCR-Engine verfügbar. Auf diesen Systemen liest das Programm nur PDFs, die Text enthalten.

Die folgenden Anweisungen legen alles in einen Ordner mit dem Namen `pdf-totals` in Ihrem Home-Ordner. Die Python-Pakete werden in einer privaten Python-Umgebung in `pdf-totals/venv` oder in `pdf-totals\python` unter Windows installiert. Nichts ist für das gesamte System installiert, außer Python selbst, wo dies angegeben ist.

### Windows 10 und Windows 11

Sie benötigen keine Administratorrechte. Öffnen Sie PowerShell und führen Sie den folgenden Befehl aus. Der Befehl lädt die einbettbare Verteilung von `Python 3.12` von python.org in Ihren Benutzerordner herunter. Es installiert dann pip und die erforderlichen Pakete und lädt die beiden Skripte herunter. Windows PowerShell 5.1 unterstützt den `&&`-Operator nicht, daher sind die Schritte mit dem `;`-Zeichen getrennt.

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Das Paket msvc-runtime kopiert die Microsoft Visual C++-Laufzeitdateien in den Python-Ordner. Die OCR-Engine benötigt diese Dateien, und einige Systeme haben sie nicht. Sie können die pip-Warnungen über Skripte ignorieren, die in einem Ordner installiert sind, der sich nicht auf PATH befindet.

Verwenden Sie `curl.exe` und nicht `curl`. In Windows PowerShell 5.1 ist `curl` ein anderer Name für `Invoke-WebRequest`, der verschiedene Optionen hat. Windows 10 enthält `curl.exe` seit Version 1803.

Um das Programm auszuführen:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

Wenn Downloads auf dem Computer blockiert sind, laden Sie die Dateien `python-3.12.10-embed-amd64.zip` und `get-pip.py` auf einem anderen Computer herunter. Kopieren Sie sie in den Downloads-Ordner. Ersetzen Sie dann die beiden `Invoke-WebRequest`-Schritte im Befehl durch die Pfade dieser Dateien.

Diese Anweisungen sind für 64-Bit-Prozessoren von Intel und AMD. Verwenden Sie für Windows auf ARM-Prozessoren stattdessen `python-3.12.10-embed-arm64.zip`. Dies wurde nicht getestet.

### macOS

Sie benötigen nur Administratorrechte, um Python zu installieren. Installieren Sie `Python 3.12` mit dem Installer von python.org (https://www.python.org/downloads/macos/) oder mit Homebrew (`brew install python@3.12`). Führen Sie dann den folgenden Befehl in der Terminal-Anwendung aus:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Wenn Sie Homebrew verwenden und der Befehl `python3.12` nicht gefunden wird, verwenden Sie stattdessen `$(brew --prefix)/bin/python3.12`. Macs mit Intel-Prozessoren und Macs mit Apple-Prozessoren werden beide unterstützt.

Um das Programm auszuführen:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

### Linux

Unter Linux wird die OCR-Engine mit dem Paket opencv-python-headless anstelle von opencv-python installiert. Das Paket opencv-python-headless benötigt nicht die Grafikbibliotheken des Systems, so dass Sie keine zusätzlichen Systempakete installieren müssen.

Wenn `Python 3.12` als Befehl `python3.12` verfügbar ist, erstellt der folgende Befehl die Umgebung für jede Distribution. Verwenden Sie in Debian 12 stattdessen den Befehl `python3`, der `Python 3.11` ist.

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Um das Programm auszuführen:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

In den folgenden Absätzen wird erläutert, wie Sie `Python 3.12` in gängigen Linux-Distributionen erhalten.

Debian 12 und Ubuntu 24.04 enthalten eine geeignete Version von Python. Installieren Sie das venv Modul und das curl Programm. Dies erfordert den root-Benutzer:

```
sudo apt install python3-venv curl
```

Auf Ubuntu 24.04 lautet der Befehl `python3.12`. Auf Debian 12 lautet der Befehl `python3`, was `Python 3.11` ist. In diesem Fall ersetzen Sie `python3.12` mit `python3` im Befehl oben. Debian 13 und neuere Versionen von Ubuntu enthalten `Python 3.13` oder eine neuere Version. Verwenden Sie auf diesen Systemen die unten beschriebene Methode mit uv.

Fedora bietet ein zusätzliches Paket für `Python 3.12`:

```
sudo dnf install python3.12
```

openSUSE Leap 15.6 und openSUSE Tumbleweed bieten ein zusätzliches Paket für `Python 3.12`:

```
sudo zypper install python312
```

Arch Linux bietet nur die neueste Version von Python. Verwenden Sie die Methode mit uv, die unten beschrieben wird.

Auf jeder Linux-Distribution, die glibc verwendet, können Sie uv verwenden und benötigen den root-Benutzer nicht. uv ist ein einzelnes Programm. Es ist in Ihrem Home-Ordner installiert und kann eine private Kopie von `Python 3.12` herunterladen.

```
curl -LsSf https://astral.sh/uv/install.sh | sh && mkdir -p ~/pdf-totals && cd ~/pdf-totals && ~/.local/bin/uv venv -p 3.12 venv && ~/.local/bin/uv pip install -p venv/bin/python pdfplumber openpyxl rapidocr_onnxruntime && ~/.local/bin/uv pip uninstall -p venv/bin/python opencv-python && ~/.local/bin/uv pip install -p venv/bin/python opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Alpine Linux verwendet die C-Bibliothek musl. Für musl ist keine OCR-Engine verfügbar, daher liest das Programm nur PDFs, die Text enthalten:

```
doas apk add python3 curl && mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv venv && venv/bin/pip install pdfplumber openpyxl && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Wenn Ihr Alpine Linux-System `sudo` anstelle von `doas` verwendet, ersetzen Sie `doas` durch `sudo`.

### FreeBSD, OpenBSD und NetBSD

Auf diesen Systemen liest das Programm nur PDFs, die Text enthalten. Die OCR-Engine ist für diese Systeme nicht verfügbar. Die PDF-Rendering-Bibliothek, die pdfplumber verwendet, ist ebenfalls nicht verfügbar, aber pdfplumber kann Text ohne sie lesen.

Das Paket pdfminer.six benötigt das Paket cryptography. Ein Rust-Compiler wird benötigt, um cryptography aus dem Quellcode zu erstellen, sodass cryptography aus den Paketen des Systems installiert wird. Alle anderen Pakete sind in reinem Python geschrieben. Sie werden ohne ihre optionalen Abhängigkeiten installiert.

Führen Sie unter FreeBSD den ersten Befehl als root-Benutzer aus, beispielsweise mit `su` oder `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Die Paketnamen enthalten die Version von Python, die in diesem Beispiel `311` ist. Wenn Ihr System eine andere Version verwendet, finden Sie die richtigen Namen mit dem Befehl `pkg search cryptography`. Verwenden Sie dann den Matching-Befehl anstelle von `python3.11`, zum Beispiel `python3.12`.

Führen Sie auf OpenBSD den ersten Befehl als root-Benutzer mit `doas` aus:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Führen Sie auf NetBSD den ersten Befehl als root-Benutzer aus:

```
pkgin install py312-pip py312-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Wie bei FreeBSD enthalten die Paketnamen die Version von Python, die in diesem Beispiel `312` ist. Finden Sie die richtigen Namen auf Ihrem System mit dem Befehl `pkgin search cryptography`.

Um das Programm auf FreeBSD, OpenBSD oder NetBSD auszuführen:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

## Aktualisierung

Um das Programm zu aktualisieren, laden Sie die beiden Skripte erneut herunter. Unter Windows:

```
curl.exe -L -o "$env:USERPROFILE\pdf-totals\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$env:USERPROFILE\pdf-totals\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Auf macOS und Linux:

```
cd ~/pdf-totals && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Verwenden Sie auf FreeBSD `fetch` auf die gleiche Weise. Verwenden Sie auf OpenBSD und NetBSD `ftp`.

Nach einer Änderung kann GitHub die vorherige Version für einige Minuten liefern. Um sofort eine bestimmte Version zu erhalten, ersetzen Sie `main` in der Adresse durch die Kennung des Commits.

## Fehlerbehebung

Die Excel-Datei kann nicht gespeichert werden. Das Programm endet mit dem Rückgabecode 1, wenn totales_lotes.xlsx in Excel geöffnet ist. Schließen Sie die Datei und führen Sie das Programm erneut aus.

Alle oder viele PDFs zeigen `0 cantidades`. Führen Sie diagnostico.py mit einem dieser PDFs aus. Wenn der angezeigte Text leer ist, wird das PDF gescannt und die OCR liest es nicht. Wenn der Text vorhanden ist, aber nichts erkannt wird, kennt das Programm die Beschriftungen oder das Layout des Dokuments nicht. Der Text zeigt, welche Wörter das Dokument verwendet. Der gleiche Text ist in sin_resultado.log. Sie können diese Wörter zur Konfiguration hinzufügen.

Eine Summe ist viel zu groß. Suchen Sie nach der Chargennummer in detalle_lotes.txt. Diese Datei zeigt jeden Betrag, der hinzugefügt wurde, und den Namen der PDF, die es enthält. Dies sind häufige Ursachen:

- Das Programm verwendete die falsche Spalte für die Menge, da das Dokument einen Spaltentitel verwendet, den das Programm nicht kennt.
- Die OCR verknüpfte zwei Zahlen, die nahe beieinander liegen, zu einer Zahl.
- Das gleiche PDF befindet sich zweimal im Ordner.

Die OCR kann unter Windows nicht geladen werden, und die Nachricht enthält `DLL load failed`. Installieren Sie das Paket msvc-runtime wie in den Anweisungen für Windows gezeigt.

Die OCR kann nicht unter Linux geladen werden, und die Nachricht enthält `libGL.so.1`. Ersetzen Sie opencv-python durch opencv-python-headless, wie in den Anweisungen für Linux gezeigt.

pip kann rapidocr_onnxruntime nicht installieren. Dieses Paket unterstützt Python-Versionen bis `Python 3.12`. Verwenden Sie `Python 3.12` oder installieren Sie stattdessen die Pakete rapidocr und onnxruntime.

## Beschränkungen

Die OCR liest handschriftliche Zahlen, wenn sie klar geschrieben sind. Es liest nicht zuverlässig Handschrift, die schwer zu lesen ist, Häkchen oder handschriftliche Abkürzungen wie `N/A`. PDFs mit Mengen, die nicht gelesen werden können, sind in revisar_a_mano.txt aufgeführt und müssen von einer Person überprüft werden.

Die OCR kann eine Ziffer falsch lesen, insbesondere wenn ein Scan verschwommen oder gekippt ist. Vergleichen Sie einige Summen aus gescannten PDFs mit detalle_lotes.txt.

Das Format einer Chargennummer wird durch `LOTE_ID` festgelegt. Das Programm ignoriert Chargennummern in anderen Formaten, bis das Muster geändert wird.

Das Programm erwartet, dass Mengen ganze Stückzahlen sind. Es konvertiert keine Dokumente, die Boxen oder andere Einheiten zählen.

## Dateien im Projekt

- lotes_auto.py ist das Programm.
- diagnostico.py ist das Diagnosewerkzeug. Es muss sich im gleichen Ordner wie lotes_auto.py befinden.
- README.md ist eine kurze Beschreibung. Es ist auch als README.de.md und README.es.md verfügbar.
- Der Ordner `docs` enthält dieses Handbuch in Englisch, Deutsch und Spanisch.
- `tools/translate_docs.py` ist das Skript, das die Dokumentation übersetzt.
- `LICENSE` enthält den Text der Lizenz.

## Lizenz

MIT-Lizenz. Siehe die Datei LICENSE.
