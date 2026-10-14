# pdf-totals-Handbuch

Diese Datei wurde automatisch aus dem englischen Original mit Argos Translate übersetzt. Wenn es sich vom englischen Original unterscheidet, gilt das englische Original.

## Name

lotes_auto.py - Addieren Sie die Anzahl der Stücke pro Chargennummer in einem Ordner von PDF-Dokumenten

diagnostico.py - zeigen, was lotes_auto.py in einem oder mehreren PDF-Dokumenten liest und erkennt

## Synopsis

```
python lotes_auto.py [-h] [--version] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] PATH
```

## Beschreibung

lotes_auto.py sucht FOLDER und alle seine Unterordner nach Dateien, die in .pdf enden. Aus jedem PDF liest es jede Chargennummer und die Anzahl der Stücke, die für diese Charge aufgezeichnet wurden. Es addiert dann die Teile jeder Charge über alle Dokumente und schreibt eine Zeile pro Charge in eine Excel-Arbeitsmappe. Batches, die in mehreren Dokumenten oder mehrmals in einem Dokument erscheinen, werden zusammengefügt.

Der Text einer PDF wird direkt gelesen, wenn die PDF Text enthält. Wenn eine Seite keinen Text enthält, wird sie als gescanntes Bild behandelt und mit optischer Zeichenerkennung (OCR) gelesen. Wenn keine OCR-Engine installiert ist, werden gescannte PDFs nicht gelesen und zur manuellen Überprüfung aufgelistet.

Das Programm wurde für Dokumente über zurückgegebene pharmazeutische Produkte in Mexiko geschrieben. Alle Nachrichten und Ausgabedateien sind in Spanisch; ihre Bedeutung wird in diesem Handbuch erklärt.

diagnostico.py verwendet genau den gleichen Lese- und Erkennungscode wie lotes_auto.py. Es druckt den Text, der aus jedem PDF gelesen wurde, was darin erkannt wurde und das Ergebnis, so dass ein PDF, das nicht richtig gelesen wird, untersucht werden kann. Es schreibt keine Dateien.

## Unterstützte Layouts

Das Programm erkennt vier Layouts. Jede Seite wird in der hier angegebenen Reihenfolge untersucht, und das erste Layout, das Ergebnisse liefert, wird für diese Seite verwendet.

1. Formulare, in denen jede Spalte ein Produkt ist. Eine Zeile, die mit einem Batchwort wie `Lote` gekennzeichnet ist, enthält eine Chargennummer pro Spalte. Andere Zeilen, z. B. `Cantidad Recibida` oder `Cantidad Facturada`, enthalten die Menge jedes Produkts in derselben Spalte. Wenn mehrere Mengenreihen vorhanden sind, wird die Zeile durch die unter Konfiguration beschriebene Priorität ausgewählt.

2. Tabellen mit einer Überschrift. Eine Überschrift wie `CANTIDAD`, `PIEZAS` oder `No. de piezas` markiert die Zahlenspalte, und die Werte darunter werden gelesen. Die Chargennummer wird einer Überschrift wie `LOTE` entnommen, falls vorhanden, ansonsten von irgendwo in derselben Zeile, z. B. innerhalb der Beschreibung, oder von den nächsten zwei Zeilen. Überschriften, die sich über zwei Zeilen erstrecken, wie `PIEZAS` über `DEVUELTAS`, werden als eine Überschrift gelesen.

3. Beschriftete Zeilen, in denen der Batch zuerst kommt, beispielsweise eine Zeile `LOTE: 25240310` gefolgt von einer Zeile `CANTIDAD EN PIEZAS: 1,250`. Jede Charge wird mit der ersten Menge gepaart, die ihr folgt.

4. Beschriftete Zeilen, in denen die Menge an erster Stelle steht. Jede Menge wird mit der ersten Charge gepaart, die ihr folgt. Ob ein Dokument Layout 3 oder Layout 4 verwendet, entscheidet sein erstes Label.

Wörter, die als Batch-Etiketten erkannt werden: `lote`, `serie`, `serial`, `n/s`, `s/n`, optional vorangestellt oder gefolgt von `No.`, `Num.` oder `Numero`. Wörter, die als Quantitätsetiketten erkannt werden: `cantidad`, `cant.`, `piezas`, `pieza`, `pzas`, `pza`, `pzs`, `pz`, `unidades`, `unid.`, `uds`. Wörter, die in Tabellen als Mengenüberschriften erkannt werden, umfassen auch `devueltas`, `devuelto`, `frascos`, `ampolletas`, `ampulas` und `viales`. Das Wort `unidad` allein ist keine Menge, weil es die Maßeinheit benennt. `Cajas` wird nie verwendet, weil es Boxen zählt, nicht Stücke.

## Regeln für jeden Wert

- Eine Chargennummer muss aus genau acht Ziffern bestehen und mit `24`, `25` oder `26` beginnen. Alles andere wird ignoriert, einschließlich Preise, Beträge, Daten und Produktcodes. Zeichen, die von der OCR an eine Chargennummer angehängt werden, z. B. ein vertikaler Balken, werden entfernt.
- Zahlen werden in der mexikanischen Notation gelesen, mit einem Komma als Tausenden Trennzeichen und einem Punkt als Dezimaltrennzeichen: `1,250` ist eintausendzweihundertfünfzig und `1,250.50` ist eintausendzweihundertfünfzigeinhalb.
- Mengen mit Dezimalstellen und Mengen, die Null oder negativ sind, werden ignoriert, weil Stücke ganze Zahlen sind. Dies verhindert auch, dass Preise als Stücke gezählt werden.
- Dateien, deren Name `FALTA REVISION` enthält, in Groß- oder Kleinbuchstaben und mit oder ohne Akzent auf dem O, werden übersprungen und zählen nicht zu einer Gesamtsumme.
- Eine einzelne Größe größer als `50,000` wird immer noch hinzugefügt, wird aber am Ende des Laufs als verdächtig gemeldet.

## Optionen

Optionen von lotes_auto.py:

- `FOLDER`: Der Ordner, der die PDF-Dateien enthält. Unterordner sind enthalten.
- `--version`: Drucken Sie die Versionsnummer und den Ausgang.
- `-h`, `--help`: Drucken Sie eine kurze Nutzungsnachricht und Exit.

Optionen von diagnostico.py:

- `PATH`: eine einzelne PDF-Datei oder ein Ordner, dessen PDF-Dateien alle untersucht werden.
- `--lineas N`: höchstens N Textzeilen pro PDF anzeigen. Der Default ist 40.
- `--version`: Drucken Sie die Versionsnummer und den Ausgang.
- `-h`, `--help`: Drucken Sie eine kurze Nutzungsnachricht und Exit.

## Ausgabedateien

lotes_auto.py schreibt die folgenden Dateien in FOLDER. Bestehende Dateien mit den gleichen Namen werden ersetzt. Eine Datei, die leer wäre, wird stattdessen entfernt.

`totales_lotes.xlsx` ist das Ergebnis. Es hat ein Blatt namens `Lotes` mit zwei Spalten, `Lote` und `Cantidad en Piezas`, eine Zeile pro Charge, sortiert nach Chargennummer. Die Mengen werden als Zahlen gespeichert und mit Tausenden Separatoren angezeigt.

`detalle_lotes.txt` listet für jeden Batch jeden Betrag auf, der hinzugefügt wurde, und das PDF, aus dem der Betrag stammt. Verwenden Sie es, um die Quelle einer Summe zu finden, die falsch aussieht.

`revisar_a_mano.txt` hat zwei Abschnitte. Die erste listet die Dateien auf, die übersprungen wurden, weil ihr Name `FALTA REVISION` enthält. Die zweite listet die PDFs auf, aus denen keine Menge gelesen werden konnte, jeweils mit dem Grund: keine lesbaren Mengen, gescannt ohne OCR oder ein Fehler beim Lesen der Datei.

`sin_resultado.log` enthält für jedes PDF, das keine Menge produziert hat, den Grund und den vollständigen Text, der daraus gelesen wurde. Es ist der erste Ort, um zu schauen, wenn ein PDF nicht richtig gelesen wird.

## Konsolennachrichten

Während es läuft, druckt lotes_auto.py eine Zeile pro PDF. Die Nachricht nach dem Pfeil hat folgende Bedeutung:

- `-> N cantidades (texto)`: N Größen wurden aus einem PDF gelesen, das Text enthält.
- `-> N cantidades (OCR)`: N Größen wurden aus einem gescannten PDF gelesen.
- `-> 0 cantidades`: Es wurde nichts gelesen. Siehe sin_resultado.log.
- `-> omitido (FALTA REVISION)`: Die Datei wurde wegen ihres Namens übersprungen.
- `-> escaneado, OCR no disponible`: Ein gescanntes PDF konnte nicht gelesen werden, da keine OCR-Engine installiert ist.
- `-> ERROR`: Die Datei konnte nicht geöffnet oder gelesen werden.

Am Ende druckt es die Anzahl der Chargen und Mengen, die Namen der Dateien, die es geschrieben hat, und, wenn eine einzelne Menge größer als `50,000` ist, eine Liste dieser Mengen, angeführt von `cantidades sospechosamente grandes`.

## Ausstiegsstatus

- `0`: Erfolg.
- `1`: Die Excel-Datei konnte nicht geschrieben werden, normalerweise weil sie in Excel geöffnet ist.
- `2`: ungültige Argumente, oder der Ordner existiert nicht.

## Konfiguration

Die Einstellungen befinden sich oben in lotes_auto.py. Bearbeiten Sie sie mit einem beliebigen Texteditor und speichern Sie die Datei in der UTF-8-Codierung.

`LOTE_WORDS` und `CANT_WORDS` sind reguläre Ausdrücke mit den Wörtern, die Chargennummern und -mengen kennzeichnen.

`PRIORIDAD_CANTIDAD` entscheidet, welche Menge Spalte oder Zeile verwendet wird, wenn ein Dokument mehrere hat. Das erste Wort in der Liste, das in einer Überschrift erscheint, gewinnt. Die Standardreihenfolge lautet `pieza`, `devuel`, `frasco`, `ampollet`, `ampula`, `vial`, `recibid`, `cantidad`, `factur`; das heißt, eine Spalte über Stücke oder zurückgegebene Waren wird bevorzugt und eine in Rechnung gestellte Menge wird nur als letzter Ausweg verwendet.

`AVISO_CANTIDAD_GRANDE` ist die Größe, oberhalb derer eine einzelne Menge als verdächtig gemeldet wird. Der Standard ist `50000`.

`OMITIR_NOMBRE` ist das Muster der Dateinamen, die übersprungen werden. Der Standard entspricht `FALTA REVISION`.

`LOTE_ID` ist das Muster einer gültigen Chargennummer. Der Standard akzeptiert genau acht Ziffern, beginnend mit `24`, `25` oder `26`. Um beispielsweise auch Chargen ab 27 zu akzeptieren, ändern Sie `2[456]` in `2[4567]`.

## Installation

### Übersicht

Das Programm benötigt Python und die Python-Pakete pdfplumber und openpyxl. Um gescannte PDFs zu lesen, braucht es auch eine OCR-Engine:

- Mit Python 3.12 oder früher: das Paket rapidocr_onnxruntime. Dies ist der bevorzugte Motor.
- Mit Python 3.13 oder höher: die Pakete rapidocr und onnxruntime. Beim Testen dieser Engine wird eine Ziffer falsch gelesen, die die bevorzugte Engine korrekt gelesen hat, so dass Python 3.12 empfohlen wird, wenn viele gescannte PDFs verarbeitet werden.

Die OCR-Engine ist nicht für FreeBSD, OpenBSD, NetBSD oder Alpine Linux verfügbar. Auf diesen Systemen werden nur PDFs gelesen, die Text enthalten.

Die folgenden Anweisungen legen alles in einem Ordner mit dem Namen `pdf-totals` in Ihrem Heimordner mit einer privaten Python-Umgebung in `pdf-totals/venv` (`pdf-totals\python` unter Windows). Nichts ist systemweit installiert, außer Python selbst.

### Windows 10 und Windows 11

Es sind keine Administratorrechte erforderlich. Öffnen Sie PowerShell und führen Sie den folgenden Befehl aus. Es lädt die einbettbare Python 3.12-Distribution von python.org in Ihren Benutzerordner herunter, installiert pip und die erforderlichen Pakete und lädt die beiden Skripte herunter. Windows PowerShell 5.1 unterstützt den `&&`-Operator nicht, daher sind die Schritte durch Semikolone getrennt.

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Das Paket msvc-runtime platziert die Microsoft Visual C++-Laufzeitdateien neben Python, die die OCR-Engine benötigt und die einigen Systemen fehlen. Warnungen von pip, dass Skripte in einem Ordner installiert sind, der sich nicht auf PATH befindet, können ignoriert werden.

Typ `curl.exe`, nicht `curl`. In Windows PowerShell 5.1 ist `curl` ein Alias für `Invoke-WebRequest`, der verschiedene Optionen auswählt. `curl.exe` ist in Windows 10 Version 1803 und höher enthalten.

Um das Programm auszuführen:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

Wenn Downloads auf dem Computer blockiert sind, laden Sie die beiden Dateien `python-3.12.10-embed-amd64.zip` und `get-pip.py` auf einem anderen Computer herunter, kopieren Sie sie in den Downloads-Ordner und ersetzen Sie die beiden `Invoke-WebRequest`-Schritte durch die entsprechenden Pfade.

Diese Anweisungen gelten für 64-Bit-Intel- und AMD-Prozessoren. Verwenden Sie unter Windows auf ARM stattdessen `python-3.12.10-embed-arm64.zip`; dies wurde nicht getestet.

### macOS

Es sind keine Administratorrechte erforderlich, die über die Installation von Python hinausgehen. Installieren Sie Python 3.12 entweder mit dem Installer von python.org (https://www.python.org/downloads/macos/) oder mit Homebrew (`brew install python@3.12`). Dann laufen Sie im Terminal:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Wenn Homebrew verwendet wird und python3.12 nicht gefunden wird, verwenden Sie stattdessen `$(brew --prefix)/bin/python3.12`. Sowohl Intel als auch Apple Silicon Macs werden unterstützt.

Um das Programm auszuführen:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

### Linux

Unter Linux wird die OCR-Engine zusammen mit opencv-python-headless anstelle von opencv-python installiert. Die Headless-Variante benötigt die Systemgrafikbibliotheken nicht, so dass keine zusätzlichen Systempakete erforderlich sind.

Alle Distributionen verwenden den gleichen Befehl zum Erstellen der Umgebung, sobald Python 3.12 als `python3.12` verfügbar ist (unter Debian 12 verwenden Sie `python3`, was Version 3.11 ist):

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Um das Programm auszuführen:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

So erhalten Sie Python 3.12 für gängige Distributionen:

Debian 12 und Ubuntu 24.04 liefern ein geeignetes Python. Installieren Sie das venv-Modul und curl als root:

```
sudo apt install python3-venv curl
```

Unter Ubuntu 24.04 lautet der Befehl `python3.12`. Unter Debian 12 lautet der Befehl `python3` (Version 3.11); Ersetzen Sie `python3.12` durch `python3` im obigen Befehl. Debian 13 und neuere Ubuntu-Veröffentlichungen liefern Python 3.13 oder höher; Verwenden Sie die uv-Methode unten.

Fedora liefert ein zusätzliches Python 3.12-Paket:

```
sudo dnf install python3.12
```

openSUSE Leap 15.6 und Tumbleweed liefern ein zusätzliches Python 3.12 Paket:

```
sudo zypper install python312
```

Arch Linux liefert nur das neueste Python. Verwenden Sie die uv-Methode unten.

Jede glibc-basierte Distribution ohne root: uv ist ein einzelnes Programm, das in Ihrem Home-Ordner installiert wird und eine private Kopie von Python 312 herunterladen kann.

```
curl -LsSf https://astral.sh/uv/install.sh | sh && mkdir -p ~/pdf-totals && cd ~/pdf-totals && ~/.local/bin/uv venv -p 3.12 venv && ~/.local/bin/uv pip install -p venv/bin/python pdfplumber openpyxl rapidocr_onnxruntime && ~/.local/bin/uv pip uninstall -p venv/bin/python opencv-python && ~/.local/bin/uv pip install -p venv/bin/python opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Alpine Linux verwendet die musl C-Bibliothek, für die keine OCR-Engine verfügbar ist. Nur PDFs, die Text enthalten, werden gelesen:

```
doas apk add python3 curl && mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv venv && venv/bin/pip install pdfplumber openpyxl && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Verwenden Sie `sudo` anstelle von `doas`, wenn Ihr Alpine-System so konfiguriert ist.

### FreeBSD, OpenBSD und NetBSD

Es werden nur PDFs gelesen, die Text enthalten. Die OCR-Engine und die von pdfplumber verwendete PDF-Rendering-Bibliothek sind für diese Systeme nicht verfügbar, und pdfplumber liest Text ohne sie. Das von pdfminer.six benötigte cryptography-Paket wird aus den Systempaketen übernommen, da das Erstellen eines Rust-Compilers erforderlich ist. Die restlichen Pakete werden in reinem Python geschrieben und ohne ihre optionalen Abhängigkeiten installiert.

FreeBSD, mit den Befehlen für root laufen durch `su` oder `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Die Paketnamen tragen die Python-Version, 311 in diesem Beispiel. Wenn Ihr System eine andere Version verwendet, finden Sie die richtigen Namen mit `pkg search cryptography` und ändern Sie `python3.11` entsprechend.

OpenBSD, wobei der erste Befehl als root durch `doas` ausgeführt wird:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

NetBSD, mit dem ersten Befehl ausgeführt als root:

```
pkgin install py312-pip py312-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Wie bei FreeBSD tragen die Paketnamen die Python-Version, 312 in diesem Beispiel; Verwenden Sie `pkgin search cryptography`, um die Namen auf Ihrem System zu finden.

Um das Programm auf einem beliebigen BSD auszuführen:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

## Aktualisierung

Laden Sie die beiden Skripte erneut herunter. Unter Windows:

```
curl.exe -L -o "$env:USERPROFILE\pdf-totals\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$env:USERPROFILE\pdf-totals\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Unter macOS und Linux:

```
cd ~/pdf-totals && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Verwenden Sie auf den BSDs `fetch` oder `ftp` auf die gleiche Weise. GitHub kann die vorherige Version nach einer Änderung einige Minuten lang bedienen. Um sofort eine bestimmte Version zu erhalten, ersetzen Sie `main` in der Adresse durch die Commit-Kennung.

## Fehlerbehebung

Die Excel-Datei kann nicht gespeichert werden. Das Programm endet mit dem Exit-Status 1, wenn totales_lotes.xlsx in Excel geöffnet ist. Schließen Sie es und führen Sie das Programm erneut aus.

Alle oder viele PDFs zeigen `0 cantidades`. Führen Sie diagnostico.py auf einem von ihnen aus. Wenn der angezeigte Text leer ist, ist das PDF ein Scan und die OCR liest ihn nicht. Wenn der Text vorhanden ist, aber nichts erkannt wird, werden die Beschriftungen oder das Layout nicht erkannt; der angezeigte Text oder der entsprechende Eintrag in sin_resultado.log zeigt an, welche Wörter das Dokument verwendet, und sie können der Konfiguration hinzugefügt werden.

Eine Summe ist viel zu groß. Suchen Sie nach dem Batch in detalle_lotes.txt, um jeden hinzugefügten Betrag und seine Quelle zu sehen. Häufige Ursachen sind eine Zahlenspalte, die falsch gewählt wurde, weil das Dokument eine unbekannte Überschrift, zwei benachbarte Zahlen, die von der OCR zusammengeführt werden, und das gleiche PDF verwendet, das zweimal im Ordner vorhanden ist.

Die OCR kann unter Windows (`DLL load failed`) nicht geladen werden. Installieren Sie msvc-runtime wie in den Windows-Anweisungen gezeigt.

Die OCR kann nicht unter Linux (`libGL.so.1: cannot open shared object file`) geladen werden. Ersetzen Sie opencv-python durch opencv-python-headless, wie in den Linux-Anweisungen gezeigt.

pip kann rapidocr_onnxruntime nicht installieren. Dieses Paket unterstützt Python bis 3.12. Verwenden Sie Python 3.12, oder installieren Sie stattdessen rapidocr und onnxruntime.

## Beschränkungen

Handschriftliche Zahlen werden von der OCR gelesen, wenn sie klar geschrieben sind. Unordentliche Handschrift, Häkchen und handschriftliche Abkürzungen wie `N/A` werden nicht zuverlässig gelesen. PDFs, deren Mengen nicht gelesen werden können, erscheinen in revisar_a_mano.txt und müssen von Hand überprüft werden.

Die OCR kann gelegentlich eine Ziffer falsch lesen, insbesondere bei verschwommenen oder verzerrten Scans. Summen aus gescannten PDFs sollten stichprobenartig mit detalle_lotes.txt verglichen werden.

Das Batch-Nummernformat wird durch `LOTE_ID` festgelegt. Dokumente, die ein anderes Format verwenden, werden ignoriert, bis das Muster geändert wird.

Es wird erwartet, dass es sich bei den Mengen um ganze Stückzahlen handelt. Dokumente, die Boxen oder andere Einheiten zählen, werden nicht konvertiert.

## Akten

- lotes_auto.py: das Programm.
- diagnostico.py: das Diagnosewerkzeug. Es muss sich im gleichen Ordner wie lotes_auto.py befinden.
- README.md: eine kurze Beschreibung, auch als README.de.md und README.es.md verfügbar.
- `docs/`: dieses Handbuch in Englisch, Deutsch und Spanisch.
- `tools/translate_docs.py`: das Skript, das die Dokumentation übersetzt.
- `LICENSE`: Der Lizenztext.

## Lizenz

MIT-Lizenz. Siehe die Datei LICENSE.
