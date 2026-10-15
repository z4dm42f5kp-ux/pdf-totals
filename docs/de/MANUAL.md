# pdf-totals-Handbuch

Diese Datei wurde automatisch aus dem englischen Original mit Argos Translate übersetzt. Wenn es sich vom englischen Original unterscheidet, gilt das englische Original.

## Name

lotes_auto.py - Berechnen Sie für jede Chargennummer die Stücke, die zurückgezogen wurden, die Stücke, die nicht verteilt wurden und die Gesamtzahl der Stücke zur Zerstörung

diagnostico.py - zeigen, was lotes_auto.py in einem oder mehreren PDF-Dokumenten liest und erkennt

## Synopsis

```
python lotes_auto.py [-h] [--version] [--excel PATH] [--retirado-esperado N] [--no-distribuida-esperada N] [--total-esperado N] [--procesos N] [--sin-cache] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] [--lectura N] PATH
```

## Beschreibung

lotes_auto.py liest alle PDF-Dateien im Ordner FOLDER und in allen Unterordnern. In den Dokumenten geht es um die Rücknahme eines Produkts vom Markt: Rückgabeformulare, Lieferquittungen, Warenquittungen, Inventarlisten und Rücknahmeberichte. Für jede Chargennummer berechnet das Programm drei Mengen:

- Die zurückgenommene Menge (`Cantidad de producto retirado`). Dies sind die Stücke, die vom Markt zurückgegeben oder zurückgewonnen wurden. Sie stammen aus Rückgabeformularen, Lieferquittungen und Warenquittungen.
- Die nicht verteilte Menge (`Cantidad no distribuida`). Dies sind die Stücke, die noch im Bestand sind und nie verteilt wurden. Sie stammen aus der Excel-Referenzdatei, die unten beschrieben wird. Ohne die Excel-Referenzdatei stammen sie aus Inventardokumenten und aus dem Auszahlungsbericht.
- Die Gesamtmenge für die Zerstörung (`Cantidad total para destrucción`). Dies ist die Summe der beiden anderen Mengen.

Das Programm überprüft dann das Ergebnis. Die Grenzwerte gelten für jede Charge selbst, nicht für die Summe einer Tabelle: Keine einzelne Charge darf mehr als `48,042`-Stücke oder mehr als `82,090`-Stücke haben, die nicht verteilt wurden. Die entnommenen Mengen aller Chargen müssen sich zu genau `209,908`, die nicht verteilten Mengen zu genau `134,543` und die Summe für die Zerstörung zu genau `344,451` addieren. Diese Werte können in der Konfiguration oder auf der Kommandozeile geändert werden.

## Die Excel-Referenzdatei

Das Programm sucht die Excel-Datei mit dem Namen `Retirado_retiro_producto_Traxicoll` (.xlsx oder .xlsm) im Ordner FOLDER und in allen Unterordnern. Wenn es keine Datei mit diesem Namen gibt, verwendet sie die erste Excel-Datei, die eine Spalte hat, deren Titel `no distribuida` enthält. Eine Datei im alten .xls-Format kann nicht gelesen werden; öffnen Sie sie in Excel und speichern Sie sie als .xlsx-Arbeitsmappe. Eine Datei kann auch mit der Option `--excel` gegeben werden. In dieser Datei findet das Programm drei Spalten: die Spalte mit den Chargennummern, die Spalte mit der nicht verteilten Menge und die Spalte, deren Titel `retirado` enthält, die nach Kundenangaben die zurückgezogene Menge ist. Andere Säulen, wie die Summe für die Zerstörung, werden ignoriert.

Die nicht verteilte Menge jeder Charge wird aus dieser Datei entnommen, nicht aus PDF-Dokumenten. Die zurückgezogene Menge wird weiterhin aus den PDF-Dokumenten berechnet und das Ergebnis für jede Charge mit der zurückgezogenen Menge in der Excel-Datei verglichen. Dieser Vergleich zeigt genau, welcher Charge nicht passt und welche PDF-Dokumente dazu beitragen.

Eine Chargennummer, die in einem PDF-Dokument, aber nicht in der Excel-Datei erscheint, wird nicht gezählt, da es sich fast immer um eine Chargennummer handelt, die falsch gelesen wurde. Diese Mengen sind in der Überprüfungsakte aufgeführt. Wenn die Excel-Datei Formeln enthält, muss sie von Excel gespeichert worden sein, damit die berechneten Werte in der Datei gespeichert werden.

Wenn ein PDF Text enthält, liest das Programm den Text direkt. Wenn eine Seite keinen Text enthält, behandelt das Programm die Seite als gescanntes Bild und liest es mit optischer Zeichenerkennung (OCR). Wenn keine OCR-Engine installiert ist, liest das Programm keine gescannten PDFs und listet sie zur manuellen Überprüfung auf.

Das Programm wurde für Dokumente über zurückgegebene pharmazeutische Produkte in Mexiko geschrieben. Alle Nachrichten und Ausgabedateien sind in Spanisch. Dieses Handbuch erklärt ihre Bedeutung.

diagnostico.py verwendet genau den gleichen Code wie lotes_auto.py, um Dokumente zu lesen und zu analysieren. Es zeigt den Text, den es aus jedem PDF gelesen hat, die Entscheidungen, die es getroffen hat und das Ergebnis. Verwenden Sie es, um ein PDF zu untersuchen, das nicht richtig gelesen wird. Es schreibt keine Dateien.

## Wie das Programm die Mengen findet

Das Programm verlässt sich nicht nur auf die Wörter in den Dokumenten. Es untersucht die Struktur jeder Tabelle und verwendet Arithmetik, um zu entscheiden, welche Zahl welche ist.

Zuerst findet das Programm die Zeilen, die eine Chargennummer enthalten. Aus den Positionen der Werte in diesen Zeilen werden die Spalten der Tabelle gebildet. Jeder Wert wird als Chargennummer, Datum, Preis, ganze Zahl, Code oder Text klassifiziert.

Dann entscheidet das Programm, welche Spalte die Anzahl der Stücke enthält. Es werden nur Spalten mit ganzen Zahlen betrachtet. Zeilennummern, Codes, die sich in jeder Zeile und in jedem Jahr wiederholen, sind ausgeschlossen. Folgende Beweise werden verwendet, vom stärksten bis zum schwächsten:

1. Die mit dem Stückpreis multiplizierten Stücke entsprechen dem Gesamtpreis in derselben Zeile. Zum Beispiel ist `4,408` multipliziert mit `$27.00` `$119,016.00`.
2. Die Summe der Spalte entspricht einer Zahl, die an anderer Stelle im Dokument gedruckt wird, beispielsweise `TOTAL DE PIEZAS 6,889` oder `No CAJAS / PIEZAS 1978`.
3. Titel der Spalte. Titel wie `PIEZAS`, `CANTIDAD EN PIEZAS`, `No. de piezas`, `RECIBIDA` oder `DEVUELTAS` zählen dafür. Titel wie `FACTURADA` oder `DISTRIBUIDA` zählen dagegen, weil sie die ursprüngliche Menge und nicht die zurückgegebene Menge beschreiben. Titel mit `CAJAS` zählen dagegen, weil sie Boxen und nicht Stücke zählen.

Ein Rückzugsbericht wird allein durch Arithmetik erkannt. In einem Auszahlungsbericht enthält jede Zeile fünf Zahlen für eine Charge und zwei Relationen in jeder Zeile:

- hergestellt minus verteilt gleich nicht verteilt;
- nicht verteilt plus zurückgezogen entspricht der Gesamtsumme für die Zerstörung.

Das Programm findet die Spalten, die diese Beziehungen erfüllen, so dass die Spaltentitel nicht benötigt werden. Die zurückgezogene Spalte ist diejenige, die niemals größer ist als die verteilte Spalte. Wenn eine Zeile des Berichts die Relationen nicht erfüllt, z. B. weil die OCR eine Ziffer falsch gelesen hat, verwendet das Programm die Relationen, um den richtigen Wert wiederherzustellen.

Die Art jeder Menge wird wie folgt festgelegt. Eine Größe wird nicht verteilt, wenn ihr Spaltentitel oder Zeilentitel dies sagt, zum Beispiel `EXISTENCIA`, `INVENTARIO` oder `NO DISTRIBUIDO`. Ansonsten entscheiden die Worte der ganzen Seite: Seiten über Rückgaben, Auszahlungen oder Quittungen geben zurückgezogene Mengen an, und Seiten über Inventar geben keine verteilten Mengen an. Wenn eine Seite keinen eindeutigen Typ hat, wird der Typ der anderen Seiten derselben PDF verwendet. Wenn das PDF keinen eindeutigen Typ hat, zählt die Menge als zurückgezogen und eine Notiz wird hinzugefügt.

Das Programm liest auch zwei andere Layouts:

- Formulare, in denen jede Spalte ein Erzeugnis beschreibt. Eine Zeile enthält die Chargennummern und andere Zeilen enthalten Mengen wie `Cantidad Recibida` oder `Cantidad Facturada`. Die Zeile wird durch die Priorität ausgewählt, die im Abschnitt Konfiguration beschrieben wird.
- Textzeilen mit Beschriftungen wie einer Zeile `LOTE: 25240310` und einer Zeile `CANTIDAD EN PIEZAS: 1,250`.

## Kontrollen und Berichtigungen

Jede Menge wird als überprüft oder nicht überprüft gekennzeichnet. Eine Menge wird überprüft, wenn die mit dem Stückpreis multiplizierte Stückzahl dem gedruckten Gesamtpreis entspricht, wenn die Summe ihrer Tabelle einer gedruckten Gesamtmenge entspricht oder wenn die Summe ihrer Charge der Rücknahmemeldung entspricht.

Wenn eine Tabelle ihrer eigenen gedruckten Stückzahl widerspricht und genau eine Zeile nicht verifiziert wird, berechnet das Programm aus der gedruckten Summe den richtigen Wert dieser Zeile. Er verwendet den korrigierten Wert nur, wenn der gedruckte Gesamtpreis ihn bestätigt. Ansonsten wird das PDF als schlecht markiert.

Zwei Seiten desselben PDFs, die genau die gleichen Chargennummern enthalten, werden als dasselbe Dokument behandelt, beispielsweise ein Rücksendeformular und dessen detaillierte Liste, oder die gleiche Quittung zweimal gedruckt. Nur eine der beiden Seiten wird gezählt. Das Programm bevorzugt die Seite, deren Summe mit der eigenen gedruckten Gesamtsumme übereinstimmt. Wenn die beiden Seiten für eine Charge nicht übereinstimmen, verwendet das Programm die Kopie, die mit der Referenz-Excel-Datei übereinstimmt.

Ein Charge, der auf mehreren anderen Seiten derselben PDF erscheint, wird einmal gezählt, wenn die Menge gleich ist. Eine Seite, deren Menge die Summe der anderen Seiten ist, wird nicht gezählt. Wenn der gleiche Charge auf mehreren Seiten mit unterschiedlichen Mengen erscheint, werden die Mengen hinzugefügt und das PDF zur Überprüfung aufgelistet. Es werden immer mehrere Zeilen des gleichen Chargen auf derselben Seite hinzugefügt, da es sich um separate Zeilen des Dokuments handelt, zum Beispiel unterschiedliche Rechnungen.

Wenn ein Charge seinen Grenzwert überschreitet, nicht mit dem Auszahlungsbericht übereinstimmt oder die Gesamtsumme nicht mit der erwarteten Gesamtsumme übereinstimmt, liest das Programm die zweifelhaften Seiten der zweifelhaft gescannten PDFs noch zweimal. Eine zweifelhafte Seite ist eine gescannte Seite mit einer Menge, die nicht verifiziert ist, die eine fehlgeschlagene Überprüfung hat, die zu einer Charge gehört, die nicht übereinstimmt oder deren Chargennummer nicht in der Excel-Datei enthalten ist; eine Seite eines Auszahlungsberichts, dessen Zeile nicht übereinstimmt; und eine Seite ohne Mengen, die Chargen oder Teile erwähnt, die mit einem geringen Vertrauen gelesen wurde oder von der fast nichts gelesen wurde. Seiten, deren Mengen alle überprüft werden, werden nicht erneut gelesen, weil eine andere Lesung sie nicht verbessern kann. Wenn ein PDF überhaupt keine Menge produziert, werden alle gescannten Seiten erneut gelesen. Die zweite Lesung verwendet eine höhere Auflösung und mehr Kontrast. Die dritte Lesung vergrößert das Bild auf das Doppelte und schärft es, was bei kleinem oder verschwommenem Text hilft. Das Programm hält die Lesung mit den meisten verifizierten Mengen und hält die anderen Lesungen für den nächsten Schritt. Dann versucht es, jeden Unterschied genau und ohne Zweideutigkeit zu erklären:

- Wenn die andere Lektüre von genau einem PDF die Charge-Match macht, wird diese Lektüre verwendet.
- Wenn eine Zeile in einem PDF eine exakte Kopie einer Zeile in einem anderen PDF ist und die Charge beim Entfernen der Kopie übereinstimmt, wird die Kopie nicht gezählt.
- Wenn eine Größe eine Chargennummer hat, die nicht in der Excel-Referenzdatei enthalten ist, sich von einer Charge in der Excel-Datei um eine Ziffer oder um zwei getauschte Ziffern unterscheidet und diese Charge genau übereinstimmt, wird die Chargennummer korrigiert.
- Wenn ein Stiftstrich oder Häkchen neben einer Zahl als zusätzliche Ziffer 1, 7 oder 4 gelesen wurde und das Entfernen es die Charge genau zusammenpasst, wird die Zahl korrigiert.
- Wenn das andere Lesen von genau einem PDF oder das Entfernen von genau einer kopierten Zeile die zurückgezogene Summe genau der erwarteten Summe entspricht, wird diese Änderung angewendet.

Ein PDF wird nur dann als schlecht markiert, wenn der Beweis schlüssig ist: wenn seine eigenen Zahlen in jeder Lesung seinen eigenen gedruckten Gesamtwerten widersprechen oder wenn es eine genaue Kopie enthält, die durch eine Charge oder durch die erwartete Gesamtsumme bewiesen wird. Alle anderen Zweifel werden als verdächtig aufgeführt und nicht verändert.

## Regeln für jeden Wert

- Eine Chargennummer hat genau acht Stellen und beginnt mit `25` oder `26`. Das Programm ignoriert alle anderen Zahlen, zum Beispiel Preise, Daten, Rechnungsnummern und Produktcodes. Wenn die OCR andere Zeichen an eine Chargennummer anhängt, z. B. einen vertikalen Balken, entfernt das Programm sie. Wenn die OCR eine handschriftliche Chargennummer in zwei Teile teilt, werden die Teile wieder zusammengefügt.
- Das Programm liest Zahlen in mexikanischer Notation. Ein Komma trennt die Tausende und ein Punkt trennt die Dezimalzahlen. Zum Beispiel ist `1,250` eintausendzweihundertfünfzig, und `1,250.50` ist eintausendzweihundertfünfzigeinhalb. Eine Zahl wie `4.408`, bei der die OCR das Komma als Punkt liest, wird als `4,408` gelesen, da Stücke immer ganze Zahlen sind.
- Handschriftliche Häkchen und Striche neben einer Zahl, wie `695 ✓` oder `/ 630`, werden ignoriert.
- Eine Spalte, deren Titel einen Grund, einen Code, ein Folio oder eine Rechnung nennt, wird nicht als Menge verwendet, und auch keine Spalte mit dem gleichen Wert in jeder Zeile.
- Das Programm überspringt jede Datei, deren Name `FALTA REVISION` enthält. Das Programm ignoriert den Unterschied zwischen Großbuchstaben und Kleinbuchstaben, und das O kann einen Akzent haben. Übersprungene Dateien zählen nicht zu einer Gesamtsumme.

## Optionen

Optionen von lotes_auto.py:

- `FOLDER` ist der Ordner, der die PDF-Dateien enthält. Unterordner sind enthalten.
- `--excel PATH` verwendet diese Excel-Referenzdatei, anstatt den Ordner zu durchsuchen.
- `--retirado-esperado N` legt die erwartete Summe der zurückgezogenen Stücke fest. Der Standardwert ist `209,908`. Der Wert 0 schaltet den Check ab.
- `--no-distribuida-esperada N` setzt die erwartete Summe der Stücke, die nicht verteilt wurden. Der Standardwert ist `134,543`. Der Wert 0 schaltet den Check ab.
- `--total-esperado N` legt die erwartete Gesamtmenge für die Zerstörung fest. Der Standardwert ist `344,451`. Der Wert 0 schaltet den Check ab.
- `--procesos N` liest N Seiten gleichzeitig, in N separaten Prozessen. Der Standard ist automatisch: ein Prozess für jeden Prozessorkern, höchstens 8, und nicht mehr als der freie Speicher erlaubt (ca. 1,2 GB für jeden Prozess). Der Wert 1 liest eine Seite nach der anderen in einem einzigen Prozess.
- `--sin-cache` liest jede Seite erneut, auch wenn sie sich im Cache befindet (siehe unten).
- `--version` zeigt die Versionsnummer an. Das Programm stoppt dann.
- `-h` oder `--help` zeigt einen kurzen Hilfetext. Das Programm stoppt dann.

Optionen von diagnostico.py:

- `PATH` ist eine einzelne PDF-Datei oder ein Ordner. Wenn es sich um einen Ordner handelt, untersucht das Programm alle darin enthaltenen PDF-Dateien.
- `--lineas N` zeigt nicht mehr als N Textzeilen für jedes PDF. Der Standardwert ist 40.
- `--lectura N` verwendet OCR-Lesung N: 1 ist der normale Messwert, 2 verwendet eine höhere Auflösung und mehr Kontrast und 3 vergrößert und schärft das Bild.
- `--version` zeigt die Versionsnummer an. Das Programm stoppt dann.
- `-h` oder `--help` zeigt einen kurzen Hilfetext. Das Programm stoppt dann.

## Ausgabedateien

lotes_auto.py erstellt den Ordner `resultados_lotes` innerhalb des Ordners FOLDER und schreibt die folgenden Dateien hinein. Es ersetzt bestehende Dateien mit den gleichen Namen. Der Ordner `resultados_lotes` wird ignoriert, wenn das Programm nach PDF-Dokumenten und Excel-Dateien sucht.

`totales_lotes.xlsx` enthält das Ergebnis auf einem Blatt mit dem Namen `Lotes`. Das Blatt enthält drei Tabellen nebeneinander: `Cantidad de producto retirado`, `Cantidad no distribuida` und `Cantidad total para destrucción`. Jede Tabelle hat die Spalten `Lote` und `Cantidad en Piezas` und eine Zeile `TOTAL`. Alle drei Tabellen listen die gleichen Chargen in der gleichen Reihenfolge auf, so dass eine Zeile über die Tabellen gelesen werden kann. Die dritte Tabelle und die Summen sind Formeln, so dass sie aktualisiert werden, wenn Sie eine Zahl von Hand ändern. Unter jeder Tabelle zeigt das Blatt die erwartete Summe und die Differenz. Zellen von Chargen, die einen Grenzwert überschreiten oder nicht mit dem Entnahmebericht übereinstimmen, sind rot.

`detalle_lotes.txt` listet für jede Charge, jede gefundene Menge, das PDF und die Seite auf, die es enthalten, ob es verifiziert ist und warum und ob es gezählt wurde. Es zeigt auch die Werte des Rückzugsberichts. Verwenden Sie diese Datei, um die Quelle einer Summe zu finden, die falsch aussieht.

`revisar_a_mano.txt` listet alles, was eine Person braucht, in nummerierten Abschnitten auf: eine Zusammenfassung mit den drei Summen und ihren erwarteten Werten; eine Liste von Aktionen in der Reihenfolge, in der sie durchgeführt werden sollten; die Chargen, die nicht übereinstimmen, mit jeder Menge, die für sie gefunden wurde, die PDFs ohne lesbare Mengen, die sie erwähnen, die übersprungenen `FALTA REVISION`-Dateien, die sie erwähnen, und eine Notiz, wenn eine solche Datei genau die fehlende Menge enthält; die Chargen der Excel-Datei ohne PDF; die Mengen, deren Chargennummer nicht in der Excel-Datei ist, mit dem ähnlichsten Charge der Excel-Datei; für jedes PDF ohne lesbare Mengen die Anzahl der Seiten, das OCR-Vertrauen, der Grund, die Batchs, die Zahlen, die falsch gelesen werden könnten Chargennummern, die ersten Textzeilen und was zu tun ist; die PDFs, die als schlecht markiert sind; die Korrekturen; die verdächtigen PDFs; die Wiederholungen in einem PDF; die gleichen Charge und Menge in mehreren PDFs; die übersprungenen Dateien mit ihren Mengen für Informationen; und eine Liste aller PDFs mit ihrem

Die Dateien, deren Name `FALTA REVISION` enthält, werden nur gelesen, um zu berichten, was sie enthalten. Sie zählen nie zu einem Total.

Der Ordner `paginas` enthält ein Bild (PNG) jeder Seite, die eine Person benötigt: Seiten ohne lesbare Mengen, Seiten mit Mengen von Stapeln, die nicht übereinstimmen, verdächtige Mengen, korrigierte Mengen, Mengen mit einer Stapelnummer, die sich nicht in der Excel-Datei befindet, und Seiten von PDFs, die als schlecht markiert sind. revisar_a_mano.txt benennt das Bild neben jedem Problem, so dass eine Seite überprüft werden kann, ohne das PDF zu öffnen. Die Bilder können nicht auf FreeBSD, OpenBSD und NetBSD erstellt werden.

Der Ordner `cache` enthält, was von jeder Seite gelesen wurde, eine kleine Datei für jede Seite und Lesen. Der nächste Durchlauf verwendet diese Dateien, anstatt die Seiten erneut zu lesen, so dass ein zweiter Durchlauf Sekunden statt Minuten dauert. Eine Datei im Cache gehört zum Inhalt einer PDF, nicht zu ihrem Namen: Wenn eine PDF umbenannt wird, z. B. um `FALTA REVISION` hinzuzufügen, werden ihre Seiten nicht erneut gelesen; wenn eine PDF durch eine neue Version ersetzt wird, werden ihre Seiten erneut gelesen. Dateien von PDFs, die sich nicht mehr im Ordner befinden, werden am Ende eines Laufs gelöscht. Sie können den Ordner `cache` jederzeit löschen.

`sin_resultado.log` enthält für jedes PDF, das keine Menge produziert hat, den Grund, den vollständigen Text, den das Programm daraus gelesen hat, und die Entscheidungen, die das Programm getroffen hat. Schauen Sie sich diese Datei zuerst an, wenn ein PDF nicht richtig gelesen wird.

## Konsolennachrichten

Während es läuft, zeigt lotes_auto.py eine Zeile für jedes PDF an, wenn alle Seiten gelesen wurden. Da mehrere Seiten gleichzeitig gelesen werden, werden die PDFs nicht immer in alphabetischer Reihenfolge angezeigt. In Klammern zeigt die Zeile die Anzahl der Seiten und die Sekunden, die zum Lesen benötigt wurden, oder `cache`, wenn die Seiten aus dem Cache kamen. Der Text nach dem Pfeil sagt, wie viele zurückgezogene und nicht verteilte Mengen gelesen wurden, ob es sich bei dem PDF um einen Entnahmebericht handelt und ob der Text direkt (`texto`) oder mit OCR (`OCR`) gelesen wurde. Weitere Meldungen:

- `-> omitido (FALTA REVISION)` bedeutet, dass das Programm die Datei aufgrund ihres Namens übersprungen hat.
- `-> escaneado, OCR no disponible` bedeutet, dass das PDF gescannt wird und keine OCR-Engine installiert ist.
- `-> ERROR` bedeutet, dass das Programm die Datei nicht öffnen oder lesen konnte.
- `Nuevas lecturas` bedeutet, dass das Programm die zweifelhaften Seiten der zweifelhaft gescannten PDFs erneut liest. Jede Zeile benennt die Seiten, die erneut gelesen wurden.
- `... sigue leyendo` bedeutet, dass eine Seite lange gedauert hat. Es wird alle 45 Sekunden beim Lesen der Seite mit dem PDF und der Seitenzahl angezeigt, so dass eine langsame Seite gefunden werden kann.
- `*** ... se deja sin leer` bedeutet, dass eine Seite mehr als 600 Sekunden dauerte. Das Programm hört auf, diese Seite zu lesen, fährt mit den anderen Seiten fort und listet das PDF in revisar_a_mano.txt mit dem Grund auf. Die Seite wird im nächsten Durchlauf erneut gelesen.
- `Tiempo` am Ende zeigt die Dauer des Laufs, die Anzahl der Seiten, die gelesen wurden, und die Anzahl der Seiten, die aus dem Cache kamen.

Am Ende zeigt das Programm die drei Summen und vergleicht die Summe für die Zerstörung mit der erwarteten Summe. Das Wort `COINCIDE` bedeutet, dass die Summen gleich sind. Ansonsten wird der Unterschied gezeigt. Dann listet es die PDFs auf, die als schlecht markiert sind.

## Rückgabewerte des Programms

- `0` bedeutet, dass das Programm beendet ist. Überprüfen Sie die Zusammenfassung, um zu sehen, ob die Summen übereinstimmen.
- `1` bedeutet, dass das Programm die Excel-Datei nicht schreiben konnte. Normalerweise ist die Datei in Excel geöffnet.
- `2` bedeutet, dass die Argumente nicht gültig sind oder dass der Ordner nicht existiert.

## Konfiguration

Die Einstellungen befinden sich am Anfang der Datei lotes_auto.py. Sie können sie mit jedem Texteditor ändern. Speichern Sie die Datei in der UTF-8-Codierung.

`TOTAL_ESPERADO`, `TOTAL_RETIRADO_ESPERADO` und `TOTAL_NO_DISTRIBUIDA_ESPERADO` sind die erwarteten Summen für die Zerstörung, für zurückgezogene Stücke und für Stücke, die nicht verteilt wurden. Die Standardwerte sind `344451`, `209908` und `134543`. Der Wert `None` schaltet einen Check aus.

`MAX_RETIRADO_POR_LOTE` und `MAX_NO_DISTRIBUIDA_POR_LOTE` sind die Grenzen für eine einzelne Charge. Sie sind keine Grenzen für die Summe einer Tabelle. Die Standardwerte sind `48042` und `82090`.

`EXCEL_NOMBRE` ist der Name der Excel-Referenzdatei, die zuerst durchsucht wird. Der Standardwert ist `Retirado_retiro_producto_Traxicoll`. `CARPETA_RESULTADOS` ist der Name des Ordners für die Ausgabedateien. Der Standardwert ist `resultados_lotes`.

`EXCEL_REFERENCIA` ist der Pfad der Excel-Referenzdatei. Der Standardwert `None` bedeutet, dass die Datei im Ordner und seinen Unterordnern durchsucht wird. `SOLO_LOTES_DEL_EXCEL` entscheidet, ob Mengen mit einer Chargennummer, die nicht in der Excel-Datei enthalten ist, ausgelassen werden. Der Standardwert ist `True`.

`USAR_REPORTE_PARA_RETIRADO` entscheidet, wie die zurückgezogenen Mengen eines Auszahlungsberichts verwendet werden. Mit dem Standardwert `False` werden sie nur zur Überprüfung der Rücksendedokumente verwendet. Mit `True` werden sie auch für Chargen verwendet, die kein Rückgabedokument im Ordner haben.

`DUPLICADOS_EN_ARCHIVO` entscheidet, wie ein Charge, der im selben PDF wiederholt wird, gezählt wird. Mit dem Standardwert `iguales` wird er nur einmal gezählt, wenn die Mengen gleich sind. Mit `todos` wird es immer einmal gezählt und die größte Menge gehalten.

`OMITIR_NOMBRE` ist das Muster für die Namen von Dateien, die übersprungen werden. Der Standardwert stimmt mit `FALTA REVISION` überein.

`LOTE_ID` ist das Muster für eine gültige Chargennummer. Der Standardwert akzeptiert genau acht Ziffern, die mit `25` oder `26` beginnen. Um beispielsweise auch Chargennummern zu akzeptieren, die mit `24` beginnen, ändern Sie `2[56]` in `2[456]`.

`OCR_DPI` ist die Auflösung des ersten und dritten OCR-Messwerts und `OCR_DPI_REINTENTO` ist die Auflösung des zweiten OCR-Messwerts. Die Seiten werden nicht in Schwarz-Weiß konvertiert, denn in Tests liest die OCR-Engine solche Seiten viel schlechter. `CONFIANZA_MINIMA` ist das OCR-Vertrauen, unterhalb dessen eine nicht verifizierte Menge als verdächtig aufgeführt wird.

`PROCESOS` ist die Anzahl der Prozesse, die Seiten gleichzeitig lesen. Der Standardwert `0` bedeutet automatisch. `LADO_MAX_PX` ist die größte Größe in Pixeln der längeren Seite eines Seitenbildes für die OCR. Der Standardwert ist `6000`. Die OCR-Engine reduziert sowieso jedes Bild auf 2000 Pixel, so dass diese Grenze das Ergebnis für normale Seiten nicht ändert. Es verhindert riesige Bilder, zum Beispiel von Scans, die mit 72 Punkten pro Zoll gespeichert wurden, deren Seiten etwa einen Meter hoch sind. `LADO_MAX_PX_AMPLIADA` ist die gleiche Grenze für die dritte Lesung, bevor das Bild vergrößert wird. `USAR_CACHE` entscheidet, ob der Cache verwendet wird. `CACHE_VERSION` ist Teil des Namens jeder Cache-Datei; ändern Sie es, damit das Programm den Cache ignoriert, zum Beispiel nach einer Änderung der OCR-Einstellungen. `AVISO_LENTO_S` ist die Anzahl der Sekunden, nach denen eine langsame Seite gemeldet wird. `LIMITE_PAGINA_S` ist die Anzahl der Sekunden, nach denen eine Seite aufgegeben wird. Der Standardwert ist `600`.

`PRIORIDAD_RENGLON` entscheidet, welche Mengenreihe in Formularen verwendet wird, in denen jede Spalte ein Produkt beschreibt.

`SEMANTICA` und `TIPO_DOCUMENTO` enthalten die Wörter, die als ein Signal unter anderem verwendet werden, um Spaltentitel zu verstehen und den Typ einer Seite zu erkennen.

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

Die Gesamtsumme entspricht nicht der erwarteten Gesamtsumme. Open revisar_a_mano.txt. Der Abschnitt über Chargen, die nicht übereinstimmen, zeigt an, welche Chargen eine Grenze überschreiten oder vom Entnahmebericht abweichen und welche PDFs dazu beitragen. Suchen Sie dann nach diesen Chargen in detalle_lotes.txt. Häufige Ursachen sind:

- Ein PDF, das zur Auszahlung gehört, fehlt im Ordner oder heißt `FALTA REVISION`.
- Die OCR hat eine Ziffer falsch gelesen. Die beiden Lesungen eines PDF werden verglichen, und ein Unterschied zwischen ihnen wird als `dudosa` vermerkt.
- Das gleiche Dokument befindet sich zweimal im Ordner unter verschiedenen Namen.
- Das Programm verwendete die falsche Spalte für die Menge, da das Dokument ein Layout verwendet, das das Programm nicht kennt. Führen Sie diagnostico.py mit diesem PDF aus, um zu sehen, welche Spalte ausgewählt wurde und warum.

Alle oder viele PDFs zeigen keine Mengen an. Führen Sie diagnostico.py mit einem dieser PDFs aus. Wenn der angezeigte Text leer ist, wird das PDF gescannt und die OCR liest es nicht. Wenn der Text vorhanden ist, aber nichts erkannt wird, kennt das Programm das Layout des Dokuments nicht.

Die OCR kann unter Windows nicht geladen werden, und die Nachricht enthält `DLL load failed`. Installieren Sie das Paket msvc-runtime wie in den Anweisungen für Windows gezeigt.

Die OCR kann nicht unter Linux geladen werden, und die Nachricht enthält `libGL.so.1`. Ersetzen Sie opencv-python durch opencv-python-headless, wie in den Anweisungen für Linux gezeigt.

pip kann rapidocr_onnxruntime nicht installieren. Dieses Paket unterstützt Python-Versionen bis `Python 3.12`. Verwenden Sie `Python 3.12` oder installieren Sie stattdessen die Pakete rapidocr und onnxruntime.

## Beschränkungen

Die OCR liest handschriftliche Zahlen, wenn sie klar geschrieben sind. Es liest nicht zuverlässig Handschrift, die schwer zu lesen ist, oder handschriftliche Abkürzungen wie `N/A`. Mengen, die von Hand in ein Formular geschrieben werden, werden daher oft nicht gelesen. Diese PDFs sind in revisar_a_mano.txt aufgeführt und müssen von einer Person überprüft werden.

Die OCR kann eine Ziffer falsch lesen, insbesondere wenn ein Scan verschwommen ist. Mengen, die durch einen Preis, eine gedruckte Gesamtsumme oder den Rückzugsbericht verifiziert werden, sind zuverlässig. Mengen, die nicht verifiziert sind, sollten mit dem PDF verglichen werden.

Wenn ein Prozess unerwartet stoppt, beispielsweise weil der Speicher voll ist, liest das Programm die verbleibenden Seiten nacheinander in einem einzigen Prozess. Die OCR dauert etwa ein bis fünf Sekunden für jede Seite und Prozess. Der erste Durchlauf über einen großen Ordner dauert daher einige Minuten. Die nächsten Läufe verwenden den Cache und dauern Sekunden. Nur die zweifelhaften Seiten werden ein zweites und ein drittes Mal gelesen.

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
