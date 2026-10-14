# Manual de pdf-totals

Este archivo se tradujo automáticamente del original inglés con Argos Translate. Si difiere del original inglés, se aplica el original inglés.

## Nombre

lotes_auto.py - agregar el número de piezas por número de lote encontrado en una carpeta de documentos PDF

diagnostico.py - mostrar lo que lotes_auto.py lee y detecta en uno o más documentos PDF

## Sinopsis

```
python lotes_auto.py [-h] [--version] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] PATH
```

## Descripción

lotes_auto.py busca FOLDER y todos sus subcarpetas para archivos que terminan en .pdf. Desde cada PDF lee cada número de lote y el número de piezas grabadas para ese lote. Luego agrega las piezas de cada lote sobre todos los documentos y escribe una fila por lote a un libro de trabajo de Excel. Los lotes que aparecen en varios documentos, o varias veces en un documento, se agregan juntos.

El texto de un PDF se lee directamente cuando el PDF contiene texto. Cuando una página no contiene texto, se trata como imagen escaneada y se lee con reconocimiento de caracteres ópticos (OCR). Si no se instala ningún motor OCR, los PDF escaneados no se leen y se enumeran para revisión manual.

El programa fue escrito para documentos sobre productos farmacéuticos devueltos en México. Todos los mensajes y archivos de salida están en español; su significado se explica en este manual.

diagnostico.py utiliza exactamente el mismo código de lectura y detección que lotes_auto.py. Imprimen el texto que fue leído en cada PDF, lo que se detectó en él y el resultado, para que se pueda examinar un PDF que no se lea correctamente. No escribe ningún archivo.

## Diseños compatibles

El programa reconoce cuatro diseños. Cada página se examina en el orden dado aquí, y el primer diseño que produce resultados se utiliza para esa página.

1. Formas en las que cada columna es un producto. Una fila etiquetada con una palabra de lote, como `Lote`, contiene un número de lote por columna. Otras filas, etiquetadas por ejemplo `Cantidad Recibida` o `Cantidad Facturada`, contienen la cantidad de cada producto en la misma columna. Si existen varias filas de cantidad, la fila es elegida por la prioridad descrita bajo Configuración.

2. Tablas con una fila encabezada. Un encabezamiento como `CANTIDAD`, `PIEZAS` o `No. de piezas` marca la columna de cantidad, y los valores debajo de ella se leen. El número de lote se toma de un título como `LOTE` cuando está presente, de otro modo desde cualquier lugar en la misma fila, por ejemplo dentro de la descripción, o desde las dos filas siguientes. Los encabezados que abarcan dos líneas, como `PIEZAS` por encima de `DEVUELTAS`, se leen como una partida.

3. Líneas etiquetadas en las que el lote viene primero, por ejemplo una línea `LOTE: 25240310` seguido de una línea `CANTIDAD EN PIEZAS: 1,250`. Cada lote está emparejado con la primera cantidad que lo sigue.

4. Líneas etiquetadas en las que la cantidad viene primero. Cada cantidad se combina con el primer lote que lo sigue. Si un documento utiliza el diseño 3 o el diseño 4 se decide por su primera etiqueta.

Palabras reconocidas como etiquetas por lotes: `lote`, `serie`, `serial`, `n/s`, `s/n`, precedida opcionalmente o seguida por `No.`, `Num.` o `Numero`. Palabras reconocidas como etiquetas de cantidad: `cantidad`, `cant.`, `piezas`, `pieza`, `pzas`, `pza`, `pzs`, `pz`, `unidades`, `unid.`, `uds`. Las palabras reconocidas como títulos de cantidad en tablas también incluyen `devueltas`, `devuelto`, `frascos`, `ampolletas`, `ampulas` y `viales`. La palabra `unidad` sola no es una cantidad, porque nombra la unidad de medida. `Cajas` nunca se utiliza, porque cuenta cajas, no piezas.

## Reglas aplicadas a cada valor

- Un número de lote debe consistir de exactamente ocho dígitos y comenzar con `24`, `25` o `26`. Cualquier otra cosa es ignorada, incluyendo precios, cantidades, fechas y códigos de producto. Los caracteres adjuntos a un número de lotes por el OCR, como una barra vertical, se eliminan.
- Los números se leen en notación mexicana, con una coma como el separador de miles y un punto como el separador decimal: `1,250` es mil doscientos cincuenta, y `1,250.50` es mil doscientos cincuenta y medio.
- Se ignoran cuantitativas con decimales, y cantidades que son cero o negativos, porque las piezas son números enteros. Esto también impide que los precios sean contados como piezas.
- Los archivos cuyo nombre contiene `FALTA REVISION`, en mayúscula o inferior y con o sin acento en el O, se saltan y no cuentan hacia ningún total.
- Todavía se añade una cantidad única mayor que `50,000`, pero se informa al final de la carrera como sospechosa.

## Opciones

Opciones de lotes_auto.py:

- `FOLDER`: la carpeta que contiene los archivos PDF. Se incluyen subcarpetas.
- `--version`: imprimir el número de versión y la salida.
- `-h`, `--help`: imprimir un mensaje de uso corto y salida.

Opciones de diagnostico.py:

- `PATH`: un solo archivo PDF, o una carpeta cuyos archivos PDF son examinados.
- `--lineas N`: mostrar en la mayoría de las líneas N de texto por PDF. El defecto es 40.
- `--version`: imprimir el número de versión y la salida.
- `-h`, `--help`: imprimir un mensaje de uso corto y salida.

## Archivo de salida

lotes_auto.py escribe los siguientes archivos en FOLDER. Los archivos existentes con los mismos nombres son reemplazados. Un archivo que estaría vacío se elimina en su lugar.

`totales_lotes.xlsx` es el resultado. Tiene una hoja llamada `Lotes` con dos columnas, `Lote` y `Cantidad en Piezas`, una fila por lote, ordenados por número de lote. Las cantidades se almacenan como números y se muestran con miles de separadores.

`detalle_lotes.txt` lista, para cada lote, cada cantidad que se agregó a él y el PDF de que la cantidad vino. Úsalo para encontrar la fuente de un total que se ve mal.

`revisar_a_mano.txt` tiene dos secciones. La primera lista de los archivos que fueron saltados porque su nombre contiene `FALTA REVISION`. La segunda enumera los PDFs de los que no se puede leer ninguna cantidad, cada uno con la razón: no hay cantidades legibles, escaneadas sin OCR disponible, o un error al leer el archivo.

`sin_resultado.log` contiene, para cada PDF que no produjo cantidad, la razón y el texto completo que fue leído de él. Es el primer lugar en mirar cuando un PDF no se lee correctamente.

## Mensajes de consola

Mientras funciona, lotes_auto.py imprime una línea por PDF. El mensaje después de la flecha tiene el siguiente significado:

- `-> N cantidades (texto)`: Se leyeron cantidades N de un PDF que contiene texto.
- `-> N cantidades (OCR)`: Se leyeron cantidades N de un PDF escaneado.
- `-> 0 cantidades`No se leía nada. See sin_resultado.log.
- `-> omitido (FALTA REVISION)`: el archivo fue saltado debido a su nombre.
- `-> escaneado, OCR no disponible`: un PDF escaneado no se puede leer porque ningún motor OCR está instalado.
- `-> ERROR`: el archivo no se puede abrir ni leer.

Al final imprime el número de lotes y cantidades, los nombres de los archivos que escribió, y, si cualquier cantidad es mayor que `50,000`, una lista de esas cantidades encabezadas por `cantidades sospechosamente grandes`.

## Situación de salida

- `0`: éxito.
- `1`: el archivo Excel no puede ser escrito, generalmente porque está abierto en Excel.
- `2`: argumentos inválidos, o la carpeta no existe.

## Configuración

Los ajustes están en la parte superior de lotes_auto.py. Editarlos con cualquier editor de texto y guardar el archivo en la codificación UTF-8.

`LOTE_WORDS` y `CANT_WORDS` son expresiones regulares con las palabras que etiquetan números y cantidades de lote.

`PRIORIDAD_CANTIDAD` decide qué columna de cantidad o fila se utiliza cuando un documento tiene varios. La primera palabra en la lista que aparece en una partida gana. El pedido predeterminado es `pieza`, `devuel`, `frasco`, `ampollet`, `ampula`, `vial`, `recibid`, `cantidad`, `factur`; es decir, se prefiere una columna sobre piezas o bienes devueltos, y una cantidad facturada se utiliza sólo como último recurso.

`AVISO_CANTIDAD_GRANDE` es el tamaño por encima del cual una sola cantidad es reportada como sospechosa. El predeterminado es `50000`.

`OMITIR_NOMBRE` es el patrón de nombres de archivos que se saltan. El predeterminado coincide con `FALTA REVISION`.

`LOTE_ID` es el patrón de un número de lote válido. El predeterminado acepta exactamente ocho dígitos comenzando con `24`, `25` o `26`. Para aceptar, por ejemplo, lotes comenzando con 27 también, cambiar `2[456]` a `2[4567]`.

## Instalación

### Sinopsis

El programa necesita Python y los paquetes Python pdfplumber y openpyxl. Para leer PDF escaneado también necesita un motor OCR:

- Con Python 3.12 o anterior: el paquete rapidocr_onnxruntime. Este es el motor preferido.
- Con Python 3.13 o posterior: los paquetes rapidocr y onnxruntime. Al probar este motor malinterpretar un dígito que el motor preferido leyó correctamente, por lo que Python 3.12 se recomienda cuando se procesan muchos PDF escaneados.

El motor OCR no está disponible para FreeBSD, OpenBSD, NetBSD o Linux alpino. En esos sistemas sólo se leen PDFs que contienen texto.

Las instrucciones a continuación colocan todo en una carpeta llamada `pdf-totals` en su carpeta de inicio, con un entorno Python privado en `pdf-totals/venv` (`pdf-totals\python` en Windows). Nada se instala a nivel de todo el sistema excepto, donde se señala, Python en sí mismo.

### Windows 10 y Windows 11

No se necesitan derechos de administrador. Abra PowerShell y ejecute el siguiente comando. Descargue la distribución embeddable Python 3.12 desde python.org a su carpeta de usuario, instala pip y los paquetes necesarios, y descarga los dos scripts. Windows PowerShell 5.1 no admite el operador `&&`, por lo que los pasos están separados por semicolonias.

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

El paquete msvc-runtime coloca los archivos de tiempo de ejecución Microsoft Visual C++ junto a Python, que el motor OCR necesita y que algunos sistemas carecen. Advertencias de pip que los scripts están instalados en una carpeta que no está en PATH pueden ser ignorados.

Tipo `curl.exe`, no `curl`. En Windows PowerShell 5.1, `curl` es un alias para `Invoke-WebRequest`, que toma diferentes opciones. `curl.exe` está incluido en Windows 10 versión 1803 y más tarde.

Para ejecutar el programa:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

Si las descargas están bloqueadas en la máquina, descargue los dos archivos `python-3.12.10-embed-amd64.zip` y `get-pip.py` en otro ordenador, copielos a la carpeta Descargas y reemplace los dos pasos `Invoke-WebRequest` con los caminos correspondientes.

Estas instrucciones son para procesadores Intel y AMD de 64 bits. En Windows en ARM, use `python-3.12.10-embed-arm64.zip` en su lugar; esto no ha sido probado.

### macOS

No se necesitan derechos de administrador más allá de instalar Python. Instala Python 3.12 ya sea con el instalador de python.org (https://www.python.org/downloads/macos/) o con Homebrew (`brew install python@3.12`). Luego corre en Terminal:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Si se utiliza Homebrew y python3.12 no se encuentra, use `$(brew --prefix)/bin/python3.12` en su lugar. Tanto Intel como Apple silicon Macs son compatibles.

Para ejecutar el programa:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

### Linux

En Linux el motor OCR se instala junto con opencv-python-headless en lugar de opencv-python. La variante sin cabeza no necesita las bibliotecas gráficas del sistema, por lo que no se requieren paquetes adicionales del sistema.

Todas las distribuciones utilizan el mismo comando para crear el entorno una vez que Python 3.12 esté disponible como `python3.12` (en Debian 12 use `python3`, que es la versión 3.11):

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Para ejecutar el programa:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

Cómo obtener Python 3.12 sobre distribuciones comunes:

Debian 12 y Ubuntu 24.04 envían un Python adecuado. Instalar el módulo de venv y rizar como raíz:

```
sudo apt install python3-venv curl
```

En Ubuntu 24.04 el comando es `python3.12`. En Debian 12 el comando es `python3` (versión 3.11); reemplazar `python3.12` con `python3` en el comando anterior. Debian 13 y el nuevo Ubuntu libera el buque Python 3.13 o posterior; utilice el método uv a continuación.

Fedora envía un paquete adicional Python 3.12:

```
sudo dnf install python3.12
```

openSUSE Leap 15.6 y Tumbleweed buque un paquete adicional Python 3.12:

```
sudo zypper install python312
```

Arch Linux envía sólo el nuevo Python. Utilice el método uv a continuación.

Cualquier distribución basada en glibc, sin root: uv es un único programa que se instala en su carpeta de inicio y puede descargar una copia privada de Python 3.12.

```
curl -LsSf https://astral.sh/uv/install.sh | sh && mkdir -p ~/pdf-totals && cd ~/pdf-totals && ~/.local/bin/uv venv -p 3.12 venv && ~/.local/bin/uv pip install -p venv/bin/python pdfplumber openpyxl rapidocr_onnxruntime && ~/.local/bin/uv pip uninstall -p venv/bin/python opencv-python && ~/.local/bin/uv pip install -p venv/bin/python opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Linux alpino utiliza la biblioteca musl C, para la cual no hay motor OCR disponible. Sólo los PDF que contienen texto son leídos:

```
doas apk add python3 curl && mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv venv && venv/bin/pip install pdfplumber openpyxl && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Utilice `sudo` en lugar de `doas` si así es como se configura el sistema Alpine.

### FreeBSD, OpenBSD y NetBSD

Sólo se leen los PDF que contienen texto. El motor OCR y la biblioteca de renderización PDF utilizada por pdfplumber no están disponibles para estos sistemas, y pdfplumber lee texto sin ellos. El paquete cryptography necesario por pdfminer.six se toma de los paquetes del sistema, porque la construcción requiere un compilador de Rust. Los paquetes restantes están escritos en Python puro y se instalan sin sus dependencias opcionales.

FreeBSD, con los comandos para root run a través de `su` o `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Los nombres de los paquetes llevan la versión Python, 311 en este ejemplo. Si su sistema utiliza otra versión, encuentre los nombres correctos con `pkg search cryptography` y cambie `python3.11` en consecuencia.

OpenBSD, con el primer comando ejecuta como root a través de `doas`:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

NetBSD, con el primer comando ejecutar como root:

```
pkgin install py312-pip py312-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Como en FreeBSD, los nombres de los paquetes llevan la versión Python, 312 en este ejemplo; use `pkgin search cryptography` para encontrar los nombres en su sistema.

Para ejecutar el programa en cualquier BSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

## Actualización

Descargar los dos scripts otra vez. En Windows:

```
curl.exe -L -o "$env:USERPROFILE\pdf-totals\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$env:USERPROFILE\pdf-totals\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

En macOS y Linux:

```
cd ~/pdf-totals && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

En las BSD, utilice `fetch` o `ftp` de la misma manera. GitHub puede servir la versión anterior durante unos minutos después de un cambio. Para obtener una versión específica inmediatamente, reemplace `main` en la dirección con el identificador de confirmación.

## Solución de problemas

El archivo Excel no se puede guardar. El programa se detiene con el estado de salida 1 cuando totales_lotes.xlsx está abierto en Excel. Cierra y vuelve a ejecutar el programa.

Todos o muchos PDFs muestran `0 cantidades`. Ejecute diagnostico.py en uno de ellos. Si el texto que muestra está vacío, el PDF es un escaneo y el OCR no lo está leyendo. Si el texto está presente pero no se detecta nada, no se reconocen las etiquetas ni el diseño; el texto mostrado, o la entrada correspondiente en sin_resultado.log, muestra qué palabras utiliza el documento, y se pueden añadir a la configuración.

Un total es demasiado grande. Busque el lote en detalle_lotes.txt para ver cada cantidad que se agregó y su fuente. Las causas comunes son una columna de cantidad elegida incorrectamente porque el documento utiliza un título desconocido, dos números vecinos fusionados por el OCR, y el mismo PDF presente dos veces en la carpeta.

El OCR no se puede cargar en Windows (`DLL load failed`). Instala msvc-runtime como se muestra en las instrucciones de Windows.

El OCR no se puede cargar en Linux (`libGL.so.1: cannot open shared object file`). Reemplaza opencv-python con opencv-python-headless como se muestra en las instrucciones de Linux.

pip no puede instalar rapidocr_onnxruntime. Este paquete admite Python hasta 3.12. Utilice Python 3.12, o instalar rapidocr y onnxruntime en su lugar.

## Limitaciones

Los números manuscritos son leídos por el OCR cuando están escritos claramente. La escritura, las marcas de verificación y abreviaturas manuscritas como `N/A` no se leen de forma fiable. PDF cuyas cantidades no se pueden leer aparecen en revisar_a_mano.txt y deben ser revisadas a mano.

El OCR puede ocasionalmente malinterpretar un dígito, en particular en los escaneos borrosos o segados. Los totales obtenidos de los PDF escaneados deben ser revisados con detalle_lotes.txt.

El formato número de lote es fijado por `LOTE_ID`. Los documentos que usan otro formato son ignorados hasta que el patrón sea cambiado.

Se espera que las cantidades sean un número entero de piezas. Los documentos que cuentan cajas u otras unidades no se convierten.

## Archivos

- lotes_auto.py: el programa.
- diagnostico.py: la herramienta de diagnóstico. Debe estar en la misma carpeta que lotes_auto.py.
- README.md: una descripción corta, también disponible como README.de.md y README.es.md.
- `docs/`: este manual en inglés, alemán y español.
- `tools/translate_docs.py`: el script que traduce la documentación.
- `LICENSE`: el texto de la licencia.

## Licencia

Licencia MIT. Vea el archivo LICENSE.
