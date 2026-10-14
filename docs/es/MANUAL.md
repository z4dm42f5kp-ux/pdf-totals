# Manual de pdf-totals

Este archivo se tradujo automáticamente del original inglés con Argos Translate. Si difiere del original inglés, se aplica el original inglés.

## Nombre

lotes_auto.py - calcular el número total de piezas para cada número de lote en una carpeta de documentos PDF

diagnostico.py - mostrar lo que lotes_auto.py lee y detecta en uno o más documentos PDF

## Sinopsis

```
python lotes_auto.py [-h] [--version] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] PATH
```

## Descripción

lotes_auto.py busca la carpeta FOLDER y todos sus subcarpetas para archivos cuyo nombre termina en .pdf. En cada PDF encuentra cada número de lotes y el número de piezas que el documento registra para ese lote. Luego calcula el número total de piezas de cada lote sobre todos los documentos y escribe una fila por lote a un libro de trabajo de Excel. Cuando un lote aparece en varios documentos, o varias veces en un solo documento, sus cantidades se agregan juntas.

Cuando un PDF contiene texto, el programa lee el texto directamente. Cuando una página no contiene texto, el programa trata la página como imagen escaneada y la lee con reconocimiento de caracteres ópticos (OCR). Si no se instala ningún motor OCR, el programa no lee PDFs escaneados y los enumera para revisión manual.

El programa fue escrito para documentos sobre productos farmacéuticos devueltos en México. Todos los mensajes y archivos de salida están en español. Este manual explica su significado.

diagnostico.py utiliza exactamente el mismo código que lotes_auto.py para leer y analizar documentos. Muestra el texto que leía de cada PDF, lo que detectó en ese texto y el resultado. Úsalo para examinar un PDF que no se lee correctamente. No escribe ningún archivo.

## Diseños compatibles

El programa reconoce cuatro diseños. Examina cada página en el orden dado aquí y utiliza el primer diseño que produce resultados para esa página.

1. Formas en las que cada columna describe un producto. Una fila tiene una etiqueta de lote, como `Lote`, y contiene un número de lote por columna. Otras filas tienen una etiqueta de cantidad, como `Cantidad Recibida` o `Cantidad Facturada`, y contienen la cantidad de cada producto en la misma columna. Si hay varias filas de cantidad, el programa elige una por la prioridad que se describe en la sección Configuración.

2. Tablas con títulos de columna. Un título de columna como `CANTIDAD`, `PIEZAS` o `No. de piezas` marca la columna de cantidad, y el programa lee los valores debajo de ella. El programa toma el número de lotes de una columna con un título como `LOTE`. Si no hay tal columna, toma el número de lote de cualquier otro lugar en la misma fila, por ejemplo de la descripción, o de las dos filas siguientes. Un título de columna que está escrito en dos líneas, como `PIEZAS` sobre `DEVUELTAS`, se lee como un título.

3. Líneas con etiquetas, en las que el lote viene primero. Un ejemplo es una línea `LOTE: 25240310` seguida de una línea `CANTIDAD EN PIEZAS: 1,250`. El programa empareja cada lote con la primera cantidad que lo sigue.

4. Líneas con etiquetas, en las que la cantidad viene primero. El programa empareja cada cantidad con el primer lote que lo sigue. La primera etiqueta del documento decide si el documento utiliza el diseño 3 o el diseño 4.

Estas palabras son reconocidas como etiquetas para lotes: `lote`, `serie`, `serial`, `n/s` y `s/n`. Pueden combinarse con `No.`, `Num.` o `Numero`.

Estas palabras se reconocen como etiquetas de cantidad: `cantidad`, `cant.`, `piezas`, `pieza`, `pzas`, `pza`, `pzs`, `pz`, `unidades`, `unid.` y `uds`.

Estas palabras también se reconocen como títulos de columna de cantidad: `devueltas`, `devuelto`, `frascos`, `ampolletas`, `ampulas` y `viales`.

La palabra `unidad` por sí sola no marca una cantidad, porque nombra la unidad de medida. La palabra `cajas` nunca se utiliza, porque cuenta cajas y no piezas.

## Reglas para cada valor

- Un número de lote tiene exactamente ocho dígitos y comienza con `24`, `25` o `26`. El programa ignora todos los demás números, por ejemplo precios, cantidades, fechas y códigos de producto. Si el OCR adjunta otros caracteres a un número de lote, como una barra vertical, el programa los elimina.
- El programa lee números en notación mexicana. Un coma separa los miles y un punto separa los decimales. Por ejemplo, `1,250` es mil doscientos cincuenta, y `1,250.50` es mil doscientos cincuenta y medio.
- El programa ignora las cantidades que tienen decimales, y las cantidades que son cero o negativo, porque un número de piezas es siempre un número entero. Esta regla también impide que los precios sean contados como piezas.
- El programa salta cada archivo cuyo nombre contiene `FALTA REVISION`. El capital y las letras pequeñas se tratan por igual, y el O puede tener un acento. Los archivos eliminados no cuentan hacia ningún total.
- Si una sola cantidad es mayor que `50,000`, el programa todavía lo añade al total. Cuando el programa termina, reporta cantidades como sospechosas.

## Opciones

Opciones de lotes_auto.py:

- `FOLDER` es la carpeta que contiene los archivos PDF. Se incluyen subcarpetas.
- `--version` muestra el número de versión. El programa se detiene.
- `-h` o `--help` muestra un texto de ayuda corta. El programa se detiene.

Opciones de diagnostico.py:

- `PATH` es un solo archivo PDF, o una carpeta. Si es una carpeta, el programa examina todos los archivos PDF en ella.
- `--lineas N` no muestra más que líneas N de texto para cada PDF. El valor predeterminado es 40.
- `--version` muestra el número de versión. El programa se detiene.
- `-h` o `--help` muestra un texto de ayuda corta. El programa se detiene.

## Archivo de salida

lotes_auto.py escribe los siguientes archivos en la carpeta FOLDER. Reemplaza los archivos existentes con los mismos nombres. Si un archivo estaría vacío, el programa lo borra en su lugar.

`totales_lotes.xlsx` contiene el resultado. Tiene una hoja con el nombre `Lotes`. La hoja tiene dos columnas, `Lote` y `Cantidad en Piezas`, y una fila para cada lote. Las filas están clasificadas por número de lote. Las cantidades se almacenan como números y se muestran con agrupación de dígitos, por ejemplo `1,406`.

`detalle_lotes.txt` enumera cada cantidad que se agregó a cada lote, junto con el nombre del PDF que contiene la cantidad. Utilice este archivo para encontrar la fuente de un total que se ve mal.

`revisar_a_mano.txt` tiene dos secciones. La primera sección enumera los archivos que fueron saltados porque su nombre contiene `FALTA REVISION`. En la segunda sección se enumeran los PDF de los que no se puede leer ninguna cantidad, cada uno con la razón. Las posibles razones son: no hay cantidades legibles; un PDF escaneado y ningún motor OCR; un error al leer el archivo.

`sin_resultado.log` contiene, para cada PDF que no produjo cantidad, la razón y el texto completo que el programa leyó de él. Mira este archivo primero cuando un PDF no se lee correctamente.

## Mensajes de consola

Mientras funciona, lotes_auto.py muestra una línea para cada PDF. El texto después de la flecha tiene el siguiente significado:

- `-> N cantidades (texto)` significa que el programa leyó cantidades N de un PDF que contiene texto.
- `-> N cantidades (OCR)` significa que el programa leyó cantidades N de un PDF escaneado.
- `-> 0 cantidades` significa que el programa no leyó cantidad. El archivo sin_resultado.log muestra el texto que lee.
- `-> omitido (FALTA REVISION)` significa que el programa saltó el archivo debido a su nombre.
- `-> escaneado, OCR no disponible` significa que el PDF es escaneado y que ningún motor OCR está instalado.
- `-> ERROR` significa que el programa no pudo abrir o leer el archivo.

Al final, el programa muestra el número de lotes y cantidades y los nombres de los archivos que escribió. Si una sola cantidad es mayor que `50,000`, también muestra una lista de estas cantidades bajo el título `cantidades sospechosamente grandes`.

## Códigos de devolución

- `0` significa que el programa terminó con éxito.
- `1` significa que el programa no pudo escribir el archivo Excel. Normalmente el archivo está abierto en Excel.
- `2` significa que los argumentos no son válidos, o que la carpeta no existe.

## Configuración

Los ajustes están al principio del archivo lotes_auto.py. Puedes cambiarlos con cualquier editor de texto. Guarda el archivo en la codificación UTF-8.

`LOTE_WORDS` y `CANT_WORDS` son expresiones regulares. Contienen las palabras que marcan números y cantidades de lote.

`PRIORIDAD_CANTIDAD` decide qué columna de cantidad o fila se utiliza cuando un documento tiene varios. El programa utiliza la primera palabra en esta lista que aparece en un título de columna. El pedido por defecto es `pieza`, `devuel`, `frasco`, `ampollet`, `ampula`, `vial`, `recibid`, `cantidad`, `factur`. Esto significa que se prefiere una columna sobre piezas o sobre bienes devueltos, y que sólo se utiliza una cantidad facturada si no existe otra cantidad.

`AVISO_CANTIDAD_GRANDE` es el límite por encima del cual una sola cantidad es reportada como sospechosa. El valor predeterminado es `50000`.

`OMITIR_NOMBRE` es el patrón para los nombres de los archivos que se saltan. El valor predeterminado coincide con `FALTA REVISION`.

`LOTE_ID` es el patrón para un número de lote válido. El valor predeterminado acepta exactamente ocho dígitos que comienzan con `24`, `25` o `26`. Por ejemplo, para aceptar también los números de lote que comienzan con `27`, cambiar `2[456]` a `2[4567]`.

## Instalación

### Información general

El programa necesita Python y los paquetes Python pdfplumber y openpyxl. Para leer PDF escaneado, también necesita un motor OCR:

- Con `Python 3.12` o una versión anterior: el paquete rapidocr_onnxruntime. Este es el motor preferido.
- Con `Python 3.13` o una versión más reciente: los paquetes rapidocr y onnxruntime. En una prueba, este motor leyó un dígito erróneamente que el motor preferido leyó correctamente. Por esta razón, `Python 3.12` es recomendable si procesa muchos PDF escaneados.

No hay motor OCR disponible para FreeBSD, OpenBSD, NetBSD o Alpine Linux. En estos sistemas, el programa sólo lee PDFs que contienen texto.

Las siguientes instrucciones ponen todo en una carpeta con el nombre `pdf-totals` en su carpeta de inicio. Los paquetes Python se instalan en un entorno Python privado en `pdf-totals/venv`, o en `pdf-totals\python` en Windows. Nada se instala para todo el sistema, excepto Python en sí donde se declara esto.

### Windows 10 y Windows 11

No necesita derechos de administrador. Abra PowerShell y ejecute el siguiente comando. El comando descarga la distribución incrustable de `Python 3.12` desde python.org en su carpeta de usuario. Luego instala pip y los paquetes necesarios, y descarga los dos scripts. Windows PowerShell 5.1 no admite el operador `&&`, por lo que los pasos están separados con el carácter `;`.

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

El paquete msvc-runtime copia los archivos de tiempo de ejecución Microsoft Visual C++ en la carpeta Python. El motor OCR necesita estos archivos, y algunos sistemas no los tienen. Puede ignorar las advertencias pip sobre scripts que se instalan en una carpeta que no está en PATH.

Usa `curl.exe` y no `curl`. En Windows PowerShell 5.1, `curl` es otro nombre para `Invoke-WebRequest`, que tiene diferentes opciones. Windows 10 incluye `curl.exe` desde la versión 1803.

Para ejecutar el programa:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

Si las descargas están bloqueadas en el ordenador, descargue los archivos `python-3.12.10-embed-amd64.zip` y `get-pip.py` en otro ordenador. Copiarlos en la carpeta Descargas. Luego, reemplace los dos pasos `Invoke-WebRequest` en el comando con los caminos de estos archivos.

Estas instrucciones son para procesadores de 64 bits de Intel y AMD. Para Windows en procesadores ARM, utilice `python-3.12.10-embed-arm64.zip` en su lugar. Esto no ha sido probado.

### macOS

Necesita derechos de administrador sólo para instalar Python. Instala `Python 3.12` con el instalador desde python.org (https://www.python.org/downloads/macos/) o con Homebrew (`brew install python@3.12`). A continuación, ejecute el siguiente comando en la aplicación Terminal:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Si utiliza Homebrew y el comando `python3.12` no se encuentra, use `$(brew --prefix)/bin/python3.12` en su lugar. Los Mac con procesadores Intel y Mac con procesadores Apple son ambos compatibles.

Para ejecutar el programa:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

### Linux

En Linux, el motor OCR se instala con el paquete opencv-python-headless en lugar de opencv-python. El paquete opencv-python-headless no necesita las bibliotecas gráficas del sistema, por lo que no necesita instalar paquetes adicionales del sistema.

Cuando `Python 3.12` está disponible como el comando `python3.12`, el siguiente comando crea el entorno en cada distribución. En Debian 12, utilice el comando `python3` en su lugar, que es `Python 3.11`.

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Para ejecutar el programa:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

Los párrafos siguientes explican cómo obtener `Python 3.12` en distribuciones comunes.

Debian 12 y Ubuntu 24.04 incluyen una versión adecuada de Python. Instala el módulo venv y el programa curl. Esto necesita el usuario root:

```
sudo apt install python3-venv curl
```

En Ubuntu 24.04, el comando es `python3.12`. En Debian 12, el comando es `python3`, que es `Python 3.11`. En ese caso, sustituya `python3.12` con `python3` en el comando anterior. Debian 13 y versiones más nuevas de Ubuntu incluyen `Python 3.13` o una versión más reciente. En estos sistemas, utilice el método con uv que se describe a continuación.

Fedora ofrece un paquete adicional para `Python 3.12`:

```
sudo dnf install python3.12
```

openSUSE Leap 15.6 y openSUSE Tumbleweed ofrecen un paquete adicional para `Python 3.12`:

```
sudo zypper install python312
```

Arch Linux ofrece sólo la versión más reciente de Python. Utilice el método con uv que se describe a continuación.

En cada distribución Linux que utiliza glibc, puede utilizar uv y no necesita el usuario root. uv es un programa único. Se instala en su carpeta de inicio y puede descargar una copia privada de `Python 3.12`.

```
curl -LsSf https://astral.sh/uv/install.sh | sh && mkdir -p ~/pdf-totals && cd ~/pdf-totals && ~/.local/bin/uv venv -p 3.12 venv && ~/.local/bin/uv pip install -p venv/bin/python pdfplumber openpyxl rapidocr_onnxruntime && ~/.local/bin/uv pip uninstall -p venv/bin/python opencv-python && ~/.local/bin/uv pip install -p venv/bin/python opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Alpine Linux utiliza la biblioteca C musl. No hay motor OCR disponible para musl, por lo que el programa sólo lee PDFs que contienen texto:

```
doas apk add python3 curl && mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv venv && venv/bin/pip install pdfplumber openpyxl && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Si su sistema Alpine Linux utiliza `sudo` en lugar de `doas`, reemplace `doas` con `sudo`.

### FreeBSD, OpenBSD y NetBSD

En estos sistemas, el programa sólo lee PDFs que contienen texto. El motor OCR no está disponible para estos sistemas. La biblioteca de renderización PDF que utiliza pdfplumber tampoco está disponible, pero pdfplumber puede leer texto sin ella.

El paquete pdfminer.six necesita el paquete cryptography. Se necesita un compilador Rust para construir cryptography del código fuente, por lo que cryptography se instala desde los paquetes del sistema. Todos los demás paquetes están escritos en Python puro. Están instalados sin sus dependencias opcionales.

FreeBSD. Ejecute el primer comando como el usuario root, por ejemplo con `su` o `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Los nombres de los paquetes contienen la versión de Python, que es `311` en este ejemplo. Si su sistema utiliza otra versión, encuentre los nombres correctos con el comando `pkg search cryptography`. Luego cambia `python3.11` para que coincida.

OpenBSD. Ejecute el primer comando como el usuario root con `doas`:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

NetBSD. Ejecute el primer comando como el usuario root:

```
pkgin install py312-pip py312-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Como en FreeBSD, los nombres de los paquetes contienen la versión de Python, que es `312` en este ejemplo. Encontrar los nombres correctos en su sistema con el comando `pkgin search cryptography`.

Para ejecutar el programa en FreeBSD, OpenBSD o NetBSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

## Actualización

Para actualizar el programa, descargar los dos scripts de nuevo. En Windows:

```
curl.exe -L -o "$env:USERPROFILE\pdf-totals\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$env:USERPROFILE\pdf-totals\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

En macOS y Linux:

```
cd ~/pdf-totals && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

En FreeBSD, utilice `fetch` de la misma manera. En OpenBSD y NetBSD, utilice `ftp`.

Después de un cambio, GitHub puede entregar la versión anterior durante unos minutos. Para obtener una versión específica inmediatamente, reemplace `main` en la dirección con el identificador del commit.

## Solución de problemas

El archivo Excel no se puede guardar. El programa se detiene con código de retorno 1 cuando totales_lotes.xlsx está abierto en Excel. Cierre el archivo y ejecute el programa de nuevo.

Todos o muchos PDFs muestran `0 cantidades`. Ejecute diagnostico.py con uno de estos PDFs. Si el texto que muestra está vacío, el PDF es escaneado y el OCR no lo lee. Si el texto está ahí pero no se detecta nada, el programa no conoce las etiquetas ni el diseño del documento. El texto muestra qué palabras usa el documento. El mismo texto está en sin_resultado.log. Puedes añadir estas palabras a la configuración.

Un total es demasiado grande. Busque el número de lote en detalle_lotes.txt. Este archivo muestra cada cantidad que se agregó y el PDF que vino de. Estas son causas comunes:

- El programa eligió la columna de cantidad equivocada, porque el documento utiliza un título de columna que el programa no sabe.
- El OCR leyó dos números que están unidos como un número.
- El mismo PDF está en la carpeta dos veces.

El OCR no se puede cargar en Windows, y el mensaje contiene `DLL load failed`. Instala el paquete msvc-runtime como se muestra en las instrucciones para Windows.

El OCR no se puede cargar en Linux, y el mensaje contiene `libGL.so.1`. Reemplaza opencv-python con opencv-python-headless como se muestra en las instrucciones para Linux.

pip no puede instalar rapidocr_onnxruntime. Este paquete admite versiones de Python hasta `Python 3.12`. Utilice `Python 3.12`, o instalar los paquetes rapidocr y onnxruntime en su lugar.

## Limitaciones

El OCR lee números manuscritos cuando están escritos claramente. No lee fiablemente la escritura que es difícil de leer, marca de verificación, o abreviaturas escritas a mano como `N/A`. Los PDF con cantidades que no se pueden leer se enumeran en revisar_a_mano.txt y deben ser revisados por una persona.

El OCR puede leer un dígito erróneamente, especialmente cuando un escaneo está borroso o inclinado. Compare algunos totales de PDF escaneados con detalle_lotes.txt.

El formato de un número de lote es establecido por `LOTE_ID`. El programa ignora los números de lotes en otros formatos hasta que el patrón sea cambiado.

El programa espera que las cantidades sean números enteros de piezas. No convierte documentos que cuenten cajas u otras unidades.

## Archivos en el proyecto

- lotes_auto.py es el programa.
- diagnostico.py es la herramienta de diagnóstico. Debe estar en la misma carpeta que lotes_auto.py.
- README.md es una descripción corta. También está disponible como README.de.md y README.es.md.
- La carpeta `docs` contiene este manual en inglés, alemán y español.
- `tools/translate_docs.py` es el script que traduce la documentación.
- `LICENSE` contiene el texto de la licencia.

## Licencia

Licencia MIT. Vea el archivo LICENSE.
