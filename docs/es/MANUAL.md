# Manual de pdf-totals

Este archivo fue traducido automáticamente del original inglés con Argos Translate. Si difiere del original inglés, se aplica el original inglés.

## Nombre

lotes_auto.py - calcular, para cada número de lote, las piezas que fueron retiradas, las piezas que no fueron distribuidas y el número total de piezas para la destrucción

diagnostico.py - mostrar lo que lotes_auto.py lee y detecta en uno o más documentos PDF

## Sinopsis

```
python lotes_auto.py [-h] [--version] [--total-esperado N] FOLDER
python diagnostico.py [-h] [--version] [--lineas N] [--segunda-lectura] PATH
```

## Descripción

lotes_auto.py lee todos los archivos PDF en la carpeta FOLDER y en todos sus subcarpetas. Los documentos se refieren a la retirada de un producto del mercado: formularios de devolución, recibos de entrega, recibos de bienes, listas de inventarios e informes de retiro. Para cada número de lote el programa calcula tres cantidades:

- La cantidad retirada (`Cantidad de producto retirado`). Estas son las piezas que fueron devueltas o recuperadas del mercado. Vienen de formularios de devolución, recibos de entrega y recibos de bienes.
- La cantidad no distribuida (`Cantidad no distribuida`). Estas son las piezas que todavía están en inventario y nunca fueron distribuidas. Vienen de documentos de inventario. Si un lote no tiene documento de inventario, la cantidad se extrae del informe de retiro.
- La cantidad total para la destrucción (`Cantidad total para destrucción`). Esta es la suma de las otras dos cantidades.

El programa revisa el resultado. Un lote no debe tener más de piezas retiradas `48,042` o más de piezas `82,090` que no fueron distribuidas. La cantidad total para la destrucción de todos los lotes debe ser exactamente `344,451`. Estos límites se pueden cambiar en la configuración.

Cuando un PDF contiene texto, el programa lee el texto directamente. Cuando una página no contiene texto, el programa trata la página como imagen escaneada y la lee con reconocimiento de caracteres ópticos (OCR). Si no se instala ningún motor OCR, el programa no lee PDFs escaneados y los enumera para revisión manual.

El programa fue escrito para documentos sobre productos farmacéuticos devueltos en México. Todos los mensajes y archivos de salida están en español. Este manual explica su significado.

diagnostico.py utiliza exactamente el mismo código que lotes_auto.py para leer y analizar documentos. Muestra el texto que leyó de cada PDF, las decisiones que tomó y el resultado. Úsalo para examinar un PDF que no se lee correctamente. No escribe ningún archivo.

## Cómo el programa encuentra las cantidades

El programa no depende solamente de las palabras en los documentos. Examina la estructura de cada tabla y utiliza aritmética para decidir qué número es cuál.

Primero, el programa encuentra las filas que contienen un número de lote. Las columnas de la tabla se forman de las posiciones de los valores en estas filas. Cada valor se clasifica como un número de lote, una fecha, un precio, un número entero, un código o texto.

Luego el programa decide qué columna contiene el número de piezas. Sólo se consideran columnas con números enteros. Números de filas, códigos que repiten en cada fila y años están excluidos. Se utilizan las siguientes pruebas, desde más fuerte hasta más débil:

1. Cantidad multiplicada por el precio unitario igual al precio total en la misma fila. Por ejemplo, `4,408` multiplicado por `$27.00` es `$119,016.00`.
2. La suma de la columna equivale a un número impreso en otro lugar del documento, por ejemplo `TOTAL DE PIEZAS 6,889` o `No CAJAS / PIEZAS 1978`.
3. El título de la columna. Títulos como `PIEZAS`, `CANTIDAD EN PIEZAS`, `No. de piezas`, `RECIBIDA` o `DEVUELTAS` cuentan a favor. Títulos como `FACTURADA` o `DISTRIBUIDA` cuentan con contra, porque describen la cantidad original y no la cantidad devuelta. Títulos con `CAJAS` cuentan contra, porque cuentan cajas y no piezas.

Un informe de retiro es reconocido solo por aritmética. En un informe de retiro, cada fila contiene cinco números para un lote, y dos relaciones mantienen en cada fila:

- manufacturado menos distribuidos iguales no distribuidos;
- no distribuido más retirado equivale al total para la destrucción.

El programa encuentra las columnas que cumplen estas relaciones, por lo que los títulos de columna no son necesarios. La columna retirada es la que nunca es más grande que la columna distribuida. Si una fila del informe no cumple las relaciones, por ejemplo porque el OCR lee un dígito erróneamente, el programa utiliza las relaciones para recuperar el valor correcto.

El tipo de cada cantidad se decide como sigue. Una cantidad no se distribuye si su título de columna o título de fila lo dice, por ejemplo `EXISTENCIA`, `INVENTARIO` o `NO DISTRIBUIDO`. De lo contrario, las palabras de toda la página deciden: páginas sobre devoluciones, retiros o recibos dan cantidades retiradas, y páginas sobre inventario dan cantidades no distribuidas. Si una página no tiene un tipo claro, se utiliza el tipo de las otras páginas del mismo PDF. Si el PDF no tiene un tipo claro, la cantidad cuenta como retirada y se añade una nota.

El programa también lee otros dos diseños:

- Formas en las que cada columna describe un producto. Una fila contiene los números de lote, y otras filas contienen cantidades como `Cantidad Recibida` o `Cantidad Facturada`. La fila es elegida por la prioridad que se describe en la sección Configuración.
- Líneas de texto con etiquetas, como una línea `LOTE: 25240310` y una línea `CANTIDAD EN PIEZAS: 1,250`.

## Comprobaciones y correcciones

Cada cantidad está marcada como verificada o no verificada. Una cantidad se verifica cuando las piezas multiplicadas por el precio unitario igual al precio total impreso, cuando la suma de su tabla equivale a un total impreso, o cuando el total de su lote equivale al informe de retirada.

Si una tabla contradice su propio total impreso de piezas, y exactamente una fila no se verifica, el programa calcula el valor correcto de esa fila del total impreso. Utiliza el valor corregido sólo si el precio total impreso lo confirma. De lo contrario, el PDF está marcado como malo.

Un lote que aparece varias veces en el mismo PDF se cuenta una vez cuando la cantidad es la misma. Esto sucede, por ejemplo, cuando un PDF contiene el mismo documento dos veces. Una página cuya cantidad es la suma de las otras páginas no se cuenta. Si el mismo lote aparece en varias páginas con diferentes cantidades, las cantidades se agregan, y el PDF se enumera para su revisión.

Si un lote excede su límite, no coincide con el informe de retirada, o el total no coincide con el total esperado, el programa lee los PDFs escaneados dudosos por segunda vez con una resolución superior y más contraste. Luego trata de explicar cada diferencia exactamente y sin ambigüedad:

- Si la otra lectura de exactamente un PDF hace que el lote coincida, esa lectura se utiliza.
- Si una fila en un PDF es una copia exacta de una fila en otro PDF, y el lote coincide cuando se retira la copia, la copia no se cuenta.
- Si la otra lectura de exactamente un PDF, o la eliminación de exactamente una fila copiada, hace el total exactamente igual al total esperado, este cambio se aplica.

Un PDF está marcado como malo sólo cuando la prueba es concluyente: cuando sus propios números contradicen sus propios totales impresos en cada lectura, o cuando contiene una copia exacta que es probada por un lote o por el total esperado. Todas las demás dudas se enumeran como sospechosas y no se cambian.

## Reglas para cada valor

- Un número de lote tiene exactamente ocho dígitos y comienza con `24`, `25` o `26`. El programa ignora todos los demás números, por ejemplo precios, fechas, números de facturas y códigos de producto. Si el OCR adjunta otros caracteres a un número de lote, como una barra vertical, el programa los elimina.
- El programa lee números en notación mexicana. Un coma separa los miles y un punto separa los decimales. Por ejemplo, `1,250` es mil doscientos cincuenta, y `1,250.50` es mil doscientos cincuenta y medio. Un número como `4.408`, donde el OCR lee la coma como punto, se lee como `4,408`, porque las piezas siempre son números enteros.
- Se ignoran las marcas de cheque manuscritas junto a un número, como `695 ✓`.
- El programa salta cada archivo cuyo nombre contiene `FALTA REVISION`. El programa ignora la diferencia entre el caso superior y el caso inferior, y el O puede tener un acento. Los archivos eliminados no cuentan hacia ningún total.

## Opciones

Opciones de lotes_auto.py:

- `FOLDER` es la carpeta que contiene los archivos PDF. Se incluyen subcarpetas.
- `--total-esperado N` establece el total previsto para la destrucción. El valor predeterminado es `344,451`. El valor 0 apaga el cheque.
- `--version` muestra el número de versión. El programa se detiene.
- `-h` o `--help` muestra un texto de ayuda corta. El programa se detiene.

Opciones de diagnostico.py:

- `PATH` es un solo archivo PDF, o una carpeta. Si es una carpeta, el programa examina todos los archivos PDF en ella.
- `--lineas N` no muestra más que líneas N de texto para cada PDF. El valor predeterminado es 40.
- `--segunda-lectura` utiliza la segunda lectura OCR, con una resolución más alta y más contraste.
- `--version` muestra el número de versión. El programa se detiene.
- `-h` o `--help` muestra un texto de ayuda corta. El programa se detiene.

## Archivo de salida

lotes_auto.py escribe los siguientes archivos en la carpeta FOLDER. Reemplaza los archivos existentes con los mismos nombres.

`totales_lotes.xlsx` contiene el resultado en una hoja con el nombre `Lotes`. La hoja contiene tres tablas lado a lado: `Cantidad de producto retirado`, `Cantidad no distribuida` y `Cantidad total para destrucción`. Cada tabla tiene las columnas `Lote` y `Cantidad en Piezas` y una fila `TOTAL`. Las tres tablas enumeran los mismos lotes en el mismo orden, por lo que una fila se puede leer a través de las tablas. La tercera tabla y los totales son fórmulas, por lo que se actualizan cuando se cambia un número a mano. Debajo de la tercera tabla, la hoja muestra el total esperado y la diferencia. Las células de lotes que exceden un límite o no coinciden con el informe de retirada son rojas.

`detalle_lotes.txt` lista, para cada lote, cada cantidad que se encontró, el PDF y la página que lo contienen, ya sea verificada y por qué, y si se cuenta. También muestra los valores del informe de retirada. Utilice este archivo para encontrar la fuente de un total que se ve mal.

`revisar_a_mano.txt` enumera todo lo que necesita una persona: un resumen con los totales y el total esperado, los PDFs que están marcados como malos y por qué, las correcciones que se aplicaron, los lotes que exceden un límite o no coinciden con el informe de retirada junto con sus PDFs, los PDFs sospechosos, los lotes que se repiten en un PDF, los lotes que aparecen con la misma cantidad en varios PDF, los PDF sin cantidades legibles, y los archivos que fueron `FALTA REVISION`.

`sin_resultado.log` contiene, para cada PDF que no produjo cantidad, la razón, el texto completo que el programa leyó de él y las decisiones que el programa tomó. Mira este archivo primero cuando un PDF no se lee correctamente.

## Mensajes de consola

Mientras funciona, lotes_auto.py muestra una línea para cada PDF. El texto después de la flecha dice cuántas cantidades retiradas y no distribuidas fueron leídas, si el PDF es un informe de retirada, y si el texto fue leído directamente (`texto`) o con OCR (`OCR`). Otros mensajes:

- `-> omitido (FALTA REVISION)` significa que el programa saltó el archivo debido a su nombre.
- `-> escaneado, OCR no disponible` significa que el PDF es escaneado y que ningún motor OCR está instalado.
- `-> ERROR` significa que el programa no pudo abrir o leer el archivo.
- `Segunda lectura` significa que el programa lee PDFs escaneados dudosos por segunda vez.

Al final, el programa muestra los tres totales y compara el total para la destrucción con el total esperado. La palabra `COINCIDE` significa que los totales son iguales. De lo contrario se muestra la diferencia. Luego enumera los PDF que están marcados como malos.

## Valores de retorno del programa

- `0` significa que el programa terminó. Compruebe el resumen para ver si los totales coinciden.
- `1` significa que el programa no pudo escribir el archivo Excel. Normalmente el archivo está abierto en Excel.
- `2` significa que los argumentos no son válidos, o que la carpeta no existe.

## Configuración

Los ajustes están al principio del archivo lotes_auto.py. Puedes cambiarlos con cualquier editor de texto. Guarda el archivo en la codificación UTF-8.

`TOTAL_ESPERADO` es el total previsto para la destrucción. El valor predeterminado es `344451`. El valor `None` apaga el cheque.

`MAX_RETIRADO_POR_LOTE` y `MAX_NO_DISTRIBUIDA_POR_LOTE` son los límites para cada lote. Los valores predeterminados son `48042` y `82090`.

`USAR_REPORTE_PARA_RETIRADO` decide cómo se utilizan las cantidades retiradas de un informe de retiro. Con el valor predeterminado `False`, sólo se utilizan para comprobar los documentos de devolución. Con `True`, también se utilizan para lotes que no tienen documento de devolución en la carpeta.

`DUPLICADOS_EN_ARCHIVO` decide cómo se cuenta un lote que se repite en el mismo PDF. Con el valor predeterminado `iguales`, se cuenta una vez sólo cuando las cantidades son iguales. Con `todos`, siempre se cuenta una vez, y se mantiene la mayor cantidad.

`OMITIR_NOMBRE` es el patrón para los nombres de los archivos que se saltan. El valor predeterminado coincide con `FALTA REVISION`.

`LOTE_ID` es el patrón para un número de lote válido. El valor predeterminado acepta exactamente ocho dígitos que comienzan con `24`, `25` o `26`. Por ejemplo, para aceptar también los números de lote que comienzan con `27`, cambiar `2[456]` a `2[4567]`.

`OCR_DPI` y `OCR_DPI_REINTENTO` son las resoluciones de la primera y segunda lectura de OCR. `CONFIANZA_MINIMA` es la confianza OCR abajo que una cantidad que no se verifica se enumera como sospechosa.

`PRIORIDAD_RENGLON` decide qué fila de cantidad se utiliza en formas en las que cada columna describe un producto.

`SEMANTICA` y `TIPO_DOCUMENTO` contienen las palabras que se utilizan como una señal entre otros para comprender los títulos de columna y reconocer el tipo de página.

## Instalación

### Información general

El programa necesita Python y los paquetes Python pdfplumber y openpyxl. Para leer PDF escaneado, también necesita un motor OCR:

- Con `Python 3.12` o una versión anterior: el paquete rapidocr_onnxruntime. Este es el motor OCR preferido.
- Con `Python 3.13` o una versión más reciente: los paquetes rapidocr y onnxruntime. En una prueba, este motor OCR leyó un dígito erróneamente. El motor OCR preferido lee correctamente el mismo dígito. Por esta razón, `Python 3.12` es recomendable si procesa muchos PDF escaneados.

No hay motor OCR disponible para FreeBSD, OpenBSD, NetBSD o Alpine Linux. En estos sistemas, el programa sólo lee PDFs que contienen texto.

Las siguientes instrucciones ponen todo en una carpeta con el nombre `pdf-totals` en su carpeta de inicio. Los paquetes Python se instalan en un entorno Python privado en `pdf-totals/venv`, o en `pdf-totals\python` en Windows. Nada se instala para todo el sistema, excepto Python en sí donde se declara esto.

### Windows 10 y Windows 11

No necesita derechos de administrador. Abra PowerShell y ejecute el siguiente comando. El comando descarga la distribución incrustable de `Python 3.12` desde python.org en su carpeta de usuario. Luego instala pip y los paquetes necesarios, y descarga los dos scripts. Windows PowerShell 5.1 no admite el operador `&&`, por lo que los pasos están separados con el carácter `;`.

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

El paquete msvc-runtime copia los archivos de tiempo de ejecución Microsoft Visual C++ en la carpeta Python. El motor OCR necesita estos archivos, y algunos sistemas no los tienen. Puede ignorar las advertencias pip sobre scripts que se instalan en una carpeta que no está en PATH.

Use `curl.exe` y no `curl`. En Windows PowerShell 5.1, `curl` es otro nombre para `Invoke-WebRequest`, que tiene diferentes opciones. Windows 10 incluye `curl.exe` desde la versión 1803.

Para ejecutar el programa:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

Si las descargas están bloqueadas en el ordenador, descargue los archivos `python-3.12.10-embed-amd64.zip` y `get-pip.py` en otro ordenador. Copiarlos en la carpeta Descargas. Luego, reemplace los dos pasos `Invoke-WebRequest` en el comando con los caminos de estos archivos.

Estas instrucciones son para procesadores de 64 bits de Intel y AMD. Para Windows en procesadores ARM, utilice `python-3.12.10-embed-arm64.zip` en su lugar. Esto no ha sido probado.

### macOS

Necesita derechos de administrador sólo para instalar Python. Instala `Python 3.12` con el instalador desde python.org (https://www.python.org/downloads/macos/) o con Homebrew (`brew install python@3.12`). Luego ejecute el siguiente comando en la aplicación Terminal:

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

Los párrafos siguientes explican cómo obtener `Python 3.12` en distribuciones comunes de Linux.

Debian 12 y Ubuntu 24.04 incluyen una versión adecuada de Python. Instala el módulo venv y el programa curl. Esto necesita el usuario de root:

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

En estos sistemas, el programa sólo lee PDFs que contienen texto. El motor OCR no está disponible para estos sistemas. La biblioteca de renderizado PDF que utiliza pdfplumber tampoco está disponible, pero pdfplumber puede leer texto sin ella.

El paquete pdfminer.six necesita el paquete cryptography. Se necesita un compilador Rust para construir cryptography del código fuente, por lo que cryptography está instalado de los paquetes del sistema. Todos los demás paquetes están escritos en Python puro. Están instalados sin sus dependencias opcionales.

En FreeBSD, ejecute el primer comando como el usuario root, por ejemplo con `su` o `sudo`:

```
pkg install py311-pip py311-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.11 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && fetch https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Los nombres de los paquetes contienen la versión de Python, que es `311` en este ejemplo. Si su sistema utiliza otra versión, encuentre los nombres correctos con el comando `pkg search cryptography`. A continuación, utilice el comando matching en lugar de `python3.11`, por ejemplo `python3.12`.

En OpenBSD, ejecute el primer comando como el usuario root con `doas`:

```
doas pkg_add py3-pip py3-cryptography
```

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3 -m venv --system-site-packages --without-pip venv && venv/bin/python -m pip install --no-deps pdfplumber pdfminer.six openpyxl et_xmlfile charset-normalizer && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && ftp https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

En NetBSD, ejecute el primer comando como el usuario root:

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

Para actualizar el programa, descargue los dos scripts de nuevo. En Windows:

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

El total no coincide con el total esperado. Abre revisar_a_mano.txt. La sección sobre lotes que no coinciden muestra qué lotes exceden un límite o difieren del informe de retirada, y qué PDFs contribuyen a ellos. Entonces busca estos lotes en detalle_lotes.txt. Las causas comunes son:

- Un PDF que pertenece a la retirada falta de la carpeta, o se llama `FALTA REVISION`.
- El OCR leyó un dígito erróneamente. Las dos lecturas de un PDF se comparan, y una diferencia entre ellas se observa como `dudosa`.
- El mismo documento está en la carpeta dos veces bajo nombres diferentes.
- El programa utilizó la columna equivocada para la cantidad, porque el documento utiliza un diseño que el programa no sabe. Ejecute diagnostico.py con este PDF para ver qué columna fue elegida y por qué.

Todos o muchos PDF no muestran cantidades. Ejecute diagnostico.py con uno de estos PDFs. Si el texto que muestra está vacío, el PDF es escaneado y el OCR no lo lee. Si el texto está ahí pero no se detecta nada, el programa no conoce el diseño del documento.

El OCR no se puede cargar en Windows, y el mensaje contiene `DLL load failed`. Instala el paquete msvc-runtime como se muestra en las instrucciones para Windows.

El OCR no se puede cargar en Linux, y el mensaje contiene `libGL.so.1`. Reemplaza opencv-python con opencv-python-headless como se muestra en las instrucciones para Linux.

pip no puede instalar rapidocr_onnxruntime. Este paquete admite versiones de Python hasta `Python 3.12`. Utilice `Python 3.12`, o instalar los paquetes rapidocr y onnxruntime en su lugar.

## Limitaciones

El OCR lee números manuscritos cuando están escritos claramente. No lee fiablemente la escritura que es difícil de leer, o abreviaturas escritas a mano como `N/A`. Por lo tanto, no se leen cuantitativas que se escriben a mano en forma. Estos PDF están listados en revisar_a_mano.txt y deben ser revisados por una persona.

El OCR puede leer un dígito erróneamente, especialmente cuando un escaneo está borroso. Las cantidades verificadas por un precio, un total impreso o el informe de retiro son fiables. Las cantidades que no se verifican deben compararse con el PDF.

La segunda lectura de OCR lleva tiempo. Cuando el total no coincide, todos los PDF escaneados con cantidades que no se verifican se leen de nuevo, por lo que una carrera puede tardar aproximadamente el doble de tiempo.

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
