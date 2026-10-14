# pdf-totals

Este archivo se tradujo automáticamente del original inglés con Argos Translate. Si difiere del original inglés, se aplica el original inglés.

Idiomas: [English](README.md), [Deutsch](README.de.md), [Español](README.es.md)

pdf-totals lee una carpeta de documentos PDF, tales como facturas, notas de entrega y formularios para productos devueltos. Encuentra cada número de lotes y el número de piezas que se registran para cada lote. A continuación, escribe el número total de piezas para cada lote a un archivo Excel.

El programa lee el texto de PDF que contiene texto. Lee PDF escaneado con reconocimiento de caracteres ópticos (OCR). Reconoce varios diseños comunes: tablas con una columna de lote y una columna de cantidad, tablas en las que el número de lote es parte de la descripción, formas en las que cada columna describe un producto, y líneas con etiquetas como `LOTE: 25240310` y `CANTIDAD: 1,250`.

El programa fue escrito para documentos sobre productos farmacéuticos devueltos en México. Sus mensajes y archivos de salida están en español. Un número de lote debe tener exactamente ocho dígitos y comenzar con `24`, `25` o `26`. Esta regla es fácil de cambiar.

## Necesidades

- Python 3.9 o una versión más nueva. `Python 3.12` se recomienda, ya que el motor OCR preferido soporta las versiones de Python hasta `Python 3.12`.
- Los paquetes Python pdfplumber y openpyxl.
- Para PDF escaneado: el paquete rapidocr_onnxruntime con `Python 3.12` o una versión anterior, o los paquetes rapidocr y onnxruntime con `Python 3.13` o una versión más reciente.

En Windows y macOS no necesita derechos de administrador. En Linux y en los sistemas BSD, necesita que el usuario root instale Python desde los paquetes del sistema. En Linux se puede evitar esto con el programa uv, como describe el manual.

OCR está disponible en Windows, en macOS y en distribuciones Linux que usan glibc. En FreeBSD, OpenBSD, NetBSD y Alpine Linux, el programa sólo lee PDFs que contienen texto. Lista PDF escaneado para revisión manual.

## Inicio rápido

Windows 10 y Windows 11, en PowerShell, sin derechos de administrador:

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

macOS, cuando se instala `Python 3.12`:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Linux, cuando se instala `Python 3.12`:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

El manual describe la instalación en otras distribuciones de Linux y en FreeBSD, OpenBSD y NetBSD.

## Usage

Windows:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

macOS, Linux, FreeBSD, OpenBSD y NetBSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

El programa también busca todos los subcarpetas. Se salta cada archivo cuyo nombre contiene `FALTA REVISION`. Cuando el programa termina, la carpeta contiene estos archivos:

- totales_lotes.xlsx contiene una fila para cada lote con el número total de piezas.
- detalle_lotes.txt enumera cada cantidad que se agregó a cada lote y el PDF que vino de.
- revisar_a_mano.txt enumera los PDFs saltados y los PDF de los que no se puede leer ninguna cantidad.
- sin_resultado.log contiene el texto que fue leído en cada PDF que no produjo cantidad.

Para ver lo que el programa lee y detecta en un PDF, ejecute diagnostico.py con el camino del PDF. diagnostico.py no escribe ningún archivo.

## Documentación

El manual describe todas las opciones, los diseños compatibles, los archivos de salida, la configuración y la instalación en cada sistema:

- Inglés: [docs/en/MANUAL.md](docs/en/MANUAL.md)
- Deutsch: [docs/de/MANUAL.md](docs/de/MANUAL.md)
- Español: [docs/es/MANUAL.md](docs/es/MANUAL.md)

Los documentos alemán y español se traducen automáticamente de los documentos ingleses con Argos Translate. Si una traducción difiere del documento inglés, se aplica el documento inglés.

## Licencia

Licencia MIT. Vea el archivo LICENSE.
