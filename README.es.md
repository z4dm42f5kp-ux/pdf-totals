# pdf-totals

Este archivo se tradujo automáticamente del original inglés con Argos Translate. Si difiere del original inglés, se aplica el original inglés.

Idiomas: [English](README.md), [Deutsch](README.de.md), [Español](README.es.md)

pdf-totals lee una carpeta de documentos PDF, tales como facturas, notas de entrega y formularios de devolución de productos, encuentra cada número de lotes y el número de piezas registradas para él, y escribe el número total de piezas por lote a un archivo Excel.

Lee PDFs que contienen texto directamente y lee PDFs escaneados con reconocimiento de caracteres ópticos (OCR). Comprende varios diseños comunes: tablas con una columna de lote y una columna de cantidad, tablas donde el número de lote aparece dentro de la descripción, formas donde cada columna es un producto, y líneas etiquetadas como `LOTE: 25240310` seguido por `CANTIDAD: 1,250`.

El programa fue escrito para documentos sobre productos farmacéuticos devueltos en México. Sus mensajes y archivos de salida están en español. Se espera que los números de lotes tengan exactamente ocho dígitos y comiencen con `24`, `25` o `26`; esta regla es fácil de cambiar.

## Necesidades

- Python 3.9 o posterior. Python 3.12 se recomienda porque el motor OCR preferido soporta Python hasta 3.12.
- Los paquetes Python pdfplumber y openpyxl.
- Para PDF escaneados: rapidocr_onnxruntime (Python 3.12 o anterior) o rapidocr junto con onnxruntime (Python 3.13 o posterior).

Los derechos de administrador no son necesarios en Windows o macOS. En Linux y BSD, instalar Python de los paquetes del sistema requiere root; en Linux esto se puede evitar con la herramienta uv, como se describe en el manual. OCR está disponible en Windows, macOS y Linux basado en glibc. En FreeBSD, OpenBSD, NetBSD y Alpine Linux sólo se pueden leer PDFs que contienen texto; PDF escaneados se enumeran para revisión manual.

## Inicio rápido

Windows 10 y 11, en PowerShell, sin derechos de administrador:

```
$d="$env:USERPROFILE\pdf-totals"; New-Item -ItemType Directory -Force "$d\python" | Out-Null; Invoke-WebRequest -UseBasicParsing https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip -OutFile "$env:TEMP\py-embed.zip"; Expand-Archive -Force "$env:TEMP\py-embed.zip" "$d\python"; (Get-Content "$d\python\python312._pth") -replace '^#\s*import site','import site' | Set-Content "$d\python\python312._pth"; Invoke-WebRequest -UseBasicParsing https://bootstrap.pypa.io/get-pip.py -OutFile "$d\python\get-pip.py"; & "$d\python\python.exe" "$d\python\get-pip.py" --no-warn-script-location; & "$d\python\python.exe" -m pip install pdfplumber openpyxl rapidocr_onnxruntime msvc-runtime; curl.exe -L -o "$d\lotes_auto.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py; curl.exe -L -o "$d\diagnostico.py" https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

macOS, con Python 3.12 instalado:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

Linux, con Python 3.12 instalado:

```
mkdir -p ~/pdf-totals && cd ~/pdf-totals && python3.12 -m venv venv && venv/bin/pip install pdfplumber openpyxl rapidocr_onnxruntime && venv/bin/pip uninstall -y opencv-python && venv/bin/pip install opencv-python-headless && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/lotes_auto.py && curl -LO https://raw.githubusercontent.com/z4dm42f5kp-ux/pdf-totals/main/diagnostico.py
```

La instalación en otras distribuciones de Linux y en los BSD se describe en el manual.

## Usage

Windows:

```
& "$env:USERPROFILE\pdf-totals\python\python.exe" "$env:USERPROFILE\pdf-totals\lotes_auto.py" "C:\path\to\pdfs"
```

macOS, Linux y BSD:

```
~/pdf-totals/venv/bin/python ~/pdf-totals/lotes_auto.py /path/to/pdfs
```

La carpeta se registra recursivamente. Los archivos cuyo nombre contiene `FALTA REVISION` son saltados. Cuando el programa termina, la carpeta contiene:

- `totales_lotes.xlsx`: una fila por lote con el número total de piezas.
- `detalle_lotes.txt`: cada cantidad que se agregó a cada lote y el PDF que provenía.
- `revisar_a_mano.txt`: PDFs que fueron saltados y PDFs de los que no se podía leer ninguna cantidad.
- `sin_resultado.log`: el texto que fue leído en cada PDF que no produjo cantidad.

Para ver lo que el programa lee y detecta en un solo PDF sin escribir ningún archivo, ejecute `diagnostico.py` con el camino del PDF.

## Documentación

El manual describe cada opción, los diseños compatibles, los archivos de salida, los ajustes de configuración y la instalación en cada plataforma:

- Inglés: [docs/en/MANUAL.md](docs/en/MANUAL.md)
- Deutsch: [docs/de/MANUAL.md](docs/de/MANUAL.md)
- Español: [docs/es/MANUAL.md](docs/es/MANUAL.md)

Los documentos alemanes y españoles se traducen automáticamente de los originales ingleses con Argos Translate. Si no están de acuerdo, los documentos ingleses son autorizados.

## Licencia

MIT. Vea el archivo LICENSE.
