"""Draw the tute-lote logo: a plot of land (a "lote") seen from above, with a dashed
boundary and a survey stake at each corner. Where a plan of a plot would show streets
and blocks, there are four PDF pages.

Usage:
    python tools/make_logo.py

Writes docs/logo.svg with the Python standard library only. If Pillow is installed, it
also writes docs/logo.png (512 x 512 pixels) from the same shapes.
"""
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
SIZE = 256

TIERRA = "#8bc34a"      # land
BORDE = "#33691e"       # boundary and stakes
ESTACA = "#ff8f00"      # stake heads
PAPEL = "#ffffff"
DOBLEZ = "#cfd8dc"      # folded corner
ROJO = "#d32f2f"        # PDF band
LINEA = "#90a4ae"       # text lines on the pages


def figuras():
    """The shapes of the logo, shared by the SVG and the PNG output."""
    f = []
    # The plot: a rounded square of land with a dashed boundary.
    f.append(("rect", 20, 20, 216, 216, 18, TIERRA, None, 0, False))
    f.append(("rect", 30, 30, 196, 196, 10, None, BORDE, 4, True))
    # Survey stakes at the corners.
    for x, y in ((30, 30), (226, 30), (30, 226), (226, 226)):
        f.append(("circle", x, y, 9, ESTACA, BORDE, 3))
    # Four PDF pages instead of streets and blocks.
    for col in (0, 1):
        for fila in (0, 1):
            x, y = 58 + col * 76, 52 + fila * 82
            w, h, d = 64, 74, 16
            f.append(("poly", [(x, y), (x + w - d, y), (x + w, y + d), (x + w, y + h), (x, y + h)],
                      PAPEL, BORDE, 2.5))
            f.append(("poly", [(x + w - d, y), (x + w - d, y + d), (x + w, y + d)], DOBLEZ, BORDE, 2.5))
            f.append(("rect", x + 8, y + 40, w - 16, 16, 3, ROJO, None, 0, False))
            f.append(("text", x + w / 2, y + 52.5, "PDF", PAPEL, 12))
            for i in range(2):
                f.append(("line", x + 9, y + 18 + i * 9, x + w - 22 + (i * 8), y + 18 + i * 9, LINEA, 3))
    return f


def svg():
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {SIZE} {SIZE}" width="{SIZE}" '
           f'height="{SIZE}" role="img" aria-label="tute-lote logo">',
           "  <title>tute-lote</title>"]
    for s in figuras():
        if s[0] == "rect":
            _, x, y, w, h, r, fill, stroke, sw, dash = s
            attrs = f'fill="{fill}"' if fill else 'fill="none"'
            if stroke:
                attrs += f' stroke="{stroke}" stroke-width="{sw}"'
            if dash:
                attrs += ' stroke-dasharray="12 8" stroke-linecap="round"'
            out.append(f'  <rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" {attrs}/>')
        elif s[0] == "circle":
            _, x, y, r, fill, stroke, sw = s
            out.append(f'  <circle cx="{x}" cy="{y}" r="{r}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>')
        elif s[0] == "poly":
            _, pts, fill, stroke, sw = s
            p = " ".join(f"{x:g},{y:g}" for x, y in pts)
            out.append(f'  <polygon points="{p}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" '
                       'stroke-linejoin="round"/>')
        elif s[0] == "line":
            _, x1, y1, x2, y2, color, sw = s
            out.append(f'  <line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" '
                       f'stroke-width="{sw}" stroke-linecap="round"/>')
        elif s[0] == "text":
            _, x, y, texto, color, size = s
            out.append(f'  <text x="{x:g}" y="{y:g}" fill="{color}" font-family="DejaVu Sans, Verdana, '
                       f'Arial, sans-serif" font-size="{size}" font-weight="bold" '
                       f'text-anchor="middle">{texto}</text>')
    out.append("</svg>")
    return "\n".join(out) + "\n"


def png(path, escala=2):
    from PIL import Image, ImageDraw, ImageFont
    k = escala * 4                                  # draw large, then reduce: smooth edges
    img = Image.new("RGBA", (SIZE * k, SIZE * k), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    def guiones(x0, y0, x1, y1, color, ancho, largo=12 * k, hueco=8 * k):
        total = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5
        pos = 0
        while pos < total:
            a, b = pos / total, min(pos + largo, total) / total
            d.line([(x0 + (x1 - x0) * a, y0 + (y1 - y0) * a), (x0 + (x1 - x0) * b, y0 + (y1 - y0) * b)],
                   fill=color, width=ancho)
            pos += largo + hueco

    for s in figuras():
        if s[0] == "rect":
            _, x, y, w, h, r, fill, stroke, sw, dash = s
            box = [x * k, y * k, (x + w) * k, (y + h) * k]
            if dash:
                x0, y0, x1, y1 = box
                for a, b in (((x0, y0), (x1, y0)), ((x1, y0), (x1, y1)), ((x1, y1), (x0, y1)), ((x0, y1), (x0, y0))):
                    guiones(*a, *b, stroke, int(sw * k))
            else:
                d.rounded_rectangle(box, r * k, fill=fill, outline=stroke, width=int(sw * k) if stroke else 0)
        elif s[0] == "circle":
            _, x, y, r, fill, stroke, sw = s
            d.ellipse([(x - r) * k, (y - r) * k, (x + r) * k, (y + r) * k], fill=fill, outline=stroke,
                      width=int(sw * k))
        elif s[0] == "poly":
            _, pts, fill, stroke, sw = s
            d.polygon([(x * k, y * k) for x, y in pts], fill=fill, outline=stroke, width=int(sw * k))
        elif s[0] == "line":
            _, x1, y1, x2, y2, color, sw = s
            d.line([(x1 * k, y1 * k), (x2 * k, y2 * k)], fill=color, width=int(sw * k))
        elif s[0] == "text":
            _, x, y, texto, color, size = s
            try:
                font = ImageFont.truetype("DejaVuSans-Bold.ttf", int(size * k))
            except OSError:
                font = ImageFont.load_default(size=int(size * k))
            d.text((x * k, y * k), texto, fill=color, font=font, anchor="ms")
    img.resize((SIZE * escala, SIZE * escala), Image.LANCZOS).save(path)


def main():
    destino = ROOT / "docs"
    (destino / "logo.svg").write_text(svg(), encoding="utf-8")
    print(f"written: {destino / 'logo.svg'}")
    try:
        png(destino / "logo.png")
        print(f"written: {destino / 'logo.png'}")
    except ImportError:
        print("Pillow is not installed: docs/logo.png was not written.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
