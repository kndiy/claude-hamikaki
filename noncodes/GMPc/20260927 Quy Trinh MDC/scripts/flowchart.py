"""SVG flowcharts for the MD&C SOPs, rendered to PNG with headless Microsoft Edge (no extra packages).

Style copies the MDC-QT-05 chart: navy title/end boxes, teal steps, white side notes,
lavender input/output boxes, yellow decision diamonds, dashed feedback arrows.
"""
import os
import subprocess
import tempfile
from xml.sax.saxutils import escape

NAVY, TEAL, LAV, YELLOW, WHITE = "#1b1464", "#00b5a5", "#e8e8f4", "#fff2cc", "#ffffff"
FONT = "Calibri, 'Segoe UI', Arial, sans-serif"
EDGE = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"]

# box kinds: (fill, text colour, corner radius)
KINDS = {
    "title": (NAVY, WHITE, 0),
    "step": (TEAL, WHITE, 14),
    "note": (WHITE, NAVY, 14),
    "io": (LAV, NAVY, 14),
}


class Chart:
    def __init__(self, w, h):
        self.w, self.h = w, h
        self.shapes, self.lines = [], []

    def _text(self, cx, cy, lines, color, size, bold_first):
        lh = size * 1.25
        y0 = cy - lh * (len(lines) - 1) / 2 + size * 0.35
        for i, s in enumerate(lines):
            weight = "bold" if (bold_first and i == 0) else "normal"
            self.shapes.append(f'<text x="{cx}" y="{y0 + i * lh:.1f}" text-anchor="middle" font-size="{size}" '
                               f'font-weight="{weight}" fill="{color}">{escape(s)}</text>')

    def box(self, cx, y, w, h, lines, kind="step", bold_first=None, size=17):
        fill, color, radius = KINDS[kind]
        if bold_first is None:
            bold_first = kind in ("step", "io") and len(lines) > 1 or kind == "step"
        x = cx - w / 2
        self.shapes.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{radius}" fill="{fill}" '
                           f'stroke="{NAVY}" stroke-width="1.6"/>')
        self._text(cx, y + h / 2, lines, color, size, bold_first)
        return {"cx": cx, "top": y, "bot": y + h, "l": x, "r": x + w, "mid": y + h / 2}

    def diamond(self, cx, y, w, h, lines, size=15):
        cy = y + h / 2
        pts = f"{cx},{y} {cx + w / 2},{cy} {cx},{y + h} {cx - w / 2},{cy}"
        self.shapes.append(f'<polygon points="{pts}" fill="{YELLOW}" stroke="{NAVY}" stroke-width="1.6"/>')
        self._text(cx, cy, lines, NAVY, size, False)
        return {"cx": cx, "top": y, "bot": y + h, "l": cx - w / 2, "r": cx + w / 2, "mid": cy}

    def arrow(self, points, dashed=False):
        d = "M " + " L ".join(f"{px},{py}" for px, py in points)
        dash = ' stroke-dasharray="7 5"' if dashed else ""
        self.lines.append(f'<path d="{d}" fill="none" stroke="{NAVY}" stroke-width="1.6"{dash} marker-end="url(#ah)"/>')

    def link(self, points):
        """Dashed connector without arrowhead (box to side note)."""
        d = "M " + " L ".join(f"{px},{py}" for px, py in points)
        self.lines.append(f'<path d="{d}" fill="none" stroke="{NAVY}" stroke-width="1.3" stroke-dasharray="4 4"/>')

    def label(self, x, y, s, anchor="middle", rotate=None, size=14):
        tr = f' transform="rotate({rotate} {x} {y})"' if rotate is not None else ""
        self.lines.append(f'<text x="{x}" y="{y}" text-anchor="{anchor}" font-size="{size}" fill="{NAVY}"{tr}>'
                          f'{escape(s)}</text>')

    def down(self, a, b):
        self.arrow([(a["cx"], a["bot"]), (b["cx"], b["top"])])

    def svg(self):
        return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.w}" height="{self.h}" '
                f'viewBox="0 0 {self.w} {self.h}" font-family="{FONT}"><defs><marker id="ah" viewBox="0 0 10 10" '
                f'refX="9" refY="5" markerWidth="9" markerHeight="9" orient="auto-start-reverse">'
                f'<path d="M0,0 L10,5 L0,10 z" fill="{NAVY}"/></marker></defs>'
                f'<rect width="{self.w}" height="{self.h}" fill="#ffffff"/>'
                + "".join(self.lines) + "".join(self.shapes) + "</svg>")

    def render(self, out_png, scale=2):
        """Write <out>.svg next to the PNG and screenshot it with headless Edge. Returns (w, h)."""
        svg = self.svg()
        with open(os.path.splitext(out_png)[0] + ".svg", "w", encoding="utf-8") as f:
            f.write(svg)
        tmp = tempfile.mkdtemp()
        html = os.path.join(tmp, "chart.html")
        with open(html, "w", encoding="utf-8") as f:
            f.write('<!doctype html><meta charset="utf-8"><style>html,body{margin:0;padding:0;'
                    f'overflow:hidden;background:#fff}}svg{{display:block}}</style>{svg}')
        edge = next(p for p in EDGE if os.path.exists(p))
        subprocess.run([edge, "--headless", "--disable-gpu", "--hide-scrollbars",
                        f"--user-data-dir={os.path.join(tmp, 'profile')}",
                        f"--force-device-scale-factor={scale}", f"--window-size={self.w},{self.h}",
                        f"--screenshot={os.path.abspath(out_png)}", "file:///" + html.replace("\\", "/")],
                       check=True, capture_output=True, timeout=120)
        return self.w, self.h
