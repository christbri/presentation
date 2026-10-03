"""Generates the two art pieces for the cover and the closing slide.

  sphere.png    a halftone sphere of dots, lit from the top left, with eight
                bright nodes wired together on its surface (the eight agents)
  question.png  a question mark built from the same dots, with loose
                particles drifting in around it

Run:  python3 pptx/art/make_art.py   (needs Pillow)
Both are white/gray on transparent, made for the black slides.
"""
import math, random
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

HERE = Path(__file__).parent
SS = 3  # supersample, then downscale for clean dot edges
random.seed(7)


def shade(v):
    """0..1 → white..gray"""
    g = int(70 + 185 * v)
    return (g, g, g, 255)


def sphere(size=1400, out="sphere.png"):
    S = size * SS
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    cx = cy = S / 2
    R = S * 0.40
    step = S / 58
    light = (-0.55, -0.62, 0.56)
    ln = math.sqrt(sum(c * c for c in light))
    light = tuple(c / ln for c in light)

    # halftone dots
    y = cy - R
    while y <= cy + R:
        x = cx - R
        while x <= cx + R:
            nx, ny = (x - cx) / R, (y - cy) / R
            rr = nx * nx + ny * ny
            if rr <= 1:
                nz = math.sqrt(1 - rr)
                lam = max(0.0, nx * light[0] + ny * light[1] + nz * light[2])
                lam = lam ** 1.15
                r = step * 0.46 * (0.10 + 0.90 * lam)
                if r > step * 0.05:
                    d.ellipse((x - r, y - r, x + r, y + r), fill=shade(0.25 + 0.75 * lam))
            x += step
        y += step

    # a thin equator-like orbit, tilted, passing in front of the sphere
    tilt = math.radians(-18)
    pts = []
    for k in range(721):
        t = 2 * math.pi * k / 720
        ex, ey = R * 1.22 * math.cos(t), R * 0.30 * math.sin(t)
        pts.append((cx + ex * math.cos(tilt) - ey * math.sin(tilt), cy + ex * math.sin(tilt) + ey * math.cos(tilt), math.sin(t)))
    for (x1, y1, z1), (x2, y2, _) in zip(pts, pts[1:]):
        front = z1 > 0 or (x1 - cx) ** 2 + (y1 - cy) ** 2 > R * R
        d.line((x1, y1, x2, y2), fill=(190, 190, 190, 255) if front else (90, 90, 90, 140), width=int(SS * 2))

    # eight agents on the lit face, wired to the first one
    nodes = []
    while len(nodes) < 8:
        u, v = random.uniform(-0.85, 0.55), random.uniform(-0.85, 0.6)
        if u * u + v * v < 0.72 and all((u - a) ** 2 + (v - b) ** 2 > 0.07 for a, b in nodes):
            nodes.append((u, v))
    nodes.sort(key=lambda p: p[0] + p[1])
    P = [(cx + u * R, cy + v * R) for u, v in nodes]
    hub = P[2]
    for p in P:
        if p != hub:
            d.line((*hub, *p), fill=(255, 255, 255, 170), width=int(SS * 2.2))
    for a, b in zip(P, P[1:]):
        d.line((*a, *b), fill=(255, 255, 255, 70), width=int(SS * 1.4))
    for i, (x, y) in enumerate(P):
        r = step * (0.95 if (x, y) == hub else 0.62)
        d.ellipse((x - r * 2.2, y - r * 2.2, x + r * 2.2, y + r * 2.2), outline=(255, 255, 255, 120), width=int(SS * 1.6))
        d.ellipse((x - r, y - r, x + r, y + r), fill=(255, 255, 255, 255))

    # the moving node on the orbit
    x, y, _ = pts[95]
    d.ellipse((x - step * 0.5, y - step * 0.5, x + step * 0.5, y + step * 0.5), fill=(255, 255, 255, 255))

    img.resize((size, size), Image.LANCZOS).save(HERE / out)


def question(w=1100, h=1400, out="question.png"):
    W, H = w * SS, h * SS
    font = ImageFont.truetype(str(HERE / "eb-garamond-latin-500-italic.woff"), int(H * 1.05))
    mask = Image.new("L", (W, H), 0)
    md = ImageDraw.Draw(mask)
    l, t, r, b = md.textbbox((0, 0), "?", font=font)
    md.text(((W - (r - l)) / 2 - l, (H - (b - t)) / 2 - t), "?", font=font, fill=255)
    near = mask.filter(ImageFilter.GaussianBlur(H / 22))

    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    step = W / 44
    y = step / 2
    while y < H:
        x = step / 2
        while x < W:
            cov = mask.getpixel((int(x), int(y))) / 255
            glow = near.getpixel((int(x), int(y))) / 255
            vert = 1 - 0.35 * (y / H)  # a little brighter at the top
            if cov > 0.5:
                rr = step * 0.44 * (0.55 + 0.45 * vert)
                d.ellipse((x - rr, y - rr, x + rr, y + rr), fill=shade(0.55 + 0.45 * vert))
            elif glow > 0.04 and random.random() < glow * 1.6:
                # particles drifting in toward the glyph
                jx, jy = random.uniform(-0.3, 0.3) * step, random.uniform(-0.3, 0.3) * step
                rr = step * (0.08 + 0.26 * glow)
                d.ellipse((x + jx - rr, y + jy - rr, x + jx + rr, y + jy + rr), fill=shade(0.1 + 0.6 * glow))
            x += step
        y += step

    img.resize((w, h), Image.LANCZOS).save(HERE / out)


if __name__ == "__main__":
    sphere()
    question()
    print("Wrote", HERE / "sphere.png", "and", HERE / "question.png")
