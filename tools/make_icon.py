"""Draws the BILLFLOW app icon (Tron style: a glowing feather pen over a light-grid and binary belt) for Windows.

    python tools/make_icon.py

Writes src/billflow/assets/billflow.ico (16-256 px, used by the .exe files, shortcuts and installer) and
src/billflow/assets/billflow.png (256 px, used for the window icon). Large sizes get the binary belt, the grid
floor and the feather's barbs; small sizes drop them and thicken the lines so the quill stays recognisable in the taskbar.
"""
import math
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "src" / "billflow" / "assets"
BLACK, CYAN, ICE, ORANGE = (0, 0, 0), (0, 216, 255), (210, 250, 255), (255, 138, 28)
S = 1024                                                     # draw big, then shrink for each size
FONT_FILES = ["C:/Windows/Fonts/consolab.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
              "/Library/Fonts/Courier New Bold.ttf"]


def font(size):
    for f in FONT_FILES:
        if Path(f).exists():
            return ImageFont.truetype(f, size)
    return ImageFont.load_default(size)


def glow(base, draw_fn, radius, strength=2):
    """Draws something plus soft blurred copies behind it, so it looks lit up."""
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw_fn(ImageDraw.Draw(layer))
    blurred = layer.filter(ImageFilter.GaussianBlur(radius))
    for _ in range(strength):
        base.alpha_composite(blurred)
    base.alpha_composite(layer)


# ---- the feather pen --------------------------------------------------------------------------------------------
NIB, TIP = (0.25 * S, 0.80 * S), (0.80 * S, 0.17 * S)        # pen point (bottom left) and feather tip (top right)


def shaft(t):
    """A point on the quill's shaft, t=0 at the nib, t=1 at the tip, with a gentle curve."""
    x = NIB[0] + (TIP[0] - NIB[0]) * t
    y = NIB[1] + (TIP[1] - NIB[1]) * t
    bend = math.sin(math.pi * t) * 0.035 * S
    nx, ny = _normal()
    return x + nx * bend, y + ny * bend


def _normal():
    dx, dy = TIP[0] - NIB[0], TIP[1] - NIB[1]
    n = math.hypot(dx, dy)
    return -dy / n, dx / n                                   # points up-left of the shaft


def vane_edge(side, notches):
    """Outline of one side of the feather. side=+1 is the wide upper-left side, -1 the narrow lower-right side."""
    nx, ny = _normal()
    start, pts = 0.24, []
    for i in range(61):
        t = start + (1 - start) * i / 60
        u = (t - start) / (1 - start)
        width = math.sin(math.pi * u ** 0.75) * (0.17 if side > 0 else 0.095) * S
        for at, depth in notches:                            # small splits in the vane, like a real feather
            if abs(u - at) < 0.035:
                width *= 1 - depth * (1 - abs(u - at) / 0.035)
        x, y = shaft(t)
        lean = -0.05 * S * u                                 # barbs sweep back towards the nib
        dx, dy = (TIP[0] - NIB[0]), (TIP[1] - NIB[1])
        d = math.hypot(dx, dy)
        pts.append((x + side * nx * width + dx / d * lean, y + side * ny * width + dy / d * lean))
    return pts


def feather(detail):
    up = vane_edge(+1, [(0.32, 0.35), (0.58, 0.28)] if detail else [(0.45, 0.3)])
    down = vane_edge(-1, [(0.45, 0.3)] if detail else [])
    return up, down


def draw(detail):
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    pad, r = 40, 210
    box = (pad, pad, S - pad, S - pad)
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle(box, r, fill=255)
    tile = Image.new("RGBA", (S, S), BLACK + (255,))

    if detail:                                               # binary belt behind everything, like the app
        d = ImageDraw.Draw(tile)
        rnd = random.Random(7)
        f = font(44)
        for i, y in enumerate(range(66, S - 40, 60)):
            shade = (40, 22, 4) if i % 5 == 3 else [(2, 32, 42), (3, 44, 57), (4, 56, 74), (2, 26, 34)][i % 4]
            d.text((rnd.randint(-220, 0), y), "".join(rnd.choice("01") for _ in range(40)), font=f, fill=shade)
    if detail:                                               # Tron light-grid floor fading towards the horizon
        d = ImageDraw.Draw(tile)
        horizon, vx = 0.56 * S, 0.5 * S
        for k in range(-9, 10):
            d.line((vx, horizon, vx + k * 150, S + 60), fill=(0, 70, 92), width=4)
        for i in range(1, 8):
            y = horizon + (S - horizon) * (i / 7) ** 1.9
            d.line((0, y, S, y), fill=(0, 52 + i * 8, 70 + i * 9), width=4)
        d.line((0, horizon, S, horizon), fill=(0, 140, 170), width=5)
    tile.putalpha(mask)
    img.alpha_composite(tile)

    w = 30 if detail else 60                                 # glowing frame
    glow(img, lambda d: d.rounded_rectangle((pad + w / 2, pad + w / 2, S - pad - w / 2, S - pad - w / 2), r - w / 2,
                                            outline=CYAN + (255,), width=w), 26 if detail else 16)

    # orange ink trail flowing from the nib, like a signature on the invoice
    dx0, dy0 = TIP[0] - NIB[0], TIP[1] - NIB[1]
    n0 = math.hypot(dx0, dy0)
    reach = 0.05 if detail else 0.09
    nib_tip = (NIB[0] - dx0 / n0 * reach * n0, NIB[1] - dy0 / n0 * reach * n0)
    ink = []
    for i in range(0, 41):
        u = i / 40
        ink.append((nib_tip[0] + u * 0.5 * S, nib_tip[1] + math.sin(u * math.pi * 2.2) * 0.04 * S + u * 0.01 * S))
    glow(img, lambda d: d.line(ink, fill=ORANGE + (255,), width=22 if detail else 46, joint="curve"), 18, 2)

    up, down = feather(detail)
    outline = up + list(reversed(down))
    fill = (0, 44, 58, 255) if detail else (0, 150, 185, 255)
    lw = 14 if detail else 34
    glow(img, lambda d: d.polygon(outline, fill=fill), 1, 0)
    if detail:                                               # barbs: fine lines across the vane
        def barbs(d):
            for edge in (up, down):
                for i in range(4, len(edge) - 3, 3):
                    t = 0.24 + 0.76 * i / 60
                    d.line((shaft(t), edge[i]), fill=(0, 150, 180, 255), width=4)
        glow(img, barbs, 4, 1)
    glow(img, lambda d: d.line(outline + [outline[0]], fill=CYAN + (255,), width=lw, joint="curve"), 22 if detail else 14, 2)
    spine = [shaft(i / 50) for i in range(51)]               # the quill's shaft, brightest line
    glow(img, lambda d: d.line(spine, fill=ICE + (255,), width=16 if detail else 30, joint="curve"), 14, 2)

    nx, ny = _normal()                                       # the pen nib: slim, with a slit and a breather hole
    dx, dy = TIP[0] - NIB[0], TIP[1] - NIB[1]
    n = math.hypot(dx, dy)
    ux, uy = dx / n, dy / n

    def at(t, side=0.0):                                     # point along the pen, pushed sideways by 'side'
        return (NIB[0] + ux * t * n + nx * side, NIB[1] + uy * t * n + ny * side)

    hw = (0.042 if detail else 0.04) * S
    tipt = -0.05 if detail else -0.09                        # longer, pointier nib at small sizes so it is not a blob
    nib = [at(tipt), at(0.06, hw), at(0.15, hw * 0.75), at(0.17, hw * 0.45), at(0.17, -hw * 0.45),
           at(0.15, -hw * 0.75), at(0.06, -hw)]
    glow(img, lambda d: d.polygon(nib, fill=(0, 30, 40, 255) if detail else ICE + (255,)), 1, 0)
    glow(img, lambda d: d.line(nib + [nib[0]], fill=ICE + (255,), width=12 if detail else 14, joint="curve"), 12 if detail else 6, 2)
    if detail:
        hole = at(0.075)
        glow(img, lambda d: (d.line((at(-0.045), hole), fill=ICE + (255,), width=7),
                             d.ellipse((hole[0] - 13, hole[1] - 13, hole[0] + 13, hole[1] + 13), fill=ICE + (255,))), 6, 1)
    return img


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    big, small = draw(True), draw(False)
    frames = [(small if n <= 32 else big).resize((n, n), Image.LANCZOS) for n in (16, 24, 32, 48, 64, 128, 256)]
    frames[-1].save(OUT / "billflow.ico", sizes=[f.size for f in frames], append_images=frames[:-1])
    frames[-1].save(OUT / "billflow.png")
    print("Saved", OUT / "billflow.ico", "and", OUT / "billflow.png")


if __name__ == "__main__":
    main()
