"""Draws the BILLFLOW app icon in the Neon style and saves it for Windows and the window title bar.

    python tools/make_icon.py

Writes src/billflow/assets/billflow.ico (16-256 px, used by the .exe, shortcuts and installer) and
src/billflow/assets/billflow.png (256 px, used for the window icon). Large sizes get the binary
"conveyor belt" in the background; small sizes drop it and thicken the lines so they stay sharp.
"""
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "src" / "billflow" / "assets"
BG, CYAN, MAG, GREEN = (5, 6, 15), (0, 240, 255), (255, 43, 214), (0, 255, 156)
FONT_FILES = ["C:/Windows/Fonts/consolab.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
              "/Library/Fonts/Courier New Bold.ttf"]
S = 1024                                                     # draw big, then shrink for each size


def font(size):
    for f in FONT_FILES:
        if Path(f).exists():
            return ImageFont.truetype(f, size)
    return ImageFont.load_default(size)


def glow(base, draw_fn, radius, strength=2):
    """Draws something, plus a soft blurred copy behind it so it looks lit up."""
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw_fn(ImageDraw.Draw(layer))
    blurred = layer.filter(ImageFilter.GaussianBlur(radius))
    for _ in range(strength):
        base.alpha_composite(blurred)
    base.alpha_composite(layer)


def draw(detail):
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    pad, r = 40, 210
    box = (pad, pad, S - pad, S - pad)
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle(box, r, fill=255)
    tile = Image.new("RGBA", (S, S), BG + (255,))

    if detail:                                               # binary belt, like the app background
        d = ImageDraw.Draw(tile)
        rnd = random.Random(7)
        f = font(44)
        for i, y in enumerate(range(70, S - 40, 62)):
            shade = [(8, 46, 56), (10, 62, 60), (11, 47, 74), (26, 16, 56)][i % 4]
            d.text((rnd.randint(-200, 0), y), "".join(rnd.choice("01") for _ in range(40)), font=f, fill=shade)
    tile.putalpha(mask)
    img.alpha_composite(tile)

    w = 34 if detail else 64                                 # neon frame
    glow(img, lambda d: d.rounded_rectangle((pad + w / 2, pad + w / 2, S - pad - w / 2, S - pad - w / 2), r - w / 2,
                                            outline=CYAN + (255,), width=w), 28 if detail else 18)

    f = font(700 if detail else 820)                         # the B: magenta shadow + cyan front, like the title
    bbox = ImageDraw.Draw(img).textbbox((0, 0), "B", font=f)
    x = (S - (bbox[2] - bbox[0])) / 2 - bbox[0]
    y = (S - (bbox[3] - bbox[1])) / 2 - bbox[1] - (40 if detail else 10)
    off = 30 if detail else 38
    glow(img, lambda d: d.text((x - off, y - off * 0.6), "B", font=f, fill=MAG + (255,)), 20, 1)
    glow(img, lambda d: d.text((x, y), "B", font=f, fill=CYAN + (255,)), 22, 1)

    if detail:                                               # a ledger line flowing under the letter
        yy = S - pad - 150
        glow(img, lambda d: d.line((S * 0.24, yy, S * 0.76, yy), fill=GREEN + (255,), width=26), 16)
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
