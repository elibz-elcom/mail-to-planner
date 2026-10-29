"""Generate the add-in icons (teal tile with a checklist + arrow glyph)."""
from pathlib import Path
from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent / "assets"
OUT.mkdir(parents=True, exist_ok=True)
SIZES = [16, 25, 32, 48, 50, 64, 75, 80, 96, 128, 144]
TEAL = (15, 118, 110, 255)
WHITE = (255, 255, 255, 255)


def draw(size: int) -> Image.Image:
    s = 512
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([16, 16, s - 16, s - 16], radius=96, fill=TEAL)
    # three checklist rows
    for i, y in enumerate((150, 256, 362)):
        d.rounded_rectangle([110, y - 34, 178, y + 34], radius=14, outline=WHITE, width=18)
        if i == 0:
            d.line([(124, y), (142, y + 20), (172, y - 22)], fill=WHITE, width=20, joint="curve")
        d.rounded_rectangle([212, y - 12, 400, y + 12], radius=12, fill=WHITE)
    return img.resize((size, size), Image.LANCZOS)


for sz in SIZES:
    draw(sz).save(OUT / f"icon-{sz}.png")
print("icons:", ", ".join(str(s) for s in SIZES))
