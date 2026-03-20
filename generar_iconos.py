import os
from PIL import Image, ImageDraw, ImageFont

ICONS = {
    "excel": "E",
    "pdf": "P", 
    "folder": "F",
    "settings": "C",
    "start": "▶",
    "target": "T",
    "save": "S",
    "load": "L",
    "close": "X",
    "prev": "◀",
    "next": "▶"
}

os.makedirs("assets/icons", exist_ok=True)

# Generate simple 64x64 PNGs with text centered
for name, char in ICONS.items():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # White text so it looks good on dark UI
    # In a real app we'd download nice vectors, but this produces valid PNGs quickly.
    d.text((20, 16), char, fill="white", font_size=24)
    img.save(f"assets/icons/{name}.png", "PNG")

print("Iconos generados.")
