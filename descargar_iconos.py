import os
import urllib.request

# We will use a reliable service (like unpkg or raw github) or a simple fallback CDN
# for free standard icons like icons8 or material design if direct PNG isn't trivial.
# Since we need PNGs (CTkImage prefers them over raw SVG without extra libs), 
# let's download real icons from a generic icon API.

ICONS = {
    # Using tabler-icons or similar format via a CDN that renders PNGs
    "excel": "https://img.icons8.com/color/64/microsoft-excel-2019--v1.png",
    "pdf": "https://img.icons8.com/color/64/pdf.png", 
    "folder": "https://img.icons8.com/color/64/folder-invoices--v1.png",
    "settings": "https://img.icons8.com/ios-filled/64/ffffff/settings.png",
    "start": "https://img.icons8.com/color/64/circled-play.png",
    "target": "https://img.icons8.com/ios-filled/64/ffffff/crosshair.png",
    "save": "https://img.icons8.com/ios-filled/64/ffffff/save--v1.png",
    "load": "https://img.icons8.com/ios-filled/64/ffffff/download--v1.png",
    "close": "https://img.icons8.com/color/64/cancel--v1.png",
    "prev": "https://img.icons8.com/ios-filled/64/ffffff/circled-left-2.png",
    "next": "https://img.icons8.com/ios-filled/64/ffffff/circled-right-2.png"
}

os.makedirs("assets/icons", exist_ok=True)

print("Descargando iconos reales de alta calidad...")
for name, url in ICONS.items():
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            with open(f"assets/icons/{name}.png", "wb") as f:
                f.write(response.read())
        print(f" - {name}.png descargado.")
    except Exception as e:
        print(f"Error descargando {name}: {e}")

print("¡Iconos reales listos!")
