import os
import urllib.request

ICONS = {
    "check": "https://img.icons8.com/color/100/checked--v1.png",
    "alert": "https://img.icons8.com/color/100/high-importance--v1.png", 
    "clock": "https://img.icons8.com/ios-filled/100/ffffff/time.png"
}

os.makedirs("assets/icons", exist_ok=True)

for name, url in ICONS.items():
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            with open(f"assets/icons/{name}.png", "wb") as f:
                f.write(response.read())
        print(f" - {name}.png descargado.")
    except Exception as e:
        print(f"Error descargando {name}: {e}")
