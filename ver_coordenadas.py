import pdfplumber
from PIL import Image, ImageDraw

# CAMBIA ESTO POR TU PDF
archivo = "D:\\COMPARTIDOS\\10_Oct-25\\1_ARD\\junto\\2025.10.01_AA_Jacana_LJS753-.pdf" 

with pdfplumber.open(archivo) as pdf:
    print(f"El PDF tiene {len(pdf.pages)} páginas.")
    
    for i, pagina in enumerate(pdf.pages):
        print(f"\n--- Analizando PÁGINA {i+1} ---")
        print(f"Dimensiones -> Ancho: {pagina.width}, Alto: {pagina.height}")
        
        # Generar imagen con las cajitas rojas
        im = pagina.to_image(resolution=150) # Aumenté resolución para ver mejor
        im.draw_rects(pagina.extract_words(), stroke="red", stroke_width=2)
        
        print(f"Mostrando página {i+1}...")
        im.show()
        
        # Pausa para que puedas verla antes de pasar a la siguiente
        if i < len(pdf.pages) - 1:
            input("Presiona ENTER en la consola para ver la siguiente página...")