import pdfplumber

ruta_pdf = r"D:\COMPARTIDOS\10_Oct-25\1_ARD\Jacana Sur\2025.10.12_ARD_Jacana Sur_Jacana 63_Atina 1001_NUX770-.pdf" # <-- CAMBIA ESTO

print("=== DIAGNÓSTICO DE TABLA (MANIFIESTO) ===")
try:
    with pdfplumber.open(ruta_pdf) as pdf:
        # Solo leemos la página 1 (índice 0)
        tabla = pdf.pages[0].extract_table()
        
        if tabla:
            for i, fila in enumerate(tabla[:15]): # Imprime las primeras 15 filas
                print(f"Fila {i}: {fila}")
        else:
            print("La página 1 no tiene una estructura de tabla reconocible.")
except Exception as e:
    print(f"Error leyendo el PDF: {e}")