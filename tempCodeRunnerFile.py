def _extraer(ruta_archivo, cols):
    datos = {}
    try:
        with pdfplumber.open(ruta_archivo) as pdf:
            texto = "\n".join(p.extract_text() for p in pdf.pages if p.extract_text())
        for col in cols:
            cl = col.lower()
            if "manifiesto" in cl or "remesa" in cl:
                    # CAMBIO CLAVE: Quitamos el \b del inicio y buscamos el patrón greedy (glotón)
                    # Esto atrapa "128252-06" incluso si el PDF lee "texto128252-06"
                    match_manifiestos = re.findall(r'(\d{5,6})\s*-\s*(\d{2})\b', texto)
                    
                    if match_manifiestos:
                        # re.findall con grupos devuelve una lista de tuplas [('128252', '06')]
                        # Tomamos la última encontrada (-1) y unimos las partes
                        ultimo_match = match_manifiestos[-1]
                        datos[col] = f"{ultimo_match[0]}-{ultimo_match[1]}"
                    else:
                        datos[col] = "No encontrado"
            elif "placa" in cl or "vehículo" in cl:
                m = re.search(r'PLACA DEL VEH[ÍI]CULO[\s\S]*?\b([A-Z]{3}[-\s]?\d{3})\b', texto, re.I)
                datos[col] = m.group(1).upper().replace(" ","").replace("-","") if m else "No encontrado"
            elif "fecha" in cl:
                fs = re.findall(r'\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4}', texto)
                datos[col] = next((f for f in fs if f != "30/08/2023"), "No encontrado")
            elif "cantidad" in cl or "volumen" in cl:
                m = re.search(r'SUBTOTAL[^\d]*(\d+[.,]?\d*)', texto, re.I)
                datos[col] = m.group(1) if m else "No encontrado"
            else:
                if any(x in cl for x in ["cédula","cedula","nit","id"]):
                    m = re.search(rf'(?i)({col})[^\d]*(\d{{7,15}})', texto)
                    datos[col] = m.group(2) if m else "No encontrado"
                else:
                    m = re.search(rf'(?i){re.escape(col)}[:\s]*([A-Za-z0-9., áéíóúÁÉÍÓÚñÑ]+)', texto)
                    datos[col] = m.group(1).strip() if m else "No encontrado"
    except Exception as e:
        log(f"Error PDF {os.path.basename(ruta_archivo)}: {e}", "error")
    return datos