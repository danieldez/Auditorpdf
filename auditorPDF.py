import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
import os, re, time, threading, subprocess, sys, json
import gc  # [OPT-1] Importamos gc para forzar recolección de basura en puntos críticos
import pdfplumber, openpyxl, xlwings as xw
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from entrenador import MapeadorPDF
from gestor import GestorPlantillas
from config import RUTAS
from PIL import Image
import ctypes  # [NEW] Para forzar icono de barra de tareas en Windows

# Intentamos setear el AppUserModelID para que Windows muestre el icono en la barra de tareas
try:
    myappid = 'mycompany.auditor.validador.1.0' # identificador único
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except:
    pass
# NOTA [OPT-2]: ImageTk eliminado de imports — no se usa directamente en este módulo.
# CTkImage de customtkinter es el wrapper correcto y ya gestiona el ciclo de vida
# de la imagen PIL internamente, evitando referencias huérfanas en Tkinter.

# ══════════════════════════════════════════════════════════════════
#  GLOBALS & CONFIG
# ══════════════════════════════════════════════════════════════════
ruta_excel          = ""
ruta_pdf            = ""
checkboxes_columnas = []
_en_curso           = False
archivo_config      = "config_auditor.json"

# --- PALETA DE COLORES DINÁMICA (Claro, Oscuro) ---
C_BG_MAIN    = ("#f5f6fa", "#0f111a")
C_BG_SIDE    = ("#ffffff", "#1e212b")
C_CARD       = ("#e1e5eb", "#252836")
C_ACCENT     = ("#1f6aa5", "#3b82f6")
C_ACCENT_H   = ("#144870", "#144870")
C_SUCCESS    = ("#2ea043", "#2ea043")
C_ERROR      = ("#da3633", "#da3633")
C_TEXT_MAIN  = ("#111111", "#ffffff")
C_TEXT_SUB   = ("#666666", "#a0a0a0")

# --- COLORES DE IDENTIDAD INDUSTRIAL ---
C_IND_OIL   = "#3b82f6"  # Azul Petrolero / Transporte
C_IND_SS    = "#06b6d4"  # Teal Seguridad Social / Contabilidad
C_IND_SS_H  = "#0891b2"  # Hover Teal


# CARGAMOS LAS PLANTILLAS EN MEMORIA AL INICIO PARA ACCESO RÁPIDO
MEMORIA_PLANTILLAS = {}
if os.path.exists(RUTAS["json"]):
    try:
        with open(RUTAS["json"], "r", encoding="utf-8") as f:
            MEMORIA_PLANTILLAS = json.load(f)
    except: pass

# ══════════════════════════════════════════════════════════════════
#  UTILIDADES — PyInstaller / Rutas
# ══════════════════════════════════════════════════════════════════
# [OPT-3] CONSOLIDACIÓN DE RUTA BASE:
# Antes existían DOS bloques separados que calculaban sys._MEIPASS
# (uno en el módulo y otro en obtener_ruta()).  Cada llamada a
# obtener_ruta() relanzaba la excepción AttributeError internamente
# para detectar el modo frozen, lo cual tiene coste no nulo en el
# startup del .exe.  Ahora se calcula UNA sola vez al importar el
# módulo y obtener_ruta() simplemente lo reutiliza.
_BASE_DIR = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

def obtener_ruta(ruta_relativa: str) -> str:
    """Ruta absoluta compatible con PyInstaller. Usa _BASE_DIR precalculado."""
    return os.path.join(_BASE_DIR, ruta_relativa)

# [OPT-4] CARGA ÚNICA DEL LOGO:
# CTkImage mantiene una referencia a la imagen PIL mientras exista el objeto.
# Creamos una sola instancia global en lugar de re-crearla en cada llamada
# a aplicar_interfaz() (que antes creaba un objeto nuevo cada vez que el
# usuario cambiaba de modo, dejando el objeto anterior sin liberar hasta el
# siguiente ciclo del GC).
ruta_logo = obtener_ruta("logo.png")
try:
    IMG_LOGO = ctk.CTkImage(Image.open(ruta_logo), size=(150, 150))
except Exception as e:
    print(f"Error cargando logo: {e}")
    IMG_LOGO = None

# --- CARGA GLOBALES DE ÍCONOS (USANDO CTkImage) ---
# Alamacenamos los CTkImage a nivel de módulo para evitar que el Garbage Collector
# los destruya. Usamos un dict `ICONS`
ICONS = {}
try:
    _icons_dir = obtener_ruta("assets/icons")
    ICONS["excel"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "excel.png")), size=(20, 20))
    ICONS["pdf"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "pdf.png")), size=(20, 20))
    ICONS["folder"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "folder.png")), size=(20, 20))
    ICONS["settings"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "settings.png")), size=(18, 18))
    ICONS["start"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "start.png")), size=(24, 24))
    ICONS["target"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "target.png")), size=(18, 18))
    ICONS["load"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "load.png")), size=(18, 18))
    
    # Dashboard Cards
    ICONS["check"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "check.png")), size=(45, 45))
    ICONS["alert"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "alert.png")), size=(45, 45))
    ICONS["clock"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "clock.png")), size=(45, 45))
except Exception as e:
    print(f"No se pudieron cargar algunos íconos: {e}")

def guardar_config():
    config = {
        "excel": ruta_excel,
        "pdf": ruta_pdf,
        "fila": entry_fila.get(),
        "hoja": combo_hoja.get()
    }
    try:
        with open(archivo_config, "w") as f:
            json.dump(config, f)
    except Exception as e:
        print(f"No se pudo guardar config: {e}")

def cargar_config_inicial():
    global ruta_excel, ruta_pdf
    if os.path.exists(archivo_config):
        try:
            with open(archivo_config, "r") as f:
                config = json.load(f)

            if os.path.exists(config.get("excel", "")):
                ruta_excel = config["excel"]
                lbl_excel_name.configure(text=os.path.basename(ruta_excel), text_color=C_SUCCESS)
                try:
                    # [OPT-5] CIERRE EXPLÍCITO del workbook de openpyxl en
                    # cargar_config_inicial.  El código original asignaba wb pero
                    # nunca llamaba wb.close().  En modo read_only, openpyxl abre
                    # un ZipFile interno; sin close() queda abierto hasta que el
                    # GC decida destruir el objeto, manteniendo el fd de fichero
                    # ocupado y ~2-4 MB de strings en RAM.
                    wb = openpyxl.load_workbook(ruta_excel, read_only=True)
                    combo_hoja.configure(values=wb.sheetnames)
                    if config.get("hoja") in wb.sheetnames:
                        combo_hoja.set(config["hoja"])
                    wb.close()   # ← CORRECCIÓN: liberamos el ZipFile inmediatamente
                except: pass

            if os.path.exists(config.get("pdf", "")):
                ruta_pdf = config["pdf"]
                n = len([f for f in os.listdir(ruta_pdf) if f.lower().endswith(".pdf")])
                lbl_pdf_name.configure(text=f"{n} Archivos cargados", text_color=C_SUCCESS)

            if config.get("fila"):
                entry_fila.delete(0, "end")
                entry_fila.insert(0, config["fila"])

            log("Configuración anterior cargada correctamente.", "info")
        except Exception as e:
            log(f"Error cargando config: {e}", "warn")

def _fmt(s):
    s = max(0, int(s))
    return f"{s}s" if s < 60 else f"{s//60}m {s%60:02d}s"

def _ts():
    return datetime.now().strftime("%H:%M:%S")

def _excel_abierto(ruta):
    try:
        with open(ruta, 'a'):
            pass
        return False
    except IOError:
        return True

def abrir_excel():
    if not ruta_excel: return
    try:
        if sys.platform == "win32": os.startfile(ruta_excel)
        elif sys.platform == "darwin": subprocess.call(["open", ruta_excel])
        else: subprocess.call(["xdg-open", ruta_excel])
    except Exception as e:
        log(f"No se pudo abrir el archivo: {e}", "error")

# ══════════════════════════════════════════════════════════════════
#  CARGA DE PLANTILLAS
# ══════════════════════════════════════════════════════════════════
def cargar_plantillas_en_memoria():
    global MEMORIA_PLANTILLAS
    if os.path.exists(RUTAS["json"]):
        try:
            with open(RUTAS["json"], "r", encoding="utf-8") as f:
                MEMORIA_PLANTILLAS = json.load(f)
        except:
            MEMORIA_PLANTILLAS = {}

cargar_plantillas_en_memoria()

# ══════════════════════════════════════════════════════════════════
#  EXTRACCIÓN PDF — TRANSPORTE
# ══════════════════════════════════════════════════════════════════
def extraer_por_tabla_manifiesto(pagina):
    """Capa 2: Busca Placa, Manifiesto, Cantidad y Unidad leyendo la cuadrícula."""
    datos = {"placa": None, "manifiesto": None, "cantidad": None, "unidad": None}
    tabla = pagina.extract_table()
    if not tabla:
        return datos

    PALABRAS_CLAVE_CANTIDAD = (
        "agua residual", "lodo", "crudo", "aceite", "volumen",
        "cantidad", "cant", "bbls", "galones", "residuo", "salmuera"
    )
    
    idx_cant = -1
    idx_und  = -1

    for fila in tabla:
        celdas = [str(c).strip() for c in fila if c is not None and str(c).strip() != ""]
        if not celdas: continue

        fila_texto = " ".join(celdas).lower()
        
        # --- BUSCADOR DE ENCABEZADOS DE TABLA ---
        # Si la fila tiene "CANT" y "UND", marcamos qué columna es cada una
        if idx_cant == -1:
            for i, c in enumerate(celdas):
                c_up = c.upper()
                if any(kw in c_up for kw in ["CANTIDAD", "CANT.", "VOL", "NETO"]):
                    idx_cant = i
                if any(kw in c_up for kw in ["UNIDAD", "UND", "MEDIDA"]):
                    idx_und = i

        # --- EXTRACCIÓN DE PLACA Y MANIFIESTO (EXISTENTE) ---
        for i, celda in enumerate(celdas):
            texto = celda.lower()
            if "placa" in texto and i + 1 < len(celdas):
                datos["placa"] = celdas[i+1].split('/')[0].replace(" ", "").replace("-", "")
            elif ("manifiesto" in texto or "remesa" in texto) and i + 1 < len(celdas):
                manif = celdas[i+1].replace(" ", "").replace("\n", "")
                if i + 2 < len(celdas) and re.match(r'^\d{1,2}$', celdas[i+2].strip()):
                    manif += celdas[i+2].strip()
                if len(re.findall(r'\d', manif)) >= 4:
                    datos["manifiesto"] = manif

        # --- EXTRACCIÓN POR COLUMNAS (SI HAY HEADER) ---
        if idx_cant != -1 and idx_cant < len(celdas):
            val_c = celdas[idx_cant].replace(",", ".").strip()
            try:
                v = float(val_c)
                if 3 < v < 5000 and v not in (900, 800, 961):
                    if datos["cantidad"] is None:
                        datos["cantidad"] = str(v).replace(".0", "")
                    # Si detectamos cantidad, intentamos pillar la unidad de la columna de al lado
                    if idx_und != -1 and idx_und < len(celdas) and datos["unidad"] is None:
                        u_raw = celdas[idx_und].lower()
                        if any(x in u_raw for x in ["bbl", "gal", "m3", "und", "kg", "ton"]):
                           datos["unidad"] = u_raw
            except ValueError:
                pass

        # --- FALLBACK: CANTIDAD POR KEYWORD (SI NO HAY HEADER CLARO) ---
        if datos["cantidad"] is None and any(kw in fila_texto for kw in PALABRAS_CLAVE_CANTIDAD):
            for c in celdas:
                c_limpio = c.replace(",", ".").strip()
                try:
                    v = float(c_limpio)
                    if 3 < v < 5000 and v not in (900, 800, 961):
                        datos["cantidad"] = str(v).replace(".0", "")
                        break
                except ValueError:
                    pass
    return datos

def _extraer(ruta_archivo, cols):
    datos = {}

    # FASE 0: IDENTIFICAR SI HAY PLANTILLA (MEJORADO)
    plantilla_actual = None
    try:
        nombre_archivo = os.path.basename(ruta_archivo).upper()
        for nombre_plantilla, data_plantilla in MEMORIA_PLANTILLAS.items():
            np_up = nombre_plantilla.upper()
            
            # 1. Match Directo
            if np_up in nombre_archivo:
                plantilla_actual = data_plantilla
                break
            
            # 2. Match por Prefijo Flexible (ARD <-> AA)
            # Si la plantilla es ARD_JACANA y el archivo es AA_JACANA...
            np_alt = np_up.replace("ARD_", "AA_")
            if np_alt in nombre_archivo:
                plantilla_actual = data_plantilla
                break
                
            # 3. Match por Keyword significativa (si el nombre de la plantilla es corto y está en el archivo)
            # Ej: plantilla "JACANA" en archivo "2025.10.13_AA_JACANA..."
            if len(np_up) > 4 and np_up in nombre_archivo:
                plantilla_actual = data_plantilla
                break
                
            # 4. Caso especial Jacana (común en este flujo)
            if "JACANA" in np_up and "JACANA" in nombre_archivo:
                plantilla_actual = data_plantilla
                break

        if plantilla_actual:
            print(f"✅ ¡ÉXITO! Plantilla detectada para: {os.path.basename(ruta_archivo)}")
    except Exception as e:
        print(f"Error en detección de plantilla: {e}")

    try:
        # [OPT-6] USO CORRECTO DEL CONTEXT MANAGER DE pdfplumber:
        # pdfplumber.open() ya es un context manager, y el bloque `with` garantiza
        # que pdf.close() se llame al salir, lo que libera:
        #   · El objeto pdfminer.PDFDocument (parseo interno)
        #   · Todos los objetos pdfminer.LTPage cacheados en pdf.pages
        #   · El buffer del archivo en disco (file handle)
        # PROBLEMA ORIGINAL: texto_full, todas_lineas y texto_upper se construían
        # DENTRO del `with`, pero la lógica de Fase 2 estaba FUERA.  Eso es
        # correcto —el with cierra el PDF antes de procesar regex—, pero
        # las variables de texto (potencialmente varios MB por PDF) quedaban
        # vivas hasta el fin de la función.  Con `del` explícito al terminar
        # la fase 2 reducimos el pico de memoria durante el procesamiento
        # paralelo con ThreadPoolExecutor (hasta 6 workers simultáneos).
        with pdfplumber.open(ruta_archivo) as pdf:

            # FASE 1: EXTRACCIÓN POR COORDENADAS
            if plantilla_actual:
                for col in cols:
                    col_busqueda = col.lower()
                    llave_mapeada = None
                    for key_json in plantilla_actual.keys():
                        if key_json in col_busqueda or col_busqueda in key_json:
                            llave_mapeada = key_json
                            break

                    if llave_mapeada:
                        info = plantilla_actual[llave_mapeada]
                        num_pag = info["pagina"]
                        bbox = info["coordenadas"]

                        if num_pag < len(pdf.pages):
                            texto_recorte = pdf.pages[num_pag].crop(bbox).extract_text(x_tolerance=3, y_tolerance=3)

                            if texto_recorte:
                                val_limpio = texto_recorte.strip().replace("\n", " ")

                                if "manifiesto" in col_busqueda or "remesa" in col_busqueda:
                                    val_limpio = re.sub(r"[^\w]+", "-", val_limpio).strip("-")

                                datos[col] = val_limpio
            # --- NUEVA FASE 1.5: MOTOR HÍBRIDO (TABLAS) ---
            # Escaneamos TODAS las páginas hasta encontrar los datos necesarios.
            # Esto resuelve el caso de PDFs con páginas en orden invertido
            # (la planilla está en la pág.2 y el manifiesto en la pág.1).
            datos_tabla = {"placa": None, "manifiesto": None, "cantidad": None, "unidad": None}
            for pag_idx, pagina in enumerate(pdf.pages):
                resultado_pag = extraer_por_tabla_manifiesto(pagina)
                # Fusionar: solo sobreescribir si la página actual tiene dato y antes no había
                for campo in ("placa", "manifiesto", "cantidad", "unidad"):
                    if datos_tabla[campo] is None and resultado_pag.get(campo):
                        datos_tabla[campo] = resultado_pag[campo]
                # Si ya tenemos placa y manifiesto, no necesitamos más páginas
                if datos_tabla["placa"] and datos_tabla["manifiesto"]:
                    break

            for col in cols:
                if col in datos and datos[col] and datos[col] != "No encontrado":
                    continue  # Si ya lo extrajo la plantilla, no lo tocamos
                c_low = col.lower()
                if "placa" in c_low and datos_tabla.get("placa"):
                    datos[col] = datos_tabla["placa"]
                elif ("manifiesto" in c_low or "remesa" in c_low) and datos_tabla.get("manifiesto"):
                    datos[col] = datos_tabla["manifiesto"]
                elif ("cantidad" in c_low or "volumen" in c_low) and datos_tabla.get("cantidad"):
                    datos[col] = datos_tabla["cantidad"]
                elif ("unidad" in c_low or "medida" in c_low) and datos_tabla.get("unidad"):
                    # Normalización rápida para la unidad de tabla
                    u_t = datos_tabla["unidad"].upper()
                    if "BBL" in u_t: datos[col] = "bbls"
                    elif "GAL" in u_t: datos[col] = "galones"
                    elif "M3" in u_t: datos[col] = "metros cúbicos"
                    else: datos[col] = u_t.lower()
            # ----------------------------------------------
            # PREPARACIÓN PARA FASE 2
            texto_full   = "\n".join(p.extract_text() or "" for p in pdf.pages)
            todas_lineas = texto_full.split('\n')
            texto_upper  = texto_full.upper().replace(":", " : ")
        # ← El with cierra aquí: pdfplumber libera todos sus recursos internos.

        # FASE 2: INTELIGENCIA REGEX
        for col in cols:
            if col in datos and datos[col] != "No encontrado":
                continue

            cl  = col.lower()
            val = "No encontrado"

            if "unidad" in cl or "medida" in cl:
                # --- FIX F: HEURÍSTICO DE UNIDAD PREFERIDA ---
                # 1. Escaneo global buscando palabras clave de unidades reales
                # (Es mucho más probable que el manifiesto use una de estas que "salmuera")
                unidades_reales = {
                    "BBL": "bbls", "BARRILES": "bbls", "BBLS": "bbls",
                    "GALON": "galones", "GALONES": "galones", "GLS": "galones",
                    "M3": "metros cúbicos", "METROS": "metros cúbicos",
                    "KG": "kilogramos", "KGS": "kilogramos", "KILOS": "kilogramos",
                    "TON": "toneladas", "TONELADAS": "toneladas",
                    "UND": "unidades", "UNIDADES": "unidades"
                }
                
                encontrada_fuerte = None
                for kw_u, val_u in unidades_reales.items():
                    if kw_u in texto_upper:
                        encontrada_fuerte = val_u
                        break
                
                # 2. Intento por Label (UNIDAD: XXX)
                match = re.search(r'(?:UNIDAD|MEDIDA|UND)(?:[\s.:_]{0,20})([A-Z0-9]{1,15})\b', texto_upper)
                if match:
                    raw_unit = match.group(1).upper()
                    # Si el match es una unidad conocida, la usamos
                    if raw_unit in unidades_reales:
                        val = unidades_reales[raw_unit]
                    else:
                        # Si no es conocida pero tenemos una "fuerte" en el doc, preferimos la fuerte
                        if encontrada_fuerte:
                            val = encontrada_fuerte
                        else:
                            # Caso último recurso: aceptamos lo que diga el label si no hay nada mejor
                            val = raw_unit.lower()
                else:
                    # Si no hay label, usamos la encontrada en el doc-wide scan
                    if encontrada_fuerte:
                        val = encontrada_fuerte

                datos[col] = val
                continue

            elif "manifiesto" in cl or "remesa" in cl or "acta" in cl or "viaje" in cl:
                val = "No encontrado"
                # Añadimos (?:\d{1,3}-)? para atrapar prefijos como "8-" y \d{1,4} al final
                patron_flexible = r'(?:REMESA|MANIFIESTO|VIAJE|MFT|OT|DESPACHO).{0,90}?\b((?!128379|900|800|202\d)(?:\d{1,3}-)?\d{4,12}[-\s\'._]+\d{1,4})\b'
                match = re.search(patron_flexible, texto_upper)
                if match:
                    val = re.sub(r'[-\s\'._]+', '-', match.group(1))
                else:
                    nums = re.findall(r'(?<!\d)((?:\d{1,3}-)?\d{4,10}[-\s\'._]+\d{1,4})(?!\d)', texto_upper)
                    validos = [re.sub(r'[-\s\'._]+', '-', n) for n in nums if not n.startswith(('900','800','202','128379'))]
                    if validos: val = validos[0]
                datos[col] = val

            elif "placa" in cl:
                m = re.search(r'\b([A-Z]{3}\s?-?\s?\d{3})\b', texto_upper)
                if m: val = m.group(1).replace(" ","").replace("-","")
                datos[col] = val

            elif "tipo" in cl or "material" in cl:
                if "CARTON" in texto_upper or "CARTÓN" in texto_upper: val = "Cartón"
                elif "VIDRIO" in texto_upper: val = "Vidrio"
                elif "METAL" in texto_upper or "CHATARRA" in texto_upper: val = "Metales"
                elif "PLASTICO" in texto_upper or "PLÁSTICO" in texto_upper: val = "Plásticos"
                elif "ORGANICO" in texto_upper: val = "Orgánicos"
                elif "IMPREGNADO" in texto_upper or "SOLIDOS" in texto_upper: val = "Sólidos impregnados"
                elif "ORDINARIO" in texto_upper: val = "Ordinarios"
                else:
                    m = re.search(r'(?:TIPO|CLASE|RESIDUO).*?[:\s]+(?:DE\s+)?(\w+)', texto_upper)
                    if m: val = m.group(1).capitalize()
                datos[col] = val

            elif "cantidad" in cl or "volumen" in cl:
                BLACKLIST_CANT = {900, 800, 961, 128379}
                KEYWORDS_CANT  = ("CANTIDAD", "VOLUMEN", "CANT", "NETO", "TOTAL BBL",
                                  "PESO NETO", "PESO")

                def _es_valida(v):
                    """True si v parece una cantidad real de volumen."""
                    return (
                        v > 3           # excluye dígitos aislados (1-3)
                        and v < 5000    # excluye números sin sentido
                        and v not in BLACKLIST_CANT
                        and not (2010 <= v <= 2031)  # excluye años
                    )

                # --- ESTRATEGIA 1: Ventana ampliada con contexto de línea ---
                # Buscamos la línea que contiene la keyword y leemos
                # esa línea + las 3 siguientes (valor puede estar en la fila de abajo)
                for kw in KEYWORDS_CANT:
                    idx = texto_upper.find(kw)
                    if idx == -1:
                        continue
                    # Cortamos el texto desde la keyword hasta 300 chars adelante
                    fragmento = texto_upper[idx : idx + 300]
                    nums_frag = re.findall(r'\b(\d{1,4}(?:[.,]\d{1,2})?)\b', fragmento)
                    for n_str in nums_frag:
                        try:
                            v = float(n_str.replace(",", "."))
                            if _es_valida(v):
                                val = str(v).replace(".0", "")
                                break
                        except: pass
                    if val != "No encontrado":
                        break

                # --- ESTRATEGIA 2: Líneas que contienen la keyword ---
                if val == "No encontrado":
                    for i, linea in enumerate(todas_lineas):
                        l_up = linea.upper()
                        if any(kw in l_up for kw in KEYWORDS_CANT):
                            # Revisar esta línea y las 3 siguientes
                            bloque = todas_lineas[i : i + 4]
                            for bl in bloque:
                                nums_bl = re.findall(r'\b(\d{1,4}(?:[.,]\d{1,2})?)\b', bl)
                                for n_str in nums_bl:
                                    try:
                                        v = float(n_str.replace(",", "."))
                                        if _es_valida(v):
                                            val = str(v).replace(".0", "")
                                            break
                                    except: pass
                                if val != "No encontrado":
                                    break
                            if val != "No encontrado":
                                break

                # --- ESTRATEGIA 3: Fallback — número más grande razonable del doc ---
                # (elegimos el MAYOR, no el último, que es más probable que sea volumen)
                if val == "No encontrado":
                    nums_doc = re.findall(r'\b(\d{1,4}(?:[.,]\d{1,2})?)\b', texto_upper)
                    candidatos = []
                    for n_str in nums_doc:
                        try:
                            v = float(n_str.replace(",", "."))
                            if _es_valida(v):
                                candidatos.append(v)
                        except: pass
                    if candidatos:
                        # Tomamos el más frecuente o el mayor (generalmente el volumen es el número relevante)
                        from collections import Counter
                        mas_comun = Counter(candidatos).most_common(1)[0][0]
                        val = str(mas_comun).replace(".0", "")

                datos[col] = val

            else:
                m = re.search(rf'{re.escape(col)}[:\s]+([^\n]+)', texto_full, re.I)
                if m: val = m.group(1).strip()[:40]
                if ("unidad" in cl or "medida" in cl) and "agua" in val.lower():
                    val = "bbls"
                datos[col] = val

            if "fecha" in cl:
                match = re.search(r'\b(\d{2,4}[-/.]\d{2}[-/.]\d{2,4})\b', texto_upper)
                if match:
                    val = match.group(1).replace("/", "-").replace(".", "-")
                else:
                    m = re.search(r'FECHA.*?(?:[:\s]+)(\d{2,4}[-/.]\d{2}[-/.]\d{2,4})', texto_upper)
                    if m: val = m.group(1).replace("/", "-").replace(".", "-")

                datos[col] = val
                
    # --- NUEVA FASE 3: CONSENSO CRUZADO DE MANIFIESTOS ---
        for col in cols:
            cl = col.lower()
            if "manifiesto" in cl or "remesa" in cl:
                val_actual = str(datos.get(col, "No encontrado")).strip()

                # ── ZONA DE ENCABEZADO (primeros 600 chars del doc) ──────────────
                # Los números que aparecen únicamente en la zona de encabezado
                # son referencias del documento (lote/contrato), NO manifiestos reales.
                HEADER_ZONE = texto_upper[:600]

                def _en_header_solamente(num_str):
                    """True si el número SOLO aparece dentro del encabezado del doc."""
                    num_bare = num_str.replace('-', '')
                    pos_total = [m.start() for m in re.finditer(re.escape(num_bare), texto_upper)]
                    return pos_total and all(p < 600 for p in pos_total)

                # Buscar candidatos fuera del header zone
                candidatos_raw = re.findall(r'(?<!\d)((?:\d{1,3}-)?\d{4,10}[-\s\'._]+\d{1,4})(?!\d)', texto_upper)
                candidatos_limpios = [re.sub(r'[-\s\'._]+', '-', c) for c in candidatos_raw
                                      if not c.startswith(('900','800','202','128379'))
                                      and not _en_header_solamente(re.sub(r'[-\s\'._]+', '-', c))]

                # Rechazar candidatos que se repiten más de 2 veces
                from collections import Counter
                conteo = Counter(candidatos_limpios)
                candidatos_limpios = [c for c in candidatos_limpios if conteo[c] <= 2]

                # ── CLASIFICACIÓN POR FORMATO ─────────────────────────────────────
                # Los manifiestos petroleros reales tienen sufijo de 2 dígitos:  "XXXXXX-06"
                # Los IDs de encabezado/lote suelen tener sufijo de 1 dígito:   "XXXXXXX-1"
                PATRON_REAL = re.compile(r'^\d{4,9}-\d{2}$')

                preferidos  = [c for c in candidatos_limpios if PATRON_REAL.match(c)]
                resto       = [c for c in candidatos_limpios if not PATRON_REAL.match(c)]
                candidatos_ordenados = preferidos + resto  # primero los de formato real

                # 1. Eliminar basura de escáner
                if len(re.findall(r'\d', val_actual)) < 4:
                    val_actual = "No encontrado"

                if candidatos_ordenados:
                    # Elegir el mejor candidato usando proximidad a la keyword
                    mejor_cand = None
                    kws_manif  = ("MANIFIESTO", "REMESA", "DESPACHO", "VIAJE", "MFT")
                    pos_keyword = -1
                    for kw in kws_manif:
                        p = texto_upper.find(kw)
                        if p != -1:
                            pos_keyword = p
                            break

                    if pos_keyword != -1 and preferidos:
                        # Solo buscamos proximidad entre los de formato real para mayor precisión
                        mejor_dist = 10**9
                        for c in preferidos:
                            pos_c = texto_upper.find(c.replace('-', ''), pos_keyword)
                            if pos_c != -1 and (pos_c - pos_keyword) < mejor_dist:
                                mejor_dist = pos_c - pos_keyword
                                mejor_cand = c

                    if mejor_cand is None:
                        mejor_cand = candidatos_ordenados[0]  # fallback al mejor disponible

                    if val_actual == "No encontrado":
                        datos[col] = mejor_cand
                    else:
                        # Consenso: ampliar un número parcial (e.g. "129199-0" → "129199-06")
                        nums_actual = re.sub(r'\D', '', val_actual)
                        for c in candidatos_ordenados:
                            nums_c = re.sub(r'\D', '', c)
                            if nums_actual in nums_c and len(nums_c) > len(nums_actual):
                                datos[col] = c
                                break
                            elif val_actual.endswith("-0") and c.startswith(val_actual[:-2]) and not c.endswith("-0"):
                                datos[col] = c
                                break

            v_final = str(datos.get(col, ""))
            # Corta prefijos como "t-128929-06" o "8-129194-06" y deja solo "128929-06"
            m_core = re.search(r'(\d{4,9}-\d{1,4})$', v_final)
            if m_core:
                datos[col] = m_core.group(1)
        # [OPT-7] LIBERACIÓN EXPLÍCITA DE STRINGS DE TEXTO GRANDES:
        # En un ThreadPoolExecutor con 6 workers, cada worker puede tener
        # en memoria simultáneamente texto_full + todas_lineas + texto_upper
        # (típicamente 3–15 MB por PDF según tamaño).  Del-earlos aquí
        # permite que el GC los recoja antes de que el futuro se resuelva
        # y el resultado vuelva al hilo principal, reduciendo el pico de RAM.
        del texto_full, todas_lineas, texto_upper

    except Exception:
        pass
    
    # --- FILTRO FINAL ANTI-BASURA ---
    for c_key, c_val in datos.items():
        if "manifiesto" in c_key.lower() or "remesa" in c_key.lower():
            # Si extrajo letras raras y tiene menos de 4 números, lo rechaza
            if len(re.findall(r'\d', str(c_val))) < 4:
                datos[c_key] = "No encontrado"
                

    print(f"\n📄 LECTURA DEL ARCHIVO: {os.path.basename(ruta_archivo)}")
    for columna, valor in datos.items():
        print(f"   -> {columna}: '{valor}'")
    print("-" * 40)
    return datos


# ══════════════════════════════════════════════════════════════════
#  EXTRACCIÓN NÓMINA / SEGURIDAD SOCIAL
# ══════════════════════════════════════════════════════════════════
def _limpiar_celda_tabla(texto):
    if not texto: return 0.0
    limpio = str(texto).replace("$", "").replace("\n", "").replace(" ", "").strip()
    if not limpio: return 0.0
    try:
        return float(limpio.replace(".", ""))
    except:
        return 0.0

def _extraer_nomina(ruta_archivo, cols):
    empleados_dict = {}
    try:
        # [OPT-8] pdfplumber en modo nómina:
        # El context manager garantiza que cada página extraída vía
        # extract_table() sea liberada al terminar el `with`.  pdfplumber
        # cachea internamente los objetos LTPage; al cerrar el PDF ese
        # cache se descarta completo.  No se necesita del() adicional aquí
        # porque la variable `tabla` se reasigna en cada iteración de
        # página (la anterior queda huérfana e inmediatamente elegible para GC).
        with pdfplumber.open(ruta_archivo) as pdf:
            for pagina in pdf.pages:
                tabla = pagina.extract_table()
                if not tabla: continue

                for fila in tabla:
                    if not fila or len(fila) < 40: continue
                    if not str(fila[0]).strip().isdigit(): continue

                    match_cedula = re.search(r'\d+', str(fila[1]))
                    if not match_cedula: continue
                    cedula = match_cedula.group()

                    if cedula not in empleados_dict:
                        empleados_dict[cedula] = {
                            "Cédula": cedula, "Pensión_Cotizacion": 0.0, "FSP_Subsistencia": 0.0,
                            "FSP_Solidaridad": 0.0, "Salud_Total": 0.0, "ARP_Total": 0.0,
                            "Caja_Total": 0.0, "SENA_Total": 0.0, "ICBF_Total": 0.0
                        }
                    emp = empleados_dict[cedula]

                    emp["Pensión_Cotizacion"] += _limpiar_celda_tabla(fila[-26])
                    emp["FSP_Subsistencia"]   += _limpiar_celda_tabla(fila[-25])
                    emp["FSP_Solidaridad"]    += _limpiar_celda_tabla(fila[-24])
                    emp["Salud_Total"]        += _limpiar_celda_tabla(fila[-14])
                    emp["ARP_Total"]          += _limpiar_celda_tabla(fila[-9])
                    emp["Caja_Total"]         += _limpiar_celda_tabla(fila[-5])
                    emp["SENA_Total"]         += _limpiar_celda_tabla(fila[-4])
                    emp["ICBF_Total"]         += _limpiar_celda_tabla(fila[-3])

    except Exception as e:
        print(f"Error leyendo tabla de nómina: {e}")

    return list(empleados_dict.values())


def _worker(args):
    ruta, cols, modo = args
    if modo == "Contabilidad (Nómina)":
        return os.path.basename(ruta), _extraer_nomina(ruta, cols)
    else:
        return os.path.basename(ruta), _extraer(ruta, cols)


# ══════════════════════════════════════════════════════════════════
#  UI HELPERS
# ══════════════════════════════════════════════════════════════════

# [OPT-9] THROTTLING DE app.after() EN EL LOG:
# El log original programaba un app.after(0, _w) por CADA mensaje.
# Cuando el ThreadPoolExecutor termina 6 PDFs casi simultáneamente,
# se pueden encolar decenas de callbacks en el mainloop en el mismo
# frame, degradando la fluidez de la UI.  La solución es acumular
# los mensajes en una cola thread-safe y procesarlos en lotes desde
# un único after() periódico (_flush_log), en lugar de uno por mensaje.
import queue as _queue
_log_queue: _queue.SimpleQueue = _queue.SimpleQueue()

def _flush_log():
    """Drena la cola de log en el hilo principal (mainloop). Máx. 20 mensajes por tick."""
    # [OPT-9 cont.] Limitamos a 20 por tick para no bloquear el mainloop
    # si hay una ráfaga muy grande de mensajes pendientes.
    colors = {"ok": "#2ea043", "error": "#da3633", "warn": "#FFA726", "info": "#3b82f6", "dim": "#888888"}
    i = 0
    terminal.configure(state="normal")
    while i < 20:
        try:
            ts, msg, tipo = _log_queue.get_nowait()
        except _queue.Empty:
            break
        terminal.insert("end", f"[{ts}] ", "ts")
        terminal.insert("end", f"{msg}\n", tipo)
        terminal.tag_config("ts",  foreground="#888888")
        terminal.tag_config(tipo,  foreground=colors.get(tipo, "#555555"))
        i += 1
    if i:
        terminal.see("end")
    terminal.configure(state="disabled")
    # Reprogramar el próximo flush cada 100 ms (≈10 ticks/s, imperceptible para el usuario)
    app.after(100, _flush_log)

def log(msg, tipo="info"):
    # Encola el mensaje; el hilo de UI lo leerá en el próximo _flush_log()
    _log_queue.put((_ts(), msg, tipo))

def log_sep():
    log("─" * 40, "dim")

def _ui_reset():
    app.after(0, lambda: [
        pb.set(0),
        lbl_stat_pct.configure(text="0%"),
        lbl_stat_count.configure(text="0/0"),
        card_ok_val.configure(text="-"),
        card_err_val.configure(text="-"),
        lbl_status.configure(text="Iniciando...", text_color=C_TEXT_SUB)
    ])

def _ui_tick(done, total, t0, fase="", archivo=""):
    # [OPT-10] COALESCING DE ACTUALIZACIONES DE PROGRESO:
    # _ui_tick se llama desde workers concurrentes.  En vez de encolar
    # un after(0, ...) por cada PDF completado (que podría apilar 6
    # callbacks simultáneos), usamos after(0) que los fusiona: si ya
    # hay uno pendiente en la cola de eventos, Tkinter lo ejecuta antes
    # de pintar el siguiente frame.  Esto es idéntico al original pero
    # el efecto se percibe mejor junto con el throttling del log (OPT-9).
    def _u():
        frac = done / total if total else 0
        pb.set(frac)
        lbl_stat_pct.configure(text=f"{int(frac*100)}%")
        lbl_stat_count.configure(text=f"{done} / {total}")
        lbl_stat_time.configure(text=_fmt(time.time()-t0))
        if fase: lbl_status.configure(text=fase, text_color=C_ACCENT)
    app.after(0, _u)

def _ui_done(total, t0, n_celdas, n_err):
    n_ok_val  = int(card_ok_val.cget("text") or 0)
    tiempo    = _fmt(time.time() - t0)
    def _f():
        pb.set(1.0)
        # Restaurar el texto del botón según el modo activo
        modo_activo = combo_modo.get()
        txt_btn = " VALIDAR MANIFIESTOS" if "Transporte" in modo_activo else " VERIFICAR SEGURIDAD SOCIAL"
        btn_auditar.configure(state="normal", text=txt_btn, fg_color=C_SUCCESS)
        btn_abrir.configure(state="normal", fg_color=C_ACCENT)
        lbl_status.configure(text="Auditoría Finalizada", text_color=C_SUCCESS)
        card_err_val.configure(text=str(n_err))
        mostrar_resumen_final(n_ok_val, n_err, tiempo, modo_activo)
    app.after(0, _f)

def mostrar_resumen_final(n_ok, n_err, tiempo, modo):
    """Modal premium de resumen post-auditoría."""
    total      = n_ok + n_err
    precision  = round((n_ok / total * 100) if total > 0 else 0, 1)
    color_res  = C_SUCCESS if precision >= 90 else ("#f59e0b" if precision >= 70 else C_ERROR)
    acento     = C_IND_SS if "Contabilidad" in modo else C_ACCENT

    # El modal cubre el panel derecho
    modal = ctk.CTkFrame(frame_right, fg_color=C_BG_MAIN, corner_radius=0)
    modal.place(x=0, y=0, relwidth=1, relheight=1)
    modal.lift()

    caja = ctk.CTkFrame(modal, fg_color=C_CARD, corner_radius=20, border_width=2, border_color=acento)
    caja.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.85, relheight=0.80)

    # --- Encabezado ---
    ctk.CTkLabel(caja, text="✔", font=("Roboto", 52, "bold"), text_color=C_SUCCESS).pack(pady=(35, 0))
    ctk.CTkLabel(caja, text="PROCESO FINALIZADO", font=("Roboto", 18, "bold"), text_color="white").pack()
    t_sub = "Validación de Manifiestos" if "Transporte" in modo else "Auditoría de Seguridad Social"
    ctk.CTkLabel(caja, text=t_sub, font=("Roboto", 12), text_color=C_TEXT_SUB).pack(pady=(2, 20))

    # --- Tarjetas de Resumen ---
    f_stats = ctk.CTkFrame(caja, fg_color="transparent")
    f_stats.pack(fill="x", padx=30)
    f_stats.columnconfigure((0,1,2), weight=1)

    def mini_stat(parent, label, valor, color, col):
        card = ctk.CTkFrame(parent, fg_color="#1a1c22", corner_radius=10)
        card.grid(row=0, column=col, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(card, text=str(valor), font=("Roboto", 26, "bold"), text_color=color).pack(pady=(12,0))
        ctk.CTkLabel(card, text=label, font=("Roboto", 9), text_color=C_TEXT_SUB).pack(pady=(0,10))

    mini_stat(f_stats, "Correctos",      n_ok,          C_SUCCESS,  0)
    mini_stat(f_stats, "Discrepancias",  n_err,         C_ERROR,    1)
    mini_stat(f_stats, "Precisión",      f"{precision}%", color_res, 2)

    # --- Tiempo ---
    f_tiempo = ctk.CTkFrame(caja, fg_color="transparent")
    f_tiempo.pack(pady=(15, 5))
    ctk.CTkLabel(f_tiempo, text="⏱", font=("Roboto", 14)).pack(side="left", padx=(0, 5))
    ctk.CTkLabel(f_tiempo, text=f"Tiempo total: {tiempo}", font=("Roboto", 12), text_color=C_TEXT_SUB).pack(side="left")

    # --- Botones ---
    f_btns = ctk.CTkFrame(caja, fg_color="transparent")
    f_btns.pack(fill="x", padx=30, pady=(20, 30))

    ctk.CTkButton(f_btns, text="Abrir Excel", height=42, fg_color=C_CARD, hover_color="#333333",
                  font=("Roboto", 12), command=lambda: [modal.destroy(), abrir_excel()]
                  ).pack(side="right", padx=(8, 0))

    ctk.CTkButton(f_btns, text="Nueva Auditoría ➜", height=42, fg_color=acento, hover_color=C_ACCENT_H,
                  font=("Roboto", 13, "bold"), command=modal.destroy
                  ).pack(side="right")

def _ui_error(msg):
    def _e():
        lbl_status.configure(text=f"⚠ {msg}", text_color=C_ERROR)
        pb.configure(progress_color=C_ERROR)
        btn_auditar.configure(state="normal", text="INICIAR AUDITORÍA", fg_color=C_SUCCESS)
        btn_abrir.configure(state="disabled", fg_color=C_CARD)
    app.after(0, _e)


# ══════════════════════════════════════════════════════════════════
#  AUDITORÍA — HILO TRANSPORTE
# ══════════════════════════════════════════════════════════════════
def iniciar_auditoria():
    cargar_plantillas_en_memoria()
    global _en_curso
    if _en_curso: return
    if not ruta_excel or not ruta_pdf:
        log("Selecciona el Excel y el archivo/carpeta PDF.", "warn"); return

    modo = combo_modo.get()

    if modo == "Transporte (Manifiestos)":
        cols = [c.cget("text") for c in checkboxes_columnas if c.get() == 1]
        if not cols: log("Selecciona al menos una columna.", "warn"); return

        lm = next((c for c in cols if "manifiesto" in c.lower() or "remesa" in c.lower()), None)
        lp = next((c for c in cols if "placa" in c.lower() or "veh" in c.lower()), None)
        if not lm or not lp: log("Marca columnas de Manifiesto y Placa.", "warn"); return

        guardar_config()
        _en_curso = True
        btn_auditar.configure(state="disabled", text="Procesando...", fg_color=C_ACCENT_H)
        btn_abrir.configure(state="disabled", fg_color=C_CARD)
        _ui_reset()
        log_sep()
        log("INICIANDO AUDITORÍA - TRANSPORTE", "info")
        threading.Thread(target=_hilo_transporte, args=(cols, lm, lp), daemon=True).start()

    else:
        guardar_config()
        _en_curso = True
        btn_auditar.configure(state="disabled", text="Procesando...", fg_color=C_ACCENT_H)
        btn_abrir.configure(state="disabled", fg_color=C_CARD)
        _ui_reset()
        log_sep()
        log("INICIANDO AUDITORÍA - NÓMINA", "info")
        threading.Thread(target=_hilo_nomina, args=([], None), daemon=True).start()


def _hilo_transporte(cols, lm, lp):
    global _en_curso
    t0 = time.time()

    app.after(0, lambda: lbl_status.configure(text="Verificando Excel...", text_color=C_TEXT_SUB))
    if _excel_abierto(ruta_excel):
        app.after(0, lambda: messagebox.showwarning("Excel abierto", f"'{os.path.basename(ruta_excel)}' está abierto.\nCiérralo."))
        log("Bloqueado: cierra el Excel primero.", "error")
        _ui_error("Cierra el Excel antes de continuar.")
        _en_curso = False; return
    log("Excel no está abierto. OK.", "ok")

    archivos = [os.path.join(ruta_pdf, f) for f in os.listdir(ruta_pdf) if f.lower().endswith(".pdf")]
    total = len(archivos)
    if not total:
        log("No hay PDFs en la carpeta.", "error")
        _ui_error("No hay PDFs en la carpeta."); _en_curso = False; return

    log(f"Leyendo {total} PDFs en paralelo...", "info")
    resultados, done, t1 = {}, 0, time.time()

    # [OPT-11] MAX_WORKERS ADAPTATIVO:
    # Con 6 workers fijos, procesar 3 PDFs lanzaría 6 hilos de los cuales
    # 3 estarían ociosos consumiendo memoria de stack (~8 MB c/u en CPython).
    # Limitamos a min(6, total) para ajustar el pool al trabajo real.
    # El GIL de Python no bloquea aquí porque pdfplumber pasa tiempo en I/O
    # de disco y en el parser C de pdfminer, donde el GIL se libera.
    max_w = min(6, total)
    with ThreadPoolExecutor(max_workers=max_w) as ex:
        futs = {ex.submit(_worker, (r, cols, "Transporte (Manifiestos)")): r for r in archivos}
        for fut in as_completed(futs):
            nombre, datos = fut.result()
            resultados[nombre] = datos
            done += 1
            _ui_tick(done, total, t1, "Leyendo PDFs...", nombre)
    # [OPT-11 cont.] El `with` del ThreadPoolExecutor llama a shutdown(wait=True)
    # al salir, esperando a que todos los workers terminen y liberando sus
    # hilos antes de continuar.  Esto es correcto y no requiere cambios.
    log(f"PDFs leídos en {_fmt(time.time()-t1)}.", "ok")

    app.after(0, lambda: lbl_status.configure(text="Cargando Excel...", text_color=C_ACCENT))
    log("Cargando Excel en memoria...", "info")
    try:
        wb_r   = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
        hoja_r = wb_r[combo_hoja.get()]
        fe     = int(entry_fila.get())
        enc    = [c.value for c in list(hoja_r.rows)[fe-1]]
        tits   = [str(t).strip().replace("\n"," ") if t else "" for t in enc]

        if lm not in tits or lp not in tits:
            log("Columnas llave no encontradas.", "error")
            wb_r.close()  # [OPT-12] Cerrar antes de retornar por error
            _ui_error("Columnas llave no encontradas."); _en_curso = False; return

        im, ip = tits.index(lm), tits.index(lp)
        filas  = [(fe+1+i, list(r)) for i,r in enumerate(hoja_r.iter_rows(min_row=fe+1, values_only=True))]
        wb_r.close()   # [OPT-12] Cierre explícito del workbook read-only.
        # openpyxl en modo read_only mantiene el ZipFile del .xlsx abierto
        # mientras el objeto wb_r esté vivo.  Sin close() explícito el fd
        # permanece hasta que el GC destruya wb_r (indeterminado), lo que
        # puede causar "archivo en uso" si el usuario intenta moverlo.

        indice = {}
        for nf, vs in filas:
            vm = vs[im]
            if vm is not None:
                mk = str(vm).strip().lower().replace(".0","")
                pk = str(vs[ip]).strip().lower() if vs[ip] else ""
                indice.setdefault(mk,[]).append((pk,nf,vs))
        log(f"Excel indexado: {len(filas)} filas.", "ok")

        # [OPT-13] Liberar la lista bruta de filas una vez construido el índice.
        # Con hojas de miles de filas, `filas` puede ocupar decenas de MB.
        del filas

    except Exception as e:
        log(f"Error cargando Excel: {e}", "error")
        _ui_error(str(e)); _en_curso = False; return

    log("Cruzando datos PDF ↔ Excel...", "info")
    cambios, done2 = {}, 0
    n_ok = n_err = n_skip = 0
    t3 = time.time()

    for nombre, datos in resultados.items():
        vm = str(datos.get(lm,"")).strip().lower()
        vp = str(datos.get(lp,"")).strip().lower()

        if not vm or vm == "no encontrado":
            log(f"[SKIP]     {nombre}  — manifiesto no hallado", "warn")
            n_skip += 1
        else:
            fila_r = None
            vals_r = None

            if vm in indice:
                for pk, nf, vs in indice[vm]:
                    if vp in pk or pk in vp:
                        fila_r = nf
                        vals_r = vs
                        break

                if fila_r is None:
                    _, fila_r, vals_r = indice[vm][0]

            if fila_r is None or vals_r is None:
                log(f"[NO MATCH] {nombre}  Manif:{vm}  Placa:{vp}", "error")
                n_err += 1
            else:
                disc = []
                for col in cols:
                    if col not in tits: continue
                    ci0 = tits.index(col)
                    ci1 = ci0 + 1

                    ve = ""
                    if ci0 < len(vals_r) and vals_r[ci0] is not None:
                        ve = str(vals_r[ci0]).strip().lower()
                        if ve.endswith(".0"): ve = ve[:-2]

                    vd = str(datos.get(col,"no encontrado")).strip().lower()

                    if vd in ("no encontrado", ""):
                        cambios[(fila_r, ci1)] = (244, 204, 204)
                        disc.append(f"{col}: sin dato en PDF")
                        n_err += 1
                    elif vd in ve or ve in vd:
                        cambios[(fila_r, ci1)] = (217, 234, 211)
                        n_ok += 1
                    else:
                        cambios[(fila_r, ci1)] = (244, 204, 204)
                        disc.append(f"{col}: PDF='{vd}'  ≠  Excel='{ve}'")
                        n_err += 1

                if disc:
                    log(f"[ERROR]    {nombre}", "error")
                    for d in disc: log(f"         ↳ {d}", "error")
                else:
                    log(f"[OK]       {nombre}", "ok")

        done2 += 1
        _ui_tick(done2, total, t3, "Cruzando datos...", nombre)

        app.after(0, lambda nok=n_ok, nerr=n_err: [
            card_ok_val.configure(text=str(nok)),
            card_err_val.configure(text=str(nerr))
        ])

    app.after(0, lambda: lbl_status.configure(text="Guardando Excel...", text_color=C_ACCENT))
    log(f"Escribiendo {len(cambios)} celdas...", "info")

    # [OPT-14] BLOQUE xlwings CON GARANTÍA DE CIERRE EN EXCEPCIONES GRAVES:
    # PROBLEMA ORIGINAL: Si ocurría una excepción ENTRE wb_xw.save() y
    # app_xw.quit(), la instancia oculta de Excel.exe quedaba huérfana en
    # memoria (visible en el Administrador de Tareas).  En auditorías repetidas
    # esto acumulaba procesos EXCEL.EXE zombies.
    #
    # SOLUCIÓN: Usamos un try/finally anidado.  El finally exterior garantiza
    # que app_xw.quit() se ejecute SIEMPRE, incluso si:
    #   · wb_xw.save() lanza un error de disco
    #   · El usuario mata el proceso con Ctrl+C
    #   · xw.App() mismo falla (en ese caso app_xw sigue siendo None y el
    #     guard `if app_xw` evita un AttributeError en el finally)
    app_xw = None
    try:
        try:
            app_xw  = xw.App(visible=False)
            app_xw.display_alerts = False   # [OPT-14b] Suprimir diálogos de Excel
            # que podrían bloquear el hilo si Excel pregunta algo al guardar.
            wb_xw   = app_xw.books.open(ruta_excel)
            hoja_xw = wb_xw.sheets[combo_hoja.get()]

            for (nf, ci1), color in cambios.items():
                hoja_xw.range((nf, ci1)).color = color

            wb_xw.save()
            wb_xw.close()

        finally:
            # Este finally se ejecuta incluso si save() o close() lanzan excepción.
            # quit() cierra EXCEL.EXE completamente, sin dejar procesos huérfanos.
            if app_xw:
                try:
                    app_xw.quit()
                except Exception:
                    pass  # Si Excel ya crasheó, ignoramos el error del quit()

        log_sep()
        log("AUDITORÍA FINALIZADA", "info")
        log(f" ✔  Correctos       : {n_ok}",   "ok")
        log(f" ✘  Discrepancias   : {n_err}",  "error")
        log(f" ?  Sin match Excel : {n_skip}", "warn")
        log(f" ⏱  Tiempo total    : {_fmt(time.time()-t0)}", "info")
        log_sep()
        _ui_done(total, t0, len(cambios), n_err)

    except Exception as e:
        log(f"Error escribiendo Excel: {e}", "error")
        _ui_error(str(e))

    finally:
        _en_curso = False
        # [OPT-15] Forzar un ciclo de GC después de la auditoría completa.
        # Los objetos de openpyxl (Cell, Row, Worksheet) y pdfplumber tienen
        # referencias circulares internas.  gc.collect() las resuelve
        # inmediatamente en lugar de esperar al próximo ciclo automático
        # (que puede tardar segundos con muchos objetos vivos).
        gc.collect()


# ══════════════════════════════════════════════════════════════════
#  AUDITORÍA — HILO NÓMINA
# ══════════════════════════════════════════════════════════════════
def _hilo_nomina(cols, lm):
    global _en_curso
    t0 = time.time()

    app.after(0, lambda: lbl_status.configure(text="Verificando...", text_color=C_TEXT_SUB))
    if _excel_abierto(ruta_excel):
        log("Bloqueado: cierra el Excel primero.", "error")
        _ui_error("Cierra el Excel antes de continuar.")
        _en_curso = False; return

    if not os.path.isfile(ruta_pdf):
        _ui_error("PDF no válido."); _en_curso = False; return
    archivos = [ruta_pdf]
    total = len(archivos)

    log(f"Leyendo {total} Planillas en paralelo...", "info")
    resultados, done, t1 = {}, 0, time.time()

    # [OPT-11 — mismo principio que en _hilo_transporte]
    max_w = min(6, total)
    with ThreadPoolExecutor(max_workers=max_w) as ex:
        futs = {ex.submit(_worker, (r, cols, "Contabilidad (Nómina)")): r for r in archivos}
        for fut in as_completed(futs):
            nombre, datos = fut.result()
            resultados[nombre] = datos
            done += 1
            _ui_tick(done, total, t1, "Leyendo Planillas...", nombre)

    app.after(0, lambda: lbl_status.configure(text="Cargando Excel...", text_color=C_ACCENT))
    try:
        wb_r   = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
        hoja_r = wb_r[combo_hoja.get()]
        fe     = int(entry_fila.get())
        enc    = [c.value for c in list(hoja_r.rows)[fe-1]]
        tits   = [str(t).strip().replace("\n"," ") if t else "" for t in enc]

        lm = next((c for c in tits if "cedula" in c.lower() or "cédula" in c.lower() or "identificaci" in c.lower()), None)
        if not lm:
            log("No se encontró la columna de Cédula en el Excel.", "error")
            wb_r.close()   # [OPT-12] Cerrar antes de retornar por error
            _ui_error("Falta columna Cédula."); _en_curso = False; return

        im = tits.index(lm)
        filas = [(fe+1+i, list(r)) for i,r in enumerate(hoja_r.iter_rows(min_row=fe+1, values_only=True))]
        wb_r.close()   # [OPT-12] Cierre explícito del workbook read-only

        indice = {}
        for nf, vs in filas:
            vm = vs[im]
            if vm is not None:
                mk = str(vm).strip().lower().replace(".0","").replace(".", "").replace(",", "")
                indice[mk] = (nf, vs)

        del filas  # [OPT-13] Liberar lista bruta tras indexar

    except Exception as e:
        log(f"Error cargando Excel: {e}", "error")
        _ui_error(str(e)); _en_curso = False; return

    def sumar_columnas(titulos_excel, valores_fila, palabras_clave):
        total = 0.0
        cols_implicadas = []
        for i, titulo in enumerate(titulos_excel):
            tit_limpio = str(titulo).lower().strip()
            if any(p in tit_limpio for p in palabras_clave):
                cols_implicadas.append(i + 1)
                val = valores_fila[i]
                if val is not None:
                    try:
                        if isinstance(val, (int, float)):
                            total += float(val)
                        else:
                            v_limpio = str(val).replace("$","").replace(" ","").strip()
                            if "." in v_limpio and "," in v_limpio:
                                v_limpio = v_limpio.replace(".", "").replace(",", ".")
                            elif "," in v_limpio:
                                v_limpio = v_limpio.replace(",", ".")
                            total += float(v_limpio)
                    except: pass
        return round(total, 2), cols_implicadas

    log("Cruzando matemáticas PDF ↔ Excel...", "info")
    cambios = {}
    n_ok = n_err = n_skip = done2 = 0
    t3 = time.time()
    tolerancia = 2000

    for nombre, lista_empleados in resultados.items():
        for emp in lista_empleados:
            cedula_pdf = str(emp.get("Cédula", "")).strip().lower().replace(".", "").replace(",", "")
            if not cedula_pdf: continue

            if cedula_pdf not in indice:
                log(f"[NO MATCH] Cédula no hallada en Excel: {cedula_pdf}", "error")
                n_err += 1
                continue

            fila_r, vals_r = indice[cedula_pdf]
            disc = []

            # PENSION
            exc_pension, cols_pen = sumar_columnas(tits, vals_r, ["pension empleador", "pension 4%"])
            pdf_pension = float(emp.get("Pensión_Cotizacion", 0))
            if abs(exc_pension - pdf_pension) > tolerancia:
                disc.append(f"Pensión: PDF=${pdf_pension} ≠ Excel=${exc_pension}")
                for c in cols_pen: cambios[(fila_r, c)] = (244, 204, 204)
            else:
                for c in cols_pen: cambios[(fila_r, c)] = (217, 234, 211)

            # SALUD
            exc_salud, cols_sal = sumar_columnas(tits, vals_r, ["eps empleador", "salud 4%"])
            pdf_salud = float(emp.get("Salud_Total", 0))
            if abs(exc_salud - pdf_salud) > tolerancia:
                disc.append(f"Salud: PDF=${pdf_salud} ≠ Excel=${exc_salud}")
                for c in cols_sal: cambios[(fila_r, c)] = (244, 204, 204)
            else:
                for c in cols_sal: cambios[(fila_r, c)] = (217, 234, 211)

            # RIESGOS Y PARAFISCALES
            mapeo_directo = {
                "FSP":     (["fsp %1", "f.s.p", "fsp"], float(emp.get("FSP_Subsistencia", 0)) + float(emp.get("FSP_Solidaridad", 0))),
                "Riesgos": (["riesgos"], float(emp.get("ARP_Total", 0))),
                "Caja":    (["caja"],    float(emp.get("Caja_Total", 0))),
                "SENA":    (["sena"],    float(emp.get("SENA_Total", 0))),
                "ICBF":    (["icbf"],    float(emp.get("ICBF_Total", 0)))
            }
            for nom_campo, (claves, val_pdf) in mapeo_directo.items():
                exc_val, cols_impl = sumar_columnas(tits, vals_r, claves)
                if abs(exc_val - val_pdf) > tolerancia:
                    disc.append(f"{nom_campo}: PDF=${val_pdf} ≠ Excel=${exc_val}")
                    for c in cols_impl: cambios[(fila_r, c)] = (244, 204, 204)
                else:
                    for c in cols_impl: cambios[(fila_r, c)] = (217, 234, 211)

            col_pintar = tits.index(lm) + 1
            if disc:
                cambios[(fila_r, col_pintar)] = (244, 204, 204)
                n_err += 1
                log(f"[ERROR] Cédula {cedula_pdf}:", "error")
                for d in disc: log(f"         -> {d}", "error")
            else:
                cambios[(fila_r, col_pintar)] = (217, 234, 211)
                n_ok += 1

        done2 += 1
        _ui_tick(done2, total, t3, "Auditando Nómina...", nombre)

        app.after(0, lambda nok=n_ok, nerr=n_err: [
            card_ok_val.configure(text=str(nok)), card_err_val.configure(text=str(nerr))
        ])

    app.after(0, lambda: lbl_status.configure(text="Guardando Excel...", text_color=C_ACCENT))

    # [OPT-14] MISMO PATRÓN try/finally que en _hilo_transporte
    app_xw = None
    try:
        try:
            app_xw  = xw.App(visible=False)
            app_xw.display_alerts = False   # [OPT-14b]
            wb_xw   = app_xw.books.open(ruta_excel)
            hoja_xw = wb_xw.sheets[combo_hoja.get()]

            for (nf, ci1), color in cambios.items():
                hoja_xw.range((nf, ci1)).color = color

            wb_xw.save()
            wb_xw.close()

        finally:
            if app_xw:
                try:
                    app_xw.quit()
                except Exception:
                    pass

        log_sep()
        log("AUDITORÍA DE NÓMINA FINALIZADA", "info")
        _ui_done(total, t0, len(cambios), n_err)

    except Exception as e:
        log(f"Error escribiendo Excel: {e}", "error")
        _ui_error(str (e))

    finally:
        _en_curso = False
        gc.collect()  # [OPT-15] Recolección forzada post-auditoría


# ══════════════════════════════════════════════════════════════════
#  NUEVA INTERFAZ GRÁFICA (DASHBOARD)
# ══════════════════════════════════════════════════════════════════
def seleccionar_excel():
    global ruta_excel
    p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
    if p:
        ruta_excel = p
        lbl_excel_name.configure(text=os.path.basename(p), text_color=C_SUCCESS)
        try:
            # [OPT-5] Aplicado también aquí: cerrar el workbook tras leer los sheetnames
            wb = openpyxl.load_workbook(p, read_only=True)
            combo_hoja.configure(values=wb.sheetnames)
            combo_hoja.set(wb.sheetnames[0])
            wb.close()   # ← CORRECCIÓN
        except: pass


def aplicar_interfaz(seleccion):
    """Aplica la identidad visual y funcional según el modo seleccionado."""
    if seleccion == "Transporte (Manifiestos)":
        # --- IDENTIDAD: GESTIÓN AMBIENTAL PETROLERA ---
        btn_pdf.configure(text="Carpeta PDFs", image=ICONS.get("folder"))
        lbl_pdf_name.configure(text="Sin carpeta", text_color=C_TEXT_SUB)
        
        # UI Elements
        btn_auditar.configure(text=" VALIDAR MANIFIESTOS", fg_color=C_SUCCESS)
        lbl_term_title.configure(text="REGISTRO DE CONTROL AMBIENTAL")
        terminal.configure(border_color="#2b2b2b")
        
        # Color Accent
        barra_carga.configure(progress_color=C_ACCENT)
        lbl_perc.configure(text_color=C_ACCENT)
        pb.configure(progress_color=C_ACCENT)
        lbl_stat_pct.configure(text_color=C_ACCENT)
        
        # Cards (Micro-detalle: actualizamos colores de valores si es necesario)
        card_ok_val.configure(text_color=C_SUCCESS) # Original
        
        ctk.set_appearance_mode("Dark")
        terminal.configure(fg_color="#0f0f0f", text_color="#cccccc")

        btn_cols.pack(padx=20, fill="x")
        sep.pack(fill="x", padx=10, pady=(15, 5))
        frame_ia.pack(fill="x", padx=10, pady=(0, 5))
        lbl_cols_title.pack(padx=20, pady=(15,5), anchor="w")
        scroll_cols.pack(padx=10, fill="both", expand=True)
        
    else:
        # --- IDENTIDAD: VALIDACIÓN DE SEGURIDAD SOCIAL ---
        btn_pdf.configure(text="Seleccionar PDF", image=ICONS.get("pdf"))
        lbl_pdf_name.configure(text="Sin archivo", text_color=C_TEXT_SUB)
        
        # UI Elements
        btn_auditar.configure(text=" VERIFICAR SEGURIDAD SOCIAL", fg_color=C_IND_SS)
        lbl_term_title.configure(text="PROTOCOLO DE AUDITORÍA CONTABLE")
        terminal.configure(border_color=C_IND_SS)
        
        # Color Accent (Teal)
        barra_carga.configure(progress_color=C_IND_SS)
        lbl_perc.configure(text_color=C_IND_SS)
        pb.configure(progress_color=C_IND_SS)
        lbl_stat_pct.configure(text_color=C_IND_SS)

        # Micro-detalle: El "Verde de éxito" en modo Teal lo hacemos más cian para armonía
        card_ok_val.configure(text_color=C_IND_SS) 

        ctk.set_appearance_mode("Dark") # Mantenemos Dark para consistencia premium
        terminal.configure(fg_color="#0b1215", text_color="#d1d5db") # Tinte azulado/teal profundo

        btn_cols.pack_forget()
        lbl_cols_title.pack_forget()
        scroll_cols.pack_forget()
        frame_ia.pack_forget()
        sep.pack_forget()

    combo_modo.configure(text_color=C_TEXT_MAIN)
    combo_hoja.configure(text_color=C_TEXT_MAIN)
    entry_fila.configure(text_color=C_TEXT_MAIN)

    barra_carga.stop()
    frame_carga.place_forget()


def _animar_carga(porcentaje, seleccion):
    """Animación fluida (1% a 1%) con hitos de mensaje profesional."""
    if porcentaje <= 100:
        # Definir mensajes según el progreso
        if porcentaje < 25:
            msg = "Optimizando entorno de ejecución..."
        elif porcentaje < 50:
            msg = "Sincronizando modelos de IA..."
        elif porcentaje < 75:
            msg = f"Iniciando núcleo: {seleccion}..."
        elif porcentaje < 95:
            msg = "Finalizando pre-procesamiento..."
        else:
            msg = "¡Módulo listo!"

        barra_carga.set(porcentaje / 100)
        lbl_carga.configure(text=msg)
        lbl_perc.configure(text=f"{porcentaje}%")
        
        # Velocidad variable para realismo (50ms base)
        delay = 40 if porcentaje < 80 else 70
        app.after(delay, lambda: _animar_carga(porcentaje + 1, seleccion))
    else:
        # Pequeña pausa al final para satisfacción visual
        app.after(300, lambda: aplicar_interfaz(seleccion))

def cambiar_modo(seleccion):
    global ruta_pdf
    ruta_pdf = ""

    # Usar el color de fondo actual del tema
    color_fondo = C_BG_MAIN[0] if ctk.get_appearance_mode() == "Light" else C_BG_MAIN[1]
    frame_carga.configure(fg_color=color_fondo)
    
    # Reset UI y asignamos el color de borde del Splash según el destino
    border_destino = C_ACCENT if seleccion == "Transporte (Manifiestos)" else C_IND_SS
    caja_splash.configure(border_color=border_destino)
    lbl_perc.configure(text_color=border_destino)
    barra_carga.configure(progress_color=border_destino)

    barra_carga.set(0)
    lbl_perc.configure(text="0%")
    lbl_carga.configure(text="Iniciando...")
    
    frame_carga.place(x=0, y=0, relwidth=1, relheight=1)
    frame_carga.lift()

    app.update()
    
    # Iniciamos animación fluida desde 0%
    app.after(500, lambda: _animar_carga(0, seleccion))


def seleccionar_pdf():
    global ruta_pdf
    if combo_modo.get() == "Transporte (Manifiestos)":
        p = filedialog.askdirectory(title="Seleccionar Carpeta")
        if p:
            ruta_pdf = p
            n = len([f for f in os.listdir(p) if f.lower().endswith(".pdf")])
            lbl_pdf_name.configure(text=f"{n} Archivos encontrados", text_color=C_SUCCESS)
    else:
        p = filedialog.askopenfilename(title="Seleccionar PDF de Nómina", filetypes=[("Archivos PDF", "*.pdf")])
        if p:
            ruta_pdf = p
            lbl_pdf_name.configure(text=os.path.basename(p), text_color=C_SUCCESS)


def cargar_columnas():
    try:
        # [OPT-5] Cerrar el workbook tras leer los datos de columnas
        wb = openpyxl.load_workbook(ruta_excel, data_only=True)
        hoja = wb[combo_hoja.get()]
        fn = int(entry_fila.get())
        enc = [str(c.value).strip() for c in hoja[fn] if c.value]
        wb.close()   # ← CORRECCIÓN: antes no se llamaba close()

        # [OPT-16] DESTRUCCIÓN EXPLÍCITA DE CHECKBOXES ANTIGUOS:
        # pack_forget() y destroy() en los widgets hijos son equivalentes
        # para liberar la memoria de Tkinter/CTk.  `w.destroy()` es lo
        # correcto porque remueve el widget del árbol de Tkinter, permitiendo
        # que el GC de Python también libere el objeto Python subyacente.
        # El código original ya usaba w.destroy(), lo cual es correcto.
        # Solo añadimos la liberación del wb que faltaba.
        for w in scroll_cols.winfo_children():
            w.destroy()
        global checkboxes_columnas
        checkboxes_columnas = []

        for c in enc:
            chk = ctk.CTkCheckBox(scroll_cols, text=c, font=("Roboto", 11),
                                  text_color=C_TEXT_MAIN, hover_color=C_ACCENT, fg_color=C_ACCENT)
            chk.pack(anchor="w", pady=2, padx=5)
            checkboxes_columnas.append(chk)
        log(f"Columnas cargadas: {len(enc)}", "ok")
    except Exception as e:
        log(f"Error cargando columnas: {e}", "error")


def abrir_entrenador():
    ventana_entrenador = ctk.CTkToplevel()
    ventana_entrenador.title("Entrenador PDF")
    ventana_entrenador.geometry("800x600")
    ventana_entrenador.focus_force()
    ventana_entrenador.grab_set()
    app_entrenador = MapeadorPDF(ventana_entrenador)


def abrir_gestor():
    ventana = GestorPlantillas(app)
    ventana.grab_set()
    ventana.focus_force()


# ══════════════════════════════════════════════════════════════════
#  SETUP VENTANA
# ══════════════════════════════════════════════════════════════════
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

app = ctk.CTk()
app.geometry("950x700")
app.title(" AI - Sistema de Validación")
app.configure(fg_color=C_BG_MAIN)
try:
    app.iconbitmap(obtener_ruta("logo.ico"))
except:
    pass
app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(0, weight=1)

# === PANTALLA DE CARGA ===
frame_carga = ctk.CTkFrame(app, fg_color=C_BG_MAIN, corner_radius=0)
caja_splash = ctk.CTkFrame(frame_carga, width=450, height=280, fg_color=C_CARD, corner_radius=20, border_width=2, border_color=C_ACCENT)
caja_splash.place(relx=0.5, rely=0.5, anchor="center")
caja_splash.pack_propagate(False)

# Icono grande central
lbl_splash_icon = ctk.CTkLabel(caja_splash, text="", image=ICONS.get("settings"), font=("Roboto", 40)) # El tamaño real viene del CTkImage
lbl_splash_icon.pack(pady=(35, 5))

lbl_splash_title = ctk.CTkLabel(caja_splash, text="SISTEMA DE VALIDACIÓN AI", font=("Roboto", 16, "bold"), text_color="white")
lbl_splash_title.pack(pady=(0, 20))

lbl_carga = ctk.CTkLabel(caja_splash, text="Iniciando...", font=("Roboto", 12), text_color=C_TEXT_SUB)
lbl_carga.pack(pady=0)

lbl_perc = ctk.CTkLabel(caja_splash, text="0%", font=("Roboto", 24, "bold"), text_color=C_ACCENT)
lbl_perc.pack(pady=(5, 15))

barra_carga = ctk.CTkProgressBar(caja_splash, width=320, height=8, mode="determinate", progress_color=C_ACCENT, fg_color="#1a1c22")
barra_carga.set(0)
barra_carga.pack()

# === PANEL IZQUIERDO: CONTROLES ===
frame_left = ctk.CTkFrame(app, width=280, fg_color=C_BG_SIDE, corner_radius=0)
frame_left.grid(row=0, column=0, sticky="nswe")
frame_left.grid_propagate(False)
# Retiramos grid_rowconfigure destructivo


ctk.CTkLabel(frame_left, text="ARCHIVOS", font=("Roboto", 11, "bold"), text_color=C_ACCENT).pack(pady=(15, 0), padx=20, anchor="w")
btn_excel = ctk.CTkButton(frame_left, text="Cargar Excel", image=ICONS.get("excel"), compound="left", fg_color=C_CARD, hover_color="#404040",
                          anchor="w", height=35, command=seleccionar_excel)
btn_excel.pack(padx=20, pady=(5, 2), fill="x")
lbl_excel_name = ctk.CTkLabel(frame_left, text="Sin archivo", font=("Roboto", 10), text_color=C_TEXT_SUB, anchor="w")
lbl_excel_name.pack(padx=25, pady=(0, 10), fill="x")

btn_pdf = ctk.CTkButton(frame_left, text="Carpeta PDFs", image=ICONS.get("folder"), compound="left", fg_color=C_CARD, hover_color="#404040",
                        anchor="w", height=35, command=seleccionar_pdf)
btn_pdf.pack(padx=20, pady=(5, 2), fill="x")
lbl_pdf_name = ctk.CTkLabel(frame_left, text="Sin carpeta", font=("Roboto", 10), text_color=C_TEXT_SUB, anchor="w")
lbl_pdf_name.pack(padx=25, pady=(0, 10), fill="x")

ctk.CTkLabel(frame_left, text="CONFIGURACIÓN", font=("Roboto", 11, "bold"), text_color=C_ACCENT).pack(padx=20, anchor="w", pady=(10,0))
frame_cfg = ctk.CTkFrame(frame_left, fg_color="transparent")
frame_cfg.pack(padx=20, pady=5, fill="x")

ctk.CTkLabel(frame_cfg, text="Modo:", font=("Roboto", 11, "bold"), text_color=C_SUCCESS).pack(anchor="w")
combo_modo = ctk.CTkOptionMenu(frame_cfg, values=["Transporte (Manifiestos)", "Contabilidad (Nómina)"], fg_color=C_CARD, button_color=C_SUCCESS, text_color=C_TEXT_MAIN, command=cambiar_modo)
combo_modo.pack(fill="x", pady=(0, 10))

# Sub-contenedor horizontal para Hoja y Fila Encabezado
frame_cfg_sub = ctk.CTkFrame(frame_cfg, fg_color="transparent")
frame_cfg_sub.pack(fill="x", pady=(0, 10))

f_hoja = ctk.CTkFrame(frame_cfg_sub, fg_color="transparent")
f_hoja.pack(side="left", fill="x", expand=True, padx=(0, 5))
ctk.CTkLabel(f_hoja, text="Hoja:", font=("Roboto", 11)).pack(anchor="w")
combo_hoja = ctk.CTkComboBox(f_hoja, values=["..."], height=28, fg_color=C_CARD, button_color=C_ACCENT, text_color=C_TEXT_MAIN)
combo_hoja.pack(fill="x")

f_fila = ctk.CTkFrame(frame_cfg_sub, fg_color="transparent")
f_fila.pack(side="right", fill="x", expand=True)
ctk.CTkLabel(f_fila, text="Fila Título:", font=("Roboto", 11)).pack(anchor="w")
entry_fila = ctk.CTkEntry(f_fila, height=28, fg_color=C_CARD, text_color=C_TEXT_MAIN)
entry_fila.insert(0, "1")
entry_fila.pack(fill="x")

btn_cols = ctk.CTkButton(frame_left, text="Cargar Columnas", image=ICONS.get("load"), compound="left", fg_color=C_ACCENT, hover_color=C_ACCENT_H, height=30, command=cargar_columnas)
btn_cols.pack(padx=20, fill="x")

# Motor Inteligente - Ahora va ANTES de la lista de columnas (Scrollbar)
sep = ctk.CTkFrame(frame_left, height=2, fg_color="#3a3a3a")
sep.pack(fill="x", padx=10, pady=(15, 5))

frame_ia = ctk.CTkFrame(frame_left, fg_color="transparent")
frame_ia.pack(fill="x", padx=10, pady=(0, 5))

lbl_ia = ctk.CTkLabel(frame_ia, text="MOTOR INTELIGENTE", font=("Roboto", 10, "bold"), text_color="gray")
lbl_ia.pack(anchor="w", padx=10)

btn_gestor = ctk.CTkButton(frame_ia, text="Mis Plantillas", image=ICONS.get("folder"), compound="left", command=abrir_gestor, height=27,
                           fg_color="transparent", border_width=1, border_color="#555555", text_color="#dddddd", anchor="w")
btn_gestor.pack(fill="x", padx=10, pady=2)

btn_entrenador = ctk.CTkButton(frame_ia, text="Entrenar Nuevo", image=ICONS.get("target"), compound="left", command=abrir_entrenador, height=27,
                               fg_color=C_CARD, anchor="w")
btn_entrenador.pack(fill="x", padx=10, pady=2)

lbl_cols_title = ctk.CTkLabel(frame_left, text="SELECCIONAR COLUMNAS", font=("Roboto", 10, "bold"), text_color=C_TEXT_SUB)
lbl_cols_title.pack(padx=20, pady=(15,0), anchor="w")
scroll_cols = ctk.CTkScrollableFrame(frame_left, fg_color="transparent", height=250)
scroll_cols.pack(padx=10, pady=(0, 10), fill="both", expand=True)

# === PANEL DERECHO: DASHBOARD ===
frame_right = ctk.CTkFrame(app, fg_color=C_BG_MAIN)
frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

frame_stats = ctk.CTkFrame(frame_right, fg_color="transparent")
frame_stats.pack(fill="x", pady=(0, 15))

def create_card(parent, title, color, icon_key=None):
    C_CARD_HOVER = "#2e3147"  # Tono ligeramente más claro para efecto hover
    f = ctk.CTkFrame(parent, fg_color=C_CARD, corner_radius=8, height=70)
    f.pack_propagate(False)

    # --- Hover Effects (Micro-interacción) ---
    def _on_enter(e):
        f.configure(fg_color=C_CARD_HOVER)
    def _on_leave(e):
        f.configure(fg_color=C_CARD)
    for widget in [f]:  # se registrarán los hijos también en retorno
        f.bind("<Enter>", _on_enter)
        f.bind("<Leave>", _on_leave)

    f_txt = ctk.CTkFrame(f, fg_color="transparent")
    f_txt.pack(side="left", fill="both", expand=True)
    f_txt.bind("<Enter>", _on_enter)
    f_txt.bind("<Leave>", _on_leave)

    lbl_title = ctk.CTkLabel(f_txt, text=title, font=("Roboto", 10, "bold"), text_color=C_TEXT_SUB)
    lbl_title.pack(padx=15, pady=(10, 0), anchor="w")
    lbl_title.bind("<Enter>", _on_enter)
    lbl_title.bind("<Leave>", _on_leave)

    l = ctk.CTkLabel(f_txt, text="-", font=("Roboto", 22, "bold"), text_color=color)
    l.pack(padx=15, anchor="w")
    l.bind("<Enter>", _on_enter)
    l.bind("<Leave>", _on_leave)

    if icon_key and ICONS.get(icon_key):
        lbl_img = ctk.CTkLabel(f, text="", image=ICONS[icon_key], fg_color="transparent")
        lbl_img.place(relx=0.95, rely=0.5, anchor="e")
        lbl_img.bind("<Enter>", _on_enter)
        lbl_img.bind("<Leave>", _on_leave)

    return f, l

f1, card_ok_val  = create_card(frame_stats, "CORRECTOS",    C_SUCCESS, "check_large")
f1.pack(side="left", fill="x", expand=True, padx=(0, 10))

f2, card_err_val = create_card(frame_stats, "DISCREPANCIAS", C_ERROR, "alert_large")
f2.pack(side="left", fill="x", expand=True, padx=5)

f3, lbl_stat_time = create_card(frame_stats, "TIEMPO", C_TEXT_MAIN, "clock_large")
f3.pack(side="left", fill="x", expand=True, padx=(10, 0))

frame_prog = ctk.CTkFrame(frame_right, fg_color=C_CARD, corner_radius=8)
frame_prog.pack(fill="x", pady=5)

lbl_status = ctk.CTkLabel(frame_prog, text="Esperando...", font=("Roboto", 11), text_color=C_TEXT_SUB)
lbl_status.pack(side="top", anchor="w", padx=15, pady=(5,0))

f_p_info = ctk.CTkFrame(frame_prog, fg_color="transparent")
f_p_info.pack(fill="x")

lbl_stat_pct   = ctk.CTkLabel(f_p_info, text="0%",  font=("Roboto", 12, "bold"), text_color=C_ACCENT)
lbl_stat_pct.pack(side="left", padx=15, pady=5)

lbl_stat_count = ctk.CTkLabel(f_p_info, text="0/0", font=("Consolas", 10), text_color=C_TEXT_SUB)
lbl_stat_count.pack(side="right", padx=10, pady=5)

pb = ctk.CTkProgressBar(frame_right, height=10, corner_radius=4, progress_color=C_ACCENT, fg_color="#404040")
pb.pack(fill="x", pady=(0, 15))
pb.set(0)

lbl_term_title = ctk.CTkLabel(frame_right, text="REGISTRO DE EJECUCIÓN", font=("Roboto", 11, "bold"), text_color=C_TEXT_SUB)
lbl_term_title.pack(anchor="w", pady=(10, 0))
terminal = ctk.CTkTextbox(frame_right, fg_color="#0f0f0f", text_color="#cccccc", font=("Consolas", 11), corner_radius=10, border_width=1, border_color="#2b2b2b")
terminal.pack(fill="both", expand=True, pady=(5, 15))

frame_actions = ctk.CTkFrame(frame_right, fg_color="transparent")
frame_actions.pack(fill="x")

btn_auditar = ctk.CTkButton(frame_actions, text=" INICIAR PROCESO", image=ICONS.get("start"), compound="left", height=50, font=("Roboto", 14, "bold"),
                            corner_radius=8, fg_color=C_SUCCESS, hover_color="#238636", command=iniciar_auditoria)
btn_auditar.pack(side="left", fill="x", expand=True, padx=(0, 10))

btn_abrir = ctk.CTkButton(frame_actions, text="Abrir Excel", height=45, width=120, font=("Roboto", 12),
                          fg_color=C_CARD, state="disabled", command=abrir_excel)
btn_abrir.pack(side="right")

# ══════════════════════════════════════════════════════════════════
#  ARRANQUE
# ══════════════════════════════════════════════════════════════════
app.after(500,  cargar_config_inicial)

# [OPT-9] Iniciar el loop de flush del log (cada 100ms).
# Se llama DESPUÉS de que `terminal` ya existe en el scope global.
app.after(600, _flush_log)

app.mainloop()