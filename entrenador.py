import tkinter as tk
from tkinter import filedialog
from CTkMessagebox import CTkMessagebox
import customtkinter as ctk
import pdfplumber
from PIL import ImageTk
import os
import json
from PIL import Image
from config import RUTAS

class MapeadorPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("Creador de Plantillas PDF - Multiusos")
        
        self.rect = None
        self.start_x = None
        self.start_y = None
        self.tk_img = None
        self.datos_plantilla = {} 
        self.pdf = None 
        
        self.escala_x = 1.0
        self.escala_y = 1.0
        self.pagina_actual = 0
        self.total_paginas = 0
        
        # --- CARGA GLOBALES DE ÍCONOS (USANDO CTkImage) ---
        self.ICONS = {}
        try:
            # Reutilizamos la misma estrategia que en auditorPDF.py
            _base_dir = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
            _icons_dir = os.path.join(_base_dir, "assets", "icons")
            self.ICONS["pdf"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "pdf.png")), size=(18, 18))
            self.ICONS["prev"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "prev.png")), size=(16, 16))
            self.ICONS["next"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "next.png")), size=(16, 16))
            self.ICONS["save"] = ctk.CTkImage(Image.open(os.path.join(_icons_dir, "save.png")), size=(18, 18))
        except Exception as e:
            print(f"No se pudieron cargar algunos íconos en Entrenador: {e}")

        # --- PANEL SUPERIOR ---
        panel = ctk.CTkFrame(root, fg_color="#1e212b", corner_radius=0)
        panel.pack(fill="x")
        
        # NUEVO: Botón para abrir explorador de archivos
        btn_abrir = ctk.CTkButton(panel, text="Abrir PDF", image=self.ICONS.get("pdf"), compound="left", font=("Roboto", 12, "bold"), fg_color="#3b82f6", hover_color="#144870", command=self.abrir_pdf)
        btn_abrir.pack(side="left", padx=10, pady=10)
        
        self.btn_prev = ctk.CTkButton(panel, text="Anterior", image=self.ICONS.get("prev"), compound="left", font=("Roboto", 12), fg_color="#252836", hover_color="#303446", command=self.pagina_anterior, state="disabled")
        self.btn_prev.pack(side="left", padx=10, pady=10)
        
        self.lbl_pag = ctk.CTkLabel(panel, text="Página 0 de 0", font=("Roboto", 12, "bold"), text_color="white")
        self.lbl_pag.pack(side="left", padx=10, pady=10)
        
        self.btn_next = ctk.CTkButton(panel, text="Siguiente", image=self.ICONS.get("next"), compound="right", font=("Roboto", 12), fg_color="#252836", hover_color="#303446", command=self.pagina_siguiente, state="disabled")
        self.btn_next.pack(side="left", padx=10, pady=10)
        
        btn_guardar = ctk.CTkButton(panel, text="Guardar Plantilla", image=self.ICONS.get("save"), compound="left", font=("Roboto", 12, "bold"), fg_color="#2ea043", hover_color="#238636", command=self.guardar_plantilla)
        btn_guardar.pack(side="right", padx=10, pady=10)
        
        # --- CANVAS CON SCROLLBARS ---
        frame_canvas = ctk.CTkFrame(root, fg_color="#0f111a", corner_radius=0)
        frame_canvas.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(frame_canvas, cursor="cross", bg="#0f111a", highlightthickness=0)
        
        vbar = ctk.CTkScrollbar(frame_canvas, orientation="vertical", command=self.canvas.yview)
        hbar = ctk.CTkScrollbar(frame_canvas, orientation="horizontal", command=self.canvas.xview)
        self.canvas.config(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        
        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.root.protocol("WM_DELETE_WINDOW", self.cerrar_seguro)
        
    def cerrar_seguro(self):
        """Libera la memoria RAM antes de cerrar la ventana"""
        if self.pdf:
            self.pdf.close() # Cierra el archivo en disco
            print("Memoria liberada: PDF cerrado.")
        self.root.destroy()
    
    # NUEVO: Función para buscar y cargar el archivo
    def abrir_pdf(self):
        ruta = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("Archivos PDF", "*.pdf")])
        if not ruta: return
        
        if self.pdf:
            self.pdf.close()
            self.tk_img = None # Borramos la imagen anterior de la memoria
            
        try:
            self.pdf = pdfplumber.open(ruta)
            self.total_paginas = len(self.pdf.pages)
            self.datos_plantilla = {} # Limpia plantillas anteriores
            self.cargar_pagina(0)
        except Exception as e:
            CTkMessagebox(title="Error", message=f"No se pudo abrir el PDF: {e}", icon="cancel")

    def cargar_pagina(self, num_pag):
        if not self.pdf: return
        self.pagina_actual = num_pag
        self.pagina = self.pdf.pages[num_pag]
        
        self.img_pil = self.pagina.to_image(resolution=150).original
        self.tk_img = ImageTk.PhotoImage(self.img_pil)
        
        self.escala_x = float(self.pagina.width) / float(self.img_pil.width)
        self.escala_y = float(self.pagina.height) / float(self.img_pil.height)
        
        self.canvas.config(scrollregion=(0, 0, self.tk_img.width(), self.tk_img.height()))
        self.canvas.delete("all") 
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        
        self.lbl_pag.configure(text=f"Página {self.pagina_actual + 1} de {self.total_paginas}")
        self.btn_prev.configure(state="normal" if self.pagina_actual > 0 else "disabled")
        self.btn_next.configure(state="normal" if self.pagina_actual < self.total_paginas - 1 else "disabled")

    def pagina_anterior(self):
        if self.pagina_actual > 0: self.cargar_pagina(self.pagina_actual - 1)

    def pagina_siguiente(self):
        if self.pagina_actual < self.total_paginas - 1: self.cargar_pagina(self.pagina_actual + 1)

    def on_press(self, event):
        if not self.pdf: return
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.rect: self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, 1, 1, outline="red", width=2)

    def on_drag(self, event):
        if not self.pdf or not self.rect: return
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_release(self, event):
        if not self.pdf or not self.rect: return
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        
        x0_screen, x1_screen = sorted([self.start_x, cur_x])
        top_screen, bottom_screen = sorted([self.start_y, cur_y])
        
        x0 = max(0.0, min(x0_screen * self.escala_x, float(self.pagina.width)))
        x1 = max(0.0, min(x1_screen * self.escala_x, float(self.pagina.width)))
        top = max(0.0, min(top_screen * self.escala_y, float(self.pagina.height)))
        bottom = max(0.0, min(bottom_screen * self.escala_y, float(self.pagina.height)))
        
        if x1 == float(self.pagina.width): x1 -= 0.1
        if bottom == float(self.pagina.height): bottom -= 0.1
        
        bbox = (x0, top, x1, bottom)
        if x1 - x0 < 2 or bottom - top < 2: return
            
        try:
            texto = self.pagina.crop(bbox).extract_text()
            if not texto: 
                texto = "[Zona vacía o no legible]"
            else:
                import re
                texto = texto.strip()
                # Detecta números separados por basura (espacios, comillas, puntos)
                # Ej: Convierte "128275 ' 06" -> "128275-06"
                if re.match(r"^\d+[-\s'_.]+\d+$", texto):
                    texto = re.sub(r"[-\s'_.]+", "-", texto)
            
            dialog = ctk.CTkInputDialog(text=f"Texto extraído: '{texto}'\n\nEscribe el nombre de la columna (ej. Manifiesto, Cantidad):", title="Nuevo Campo")
            campo = dialog.get_input()
            
            if campo:
                self.datos_plantilla[campo.lower()] = {"pagina": self.pagina_actual, "coordenadas": bbox}
                
        except Exception as e:
            CTkMessagebox(title="Error", message=f"No se pudo extraer texto. Error: {e}", icon="cancel")

    def guardar_plantilla(self):
        if not self.datos_plantilla:
            CTkMessagebox(title="Aviso", message="No has mapeado ningún campo todavía.", icon="warning")
            return
            
        dialog = ctk.CTkInputDialog(text="Nombre de la plantilla (ej. FORMATO_A):", title="Guardar")
        nombre = dialog.get_input()
        if not nombre: return
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
        # --- CORRECCIÓN DE RUTA: USAR SIEMPRE APPDATA ---
        from config import APP_NAME
        import os
        import json
        
        ruta_json = RUTAS["json"] # Usamos la ruta unificada desde config.py
        # ------------------------------------------------
        
        
        datos = {}
        # Intentamos leer lo que ya existe para no borrar lo anterior
        if os.path.exists(ruta_json):
            try:
                with open(ruta_json, "r", encoding="utf-8") as f:
                    datos = json.load(f)
            except:
                pass

        datos[nombre.upper()] = self.datos_plantilla
        
        try:
            with open(ruta_json, "w", encoding="utf-8") as f:
                json.dump(datos, f, indent=4)
            
            CTkMessagebox(title="Éxito", message=f"Plantilla '{nombre}' guardada correctamente.", icon="check")
            
            # 2. AUTO-CIERRE: Cerramos el PDF y la ventana automáticamente tras guardar
            self.cerrar_seguro() 

        except Exception as e:
            CTkMessagebox(title="Error Crítico", message=f"Error de guardado: {e}", icon="cancel")
            

if __name__ == "__main__":
    import sys
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("dark-blue")
    root = ctk.CTk()
    root.geometry("800x600") 
    app = MapeadorPDF(root)
    root.mainloop()