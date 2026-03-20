import customtkinter as ctk
from tkinter import messagebox
import json
import os
from config import RUTAS
import os
import sys
from config import RUTAS
from PIL import Image

class GestorPlantillas(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Mis Plantillas")
        self.geometry("400x500")
        self.resizable(False, False)
        
        # Color de fondo coherente con tu App
        self.configure(fg_color="#0f111a")
        
        # Hacerla modal y traer al frente
        self.transient(parent)
        self.grab_set()
        self.focus_force()

        # --- CARGA GLOBAL DEL ÍCONO ---
        try:
            _base_dir = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
            _icon_path = os.path.join(_base_dir, "assets", "icons", "pdf.png")
            self.icon_doc = ctk.CTkImage(Image.open(_icon_path), size=(18, 18))
        except Exception:
            self.icon_doc = None
        
        # --- HEADER ---
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=20, padx=20)
        
        ctk.CTkLabel(header, text="Gestor de Plantillas", font=("Roboto", 20, "bold"), text_color="white").pack(side="left")
        
        # --- LISTA SCROLLABLE ESTILIZADA ---
        self.frame_lista = ctk.CTkScrollableFrame(
            self, 
            width=360, 
            height=350, 
            fg_color="#1e212b",  # Fondo oscuro para la lista
            corner_radius=10
        )
        self.frame_lista.pack(padx=20, pady=(0, 20), fill="both", expand=True)
        
        self.cargar_lista()

    def cargar_lista(self):
        for widget in self.frame_lista.winfo_children():
            widget.destroy()
            
        datos = {}
        if os.path.exists(RUTAS["json"]):
            try:
                with open(RUTAS["json"], "r", encoding="utf-8") as f:
                    datos = json.load(f)
            except: pass
        
        if not datos:
            ctk.CTkLabel(self.frame_lista, text="\nNo hay plantillas guardadas.\nUsa el Entrenador para crear una.", text_color="gray").pack()
            return

        # Renderizar cada plantilla como una "Tarjeta"
        for nombre in sorted(datos.keys()):
            self.crear_tarjeta(nombre, datos[nombre])

    def crear_tarjeta(self, nombre, data):
        # Tarjeta contenedora
        card = ctk.CTkFrame(self.frame_lista, fg_color="#252836", corner_radius=8)
        card.pack(fill="x", pady=5, padx=5)
        
        # Columna Izquierda: Nombre y detalles
        info_frame = ctk.CTkFrame(card, fg_color="transparent")
        info_frame.pack(side="left", padx=10, pady=10)
        
        lbl_nombre = ctk.CTkLabel(info_frame, text=f" {nombre}", font=("Roboto", 14, "bold"), text_color="white")
        if getattr(self, "icon_doc", None):
            lbl_nombre.configure(image=self.icon_doc, compound="left")
        else:
            lbl_nombre.configure(text=f"📄 {nombre}")
            
        lbl_nombre.pack(anchor="w")
        
        # Mostrar cuántos campos tiene mapeados (Detalle visual útil)
        num_campos = len(data) if data else 0
        ctk.CTkLabel(info_frame, text=f"{num_campos} campos configurados", font=("Roboto", 10), text_color="gray").pack(anchor="w")
        
        # Columna Derecha: Botón Eliminar
        btn_borrar = ctk.CTkButton(
            card, 
            text="×", 
            width=30, 
            height=30,
            fg_color="#cf3a3a", # Rojo suave
            hover_color="#8a1c1c",
            font=("Roboto", 16, "bold"),
            command=lambda n=nombre: self.eliminar_plantilla(n)
        )
        btn_borrar.pack(side="right", padx=10)

    def eliminar_plantilla(self, nombre):
        # El parámetro parent=self es la CLAVE para que funcione en ventanas modales
        if messagebox.askyesno("Eliminar", f"¿Borrar plantilla '{nombre}'?", parent=self):
            try:
                from config import RUTAS
                import json
                
                # 1. Leer archivo actual
                with open(RUTAS["json"], "r", encoding="utf-8") as f:
                    datos = json.load(f)
                
                # 2. Borrar la plantilla del diccionario
                if nombre in datos:
                    del datos[nombre]
                
                # 3. Guardar el archivo actualizado en el disco duro
                with open(RUTAS["json"], "w", encoding="utf-8") as f:
                    json.dump(datos, f, indent=4)
                
                # 4. Recargar la interfaz visual
                self.cargar_lista()
                
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo borrar: {e}", parent=self)