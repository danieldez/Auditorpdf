import os

# 1. Update auditorPDF.py
with open("auditorPDF.py", "r", encoding="utf-8") as f:
    text = f.read()

# Colors
text = text.replace('C_BG_MAIN    = ("#f5f6fa", "#1a1a1a")', 'C_BG_MAIN    = ("#f5f6fa", "#0f111a")')
text = text.replace('C_BG_SIDE    = ("#ffffff", "#2b2b2b")', 'C_BG_SIDE    = ("#ffffff", "#1e212b")')
text = text.replace('C_CARD       = ("#e1e5eb", "#333333")', 'C_CARD       = ("#e1e5eb", "#252836")')
text = text.replace('C_ACCENT     = ("#1f6aa5", "#1f6aa5")', 'C_ACCENT     = ("#1f6aa5", "#3b82f6")')
text = text.replace('"info": "#1f6aa5"', '"info": "#3b82f6"')  # update log accent color

# Fonts
text = text.replace('"Segoe UI"', '"Roboto"').replace("'Segoe UI'", '"Roboto"')
text = text.replace('"Arial"', '"Roboto"').replace("'Arial'", '"Roboto"')

# Terminal Widget
text = text.replace('terminal.configure(bg="#0f0f0f", fg="#cccccc")', 'terminal.configure(fg_color="#0f0f0f", text_color="#cccccc")')
text = text.replace('terminal = tk.Text(frame_right, bg="#0f0f0f", fg="#cccccc", font=("Consolas", 10), relief="flat", padx=10, pady=10)',
                    'terminal = ctk.CTkTextbox(frame_right, fg_color="#0f0f0f", text_color="#cccccc", font=("Consolas", 10), corner_radius=8)')

with open("auditorPDF.py", "w", encoding="utf-8") as f:
    f.write(text)


# 2. Update entrenador.py
with open("entrenador.py", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace('fg_color="#2b2b2b"', 'fg_color="#1e212b"')
text = text.replace('fg_color="#1a1a1a"', 'fg_color="#0f111a"')
text = text.replace('bg="#1a1a1a"', 'bg="#0f111a"')  # For canvas
text = text.replace('fg_color="#3a3a3a"', 'fg_color="#252836"')
text = text.replace('hover_color="#404040"', 'hover_color="#303446"')
text = text.replace('fg_color="#1f6aa5"', 'fg_color="#3b82f6"') # accent

with open("entrenador.py", "w", encoding="utf-8") as f:
    f.write(text)


# 3. Update gestor.py
with open("gestor.py", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace('fg_color="#1a1a1a"', 'fg_color="#0f111a"')
text = text.replace('fg_color="#2b2b2b"', 'fg_color="#1e212b"')
text = text.replace('fg_color="#3a3a3a"', 'fg_color="#252836"')
text = text.replace('"Arial"', '"Roboto"').replace("'Arial'", '"Roboto"')

with open("gestor.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Actualización de colores y tipografía completada.")
