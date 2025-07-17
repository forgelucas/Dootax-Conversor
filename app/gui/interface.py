import tkinter as tk
from tkinter import ttk
import os
import sys
from PIL import Image, ImageTk
from app.conversor.utils import centralizar_janela
from app.controladores.seletores import selecionar_docs, selecionar_pptx, selecionar_excel, selecionar_pdf

import os
import sys
from PIL import Image, ImageTk

def resource_path(relative_path):
    """Resolve caminho relativo para uso no .exe ou em desenvolvimento"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)

def iniciar_interface():
    janela = tk.Tk()
    janela.title("Dootax Converter")
    janela.configure(bg="white")
    centralizar_janela(janela)
    janela.geometry("400x460")
    janela.resizable(False, False)

    estilo = ttk.Style()
    estilo.theme_use("default")
    estilo.configure("TButton",
                     font=("Segoe UI", 10, "bold"),
                     background="#B7E4C7",
                     foreground="black",
                     borderwidth=0)
    estilo.map("TButton", background=[('active', '#A7DCC2')])

    caminho_logo = resource_path("app/gui/logo_dootax.png")
    img_logo = Image.open(caminho_logo)
    img_logo = img_logo.resize((120, 120), Image.Resampling.LANCZOS)
    logo = ImageTk.PhotoImage(img_logo)
    label_logo = tk.Label(janela, image=logo, bg="white")
    label_logo.image = logo
    label_logo.pack(pady=(20, 5))

    label_titulo = tk.Label(janela, text="Conversor Dootax", font=("Segoe UI", 16, "bold"), bg="white", fg="#5A5A5A")
    label_titulo.pack(pady=(0, 15))

    frame_botoes = tk.Frame(janela, bg="white")
    frame_botoes.pack(pady=(0, 15))

    botoes = [
        ("Converter DOC para DOCX", selecionar_docs),
        ("Converter PPTX para DOCX", selecionar_pptx),
        ("Converter XLS para DOCX", selecionar_excel),
        ("Converter PDF para DOCX", selecionar_pdf),
    ]

    for i, (texto, comando) in enumerate(botoes):
        btn = ttk.Button(frame_botoes, text=texto, command=comando, width=30)
        btn.grid(row=i, column=0, padx=10, pady=5, ipady=4)

    janela.mainloop()