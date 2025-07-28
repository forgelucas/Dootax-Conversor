import customtkinter as ctk
from PIL import Image
from app.conversor.utils import centralizar_janela
from app.controladores.seletores import (
    selecionar_docs, selecionar_pptx,
    selecionar_excel, selecionar_pdf,
    selecionar_para_markdown
)
import os
import sys

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)

def iniciar_interface():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("green")

    janela = ctk.CTk()
    janela.title("Conversor Dootax")
    janela.geometry("420x480")
    janela.resizable(False, False)
    centralizar_janela(janela)

    caminho_logo = resource_path("logo_dootax.png")
    img_logo = Image.open(caminho_logo).resize((200, 70))
    logo = ctk.CTkImage(light_image=img_logo, size=(200, 70))
    label_logo = ctk.CTkLabel(janela, image=logo, text="")
    label_logo.pack(pady=(20, 10))

    label_logo.pack(pady=(20, 20))

    botoes = [
        ("Converter DOC para DOCX", selecionar_docs),
        ("Converter PPTX para DOCX", selecionar_pptx),
        ("Converter XLS para DOCX", selecionar_excel),
        ("Converter PDF para DOCX", selecionar_pdf),
        ("Converter para Markdown (.md)", selecionar_para_markdown)
    ]

    for texto, comando in botoes:
        btn = ctk.CTkButton(janela, text=texto, command=comando, width=280, height=40)
        btn.pack(pady=12)

    janela.mainloop()