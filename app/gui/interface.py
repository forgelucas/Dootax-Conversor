import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from conversor.conversor_doc import converter_docs_para_docx
from conversor.conversor_pptx import converter_pptx_para_docx
from conversor.utils import centralizar_janela

def selecionar_docs():
    caminhos = filedialog.askopenfilenames(title="Selecione os arquivos .doc",
                                           filetypes=[("Arquivos DOC", "*.doc")])
    if caminhos:
        threading.Thread(target=converter_docs_para_docx, args=(caminhos,)).start()

def selecionar_pptx():
    caminhos = filedialog.askopenfilenames(title="Selecione os arquivos .pptx",
                                           filetypes=[("Apresentações PowerPoint", "*.pptx")])
    if caminhos:
        pasta_destino = os.path.join(os.getcwd(), "arquivos_convertidos")
        os.makedirs(pasta_destino, exist_ok=True)

        for caminho_pptx in caminhos:
            nome_base = os.path.splitext(os.path.basename(caminho_pptx))[0]
            caminho_docx = os.path.join(pasta_destino, f"{nome_base}.docx")
            threading.Thread(target=converter_pptx_para_docx, args=(caminho_pptx, caminho_docx)).start()

def iniciar_interface():
    janela = tk.Tk()
    janela.title("Conversor DOC/PPTX para DOCX")
    centralizar_janela(janela)
    janela.geometry("400x200")
    janela.resizable(False, False)

    label = tk.Label(janela, text="Selecione o tipo de conversão", font=("Arial", 14))
    label.pack(pady=20)

    tk.Button(janela, text="Converter DOC para DOCX", command=selecionar_docs, width=30, height=2).pack(pady=5)
    tk.Button(janela, text="Converter PPTX para DOCX", command=selecionar_pptx, width=30, height=2).pack(pady=5)

    janela.mainloop()
