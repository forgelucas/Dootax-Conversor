import os
import random
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import win32com.client
import pathlib

def gerar_nome_unico(pasta_destino, nome_arquivo_base):
    while True:
        sufixo = f"_conversorcopy{random.randint(1000, 9999)}"
        novo_nome = f"{nome_arquivo_base}{sufixo}.docx"
        novo_caminho = os.path.join(pasta_destino, novo_nome)
        if not os.path.exists(novo_caminho):
            return novo_caminho

def converter_docs_para_docx(caminhos):
    pasta_destino = os.path.join(os.getcwd(), "arquivos_convertidos")
    os.makedirs(pasta_destino, exist_ok=True)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    for caminho in caminhos:
        try:
            caminho_resolvido = str(pathlib.Path(caminho).resolve())
            doc = word.Documents.Open(caminho_resolvido)

            nome_arquivo_base = os.path.splitext(os.path.basename(caminho))[0]
            novo_nome = f"{nome_arquivo_base}.docx"
            novo_caminho = os.path.join(pasta_destino, novo_nome)

            # Se o nome já existir, gera um novo com número aleatório garantido único
            if os.path.exists(novo_caminho):
                novo_caminho = gerar_nome_unico(pasta_destino, nome_arquivo_base)

            doc.SaveAs(novo_caminho, FileFormat=16)
            doc.Close()
        except Exception as e:
            print(f"Erro ao converter {caminho}: {e}")

    word.Quit()

# Função com a lógica para conversão com barra de progresso
def processar_conversao(caminhos):
    progress_bar.start()
    btn_selecionar.config(state=tk.DISABLED)

    try:
        converter_docs_para_docx(caminhos)
        messagebox.showinfo("Sucesso", f"Conversão concluída!\nArquivos salvos em 'arquivos_convertidos'.")
        pasta_destino = os.path.join(os.getcwd(), "arquivos_convertidos")
        os.startfile(pasta_destino)
    finally:
        progress_bar.stop()
        btn_selecionar.config(state=tk.NORMAL)

# Função disparada pelo botão
def selecionar_arquivos():
    caminhos = filedialog.askopenfilenames(
        title="Selecione os arquivos .doc",
        filetypes=[("Arquivos DOC", "*.doc")]
    )
    if caminhos:
        threading.Thread(target=processar_conversao, args=(caminhos,), daemon=True).start()

# Interface gráfica
def centralizar_janela(janela):
    janela_largura = janela.winfo_width()
    janela_altura = janela.winfo_height()

    monitor_largura = janela.winfo_screenwidth()
    monitor_altura = janela.winfo_screenheight()

    x = (monitor_largura // 2) - (janela_largura // 2)
    y = (monitor_altura // 2) - (janela_altura // 2)

    janela.geometry(f"{janela_largura}x{janela_altura}+{x}+{y}")


janela = tk.Tk()
janela.title("Conversor DOC para DOCX")
centralizar_janela(janela)
janela.geometry("400x180")

label = tk.Label(janela, text="Selecione os arquivos .doc para converter", font=("Arial", 12))
label.pack(pady=10)

btn_selecionar = tk.Button(janela, text="Selecionar Arquivos", command=selecionar_arquivos, height=2, width=20)
btn_selecionar.pack(pady=10)

progress_bar = ttk.Progressbar(janela, mode="indeterminate", length=300)
progress_bar.pack(pady=10)

janela.mainloop()
