import os
import random
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import win32com.client

# Função para converter arquivos DOC para DOCX
def converter_docs_para_docx(caminhos):
    pasta_destino = os.path.join(os.getcwd(), "arquivos_convertidos")
    os.makedirs(pasta_destino, exist_ok=True)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    for caminho in caminhos:
        try:
            doc = word.Documents.Open(caminho)
            nome_arquivo = os.path.splitext(os.path.basename(caminho))[0]
            novo_nome = f"{nome_arquivo}.docx"
            novo_caminho = os.path.join(pasta_destino, novo_nome)

            contador = 1
            while os.path.exists(novo_caminho):
                novo_nome = f"{nome_arquivo}_conversorcopy{random.randint(1000,9999)}.docx"
                novo_caminho = os.path.join(pasta_destino, novo_nome)

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
janela = tk.Tk()
janela.title("Conversor DOC para DOCX")
janela.geometry("400x180")

label = tk.Label(janela, text="Selecione os arquivos .doc para converter", font=("Arial", 12))
label.pack(pady=10)

btn_selecionar = tk.Button(janela, text="Selecionar Arquivos", command=selecionar_arquivos, height=2, width=20)
btn_selecionar.pack(pady=10)

progress_bar = ttk.Progressbar(janela, mode="indeterminate", length=300)
progress_bar.pack(pady=10)

janela.mainloop()
