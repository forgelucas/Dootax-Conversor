import os
import random
import re

def gerar_nome_unico(pasta_destino, nome_arquivo_base):
    while True:
        sufixo = f"_conversorcopy{random.randint(1000, 9999)}"
        novo_nome = f"{nome_arquivo_base}{sufixo}.docx"
        novo_caminho = os.path.join(pasta_destino, novo_nome)
        if not os.path.exists(novo_caminho):
            return novo_caminho

def clean_text(texto: str) -> str:
    return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', texto)

def centralizar_janela(janela):
    janela.update_idletasks()
    largura = janela.winfo_width()
    altura = janela.winfo_height()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
