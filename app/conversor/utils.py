import os
import re

def gerar_nome_unico(pasta_destino, nome_arquivo_base):
    padrao = re.compile(re.escape(nome_arquivo_base) + r'_conversorcopy(\d+)\.docx')
    maior_numero = 0

    for nome in os.listdir(pasta_destino):
        match = padrao.match(nome)
        if match:
            numero = int(match.group(1))
            if numero > maior_numero:
                maior_numero = numero

    novo_numero = maior_numero + 1
    novo_nome = f"{nome_arquivo_base}_conversorcopy{novo_numero}.docx"
    novo_caminho = os.path.join(pasta_destino, novo_nome)
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
