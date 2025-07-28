import re

def clean_text(texto: str) -> str:
    """Remove caracteres de controle invisíveis do texto."""
    return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', texto)

def centralizar_janela(janela):
    """Centraliza uma janela Tkinter na tela."""
    janela.update_idletasks()
    largura = janela.winfo_width()
    altura = janela.winfo_height()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")