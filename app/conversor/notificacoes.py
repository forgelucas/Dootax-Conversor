# mypy: ignore-errors
import sys
import tkinter as tk
from tkinter import messagebox

def notificar_sucesso(caminho_saida: str, titulo="Conversão concluída"):
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(titulo, f"Arquivo convertido com sucesso!\n\n{caminho_saida}")
        root.destroy()
    except Exception:
        if sys.platform.startswith("win"):
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0,
                f"Arquivo convertido com sucesso!\n\n{caminho_saida}",
                titulo,
                0x40  
            )
        else:
            print(f"[INFO] {titulo}: {caminho_saida}")

def notificar_erro(mensagem: str, titulo="Erro na conversão"):
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(titulo, mensagem)
        root.destroy()
    except Exception:
        if sys.platform.startswith("win"):
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0,
                mensagem,
                titulo,
                0x10  
            )
        else:
            print(f"[ERRO] {titulo}: {mensagem}")