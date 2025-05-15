# conversor/notificacoes.py
import os
import sys

def notificar_sucesso(caminho_saida: str, titulo="Conversão concluída"):
    """
    Exibe uma mensagem de sucesso e, quando o utilizador confirmar,
    abre a pasta onde o ficheiro convertido foi salvo.
    """
    pasta = os.path.dirname(os.path.abspath(caminho_saida))

    # --- TKINTER: rápido e sem dependências ---
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()                       # esconde a janela principal
        messagebox.showinfo(titulo,
                            f"Arquivo convertido com sucesso!\n\n{caminho_saida}")
        root.destroy()
    except Exception:
        # --- Windows nativo (fallback) ---
        if sys.platform.startswith("win"):
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0,
                f"Arquivo convertido com sucesso!\n\n{caminho_saida}",
                titulo,
                0x40  # MB_ICONINFORMATION
            )
        else:
            # *Última linha de defesa*: imprime no terminal
            print(f"[INFO] {titulo}: {caminho_saida}")

    # Abre a pasta convertida (explorador/shell)
    os.startfile(pasta) if sys.platform.startswith("win") else os.system(f'xdg-open "{pasta}"')
