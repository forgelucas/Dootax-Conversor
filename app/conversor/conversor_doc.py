import pathlib
import win32com.client
from app.conversor.notificacoes import notificar_sucesso, notificar_erro

def converter_doc_para_docx(caminho_origem, caminho_destino):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        caminho_resolvido = str(pathlib.Path(caminho_origem).resolve())
        doc = word.Documents.Open(caminho_resolvido)
        doc.SaveAs(caminho_destino, FileFormat=16)
        doc.Close()
        notificar_sucesso(caminho_destino)
    except Exception as e:
        erro_msg = f"Erro ao converter o arquivo:\n{caminho_origem}\n\nDetalhes: {str(e)}"
        print(erro_msg)
        notificar_erro(erro_msg)
    finally:
        word.Quit()