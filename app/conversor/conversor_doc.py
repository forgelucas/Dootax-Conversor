import os
import pathlib
import random
import win32com.client
from conversor.utils import gerar_nome_unico
from conversor.notificacoes import notificar_sucesso

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
            novo_caminho = os.path.join(pasta_destino, f"{nome_arquivo_base}.docx")

            if os.path.exists(novo_caminho):
                novo_caminho = gerar_nome_unico(pasta_destino, nome_arquivo_base)

            doc.SaveAs(novo_caminho, FileFormat=16)
            doc.Close()

            notificar_sucesso(novo_caminho)  # <-- agora dentro do try

        except Exception as e:
            print(f"Erro ao converter {caminho}: {e}")

    word.Quit()
