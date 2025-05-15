import os
from conversor.pptx_utils import extrair_texto_formatado_do_pptx, extrair_texto_imagens_pptx
from conversor.docx_helpers import criar_docx_com_formatacao_formatada
from conversor.notificacoes import notificar_sucesso, notificar_erro

def converter_pptx_para_docx(caminho_pptx, caminho_docx):
    try:
        pasta_destino_imagens = os.path.join(os.getcwd(), "imagens_extraidas")

        slides_formatados = extrair_texto_formatado_do_pptx(caminho_pptx)
        extrair_texto_imagens_pptx(caminho_pptx, pasta_destino_imagens)
        criar_docx_com_formatacao_formatada(slides_formatados, pasta_destino_imagens, caminho_docx)

        notificar_sucesso(caminho_docx)

    except Exception as e:
        erro_msg = f"Erro ao converter o arquivo PPTX:\n{caminho_pptx}\n\nDetalhes: {str(e)}"
        print(erro_msg)
        notificar_erro(erro_msg)
