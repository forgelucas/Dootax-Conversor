import os
from conversor.pptx_utils import extrair_texto_formatado_do_pptx, extrair_texto_imagens_pptx
from conversor.docx_helpers import criar_docx_com_formatacao_formatada
from conversor.notificacoes import notificar_sucesso

def converter_pptx_para_docx(caminho_pptx, caminho_docx):
    pasta_destino_imagens = os.path.join(os.getcwd(), "imagens_extraidas")
    
    slides_formatados = extrair_texto_formatado_do_pptx(caminho_pptx)
    extrair_texto_imagens_pptx(caminho_pptx, pasta_destino_imagens)
    criar_docx_com_formatacao_formatada(slides_formatados, pasta_destino_imagens, caminho_docx)
    notificar_sucesso(caminho_docx)

    print(f"Conversão concluída! Documento salvo em: {caminho_docx}")

