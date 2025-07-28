from docx import Document
from docx.shared import Inches
from io import BytesIO
from PIL import Image

EXTENSOES_VALIDAS = {
    "jpg": "JPEG", "jpeg": "JPEG", "png": "PNG",
    "gif": "GIF", "bmp": "BMP", "tiff": "TIFF", "webp": "WEBP"
}

def criar_docx_com_formatacao_formatada(slides_formatados, imagens_por_slide, caminho_docx):
    doc = Document()
    for indice, slide in enumerate(slides_formatados):
        for bloco in slide:
            if bloco['tipo'] == 'titulo':
                doc.add_paragraph(bloco['texto'])
            elif bloco['tipo'] == 'paragrafo':
                par = doc.add_paragraph(style='List Bullet' if bloco.get('bullet') else None)
                for run in bloco['runs']:
                    r = par.add_run(run['texto'])
                    if run.get('bold'):
                        r.bold = True

        # Inserir imagens do slide atual, se houver
        if imagens_por_slide and indice < len(imagens_por_slide):
            for img in imagens_por_slide[indice]:
                try:
                    extensao = img['ext'].lower()
                    formato = EXTENSOES_VALIDAS.get(extensao, "PNG")

                    image = Image.open(BytesIO(img['dados']))
                    buffer = BytesIO()
                    image.save(buffer, format=formato)
                    buffer.seek(0)

                    doc.add_picture(buffer, width=Inches(4.0))
                except Exception as e:
                    print(f"Erro ao inserir imagem do slide {indice+1}: {e}")
                    continue

        doc.add_paragraph()

    doc.save(caminho_docx)