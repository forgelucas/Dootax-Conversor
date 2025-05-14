import os
import re
from pptx import Presentation

def extrair_texto_imagens_pptx(caminho_pptx, pasta_destino):
    prs = Presentation(caminho_pptx)
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    
    contador_imagens = 0
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if hasattr(shape, "image"):
                image = shape.image
                ext = image.ext
                path_img = os.path.join(pasta_destino, f"imagem_{i+1}_{contador_imagens+1}.{ext}")
                with open(path_img, "wb") as f:
                    f.write(image.blob)
                contador_imagens += 1

def extrair_texto_formatado_do_pptx(caminho_pptx):
    prs = Presentation(caminho_pptx)
    slides_formatados = []

    for idx, slide in enumerate(prs.slides, start=1):
        conteudo_slide = [{'tipo': 'titulo', 'texto': f"Slide {idx}:"}]
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for par in shape.text_frame.paragraphs:
                par_dict = {'tipo': 'paragrafo', 'runs': [], 'bullet': par.level > 0}
                for run in par.runs:
                    par_dict['runs'].append({'texto': run.text, 'bold': run.font.bold})
                conteudo_slide.append(par_dict)
        slides_formatados.append(conteudo_slide)

    return slides_formatados
