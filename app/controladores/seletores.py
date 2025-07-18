from tkinter import filedialog
from app.conversor.conversor_pdf import converter_pdf_para_docx
from app.conversor.conversor_pptx import converter_pptx_para_docx
from app.conversor.conversor_doc import converter_doc_para_docx
from app.conversor.conversor_excel import converter_excel_para_docx
from markitdown import MarkItDown
import os


def selecionar_pdf():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    for caminho_pdf in caminhos:
        converter_pdf_para_docx(caminhos)

def selecionar_pptx():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos PPTX", filetypes=[("Apresentações", "*.pptx")])
    for caminho_pptx in caminhos:
        converter_pptx_para_docx(caminho_pptx)

def selecionar_docs():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos DOC", filetypes=[("Documentos", "*.doc")])
    converter_doc_para_docx(caminhos)

def selecionar_excel():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos Excel", filetypes=[("Excel (.xlsx e .xls)", "*.xlsx *.xls")])
    for caminho_xlsx in caminhos:
        converter_excel_para_docx(caminhos)

def selecionar_para_markdown():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos para Markdown")
    md = MarkItDown()
    for caminho in caminhos:
        result = md.convert(caminho)
        nome = os.path.splitext(os.path.basename(caminho))[0] + ".md"
        pasta_destino = os.path.join(os.getcwd(), "arquivos_convertidos")
        os.makedirs(pasta_destino, exist_ok=True)
        caminho_final = os.path.join(pasta_destino, nome)
        with open(caminho_final, "w", encoding="utf-8") as f:
            f.write(result.text_content)
