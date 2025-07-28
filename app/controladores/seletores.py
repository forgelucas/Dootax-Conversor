from tkinter import filedialog
from app.conversor.conversor_pdf import converter_pdf_para_docx
from app.conversor.conversor_pptx import converter_pptx_para_docx
from app.conversor.conversor_doc import converter_doc_para_docx
from app.conversor.conversor_excel import converter_excel_para_docx
from markitdown import MarkItDown

def selecionar_pdf():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    for caminho_pdf in caminhos:
        caminho_destino = filedialog.asksaveasfilename(
            title="Salvar como DOCX",
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if caminho_destino:
            converter_pdf_para_docx(caminho_pdf, caminho_destino)

def selecionar_pptx():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos PPTX", filetypes=[("Apresentações", "*.pptx")])
    for caminho_pptx in caminhos:
        caminho_destino = filedialog.asksaveasfilename(
            title="Salvar como DOCX",
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if caminho_destino:
            converter_pptx_para_docx(caminho_pptx, caminho_destino)

def selecionar_docs():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos DOC", filetypes=[("Documentos", "*.doc")])
    for caminho_doc in caminhos:
        caminho_destino = filedialog.asksaveasfilename(
            title="Salvar como DOCX",
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if caminho_destino:
            converter_doc_para_docx(caminho_doc, caminho_destino)

def selecionar_excel():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos Excel", filetypes=[("Excel (.xlsx e .xls)", "*.xlsx *.xls")])
    for caminho_xlsx in caminhos:
        caminho_destino = filedialog.asksaveasfilename(
            title="Salvar como DOCX",
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if caminho_destino:
            converter_excel_para_docx(caminho_xlsx, caminho_destino)

def selecionar_para_markdown():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos para Markdown")
    md = MarkItDown()
    for caminho in caminhos:
        result = md.convert(caminho)
        caminho_destino = filedialog.asksaveasfilename(
            title="Salvar como Markdown",
            defaultextension=".md",
            filetypes=[("Arquivo Markdown", "*.md")]
        )
        if caminho_destino:
            with open(caminho_destino, "w", encoding="utf-8") as f:
                f.write(result.text_content)