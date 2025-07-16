import tkinter as tk
from tkinter import filedialog
from app.conversor.conversor_pdf import converter_pdf_para_docx
from app.conversor.conversor_pptx import converter_pptx_para_docx
from app.conversor.conversor_doc import converter_doc_para_docx
from app.conversor.conversor_excel import converter_excel_para_docx

def selecionar_pdf():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    for caminho_pdf in caminhos:
        caminho_docx = caminho_pdf.replace(".pdf", ".docx")
        converter_pdf_para_docx(caminhos)

def selecionar_pptx():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos PPTX", filetypes=[("Apresentações", "*.pptx")])
    for caminho_pptx in caminhos:
        caminho_docx = caminho_pptx.replace(".pptx", ".docx")
        converter_pptx_para_docx(caminho_pptx)

def selecionar_docs():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos DOC", filetypes=[("Documentos", "*.doc")])
    converter_doc_para_docx(caminhos)

def selecionar_excel():
    caminhos = filedialog.askopenfilenames(title="Selecionar arquivos Excel", filetypes=[("Excel (.xlsx e .xls)", "*.xlsx *.xls")])
    for caminho_xlsx in caminhos:
        caminho_docx = caminho_xlsx.replace(".xlsx", ".docx")
        converter_excel_para_docx(caminhos)
