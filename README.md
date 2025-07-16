📝 Dootax Conversor
Ferramenta em Python com interface gráfica para conversão de arquivos .doc, .pptx, .pdf e .xlsx para .docx. Ideal para analistas, contadores e desenvolvedores que precisam migrar documentações de forma rápida, limpa e automatizada.

🚀 Funcionalidades
- Conversão automática dos formatos .doc, .pptx, .pdf e .xlsx para .docx
- Interface gráfica amigável com botões organizados por tipo de arquivo
- Suporte à seleção de múltiplos arquivos simultaneamente
- Extração de texto com formatação e imagens (sem salvar imagens em disco)
- Renomeação automática caso arquivos tenham nomes repetidos
- Notificação visual ao término da conversão
- Abertura automática da pasta de arquivos convertidos
- Distribuído como executável .exe (dispensa instalação do Python)

📦 Como executar
✔️ Versão Executável (recomendado)
- Baixe o arquivo main.exe na pasta dist/ ou na aba de releases aqui no GitHub
- Dê dois cliques para executar
- Importe os arquivos desejados (.doc, .pptx, .pdf, .xlsx) e aguarde a conversão
⚠️ Pré-requisito: É necessário ter Microsoft Word, Excel e PowerPoint instalados no computador, pois a conversão utiliza automação COM via pywin32.


💻 Versão Código-Fonte (para desenvolvedores)
git clone https://github.com/seu-usuario/Dootax-Conversor.git
cd Dootax-Conversor
python -m venv dootax_env
dootax_env\Scripts\activate
pip install -r requirements.txt
python -m app.main



📂 Estrutura do projeto
Dootax-Conversor/
├── app/
│   ├── main.py
│   ├── gui/
│   │   └── interface.py
│   ├── controladores/
│   │   └── seletores.py
│   ├── conversor/
│   │   ├── conversor_doc.py
│   │   ├── conversor_pdf.py
│   │   ├── conversor_pptx.py
│   │   ├── conversor_excel.py
│   │   ├── pptx_utils.py
│   │   ├── notificacoes.py
│   │   └── utils.py
├── arquivos_convertidos/
├── requirements.txt
└── .gitignore



🧪 Tecnologias
- python-docx
- python-pptx
- PyMuPDF
- xlwings
- pywin32
- Pillow
- Tkinter (interface nativa)
