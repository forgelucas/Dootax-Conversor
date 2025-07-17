# 📝 Dootax Conversor

Ferramenta em Python com interface gráfica para conversão de arquivos `.doc`, `.pptx`, `.pdf` e `.xlsx` para `.docx`. Ideal para analistas, contadores e desenvolvedores que precisam migrar documentações de forma rápida, limpa e automatizada.

---

## 🚀 Funcionalidades

- ✅ Conversão automática dos formatos `.doc`, `.pptx`, `.pdf` e `.xlsx` para `.docx`
- ✅ Interface gráfica amigável com botões por tipo de arquivo
- ✅ Suporte à seleção de múltiplos arquivos
- ✅ Extração de imagens (sem salvar no disco) e textos com formatação
- ✅ Renomeação automática de arquivos com nomes repetidos
- ✅ Notificações visuais ao término da conversão
- ✅ Abertura automática da pasta de saída
- ✅ Distribuído como `.exe` — não requer Python instalado

---

## 📦 Como executar

### ✔️ Versão Executável

1. Baixe o arquivo `DootaxConversor.exe` na pasta `dist/` ou na aba **Releases**
2. Execute com duplo clique
3. Importe seus arquivos (`.doc`, `.pptx`, `.pdf`, `.xlsx`) e aguarde o resultado

> ⚠️ Requisitos: É necessário ter o **Microsoft Word, Excel e PowerPoint instalados**, pois o projeto utiliza automação COM (`pywin32`).

---

### 💻 Versão Código-Fonte (Desenvolvedores)

```bash
git clone https://github.com/seu-usuario/Dootax-Conversor.git
cd Dootax-Conversor

python -m venv dootax_env
dootax_env\Scripts\activate

pip install -r requirements.txt
python -m app.main
```
---

### 📂 Estrutura do Projeto

```plaintext
Dootax-Conversor/
├── .github/
│   └── workflows/
│       └── ci.yml                # Pipeline de CI com lint, type checking, testes etc.
├── app/
│   ├── __init__.py              # Torna 'app' um pacote Python
│   ├── main.py                  # Ponto de entrada da aplicação
│
│   ├── controladores/
│   │   ├── __init__.py
│   │   └── seletores.py         # Controla seleção dos arquivos na interface
│
│   ├── conversor/
│   │   ├── __init__.py
│   │   ├── conversor_doc.py
│   │   ├── conversor_excel.py
│   │   ├── conversor_pdf.py
│   │   ├── conversor_pptx.py
│   │   ├── docx_helpers.py      # Auxiliares para gerar e formatar docx
│   │   ├── notificacoes.py      # Toasts, alertas e abertura de pasta convertida
│   │   ├── pptx_utils.py        # Extração de texto/imagem dos slides
│   │   └── utils.py             # Funções utilitárias compartilhadas
│
│   ├── gui/
│   │   ├── __init__.py
│   │   ├── interface.py         # Interface gráfica principal
│   │   └── logo_dootax.png      # Imagem da logo usada no app
│
├── arquivos_convertidos/        # Pasta onde os arquivos convertidos são salvos
├── requirements.txt             # Dependências do projeto
├── setup.cfg                    # Configurações do flake8, mypy, etc.
└── README.md                    # Documentação do projeto
```


---

### 🧪 Frameworks

- python-docx
- python-pptx
- PyMuPDF
- xlwings
- pywin32
- Pillow
- Tkinter
