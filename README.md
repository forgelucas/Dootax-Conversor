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

Dootax-Conversor/
├── app/
│   ├── main.py
│   ├── gui/interface.py
│   ├── controladores/seletores.py
│   └── conversor/
│       ├── conversor_doc.py
│       ├── conversor_pdf.py
│       ├── conversor_pptx.py
│       ├── conversor_excel.py
│       ├── pptx_utils.py
│       ├── notificacoes.py
│       └── utils.py
├── arquivos_convertidos/
├── requirements.txt
└── .gitignore

---

### 🧪 Frameworks

- python-docx
- python-pptx
- PyMuPDF
- xlwings
- pywin32
- Pillow
- Tkinter
