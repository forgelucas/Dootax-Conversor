# DOCX-Conversor

# 📝 Docx-Conversor

Ferramenta em Python com interface gráfica para conversão de arquivos `.doc` para `.docx`. Ideal para analistas que precisam migrar documentações de forma rápida e automatizada.

## 🚀 Funcionalidades

- Conversão automática de arquivos `.doc` para `.docx`;
- Suporte à seleção de múltiplos arquivos de uma vez;
- Renomeação automática se arquivos tiverem o mesmo nome;
- Notificação ao término da conversão;
- Abertura automática da pasta de arquivos convertidos;
- Barra de carregamento para feedback visual durante a execução;
- Distribuído como executável `.exe` (não requer instalação do Python).

## 🖼️ Interface

A interface é simples e intuitiva:

- Botão para importar arquivos;
- Mensagens informativas sobre o andamento da conversão;
- Barra de progresso animada durante a conversão;
- Notificação ao final indicando sucesso da operação.

## 📦 Como executar

### ✔️ Versão Executável (recomendado)

1. Baixe o arquivo `Conversor.exe` da pasta `dist` (ou release no GitHub).
2. Dê dois cliques para executar.
3. Importe os arquivos `.doc` desejados e aguarde a conversão.

> **Pré-requisito:** É necessário ter o **Microsoft Word** instalado no computador, pois o processo de conversão utiliza o Word via automação COM.

---

### 💻 Versão Código-Fonte (desenvolvedores)

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/Docx-Conversor.git
   cd Docx-Conversor
