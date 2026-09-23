# Roover Python - Processador de Documentos e PDFs

Projeto local para reconciliação de planilhas e consolidação de documentos em arquivos PDF unificados por pessoa.

---

## 🔒 Diretrizes de Segurança

* **Pasta de Origem (Apenas Leitura no Google Drive):**
  `G:\.shortcut-targets-by-id\11GxyE8SbWQbCXzoP8f3e9cGFfVK0TIFn\Roover`
  > **IMPORTANTE:** Nenhum arquivo ou pasta dentro do Drive `G:\` é modificado, criado ou apagado. O acesso é 100% de leitura.

* **Pasta de Destino (Local no Computador):**
  `C:\Users\hiurkiv\Documents\roover-python\`
  > Todos os scripts, dependências, arquivos processados e PDFs finais são salvos exclusivamente nesta pasta.

---

## 📁 Estrutura de Arquivos Locais

| Arquivo / Pasta | Descrição |
| :--- | :--- |
| **`reconciliar_bases.py`** | Script que lê as planilhas do Drive, cruza `CPF`, completa campos faltantes (`NOME`, `MATRICULA`) e salva a base unificada. |
| **`gerador_pdf_pessoa.py`** | Script que gera o PDF consolidado de uma pessoa (`{CPF}-{NOME}-roover.pdf`), convertendo imagens e unificando contratos. |
| **`base_consolidada_roover.csv`** | Planilha consolidada com 21.533 colaboradores, status de documentos e dados cruzados. |
| **`lista_admissao.txt`** | Índice local com a lista de arquivos da pasta `arquivos-admissao` para permitir busca instantânea sem gargalos de rede. |
| **`pdfs_gerados/`** | Pasta onde são salvos os arquivos PDF gerados. |
| **`.venv/`** | Ambiente virtual Python com `pandas`, `pypdf`, `Pillow` instalados. |
| **`requirements.txt`** | Lista de pacotes Python instalados. |

---

## 🚀 Como Executar

Abra o terminal (PowerShell ou CMD) na pasta `C:\Users\hiurkiv\Documents\roover-python`:

### 1. Gerar PDF de uma pessoa específica por CPF:
```powershell
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --cpf 12345678901
```

### 2. Gerar PDF por Matrícula:
```powershell
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --matricula 77344
```

### 3. Modo Interativo:
```powershell
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py
```
*(O script solicitará o CPF ou Matrícula diretamente no terminal)*

### 4. Reexecutar a reconciliação das bases (se as planilhas do Drive forem atualizadas):
```powershell
.\.venv\Scripts\python.exe reconciliar_bases.py
```
