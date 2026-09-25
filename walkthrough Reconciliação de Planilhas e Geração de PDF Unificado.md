# Walkthrough: Reconciliação de Planilhas e Geração de PDF Unificado

Todas as implementações foram concluídas e validadas com sucesso, respeitando rigorosamente a regra de **acesso somente leitura ao Google Drive**.

---

## 🔒 1. Segurança e Integridade do Google Drive
* Os arquivos e diretórios originais em `G:\.shortcut-targets-by-id\11GxyE8SbWQbCXzoP8f3e9cGFfVK0TIFn\Roover` **permaneceram 100% inalterados**.
* Todas as saídas (dados consolidados, índices locais e arquivos PDF) foram salvas exclusivamente na pasta local:  
  `C:\Users\hiurkiv\Documents\roover-python\`.

---

## 📊 2. Reconciliação das Bases (`reconciliar_bases.py`)

O script [reconciliar_bases.py](file:///C:/Users/hiurkiv/Documents/roover-python/reconciliar_bases.py) executou o cruzamento entre `roover-docs-assinados.csv` e `roover-docs-pessoais.csv`:

* **Normalização de CPF:** Formatação com 11 dígitos numéricos com preenchimento de zeros à esquerda (`zfill(11)`).
* **Tratamento de Lacunas:**
  * Onde `NOME` ou `MATRICULA` estavam em branco, nulos ou preenchidos como `"roover"`, foram completados com o dado correspondente da outra planilha.
  * Informações ausentes em ambas as bases foram preenchidas com `"roover"`.
* **Resultado Consolidado:**
  * Salvo em: [base_consolidada_roover.csv](file:///C:/Users/hiurkiv/Documents/roover-python/base_consolidada_roover.csv)
  * **21.533 CPFs únicos** mapeados.
  * **10.917 nomes e matrículas enriquecidos** pelo cruzamento.
  * **10.615 pessoas** possuem tanto documentos de admissão quanto contratos assinados.

---

## 📑 3. Gerador de PDF Unificado (`gerador_pdf_pessoa.py`)

O script [gerador_pdf_pessoa.py](file:///C:/Users/hiurkiv/Documents/roover-python/gerador_pdf_pessoa.py) unifica todos os documentos de um colaborador:

1. **Ordem dos Documentos:**
   * **1º Bloco (Admissão / Pessoais):** Busca os arquivos na pasta `arquivos-admissao` pelo prefixo `{ID_SISTEMICO}_`. Converte dinamicamente imagens (`.jpg`, `.jpeg`, `.png`, `.bmp`) em páginas de PDF e anexa PDFs existentes.
   * **2º Bloco (Contratos e Termos Assinados):** Anexa os documentos da pasta `assinados` registrados para o CPF/Matrícula.
2. **Padrão de Nome:** `"{CPF}-{NOME COMPLETO}-roover.pdf"`.
3. **Destino:** Salvo na pasta `C:\Users\hiurkiv\Documents\roover-python\pdfs_gerados\`.

---

## ✅ 4. Validação Realizada com Caso de Teste

Executamos o teste para a colaboradora:
* **Nome:** COLABORADOR EXEMPLO SILVA
* **CPF:** `12345678901`
* **Matrícula:** `77344`
* **ID Sistêmico:** `6364`

### Resultado da Geração:
* **Arquivo Gerado:** `12345678901-COLABORADOR EXEMPLO SILVA-roover.pdf`
* **Localização:** `C:\Users\hiurkiv\Documents\roover-python\pdfs_gerados\`
* **Total de Páginas:** **13 páginas**
* **Composição do PDF:**
  1. `6364_cre_550036d1c8ff9478a45be22585b35b03.jpeg` (Imagem convertida)
  2. `6364_out_affff6f8f1a573a13131e4c2b948ed08.jpeg` (Imagem convertida)
  3. `6364_out_cac25ca925dbafc6e00a2ce919c35e33.jpeg` (Imagem convertida)
  4. `6364_rg_442dcd929ca40ec998c3ee332540cf9a.jpg` (Imagem convertida)
  5. `sign_FPDO103-0001-1-000077344-12345678901.PDF` (2 páginas)
  6. `sign_FPDO104-0001-1-000077344-12345678901.PDF` (3 páginas)
  7. `sign_FPDO105-0001-1-000077344-12345678901.PDF` (2 páginas)
  8. `sign_FPDO106-0001-1-000077344-12345678901.PDF` (2 páginas)
* **Tamanho Final:** 1.40 MB

---

## 🚀 Como Usar

Para gerar o PDF de qualquer pessoa, abra o terminal no diretório `C:\Users\hiurkiv\Documents\roover-python` e execute:

```powershell
# 1. Gerar para uma pessoa específica por CPF:
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --cpf 12345678901

# 2. Gerar para uma pessoa específica por Matrícula:
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --matricula 77344

# 3. Gerar interativamente (o script pedirá CPF ou Matrícula na tela):
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py

# 4. GERAÇÃO EM LOTE (TODOS OS 21.533 REGISTROS):
# Por padrão utiliza 4 workers paralelos e pula arquivos já gerados:
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --todos

# 5. Testar lote menor primeiro (ex: apenas os primeiros 10):
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --todos --limite 10

# 6. Forçar regeração de arquivos já existentes:
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --todos --sobrescrever
```

### 📊 Relatório Automático de Execução
Ao término de qualquer processamento em lote, é gerado automaticamente o arquivo:
* `C:\Users\hiurkiv\Documents\roover-python\relatorio_geracao.csv`
Contendo as colunas: `CPF;MATRICULA;NOME;STATUS;PAGINAS;TAMANHO_MB;ERRO`.
