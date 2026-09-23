# Walkthrough: Geração de PDFs em Lote (4 Workers Paralelos)

A funcionalidade de **geração de PDFs em lote** para todos os cadastros foi implementada e validada com sucesso no script [gerador_pdf_pessoa.py](file:///c:/Users/hiurkiv/Documents/roover-python/gerador_pdf_pessoa.py).

---

## ⚙️ Principais Características Implementadas

1. **4 Workers Paralelos por Padrão:**
   * O lote roda distribuído em **4 threads simultâneas**, aproveitando o processamento sem sobrecarregar a máquina.
2. **Retomada Incremental (Skip de Arquivos Existentes):**
   * Se o arquivo `pdfs_gerados/{CPF}-{NOME}-roover.pdf` já existir com tamanho válido, o script pula imediatamente (`JA EXISTE`), permitindo pausar e continuar o lote sem retrabalho.
3. **Log Limpo e Acompanhamento Percentual:**
   * Exibe em tempo real o contador e percentual:  
     `[  142/21533] (  0.7%) | 12345678901 | COLABORADOR EXEMPLO.. | OK (13 pags, 1.40 MB)`
4. **Relatório CSV Automático:**
   * Ao finalizar o lote, salva o status de todos os registros em [relatorio_geracao.csv](file:///C:/Users/hiurkiv/Documents/roover-python/relatorio_geracao.csv).
5. **Gravação Atômica e Segura:**
   * Os arquivos são gerados primeiro em `.tmp` e renomeados apenas após o término da gravação, evitando PDFs corrompidos caso o processo seja interrompido.

---

## 🧪 Validação dos Testes Realizados

| Teste | Comando Executado | Resultado |
| :--- | :--- | :--- |
| **Lote Pequeno (5 cadastros)** | `gerador_pdf_pessoa.py --todos --limite 5` | **Sucesso:** 4 novos PDFs criados em 13s, 1 ignorado (já existia). |
| **Teste de Retomada (Resume)** | `gerador_pdf_pessoa.py --todos --limite 5` | **Sucesso:** 5 ignorados instantaneamente em 0s. |
| **Consulta Individual** | `gerador_pdf_pessoa.py --cpf 12345678901 --sobrescrever` | **Sucesso:** 13 páginas consolidadas, sobrescrita confirmada. |

---

## 🚀 Como Executar

Abra o terminal no diretório `C:\Users\hiurkiv\Documents\roover-python` e execute:

```powershell
# Iniciar a geração de TODA a base consolidada (4 workers padrão):
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --todos

# Testar com um limite específico de pessoas antes (ex: 20 pessoas):
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --todos --limite 20

# Consultar uma pessoa individualmente:
.\.venv\Scripts\python.exe gerador_pdf_pessoa.py --cpf 12345678901
```
