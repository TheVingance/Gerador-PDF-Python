# -*- coding: utf-8 -*-
import os
import sys

if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

PASTA_LOCAL = r"C:\Users\hiurkiv\Documents\roover-python"
IMG_GRAFICO = os.path.join(PASTA_LOCAL, "grafico_distribuicao_cadastros.png")
DOCX_RELATORIO = os.path.join(PASTA_LOCAL, "docs", "Relatorio_Comparativo_docs-roover.docx") if os.path.isdir(os.path.join(PASTA_LOCAL, "docs")) else os.path.join(PASTA_LOCAL, "Relatorio_Comparativo_docs-roover.docx")
MD_RELATORIO = os.path.join(PASTA_LOCAL, "Relatorio_Comparativo_Auditoria.md")

# ==============================================================================
# 1. GERAR GRÁFICO EM PNG ALTA RESOLUÇÃO (MATPLOTLIB)
# ==============================================================================
def gerar_grafico():
    print("[1/3] Gerando gráfico de alta resolução em PNG...")
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6), dpi=300)
    fig.patch.set_facecolor('#FFFFFF')

    # Gráfico 1: Pizza / Donut dos 20.163 colaboradores de Assinados
    labels1 = ['Completos\n(Assinados + Pessoais)\n10.615 (52,6%)', 
               'Apenas Assinados\n(Ausentes em Pessoais)\n9.548 (47,4%)']
    sizes1 = [10615, 9548]
    colors1 = ['#2E6B9E', '#D9534F']
    explode1 = (0.02, 0.05)

    wedges, texts, autotexts = ax1.pie(
        sizes1, 
        explode=explode1, 
        labels=labels1, 
        autopct='%1.1f%%',
        startangle=140, 
        colors=colors1,
        pctdistance=0.75,
        textprops={'fontsize': 10, 'color': '#2C3E50', 'weight': 'bold'},
        wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2)
    )
    for at in autotexts:
        at.set_color('white')
        at.set_fontsize(11)
        at.set_weight('bold')

    ax1.set_title("Composição da Base de Assinados\n(Total: 20.163 Colaboradores)", 
                  fontsize=12, fontweight='bold', color='#1F4E78', pad=15)

    # Gráfico 2: Barras Comparativas de Volumes
    categorias = [
        'Cadastros em\nAssinados',
        'Cadastros em\nPessoais',
        'Presentes em\nAmbas',
        'Apenas em\nAssinados',
        'Apenas em\nPessoais'
    ]
    valores = [20163, 11985, 10615, 9548, 1370]
    cores_barras = ['#1F4E78', '#418AB3', '#2E6B9E', '#D9534F', '#F0AD4E']

    bars = ax2.bar(categorias, valores, color=cores_barras, width=0.55, edgecolor='#E0E0E0', linewidth=1)
    ax2.set_facecolor('#F8F9FA')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['left'].set_color('#CCCCCC')
    ax2.spines['bottom'].set_color('#CCCCCC')
    ax2.grid(axis='y', linestyle='--', alpha=0.5, color='#D0D0D0')
    ax2.set_title("Comparativo Geral de Volumes por Base", 
                  fontsize=12, fontweight='bold', color='#1F4E78', pad=15)
    ax2.tick_params(axis='x', labelsize=9)
    ax2.tick_params(axis='y', labelsize=9)

    for bar in bars:
        h = bar.get_height()
        ax2.annotate(f'{h:,}'.replace(',', '.'),
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 4),  
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=9.5, fontweight='bold', color='#2C3E50')

    plt.tight_layout()
    plt.savefig(IMG_GRAFICO, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Gráfico salvo com sucesso: {IMG_GRAFICO}")

# ==============================================================================
# 2. GERAR RELATÓRIO FORMATADO EM WORD (.DOCX)
# ==============================================================================
def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def gerar_word():
    print("[2/3] Gerando relatório executivo em Word (.docx)...")
    doc = docx.Document()

    # Configuração de Margens (Normal: 2.54 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Cores da Paleta
    COLOR_PRIMARY = RGBColor(0x1F, 0x4E, 0x78)   # Navy Blue
    COLOR_SECONDARY = RGBColor(0x59, 0x59, 0x59) # Gray
    COLOR_RED = RGBColor(0xC0, 0x00, 0x00)       # Dark Red
    HEX_HEADER = "1F4E78"
    HEX_ZEBRA = "F9FAFC"
    HEX_ALERT = "FFF3CD"

    # Título Principal
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_title = p_title.add_run("Relatório de Auditoria Comparativa")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("Conferência de Integridade Cadastral: Documentos Assinados vs Documentos Pessoais")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(13)
    r_sub.font.bold = True
    r_sub.font.color.rgb = COLOR_SECONDARY

    p_meta = doc.add_paragraph()
    r_meta = p_meta.add_run(f"Data de Emissão: {datetime.now().strftime('%d/%m/%Y')} | Fonte: Google Drive (Somente Leitura) | Sistema: Roover")
    r_meta.font.name = "Calibri"
    r_meta.font.size = Pt(9.5)
    r_meta.font.italic = True
    r_meta.font.color.rgb = COLOR_SECONDARY

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Caixa de Segurança / Diretriz
    p_sec = doc.add_paragraph()
    p_sec.paragraph_format.space_before = Pt(4)
    p_sec.paragraph_format.space_after = Pt(12)
    r_sec_title = p_sec.add_run("DIRETRIZ DE SEGURANÇA E LEITURA: ")
    r_sec_title.font.bold = True
    r_sec_title.font.size = Pt(10)
    r_sec_title.font.color.rgb = COLOR_PRIMARY
    r_sec_body = p_sec.add_run(
        "A pasta compartilhada do Google Drive foi acessada de maneira rigorosamente em modo leitura. "
        "Nenhum arquivo na nuvem foi modificado, mantendo a integridade absoluta dos documentos originais. "
        "Todos os processamentos, gráficos e relatórios foram gerados localmente."
    )
    r_sec_body.font.size = Pt(9.5)
    r_sec_body.font.italic = True

    # 1. Visão Geral e Métricas
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Visão Geral e Métricas Consolidadas")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = COLOR_PRIMARY

    p_intro = doc.add_paragraph(
        "Durante a validação dos PDFs gerados em lote, foi identificada a ausência de documentos "
        "de admissão para determinados colaboradores admitidos antes da digitalização dos processos de onboarding. "
        "O cruzamento das duas bases primárias revelou a seguinte composição:"
    )
    p_intro.paragraph_format.space_after = Pt(10)

    # Inserir Gráfico PNG
    if os.path.exists(IMG_GRAFICO):
        doc.add_picture(IMG_GRAFICO, width=Inches(6.5))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cap = p_cap.add_run("Figura 1: Distribuição dos cadastros e comparação de volumes entre bases.")
        r_cap.font.size = Pt(9)
        r_cap.font.italic = True
        r_cap.font.color.rgb = COLOR_SECONDARY
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Tabela de Métricas
    tabela_dados = [
        ("Total de Cadastros Únicos em Assinados", "20.163", "100,0% da base com contratos assinados"),
        ("Total de Cadastros Únicos em Pessoais", "11.985", "Colaboradores com processo de admissão digital"),
        ("Cadastros Presentes em AMBAS as Planilhas", "10.615", "52,6% possuem admissão + contratos assinados"),
        ("Cadastros APENAS em Assinados (Inconsistência)", "9.548", "47,4% dos contratos NÃO têm documentos de admissão"),
        ("  ↳ Totalmente ausentes em pessoais (ex: Admissões históricas)", "9.444", "Sem qualquer cadastro no repositório de admissão"),
        ("  ↳ Divergência de CPF (mesma matrícula em pessoais)", "42", "Funcionário com matrícula idêntica e erro no CPF"),
        ("  ↳ Divergência de CPF (mesmo nome completo em pessoais)", "62", "Mesmo nome completo cadastrado com CPF divergente"),
        ("Cadastros APENAS em Pessoais", "1.370", "Colaboradores admitidos sem contratos digitais assinados")
    ]

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    col_widths = [Inches(3.3), Inches(1.1), Inches(2.1)]

    # Cabeçalho da Tabela
    hdr_cells = table.rows[0].cells
    hdr_titles = ["Métrica de Auditoria", "Quantidade", "Percentual / Detalhe"]
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], HEX_HEADER)
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=150, right=150)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 1 else WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size = Pt(10)

    # Linhas de Dados
    for row_idx, (m, q, d) in enumerate(tabela_dados):
        row_cells = table.add_row().cells
        row_cells[0].text = m
        row_cells[1].text = q
        row_cells[2].text = d

        is_inconsistencia = "INCONSISTÊNCIA" in m.upper()
        bg_color = HEX_ZEBRA if row_idx % 2 == 1 else "FFFFFF"

        for i in range(3):
            set_cell_background(row_cells[i], bg_color)
            set_cell_margins(row_cells[i], top=90, bottom=90, left=150, right=150)
            p = row_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i == 1 else WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.size = Pt(9.5)
                if i == 0 and ("Total" in m or is_inconsistencia):
                    run.font.bold = True
                if i == 1:
                    run.font.bold = True
                    if is_inconsistencia:
                        run.font.color.rgb = COLOR_RED
                if i == 2:
                    run.font.italic = True

    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. Exemplo Prático de Admissão Histórica
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Diagnóstico Prático: Exemplo de Admissões Históricas")
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = COLOR_PRIMARY

    p_ros = doc.add_paragraph(
        "Durante a validação com a liderança, foi analisada a situação de colaboradores com admissão anterior "
        "à implantação dos processos digitais. Como exemplo ilustrativo, foram identificados colaboradores na base de assinados "
        "que não possuem cadastro na planilha de documentos pessoais:"
    )
    p_ros.paragraph_format.space_after = Pt(8)

    # Tabela Exemplo Admissao Historica
    table_ros = doc.add_table(rows=1, cols=6)
    table_ros.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_ros.autofit = False

    ros_headers = ["CPF", "Matrícula", "Nome Completo", "Admissão", "Docs Assinados", "Situação em Pessoais"]
    ros_widths = [Inches(1.1), Inches(0.8), Inches(2.1), Inches(0.9), Inches(0.8), Inches(1.8)]

    hdr_ros = table_ros.rows[0].cells
    for i, title in enumerate(ros_headers):
        hdr_ros[i].text = title
        set_cell_background(hdr_ros[i], "333F48")
        set_cell_margins(hdr_ros[i], top=100, bottom=100, left=100, right=100)
        p = hdr_ros[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 1, 3, 4] else WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size = Pt(9.5)

    dados_ros = [
        ("11122233344", "1001", "COLABORADOR EXEMPLO A", "01/08/2003", "7", "NÃO CONSTA em roover-docs-pessoais"),
        ("55566677788", "1002", "COLABORADOR EXEMPLO B", "25/11/2008", "5", "NÃO CONSTA em roover-docs-pessoais")
    ]

    for row_idx, r_vals in enumerate(dados_ros):
        r_cells = table_ros.add_row().cells
        for i, val in enumerate(r_vals):
            r_cells[i].text = val
            set_cell_background(r_cells[i], "FFFFFF")
            set_cell_margins(r_cells[i], top=80, bottom=80, left=100, right=100)
            p = r_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 1, 3, 4] else WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.size = Pt(9)
                if i == 5:
                    run.font.bold = True
                    run.font.color.rgb = COLOR_RED

    for row in table_ros.rows:
        for i, w in enumerate(ros_widths):
            row.cells[i].width = w

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    p_exp = doc.add_paragraph()
    r_exp_t = p_exp.add_run("Justificativa Técnica para a Chefia:\n")
    r_exp_t.font.bold = True
    r_exp_t.font.size = Pt(10)
    p_exp.add_run(
        "1. As admissões ocorreram em anos anteriores à implantação do sistema digital de onboarding.\n"
        "2. Como não passaram pelo fluxo de admissão eletrônico, seus documentos físicos não foram carregados no repositório de admissão (arquivos-admissao).\n"
        "3. Posteriormente, foram assinados contratos e aditivos eletrônicos, motivo pelo qual constam normalmente na planilha de contratos assinados.\n"
        "4. O gerador de PDF agiu com total integridade: como não havia fotos de admissão, consolidou com sucesso todos os contratos assinados disponíveis, gerando o PDF íntegro."
    ).font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 3. Casos de Divergência de Digitação
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Alerta de Saneamento: Divergências de CPF (104 Casos)")
    r_h3.font.name = "Calibri"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = COLOR_PRIMARY

    p_div = doc.add_paragraph(
        "A auditoria identificou um subconjunto de 104 colaboradores que constam em ambas as bases, "
        "mas não foram correlacionados pelo CPF devido a erros de digitação ou cadastros divergentes entre sistemas:\n\n"
        "• 42 colaboradores possuem Matrícula idêntica em ambas as planilhas, mas o CPF difere.\n"
        "• 62 colaboradores possuem Nome Completo idêntico, mas com CPF registrado diferente.\n\n"
        "Exemplo ilustrativo detectado: Um colaborador com Matrícula 12345 consta em assinados com CPF 11122233344 e em pessoais com CPF 11122233345 (diferença de apenas 1 dígito de digitação). "
        "Todos esses 104 casos foram marcados com destaque visual em amarelo na planilha Excel gerada para saneamento pelo RH/TI."
    )
    p_div.paragraph_format.space_after = Pt(12)

    # 4. Arquivos Gerados
    h4 = doc.add_heading(level=1)
    r_h4 = h4.add_run("4. Entregáveis e Arquivos Gerados no Repositório")
    r_h4.font.name = "Calibri"
    r_h4.font.size = Pt(14)
    r_h4.font.bold = True
    r_h4.font.color.rgb = COLOR_PRIMARY

    p_ent = doc.add_paragraph(
        "Para consulta da chefia e auditoria contínua, os seguintes arquivos foram gerados e salvos localmente:\n\n"
        "1. comparativo_cadastros_assinados_vs_pessoais.xlsx: Planilha Excel oficial com 3 abas estilizadas (Resumo Executivo, Cadastros Apenas Assinados e Cadastros Apenas Pessoais).\n"
        "2. cadastros_apenas_em_assinados.csv: Exportação direta em CSV delimitado por ponto e vírgula com os 9.548 registros ausentes.\n"
        "3. grafico_distribuicao_cadastros.png: Gráfico comparativo de alta resolução.\n"
        "4. gerar_planilha_comparativa.py: Script em Python documentado para reexecução automática sob demanda.\n"
        "5. Relatorio_Comparativo_Auditoria.docx: Este documento formal de auditoria."
    )
    p_ent.paragraph_format.space_after = Pt(12)

    doc.save(DOCX_RELATORIO)
    print(f"Relatório Word salvo com sucesso: {DOCX_RELATORIO}")

# ==============================================================================
# 3. GERAR VERSÃO MARKDOWN (.MD) COM MERMAID CORRIGIDO
# ==============================================================================
def gerar_markdown():
    print("[3/3] Gerando relatório em Markdown (.md) para o repositório...")
    md_content = """# Relatório de Auditoria Comparativa: Assinados vs Pessoais

> **Diretriz de Segurança:**  
> A pasta compartilhada no Google Drive (`G:\\.shortcut-targets-by-id\\...\\Roover`) foi acessada em **modo estritamente de leitura**. Nenhum arquivo na nuvem foi modificado. Todas as planilhas e relatórios foram gerados localmente no repositório do projeto.

---

## 📌 1. Visão Geral da Análise

Durante a conferência dos PDFs gerados com a liderança, foi identificada a ausência de documentos pessoais/admissão para determinados colaboradores admitidos antes da digitalização dos processos de onboarding.

Ao cruzar as duas fontes primárias do Google Drive:
1. `roover-docs-assinados.csv` (231.729 registros de contratos e termos assinados)
2. `roover-docs-pessoais.csv` (11.986 registros de documentos de admissão)

Foram identificados os totais consolidados de colaboradores (CPFs únicos) em cada situação:

### Gráfico Mermaid (Corrigido)

```mermaid
pie title "Cadastros em Assinados (Total: 20.163)"
    "Completos (Assinados + Pessoais)" : 10615
    "Apenas Assinados (Sem Docs Pessoais)" : 9548
```

### Gráfico em Alta Resolução (PNG)

![Gráfico Comparativo de Auditoria](grafico_distribuicao_cadastros.png)

---

## 📊 2. Tabela de Métricas Consolidadas

| Métrica de Auditoria | Quantidade | Percentual / Detalhe | Status |
| :--- | :---: | :---: | :--- |
| **Total de Cadastros Únicos em Assinados** | **20.163** | 100,0% | Base total com contratos assinados |
| **Total de Cadastros Únicos em Pessoais** | **11.985** | — | Base de colaboradores com admissão digital |
| **Cadastros Presentes em AMBAS as Planilhas** | **10.615** | 52,6% | Possuem pasta de admissão e contratos assinados |
| **Cadastros APENAS em Assinados (Inconsistência)** | **9.548** | **47,4%** | **Constam em assinados mas NÃO em pessoais** |
| ↳ *Totalmente ausentes em pessoais (ex: Admissões históricas)* | 9.444 | 46,8% | Sem correspondência de CPF, Matrícula ou Nome |
| ↳ *Divergência de CPF (Mesma matrícula em pessoais)* | 42 | 0,2% | Matrícula idêntica, mas CPF divergente ou com erro de digitação |
| ↳ *Divergência de CPF (Mesmo nome em pessoais)* | 62 | 0,3% | Mesmo nome completo cadastrado com outro CPF |
| **Cadastros APENAS em Pessoais** | **1.370** | — | Colaboradores com admissão sem contratos assinados |

---

## 🔍 3. Análise Prática: Exemplo de Admissões Históricas

Na planilha de contratos assinados (`roover-docs-assinados.csv`), constam colaboradores admitidos em períodos históricos:

| CPF | Matrícula | Nome Completo | Data Admissão | Qtd Docs Assinados | Presença em `docs-pessoais` |
| :---: | :---: | :--- | :---: | :---: | :--- |
| `11122233344` | `1001` | **COLABORADOR EXEMPLO A** | 01/08/2003 | 7 contratos | ❌ **Ausente** (não consta CPF, Matrícula nem Nome) |
| `55566677788` | `1002` | **COLABORADOR EXEMPLO B** | 25/11/2008 | 5 contratos | ❌ **Ausente** (não consta CPF, Matrícula nem Nome) |

### Motivo Técnico da Inconsistência
1. **Admissões Históricas:** As admissões ocorreram há muitos anos (**2003** e **2008**), data anterior à digitalização do sistema de onboarding (`arquivos-admissao`).
2. **Ausência de Documentos Físicos Digitalizados:** Por não terem passado pelo processo eletrônico de admissão, não há pastas com prefixo `{ID_SISTEMICO}_` na nuvem.
3. **Assinaturas Eletrônicas Recentes:** Posteriormente, foram assinados termos e contratos eletrônicos, motivo pelo qual constam com 7 e 5 arquivos na planilha de assinados.
4. **Comportamento do Gerador:** O gerador de PDF agiu corretamente: unificou com sucesso todos os contratos assinados existentes, ficando apenas sem a seção de admissão.

---

## 🎯 4. Diagnóstico de Saneamento: Divergências de CPF (104 Casos)

Além dos colaboradores ausentes por admissão antiga, o cruzamento identificou **104 casos** onde o colaborador existe em ambas as planilhas, mas não se cruzavam automaticamente devido a divergências no CPF:

> [!TIP]
> **Exemplo Ilustrativo Detectado:**  
> Um colaborador com Matrícula `12345`:  
> - Em `roover-docs-assinados`: CPF cadastrado como `11122233344`  
> - Em `roover-docs-pessoais`: CPF cadastrado como `11122233345` (diferença de apenas 1 dígito de digitação)  
> 
> Todos esses casos foram sinalizados com status destacado em amarelo na planilha gerada, com o detalhe da matrícula e ID Sistêmico para que a equipe de RH/TI possa efetuar o saneamento cadastral.

---

## 📁 5. Arquivos Gerados no Repositório

1. **Documento Formal em Word (.docx):**  
   [Relatorio_Comparativo_Auditoria.docx](Relatorio_Comparativo_Auditoria.docx)
2. **Planilha Excel Executiva com 3 Abas:**  
   [comparativo_cadastros_assinados_vs_pessoais.xlsx](comparativo_cadastros_assinados_vs_pessoais.xlsx)
3. **Arquivo CSV Direto:**  
   [cadastros_apenas_em_assinados.csv](cadastros_apenas_em_assinados.csv)
4. **Gráfico em Alta Resolução (PNG):**  
   [grafico_distribuicao_cadastros.png](grafico_distribuicao_cadastros.png)
5. **Script de Auditoria Automatizado:**  
   [gerar_planilha_comparativa.py](gerar_planilha_comparativa.py)
"""
    with open(MD_RELATORIO, "w", encoding="utf-8") as f:
        f.write(md_content)
    print(f"Relatório Markdown salvo com sucesso: {MD_RELATORIO}")

if __name__ == "__main__":
    gerar_grafico()
    gerar_word()
    gerar_markdown()
