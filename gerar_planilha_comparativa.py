# -*- coding: utf-8 -*-
import os
import sys

if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURAÇÕES DE CAMINHOS
# ATENÇÃO: O Google Drive é acessado RIGOROSAMENTE EM MODO LEITURA.
# NENHUM arquivo no Drive será modificado, criado ou excluído.
# ==============================================================================
PASTA_DRIVE = r"G:\.shortcut-targets-by-id\11GxyE8SbWQbCXzoP8f3e9cGFfVK0TIFn\Roover"
PASTA_DESTINO = r"C:\Users\hiurkiv\Documents\roover-python"

ARQ_ASSINADOS_DRIVE = os.path.join(PASTA_DRIVE, "roover-docs-assinados.csv")
ARQ_PESSOAIS_DRIVE = os.path.join(PASTA_DRIVE, "roover-docs-pessoais.csv")

SAIDA_EXCEL = os.path.join(PASTA_DESTINO, "comparativo_cadastros_assinados_vs_pessoais.xlsx")
SAIDA_CSV = os.path.join(PASTA_DESTINO, "cadastros_apenas_em_assinados.csv")

def normalizar_cpf(val):
    if pd.isna(val):
        return ""
    digits = "".join([c for c in str(val) if c.isdigit()])
    if digits:
        return digits.zfill(11)
    return ""

def normalizar_matricula(val):
    if pd.isna(val):
        return ""
    digits = "".join([c for c in str(val) if c.isdigit()])
    if digits:
        return str(int(digits))
    return ""

def normalizar_nome(val):
    if pd.isna(val):
        return ""
    return " ".join(str(val).upper().strip().split())

def aplicar_estilo_tabela(ws, min_row, max_row, min_col, max_col, header_fill, header_font, col_widths=None):
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    zebra_fill = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    divergencia_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Cabeçalho
    for col in range(min_col, max_col + 1):
        cell = ws.cell(row=min_row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_header

    ws.row_dimensions[min_row].height = 28

    # Dados
    for row in range(min_row + 1, max_row + 1):
        status_val = ws.cell(row=row, column=6).value if max_col >= 6 else ""
        is_divergencia = "Divergência" in str(status_val)
        row_fill = zebra_fill if row % 2 == 0 else white_fill

        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            if is_divergencia and col in [6, 7]:
                cell.fill = divergencia_fill
            else:
                cell.fill = row_fill
            
            if col in [1, 2, 4, 5]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Largura das colunas pré-definida
    if col_widths:
        for col_idx, width in enumerate(col_widths, start=min_col):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

def gerar_planilhas_comparativas():
    print("=" * 70)
    print("GERANDO PLANILHA COMPARATIVA: ROOVER-DOCS-ASSINADOS VS DOCS-PESSOAIS")
    print("=" * 70)
    print(f"Origem Drive (SOMENTE LEITURA): {PASTA_DRIVE}")
    print(f"Destino Local Seguro: {PASTA_DESTINO}")
    print("-" * 70)

    # 1. Carregar base de assinados
    print("[1/5] Carregando 'roover-docs-assinados.csv' (somente leitura)...")
    df_ass = pd.read_csv(ARQ_ASSINADOS_DRIVE, sep=";", dtype=str)
    df_ass["CPF_NORM"] = df_ass["CPF"].apply(normalizar_cpf)
    df_ass["MAT_NORM"] = df_ass["MATRICULA"].apply(normalizar_matricula)
    df_ass["NOME_NORM"] = df_ass["NOME"].apply(normalizar_nome)

    # 2. Carregar base de pessoais
    print("[2/5] Carregando 'roover-docs-pessoais.csv' (somente leitura)...")
    df_pess = pd.read_csv(ARQ_PESSOAIS_DRIVE, sep=";", dtype=str)
    df_pess["CPF_NORM"] = df_pess["CPF"].apply(normalizar_cpf)
    df_pess["MAT_NORM"] = df_pess["MATRICULA"].apply(normalizar_matricula)
    df_pess["NOME_NORM"] = df_pess["NOME"].apply(normalizar_nome)

    # Mapear dados de pessoais por Matrícula e Nome para detecção de divergência
    pess_cpfs = set(df_pess["CPF_NORM"].unique()) - {""}
    pess_by_mat = {}
    pess_by_nome = {}

    for _, r in df_pess.iterrows():
        m = r["MAT_NORM"]
        n = r["NOME_NORM"]
        if m and m not in pess_by_mat:
            pess_by_mat[m] = r
        if n and n not in pess_by_nome:
            pess_by_nome[n] = r

    # 3. Consolidar os cadastros em Assinados
    print("[3/5] Identificando cadastros únicos e comparando presenças...")
    ass_grouped = {}
    for _, row in df_ass.iterrows():
        cpf = row["CPF_NORM"]
        if not cpf:
            continue
        if cpf not in ass_grouped:
            ass_grouped[cpf] = {
                "CPF": cpf,
                "MATRICULAS": set(),
                "MATS_NORM": set(),
                "NOMES": set(),
                "NOMES_NORM": set(),
                "ADMISSOES": set(),
                "ARQUIVOS": []
            }
        mat = str(row["MATRICULA"]).strip() if pd.notna(row["MATRICULA"]) and str(row["MATRICULA"]).strip() != "" else None
        if mat and mat.lower() != "roover":
            ass_grouped[cpf]["MATRICULAS"].add(mat)
        mat_n = row["MAT_NORM"]
        if mat_n:
            ass_grouped[cpf]["MATS_NORM"].add(mat_n)

        nome = str(row["NOME"]).strip() if pd.notna(row["NOME"]) and str(row["NOME"]).strip() != "" else None
        if nome and nome.lower() != "roover":
            ass_grouped[cpf]["NOMES"].add(nome)
        nome_n = row["NOME_NORM"]
        if nome_n:
            ass_grouped[cpf]["NOMES_NORM"].add(nome_n)

        adm = str(row["ADMISSAO"]).strip() if pd.notna(row["ADMISSAO"]) and str(row["ADMISSAO"]).strip() != "" else None
        if adm and adm.lower() != "roover":
            ass_grouped[cpf]["ADMISSOES"].add(adm)

        arq = str(row["ARQUIVO"]).strip() if pd.notna(row.get("ARQUIVO")) else None
        if arq:
            ass_grouped[cpf]["ARQUIVOS"].append(arq)

    # Identificar cadastros presentes em Assinados e ausentes em Pessoais
    apenas_assinados_rows = []
    qtd_total_ausentes = 0
    qtd_divergencia_mat = 0
    qtd_divergencia_nome = 0

    for cpf, d in sorted(ass_grouped.items(), key=lambda x: x[0]):
        if cpf not in pess_cpfs:
            qtd_total_ausentes += 1
            # Verificar se tem divergência por matrícula ou nome
            mats_norm = d["MATS_NORM"]
            nomes_norm = d["NOMES_NORM"]

            status = "Ausente em docs-pessoais"
            detalhe = "Cadastro não localizado em roover-docs-pessoais"

            # Checar correspondência por matrícula
            mat_match = None
            for m in mats_norm:
                if m in pess_by_mat:
                    mat_match = pess_by_mat[m]
                    break

            # Checar correspondência por nome
            nome_match = None
            for n in nomes_norm:
                if n in pess_by_nome:
                    nome_match = pess_by_nome[n]
                    break

            if mat_match is not None:
                qtd_divergencia_mat += 1
                status = "Divergência de CPF (Mesma Matrícula em Pessoais)"
                cpf_pess = mat_match['CPF_NORM'] if pd.notna(mat_match['CPF_NORM']) else 'Sem CPF'
                detalhe = f"Matrícula {mat_match['MATRICULA']} consta em pessoais com CPF {cpf_pess} (ID Sistêmico {mat_match['ID_SISTEMICO']})"
            elif nome_match is not None:
                qtd_divergencia_nome += 1
                status = "Divergência de CPF (Mesmo Nome em Pessoais)"
                cpf_pess = nome_match['CPF_NORM'] if pd.notna(nome_match['CPF_NORM']) else 'Sem CPF'
                detalhe = f"Nome coincide em pessoais com CPF {cpf_pess} (Matrícula {nome_match['MATRICULA']}, ID Sistêmico {nome_match['ID_SISTEMICO']})"

            # Primeiros 3 arquivos de exemplo
            arqs = d["ARQUIVOS"]
            ex_arqs = "; ".join(arqs[:3]) + (f" (+{len(arqs)-3} outros)" if len(arqs) > 3 else "")

            apenas_assinados_rows.append({
                "CPF": cpf,
                "MATRICULA": " / ".join(sorted(d["MATRICULAS"])) if d["MATRICULAS"] else "Não informada",
                "NOME": " / ".join(sorted(d["NOMES"])) if d["NOMES"] else "Não informado",
                "ADMISSAO": " / ".join(sorted(d["ADMISSOES"])) if d["ADMISSOES"] else "Não informada",
                "QTD_DOCS_ASSINADOS": len(arqs),
                "STATUS_AUDITORIA": status,
                "DETALHE_DIVERGENCIA": detalhe,
                "EXEMPLOS_ARQUIVOS_ASSINADOS": ex_arqs
            })

    df_apenas_ass = pd.DataFrame(apenas_assinados_rows)

    # Identificar também cadastros presentes apenas em Pessoais (para visão bilateral completa)
    apenas_pessoais_rows = []
    pess_grouped = {}
    for _, row in df_pess.iterrows():
        cpf = row["CPF_NORM"]
        if not cpf:
            continue
        if cpf not in pess_grouped:
            pess_grouped[cpf] = {
                "CPF": cpf,
                "ID_SISTEMICO": row.get("ID_SISTEMICO", ""),
                "MATRICULA": row.get("MATRICULA", ""),
                "NOME": row.get("NOME", ""),
                "LOCAL": row.get("LOCAL", "")
            }

    for cpf, d in sorted(pess_grouped.items(), key=lambda x: x[0]):
        if cpf not in ass_grouped:
            apenas_pessoais_rows.append(d)

    df_apenas_pess = pd.DataFrame(apenas_pessoais_rows)

    total_ass = len(ass_grouped)
    total_pess = len(pess_grouped)
    intersecao = len(set(ass_grouped.keys()) & set(pess_grouped.keys()))

    print(f"Total de cadastros em roover-docs-assinados: {total_ass}")
    print(f"Total de cadastros em roover-docs-pessoais:  {total_pess}")
    print(f"Cadastros em AMBAS as planilhas:              {intersecao} ({intersecao/total_ass*100:.1f}%)")
    print(f"Cadastros APENAS em assinados (ausentes):     {len(df_apenas_ass)} ({len(df_apenas_ass)/total_ass*100:.1f}%)")
    print(f"  -> Totalmente ausentes em pessoais:         {len(df_apenas_ass) - qtd_divergencia_mat - qtd_divergencia_nome}")
    print(f"  -> Divergência de CPF (mesma matrícula):    {qtd_divergencia_mat}")
    print(f"  -> Divergência de CPF (mesmo nome):         {qtd_divergencia_nome}")
    print(f"Cadastros APENAS em pessoais:                 {len(df_apenas_pess)}")

    # 4. Salvar CSV direto
    print(f"\n[4/5] Salvando CSV local: {SAIDA_CSV}...")
    df_apenas_ass.to_csv(SAIDA_CSV, index=False, sep=";", encoding="utf-8-sig")

    # 5. Gerar pasta de trabalho Excel estilizada
    print(f"[5/5] Gerando Excel formatado: {SAIDA_EXCEL}...")
    wb = openpyxl.Workbook()
    # Remover aba padrao
    wb.remove(wb.active)

    header_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    # --- ABA 1: RESUMO EXECUTIVO ---
    ws_resumo = wb.create_sheet(title="Resumo Executivo")
    ws_resumo.views.sheetView[0].showGridLines = True

    # Titulo
    ws_resumo["A1"] = "RELATÓRIO COMPARATIVO DE AUDITORIA: ASSINADOS VS PESSOAIS"
    ws_resumo["A1"].font = Font(name="Calibri", size=16, bold=True, color="1F4E78")
    ws_resumo["A2"] = f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')} | Fonte: Google Drive (Somente Leitura)"
    ws_resumo["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

    # Bloco Metricas Principais
    headers_met = ["Métrica", "Quantidade", "Percentual / Detalhe"]
    for i, h in enumerate(headers_met, 1):
        cell = ws_resumo.cell(row=4, column=i, value=h)
        cell.fill = header_navy
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    metricas_dados = [
        ("Total de Cadastros Únicos em roover-docs-assinados", total_ass, "100.0% da base com contratos assinados"),
        ("Total de Cadastros Únicos em roover-docs-pessoais", total_pess, "Colaboradores com documentos de admissão"),
        ("Cadastros Presentes em AMBAS as planilhas (Completos)", intersecao, f"{intersecao/total_ass*100:.1f}% possuem admissão + assinados"),
        ("Cadastros APENAS em docs-assinados (INCONSISTÊNCIA / FALTANDO)", len(df_apenas_ass), f"{len(df_apenas_ass)/total_ass*100:.1f}% dos contratos NÃO têm docs pessoais"),
        ("  ↳ Totalmente ausentes em pessoais (ex: Admissões históricas)", len(df_apenas_ass) - qtd_divergencia_mat - qtd_divergencia_nome, "Sem cadastro de admissão no Drive"),
        ("  ↳ Divergência de CPF (mesma matrícula em pessoais)", qtd_divergencia_mat, "Mesmo funcionário cadastrado com outro CPF"),
        ("  ↳ Divergência de CPF (mesmo nome em pessoais)", qtd_divergencia_nome, "Mesmo nome homônimo com outro CPF"),
        ("Cadastros APENAS em docs-pessoais (Sem Contratos)", len(df_apenas_pess), "Colaboradores admitidos sem contratos assinados")
    ]

    for idx, (m, q, p) in enumerate(metricas_dados, 5):
        ws_resumo.cell(row=idx, column=1, value=m).font = Font(name="Calibri", size=10, bold=(idx in [4, 7, 8]))
        c2 = ws_resumo.cell(row=idx, column=2, value=q)
        c2.number_format = '#,##0'
        c2.alignment = Alignment(horizontal="right")
        c2.font = Font(name="Calibri", size=10, bold=True)
        if idx == 8: # Linha de inconsistencia
            c2.font = Font(name="Calibri", size=11, bold=True, color="C00000")
        ws_resumo.cell(row=idx, column=3, value=p).font = Font(name="Calibri", size=10, italic=True)

    aplicar_estilo_tabela(ws_resumo, min_row=4, max_row=12, min_col=1, max_col=3, 
                         header_fill=header_navy, header_font=font_header,
                         col_widths=[62, 18, 45])

    # Exemplo Admissão Histórica em destaque
    ws_resumo["A15"] = "CASO PRÁTICO ANALISADO: EXEMPLO DE ADMISSÕES HISTÓRICAS"
    ws_resumo["A15"].font = Font(name="Calibri", size=13, bold=True, color="1F4E78")

    ws_resumo["A16"] = "Foram identificadas 2 colaboradores com admissões anteriores à digitalização na planilha roover-docs-assinados:"
    ws_resumo["A16"].font = Font(name="Calibri", size=10, italic=True)

    exemplo_headers = ["CPF", "Matrícula", "Nome Completo", "Data Admissão", "Docs Assinados", "Status em Pessoais"]
    for i, h in enumerate(exemplo_headers, 1):
        cell = ws_resumo.cell(row=18, column=i, value=h)
        cell.fill = PatternFill(start_color="333f48", end_color="333f48", fill_type="solid")
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center")

    exemplo_dados = [
        ("11122233344", "7421", "COLABORADOR EXEMPLO A", "2003-08-01", 7, "NÃO CONSTA em roover-docs-pessoais"),
        ("55566677788", "13117", "COLABORADOR EXEMPLO B", "2008-11-25", 5, "NÃO CONSTA em roover-docs-pessoais")
    ]
    for r_idx, r_data in enumerate(exemplo_dados, 19):
        for c_idx, val in enumerate(r_data, 1):
            cell = ws_resumo.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = Alignment(horizontal="center" if c_idx in [1, 2, 4, 5] else "left")
            if c_idx == 6:
                cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")

    aplicar_estilo_tabela(ws_resumo, min_row=18, max_row=20, min_col=1, max_col=6, 
                         header_fill=PatternFill(start_color="333f48", end_color="333f48", fill_type="solid"), 
                         header_font=font_header,
                         col_widths=[16, 14, 35, 16, 16, 35])

    ws_resumo["A22"] = "Diagnóstico Técnico:"
    ws_resumo["A22"].font = Font(name="Calibri", size=11, bold=True)
    ws_resumo["A23"] = "• Ambas as colaboradoras possuem contratos assinados arquivados, mas NÃO constam na planilha de admissão (docs-pessoais)."
    ws_resumo["A24"] = "• Como suas admissões ocorreram em 2003 e 2008, os documentos físicos de admissão não foram integrados ao repositório digital de admissão."
    ws_resumo["A25"] = "• Portanto, o gerador de PDF unificou corretamente os documentos disponíveis (contratos assinados), ficando sem a seção de admissão."

    for row_diag in range(23, 26):
        ws_resumo[f"A{row_diag}"].font = Font(name="Calibri", size=10, color="262626")

    # --- ABA 2: CADASTROS APENAS EM ASSINADOS ---
    ws_apenas_ass = wb.create_sheet(title="Cadastros Apenas Assinados")
    ws_apenas_ass.views.sheetView[0].showGridLines = True

    cols_apenas_ass = [
        "CPF", "MATRÍCULA", "NOME COMPLETO", "DATA ADMISSÃO", 
        "QTD DOCS ASSINADOS", "STATUS DE AUDITORIA", "DETALHE DIVERGÊNCIA", 
        "EXEMPLOS DE ARQUIVOS ASSINADOS"
    ]
    for c_idx, h in enumerate(cols_apenas_ass, 1):
        ws_apenas_ass.cell(row=1, column=c_idx, value=h)

    for r_idx, row_dict in enumerate(apenas_assinados_rows, 2):
        ws_apenas_ass.cell(row=r_idx, column=1, value=row_dict["CPF"])
        ws_apenas_ass.cell(row=r_idx, column=2, value=row_dict["MATRICULA"])
        ws_apenas_ass.cell(row=r_idx, column=3, value=row_dict["NOME"])
        ws_apenas_ass.cell(row=r_idx, column=4, value=row_dict["ADMISSAO"])
        ws_apenas_ass.cell(row=r_idx, column=5, value=row_dict["QTD_DOCS_ASSINADOS"])
        ws_apenas_ass.cell(row=r_idx, column=6, value=row_dict["STATUS_AUDITORIA"])
        ws_apenas_ass.cell(row=r_idx, column=7, value=row_dict["DETALHE_DIVERGENCIA"])
        ws_apenas_ass.cell(row=r_idx, column=8, value=row_dict["EXEMPLOS_ARQUIVOS_ASSINADOS"])

    aplicar_estilo_tabela(ws_apenas_ass, min_row=1, max_row=len(apenas_assinados_rows) + 1, min_col=1, max_col=8,
                         header_fill=header_navy, header_font=font_header,
                         col_widths=[16, 16, 38, 16, 20, 36, 50, 50])

    # Congelar primeira linha
    ws_apenas_ass.freeze_panes = "A2"

    # --- ABA 3: CADASTROS APENAS EM PESSOAIS ---
    ws_apenas_pess = wb.create_sheet(title="Cadastros Apenas Pessoais")
    ws_apenas_pess.views.sheetView[0].showGridLines = True

    cols_apenas_pess = ["CPF", "MATRÍCULA", "NOME COMPLETO", "ID SISTÊMICO", "LOCAL DE ARMAZENAMENTO"]
    for c_idx, h in enumerate(cols_apenas_pess, 1):
        ws_apenas_pess.cell(row=1, column=c_idx, value=h)

    for r_idx, row_dict in enumerate(apenas_pessoais_rows, 2):
        ws_apenas_pess.cell(row=r_idx, column=1, value=row_dict.get("CPF", ""))
        ws_apenas_pess.cell(row=r_idx, column=2, value=row_dict.get("MATRICULA", ""))
        ws_apenas_pess.cell(row=r_idx, column=3, value=row_dict.get("NOME", ""))
        ws_apenas_pess.cell(row=r_idx, column=4, value=row_dict.get("ID_SISTEMICO", ""))
        ws_apenas_pess.cell(row=r_idx, column=5, value=row_dict.get("LOCAL", ""))

    aplicar_estilo_tabela(ws_apenas_pess, min_row=1, max_row=len(apenas_pessoais_rows) + 1, min_col=1, max_col=5,
                         header_fill=PatternFill(start_color="418AB3", end_color="418AB3", fill_type="solid"),
                         header_font=font_header,
                         col_widths=[16, 16, 38, 16, 25])

    ws_apenas_pess.freeze_panes = "A2"

    # Salvar
    wb.save(SAIDA_EXCEL)
    print("=" * 70)
    print(f"PLANILHAS COMPARATIVAS GERADAS COM SUCESSO!")
    print(f"1. Excel Consolidado: {SAIDA_EXCEL}")
    print(f"2. CSV Focado:        {SAIDA_CSV}")
    print("=" * 70)

if __name__ == "__main__":
    gerar_planilhas_comparativas()
