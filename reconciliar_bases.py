import os
import pandas as pd

# ORIGEM (Apenas leitura no Google Drive - NUNCA MODIFICAR)
PASTA_DRIVE = r"G:\.shortcut-targets-by-id\11GxyE8SbWQbCXzoP8f3e9cGFfVK0TIFn\Roover"

# DESTINO LOCAL SEGURO
PASTA_DESTINO = r"C:\Users\hiurkiv\Documents\roover-python"

def normalizar_cpf(val):
    if pd.isna(val):
        return ""
    digits = "".join([c for c in str(val) if c.isdigit()])
    if digits:
        return digits.zfill(11)
    return ""

def eh_invalido(val):
    if val is None or pd.isna(val):
        return True
    s = str(val).strip()
    return s == "" or s.lower() in ["roover", "nan", "none", "null"]

def escolher_valor(val_pess, val_ass, fallback="roover"):
    if not eh_invalido(val_pess):
        return str(val_pess).strip()
    if not eh_invalido(val_ass):
        return str(val_ass).strip()
    return fallback

def executar_reconciliacao():
    print("=" * 60)
    print("INICIANDO RECONCILIACAO DAS PLANILHAS (SOMENTE LEITURA NO DRIVE)")
    print("=" * 60)

    # 1. Carregar planilha de pessoais
    arq_pessoais = os.path.join(PASTA_DRIVE, "roover-docs-pessoais.csv")
    print(f"[1/4] Lendo {arq_pessoais}...")
    df_pess = pd.read_csv(arq_pessoais, sep=";", dtype=str)
    df_pess["CPF_NORM"] = df_pess["CPF"].apply(normalizar_cpf)

    # 2. Carregar planilha de assinados
    arq_assinados = os.path.join(PASTA_DRIVE, "roover-docs-assinados.csv")
    print(f"[2/4] Lendo {arq_assinados}...")
    df_ass = pd.read_csv(arq_assinados, sep=";", dtype=str)
    df_ass["CPF_NORM"] = df_ass["CPF"].apply(normalizar_cpf)

    print("[3/4] Cruzando dados e resolvendo campos faltantes...")

    # Agrupar assinados por CPF_NORM para consolidar documentos e metadados
    ass_grouped = {}
    for _, row in df_ass.iterrows():
        cpf = row["CPF_NORM"]
        if not cpf:
            continue
        if cpf not in ass_grouped:
            ass_grouped[cpf] = {
                "MATRICULA": row["MATRICULA"] if not eh_invalido(row["MATRICULA"]) else None,
                "NOME": row["NOME"] if not eh_invalido(row["NOME"]) else None,
                "ADMISSAO": row["ADMISSAO"] if not eh_invalido(row["ADMISSAO"]) else None,
                "ARQUIVOS": []
            }
        else:
            if ass_grouped[cpf]["MATRICULA"] is None and not eh_invalido(row["MATRICULA"]):
                ass_grouped[cpf]["MATRICULA"] = row["MATRICULA"]
            if ass_grouped[cpf]["NOME"] is None and not eh_invalido(row["NOME"]):
                ass_grouped[cpf]["NOME"] = row["NOME"]
            if ass_grouped[cpf]["ADMISSAO"] is None and not eh_invalido(row["ADMISSAO"]):
                ass_grouped[cpf]["ADMISSAO"] = row["ADMISSAO"]

        arq = row.get("ARQUIVO")
        if arq and not eh_invalido(arq):
            ass_grouped[cpf]["ARQUIVOS"].append(str(arq).strip())

    # Agrupar pessoais por CPF_NORM
    pess_grouped = {}
    for _, row in df_pess.iterrows():
        cpf = row["CPF_NORM"]
        if not cpf:
            continue
        pess_grouped[cpf] = {
            "ID_SISTEMICO": row["ID_SISTEMICO"] if not eh_invalido(row["ID_SISTEMICO"]) else None,
            "NOME": row["NOME"] if not eh_invalido(row["NOME"]) else None,
            "MATRICULA": row["MATRICULA"] if not eh_invalido(row["MATRICULA"]) else None
        }

    # Unir todos os CPFs unicos
    todos_cpfs = sorted(list(set(list(pess_grouped.keys()) + list(ass_grouped.keys()))))
    print(f"Total de CPFs unicos consolidados: {len(todos_cpfs)}")

    consolidado = []
    qtd_nome_completado = 0
    qtd_matricula_completada = 0

    for cpf in todos_cpfs:
        dados_pess = pess_grouped.get(cpf, {})
        dados_ass = ass_grouped.get(cpf, {})

        pess_nome = dados_pess.get("NOME")
        ass_nome = dados_ass.get("NOME")
        if eh_invalido(pess_nome) and not eh_invalido(ass_nome):
            qtd_nome_completado += 1
        elif eh_invalido(ass_nome) and not eh_invalido(pess_nome):
            qtd_nome_completado += 1
        nome_final = escolher_valor(pess_nome, ass_nome, fallback="roover")

        pess_mat = dados_pess.get("MATRICULA")
        ass_mat = dados_ass.get("MATRICULA")
        if eh_invalido(pess_mat) and not eh_invalido(ass_mat):
            qtd_matricula_completada += 1
        elif eh_invalido(ass_mat) and not eh_invalido(pess_mat):
            qtd_matricula_completada += 1
        matricula_final = escolher_valor(pess_mat, ass_mat, fallback="roover")

        id_sistemico_final = dados_pess.get("ID_SISTEMICO") or "roover"
        admissao_final = dados_ass.get("ADMISSAO") or "roover"
        arqs_assinados = dados_ass.get("ARQUIVOS", [])

        consolidado.append({
            "CPF": cpf,
            "NOME": nome_final,
            "MATRICULA": matricula_final,
            "ID_SISTEMICO": id_sistemico_final,
            "ADMISSAO": admissao_final,
            "TOTAL_DOCS_ASSINADOS": len(arqs_assinados),
            "LISTA_DOCS_ASSINADOS": ";".join(arqs_assinados)
        })

    df_consolidado = pd.DataFrame(consolidado)

    # 4. Salvar EXCLUSIVAMENTE no destino local
    destino_csv = os.path.join(PASTA_DESTINO, "base_consolidada_roover.csv")
    print(f"[4/4] Salvando resultado unificado em {destino_csv}...")
    df_consolidado.to_csv(destino_csv, index=False, sep=";", encoding="utf-8-sig")

    total_com_id = (df_consolidado["ID_SISTEMICO"] != "roover").sum()
    total_com_ass = (df_consolidado["TOTAL_DOCS_ASSINADOS"] > 0).sum()
    total_ambos = ((df_consolidado["ID_SISTEMICO"] != "roover") & (df_consolidado["TOTAL_DOCS_ASSINADOS"] > 0)).sum()

    print("\n" + "=" * 60)
    print("RESUMO DA RECONCILIACAO")
    print("=" * 60)
    print(f"Total de registros na base consolidada: {len(df_consolidado)}")
    print(f"Campos com NOME enriquecidos por cruzamento: {qtd_nome_completado}")
    print(f"Campos com MATRICULA enriquecidos por cruzamento: {qtd_matricula_completada}")
    print(f"Registros com ID_SISTEMICO presente: {total_com_id}")
    print(f"Registros com Documentos Assinados: {total_com_ass}")
    print(f"Registros com ambos (Admissao + Assinados): {total_ambos}")
    print("=" * 60)
    print("CONCLUIDO COM SUCESSO!")
    print("=" * 60)

if __name__ == "__main__":
    executar_reconciliacao()
