import os
import re
import io
import sys
import time
import argparse
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    pass

# ORIGEM (Apenas leitura no Google Drive - NUNCA MODIFICAR)
PASTA_DRIVE = r"G:\.shortcut-targets-by-id\11GxyE8SbWQbCXzoP8f3e9cGFfVK0TIFn\Roover"
DIR_ADMISSAO = os.path.join(PASTA_DRIVE, "arquivos-admissao")
DIR_ASSINADOS = os.path.join(PASTA_DRIVE, "assinados")

# DESTINO LOCAL
PASTA_DESTINO = r"C:\Users\hiurkiv\Documents\roover-python"
DIR_SAIDA_PDFS = os.path.join(PASTA_DESTINO, "pdfs_gerados")
CSV_CONSOLIDADO = os.path.join(PASTA_DESTINO, "base_consolidada_roover.csv")
LISTA_ADMISSAO_TXT = os.path.join(PASTA_DESTINO, "lista_admissao.txt")

def sanitizar_nome_arquivo(nome):
    """Remove caracteres invalidos para nomes de arquivos no Windows."""
    if not nome:
        return "roover"
    limpo = re.sub(r'[\\/*?:"<>|]', "", str(nome))
    limpo = re.sub(r'\s+', " ", limpo).strip()
    return limpo if limpo else "roover"

def normalizar_cpf(val):
    if not val or pd.isna(val):
        return ""
    digits = "".join([c for c in str(val) if c.isdigit()])
    return digits.zfill(11) if digits else ""

def converter_imagem_para_pdf_bytes(caminho_imagem):
    """
    Converte imagem (JPG, PNG, JPEG) ajustando proporcionalmente ao tamanho padrao A4.
    Corrige tambem a orientacao via EXIF e centraliza sobre folha branca.
    """
    try:
        with Image.open(caminho_imagem) as img:
            # Corrige orientacao de fotos de celular que usam tag EXIF
            img = ImageOps.exif_transpose(img)
            if img.mode != "RGB":
                img = img.convert("RGB")

            # Padrao A4 a 150 DPI (1240 x 1754 px) -> exatamente 595 x 842 pt no PDF
            A4_W, A4_H = 1240, 1754
            is_landscape = img.width > img.height
            pw, ph = (A4_H, A4_W) if is_landscape else (A4_W, A4_H)

            # Redimensiona proporcionalmente para caber na folha sem distorcer
            img.thumbnail((pw, ph), Image.Resampling.LANCZOS)

            # Cria folha A4 branca e centraliza a imagem
            canvas = Image.new("RGB", (pw, ph), (255, 255, 255))
            offset = ((pw - img.width) // 2, (ph - img.height) // 2)
            canvas.paste(img, offset)

            pdf_bytes = io.BytesIO()
            canvas.save(pdf_bytes, format="PDF", resolution=150.0, quality=85)
            pdf_bytes.seek(0)
            return pdf_bytes
    except Exception as e:
        print(f"    [AVISO] Falha ao converter imagem {os.path.basename(caminho_imagem)}: {e}", flush=True)
        return None

def carregar_mapa_admissao():
    """
    Carrega o mapa {ID_SISTEMICO: [arquivos]} a partir do arquivo local lista_admissao.txt.
    Executa em fracao de segundo e elimina chamadas pesadas ao Google Drive.
    """
    mapa = {}
    if not os.path.exists(LISTA_ADMISSAO_TXT):
        print(f"[AVISO] {LISTA_ADMISSAO_TXT} nao encontrado.", flush=True)
        return mapa

    t0 = time.time()
    total_linhas = 0
    with open(LISTA_ADMISSAO_TXT, "r", encoding="utf-8", errors="ignore") as f:
        for linha in f:
            nome = linha.strip()
            if not nome:
                continue
            total_linhas += 1
            if "_" in nome:
                sid = nome.split("_")[0]
                if sid not in mapa:
                    mapa[sid] = []
                mapa[sid].append(nome)
    
    print(f"Mapa de admissao carregado ({total_linhas} arquivos, {len(mapa)} IDs) em {time.time() - t0:.2f}s.", flush=True)
    return mapa

def gerar_pdf_pessoa(registro, dir_saida=DIR_SAIDA_PDFS, mapa_admissao=None, verbose=True, sobrescrever=False):
    """
    Gera o PDF unificado com a ordem:
    1º: Documentos de Admissao / Pessoais (arquivos-admissao)
    2º: Documentos Assinados (assinados)
    """
    os.makedirs(dir_saida, exist_ok=True)

    cpf = registro["CPF"]
    nome = registro["NOME"]
    matricula = registro["MATRICULA"]
    id_sistemico = registro["ID_SISTEMICO"]

    nome_seguro = sanitizar_nome_arquivo(nome)
    nome_pdf_final = f"{cpf}-{nome_seguro}-roover.pdf"
    caminho_pdf_final = os.path.join(dir_saida, nome_pdf_final)

    # Verificar se ja existe (capacidade de retomada/skip)
    if not sobrescrever and os.path.exists(caminho_pdf_final) and os.path.getsize(caminho_pdf_final) > 0:
        tamanho_mb = os.path.getsize(caminho_pdf_final) / (1024 * 1024)
        if verbose:
            print(f"  [IGNORADO] Arquivo ja existe: {nome_pdf_final} ({tamanho_mb:.2f} MB)", flush=True)
        return {
            "cpf": cpf,
            "matricula": matricula,
            "nome": nome,
            "status": "IGNORADO_EXISTENTE",
            "caminho": caminho_pdf_final,
            "paginas": 0,
            "tamanho_mb": round(tamanho_mb, 2),
            "docs_adm": 0,
            "docs_ass": 0,
            "erro": ""
        }

    if verbose:
        print("-" * 60, flush=True)
        print(f"Processando: {nome}", flush=True)
        print(f"  CPF: {cpf} | Matrícula: {matricula} | ID Sistêmico: {id_sistemico}", flush=True)
        print(f"  Arquivo destino: {nome_pdf_final}", flush=True)

    writer = PdfWriter()
    total_paginas = 0
    docs_admissao_incluidos = 0
    docs_assinados_incluidos = 0

    # -------------------------------------------------------------
    # 1º Bloco: Documentos de Admissao / Pessoais
    # -------------------------------------------------------------
    if id_sistemico and id_sistemico.lower() != "roover":
        arqs_adm = []
        if mapa_admissao is not None:
            arqs_adm = sorted(mapa_admissao.get(id_sistemico, []))
        
        if verbose:
            print(f"  [1/2] Documentos de admissao encontrados: {len(arqs_adm)}", flush=True)
        for arq in arqs_adm:
            caminho_completo = os.path.join(DIR_ADMISSAO, arq)
            ext = os.path.splitext(arq)[1].lower()
            try:
                if ext == ".pdf":
                    reader = PdfReader(caminho_completo)
                    for page in reader.pages:
                        writer.add_page(page)
                        total_paginas += 1
                    docs_admissao_incluidos += 1
                    if verbose:
                        print(f"    + Admissao (PDF): {arq} ({len(reader.pages)} paginas)", flush=True)
                elif ext in [".jpg", ".jpeg", ".png", ".bmp", ".webp", ".heic"]:
                    img_bytes = converter_imagem_para_pdf_bytes(caminho_completo)
                    if img_bytes:
                        reader = PdfReader(img_bytes)
                        for page in reader.pages:
                            writer.add_page(page)
                            total_paginas += 1
                        docs_admissao_incluidos += 1
                        if verbose:
                            print(f"    + Admissao (Imagem convertida): {arq}", flush=True)
            except Exception as e:
                if verbose:
                    print(f"    [ERRO] Nao foi possivel ler {arq}: {e}", flush=True)
    else:
        if verbose:
            print("  [1/2] Sem ID Sistemico para documentos de admissao.", flush=True)

    # -------------------------------------------------------------
    # 2º Bloco: Documentos Assinados
    # -------------------------------------------------------------
    lista_assinados_raw = registro.get("LISTA_DOCS_ASSINADOS", "")
    arqs_assinados = [a.strip() for a in str(lista_assinados_raw).split(";") if a.strip() and a.strip().lower() != "nan"] if pd.notna(lista_assinados_raw) else []
    if verbose:
        print(f"  [2/2] Documentos assinados registrados: {len(arqs_assinados)}", flush=True)

    for arq in arqs_assinados:
        caminho_completo = os.path.join(DIR_ASSINADOS, arq)
        if not os.path.exists(caminho_completo):
            if verbose:
                print(f"    [AVISO] Arquivo assinado nao encontrado no Drive: {arq}", flush=True)
            continue
        try:
            reader = PdfReader(caminho_completo)
            for page in reader.pages:
                writer.add_page(page)
                total_paginas += 1
            docs_assinados_incluidos += 1
            if verbose:
                print(f"    + Assinado (PDF): {arq} ({len(reader.pages)} paginas)", flush=True)
        except Exception as e:
            if verbose:
                print(f"    [ERRO] Falha ao processar PDF assinado {arq}: {e}", flush=True)

    # -------------------------------------------------------------
    # Salvar PDF unificado
    # -------------------------------------------------------------
    if total_paginas == 0:
        if verbose:
            print(f"  [ATENCAO] Nenhum documento valido encontrado para {nome}. Nenhum PDF gerado.", flush=True)
        return {
            "cpf": cpf,
            "matricula": matricula,
            "nome": nome,
            "status": "SEM_DOCUMENTOS",
            "caminho": "",
            "paginas": 0,
            "tamanho_mb": 0.0,
            "docs_adm": 0,
            "docs_ass": 0,
            "erro": "Nenhum documento encontrado ou legivel"
        }

    caminho_tmp = caminho_pdf_final + ".tmp"
    try:
        with open(caminho_tmp, "wb") as f_out:
            writer.write(f_out)
        if os.path.exists(caminho_pdf_final):
            os.remove(caminho_pdf_final)
        os.rename(caminho_tmp, caminho_pdf_final)
    except Exception as e:
        if os.path.exists(caminho_tmp):
            try:
                os.remove(caminho_tmp)
            except Exception:
                pass
        if verbose:
            print(f"    [ERRO] Falha ao salvar PDF final: {e}", flush=True)
        return {
            "cpf": cpf,
            "matricula": matricula,
            "nome": nome,
            "status": "ERRO",
            "caminho": "",
            "paginas": 0,
            "tamanho_mb": 0.0,
            "docs_adm": docs_admissao_incluidos,
            "docs_ass": docs_assinados_incluidos,
            "erro": str(e)
        }

    tamanho_mb = os.path.getsize(caminho_pdf_final) / (1024 * 1024)
    if verbose:
        print(f"  [SUCESSO] PDF gerado com sucesso!", flush=True)
        print(f"    - Salvo em: {caminho_pdf_final}", flush=True)
        print(f"    - Total de paginas: {total_paginas}", flush=True)
        print(f"    - Docs Admissao anexados: {docs_admissao_incluidos}", flush=True)
        print(f"    - Docs Assinados anexados: {docs_assinados_incluidos}", flush=True)
        print(f"    - Tamanho do arquivo: {tamanho_mb:.2f} MB", flush=True)
        print("-" * 60, flush=True)

    return {
        "cpf": cpf,
        "matricula": matricula,
        "nome": nome,
        "status": "SUCESSO",
        "caminho": caminho_pdf_final,
        "paginas": total_paginas,
        "tamanho_mb": round(tamanho_mb, 2),
        "docs_adm": docs_admissao_incluidos,
        "docs_ass": docs_assinados_incluidos,
        "erro": ""
    }

def processar_lote(df, mapa_adm, workers=4, limite=None, sobrescrever=False):
    """Executa a geracao de PDFs em lote para todos os colaboradores com paralelismo controlado."""
    if limite and limite > 0:
        df_alvo = df.head(limite).copy()
        print(f"Modo lote iniciado: limitando aos primeiros {limite} registros.", flush=True)
    else:
        df_alvo = df.copy()
        print(f"Modo lote iniciado: total de {len(df_alvo)} registros.", flush=True)

    print(f"Trabalhadores paralelos (workers): {workers}", flush=True)
    print(f"Sobrescrever existentes: {'Sim' if sobrescrever else 'Nao (pula)'}", flush=True)
    print("=" * 70, flush=True)

    total = len(df_alvo)
    concluidos = 0
    sucessos = 0
    ignorados = 0
    erros = 0
    sem_docs = 0

    lock_contador = threading.Lock()
    t_inicio = time.time()
    resultados = []

    def tarefa(reg_dict):
        nonlocal concluidos, sucessos, ignorados, erros, sem_docs
        try:
            res = gerar_pdf_pessoa(
                reg_dict,
                dir_saida=DIR_SAIDA_PDFS,
                mapa_admissao=mapa_adm,
                verbose=False,
                sobrescrever=sobrescrever
            )
        except Exception as e:
            res = {
                "cpf": reg_dict.get("CPF", ""),
                "matricula": reg_dict.get("MATRICULA", ""),
                "nome": reg_dict.get("NOME", ""),
                "status": "ERRO",
                "caminho": "",
                "paginas": 0,
                "tamanho_mb": 0.0,
                "docs_adm": 0,
                "docs_ass": 0,
                "erro": str(e)
            }

        with lock_contador:
            concluidos += 1
            pct = (concluidos / total) * 100
            status = res["status"]
            if status == "SUCESSO":
                sucessos += 1
                detalhe = f"OK ({res['paginas']} pags, {res['tamanho_mb']:.2f} MB)"
            elif status == "IGNORADO_EXISTENTE":
                ignorados += 1
                detalhe = f"JA EXISTE ({res['tamanho_mb']:.2f} MB)"
            elif status == "SEM_DOCUMENTOS":
                sem_docs += 1
                detalhe = "SEM DOCUMENTOS"
            else:
                erros += 1
                detalhe = f"ERRO: {res.get('erro', 'Falha')}"

            nome_curto = (res['nome'][:28] + "..") if len(res['nome']) > 30 else res['nome']
            print(f"[{concluidos:5d}/{total}] ({pct:5.1f}%) | {res['cpf']} | {nome_curto:<30} | {detalhe}", flush=True)

        return res

    registros = [r.to_dict() for _, r in df_alvo.iterrows()]

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(tarefa, reg) for reg in registros]
        for future in as_completed(futures):
            resultados.append(future.result())

    tempo_total = time.time() - t_inicio
    mins = int(tempo_total // 60)
    segs = int(tempo_total % 60)

    print("=" * 70, flush=True)
    print("RESUMO DA GERACAO EM LOTE:", flush=True)
    print(f"  Total processados: {total}", flush=True)
    print(f"  Sucessos: {sucessos}", flush=True)
    print(f"  Ignorados (ja existentes): {ignorados}", flush=True)
    print(f"  Sem documentos: {sem_docs}", flush=True)
    print(f"  Erros: {erros}", flush=True)
    print(f"  Tempo decorrido: {mins}m {segs}s", flush=True)
    print("=" * 70, flush=True)

    caminho_relatorio = os.path.join(PASTA_DESTINO, "relatorio_geracao.csv")
    try:
        df_rel = pd.DataFrame(resultados)
        df_rel.to_csv(caminho_relatorio, sep=";", index=False, encoding="utf-8-sig")
        print(f"Relatorio de execucao salvo em: {caminho_relatorio}", flush=True)
    except Exception as e:
        print(f"[AVISO] Nao foi possivel salvar relatorio CSV: {e}", flush=True)

def main():
    parser = argparse.ArgumentParser(description="Gerador de PDF Consolidado Roover")
    parser.add_argument("--cpf", type=str, help="CPF da pessoa (ex: 12345678901)")
    parser.add_argument("--matricula", type=str, help="Matrícula da pessoa")
    parser.add_argument("--todos", action="store_true", help="Gera os PDFs para todos os registros da base consolidada em lote")
    parser.add_argument("--workers", type=int, default=4, help="Quantidade de trabalhadores paralelos no lote (padrao: 4)")
    parser.add_argument("--limite", type=int, default=None, help="Limita a quantidade de registros a processar no lote (ex: --limite 10)")
    parser.add_argument("--sobrescrever", action="store_true", help="Sobrescreve os arquivos PDF existentes em vez de pular")
    args = parser.parse_args()

    if not os.path.exists(CSV_CONSOLIDADO):
        print(f"Base consolidada nao encontrada em {CSV_CONSOLIDADO}.", flush=True)
        print("Execute primeiro: python reconciliar_bases.py", flush=True)
        sys.exit(1)

    print("Carregando base consolidada...", flush=True)
    df = pd.read_csv(CSV_CONSOLIDADO, sep=";", dtype=str)

    # Carregar mapa de admissao a partir do arquivo texto local
    mapa_adm = carregar_mapa_admissao()

    # MODO LOTE
    if args.todos:
        processar_lote(
            df=df,
            mapa_adm=mapa_adm,
            workers=args.workers,
            limite=args.limite,
            sobrescrever=args.sobrescrever
        )
        return

    # MODO INDIVIDUAL
    filtro = None
    if args.cpf:
        cpf_buscado = normalizar_cpf(args.cpf)
        filtro = df[df["CPF"] == cpf_buscado]
    elif args.matricula:
        filtro = df[df["MATRICULA"] == str(args.matricula).strip()]
    else:
        entrada = input("Digite o CPF ou Matricula da pessoa para gerar o PDF: ").strip()
        cpf_norm = normalizar_cpf(entrada)
        filtro = df[(df["CPF"] == cpf_norm) | (df["MATRICULA"] == entrada)]

    if filtro is None or filtro.empty:
        print(f"Nenhum registro encontrado na base consolidada para a busca informada.", flush=True)
        sys.exit(1)

    registro = filtro.iloc[0].to_dict()
    gerar_pdf_pessoa(registro, mapa_admissao=mapa_adm, verbose=True, sobrescrever=args.sobrescrever)

if __name__ == "__main__":
    main()
