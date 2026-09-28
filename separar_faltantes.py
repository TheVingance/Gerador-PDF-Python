import os
import sys
import shutil
import re

# Garante suporte a UTF-8
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

DRIVE_FOLDER = r"G:\.shortcut-targets-by-id\11GxyE8SbWQbCXzoP8f3e9cGFfVK0TIFn\Roover\Arquivos Gerados\pdfs_gerados"
LOCAL_FOLDER = r"c:\Users\hiurkiv\Documents\roover-python\pdfs_gerados"
OUTPUT_DIR = r"c:\Users\hiurkiv\Documents\roover-python\faltantes_para_upload"

def main():
    if not os.path.exists(DRIVE_FOLDER):
        print(f"[ERRO] Pasta do Google Drive não encontrada: {DRIVE_FOLDER}")
        return

    if not os.path.exists(LOCAL_FOLDER):
        print(f"[ERRO] Pasta local não encontrada: {LOCAL_FOLDER}")
        return

    print("[INFO] Analisando arquivos no Google Drive e no computador local...", flush=True)

    drive_files = os.listdir(DRIVE_FOLDER)
    local_files = os.listdir(LOCAL_FOLDER)

    # Identificar duplicados com ' (1).pdf', ' (2).pdf' no Drive
    padrao_dup = re.compile(r" \(\d+\)\.pdf$")
    drive_unicos = {f for f in drive_files if not padrao_dup.search(f)}
    local_set = set(local_files)

    # Arquivos que faltam no Drive
    faltantes = sorted(list(local_set - drive_unicos))
    total_faltantes = len(faltantes)

    print(f"[INFO] Total esperado no local:     {len(local_set):,}".replace(",", "."))
    print(f"[INFO] Total único presente no Drive: {len(drive_unicos):,}".replace(",", "."))
    print(f"[INFO] Total de arquivos FALTANTES:   {total_faltantes:,}".replace(",", "."), flush=True)

    if total_faltantes == 0:
        print("\n🎉 PARABÉNS! Todos os 21.146 arquivos já estão presentes no Google Drive!")
        return

    # Limpar pasta de saída se já existir
    if os.path.exists(OUTPUT_DIR):
        print(f"[INFO] Limpando pasta anterior: {OUTPUT_DIR}...", flush=True)
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Salvar lista em TXT
    txt_path = os.path.join(OUTPUT_DIR, "lista_arquivos_faltantes.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(f"Total de arquivos faltantes: {total_faltantes}\n\n")
        for nome in faltantes:
            f.write(nome + "\n")
    print(f"[INFO] Lista completa salva em: {txt_path}", flush=True)

    # Separar em lotes organizados de 1.000 arquivos (evita travamento do navegador)
    TAMANHO_LOTE = 1000
    total_lotes = (total_faltantes + TAMANHO_LOTE - 1) // TAMANHO_LOTE

    print(f"[INFO] Organizando os {total_faltantes:,} arquivos em {total_lotes} lotes de até 1.000 itens...".replace(",", "."), flush=True)

    for lote_num in range(1, total_lotes + 1):
        inicio = (lote_num - 1) * TAMANHO_LOTE
        fim = min(inicio + TAMANHO_LOTE, total_faltantes)
        itens_lote = faltantes[inicio:fim]

        nome_pasta_lote = f"Lote_{lote_num:02d}_({len(itens_lote)}_arquivos)"
        pasta_lote = os.path.join(OUTPUT_DIR, nome_pasta_lote)
        os.makedirs(pasta_lote, exist_ok=True)

        for nome_arquivo in itens_lote:
            src = os.path.join(LOCAL_FOLDER, nome_arquivo)
            dst = os.path.join(pasta_lote, nome_arquivo)

            # Usar Hardlink (0 bytes adicionais de disco e criação instantânea)
            try:
                os.link(src, dst)
            except Exception:
                shutil.copy2(src, dst)

        print(f"   ✅ {nome_pasta_lote} criado com sucesso ({len(itens_lote)} arquivos)")

    print("\n" + "=" * 65)
    print("📁 PASTAS PRONTAS PARA ENVIO NO NAVEGADOR!")
    print("=" * 65)
    print(f"Caminho principal: {OUTPUT_DIR}")
    print("\nComo enviar sem travar o navegador:")
    print("1. Abra a pasta do primeiro lote:")
    print("   Lote_01_...")
    print("2. Pressione Ctrl + A e arraste todos os arquivos para o Google Drive no navegador.")
    print("3. Aguarde o lote concluir (leva ~1 a 2 minutos sem travamentos).")
    print("4. Repita para o próximo lote até o final.")
    print("=" * 65 + "\n")

if __name__ == "__main__":
    main()
