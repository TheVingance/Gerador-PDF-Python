import os
import sys
import argparse

# Garante suporte a UTF-8 e auto-flush
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Caminho padrão da pasta no Google Drive sincronizado
DEFAULT_DRIVE_FOLDER = r"G:\.shortcut-targets-by-id\11GxyE8SbWQbCXzoP8f3e9cGFfVK0TIFn\Roover\Arquivos Gerados\pdfs_gerados"

def analisar_duplicados(pasta: str):
    """
    Analisa a pasta e identifica arquivos com sufixo estrito ' (1).pdf',
    verificando se o par original existe. Execução ultrarrápida em memória O(1).
    """
    if not os.path.exists(pasta):
        print(f"[ERRO] Pasta não encontrada: {pasta}", flush=True)
        return None, None, None

    print(f"\n[INFO] Lendo arquivos da pasta no Google Drive...", flush=True)
    todos_arquivos = os.listdir(pasta)
    total_itens = len(todos_arquivos)
    print(f"[INFO] Total de arquivos encontrados: {total_itens:,}".replace(",", "."), flush=True)

    # Conjunto para busca O(1) instantânea em memória
    set_arquivos = set(todos_arquivos)

    # 1. Filtro estrito: DEVE terminar exatamente com ' (1).pdf'
    candidatos_duplicados = [f for f in todos_arquivos if f.endswith(" (1).pdf")]
    
    seguros_para_remover = []
    sem_original_alerta = []

    for dup in candidatos_duplicados:
        # Nome esperado do arquivo original sem o ' (1)'
        nome_original = dup[:-8] + ".pdf"

        if nome_original in set_arquivos:
            seguros_para_remover.append({
                "dup_nome": dup,
                "dup_path": os.path.join(pasta, dup),
                "orig_nome": nome_original,
                "orig_path": os.path.join(pasta, nome_original)
            })
        else:
            sem_original_alerta.append(dup)

    return total_itens, seguros_para_remover, sem_original_alerta


def main():
    parser = argparse.ArgumentParser(
        description="Limpeza Segura de Arquivos Duplicados ' (1).pdf' no Google Drive."
    )
    parser.add_argument(
        "--pasta",
        default=DEFAULT_DRIVE_FOLDER,
        help="Caminho da pasta pdfs_gerados no Google Drive"
    )
    parser.add_argument(
        "--executar",
        action="store_true",
        help="Executa a exclusão dos duplicados seguros (sem esta flag, roda em modo SIMULAÇÃO)"
    )
    parser.add_argument(
        "--forcar",
        action="store_true",
        help="Pula o prompt interativo de confirmação (útil para automação não interativa)"
    )

    args = parser.parse_args()

    total_itens, seguros, alertas = analisar_duplicados(args.pasta)
    if total_itens is None:
        return

    print("\n" + "=" * 65, flush=True)
    print("📋 RELATÓRIO DE AUDITORIA DE DUPLICADOS", flush=True)
    print("=" * 65, flush=True)
    print(f"Total de itens na pasta:               {total_itens:,}".replace(",", "."), flush=True)
    print(f"Total de arquivos com ' (1).pdf':      {len(seguros) + len(alertas):,}".replace(",", "."), flush=True)
    print(f"✅ Duplicados seguros para remoção:     {len(seguros):,}".replace(",", "."), flush=True)
    print(f"⚠️ Alertas (sem par original):          {len(alertas):,}".replace(",", "."), flush=True)

    if alertas:
        print("\n⚠️ ATENÇÃO: Os seguintes arquivos NÃO possuem o original correspondente", flush=True)
        print("   e NÃO serão tocados por motivo de segurança:", flush=True)
        for nome in alertas[:10]:
            print(f"   - {nome}", flush=True)

    if not seguros:
        print("\n✨ Nenhum arquivo duplicado ' (1).pdf' encontrado para remoção.", flush=True)
        return

    print("\n🔍 Exemplos de arquivos identificados para remoção e preservação:", flush=True)
    for item in seguros[:5]:
        print(f"   ❌ Excluir:   {item['dup_nome']}", flush=True)
        print(f"   ✅ Preservar: {item['orig_nome']}", flush=True)
        print("   " + "-" * 55, flush=True)

    # Decisão de Execução
    if not args.executar:
        print("\n" + "=" * 65, flush=True)
        print("🛡️ MODO SIMULAÇÃO (DRY RUN) CONCLUÍDO COM SUCESSO!", flush=True)
        print("Nenhum arquivo foi modificado ou excluído.", flush=True)
        print(f"Total validado: {len(seguros):,} arquivos duplicados prontos para remoção com segurança.".replace(",", "."), flush=True)
        print("=" * 65 + "\n", flush=True)
        return

    # Confirmação de segurança
    print("\n" + "!" * 65, flush=True)
    print("⚠️ ATENÇÃO: MODO DE EXCLUSÃO ATIVADO!", flush=True)
    print(f"Serão excluídos {len(seguros):,} arquivos duplicados no Google Drive.".replace(",", "."), flush=True)
    print("Os arquivos originais serão 100% PRESERVADOS.", flush=True)
    print("!" * 65, flush=True)
    
    if not args.forcar:
        resposta = input("\nDigite 'CONFIRMAR' para prosseguir com a remoção: ").strip()
        if resposta != "CONFIRMAR":
            print("[CANCELADO] Operação cancelada. Nenhum arquivo foi alterado.", flush=True)
            return

    print("\n🗑️  Iniciando remoção segura dos duplicados...", flush=True)
    removidos = 0
    erros = 0

    for i, item in enumerate(seguros, 1):
        try:
            os.remove(item["dup_path"])
            removidos += 1
        except Exception as e:
            erros += 1
            if erros <= 5:
                print(f"[ERRO] Falha ao remover {item['dup_nome']}: {e}", flush=True)

        if i % 500 == 0 or i == len(seguros):
            print(f"   Progresso: {i:,}/{len(seguros):,} ({i/len(seguros)*100:.1f}%)".replace(",", "."), flush=True)

    print("\n" + "=" * 65, flush=True)
    print("🎉 LIMPEZA CONCLUÍDA COM SUCESSO!", flush=True)
    print("=" * 65, flush=True)
    print(f"Arquivos duplicados removidos: {removidos:,}".replace(",", "."), flush=True)
    print(f"Erros encontrados:             {erros:,}".replace(",", "."), flush=True)
    print(f"Arquivos originais mantidos:   {len(seguros):,}".replace(",", "."), flush=True)
    print("=" * 65 + "\n", flush=True)

if __name__ == "__main__":
    main()
