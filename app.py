import concurrent.futures
import time
import signal
import sys
import pandas as pd
from tqdm import tqdm
from datetime import datetime

# Importações do Core
try:
    from src.config import INPUT_FILE
    from src.logger import logger
    from src.excel_manager import ExcelManager
    from src.tracker_service import JadlogTracker
except ImportError as e:
    print(f"❌ Erro crítico de importação: {e}")
    sys.exit(1)

# Configurações de Performance
MAX_WORKERS = 10 

# Variável global para garantir salvamento em caso de interrupção abrupta
resultados_globais = []

def signal_handler(sig, frame):
    """Captura Ctrl+C para salvar dados antes de sair."""
    print("\n\n⚠️ Interrupção detectada! Salvando dados processados até agora...")
    salvar_emergencia()
    sys.exit(0)

def salvar_emergencia():
    """Função de salvamento de segurança."""
    if resultados_globais:
        arquivo_rescue = f"data/resgate_rastreio_{datetime.now().strftime('%H%M%S')}.xlsx"
        try:
            pd.DataFrame(resultados_globais).to_excel(arquivo_rescue, index=False)
            print(f"💾 DADOS SALVOS EM EMERGÊNCIA: {arquivo_rescue}")
        except Exception as e:
            print(f"❌ Falha ao salvar arquivo de resgate: {e}")
    else:
        print("ℹ️ Nenhum dado para salvar.")

# Registra o listener de Ctrl+C
signal.signal(signal.SIGINT, signal_handler)

def main():
    global resultados_globais
    
    print("\n" + "="*60)
    print("🚀 SISTEMA DE RASTREAMENTO JADLOG - CLI MODE")
    print("="*60 + "\n")
    logger.info("Iniciando processo CLI...")

    # 1. Inicialização
    excel_mgr = ExcelManager()
    tracker = JadlogTracker()

    # 2. Leitura
    print(f"📂 Lendo arquivo de entrada: {INPUT_FILE}...")
    try:
        codigos = excel_mgr.carregar_dados()
    except Exception as e:
        logger.critical(f"Erro ao abrir Excel: {e}")
        return

    if not codigos:
        logger.warning("Lista de códigos vazia.")
        print("❌ Nenhum código encontrado. Verifique a coluna no Excel.")
        return

    total_pacotes = len(codigos)
    print(f"📦 Carga detectada: {total_pacotes} encomendas.")
    print(f"⚡ Workers Ativos: {MAX_WORKERS}")

    # 3. Processamento Paralelo Seguro
    start_time = time.time()
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # Mapeia futures
        future_to_code = {
            executor.submit(tracker.consultar_encomenda, code): code 
            for code in codigos
        }

        # Barra de progresso profissional
        with tqdm(total=total_pacotes, unit="pkg", desc="Rastreando", dynamic_ncols=True) as pbar:
            for future in concurrent.futures.as_completed(future_to_code):
                codigo = future_to_code[future]
                try:
                    data = future.result()
                    
                    # Validação básica de integridade
                    if not data:
                        raise ValueError("Retorno vazio do Tracker")
                        
                    resultados_globais.append(data)
                    
                    # Atualiza descrição da barra com último status (opcional, visualmente legal)
                    status_curto = str(data.get('Status', 'ND'))[:15]
                    pbar.set_postfix(last=f"{codigo}: {status_curto}")

                except Exception as exc:
                    logger.error(f"Falha em {codigo}: {exc}")
                    resultados_globais.append({
                        "CODIGO": codigo,
                        "Status": "ERRO PROCESSAMENTO",
                        "Detalhes": str(exc)
                    })
                
                finally:
                    pbar.update(1)

    # 4. Finalização e Relatório
    duration = time.time() - start_time
    print(f"\n✅ Processamento finalizado em {duration:.2f} segundos.")
    print(f"📊 Sucesso: {len(resultados_globais)}/{total_pacotes}")

    print("💾 Gerando relatório final Excel...")
    try:
        excel_mgr.salvar_relatorio(resultados_globais)
        print("✅ Relatório salvo com sucesso na pasta de saída.")
    except Exception as e:
        logger.error(f"Erro ao salvar relatório final: {e}")
        salvar_emergencia() # Tenta salvar o backup bruto

if __name__ == "__main__":
    main()