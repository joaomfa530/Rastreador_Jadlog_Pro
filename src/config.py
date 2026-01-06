import os
from pathlib import Path
from dotenv import load_dotenv
from datetime import datetime

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

# --- CAMINHOS DO SISTEMA ---
# Define a raiz do projeto de forma dinâmica
ROOT_DIR = Path(__file__).parent.parent

# Pastas principais
DATA_DIR = ROOT_DIR / "data"
LOGS_DIR = ROOT_DIR / "logs"

# Arquivos Específicos (AQUI ESTAVA O ERRO, ESTAS LINHAS PRECISAM EXISTIR)
INPUT_FILE = DATA_DIR / "base.xlsx"

# Gera nome dinâmico: resultado_rastreio_2023-10-27_15-30.xlsx
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
OUTPUT_FILE = DATA_DIR / f"resultado_rastreio_{timestamp}.xlsx"

# --- CONFIGURAÇÕES DA JADLOG ---
JADLOG_USER = os.getenv("JADLOG_USER")
JADLOG_PASSWORD = os.getenv("JADLOG_PASSWORD")
JADLOG_URL = os.getenv("JADLOG_URL", "https://www.jadlog.com.br/tracking")

# Validação Básica
if not JADLOG_USER or not JADLOG_PASSWORD:
    # Não vamos travar o código aqui se estivermos apenas testando, 
    # mas em produção isso seria um erro.
    pass 

# --- CONTROLE DE AMBIENTE ---
# Se True, usa dados falsos para teste.
SIMULATION_MODE = True