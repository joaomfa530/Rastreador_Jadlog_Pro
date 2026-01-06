import logging
import sys
from logging.handlers import RotatingFileHandler
from .config import LOGS_DIR

def setup_logger(name="RastreadorLogger"):
    # Garante que a pasta de logs existe
    LOGS_DIR.mkdir(exist_ok=True)
    
    # Cria o logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    
    # Formato da mensagem: [DATA] [NIVEL] Mensagem
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # 1. Handler para Arquivo (Salva o histórico, rotaciona a cada 5MB)
    file_handler = RotatingFileHandler(
        LOGS_DIR / "app.log", maxBytes=5*1024*1024, backupCount=3, encoding='utf-8'
    )
    file_handler.setFormatter(formatter)
    
    # 2. Handler para Tela (Console)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    
    # Adiciona ao logger apenas se ele não tiver handlers ainda (evita duplicação)
    if not logger.handlers:
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
    return logger
logger = setup_logger()