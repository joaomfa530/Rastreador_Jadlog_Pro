import time
import random
import requests
from datetime import datetime, timedelta
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from src.logger import logger
# Se tiver config.py, certifique-se que SIMULATION_MODE existe.
# Caso não tenha, vamos assumir True para a apresentação por segurança.
try:
    from src.config import SIMULATION_MODE
except ImportError:
    SIMULATION_MODE = True

class JadlogTracker:
    def __init__(self):
        """
        Inicializa a sessão HTTP com Retry Strategy (Resiliência Enterprise).
        Isso garante que, quando formos para produção, a rede instável não derrube o app.
        """
        self.session = requests.Session()
        
        # Estratégia de Retry: Tenta 3x com espera exponencial (Backoff)
        # Fundamental para scrapers e integrações de API
        retries = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[500, 502, 503, 504]
        )
        adapter = HTTPAdapter(max_retries=retries)
        self.session.mount('http://', adapter)
        self.session.mount('https://', adapter)
        
        self.token = None 

    def consultar_encomenda(self, codigo):
        """
        Orquestrador principal. Decide se vai para o Mock (Apresentação) ou Real (Prod).
        """
        # Se for apresentação, usamos o Mock Inteligente Determinístico
        if SIMULATION_MODE:
            return self._mock_response_deterministico(codigo)
        else:
            return self._api_response(codigo)

    def _mock_response_deterministico(self, codigo):
        """
        SIMULAÇÃO INTELIGENTE: Gera dados consistentes baseados no próprio código.
        Objetivo: Permitir uma apresentação controlada onde o apresentador sabe o resultado.
        """
        # Jitter de UX: Pequeno delay variável para a barra de progresso parecer natural
        time.sleep(random.uniform(0.1, 0.4))
        
        # Tratamento de entrada (Excel às vezes manda float/científico)
        try:
            if isinstance(codigo, float):
                codigo_str = str(int(codigo))
            else:
                codigo_str = str(codigo).strip()
            # Limpeza básica
            codigo_str = codigo_str.replace(".", "").replace("-", "").replace("/", "")
        except:
            codigo_str = str(codigo)

        # SEED: O segredo da consistência. Usamos o último dígito numérico.
        # Se o código termina em 1, sempre será o caso 1.
        try:
            seed = int(codigo_str[-1]) 
        except:
            seed = 5 # Fallback genérico

        # Base de Unidades Realistas da Jadlog
        unidades = [
            "CO SÃO PAULO 01", "HUB MATRIZ (SP)", "TECA RIO DE JANEIRO", 
            "CO CURITIBA 01", "CO PORTO ALEGRE", "CO SALVADOR BASE", 
            "CD MANAUS", "CO FORTALEZA", "AG RECIFE", "FL BELO HORIZONTE"
        ]

        # Lógica de Negócio Simulada (Baseada no final do CTE)
        
        # Cenário 1: Sucesso/Entregue (Finais 0, 1, 2, 3)
        if seed <= 3:
            status = "Entregue"
            detalhe = "Entrega realizada no destinatário (Baixa Mobile)."
            unidade = unidades[seed % 5]
            dias_postagem = 4
            dias_prevista = -1 # Ontem

        # Cenário 2: Em Trânsito Normal (Finais 4, 5, 6, 7)
        elif 4 <= seed <= 7:
            status = "Em Trânsito"
            detalhe = "Transferência entre unidades (Carreta)."
            unidade = unidades[(seed % 5) + 5]
            dias_postagem = 2
            dias_prevista = 2 # Daqui a 2 dias

        # Cenário 3: Problema/Atraso (Finais 8, 9)
        else:
            status = "Retido / Atrasado"
            detalhe = "Endereço não localizado / Número inexistente."
            unidade = "PENDÊNCIA (SAC)"
            dias_postagem = 10
            dias_prevista = -3 # Atrasado há 3 dias

        # Geração de Datas Coerentes
        data_postagem = datetime.now() - timedelta(days=dias_postagem)
        data_prevista = datetime.now() + timedelta(days=dias_prevista)

        return {
            "CODIGO": codigo_str,
            "Status": status,
            "Detalhes": detalhe,
            "Unidade_Atual": unidade,
            "Data_Postagem": data_postagem.strftime("%d/%m/%Y"),
            "Data_Prevista": data_prevista.strftime("%d/%m/%Y"),
            "Ultima_Atualizacao": datetime.now().strftime("%d/%m/%Y %H:%M")
        }

    def _api_response(self, codigo):
        """
        [PLACEHOLDER] Integração Real (Pós-Aprovação).
        Aqui entrará a lógica de Requests/Scraping com os headers já configurados no __init__.
        """
        try:
            # Jitter de segurança contra bloqueio de WAF
            time.sleep(random.uniform(0.5, 1.5))
            
            logger.warning(f"Chamada Real não implementada para {codigo}")
            
            return {
                "CODIGO": codigo,
                "Status": "Erro Config",
                "Detalhes": "Aguardando aprovação para integração API",
                "Unidade_Atual": "N/A",
                "Data_Postagem": datetime.now().strftime("%d/%m/%Y"),
                "Data_Prevista": "",
                "Ultima_Atualizacao": datetime.now().strftime("%d/%m/%Y %H:%M")
            }

        except Exception as e:
            logger.error(f"Erro na API Real: {e}")
            return {
                "CODIGO": codigo,
                "Status": "Erro Conexão",
                "Detalhes": str(e),
                "Unidade_Atual": "ERRO",
                "Data_Postagem": "",
                "Data_Prevista": "",
                "Ultima_Atualizacao": ""
            }