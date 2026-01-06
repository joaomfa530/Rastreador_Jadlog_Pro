import pandas as pd
import os
from datetime import datetime
import shutil

# ==============================================================================
# ⚙️ CONFIGURAÇÕES
# ==============================================================================
ARQUIVO_DIARIO = "base.xlsx"          # O arquivo que o funcionário baixou HOJE
ARQUIVO_MESTRE = "BASE_MESTRA_JADLOG.xlsx" # O seu "Banco de Dados" em Excel
BACKUP_DIR = "backups_base"           # Segurança: vamos salvar versões anteriores

# Mapeamento para normalizar (Igual ao do dashboard para falar a mesma língua)
MAPA_COLUNAS = {
    'CTE': 'CTE',
    'TIPO': 'TIPO',
    'STATUS': 'STATUS',
    'DATA_REF': ['DATA_EVENTO', 'Data_Entrega_Prevista', 'DATA'], 
    'UNIDADE': 'PONTO_ATUAL',
    'DETALHES': 'STATUS' 
}

# ==============================================================================
# 🛠️ FUNÇÕES DE ENGENHARIA DE DADOS
# ==============================================================================

def normalizar_colunas(df):
    df.columns = df.columns.str.strip().str.upper()
    return df

def obter_coluna(df, chaves):
    if isinstance(chaves, str): chaves = [chaves]
    for k in chaves:
        if k.upper() in df.columns: return k.upper()
    return None

def criar_backup():
    """Engenharia Sênior: Nunca sobrescreva um banco sem backup."""
    if not os.path.exists(ARQUIVO_MESTRE): return
    
    if not os.path.exists(BACKUP_DIR): os.makedirs(BACKUP_DIR)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"backup_mestra_{timestamp}.xlsx")
    shutil.copy2(ARQUIVO_MESTRE, backup_path)
    print(f"🛡️ Backup criado: {backup_path}")

def limpar_cte(valor):
    import re
    if pd.isna(valor): return ""
    return re.sub(r'\D', '', str(valor))

# ==============================================================================
# 🚀 PIPELINE DE UPSERT (ATUALIZAÇÃO INTELIGENTE)
# ==============================================================================

def atualizar_base():
    print("🚀 Iniciando Pipeline de Atualização da Base Mestra...")

    # 1. Carregar o Input Diário (A Novidade)
    print(f"📥 Lendo arquivo diário: {ARQUIVO_DIARIO}")
    try:
        df_diario = pd.read_excel(ARQUIVO_DIARIO)
        df_diario = normalizar_colunas(df_diario)
    except Exception as e:
        print(f"❌ Erro ao ler diário: {e}"); return

    # Identificar colunas chaves no Diário
    col_cte = obter_coluna(df_diario, MAPA_COLUNAS['CTE'])
    col_tipo = obter_coluna(df_diario, MAPA_COLUNAS['TIPO'])
    
    if not col_cte: print("❌ CTE não encontrado no diário."); return

    # Filtros de Qualidade (Data Quality)
    # Apenas Tipo F e CTEs válidos entram na Mestra
    if col_tipo:
        df_diario = df_diario[df_diario[col_tipo].astype(str).str.upper().str.strip() == 'F'].copy()
    
    # Padroniza CTE para usar como CHAVE PRIMÁRIA (ID)
    df_diario['ID_CTE'] = df_diario[col_cte].apply(limpar_cte)
    df_diario = df_diario[df_diario['ID_CTE'].str.len() > 5].copy()
    
    # Remove duplicatas no diário (pega o mais recente se houver duplicidade)
    df_diario = df_diario.drop_duplicates(subset='ID_CTE', keep='last')

    print(f"📊 Dados válidos no diário de hoje: {len(df_diario)}")

    # 2. Carregar a Base Mestra (O Histórico)
    if os.path.exists(ARQUIVO_MESTRE):
        print(f"📖 Lendo Base Mestra existente...")
        criar_backup() # Segurança primeiro
        df_mestra = pd.read_excel(ARQUIVO_MESTRE)
        # Garante que ID_CTE seja string para o merge funcionar
        df_mestra['ID_CTE'] = df_mestra['ID_CTE'].astype(str).apply(limpar_cte)
    else:
        print(f"✨ Base Mestra não existe. Criando nova a partir do diário.")
        # Se é a primeira vez, adiciona a coluna de anotações vazia
        df_mestra = df_diario.copy()
        df_mestra['ANOTACOES_INTERNAS'] = "" 
        df_mestra['DATA_PRIMEIRA_IMPORTACAO'] = datetime.now().strftime("%d/%m/%Y")
        
        # Salva e sai, pois não há merge a fazer
        df_mestra.to_excel(ARQUIVO_MESTRE, index=False)
        print("✅ Base Mestra criada com sucesso!")
        return

    # 3. O Grande Merge (Upsert Logic)
    # Vamos usar ID_CTE como índice para facilitar
    df_mestra.set_index('ID_CTE', inplace=True)
    df_diario.set_index('ID_CTE', inplace=True)

    # Separa as colunas que queremos preservar da Mestra (Notas, Histórico)
    # E as colunas que queremos atualizar do Diário (Status, Data, Local)
    
    print("🔄 Cruzando dados...")

    # Identifica novos e existentes
    ids_novos = df_diario.index.difference(df_mestra.index)
    ids_existentes = df_diario.index.intersection(df_mestra.index)

    print(f"   -> Novos itens: {len(ids_novos)}")
    print(f"   -> Itens atualizados: {len(ids_existentes)}")

    # A. ATUALIZAÇÃO (UPDATE)
    # Para os existentes, atualizamos TODAS as colunas vindas do diário, 
    # mas mantemos as colunas exclusivas da mestra (como ANOTACOES_INTERNAS)
    
    # O comando update do pandas faz exatamente isso: atualiza valores onde o índice bate
    df_mestra.update(df_diario)
    
    # Atualiza coluna de data de atualização
    df_mestra.loc[ids_existentes, 'ULTIMA_ATUALIZACAO'] = datetime.now().strftime("%d/%m/%Y %H:%M")

    # B. INSERÇÃO (INSERT)
    # Para os novos, adicionamos à base
    if not ids_novos.empty:
        df_novos = df_diario.loc[ids_novos].copy()
        df_novos['ANOTACOES_INTERNAS'] = "" # Cria campo vazio para novos
        df_novos['DATA_PRIMEIRA_IMPORTACAO'] = datetime.now().strftime("%d/%m/%Y")
        df_novos['ULTIMA_ATUALIZACAO'] = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        df_mestra = pd.concat([df_mestra, df_novos])

    # 4. Limpeza e Salvamento
    df_mestra.reset_index(inplace=True)
    
    # Reordenar colunas para deixar 'ANOTACOES_INTERNAS' no começo ou fim fácil de ver
    cols = list(df_mestra.columns)
    if 'ANOTACOES_INTERNAS' in cols:
        cols.remove('ANOTACOES_INTERNAS')
        cols.append('ANOTACOES_INTERNAS') # Joga pro fim
    df_mestra = df_mestra[cols]

    print(f"💾 Salvando Base Mestra atualizada ({len(df_mestra)} registros totais)...")
    df_mestra.to_excel(ARQUIVO_MESTRE, index=False)
    print("✅ Processo de Atualização Concluído com Sucesso!")
    
    print("\n💡 PRÓXIMO PASSO: Rode o 'gerar_dashboard_final.py'.")
    print("   (Lembre-se de alterar o ARQUIVO_ENTRADA dele para 'BASE_MESTRA_JADLOG.xlsx')")

if __name__ == "__main__":
    atualizar_base()