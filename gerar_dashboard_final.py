import pandas as pd
import re
from datetime import date, datetime
import os

# ==============================================================================
# ⚙️ 1. CONFIGURAÇÕES E MAPEAMENTO
# ==============================================================================
ARQUIVO_ENTRADA = "BASE_MESTRA_JADLOG.xlsx"
ARQUIVO_SAIDA = "Relatorio_Logistico_Final.xlsx"

# Mapeamento: Python Name <- Nome REAL no seu Excel
MAPA_COLUNAS = {
    'CTE': 'CTE',
    'TIPO': 'TIPO',
    'STATUS': 'STATUS',
    # Procura a data nesta ordem
    'DATA_REF': ['DATA_EVENTO', 'Data_Entrega_Prevista', 'Data_Verificacao', 'DATA', 'Data_Postagem'], 
    'UNIDADE': 'PONTO_ATUAL',
    'DETALHES': 'STATUS' 
}

# ==============================================================================
# 🧠 2. LÓGICA DE NEGÓCIO (ENGINEERING CORE)
# ==============================================================================

def limpar_cte(valor):
    if pd.isna(valor): return ""
    try:
        if isinstance(valor, float): valor = int(valor)
    except: pass
    return re.sub(r'\D', '', str(valor))

def normalizar_cabecalho(df):
    """Padroniza colunas para evitar erros de 'KeyError'."""
    df.columns = df.columns.str.strip().str.upper()
    return df

def obter_coluna_segura(df, chaves_possiveis):
    if isinstance(chaves_possiveis, str): chaves_possiveis = [chaves_possiveis]
    for tentativa in chaves_possiveis:
        tentativa = tentativa.upper().strip()
        if tentativa in df.columns: return tentativa
    return None

def calcular_gravidade(row, col_status, col_data):
    """
    Calcula o SLA (Atraso) INDEPENDENTE do status de erro.
    Se não entregou e a data passou, calcula dias.
    """
    hoje = pd.to_datetime(date.today())
    
    # Recuperação segura
    status_val = row.get(col_status)
    status = str(status_val).strip().upper() if pd.notna(status_val) else ""
    data_ref = row.get(col_data)
    
    # 1. Se entregue, acabou o SLA
    if status == 'ENTREGUE': return 'Concluído'
    
    # 2. Se não tem data, não dá pra calcular
    if pd.isna(data_ref): return 'Em Trânsito (Sem Data)'
    
    # 3. Análise de Prazo (Vale para Erros e Em Trânsito)
    if data_ref >= hoje: return 'No Prazo'
    
    # 4. Cálculo de Dias de Atraso
    dias = (hoje - data_ref).days
    if dias <= 3: return 'Leve (1-3 dias)'
    if dias <= 7: return 'Médio (4-7 dias)'
    return 'Crítico (7+ dias)'

# ==============================================================================
# 🎨 3. MOTOR VISUAL (XLSXWRITER)
# ==============================================================================

def aplicar_estilo_meta(writer, df, sheet_name, start_row=0, start_col=0):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    (max_row, max_col) = df.shape

    # --- Estilos Corporativos ---
    # Fonte Arial 10 em tudo
    fmt_header = workbook.add_format({
        'bold': True, 'font_name': 'Arial', 'font_size': 10, 
        'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    fmt_cell = workbook.add_format({
        'font_name': 'Arial', 'font_size': 10, 'border': 1, 'align': 'left'
    })
    fmt_date = workbook.add_format({
        'font_name': 'Arial', 'font_size': 10, 'border': 1, 'num_format': 'dd/mm/yyyy', 'align': 'center'
    })
    
    # Cores Semânticas
    fmt_verde = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'font_name': 'Arial', 'font_size': 10, 'border': 1})
    fmt_vermelho = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'font_name': 'Arial', 'font_size': 10, 'border': 1})
    fmt_amarelo = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'font_name': 'Arial', 'font_size': 10, 'border': 1})

    # Aplica formatação de cabeçalho e corpo
    for i, col in enumerate(df.columns):
        col_idx = start_col + i
        # Header
        worksheet.write(start_row, col_idx, col, fmt_header)
        
        # Ajuste de largura (Smart Width)
        largura_texto = max(len(str(col)) + 4, 15)
        worksheet.set_column(col_idx, col_idx, largura_texto)
        
        # Corpo
        for j, valor in enumerate(df[col]):
            row_idx = start_row + 1 + j
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                worksheet.write(row_idx, col_idx, valor, fmt_date)
            else:
                worksheet.write(row_idx, col_idx, valor, fmt_cell)

    # Definição do Range para Formatação Condicional
    # Convertendo índices numéricos para letras (ex: 0->A, 1->B)
    def num_to_col(n): return chr(65 + n) if n < 26 else 'Z' # Simplificado para tabelas normais
    
    col_inicial = num_to_col(start_col)
    col_final = num_to_col(start_col + max_col - 1)
    rng = f"{col_inicial}{start_row+2}:{col_final}{start_row + max_row + 1}"

    # Regras de Cores (Status e Gravidade)
    worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'Concluído', 'format': fmt_verde})
    worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'Entregue', 'format': fmt_verde})
    
    worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'Crítico', 'format': fmt_vermelho})
    worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'Erro', 'format': fmt_vermelho})
    
    worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'Médio', 'format': fmt_amarelo})
    worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'Leve', 'format': fmt_amarelo})

# ==============================================================================
# 🚀 4. PIPELINE PRINCIPAL
# ==============================================================================

def gerar_relatorio():
    print(f"🔄 Lendo arquivo base: {ARQUIVO_ENTRADA}")
    
    # 1. Leitura
    try:
        if ARQUIVO_ENTRADA.endswith('.xlsx'): df = pd.read_excel(ARQUIVO_ENTRADA)
        else: df = pd.read_csv(ARQUIVO_ENTRADA)
        df = normalizar_cabecalho(df)
    except Exception as e:
        print(f"❌ Erro leitura: {e}"); return

    # 2. Binding de Colunas
    c_cte = obter_coluna_segura(df, MAPA_COLUNAS['CTE'])
    c_tipo = obter_coluna_segura(df, MAPA_COLUNAS['TIPO'])
    c_status = obter_coluna_segura(df, MAPA_COLUNAS['STATUS'])
    c_data = obter_coluna_segura(df, MAPA_COLUNAS['DATA_REF'])
    c_unidade = obter_coluna_segura(df, MAPA_COLUNAS['UNIDADE'])
    c_detalhes = obter_coluna_segura(df, MAPA_COLUNAS['DETALHES'])

    if not c_cte:
        print("❌ Coluna CTE não encontrada. Verifique o mapeamento."); return

    # 3. Sanitização e Filtros
    # Tipo F
    if c_tipo:
        df = df[df[c_tipo].astype(str).str.upper().str.strip() == 'F'].copy()
    
    # CTE > 5
    df['__temp_cte'] = df[c_cte].apply(limpar_cte)
    df = df[df['__temp_cte'].str.len() > 5].copy()
    
    if df.empty: print("⚠️ Filtros esvaziaram a base."); return

    # Tratamento de Data
    if c_data: df[c_data] = pd.to_datetime(df[c_data], errors='coerce')
    else: df['__DATA_FAKE'] = pd.NaT; c_data = '__DATA_FAKE'

    # 4. Cálculo de Gravidade (Inclusive para Erros, conforme referência)
    df['__gravidade_calc'] = df.apply(lambda row: calcular_gravidade(row, c_status, c_data), axis=1)

    # 5. Construção do DataFrame Final (Layout da Tabela de Dados)
    df_final = pd.DataFrame()
    df_final['CODIGO'] = df[c_cte]
    df_final['Status'] = df[c_status].fillna('Desconhecido') if c_status else 'Desconhecido'
    df_final['Detalhes'] = df[c_detalhes].fillna('') if c_detalhes else ''
    df_final['Unidade'] = df[c_unidade].fillna('') if c_unidade else ''
    # Incluindo colunas de data para referência
    df_final['Data_Ref'] = df[c_data] if c_data != '__DATA_FAKE' else pd.NaT
    df_final['Gravidade_Atraso'] = df['__gravidade_calc']

    # 6. Separação de Abas
    print("📂 Gerando abas de KPI...")
    
    # Normalização para filtros
    status_series = df_final['Status'].astype(str).str.upper().str.strip()
    grav_series = df_final['Gravidade_Atraso'].astype(str)

    # Aba Entregues
    aba_entregues = df_final[status_series == 'ENTREGUE'].copy()
    
    # Aba Erros (Status contém Erro, Devolvido, etc)
    # Lista de palavras-chave de erro
    termos_erro = ['ERRO', 'DEVOLVIDO', 'EXTRAVIADO', 'RETIDO', 'RECUSADO']
    aba_erros = df_final[status_series.isin(termos_erro) | status_series.str.contains('ERRO', na=False)].copy()
    
    # Aba Atrasadas
    # Regra: Não Entregue E Não é Erro (para não duplicar com a aba erros) E Tem atraso
    # OU, se a referência mostra atraso mesmo em erro, podemos incluir.
    # Mas para "Atrasadas" geralmente queremos "Em Trânsito Atrasado".
    aba_atrasadas = df_final[
        (status_series != 'ENTREGUE') & 
        (~status_series.isin(termos_erro)) &
        (grav_series.str.contains('dias', na=False))
    ].copy()

    # 7. Construção do Dashboard (Layout Específico 4 Blocos)
    print("📊 Calculando Dashboard...")
    
    total = len(df_final)
    entregues = len(aba_entregues)
    erros = len(aba_erros)
    # Em transito puro = Total - Entregues - Erros
    em_transito = total - entregues - erros
    sucesso = (entregues / total) if total > 0 else 0

    # Bloco 1: Indicadores
    df_kpi_geral = pd.DataFrame({
        'Indicador': ['Volume Total', 'Entregas Realizadas', 'Em Trânsito', 'Erros / Bloqueios', 'Taxa de Sucesso'],
        'Valor': [total, entregues, em_transito, erros, f"{sucesso:.1%}"]
    })

    # Bloco 2: Gargalos (Top Unidades onde NÃO está entregue)
    gargalos = df_final[status_series != 'ENTREGUE']['Unidade'].value_counts().head(5).reset_index()
    gargalos.columns = ['Top 5 Unidades (Gargalo)', 'Vol. Pendente']

    # Bloco 3: Ocorrências (Top Detalhes onde NÃO está entregue)
    ocorrencias = df_final[status_series != 'ENTREGUE']['Detalhes'].value_counts().head(5).reset_index()
    ocorrencias.columns = ['Top 5 Motivos de Retenção', 'Ocorrências']

    # Bloco 4: SLA / Gravidade (Resumo da coluna Gravidade)
    # Ordem fixa para apresentação
    sla_order = ['Leve (1-3 dias)', 'Médio (4-7 dias)', 'Crítico (7+ dias)']
    sla_counts = df_final['Gravidade_Atraso'].value_counts()
    
    sla_data = []
    for sla in sla_order:
        val = sla_counts.get(sla, 0)
        if val > 0: sla_data.append([sla, val])
    
    if not sla_data: # Se não tiver nada, bota vazio
         sla_data = [['Sem atrasos', 0]]

    df_sla = pd.DataFrame(sla_data, columns=['Gravidade do Atraso', 'Qtd Pacotes'])

    # 8. Exportação
    print(f"💾 Salvando: {ARQUIVO_SAIDA}")
    
    try:
        with pd.ExcelWriter(ARQUIVO_SAIDA, engine='xlsxwriter') as writer:
            wb = writer.book
            
            # --- ABA DASHBOARD ---
            # Escrevemos os 4 blocos em posições específicas para imitar o layout visual
            ws_dash = wb.add_worksheet('DASHBOARD')
            writer.sheets['DASHBOARD'] = ws_dash
            
            # Bloco 1 (A1)
            aplicar_estilo_meta_custom(writer, df_kpi_geral, ws_dash, 1, 1) # B2
            
            # Bloco 2 (E1) - Gargalos
            # Cabeçalho da Seção
            fmt_title_sec = wb.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'bg_color': '#D9D9D9', 'border': 1})
            ws_dash.write(1, 4, "ANÁLISE DE GARGALOS (UNIDADES)", fmt_title_sec)
            aplicar_estilo_meta_custom(writer, gargalos, ws_dash, 2, 4) # E3
            
            # Bloco 3 (I1) - Ocorrências
            ws_dash.write(1, 7, "ANÁLISE DE OCORRÊNCIAS", fmt_title_sec)
            aplicar_estilo_meta_custom(writer, ocorrencias, ws_dash, 2, 7) # H3
            
            # Bloco 4 (A10) - SLA
            row_sla = 2 + len(df_kpi_geral) + 2 # Espaço abaixo do primeiro bloco
            ws_dash.write(row_sla, 1, "SLA / GRAVIDADE DO ATRASO", fmt_title_sec)
            aplicar_estilo_meta_custom(writer, df_sla, ws_dash, row_sla + 1, 1) # B10

            # --- ABAS DE DADOS ---
            df_final.to_excel(writer, sheet_name='DADOS_COMPLETOS', index=False)
            aplicar_estilo_meta(writer, df_final, 'DADOS_COMPLETOS')
            
            if not aba_entregues.empty:
                aba_entregues.to_excel(writer, sheet_name='ENTREGUES', index=False)
                aplicar_estilo_meta(writer, aba_entregues, 'ENTREGUES')
                
            if not aba_erros.empty:
                aba_erros.to_excel(writer, sheet_name='ERROS', index=False) # A ABA ERROS VOLTOU!
                aplicar_estilo_meta(writer, aba_erros, 'ERROS')
                
            if not aba_atrasadas.empty:
                aba_atrasadas.to_excel(writer, sheet_name='ATRASADAS', index=False)
                aplicar_estilo_meta(writer, aba_atrasadas, 'ATRASADAS')

            # Título do Dashboard
            ws_dash.set_column('A:A', 2) # Margem
            ws_dash.write(0, 1, "DASHBOARD EXECUTIVO - STATUS DE ENTREGAS", wb.add_format({'bold': True, 'font_size': 12, 'font_name': 'Arial'}))

        print("✅ Relatório Final Gerado. Estrutura de KPIs e Abas restaurada.")

    except Exception as e:
        print(f"❌ Erro ao salvar: {e}")

# Função Auxiliar para escrever DataFrames soltos no Dashboard sem usar to_excel direto (mais controle)
def aplicar_estilo_meta_custom(writer, df, worksheet, start_row, start_col):
    wb = writer.book
    fmt_header = wb.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
    fmt_cell = wb.add_format({'font_name': 'Arial', 'font_size': 10, 'border': 1, 'align': 'left'})
    
    # Headers
    for i, col in enumerate(df.columns):
        worksheet.write(start_row, start_col + i, col, fmt_header)
        worksheet.set_column(start_col + i, start_col + i, max(len(str(col))+2, 18))
    
    # Dados
    for i, row in enumerate(df.values):
        for j, val in enumerate(row):
            worksheet.write(start_row + 1 + i, start_col + j, val, fmt_cell)

if __name__ == "__main__":
    gerar_relatorio()