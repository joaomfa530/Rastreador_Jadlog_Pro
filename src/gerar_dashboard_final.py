import pandas as pd
import xlsxwriter
from datetime import date, datetime, timedelta
import re

# ==============================================================================
# 🧠 ENGINEERING CORE: Lógica de Negócio e Tratamento de Dados
# ==============================================================================

def limpar_cte(valor):
    """Remove caracteres não numéricos e garante string limpa."""
    if pd.isna(valor): return ""
    return re.sub(r'\D', '', str(valor))

def calcular_gravidade(row):
    """
    Algoritmo de SLA (Service Level Agreement).
    Define se o atraso é Leve, Médio ou Crítico baseado na Data Prevista.
    """
    # Recupera valores com segurança
    status = str(row.get('Status', '')).strip().upper()
    data_ref = row.get('Data_Ref')
    
    # 1. Se já foi entregue, SLA cumprido (ou encerrado)
    if 'ENTREGUE' in status:
        return 'Concluído'
    
    # 2. Se não tem data de referência, é impossível calcular
    if pd.isna(data_ref) or str(data_ref) == '' or str(data_ref) == 'NaT':
        return 'Sem Data Prevista'
    
    # Converte para datetime se ainda não for
    try:
        if not isinstance(data_ref, datetime):
            data_ref = pd.to_datetime(data_ref)
    except:
        return 'Data Inválida'

    hoje = pd.to_datetime(date.today())
    
    # 3. Se a data prevista é futura ou hoje, está no prazo
    if data_ref >= hoje:
        return 'No Prazo'
    
    # 4. Cálculo de Dias de Atraso
    dias_atraso = (hoje - data_ref).days
    
    if dias_atraso <= 3: return 'Leve (1-3 dias)'
    if dias_atraso <= 7: return 'Médio (4-7 dias)'
    return 'Crítico (7+ dias)'

# ==============================================================================
# 🎨 VISUAL ENGINE: Geração do Relatório "De Encher os Olhos"
# ==============================================================================

def gerar_dashboard_kpis(df_entrada, caminho_saida):
    """
    Função Principal chamada pela Interface Gráfica.
    Orquestra o tratamento de dados e a pintura do Excel.
    """
    print(f"🚀 Iniciando Engine Gráfica Meta: {caminho_saida}")
    
    # --- 1. PREPARAÇÃO DOS DADOS (DATA MUNGIN) ---
    
    # Normalização de Colunas (Evita KeyError)
    df = df_entrada.copy()
    
    # Mapeamento inteligente de colunas (Buscando as colunas reais do seu scraping)
    # Tenta achar a coluna de STATUS
    col_status = next((c for c in df.columns if c.lower() == 'status'), None)
    # Tenta achar a coluna de DATA (Postagem ou Prevista)
    col_data = next((c for c in df.columns if 'data' in c.lower() and 'prevista' in c.lower()), 
               next((c for c in df.columns if 'data' in c.lower()), None))
    # Tenta achar coluna TIPO (Seu requisito da coluna AE)
    col_tipo = next((c for c in df.columns if c.strip().upper() == 'TIPO'), None)

    # Padronização final do DataFrame para o Relatório
    df_final = pd.DataFrame()
    df_final['CODIGO'] = df.iloc[:, 0] # Assume que a primeira é o CTE/Código
    
    if col_status:
        df_final['Status'] = df[col_status].fillna('Desconhecido')
    else:
        df_final['Status'] = 'Não Verificado'

    if col_data:
        df_final['Data_Ref'] = pd.to_datetime(df[col_data], errors='coerce')
    else:
        df_final['Data_Ref'] = pd.NaT

    # Traz outras colunas úteis se existirem
    for col in ['Unidade', 'Detalhes', 'UNIDADE', 'DETALHES']:
        c_orig = next((c for c in df.columns if c.upper() == col.upper()), None)
        if c_orig:
            df_final[col.capitalize()] = df[c_orig]

    # --- REGRA DE NEGÓCIO: FILTRO "TIPO F" (Coluna AE) ---
    # Você pediu para manter essa função. Se existir a coluna TIPO, filtramos.
    if col_tipo:
        print("🔍 Aplicando filtro de integridade: TIPO == F")
        # Mantém apenas registros "F" (Físico/Final)
        mask_tipo = df[col_tipo].astype(str).str.strip().str.upper() == 'F'
        df_final = df_final[mask_tipo].copy()
    
    # Calcula SLA (Gravidade)
    df_final['SLA_Status'] = df_final.apply(calcular_gravidade, axis=1)

    # --- SEPARAÇÃO DAS 5 ABAS ---
    s_status = df_final['Status'].astype(str).str.upper().str.strip()
    
    # 1. Entregues
    df_entregues = df_final[s_status == 'ENTREGUE'].copy()
    
    # 2. Erros (Termos chave de falha)
    termos_erro = ['ERRO', 'DEVOLVIDO', 'RECUSADO', 'EXTRAVIADO', 'FALHA', 'CANCELADO', 'RETIDO']
    df_erros = df_final[s_status.isin(termos_erro) | s_status.str.contains('ERRO', na=False)].copy()
    
    # 3. Atrasadas (Não Entregue, Não Erro, mas SLA != No Prazo)
    # Exclui o que já está na aba de Erros para não duplicar
    df_atrasadas = df_final[
        (s_status != 'ENTREGUE') & 
        (~df_final.index.isin(df_erros.index)) &
        (df_final['SLA_Status'].str.contains('dias', na=False)) # Pega Leve, Médio, Crítico
    ].copy()

    # --- 2. MOTOR DE ESCRITA (XLSXWRITER) ---
    try:
        with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # === DESIGN SYSTEM (PALETA DE CORES META/FACEBOOK INSPIRED) ===
            COR_HEADER = '#1877F2' # Azul Facebook
            COR_BG_KPI = '#F0F2F5' # Cinza Claro Fundo
            
            # Formatos Globais
            fmt_card_titulo = workbook.add_format({'bold': True, 'font_color': '#65676B', 'bg_color': 'white', 'font_size': 11, 'align': 'center', 'border': 1})
            fmt_card_valor = workbook.add_format({'bold': True, 'font_color': '#1C1E21', 'bg_color': 'white', 'font_size': 22, 'align': 'center', 'border': 1, 'bottom': 6, 'bottom_color': COR_HEADER})
            
            fmt_header_tab = workbook.add_format({'bold': True, 'bg_color': '#2C3E50', 'font_color': 'white', 'border': 1, 'align': 'center'})
            fmt_cell = workbook.add_format({'border': 1, 'font_size': 10, 'align': 'left'})
            fmt_date = workbook.add_format({'border': 1, 'font_size': 10, 'num_format': 'dd/mm/yyyy', 'align': 'center'})
            
            # Formatos de Status
            fmt_success = workbook.add_format({'bg_color': '#E7F3FF', 'font_color': '#1877F2', 'bold': True}) # Azul claro
            fmt_danger = workbook.add_format({'bg_color': '#FFEBE8', 'font_color': '#DC3545', 'bold': True}) # Vermelho
            fmt_warning = workbook.add_format({'bg_color': '#FFF3CD', 'font_color': '#856404', 'bold': True}) # Amarelo

            # ==================================================================
            # ABA 1: KPIs (DASHBOARD)
            # ==================================================================
            ws_kpi = workbook.add_worksheet("KPIs")
            ws_kpi.hide_gridlines(2)
            ws_kpi.set_column('A:Z', 2) # Margem estreita
            
            # Métricas
            total = len(df_final)
            entregues = len(df_entregues)
            problemas = len(df_erros) + len(df_atrasadas)
            sucesso_pct = (entregues / total) if total > 0 else 0

            # Renderiza Cards (Simulando HTML/CSS no Excel)
            # Card 1: Total
            ws_kpi.merge_range('B2:E2', "TOTAL MONITORADO", fmt_card_titulo)
            ws_kpi.merge_range('B3:E5', total, fmt_card_valor)
            
            # Card 2: Sucesso
            ws_kpi.merge_range('G2:J2', "TAXA DE SUCESSO", fmt_card_titulo)
            ws_kpi.merge_range('G3:J5', f"{sucesso_pct:.1%}", fmt_card_valor)
            
            # Card 3: Atenção
            fmt_card_valor_red = workbook.add_format({'bold': True, 'font_color': '#DC3545', 'bg_color': 'white', 'font_size': 22, 'align': 'center', 'border': 1, 'bottom': 6, 'bottom_color': '#DC3545'})
            ws_kpi.merge_range('L2:O2', "ATENÇÃO NECESSÁRIA", fmt_card_titulo)
            ws_kpi.merge_range('L3:O5', problemas, fmt_card_valor_red)

            # Gráfico de Rosca
            ws_kpi.write('B8', "Distribuição de Status", workbook.add_format({'bold':True, 'font_size':14, 'font_color':'#444'}))
            
            # Dados ocultos para o gráfico
            ws_kpi.write_column('AA1', ['Status', 'Entregues', 'Atrasados', 'Erros', 'Outros'])
            ws_kpi.write_column('AB1', ['', entregues, len(df_atrasadas), len(df_erros), total - (entregues + len(df_atrasadas) + len(df_erros))])
            
            chart = workbook.add_chart({'type': 'doughnut'})
            chart.add_series({
                'name': 'Status',
                'categories': '=KPIs!$AA$2:$AA$5',
                'values':     '=KPIs!$AB$2:$AB$5',
                'points': [
                    {'fill': {'color': '#28A745'}}, # Verde
                    {'fill': {'color': '#FFC107'}}, # Amarelo
                    {'fill': {'color': '#DC3545'}}, # Vermelho
                    {'fill': {'color': '#6C757D'}}, # Cinza
                ],
            })
            chart.set_style(10)
            ws_kpi.insert_chart('B9', chart, {'x_scale': 1.5, 'y_scale': 1.5})

            # ==================================================================
            # ABAS DE DADOS (Visão Geral, Entregues, Atrasadas, Erros)
            # ==================================================================
            
            def criar_aba(dataframe, nome_aba, cor_tab, format_condicional=None):
                if dataframe.empty: return
                
                dataframe.to_excel(writer, sheet_name=nome_aba, index=False)
                ws = writer.sheets[nome_aba]
                ws.set_tab_color(cor_tab)
                
                # Formatação de Cabeçalho e Largura
                for idx, col in enumerate(dataframe.columns):
                    ws.write(0, idx, col.upper(), fmt_header_tab)
                    largura = max(len(str(col)) + 5, 18)
                    ws.set_column(idx, idx, largura)
                
                # Corpo (Data e Texto)
                # O Pandas já escreve os dados, vamos aplicar formatação condicional por cima
                last_row = len(dataframe) + 1
                last_col_char = chr(65 + len(dataframe.columns) - 1)
                rng = f"A2:{last_col_char}{last_row}"
                
                ws.autofilter(0, 0, last_row, len(dataframe.columns) - 1)
                ws.freeze_panes(1, 0)
                
                # Formatação Condicional Inteligente
                if format_condicional:
                    # Aplica na coluna de Status (busca dinâmica)
                    col_idx = dataframe.columns.get_loc("Status") if "Status" in dataframe.columns else -1
                    if col_idx >= 0:
                        col_char = chr(65 + col_idx)
                        rng_status = f"{col_char}2:{col_char}{last_row}"
                        ws.conditional_format(rng_status, {'type': 'no_errors', 'format': format_condicional})

            # 1. Visão Geral (Azul)
            criar_aba(df_final, "Visão Geral", "#1877F2")
            
            # 2. Entregues (Verde)
            criar_aba(df_entregues, "Entregues", "#28A745", fmt_success)
            
            # 3. Atrasadas (Amarelo)
            criar_aba(df_atrasadas, "Atrasadas", "#FFC107", fmt_warning)
            
            # 4. Erros (Vermelho)
            criar_aba(df_erros, "Erros", "#DC3545", fmt_danger)

        print("✅ Relatório Premium gerado com sucesso!")
        return True

    except Exception as e:
        print(f"❌ Erro crítico no XlsxWriter: {e}")
        # Fallback de segurança: salva sem formatação se der erro visual
        df_entrada.to_excel(caminho_saida, index=False)
        return False