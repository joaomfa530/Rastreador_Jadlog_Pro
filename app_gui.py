import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import concurrent.futures
import pandas as pd
import numpy as np # Necessário para cálculos
from datetime import datetime
import os
import sys
from pathlib import Path

# ==============================================================================
# 🔧 CONFIGURAÇÃO DE AMBIENTE
# ==============================================================================
BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / 'src'
sys.path.append(str(SRC_DIR))

# Tenta importar o Tracker. Se não tiver, avisa (pois ele é essencial)
try:
    try:
        from src.tracker_service import JadlogTracker
    except ImportError:
        from tracker_service import JadlogTracker
except ImportError as e:
    messagebox.showerror("Erro Crítico", f"Módulo 'tracker_service' não encontrado.\nCertifique-se que o arquivo tracker_service.py existe.\nErro: {e}")
    sys.exit(1)

# ==============================================================================
# 🧠 LÓGICA DE INTELIGÊNCIA DE DADOS (NOVA ENGINE DE RELATÓRIO)
# ==============================================================================
import unicodedata # Adicione isso no topo se não tiver, ou o Python resolve dentro da função

def gerar_relatorio_avancado(df_input, output_file):
    """
    Gera relatório executivo High-Fidelity.
    ATUALIZAÇÃO CRÍTICA: Lógica de detecção de erros via Normalização de Texto e Keywords.
    """
    print("--- Iniciando Engine de Relatório (Filtro de Erros Refatorado) ---")
    
    # 0. DEEP COPY & SANITIZATION
    df = df_input.copy()
    df.columns = df.columns.str.strip().str.upper()
    df = df.loc[:, ~df.columns.duplicated()]

    # 1. DEFINIÇÃO DE ESCOPO
    cols_core = ['CTE', 'PONTO_ATUAL', 'STATUS', 'DATA_PREVISTA', 'DATA_ENTREGA', 'UNIDADE']
    
    if 'PONTO_ATUAL' not in df.columns and 'STATUS_DETALHADO' in df.columns:
         df.rename(columns={'STATUS_DETALHADO': 'PONTO_ATUAL'}, inplace=True)
    
    cols_existentes = [c for c in cols_core if c in df.columns]

    # 2. BUSINESS LOGIC ENGINE (Cálculos de Data)
    hoje = pd.to_datetime(datetime.now().date())
    col_prevista = 'DATA_PREVISTA'
    col_entrega = 'DATA_ENTREGA'

    for col in [col_prevista, col_entrega]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

    def get_atraso(row):
        if col_prevista not in row or pd.isna(row[col_prevista]): return 0
        d_prev = row[col_prevista]
        if col_entrega in row and pd.notna(row[col_entrega]):
            delta = (row[col_entrega] - d_prev).days
        else:
            delta = (hoje - d_prev).days
        return delta if delta > 0 else 0

    df['DIAS_ATRASO'] = df.apply(get_atraso, axis=1)
    df['DIAS_ATRASO'] = df['DIAS_ATRASO'].fillna(0).replace([np.inf, -np.inf], 0).astype(int)

    def get_gravidade(dias):
        if dias == 0: return 'No Prazo'
        if dias <= 3: return 'Leve'
        elif dias <= 7: return 'Médio'
        else: return 'Alto'

    df['GRAVIDADE'] = df['DIAS_ATRASO'].apply(get_gravidade)

    # 3. PREPARAÇÃO DO DATASET MESTRE
    cols_finais = cols_existentes + ['DIAS_ATRASO', 'GRAVIDADE']
    if 'MOTIVO' in df.columns: cols_finais.append('MOTIVO')
    if 'TIPO' in df.columns: cols_finais.append('TIPO')

    df_clean = df[cols_finais].copy()
    df_clean.fillna('', inplace=True)

    # ==============================================================================
    # 4. SEGMENTAÇÃO DE DADOS (LÓGICA BLINDADA)
    # ==============================================================================
    
    # Função auxiliar para normalizar texto (Remove acentos: Ç -> C, Ã -> A)
    def normalizar_texto(texto):
        if not isinstance(texto, str): return str(texto)
        # Normaliza para formatação NFKD e remove caracteres não-ASCII (acentos)
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII').upper()

    # Cria uma coluna temporária normalizada para filtrar (não vai para o Excel)
    df_clean['__STATUS_NORM__'] = df_clean['STATUS'].apply(normalizar_texto)
    
    # LISTA DE GATILHOS (Keywords Raiz)
    # Usamos partes das palavras para garantir match. 
    # Ex: "DEVOLU" pega "DEVOLUCAO", "DEVOLUÇÃO", "EM DEVOLUCAO"
    gatilhos_erro = [
        'AVARIA',          # Pega AVARIA, AVARIA PARCIAL, DANO
        'IRREGULARIDADE',  # Pega TERMO DE IRREGULARIDADE
        'DEVOLU',          # Pega EM DEVOLUCAO, DEVOLVIDO
        'FALHA',           # Pega FALHA NA ENTREGA
        'RECUSADO',        # Pega RECUSADO
        'FUTURA',          # Pega SOLICITACAO ENTREGA FUTURA
        'BUSCA',           # Pega BUSCA
        'TRAVADO',         # Pega TRAVADO
        'EXTRAVIO',        # Pega EXTRAVIO
        'ROUBO',           # Extra: Segurança
        'SINISTRO'         # Extra: Segurança
    ]
    
    # Cria Regex: AVARIA|IRREGULARIDADE|DEVOLU|FALHA...
    regex_erro = '|'.join(gatilhos_erro)
    
    # 1. Aplica o filtro na coluna normalizada
    mask_erro_status = df_clean['__STATUS_NORM__'].str.contains(regex_erro, regex=True, na=False)
    
    # 2. Reforço: Verifica coluna TIPO se existir
    mask_erro_tipo = pd.Series(False, index=df_clean.index)
    if 'TIPO' in df.columns:
        # Também normaliza o tipo para garantir
        mask_erro_tipo = df_clean['TIPO'].apply(normalizar_texto).str.contains("ERRO|FALHA", na=False)
    
    mask_erros_final = mask_erro_status | mask_erro_tipo

    # 3. Separação dos DataFrames
    df_erros = df_clean[mask_erros_final].copy()
    
    # Para as outras abas, EXCLUÍMOS o que já é erro para não duplicar e limpar a visão
    df_restante = df_clean[~mask_erros_final]
    
    mask_entregue = df_restante['STATUS'].astype(str).str.upper().str.contains('ENTREGUE', na=False)
    
    df_entregues = df_restante[mask_entregue].copy()
    
    # Atrasadas: O que sobrou (Não é Erro, Não é Entregue) e tem atraso > 0
    mask_atraso_real = (df_restante['DIAS_ATRASO'] > 0)
    df_atrasadas = df_restante[(~mask_entregue) & mask_atraso_real].copy()
    
    # Remove a coluna temporária antes de salvar
    for d in [df_erros, df_entregues, df_atrasadas]:
        if '__STATUS_NORM__' in d.columns:
            d.drop(columns=['__STATUS_NORM__'], inplace=True)
            
    # Ordenação Atrasadas
    ordem_map = {'Alto': 0, 'Médio': 1, 'Leve': 2, 'No Prazo': 3}
    df_atrasadas['__sort'] = df_atrasadas['GRAVIDADE'].map(ordem_map)
    df_atrasadas = df_atrasadas.sort_values('__sort').drop(columns=['__sort'])
    
    # ==============================================================================

    # 5. KPI ENGINE
    total_ops = len(df_clean)
    total_erros = len(df_erros)
    total_entregue = len(df_entregues)
    total_atraso = len(df_atrasadas)
    taxa_sucesso = (total_entregue / total_ops) if total_ops > 0 else 0

    kpi_resumo = pd.DataFrame([
        ['Total de Operações', total_ops],
        ['Entregues (Sucesso)', total_entregue],
        ['Em Atraso (Foco Operacional)', total_atraso],
        ['Problemas / Devoluções', total_erros],
        ['Taxa de Eficiência', taxa_sucesso]
    ], columns=['Métrica', 'Valor'])

    kpi_risk = df_atrasadas['GRAVIDADE'].value_counts().reindex(['Alto', 'Médio', 'Leve']).reset_index()
    kpi_risk.columns = ['Nível de Risco', 'Volume']
    kpi_risk.fillna(0, inplace=True)

    kpi_gargalos = pd.DataFrame(columns=['Unidade', 'Qtd'])
    if 'UNIDADE' in df_atrasadas.columns:
        kpi_gargalos = df_atrasadas['UNIDADE'].value_counts().head(5).reset_index()
        kpi_gargalos.columns = ['Unidade', 'Qtd']
        
    kpi_motivos = pd.DataFrame(columns=['Motivo', 'Qtd'])
    if 'MOTIVO' in df_atrasadas.columns:
         kpi_motivos = df_atrasadas['MOTIVO'].value_counts().head(5).reset_index()
         kpi_motivos.columns = ['Motivo Principal', 'Qtd']

    # 6. EXCEL VISUAL LAYER (XLSXWRITER - MANTIDO VISUAL PROFISSIONAL)
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- Paleta e Estilos (Mantidos idênticos) ---
            c_primary, c_header = '#1F4E78', '#2C3E50'
            c_alert_high, c_alert_med, c_alert_low = '#E74C3C', '#F39C12', '#27AE60'

            s_title = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': c_primary, 'font_name': 'Segoe UI'})
            s_subtitle = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#7F8C8D', 'font_name': 'Segoe UI', 'bottom': 2, 'bottom_color': c_primary})
            s_header = workbook.add_format({'bold': True, 'fg_color': c_header, 'font_color': 'white', 'border': 1, 'valign': 'vcenter', 'font_name': 'Segoe UI', 'font_size': 10})
            s_cell_center = workbook.add_format({'border': 1, 'font_name': 'Segoe UI', 'font_size': 10, 'align': 'center'})
            s_pct = workbook.add_format({'num_format': '0.0%', 'border': 1, 'font_name': 'Segoe UI', 'align': 'center'})
            s_high = workbook.add_format({'bg_color': c_alert_high, 'font_color': 'white', 'bold': True})
            s_med = workbook.add_format({'bg_color': c_alert_med, 'font_color': 'white', 'bold': True})
            s_low = workbook.add_format({'bg_color': c_alert_low, 'font_color': 'white', 'bold': True})

            # --- ABA 1: DASHBOARD ---
            ws_dash = workbook.add_worksheet('Dashboard Executivo')
            ws_dash.hide_gridlines(2)
            ws_dash.set_column('A:A', 2)
            
            ws_dash.write('B2', "Relatório de Performance Logística", s_title)
            ws_dash.write('B3', f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", workbook.add_format({'font_color': 'gray', 'italic': True, 'font_size': 9}))

            def draw_kpi_table(row, col, title, df_kpi):
                ws_dash.write(row, col, title, s_subtitle)
                for i, c in enumerate(df_kpi.columns): ws_dash.write(row+2, col+i, c, s_header)
                for r_idx, row_data in df_kpi.iterrows():
                    for c_idx, val in enumerate(row_data):
                        fmt = s_pct if isinstance(val, float) and val <= 1.0 else s_cell_center
                        ws_dash.write(row+3+r_idx, col+c_idx, val, fmt)
                ws_dash.set_column(col, col+len(df_kpi.columns)-1, 22)

            draw_kpi_table(4, 1, "RESUMO OPERACIONAL", kpi_resumo)
            draw_kpi_table(4, 4, "ANÁLISE DE RISCO", kpi_risk)
            draw_kpi_table(12, 1, "TOP GARGALOS", kpi_gargalos)
            draw_kpi_table(12, 4, "PRINCIPAIS MOTIVOS", kpi_motivos)

            # --- ABAS DE DADOS ---
            sheets_config = {
                'Visão Geral': df_clean,
                'Erros e Devoluções': df_erros, # Nome claro para a aba
                'Atrasadas': df_atrasadas,
                'Entregues': df_entregues
            }

            for sheet_name, dframe in sheets_config.items():
                if '__STATUS_NORM__' in dframe.columns: dframe.drop(columns=['__STATUS_NORM__'], inplace=True)
                dframe.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                ws.hide_gridlines(2)
                ws.set_tab_color(c_primary)
                ws.set_column(0, len(dframe.columns)-1, 18)
                ws.freeze_panes(1, 0)

                for col_num, value in enumerate(dframe.columns.values):
                    ws.write(0, col_num, value, s_header)

                if 'GRAVIDADE' in dframe.columns:
                    col_idx = dframe.columns.get_loc('GRAVIDADE')
                    last_row = len(dframe) + 1
                    ws.conditional_format(1, col_idx, last_row, col_idx, {'type': 'text', 'criteria': 'begins with', 'value': 'Alto', 'format': s_high})
                    ws.conditional_format(1, col_idx, last_row, col_idx, {'type': 'text', 'criteria': 'begins with', 'value': 'Médio', 'format': s_med})
                    ws.conditional_format(1, col_idx, last_row, col_idx, {'type': 'text', 'criteria': 'begins with', 'value': 'Leve', 'format': s_low})
                
                for col_name in ['DATA_PREVISTA', 'DATA_ENTREGA']:
                    if col_name in dframe.columns:
                        idx = dframe.columns.get_loc(col_name)
                        ws.set_column(idx, idx, 12, workbook.add_format({'num_format': 'dd/mm/yy', 'border': 1, 'align': 'center'}))

        return True

    except Exception as e:
        print(f"Erro na geração do relatório: {e}")
        return False
# ==============================================================================
# 🖥️ INTERFACE GRÁFICA (GUI)
# ==============================================================================
class AppApresentacao:
    def __init__(self, root):
        self.root = root
        self.root.title("Rastreador Jadlog Pro (Enterprise Edition)")
        self.root.geometry("720x600")
        self.caminho_arquivo = tk.StringVar()
        self.status = tk.StringVar(value="Sistema Pronto.")
        
        self.cor_bg = "#F0F2F5"
        self.cor_header = "#1877F2" 
        self.root.configure(bg=self.cor_bg)
        
        self._criar_interface()

    def _criar_interface(self):
        main_frame = tk.Frame(self.root, padx=25, pady=25, bg=self.cor_bg)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="📊 Dashboard Logístico Jadlog", font=("Segoe UI", 18, "bold"), fg=self.cor_header, bg=self.cor_bg).pack(pady=(0, 5))
        tk.Label(main_frame, text="Engine Visual: XlsxWriter | Engine de Dados: Pandas", font=("Segoe UI", 9), fg="#65676B", bg=self.cor_bg).pack(pady=(0, 20))
        
        lbl_frame = tk.LabelFrame(main_frame, text=" Fonte de Dados (Excel/CSV) ", padx=15, pady=15, bg="white", font=("Segoe UI", 10, "bold"))
        lbl_frame.pack(fill=tk.X, pady=5)
        
        tk.Entry(lbl_frame, textvariable=self.caminho_arquivo, width=50, font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        tk.Button(lbl_frame, text="📂 Selecionar", command=self.selecionar_arquivo, bg="#E4E6EB", relief="flat").pack(side=tk.LEFT, padx=5)
        
        self.pbar = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.pbar.pack(fill=tk.X, pady=(20, 5))
        
        lbl_status = tk.Label(main_frame, textvariable=self.status, fg="#65676B", bg=self.cor_bg, font=("Segoe UI", 9))
        lbl_status.pack(anchor="w")
        
        self.txt_log = scrolledtext.ScrolledText(main_frame, height=10, font=("Consolas", 9), state='disabled', bg="white", relief="flat")
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=15)
        
        self.btn_iniciar = tk.Button(main_frame, text="GERAR RELATÓRIO PREMIUM", command=self.iniciar_thread, 
                                     bg=self.cor_header, fg="white", font=("Segoe UI", 12, "bold"), height=2, cursor="hand2", relief="flat")
        self.btn_iniciar.pack(fill=tk.X)

    def log(self, msg):
        self.txt_log.config(state='normal')
        self.txt_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    def selecionar_arquivo(self):
        f = filedialog.askopenfilename(filetypes=[("Arquivos Excel/CSV", "*.xlsx;*.csv"), ("Todos", "*.*")])
        if f: self.caminho_arquivo.set(f)

    def iniciar_thread(self):
        if not self.caminho_arquivo.get():
            messagebox.showwarning("Atenção", "Por favor, selecione uma planilha para processar.")
            return
        self.btn_iniciar.config(state='disabled', bg="#B0B3B8", text="Processando Dados...")
        threading.Thread(target=self.processar, daemon=True).start()

    def processar(self):
        try:
            tracker = JadlogTracker()
            path = self.caminho_arquivo.get()
            self.log(f"Iniciando pipeline para: {os.path.basename(path)}")
            
            # --- FASE 1: INGESTÃO DE DADOS ---
            if path.lower().endswith('.csv'):
                try: df = pd.read_csv(path, sep=',', encoding='latin1')
                except: df = pd.read_csv(path, sep=';', encoding='latin1')
            else:
                df = pd.read_excel(path)

            qtd_total = len(df)
            self.log(f"Carga inicial: {qtd_total} registros.")

            # --- FASE 2: FILTRO TIPO 'F' ---
            col_tipo = next((c for c in df.columns if c.strip().upper() == "TIPO"), None)
            
            if col_tipo:
                df = df[df[col_tipo].astype(str).str.strip().str.upper() == 'F'].copy()
                qtd_filtrada = len(df)
                self.log(f"Filtro 'TIPO F' aplicado. Restaram: {qtd_filtrada}.")
            else:
                self.log("⚠️ Coluna 'TIPO' não detectada. Processando tudo.")

            # Identificação CTE
            col_cte = next((c for c in df.columns if c.upper().strip() in ['CTE', 'REMESSA', 'CODIGO']), None)
            if not col_cte:
                raise ValueError("Coluna CTE/REMESSA não encontrada na planilha!")

            def limpar(x):
                try: return str(int(float(str(x).replace(',','.')))).strip()
                except: return str(x).strip()

            codigos = df[col_cte].apply(limpar).tolist()
            
            # --- FASE 3: ENRIQUECIMENTO (TRACKING) ---
            self.log("Consultando API Jadlog...")
            resultados = []
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                future_to_code = {executor.submit(tracker.consultar_encomenda, c): c for c in codigos}
                total_track = len(codigos)
                
                for i, future in enumerate(concurrent.futures.as_completed(future_to_code)):
                    try:
                        res = future.result()
                        resultados.append(res)
                    except Exception as e:
                        self.log(f"Erro item {i}: {e}")

                    if i % 5 == 0 or i == total_track - 1:
                        perc = (i + 1) / total_track * 100
                        self.pbar['value'] = perc
                        self.status.set(f"Rastreando: {i+1}/{total_track}")
                        self.root.update_idletasks()

            # --- FASE 4: CONSOLIDAÇÃO ---
            self.log("Consolidando dados...")
            df_rastreio = pd.DataFrame(resultados)
            
            df['__key__'] = df[col_cte].apply(limpar)
            df_final = pd.merge(df, df_rastreio, left_on='__key__', right_on='CODIGO', how='left')
            cols_drop = ['__key__', 'CODIGO'] 
            df_final.drop(columns=[c for c in cols_drop if c in df_final.columns], inplace=True)

            # --- FASE 5: RELATÓRIO VISUAL (A MÁGICA) ---
            nome_saida = path.replace(".csv", "").replace(".xlsx", "") + "_RELATORIO_PREMIUM.xlsx"
            self.log(f"Gerando Excel Avançado: {os.path.basename(nome_saida)}")
            
            # CHAMA A NOVA FUNÇÃO INTEGRADA
            sucesso = gerar_relatorio_avancado(df_final, nome_saida)
            
            if sucesso:
                self.log("✅ SUCESSO! Relatório Gerado.")
                messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{nome_saida}")
            else:
                self.log("⚠️ Erro ao salvar Excel.")
                messagebox.showwarning("Erro", "Não foi possível salvar o arquivo final.")

            self.status.set("Concluído.")

        except Exception as e:
            self.log(f"❌ ERRO FATAL: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            
        finally:
            self.btn_iniciar.config(state='normal', bg=self.cor_header, text="GERAR RELATÓRIO PREMIUM")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppApresentacao(root)
    root.mainloop()