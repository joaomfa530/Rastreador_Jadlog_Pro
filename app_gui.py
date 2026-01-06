import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import concurrent.futures
import pandas as pd
from datetime import datetime
import os
import sys

# Ajuste de path para PyInstaller
if getattr(sys, 'frozen', False):
    sys.path.append(os.path.join(sys._MEIPASS, 'src'))

try:
    from src.tracker_service import JadlogTracker
except ImportError:
    sys.exit(1)

class AppApresentacao:
    def __init__(self, root):
        self.root = root
        self.root.title("Rastreador Jadlog Pro (Filtro Inteligente)")
        self.root.geometry("680x580")
        self.caminho_arquivo = tk.StringVar()
        self.status = tk.StringVar(value="Aguardando...")
        self._criar_interface()

    def _criar_interface(self):
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="📦 Gerador de Relatório Logístico", font=("Segoe UI", 16, "bold"), fg="#2c3e50").pack(pady=(0, 5))
        tk.Label(main_frame, text="Processa apenas CTEs do TIPO 'F'", font=("Segoe UI", 9, "italic"), fg="gray").pack(pady=(0, 20))
        
        lbl_frame = tk.LabelFrame(main_frame, text="Base de Dados (.csv / .xlsx)", padx=10, pady=10)
        lbl_frame.pack(fill=tk.X)
        
        tk.Entry(lbl_frame, textvariable=self.caminho_arquivo, width=55).pack(side=tk.LEFT, padx=5)
        tk.Button(lbl_frame, text="Selecionar", command=self.selecionar_arquivo).pack(side=tk.LEFT)
        
        self.pbar = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.pbar.pack(fill=tk.X, pady=20)
        tk.Label(main_frame, textvariable=self.status, fg="#7f8c8d").pack()
        
        self.txt_log = scrolledtext.ScrolledText(main_frame, height=12, font=("Consolas", 8), state='disabled')
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.btn_iniciar = tk.Button(main_frame, text="GERAR RELATÓRIO FILTRADO", command=self.iniciar_thread, 
                                     bg="#2980b9", fg="white", font=("Segoe UI", 11, "bold"), height=2)
        self.btn_iniciar.pack(fill=tk.X)

    def log(self, msg):
        self.txt_log.config(state='normal')
        self.txt_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    def selecionar_arquivo(self):
        f = filedialog.askopenfilename(filetypes=[("Dados", "*.csv;*.xlsx"), ("Todos", "*.*")])
        if f: self.caminho_arquivo.set(f)

    def iniciar_thread(self):
        if not self.caminho_arquivo.get():
            messagebox.showwarning("Atenção", "Selecione a planilha base!")
            return
        self.btn_iniciar.config(state='disabled', bg="#95a5a6", text="Processando...")
        threading.Thread(target=self.processar, daemon=True).start()

    def processar(self):
        try:
            tracker = JadlogTracker()
            path = self.caminho_arquivo.get()
            
            self.log("Lendo arquivo...")
            
            # Leitura Inteligente
            if path.lower().endswith('.csv'):
                try:
                    df = pd.read_csv(path, sep=',')
                except:
                    df = pd.read_csv(path, sep=';')
            else:
                df = pd.read_excel(path)

            qtd_original = len(df)
            self.log(f"Registros carregados: {qtd_original}")

            # --- APLICAÇÃO DO FILTRO TIPO 'F' ---
            col_tipo = next((c for c in df.columns if c.strip().upper() == "TIPO"), None)
            
            if col_tipo:
                df = df[df[col_tipo].astype(str).str.strip().str.upper() == 'F']
                qtd_filtrada = len(df)
                removidos = qtd_original - qtd_filtrada
                self.log(f"Filtro 'TIPO F' aplicado.")
                self.log(f"MANTIDOS: {qtd_filtrada} | DESCARTADOS: {removidos}")
                
                if df.empty:
                    raise Exception("O filtro removeu todos os registros! Verifique se há TIPO 'F' na planilha.")
            else:
                self.log("⚠️ Coluna 'TIPO' não encontrada! Processando tudo.")

            # Identificação CTE
            col_cte = next((c for c in df.columns if c.upper().strip() in ['CTE', 'REMESSA', 'CODIGO']), None)
            if not col_cte:
                raise Exception("Coluna CTE/REMESSA não encontrada!")

            def limpar(x):
                try: return str(int(float(str(x).replace(',','.')))).strip()
                except: return str(x).strip()

            codigos = df[col_cte].apply(limpar).tolist()
            total = len(codigos)
            
            resultados = []
            max_workers = 10
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_code = {executor.submit(tracker.consultar_encomenda, c): c for c in codigos}
                for i, future in enumerate(concurrent.futures.as_completed(future_to_code)):
                    res = future.result()
                    resultados.append(res)
                    
                    if i % 10 == 0 or i == total - 1:
                        perc = (i + 1) / total * 100
                        self.pbar['value'] = perc
                        self.status.set(f"Enriquecendo linha {i+1} de {total}")
                        self.root.update_idletasks()

            self.log("Gerando arquivo final...")
            df_rastreio = pd.DataFrame(resultados)
            
            # Merge
            df['__temp_key__'] = df[col_cte].apply(limpar)
            df_final = pd.merge(df, df_rastreio, left_on='__temp_key__', right_on='CODIGO', how='left')
            if '__temp_key__' in df_final.columns: del df_final['__temp_key__']
            if 'CODIGO' in df_final.columns: del df_final['CODIGO']

            nome_saida = path.replace(".csv", "").replace(".xlsx", "") + "_RELATORIO_FINAL.xlsx"
            df_final.to_excel(nome_saida, index=False)
            
            self.log("SUCESSO!")
            self.status.set("Concluído")
            messagebox.showinfo("Sucesso", f"Relatório gerado:\n{os.path.basename(nome_saida)}")

        except Exception as e:
            self.log(f"ERRO: {e}")
            messagebox.showerror("Erro", str(e))
        finally:
            self.btn_iniciar.config(state='normal', bg="#2980b9", text="GERAR RELATÓRIO FILTRADO")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppApresentacao(root)
    root.mainloop()