import os
import sys
import sqlite3
import pandas as pd
import re
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from tkinter import messagebox, Toplevel, Tk, simpledialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime

# ------------------- CONFIGURAÇÕES GLOBAIS -------------------
DB_NAME = "erp_database.db"
LOW_STOCK_THRESHOLD = 5
THEME = "darkly"

# ------------------- CLASSE DE BANCO DE DADOS (SQLite) -------------------

class DatabaseManager:
    def __init__(self, db_name):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._setup_db()

    def _setup_db(self):
        # Produtos
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS Produtos (
                codigo_produto TEXT PRIMARY KEY,
                nome_produto TEXT NOT NULL,
                categoria TEXT,
                quantidade INTEGER DEFAULT 0,
                volume TEXT,
                valor_compra REAL DEFAULT 0.0,
                valor_venda REAL DEFAULT 0.0,
                valor_mercado REAL DEFAULT 0.0
            )
        """)
        # Vendedores
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS Vendedores (
                id_vendedor TEXT PRIMARY KEY,
                nome TEXT NOT NULL,
                telefone TEXT,
                email TEXT
            )
        """)
        # Vendas (Adicionado TIMESTAMP para auditoria)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS Vendas (
                codigo_venda TEXT PRIMARY KEY,
                codigo_produto TEXT,
                nome_produto TEXT,
                id_vendedor TEXT,
                qnt_vendida INTEGER DEFAULT 1,
                data_venda TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.conn.commit()
        self._ensure_initial_data()

    def _ensure_initial_data(self):
        # Inserir dados de exemplo se as tabelas estiverem vazias
        if self.cursor.execute("SELECT COUNT(*) FROM Produtos").fetchone()[0] == 0:
            produtos = [
                ("001", "Vanish 1L", "Limpeza", 10, "1L", 10.0, 20.0, 35.0),
                ("002", "Água Sanitária 1L", "Limpeza", 20, "1L", 1.5, 3.0, 4.0),
            ]
            self.cursor.executemany("INSERT INTO Produtos VALUES (?, ?, ?, ?, ?, ?, ?, ?)", produtos)
        
        if self.cursor.execute("SELECT COUNT(*) FROM Vendedores").fetchone()[0] == 0:
            vendedores = [("V001", "Fulano", "", "")]
            self.cursor.executemany("INSERT INTO Vendedores VALUES (?, ?, ?, ?)", vendedores)
            
        self.conn.commit()

    def fetch_data(self, table_name, columns=None):
        cols_str = "*"
        if columns:
            cols_str = ", ".join(columns)
        query = f"SELECT {cols_str} FROM {table_name}"
        df = pd.read_sql_query(query, self.conn)
        # Padronizar nomes de colunas para a interface (opcional, mas bom para consistência)
        # O nome do DB 'codigo_produto' se torna 'Codigo Produto' (sem acento no "o")
        df.columns = [c.replace('_', ' ').title() if c != 'codigo_produto' else 'Codigo Produto' for c in df.columns]
        return df

    def execute_query(self, query, params=()):
        try:
            self.cursor.execute(query, params)
            self.conn.commit()
            return True
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed" in str(e):
                messagebox.showerror("Erro de Integridade", "Chave Duplicada. Este Código/ID já existe.")
            else:
                messagebox.showerror("Erro no BD", f"Erro de banco de dados: {e}")
            return False
        except Exception as e:
            messagebox.showerror("Erro no BD", f"Erro inesperado no banco de dados: {e}")
            return False

    def close(self):
        self.conn.close()

# ------------------- FUNÇÕES DE UTILIDADE -------------------

def padronizar_texto(col, valor):
    if valor is None: return ""
    val = str(valor).strip()
    # Mudança: 'Codigo Produto' sem acento
    if col in ["Codigo Produto", "Id Vendedor", "Codigo Venda"]:
        return val.upper()
    if col == "Telefone":
        nums = re.sub(r"\D", "", val)[:11]
        if len(nums) == 11:
            return f"({nums[:2]}) {nums[2:7]}-{nums[7:]}"
        elif len(nums) == 10:
            return f"({nums[:2]}) {nums[2:6]}-{nums[6:]}"
        else:
            return nums
    return val

def gerar_pdf_nota_fiscal(nf_dados):
    # Lógica de PDF simplificada (igual à original)
    try:
        # Mudança: 'Codigo Produto' sem acento
        nome_pdf = f"NotaFiscal_{nf_dados.get('Codigo Produto','')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
        c = canvas.Canvas(nome_pdf, pagesize=A4)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, 800, f"NOTA FISCAL - Nº {nf_dados.get('Número Nf','0001')}")

        c.setFont("Helvetica", 10)
        y = 770
        for k,v in nf_dados.items():
            c.drawString(50, y, f"{k}: {v}")
            y -= 15
        c.save()
        messagebox.showinfo("PDF gerado", f"Arquivo gerado com sucesso: {nome_pdf}")
    except Exception as e:
        messagebox.showerror("Erro PDF", str(e))

# ------------------- CLASSE PRINCIPAL DA APLICAÇÃO -------------------

class App:
    def __init__(self, master):
        self.master = master
        self.db = DatabaseManager(DB_NAME)
        self.master.title("ERP Moderno - Powered by SQLite")
        self.master.geometry("1300x750")
        self.master.withdraw() # Esconde a janela principal até o login

        # Carrega os dados iniciais do BD para DataFrames temporários
        self.dfs = {
            "produtos": self.db.fetch_data("Produtos"),
            "vendas": self.db.fetch_data("Vendas"),
            "vendedores": self.db.fetch_data("Vendedores"),
        }

        if self._tela_login():
            self._criar_janela_principal()
        else:
            self.master.destroy()

    def _tela_login(self):
        login_win = Toplevel(self.master)
        login_win.title("Login")
        login_win.geometry("360x240")
        login_win.resizable(False, False)
        login_win.grab_set()

        ttkb.Label(login_win, text="Usuário:", font=("Helvetica", 12)).pack(pady=(20, 5))
        ent_user = ttkb.Entry(login_win, bootstyle=PRIMARY)
        ent_user.pack(padx=20, fill="x")

        ttkb.Label(login_win, text="Senha:", font=("Helvetica", 12)).pack(pady=(10, 5))
        ent_pass = ttkb.Entry(login_win, show="*", bootstyle=PRIMARY)
        ent_pass.pack(padx=20, fill="x")

        resultado = {"ok": False}

        def tentar_login():
            user = ent_user.get().strip()
            pwd = ent_pass.get().strip()
            # Credenciais fixas (Apenas para demonstração)
            if user == "admin" and pwd == "1234":
                resultado["ok"] = True
                messagebox.showinfo("Bem-vindo!", f"Login realizado como {user}")
                login_win.destroy()
            else:
                messagebox.showerror("Erro", "Credenciais inválidas.")

        ttkb.Button(login_win, text="Entrar", bootstyle=SUCCESS, command=tentar_login).pack(pady=20)
        login_win.bind("<Return>", lambda e: tentar_login())
        self.master.wait_window(login_win)
        return resultado["ok"]

    def _criar_janela_principal(self):
        self.master.deiconify()
        self.notebook = ttkb.Notebook(self.master, bootstyle="info")
        self.notebook.pack(expand=True, fill="both")
        
        # Cria as abas e armazena as Treeviews para futuras atualizações
        self.trees = {}
        
        # Aba Produtos
        frame_prod = ttkb.Frame(self.notebook)
        self.trees["produtos"] = self._criar_aba(frame_prod, "Produtos")
        self.notebook.add(frame_prod, text="Produtos")
        self.trees["produtos"] = self._criar_aba(frame_prod, "Produtos")
        # CHAME A ATUALIZAÇÃO DEPOIS DA ATRIBUIÇÃO
        self._atualizar_tree("produtos") 

        # Aba Vendas
        frame_vend = ttkb.Frame(self.notebook)
        self.trees["vendas"] = self._criar_aba(frame_vend, "Vendas")
        self.notebook.add(frame_vend, text="Vendas")
        # CHAME A ATUALIZAÇÃO DEPOIS DA ATRIBUIÇÃO
        self._atualizar_tree("vendas") 

        # Aba Vendedores
        frame_vdr = ttkb.Frame(self.notebook)
        self.trees["vendedores"] = self._criar_aba(frame_vdr, "Vendedores")
        self.notebook.add(frame_vdr, text="Vendedores")
        # CHAME A ATUALIZAÇÃO DEPOIS DA ATRIBUIÇÃO
        self._atualizar_tree("vendedores") 

        # Aba Dashboard
        self.frame_dash = ttkb.Frame(self.notebook)
        self.notebook.add(self.frame_dash, text="Dashboard")
        
        # Atualiza o Dashboard apenas ao selecionar a aba
        self.notebook.bind("<<NotebookTabChanged>>", self._carregar_dash_se_necessario)
        
    def _carregar_dash_se_necessario(self, event=None):
        if self.notebook.index("current") == 3:
            self._criar_dashboard()

    def _criar_dashboard(self):
        for widget in self.frame_dash.winfo_children():
            widget.destroy()

        df_produtos = self.dfs["produtos"]
        df_vendas = self.dfs["vendas"]
        
        fig, ax = plt.subplots(1, 2, figsize=(10, 4))
        
        # Estoque Atual
        if "Quantidade" in df_produtos.columns and not df_produtos.empty:
            # Mudança: 'Codigo Produto' sem acento
            cores = ['#f44336' if q < LOW_STOCK_THRESHOLD else '#2196F3' 
                     for q in df_produtos["Quantidade"].fillna(0)]
            ax[0].bar(df_produtos["Nome Produto"], df_produtos["Quantidade"], color=cores)
            ax[0].set_title("Estoque Atual")
            ax[0].set_ylabel("Quantidade")
            ax[0].tick_params(axis='x', rotation=45)
        else:
            ax[0].text(0.5, 0.5, "Sem informações de estoque", ha="center", va="center")

        # Vendas Totais
        if not df_vendas.empty and "Qnt Vendida" in df_vendas.columns:
            # Usar 'Nome Produto' da venda para agrupar
            vendas_prod = df_vendas.groupby("Nome Produto")["Qnt Vendida"].sum()
            ax[1].bar(vendas_prod.index, vendas_prod.values, color="#4CAF50")
            ax[1].set_title("Vendas Totais")
            ax[1].set_ylabel("Quantidade")
            ax[1].tick_params(axis='x', rotation=45)
        else:
            ax[1].text(0.5, 0.5, "Sem vendas registradas", ha="center", va="center", fontsize=12)
            ax[1].set_xticks([]); ax[1].set_yticks([])

        plt.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.frame_dash)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both")

    def _atualizar_tree(self, tipo):
        # Recarrega o DataFrame do DB e atualiza a Treeview
        table_name = tipo.capitalize()
        # Se self.dfs[tipo] já foi carregado, apenas atualiza.
        # Se for a primeira vez (na inicialização), ele já foi carregado em __init__.
        self.dfs[tipo] = self.db.fetch_data(table_name)
        df = self.dfs[tipo]
        
        # O KeyError foi corrigido porque esta função só é chamada agora após a atribuição em self.trees
        tree = self.trees[tipo] 
        
        tree.delete(*tree.get_children())
        col_qtd = "Quantidade" if tipo == "produtos" else "Qnt Vendida"
        
        for _, row in df.iterrows():
            tag = ""
            if tipo == "produtos":
                qtd = int(row.get(col_qtd, 0))
                tag = "baixo" if qtd < LOW_STOCK_THRESHOLD else ""
            tree.insert("", "end", values=list(row), tags=(tag,))
        
        tree.tag_configure("baixo", background="#ffcccc")
        
    def _criar_aba(self, frame, tipo):
        tipo_lower = tipo.lower()
        df = self.dfs[tipo_lower]
        cols = list(df.columns)
        
        left = ttkb.Frame(frame); left.pack(side="left", expand=True, fill="both", padx=(6,3), pady=6)
        right = ttkb.Frame(frame, width=360); right.pack(side="right", fill="y", padx=(3,6), pady=6)
        right.pack_propagate(False)

        tree = ttkb.Treeview(left, columns=cols, show="headings")
        tree.pack(side="left", expand=True, fill="both")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, anchor="center", width=140)

        scroll = ttkb.Scrollbar(left, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")
        
        # Botões
        ttkb.Button(right, text=f"➕ Adicionar {tipo[:-1]}", bootstyle=SUCCESS, 
                    command=lambda: self._abrir_popup_adicionar(tipo_lower, cols)).grid(
                        row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=(8,3))
        ttkb.Button(right, text=f"✎ Editar {tipo[:-1]}", bootstyle=PRIMARY, 
                    command=lambda: self._abrir_popup_editar(tipo_lower, cols, tree)).grid(
                        row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        ttkb.Button(right, text=f"🗑 Excluir {tipo[:-1]}", bootstyle=DANGER, 
                    command=lambda: self._excluir_registro(tipo_lower, tree)).grid(
                        row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        
        # A LINHA self._atualizar_tree(tipo_lower) FOI REMOVIDA DAQUI
        return tree

    # ------------------- LÓGICA CRUD -------------------
    
    def _abrir_popup_adicionar(self, tipo, cols):
        popup = Toplevel(self.master); popup.title(f"Adicionar {tipo[:-1].capitalize()}"); popup.geometry("500x500"); popup.grab_set()
        entries_local = self._criar_campos_popup(popup, cols, tipo)
        
        if tipo == "produtos":
            ttkb.Button(popup, text="📄 Nota Fiscal", bootstyle=INFO, 
                        command=lambda: self._abrir_popup_nota_fiscal(popup, entries_local)).grid(
                            row=len(cols)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        
        ttkb.Button(popup, text="Salvar", bootstyle=SUCCESS, 
                    command=lambda: self._salvar_registro(tipo, entries_local, popup)).grid(
                        row=len(cols)+(2 if tipo=="produtos" else 1), column=0, columnspan=2, sticky="ew", padx=8, pady=8)

    def _abrir_popup_editar(self, tipo, cols, tree):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning(tipo.capitalize(), f"Selecione um(a) {tipo[:-1]} para editar.")
            return
        
        idx_tree = tree.index(sel[0])
        # AQUI usamos o valor da CHAVE PRIMÁRIA da Treeview (primeira coluna) para buscar no DB
        pk_value = tree.item(sel[0], 'values')[0] 
        # Mantém a lógica de conversão para o nome da coluna no DB (snake_case)
        pk_col_name = cols[0].replace(' ', '_').lower() # ex: 'Codigo Produto' -> 'codigo_produto'
        
        popup = Toplevel(self.master); popup.title(f"Editar {tipo[:-1].capitalize()}"); popup.geometry("500x500"); popup.grab_set()
        
        # Busca o registro original
        query = f"SELECT * FROM {tipo.capitalize()} WHERE {pk_col_name} = ?"
        original_record = self.db.cursor.execute(query, (pk_value,)).fetchone()
        
        entries_local = self._criar_campos_popup(popup, cols, tipo, original_record)
        
        if tipo == "produtos":
            ttkb.Button(popup, text="📄 Nota Fiscal", bootstyle=INFO, 
                        command=lambda: self._abrir_popup_nota_fiscal(popup, entries_local)).grid(
                            row=len(cols)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        
        ttkb.Button(popup, text="Salvar Edição", bootstyle=SUCCESS, 
                    command=lambda: self._salvar_edicao(tipo, cols, entries_local, popup, pk_col_name, pk_value)).grid(
                        row=len(cols)+(2 if tipo=="produtos" else 1), column=0, columnspan=2, sticky="ew", padx=8, pady=8)

    def _criar_campos_popup(self, popup, cols, tipo, record_data=None):
        entries_local = {}
        for i, col in enumerate(cols):
            # Ignora a coluna de Data da Venda na edição/adição
            if col == "Data Venda" and tipo == "vendas": continue 
                
            ttkb.Label(popup, text=col + ":").grid(row=i, column=0, sticky="w", padx=8, pady=6)
            ent = ttkb.Entry(popup)
            ent.grid(row=i, column=1, sticky="ew", padx=8, pady=6)
            
            # Preenche o valor original na edição
            if record_data and i < len(record_data):
                ent.insert(0, str(record_data[i]))
            
            # Adiciona o preenchimento automático
            # Mudança: 'Codigo Produto' sem acento
            if col in ["Id Vendedor", "Codigo Produto"]:
                ent.bind("<FocusOut>", lambda e, en=entries_local: self._preencher_por_chave(en))
                
            entries_local[col] = ent
        
        popup.columnconfigure(1, weight=1)
        return entries_local

    def _preencher_por_chave(self, entries_local):
        # Busca no DataFrame local (cache)
        try:
            if "Id Vendedor" in entries_local:
                vid = padronizar_texto("Id Vendedor", entries_local["Id Vendedor"].get())
                df_vdr = self.dfs["vendedores"]
                registro = df_vdr[df_vdr["Id Vendedor"] == vid]
                if not registro.empty:
                    idx = registro.index[0]
                    for col in ["Nome","Telefone","Email"]:
                        if col in entries_local:
                            entries_local[col].delete(0, 'end')
                            entries_local[col].insert(0, str(df_vdr.at[idx,col]))
                            
            # Mudança: 'Codigo Produto' sem acento
            if "Codigo Produto" in entries_local:
                pid = padronizar_texto("Codigo Produto", entries_local["Codigo Produto"].get())
                df_prod = self.dfs["produtos"]
                registro = df_prod[df_prod["Codigo Produto"] == pid]
                if not registro.empty:
                    idx = registro.index[0]
                    for col in ["Nome Produto","Categoria","Valor Compra","Valor Venda","Valor Mercado"]:
                        if col in entries_local:
                            entries_local[col].delete(0, 'end')
                            entries_local[col].insert(0, str(df_prod.at[idx,col]))
        except Exception as e:
            print("Erro no preenchimento automático:", e)

    def _abrir_popup_nota_fiscal(self, parent_popup, entries_local):
        nf_popup = Toplevel(parent_popup); nf_popup.title("Nota Fiscal"); nf_popup.geometry("400x400"); nf_popup.grab_set()
        campos_nf = ["Número NF","Série","Data","CNPJ Emitente","Destinatário","CFOP","NCM","Quantidade","Valor Unitário","ICMS","IPI","Frete","Placa"]
        entries_nf = {}
        for i, c in enumerate(campos_nf):
            ttkb.Label(nf_popup, text=c+":").grid(row=i, column=0, sticky="w", padx=8, pady=4)
            ent_nf = ttkb.Entry(nf_popup)
            ent_nf.grid(row=i, column=1, sticky="ew", padx=8, pady=4)
            entries_nf[c] = ent_nf
        nf_popup.columnconfigure(1, weight=1)
        
        def salvar_nf_pdf():
            nf_dados = {k:v.get() for k,v in entries_nf.items()}
            # Mudança: 'Codigo Produto' sem acento
            nf_dados["Codigo Produto"] = entries_local.get("Codigo Produto").get() if entries_local.get("Codigo Produto") else "N/A"
            gerar_pdf_nota_fiscal(nf_dados)
            nf_popup.destroy()
            
        ttkb.Button(nf_popup, text="Gerar PDF", bootstyle=SUCCESS, command=salvar_nf_pdf).grid(
            row=len(campos_nf)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=8)

    def _validar_dados(self, tipo, data):
        # Validação básica
        required_fields = {
            # Mudança: 'Codigo Produto' sem acento
            "produtos": ["Codigo Produto", "Nome Produto"],
            "vendas": ["Codigo Venda", "Codigo Produto", "Nome Produto"],
            "vendedores": ["Id Vendedor", "Nome"]
        }
        
        for field in required_fields.get(tipo, []):
            if not data.get(field) or str(data.get(field)).strip() == "":
                messagebox.showerror("Erro de Validação", f"O campo '{field}' é obrigatório.")
                return False
        
        # Validação de tipos
        numeric_fields = ["Quantidade", "Qnt Vendida"]
        float_fields = ["Valor Compra", "Valor Venda", "Valor Mercado"]
        
        for k, v in data.items():
            try:
                if k in numeric_fields:
                    data[k] = int(v)
                elif k in float_fields:
                    # Permite "," ou "." como separador decimal
                    v_clean = str(v).replace(',', '.')
                    data[k] = float(v_clean)
            except ValueError:
                messagebox.showerror("Erro de Validação", f"O campo '{k}' deve ser um número válido.")
                return False

        # Validação de Estoque (apenas para Vendas)
        if tipo == "vendas" and data.get("Qnt Vendida", 0) <= 0:
             messagebox.showerror("Erro de Venda", "A quantidade vendida deve ser positiva.")
             return False
        if tipo == "vendas":
            # Mudança: 'Codigo Produto' sem acento
            pid = data.get("Codigo Produto")
            qnt_venda = data.get("Qnt Vendida", 0)
            df_prod = self.dfs["produtos"]
            registro = df_prod[df_prod["Codigo Produto"] == pid]
            if registro.empty:
                 messagebox.showerror("Erro de Venda", "Codigo do Produto não encontrado.")
                 return False
            
            estoque_atual = registro["Quantidade"].iloc[0]
            if qnt_venda > estoque_atual:
                messagebox.showwarning("Estoque Insuficiente", f"Estoque disponível: {estoque_atual}. Venda não registrada.")
                return False

        return True

    def _salvar_registro(self, tipo, entries_local, popup):
        # 1. Coleta e Padroniza os dados
        new_data = {}
        for col, ent in entries_local.items():
            new_data[col] = padronizar_texto(col, ent.get())
            
        # 2. Validações
        if not self._validar_dados(tipo, new_data):
            return

        # 3. Prepara a Query (conversão para snake_case do DB)
        cols_db = [c.replace(' ', '_').lower() for c in new_data.keys()]
        values = [new_data[c] for c in entries_local.keys()]
        
        table_name = tipo.capitalize()
        placeholders = ', '.join(['?'] * len(cols_db))
        cols_str = ', '.join(cols_db)
        
        query = f"INSERT INTO {table_name} ({cols_str}) VALUES ({placeholders})"
        
        # 4. Executa e Atualiza o Estoque (Se for venda)
        if self.db.execute_query(query, values):
            if tipo == "vendas":
                # Atualiza o estoque do produto vendido
                # Mudança: 'Codigo Produto' sem acento
                self._atualizar_estoque(new_data.get("Codigo Produto"), -new_data.get("Qnt Vendida"))
            
            messagebox.showinfo(tipo.capitalize(), f"{tipo[:-1].capitalize()} adicionado(a).")
            self._atualizar_tree(tipo)
            popup.destroy()

    def _salvar_edicao(self, tipo, cols, entries_local, popup, pk_col_name, pk_value):
        # 1. Coleta e Padroniza os dados
        updated_data = {}
        for col, ent in entries_local.items():
            updated_data[col] = padronizar_texto(col, ent.get())
        
        # 2. Validações
        if not self._validar_dados(tipo, updated_data):
            return
        
        # 3. Prepara a Query (conversão para snake_case do DB)
        set_clauses = []
        values = []
        for col, val in updated_data.items():
            col_db = col.replace(' ', '_').lower()
            set_clauses.append(f"{col_db} = ?")
            values.append(val)
            
        query = f"UPDATE {tipo.capitalize()} SET {', '.join(set_clauses)} WHERE {pk_col_name} = ?"
        values.append(pk_value)
        
        # 4. Executa e Atualiza o Estoque (Se for produto ou venda)
        if tipo == "produtos":
            # Assume que a 'Quantidade' é o único campo que afeta o estoque diretamente.
            self._atualizar_tree(tipo) # Apenas para forçar o recarregamento dos dados no DF de cache
            
        if self.db.execute_query(query, values):
            messagebox.showinfo(tipo.capitalize(), f"{tipo[:-1].capitalize()} atualizado(a).")
            self._atualizar_tree(tipo)
            popup.destroy()

    def _excluir_registro(self, tipo, tree):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning(tipo.capitalize(), f"Selecione um(a) {tipo[:-1]} para excluir.")
            return
        
        pk_value = tree.item(sel[0], 'values')[0]
        pk_col_name = self.dfs[tipo].columns[0].replace(' ', '_').lower()
        table_name = tipo.capitalize()
        
        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o registro com {pk_col_name.upper()} = {pk_value}?"):
            query = f"DELETE FROM {table_name} WHERE {pk_col_name} = ?"
            
            # Se for venda, reverte o estoque antes de excluir o registro
            if tipo == "vendas":
                # Mudança: 'Codigo Produto' sem acento
                qnt_vendida = self.dfs["vendas"][self.dfs["vendas"][self.dfs["vendas"].columns[0]] == pk_value]["Qnt Vendida"].iloc[0]
                cod_produto = self.dfs["vendas"][self.dfs["vendas"][self.dfs["vendas"].columns[0]] == pk_value]["Codigo Produto"].iloc[0]
                self._atualizar_estoque(cod_produto, qnt_vendida) # Reverte o estoque
                
            if self.db.execute_query(query, (pk_value,)):
                messagebox.showinfo(tipo.capitalize(), f"{tipo[:-1].capitalize()} excluído(a).")
                self._atualizar_tree(tipo)

    def _atualizar_estoque(self, codigo_produto, delta_quantidade):
        # NOTA: Esta função não faz validação, assume que a validação de venda já ocorreu.
        query = "UPDATE Produtos SET quantidade = quantidade + ? WHERE codigo_produto = ?"
        # O delta é negativo para vendas (-qnt_vendida) e positivo para entradas
        self.db.execute_query(query, (delta_quantidade, codigo_produto))
        # Recarrega o DataFrame de produtos para manter o cache atualizado
        self._atualizar_tree("produtos") 

# ------------------- FLUXO PRINCIPAL -------------------
if __name__=="__main__":
    root = ttkb.Window(themename=THEME)
    app = App(root)
    root.mainloop()