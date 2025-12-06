# app.py
import os
import sys
import pandas as pd
import re
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from tkinter import messagebox, Toplevel, Tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime

# ------------------- CONFIGURAÇÕES -------------------
EXCEL_FILE = "produtos.xlsx"
LOW_STOCK_THRESHOLD = 5

# ------------------- FUNÇÕES DE ARQUIVO -------------------
def criar_arquivo_modelo_if_missing():
    if os.path.exists(EXCEL_FILE):
        return
    df_produtos = pd.DataFrame({
        "Código do Produto": ["001", "002"],
        "Nome do Produto": ["Vanish 1L", "Água Sanitária 1L"],
        "Categoria": ["Limpeza", "Limpeza"],
        "Quantidade": [10, 20],
        "Volume": ["1L", "1L"],
        "Valor de Compra": [10.0, 1.5],
        "Valor de Venda": [20.0, 3.0],
        "Valor de Mercado": [35.0, 4.0]
    })
    df_vendas = pd.DataFrame(columns=[
        "Código de Venda", "Código do Produto", "Nome do Produto", "ID do Vendedor", "Qnt. Vendida"
    ])
    df_vendedores = pd.DataFrame({
        "ID do Vendedor": ["V001"],
        "Nome": ["Fulano"],
        "Telefone": [""],
        "Email": [""]
    })
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_produtos.to_excel(writer, sheet_name="Produtos", index=False)
        df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
        df_vendedores.to_excel(writer, sheet_name="Vendedores", index=False)

def carregar_planilhas():
    try:
        df_produtos = pd.read_excel(EXCEL_FILE, sheet_name="Produtos")
        df_vendas = pd.read_excel(EXCEL_FILE, sheet_name="Vendas")
        df_vendedores = pd.read_excel(EXCEL_FILE, sheet_name="Vendedores")
        df_produtos.columns = df_produtos.columns.astype(str)
        df_vendas.columns = df_vendas.columns.astype(str)
        df_vendedores.columns = df_vendedores.columns.astype(str)
        return df_produtos, df_vendas, df_vendedores
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir o Excel:\n{e}")
        return None, None, None

def salvar_planilhas(df_produtos, df_vendas, df_vendedores):
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_produtos.to_excel(writer, sheet_name="Produtos", index=False)
            df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
            df_vendedores.to_excel(writer, sheet_name="Vendedores", index=False)
    except Exception as e:
        messagebox.showerror("Erro ao salvar", str(e))

# ------------------- FUNÇÃO DE REFRESH -------------------
def atualizar_tree_com_estoque(tree, df, coluna_quantidade="Quantidade"):
    tree.delete(*tree.get_children())
    for _, row in df.iterrows():
        tag = ""
        if coluna_quantidade in df.columns:
            qtd = int(row[coluna_quantidade])
            tag = "baixo" if qtd < LOW_STOCK_THRESHOLD else ""
        tree.insert("", "end", values=list(row), tags=(tag,))
    tree.tag_configure("baixo", background="#ffcccc")

# ------------------- PADRONIZAÇÃO DE DADOS -------------------
def padronizar_texto(col, valor):
    if valor is None:
        return ""
    val = str(valor).strip()
    if col in ["ID do Vendedor", "Código do Produto", "Código de Venda"]:
        val = val.upper()
    if col == "Telefone":
        nums = re.sub(r"\D", "", val)[:11]
        if len(nums) == 11:
            val = f"({nums[:2]}) {nums[2:7]}-{nums[7:]}"
        elif len(nums) == 10:
            val = f"({nums[:2]}) {nums[2:6]}-{nums[6:]}"
        else:
            val = nums
    return val

# ------------------- TELA DE LOGIN -------------------
def tela_login(root):
    login_win = Toplevel(root)
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
        if user == "admin" and pwd == "1234":
            resultado["ok"] = True
            messagebox.showinfo("Bem-vindo!", f"Login realizado como {user}")
            login_win.destroy()
        else:
            messagebox.showerror("Erro", "Credenciais inválidas.")

    ttkb.Button(login_win, text="Entrar", bootstyle=SUCCESS, command=tentar_login).pack(pady=20)
    login_win.bind("<Return>", lambda e: tentar_login())
    root.wait_window(login_win)
    return resultado["ok"]

# ------------------- DASHBOARD -------------------
def criar_dashboard(frame_dash, df_produtos, df_vendas):
    for widget in frame_dash.winfo_children():
        widget.destroy()

    fig, ax = plt.subplots(1, 2, figsize=(10, 4))

    # Estoque Atual
    if "Quantidade" in df_produtos.columns:
        cores = ['#f44336' if q < LOW_STOCK_THRESHOLD else '#2196F3' for q in df_produtos["Quantidade"]]
        ax[0].bar(df_produtos["Nome do Produto"], df_produtos["Quantidade"], color=cores)
        ax[0].set_title("Estoque Atual")
        ax[0].set_ylabel("Quantidade")
        ax[0].tick_params(axis='x', rotation=45)
    else:
        ax[0].text(0.5, 0.5, "Sem informações de estoque", ha="center", va="center")

    # Vendas Totais
    if not df_vendas.empty and "Qnt. Vendida" in df_vendas.columns:
        vendas_prod = df_vendas.groupby("Nome do Produto")["Qnt. Vendida"].sum()
        ax[1].bar(vendas_prod.index, vendas_prod.values, color="#4CAF50")
        ax[1].set_title("Vendas Totais")
        ax[1].set_ylabel("Quantidade")
        ax[1].tick_params(axis='x', rotation=45)
    else:
        ax[1].text(0.5, 0.5, "Sem vendas registradas", ha="center", va="center", fontsize=12)
        ax[1].set_xticks([])
        ax[1].set_yticks([])

    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=frame_dash)
    canvas.draw()
    canvas.get_tk_widget().pack(expand=True, fill="both")

# ------------------- FUNÇÃO DE PDF DA NOTA FISCAL -------------------
def gerar_pdf_nota_fiscal(nf_dados):
    try:
        nome_pdf = f"NotaFiscal_{nf_dados.get('Código do Produto','')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
        c = canvas.Canvas(nome_pdf, pagesize=A4)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, 800, f"NOTA FISCAL - Nº {nf_dados.get('Número NF','0001')}")

        c.setFont("Helvetica", 10)
        y = 770
        for k,v in nf_dados.items():
            c.drawString(50, y, f"{k}: {v}")
            y -= 15
        c.save()
        messagebox.showinfo("PDF gerado", f"Arquivo gerado com sucesso: {nome_pdf}")
    except Exception as e:
        messagebox.showerror("Erro PDF", str(e))

# ------------------- FUNÇÃO PARA CRIAR ABAS COM POP-UP EDITAR -------------------
def criar_aba(frame, df, tipo, df_produtos=None, df_vendas=None, df_vendedores=None, atualizar_dash=None):
    left = ttkb.Frame(frame)
    left.pack(side="left", expand=True, fill="both", padx=(6,3), pady=6)
    right = ttkb.Frame(frame, width=360)
    right.pack(side="right", fill="y", padx=(3,6), pady=6)
    right.pack_propagate(False)

    cols = list(df.columns)
    tree = ttkb.Treeview(left, columns=cols, show="headings")
    tree.pack(side="left", expand=True, fill="both")
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, anchor="center", width=140)

    scroll = ttkb.Scrollbar(left, command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    scroll.pack(side="left", fill="y")

    def atualizar_popular_tree():
        atualizar_tree_com_estoque(tree, df)

    # ----------- FUNÇÕES ADICIONAR / EDITAR / EXCLUIR -----------
    def preencher_por_chave(entries_local):
        try:
            if "ID do Vendedor" in entries_local:
                vid = padronizar_texto("ID do Vendedor", entries_local["ID do Vendedor"].get())
                idxs = df_vendedores.index[df_vendedores["ID do Vendedor"].astype(str)==vid].tolist()
                if idxs:
                    idx = idxs[0]
                    for col in ["Nome","Telefone","Email"]:
                        if col in entries_local:
                            entries_local[col].delete(0, 'end')
                            entries_local[col].insert(0, str(df_vendedores.at[idx,col]))
            if "Código do Produto" in entries_local:
                pid = padronizar_texto("Código do Produto", entries_local["Código do Produto"].get())
                idxs = df_produtos.index[df_produtos["Código do Produto"].astype(str)==pid].tolist()
                if idxs:
                    idx = idxs[0]
                    for col in ["Nome do Produto","Categoria","Valor de Compra","Valor de Venda","Valor de Mercado"]:
                        if col in entries_local:
                            entries_local[col].delete(0, 'end')
                            entries_local[col].insert(0, str(df_produtos.at[idx,col]))
        except Exception as e:
            print("Erro no preenchimento automático:", e)

    def adicionar():
        popup = Toplevel(frame)
        popup.title(f"Adicionar {tipo[:-1].capitalize()}")
        popup.geometry("500x500")
        popup.grab_set()
        entries_local = {}
        for i, col in enumerate(cols):
            ttkb.Label(popup, text=col + ":").grid(row=i, column=0, sticky="w", padx=8, pady=6)
            ent = ttkb.Entry(popup)
            ent.grid(row=i, column=1, sticky="ew", padx=8, pady=6)
            ent.bind("<FocusOut>", lambda e, en=entries_local: preencher_por_chave(en))
            entries_local[col] = ent
        popup.columnconfigure(1, weight=1)

        # ---------- Botão Nota Fiscal ----------
        if tipo=="produtos":
            def abrir_nota_fiscal():
                nf_popup = Toplevel(popup)
                nf_popup.title("Nota Fiscal")
                nf_popup.geometry("400x400")
                nf_popup.grab_set()
                campos_nf = ["Número NF","Série","Data","CNPJ Emitente","Destinatário","CFOP","NCM",
                             "Quantidade","Valor Unitário","ICMS","IPI","Frete","Placa"]
                entries_nf = {}
                for i, c in enumerate(campos_nf):
                    ttkb.Label(nf_popup, text=c+":").grid(row=i, column=0, sticky="w", padx=8, pady=4)
                    ent_nf = ttkb.Entry(nf_popup)
                    ent_nf.grid(row=i, column=1, sticky="ew", padx=8, pady=4)
                    entries_nf[c] = ent_nf
                nf_popup.columnconfigure(1, weight=1)
                def salvar_nf_pdf():
                    nf_dados = {k:v.get() for k,v in entries_nf.items()}
                    nf_dados["Código do Produto"] = entries_local["Código do Produto"].get()
                    gerar_pdf_nota_fiscal(nf_dados)
                    nf_popup.destroy()
                ttkb.Button(nf_popup, text="Gerar PDF", bootstyle=SUCCESS, command=salvar_nf_pdf).grid(
                    row=len(campos_nf)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=8
                )
            ttkb.Button(popup, text="📄 Nota Fiscal", bootstyle=INFO, command=abrir_nota_fiscal).grid(
                row=len(cols)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=8
            )

        def salvar_novo():
            try:
                novo = {}
                for col in cols:
                    val = padronizar_texto(col, entries_local[col].get())
                    if col in ["ID do Vendedor","Código do Produto","Código de Venda"]:
                        if (df[col].astype(str) == val).any():
                            messagebox.showerror("Erro", f"{col} já existe!")
                            return
                    if col in ["Quantidade","Qnt. Vendida"]:
                        val = int(val) if val != "" else 0
                    elif col in ["Valor de Compra","Valor de Venda","Valor de Mercado"]:
                        val = float(val) if val != "" else 0.0
                    novo[col] = val
                df.loc[len(df)] = [novo[col] for col in cols]
                salvar_planilhas(df_produtos if df_produtos is not None else df,
                                df_vendas if df_vendas is not None else df,
                                df_vendedores if df_vendedores is not None else df)
                atualizar_popular_tree()
                messagebox.showinfo(tipo.capitalize(), f"{tipo[:-1].capitalize()} adicionado(a).")
                popup.destroy()
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        ttkb.Button(popup, text="Salvar", bootstyle=SUCCESS, command=salvar_novo).grid(
            row=len(cols)+2, column=0, columnspan=2, sticky="ew", padx=8, pady=8
        )

    # ---------- Editar / Excluir ----------
    def editar():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning(tipo.capitalize(), f"Selecione um(a) {tipo[:-1]} para editar.")
            return
        idx = tree.index(sel[0])
        popup = Toplevel(frame)
        popup.title(f"Editar {tipo[:-1].capitalize()}")
        popup.geometry("500x500")
        popup.grab_set()
        entries_local = {}
        for i, col in enumerate(cols):
            ttkb.Label(popup, text=col + ":").grid(row=i, column=0, sticky="w", padx=8, pady=6)
            ent = ttkb.Entry(popup)
            ent.grid(row=i, column=1, sticky="ew", padx=8, pady=6)
            ent.insert(0, str(df.at[idx,col]))
            ent.bind("<FocusOut>", lambda e, en=entries_local: preencher_por_chave(en))
            entries_local[col] = ent
        popup.columnconfigure(1, weight=1)

        if tipo=="produtos":
            def abrir_nota_fiscal():
                nf_popup = Toplevel(popup)
                nf_popup.title("Nota Fiscal")
                nf_popup.geometry("400x400")
                nf_popup.grab_set()
                campos_nf = ["Número NF","Série","Data","CNPJ Emitente","Destinatário","CFOP","NCM",
                             "Quantidade","Valor Unitário","ICMS","IPI","Frete","Placa"]
                entries_nf = {}
                for i, c in enumerate(campos_nf):
                    ttkb.Label(nf_popup, text=c+":").grid(row=i, column=0, sticky="w", padx=8, pady=4)
                    ent_nf = ttkb.Entry(nf_popup)
                    ent_nf.grid(row=i, column=1, sticky="ew", padx=8, pady=4)
                    entries_nf[c] = ent_nf
                nf_popup.columnconfigure(1, weight=1)
                def salvar_nf_pdf():
                    nf_dados = {k:v.get() for k,v in entries_nf.items()}
                    nf_dados["Código do Produto"] = entries_local["Código do Produto"].get()
                    gerar_pdf_nota_fiscal(nf_dados)
                    nf_popup.destroy()
                ttkb.Button(nf_popup, text="Gerar PDF", bootstyle=SUCCESS, command=salvar_nf_pdf).grid(
                    row=len(campos_nf)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=8
                )
            ttkb.Button(popup, text="📄 Nota Fiscal", bootstyle=INFO, command=abrir_nota_fiscal).grid(
                row=len(cols)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=8
            )

        def salvar_edicao():
            try:
                for col in cols:
                    val = padronizar_texto(col, entries_local[col].get())
                    if col in ["Quantidade","Qnt. Vendida"]:
                        val = int(val)
                    elif col in ["Valor de Compra","Valor de Venda","Valor de Mercado"]:
                        val = float(val)
                    if col in ["ID do Vendedor","Código do Produto","Código de Venda"]:
                        if (df[col].astype(str) == val).any() and str(df.at[idx,col]) != val:
                            messagebox.showerror("Erro", f"{col} já existe!")
                            return
                    df.at[idx,col] = val
                salvar_planilhas(df_produtos if df_produtos is not None else df,
                                df_vendas if df_vendas is not None else df,
                                df_vendedores if df_vendedores is not None else df)
                atualizar_popular_tree()
                messagebox.showinfo(tipo.capitalize(), f"{tipo[:-1].capitalize()} atualizado(a).")
                popup.destroy()
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        ttkb.Button(popup, text="Salvar", bootstyle=SUCCESS, command=salvar_edicao).grid(
            row=len(cols)+2, column=0, columnspan=2, sticky="ew", padx=8, pady=8
        )

    def excluir():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning(tipo.capitalize(), f"Selecione um(a) {tipo[:-1]} para excluir.")
            return
        idx = tree.index(sel[0])
        df.drop(df.index[idx], inplace=True)
        df.reset_index(drop=True, inplace=True)
        salvar_planilhas(df_produtos if df_produtos is not None else df,
                        df_vendas if df_vendas is not None else df,
                        df_vendedores if df_vendedores is not None else df)
        atualizar_popular_tree()

    ttkb.Button(right, text="➕ Adicionar", bootstyle=SUCCESS, command=adicionar).grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=(8,3))
    ttkb.Button(right, text="✎ Editar", bootstyle=PRIMARY, command=editar).grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
    ttkb.Button(right, text="🗑 Excluir", bootstyle=DANGER, command=excluir).grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=3)

    atualizar_popular_tree()
    return tree

# ------------------- JANELA PRINCIPAL -------------------
def abrir_janela_principal(root, df_produtos, df_vendas, df_vendedores):
    root.deiconify()
    root.title("ERP Moderno")
    root.geometry("1300x750")

    notebook = ttkb.Notebook(root, bootstyle="info")
    notebook.pack(expand=True, fill="both")

    # Aba Produtos
    frame_prod = ttkb.Frame(notebook)
    tree_prod = criar_aba(frame_prod, df_produtos, "produtos", df_produtos, df_vendas, df_vendedores)
    notebook.add(frame_prod, text="Produtos")

    # Aba Vendas
    frame_vend = ttkb.Frame(notebook)
    tree_vend = criar_aba(frame_vend, df_vendas, "vendas", df_produtos, df_vendas, df_vendedores)
    notebook.add(frame_vend, text="Vendas")

    # Aba Vendedores
    frame_vdr = ttkb.Frame(notebook)
    tree_vdr = criar_aba(frame_vdr, df_vendedores, "vendedores", df_produtos, df_vendas, df_vendedores)
    notebook.add(frame_vdr, text="Vendedores")

    # Aba Dashboard (carregado apenas ao selecionar)
    frame_dash = ttkb.Frame(notebook)
    notebook.add(frame_dash, text="Dashboard")

    def carregar_dash(event=None):
        if notebook.index("current") == 3:
            criar_dashboard(frame_dash, df_produtos, df_vendas)
    notebook.bind("<<NotebookTabChanged>>", carregar_dash)

# ------------------- FLUXO PRINCIPAL -------------------
if __name__=="__main__":
    criar_arquivo_modelo_if_missing()
    df_produtos, df_vendas, df_vendedores = carregar_planilhas()
    if df_produtos is None:
        sys.exit(1)
    root = ttkb.Window(themename="darkly")
    root.withdraw()
    if tela_login(root):
        abrir_janela_principal(root, df_produtos, df_vendas, df_vendedores)
        root.mainloop()