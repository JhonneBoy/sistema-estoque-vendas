# app.py — Versão final (tema: darkly)
import os
import sys
import pandas as pd
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from datetime import datetime

# ---------- CONFIGURAÇÕES ----------
EXCEL_FILE = "produtos.xlsx"   # um único arquivo com 3 sheets: Produtos, Vendas, Vendedores
LOW_STOCK_THRESHOLD = 5
THEME_NAME = "darkly"

# ---------- UTILITÁRIOS DE I/O ----------
def criar_arquivo_modelo_if_missing():
    """Cria um arquivo Excel com as 3 sheets padrão, caso não exista."""
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
        "Código de Venda", "Código do Produto", "Nome do Produto", "ID do Vendedor", "Qnt. Vendida", "Data"
    ])
    df_vendedores = pd.DataFrame({
        "ID do Vendedor": ["V001"],
        "Nome": ["Fulano"],
        "Telefone": [""],
        "Email": [""]
    })
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_produtos.to_excel(writer, sheet_name="Produtos", index=False)
            df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
            df_vendedores.to_excel(writer, sheet_name="Vendedores", index=False)
    except Exception as e:
        messagebox.showerror("Erro ao criar arquivo", str(e))

def carregar_planilhas():
    """Carrega as 3 sheets; se faltar alguma coluna cria as colunas necessárias com valores padrão."""
    try:
        xls = pd.read_excel(EXCEL_FILE, sheet_name=None)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir o Excel:\n{e}")
        return None, None, None

    # Carrega ou cria padrões
    df_produtos = xls.get("Produtos", pd.DataFrame())
    df_vendas = xls.get("Vendas", pd.DataFrame())
    df_vendedores = xls.get("Vendedores", pd.DataFrame())

    # Normaliza colunas esperadas para cada DataFrame (evita KeyError)
    cols_prod = ["Código do Produto", "Nome do Produto", "Categoria", "Quantidade",
                 "Volume", "Valor de Compra", "Valor de Venda", "Valor de Mercado"]
    for c in cols_prod:
        if c not in df_produtos.columns:
            # cria coluna com valores vazios ou 0
            df_produtos[c] = 0 if c == "Quantidade" else (0.0 if "Valor" in c else "")

    # Força ordem das colunas
    df_produtos = df_produtos[cols_prod]

    cols_vendas = ["Código de Venda", "Código do Produto", "Nome do Produto", "ID do Vendedor", "Qnt. Vendida", "Data"]
    for c in cols_vendas:
        if c not in df_vendas.columns:
            df_vendas[c] = "" if c != "Qnt. Vendida" else 0
    df_vendas = df_vendas[cols_vendas]

    cols_vendedores = ["ID do Vendedor", "Nome", "Telefone", "Email"]
    for c in cols_vendedores:
        if c not in df_vendedores.columns:
            df_vendedores[c] = ""
    df_vendedores = df_vendedores[cols_vendedores]

    # Tipos coerentes
    try:
        df_produtos["Quantidade"] = pd.to_numeric(df_produtos["Quantidade"], errors="coerce").fillna(0).astype(int)
    except Exception:
        df_produtos["Quantidade"] = df_produtos["Quantidade"].apply(lambda v: int(v) if str(v).isdigit() else 0)

    df_vendas["Qnt. Vendida"] = pd.to_numeric(df_vendas["Qnt. Vendida"], errors="coerce").fillna(0).astype(int)

    return df_produtos.reset_index(drop=True), df_vendas.reset_index(drop=True), df_vendedores.reset_index(drop=True)

def salvar_planilhas(df_produtos, df_vendas, df_vendedores):
    """Salva as três tabelas no arquivo Excel (atomicamente com ExcelWriter)."""
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_produtos.to_excel(writer, sheet_name="Produtos", index=False)
            df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
            df_vendedores.to_excel(writer, sheet_name="Vendedores", index=False)
    except Exception as e:
        messagebox.showerror("Erro ao salvar", str(e))

# ---------- HELPERS UI ----------
def atualizar_tree(tree, df):
    """Preenche um Treeview com os dados do DataFrame (limpa antes)."""
    tree.delete(*tree.get_children())
    for _, row in df.iterrows():
        vals = [row[c] for c in df.columns]
        tree.insert("", "end", values=vals)

def gerar_codigo_venda(df_vendas):
    """Gera código de venda incremental como V001, V002..."""
    if df_vendas is None or df_vendas.empty:
        return "V001"
    try:
        numeros = []
        for c in df_vendas["Código de Venda"].astype(str).tolist():
            if c and c.startswith("V"):
                try:
                    numeros.append(int(c[1:]))
                except:
                    pass
        prox = (max(numeros) + 1) if numeros else (len(df_vendas) + 1)
    except Exception:
        prox = len(df_vendas) + 1
    return f"V{prox:03d}"

# ---------- APP PRINCIPAL ----------
def main():
    criar_arquivo_modelo_if_missing()
    df_produtos, df_vendas, df_vendedores = carregar_planilhas()
    if df_produtos is None:
        sys.exit(1)

    # ----------------- Janela (ttkbootstrap) -----------------
    root = ttkb.Window(themename=THEME_NAME)
    root.title("ERP - Gestão (Produtos / Vendas / Vendedores)")
    root.geometry("1300x760")
    root.minsize(980, 560)

    # centralizar? opcional
    try:
        root.eval('tk::PlaceWindow %s center' % root.winfo_pathname(root.winfo_id()))
    except Exception:
        pass

    # ---------- Tela de login modal ----------
    def tela_login():
        login = ttkb.Toplevel(root)
        login.title("Login")
        login.resizable(False, False)
        login.grab_set()
        login.geometry("360x230")

        frame = ttkb.Frame(login, padding=12)
        frame.pack(expand=True, fill="both")

        ttkb.Label(frame, text="Usuário").grid(row=0, column=0, sticky="w", pady=(6, 2))
        ent_user = ttkb.Entry(frame)
        ent_user.grid(row=1, column=0, sticky="ew", padx=2)

        ttkb.Label(frame, text="Senha").grid(row=2, column=0, sticky="w", pady=(8, 2))
        ent_pass = ttkb.Entry(frame, show="*")
        ent_pass.grid(row=3, column=0, sticky="ew", padx=2)

        # ajustar responsividade
        frame.columnconfigure(0, weight=1)

        resultado = {"ok": False}

        def tentar_login(_=None):
            user = ent_user.get().strip()
            pwd = ent_pass.get().strip()
            # credenciais padrão (pode ser estendido)
            if user == "admin" and pwd == "1234":
                resultado["ok"] = True
                login.destroy()
            else:
                messagebox.showerror("Login", "Credenciais inválidas.")

        btn = ttkb.Button(frame, text="Entrar", bootstyle=SUCCESS, command=tentar_login)
        btn.grid(row=4, column=0, pady=12, sticky="ew")
        login.bind("<Return>", tentar_login)

        root.wait_window(login)
        return resultado["ok"]

    # ---------- Janela principal (conteúdo) ----------
    notebook = ttkb.Notebook(root)
    notebook.pack(expand=True, fill="both", padx=8, pady=8)

    # Frames das abas (criados vazios; preenchidos a seguir)
    frame_dash = ttkb.Frame(notebook)
    frame_prod = ttkb.Frame(notebook)
    frame_vend = ttkb.Frame(notebook)
    frame_vdr = ttkb.Frame(notebook)

    notebook.add(frame_prod, text="Produtos")
    notebook.add(frame_vend, text="Vendas")
    notebook.add(frame_vdr, text="Vendedores")
    notebook.add(frame_dash, text="Dashboard")  # Dashboard por último

    # ---------- FUNÇÕES PARA CRIAR ABAS ----------
    # Produtos
    def criar_aba_produtos(parent):
        parent.pack_propagate(False)
        # Left: Tree, Right: painel de edição/ações
        left = ttkb.Frame(parent)
        left.pack(side="left", expand=True, fill="both", padx=(6,3), pady=6)
        right = ttkb.Frame(parent, width=360)
        right.pack(side="right", fill="y", padx=(3,6), pady=6)
        right.pack_propagate(False)

        cols = list(df_produtos.columns)
        tree = ttkb.Treeview(left, columns=cols, show="headings", bootstyle="secondary")
        tree.pack(side="left", expand=True, fill="both")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, anchor="center", width=120)

        scroll = ttkb.Scrollbar(left, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        # Entradas
        entries = {}
        for i, col in enumerate(cols):
            ttkb.Label(right, text=col + ":").grid(row=i, column=0, sticky="w", padx=8, pady=6)
            ent = ttkb.Entry(right)
            ent.grid(row=i, column=1, sticky="ew", padx=8, pady=6)
            entries[col] = ent
        right.columnconfigure(1, weight=1)

        # Preencher campos ao selecionar
        def on_select(event=None):
            sel = tree.selection()
            if not sel:
                return
            # pega código do produto (primeira coluna) para localizar index real
            values = tree.item(sel[0], "values")
            cod = str(values[0])
            idxs = df_produtos.index[df_produtos["Código do Produto"].astype(str) == cod].tolist()
            if not idxs:
                return
            idx = idxs[0]
            for col in cols:
                entries[col].delete(0, tk.END)
                entries[col].insert(0, str(df_produtos.at[idx, col]))

        tree.bind("<<TreeviewSelect>>", on_select)

        # Ações: adicionar / editar / excluir
        def adicionar():
            nonlocal df_produtos, df_vendas, df_vendedores
            try:
                novo = {}
                for col in cols:
                    val = entries[col].get().strip()
                    if col == "Quantidade":
                        val = int(val) if val != "" else 0
                    elif col in ("Valor de Compra", "Valor de Venda", "Valor de Mercado"):
                        val = float(val) if val != "" else 0.0
                    novo[col] = val
                # validação básica
                if not novo["Código do Produto"] or not novo["Nome do Produto"]:
                    messagebox.showwarning("Produto", "Código e Nome são obrigatórios.")
                    return
                if (df_produtos["Código do Produto"].astype(str) == str(novo["Código do Produto"])).any():
                    messagebox.showerror("Produto", "Já existe produto com esse código.")
                    return
                df_produtos = pd.concat([df_produtos, pd.DataFrame([novo])], ignore_index=True)
                salvar_planilhas(df_produtos, df_vendas, df_vendedores)
                atualizar_tree(tree, df_produtos)
                messagebox.showinfo("Produto", "Produto adicionado.")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        def editar():
            nonlocal df_produtos, df_vendas, df_vendedores
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Produto", "Selecione um produto para editar.")
                return
            values = tree.item(sel[0], "values")
            cod = str(values[0])
            idxs = df_produtos.index[df_produtos["Código do Produto"].astype(str) == cod].tolist()
            if not idxs:
                messagebox.showerror("Produto", "Índice do produto não encontrado.")
                return
            idx = idxs[0]
            try:
                for col in cols:
                    val = entries[col].get().strip()
                    if col == "Quantidade":
                        val = int(val) if val != "" else 0
                    elif col in ("Valor de Compra", "Valor de Venda", "Valor de Mercado"):
                        val = float(val) if val != "" else 0.0
                    df_produtos.at[idx, col] = val
                salvar_planilhas(df_produtos, df_vendas, df_vendedores)
                atualizar_tree(tree, df_produtos)
                messagebox.showinfo("Produto", "Produto atualizado.")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        def excluir():
            nonlocal df_produtos, df_vendas, df_vendedores
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Produto", "Selecione um produto para excluir.")
                return
            values = tree.item(sel[0], "values")
            cod = str(values[0])
            if (df_vendas["Código do Produto"].astype(str) == cod).any():
                if not messagebox.askyesno("Confirma", "Existem vendas vinculadas a este produto. Excluir mesmo assim?"):
                    return
            if not messagebox.askyesno("Confirma", f"Excluir produto {cod}?"):
                return
            idxs = df_produtos.index[df_produtos["Código do Produto"].astype(str) == cod].tolist()
            if not idxs:
                messagebox.showerror("Produto", "Produto não encontrado.")
                return
            df_produtos.drop(index=idxs[0], inplace=True)
            df_produtos.reset_index(drop=True, inplace=True)
            salvar_planilhas(df_produtos, df_vendas, df_vendedores)
            atualizar_tree(tree, df_produtos)
            messagebox.showinfo("Produto", "Produto excluído.")

        # Botões
        btn_add = ttkb.Button(right, text="➕ Adicionar", bootstyle=SUCCESS, command=adicionar)
        btn_edit = ttkb.Button(right, text="✎ Editar", bootstyle=PRIMARY, command=editar)
        btn_del = ttkb.Button(right, text="🗑 Excluir", bootstyle=DANGER, command=excluir)
        btn_add.grid(row=len(cols)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=(8,3))
        btn_edit.grid(row=len(cols)+2, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        btn_del.grid(row=len(cols)+3, column=0, columnspan=2, sticky="ew", padx=8, pady=3)

        atualizar_tree(tree, df_produtos)
        return tree

    # Vendas
    def criar_aba_vendas(parent):
        parent.pack_propagate(False)
        left = ttkb.Frame(parent)
        left.pack(side="left", expand=True, fill="both", padx=(6,3), pady=6)
        right = ttkb.Frame(parent, width=360)
        right.pack(side="right", fill="y", padx=(3,6), pady=6)
        right.pack_propagate(False)

        cols = list(df_vendas.columns)
        tree = ttkb.Treeview(left, columns=cols, show="headings")
        tree.pack(side="left", expand=True, fill="both")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, anchor="center", width=140)

        scroll = ttkb.Scrollbar(left, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        entries = {}
        for i, col in enumerate(cols):
            ttkb.Label(right, text=col + ":").grid(row=i, column=0, sticky="w", padx=8, pady=6)
            ent = ttkb.Entry(right)
            ent.grid(row=i, column=1, sticky="ew", padx=8, pady=6)
            entries[col] = ent
        right.columnconfigure(1, weight=1)

        def on_select(event=None):
            sel = tree.selection()
            if not sel:
                return
            values = tree.item(sel[0], "values")
            cod_venda = str(values[0])
            idxs = df_vendas.index[df_vendas["Código de Venda"].astype(str) == cod_venda].tolist()
            if not idxs:
                return
            idx = idxs[0]
            for col in cols:
                entries[col].delete(0, tk.END)
                entries[col].insert(0, str(df_vendas.at[idx, col]))

        tree.bind("<<TreeviewSelect>>", on_select)

        def adicionar_venda():
            nonlocal df_produtos, df_vendas, df_vendedores
            try:
                novo = {}
                for col in cols:
                    novo[col] = entries[col].get().strip()
                # campos obrigatórios
                if not novo["Código do Produto"] or not novo["ID do Vendedor"]:
                    messagebox.showwarning("Venda", "Código do Produto e ID do Vendedor obrigatórios.")
                    return
                novo["Qnt. Vendida"] = int(novo.get("Qnt. Vendida") or 0)
                # verificar produto
                idxp = df_produtos.index[df_produtos["Código do Produto"].astype(str) == str(novo["Código do Produto"])].tolist()
                if not idxp:
                    messagebox.showerror("Venda", "Produto não encontrado.")
                    return
                idxp = idxp[0]
                if novo["Qnt. Vendida"] <= 0:
                    messagebox.showwarning("Venda", "Quantidade deve ser maior que zero.")
                    return
                if novo["Qnt. Vendida"] > int(df_produtos.at[idxp, "Quantidade"]):
                    messagebox.showwarning("Venda", f"Estoque insuficiente. Disponível: {df_produtos.at[idxp, 'Quantidade']}")
                    return
                # decrementar estoque
                df_produtos.at[idxp, "Quantidade"] = int(df_produtos.at[idxp, "Quantidade"]) - novo["Qnt. Vendida"]
                novo["Nome do Produto"] = df_produtos.at[idxp, "Nome do Produto"]
                novo["Código de Venda"] = gerar_codigo_venda(df_vendas)
                novo["Data"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # adicionar
                df_vendas = pd.concat([df_vendas, pd.DataFrame([novo])], ignore_index=True)
                salvar_planilhas(df_produtos, df_vendas, df_vendedores)
                atualizar_tree(tree, df_vendas)
                # se existir tree de produtos, atualiza
                try:
                    atualizar_tree(tree_produtos_ref[0], df_produtos)
                except Exception:
                    pass
                messagebox.showinfo("Venda", f"Venda {novo['Código de Venda']} registrada.")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        def editar_venda():
            nonlocal df_produtos, df_vendas, df_vendedores
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Venda", "Selecione uma venda para editar.")
                return
            cod_venda = str(tree.item(sel[0], "values")[0])
            idxs = df_vendas.index[df_vendas["Código de Venda"].astype(str) == cod_venda].tolist()
            if not idxs:
                messagebox.showerror("Venda", "Venda não encontrada.")
                return
            idx = idxs[0]
            try:
                # salve estado anterior para ajustar estoque se necessário
                old_prod = str(df_vendas.at[idx, "Código do Produto"])
                old_q = int(df_vendas.at[idx, "Qnt. Vendida"])

                new_prod = entries["Código do Produto"].get().strip()
                new_q = int(entries["Qnt. Vendida"].get().strip() or 0)
                new_vendedor = entries["ID do Vendedor"].get().strip()

                if new_q <= 0:
                    messagebox.showwarning("Venda", "Quantidade deve ser maior que zero.")
                    return

                # localizar índices de produtos para ajuste
                idx_oldp = df_produtos.index[df_produtos["Código do Produto"].astype(str) == old_prod].tolist()
                idx_newp = df_produtos.index[df_produtos["Código do Produto"].astype(str) == new_prod].tolist()
                if not idx_newp:
                    messagebox.showerror("Venda", "Produto novo não encontrado.")
                    return
                idx_newp = idx_newp[0]

                # devolver estoque do antigo (se existir)
                if idx_oldp:
                    df_produtos.at[idx_oldp[0], "Quantidade"] = int(df_produtos.at[idx_oldp[0], "Quantidade"]) + old_q

                # checar disponibilidade do novo produto
                if new_q > int(df_produtos.at[idx_newp, "Quantidade"]):
                    # re-decremento reversão já feita ao devolver; apenas warn e sair
                    messagebox.showwarning("Venda", f"Estoque insuficiente no produto {new_prod}. Disponível: {df_produtos.at[idx_newp, 'Quantidade']}")
                    # Como já devolvemos o estoque antigo, salvar e atualizar e sair
                    salvar_planilhas(df_produtos, df_vendas, df_vendedores)
                    atualizar_tree(tree, df_vendas)
                    try:
                        atualizar_tree(tree_produtos_ref[0], df_produtos)
                    except Exception:
                        pass
                    return

                # aplicar nova venda
                df_produtos.at[idx_newp, "Quantidade"] = int(df_produtos.at[idx_newp, "Quantidade"]) - new_q

                # atualizar venda
                df_vendas.at[idx, "Código do Produto"] = new_prod
                df_vendas.at[idx, "Nome do Produto"] = df_produtos.at[idx_newp, "Nome do Produto"]
                df_vendas.at[idx, "ID do Vendedor"] = new_vendedor
                df_vendas.at[idx, "Qnt. Vendida"] = new_q
                df_vendas.at[idx, "Data"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                salvar_planilhas(df_produtos, df_vendas, df_vendedores)
                atualizar_tree(tree, df_vendas)
                try:
                    atualizar_tree(tree_produtos_ref[0], df_produtos)
                except Exception:
                    pass
                messagebox.showinfo("Venda", "Venda atualizada.")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        def excluir_venda():
            nonlocal df_produtos, df_vendas, df_vendedores
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Venda", "Selecione uma venda para excluir.")
                return
            cod_venda = str(tree.item(sel[0], "values")[0])
            idxs = df_vendas.index[df_vendas["Código de Venda"].astype(str) == cod_venda].tolist()
            if not idxs:
                messagebox.showerror("Venda", "Venda não encontrada.")
                return
            idx = idxs[0]
            # devolver estoque
            cod_prod = str(df_vendas.at[idx, "Código do Produto"])
            q = int(df_vendas.at[idx, "Qnt. Vendida"])
            idxp = df_produtos.index[df_produtos["Código do Produto"].astype(str) == cod_prod].tolist()
            if idxp:
                df_produtos.at[idxp[0], "Quantidade"] = int(df_produtos.at[idxp[0], "Quantidade"]) + q
            # remover venda
            df_vendas.drop(index=idx, inplace=True)
            df_vendas.reset_index(drop=True, inplace=True)
            salvar_planilhas(df_produtos, df_vendas, df_vendedores)
            atualizar_tree(tree, df_vendas)
            try:
                atualizar_tree(tree_produtos_ref[0], df_produtos)
            except Exception:
                pass
            messagebox.showinfo("Venda", "Venda excluída e estoque ajustado.")

        # Botões
        btn_add = ttkb.Button(right, text="➕ Adicionar", bootstyle=SUCCESS, command=adicionar_venda)
        btn_edit = ttkb.Button(right, text="✎ Editar", bootstyle=PRIMARY, command=editar_venda)
        btn_del = ttkb.Button(right, text="🗑 Excluir", bootstyle=DANGER, command=excluir_venda)
        btn_add.grid(row=len(cols)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=(8,3))
        btn_edit.grid(row=len(cols)+2, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        btn_del.grid(row=len(cols)+3, column=0, columnspan=2, sticky="ew", padx=8, pady=3)

        atualizar_tree(tree, df_vendas)
        return tree

    # Vendedores
    def criar_aba_vendedores(parent):
        parent.pack_propagate(False)
        left = ttkb.Frame(parent)
        left.pack(side="left", expand=True, fill="both", padx=(6,3), pady=6)
        right = ttkb.Frame(parent, width=360)
        right.pack(side="right", fill="y", padx=(3,6), pady=6)
        right.pack_propagate(False)

        cols = list(df_vendedores.columns)
        tree = ttkb.Treeview(left, columns=cols, show="headings")
        tree.pack(side="left", expand=True, fill="both")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, anchor="center", width=140)

        scroll = ttkb.Scrollbar(left, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        entries = {}
        for i, col in enumerate(cols):
            ttkb.Label(right, text=col + ":").grid(row=i, column=0, sticky="w", padx=8, pady=6)
            ent = ttkb.Entry(right)
            ent.grid(row=i, column=1, sticky="ew", padx=8, pady=6)
            entries[col] = ent
        right.columnconfigure(1, weight=1)

        def on_select(event=None):
            sel = tree.selection()
            if not sel:
                return
            values = tree.item(sel[0], "values")
            cod = str(values[0])
            idxs = df_vendedores.index[df_vendedores["ID do Vendedor"].astype(str) == cod].tolist()
            if not idxs:
                return
            idx = idxs[0]
            for col in cols:
                entries[col].delete(0, tk.END)
                entries[col].insert(0, str(df_vendedores.at[idx, col]))

        tree.bind("<<TreeviewSelect>>", on_select)

        def adicionar_vdr():
            nonlocal df_vendedores, df_produtos, df_vendas
            novo = {}
            for col in cols:
                novo[col] = entries[col].get().strip()
            if not novo["ID do Vendedor"] or not novo["Nome"]:
                messagebox.showwarning("Vendedor", "ID e Nome são obrigatórios.")
                return
            if (df_vendedores["ID do Vendedor"].astype(str) == novo["ID do Vendedor"]).any():
                messagebox.showerror("Vendedor", "Já existe vendedor com esse ID.")
                return
            df_vendedores = pd.concat([df_vendedores, pd.DataFrame([novo])], ignore_index=True)
            salvar_planilhas(df_produtos, df_vendas, df_vendedores)
            atualizar_tree(tree, df_vendedores)
            messagebox.showinfo("Vendedor", "Vendedor adicionado.")

        def editar_vdr():
            nonlocal df_vendedores, df_produtos, df_vendas
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Vendedor", "Selecione um vendedor para editar.")
                return
            cod = str(tree.item(sel[0], "values")[0])
            idxs = df_vendedores.index[df_vendedores["ID do Vendedor"].astype(str) == cod].tolist()
            if not idxs:
                messagebox.showerror("Vendedor", "Vendedor não encontrado.")
                return
            idx = idxs[0]
            for col in cols:
                df_vendedores.at[idx, col] = entries[col].get().strip()
            salvar_planilhas(df_produtos, df_vendas, df_vendedores)
            atualizar_tree(tree, df_vendedores)
            messagebox.showinfo("Vendedor", "Vendedor atualizado.")

        def excluir_vdr():
            nonlocal df_vendedores, df_produtos, df_vendas
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Vendedor", "Selecione um vendedor para excluir.")
                return
            cod = str(tree.item(sel[0], "values")[0])
            if (df_vendas["ID do Vendedor"].astype(str) == cod).any():
                if not messagebox.askyesno("Confirma", "Existem vendas com esse vendedor. Excluir mesmo assim?"):
                    return
            idxs = df_vendedores.index[df_vendedores["ID do Vendedor"].astype(str) == cod].tolist()
            if not idxs:
                messagebox.showerror("Vendedor", "Vendedor não encontrado.")
                return
            df_vendedores.drop(index=idxs[0], inplace=True)
            df_vendedores.reset_index(drop=True, inplace=True)
            salvar_planilhas(df_produtos, df_vendas, df_vendedores)
            atualizar_tree(tree, df_vendedores)
            messagebox.showinfo("Vendedor", "Vendedor excluído.")

        btn_add = ttkb.Button(right, text="➕ Adicionar", bootstyle=SUCCESS, command=adicionar_vdr)
        btn_edit = ttkb.Button(right, text="✎ Editar", bootstyle=PRIMARY, command=editar_vdr)
        btn_del = ttkb.Button(right, text="🗑 Excluir", bootstyle=DANGER, command=excluir_vdr)
        btn_add.grid(row=len(cols)+1, column=0, columnspan=2, sticky="ew", padx=8, pady=(8,3))
        btn_edit.grid(row=len(cols)+2, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        btn_del.grid(row=len(cols)+3, column=0, columnspan=2, sticky="ew", padx=8, pady=3)

        atualizar_tree(tree, df_vendedores)
        return tree

    # ---------- Dashboard (carregado sob demanda) ----------
    dashboard_canvas = {"loaded": False}

    def criar_dashboard(frame):
        # apaga widgets prévios
        for w in frame.winfo_children():
            w.destroy()
        fig, axes = plt.subplots(1, 2, figsize=(10, 4))

        # gráfico estoque
        try:
            nomes = df_produtos["Nome do Produto"].astype(str).tolist()
            qtds = df_produtos["Quantidade"].astype(int).tolist()
            cores = ['#f44336' if q < LOW_STOCK_THRESHOLD else '#2196F3' for q in qtds]
            axes[0].bar(nomes, qtds, color=cores)
            axes[0].set_title("Estoque Atual")
            axes[0].tick_params(axis='x', rotation=30)
        except Exception as e:
            axes[0].text(0.5, 0.5, "Erro gerando gráfico de estoque", ha="center")
        # gráfico vendas agregadas
        try:
            if not df_vendas.empty:
                vendas_agr = df_vendas.groupby("Nome do Produto")["Qnt. Vendida"].sum()
                axes[1].bar(vendas_agr.index.astype(str), vendas_agr.values, color="#4CAF50")
                axes[1].set_title("Vendas (total por produto)")
                axes[1].tick_params(axis='x', rotation=30)
            else:
                axes[1].text(0.5, 0.5, "Sem vendas registradas", ha="center")
        except Exception:
            axes[1].text(0.5, 0.5, "Erro gerando gráfico de vendas", ha="center")

        plt.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both")
        dashboard_canvas["loaded"] = True

    # ---------- Montagem das abas (cria controles e retorna referências) ----------
    # Guardar referência ao tree de produtos para permitir atualizações a partir de vendas
    global tree_produtos_ref
    tree_produtos_ref = [None]  # lista mutável para referenciar dentro de closures

    tree_prod = criar_aba_produtos(frame_prod)
    tree_produtos_ref[0] = tree_prod  # referência usada pela aba de vendas para atualizar produtos
    tree_vendas = criar_aba_vendas(frame_vend)
    tree_vendedores = criar_aba_vendedores(frame_vdr)

    # Dashboard: criar somente quando a aba for mostrada (evita crashes ao abrir)
    def on_tab_changed(event):
        sel = event.widget.select()
        tab_text = event.widget.tab(sel, "text")
        if tab_text == "Dashboard":
            # recria dashboard a cada abertura para garantir dados atualizados
            criar_dashboard(frame_dash)

    notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

    # ---------- Menu top (opcional) ----------
    menubar = tk.Menu(root)
    root.config(menu=menubar)
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Arquivo", menu=file_menu)
    file_menu.add_command(label="Recarregar dados", command=lambda: recarregar_tabelas())
    file_menu.add_separator()
    file_menu.add_command(label="Sair", command=root.destroy)

    # ---------- Função para recarregar tudo do disco ----------
    def recarregar_tabelas():
        nonlocal df_produtos, df_vendas, df_vendedores
        dfp, dfv, dfvd = carregar_planilhas()
        if dfp is None:
            return
        df_produtos = dfp
        df_vendas = dfv
        df_vendedores = dfvd
        atualizar_tree(tree_prod, df_produtos)
        atualizar_tree(tree_vendas, df_vendas)
        atualizar_tree(tree_vendedores, df_vendedores)
        messagebox.showinfo("Recarregado", "Dados recarregados do arquivo Excel.")

    # ---------- Responsividade adicional ----------
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    # ---------- Fluxo de inicialização ----------
    root.withdraw()
    ok = tela_login()
    if not ok:
        root.destroy()
        return

    # mostrando janela principal
    root.deiconify()
    root.mainloop()

if __name__ == "__main__":
    main()