# database.py
import os
import shutil
import pandas as pd
from datetime import datetime
from config import EXCEL_FILE, BACKUP_DIR, LOGIN_SHEET

def backup_excel():
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        dt = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy(EXCEL_FILE, os.path.join(BACKUP_DIR, f"backup_{dt}.xlsx"))
    except Exception:
        # não interrompe a execução se backup falhar
        pass

def carregar_tabelas():
    """
    Retorna um dict com dataframes: 'produtos', 'vendas', 'vendedores', 'login'.
    Se a aba 'Login' não existir, retorna um DataFrame com usuário admin padrão.
    """
    try:
        df_produtos = pd.read_excel(EXCEL_FILE, sheet_name="Produtos")
    except Exception:
        df_produtos = pd.DataFrame(columns=[
            "Código do Produto", "Nome do Produto", "Categoria", "Quantidade",
            "Volume", "Valor de Compra", "Valor de Venda", "Valor de Mercado"
        ])

    try:
        df_vendas = pd.read_excel(EXCEL_FILE, sheet_name="Vendas")
    except Exception:
        df_vendas = pd.DataFrame(columns=["Código de Venda", "Código do Produto", "Nome do Produto", "ID do Vendedor", "Qnt. Vendida"])

    try:
        df_vendedores = pd.read_excel(EXCEL_FILE, sheet_name="Vendedores")
    except Exception:
        df_vendedores = pd.DataFrame(columns=["ID do Vendedor", "Nome", "Telefone", "Email"])

    # login sheet: se não existir, cria um DF com o admin padrão (mas não salva)
    try:
        df_login = pd.read_excel(EXCEL_FILE, sheet_name=LOGIN_SHEET)
    except Exception:
        df_login = pd.DataFrame({"user": ["admin"], "pass": ["1234"]})

    # garantir colunas string
    for df in (df_produtos, df_vendas, df_vendedores, df_login):
        df.columns = df.columns.astype(str)

    return {
        "produtos": df_produtos,
        "vendas": df_vendas,
        "vendedores": df_vendedores,
        "login": df_login
    }

def salvar_tabelas(tabelas):
    """
    tabelas: dict com chaves 'produtos','vendas','vendedores' (padrão).
    Faz backup antes de salvar.
    """
    try:
        backup_excel()
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
            # escrever apenas as folhas esperadas
            if "produtos" in tabelas:
                tabelas["produtos"].to_excel(writer, sheet_name="Produtos", index=False)
            if "vendas" in tabelas:
                tabelas["vendas"].to_excel(writer, sheet_name="Vendas", index=False)
            if "vendedores" in tabelas:
                tabelas["vendedores"].to_excel(writer, sheet_name="Vendedores", index=False)
            # se quiser salvar login, incluir 'login' também
            if "login" in tabelas:
                tabelas["login"].to_excel(writer, sheet_name=LOGIN_SHEET, index=False)
    except Exception as e:
        raise
