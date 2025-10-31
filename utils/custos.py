# utils/custos.py
import pandas as pd
from pathlib import Path

ARQUIVO_CUSTOS = Path("dados/custos_salvos.xlsx")

def carregar_custos(uploaded_file=None):
    """Lê o arquivo enviado ou o salvo localmente."""
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        df.to_excel(ARQUIVO_CUSTOS, index=False)
        return df, True
    elif ARQUIVO_CUSTOS.exists():
        df = pd.read_excel(ARQUIVO_CUSTOS)
        return df, False
    else:
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"]), False

def salvar_custos(df):
    """Salva os custos atualizados localmente."""
    df.to_excel(ARQUIVO_CUSTOS, index=False)
