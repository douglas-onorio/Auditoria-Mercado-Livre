# sku_utils.py
import re
import pandas as pd

def calcular_custo_sku(sku, mapa_custos):
    if not isinstance(sku, str) or not sku.strip():
        return 0.0
    partes = re.split(r"[-/+,; ]+", sku)
    total = 0.0
    for p in partes:
        p = p.strip()
        if not p:
            continue
        c = mapa_custos.get(p, 0.0)
        if c == 0 and p.isdigit() and p.startswith("0"):
            c = mapa_custos.get(p.lstrip("0"), 0.0)
        total += float(c or 0.0)
    return round(total, 2)

def aplicar_custos(df, custo_df, coluna_unidades="Unidades"):
    if df is None or df.empty or custo_df is None or custo_df.empty:
        return df
    custo_df = custo_df.copy()
    custo_df["SKU"] = custo_df["SKU"].astype(str).str.strip()
    mapa_custos = dict(zip(custo_df["SKU"], custo_df["Custo_Produto"]))

    df["Custo_Produto"] = df["SKU"].astype(str).apply(lambda s: calcular_custo_sku(s, mapa_custos))
    df["Custo_Produto_Total"] = df["Custo_Produto"].fillna(0) * df[coluna_unidades]
    return df
