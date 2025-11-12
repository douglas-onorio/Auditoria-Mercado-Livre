import pandas as pd

def aplicar_custos(df_vendas: pd.DataFrame, df_custos: pd.DataFrame, coluna_unidades: str) -> pd.DataFrame:
    """
    Realiza o merge dos custos unitários dos produtos com o DataFrame de vendas
    e calcula o Custo_Produto_Total (Custo Unitário * Unidades).

    Args:
        df_vendas: DataFrame principal de vendas.
        df_custos: DataFrame com as colunas 'SKU' e 'Custo_Produto'.
        coluna_unidades: Nome da coluna que contém a quantidade de unidades vendidas.

    Returns:
        O DataFrame de vendas atualizado com as colunas de custo.
    """
    # Garante que as colunas de merge existam e estejam limpas
    df_custos["SKU"] = df_custos["SKU"].astype(str).str.strip()

    # Merge com os custos (mantendo todas as vendas)
    # Renomeia "Custo_Produto" do df_custos para evitar conflito se já existir
    df_custos_merge = df_custos[["SKU", "Custo_Produto"]].rename(columns={"Custo_Produto": "Custo_Produto_Unitario_Manual"})
    
    df_vendas = df_vendas.merge(
        df_custos_merge,
        on="SKU",
        how="left"
    )

    # Preenche custos ausentes com 0 e garante que seja numérico
    df_vendas["Custo_Produto_Unitario"] = pd.to_numeric(
        df_vendas["Custo_Produto_Unitario_Manual"], errors="coerce"
    ).fillna(0)
    
    # Remove a coluna temporária
    df_vendas.drop(columns=["Custo_Produto_Unitario_Manual"], errors="ignore", inplace=True)

    # Calcula o custo total do produto (unidades * custo unitário)
    df_vendas["Custo_Produto_Total"] = (
        df_vendas["Custo_Produto_Unitario"] * df_vendas[coluna_unidades]
    ).round(2)

    return df_vendas
