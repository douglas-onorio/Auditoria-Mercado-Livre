# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import re

st.set_page_config(page_title="📊 Auditoria de Vendas ML", layout="wide")
st.title("📦 Auditoria Financeira Mercado Livre")

# === CONFIGURAÇÕES ===
st.sidebar.header("⚙️ Configurações")
margem_limite = st.sidebar.number_input("Margem limite (%)", min_value=0, max_value=100, value=30, step=1)
custo_embalagem = st.sidebar.number_input("Custo fixo de embalagem (R$)", min_value=0.0, value=3.0, step=0.5)
custo_fiscal = st.sidebar.number_input("Custo fiscal (%)", min_value=0.0, value=10.0, step=0.5)

st.sidebar.markdown(
    f"""
💡 **Lógica da análise de margem:**

> **Diferença (%) = (1 - (Valor Recebido ÷ Valor da Venda)) × 100**

Vendas com diferença **acima de {margem_limite}%** são classificadas como **anormais**.
"""
)

# === UPLOAD ===
uploaded_file = st.file_uploader("Envie o arquivo Excel de vendas (.xlsx)", type=["xlsx"])
uploaded_custo = st.sidebar.file_uploader("📦 Planilha de custos (opcional)", type=["xlsx"])

if uploaded_file:
    # --- LEITURA COMPLETA ---
    df = pd.read_excel(uploaded_file, sheet_name="Vendas BR", header=5)
    df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)

    # --- MAPEAMENTO PRINCIPAL ---
    col_map = {
        "N.º de venda": "Venda",
        "Data da venda": "Data",
        "Estado": "Estado",
        "Receita por produtos (BRL)": "Valor_Venda",
        "Total (BRL)": "Valor_Recebido",
        "Tarifa de venda e impostos (BRL)": "Tarifa_Venda",
        "Tarifas de envio (BRL)": "Tarifa_Envio",
        "Cancelamentos e reembolsos (BRL)": "Cancelamentos",
        "Preço unitário de venda do anúncio (BRL)": "Preco_Unitario",
        "SKU": "SKU",
        "# de anúncio": "Anuncio",
        "Título do anúncio": "Produto",
        "Tipo de anúncio": "Tipo_Anuncio"
    }

    # Renomeia apenas o que consta no mapeamento, preservando o resto (como "Unidades")
    df.rename(columns={c: col_map[c] for c in col_map if c in df.columns}, inplace=True)

    # === IDENTIFICA E NORMALIZA COLUNA DE UNIDADES ===
    possiveis_colunas_unidades = ["Unidades", "Quantidade", "Qtde", "Qtd"]
    coluna_unidades = next((c for c in possiveis_colunas_unidades if c in df.columns), None)
    if coluna_unidades:
        df[coluna_unidades] = (
            df[coluna_unidades]
            .astype(str)
            .str.strip()
            .replace({"": "1", "-": "1", "–": "1", "—": "1", "nan": "1"}, regex=True)
            .str.extract(r"(\d+)", expand=False)
            .fillna("1")
            .astype(int)
        )
    else:
        df["Unidades"] = 1
        coluna_unidades = "Unidades"

    st.caption(f"🧩 Coluna de unidades detectada e normalizada: **{coluna_unidades}**")

    # === CONVERSÕES ===
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).abs()

    # === AJUSTE SKU ===
    def limpar_sku(valor):
        if pd.isna(valor):
            return ""
        try:
            return str(int(float(str(valor).replace(",", ".").strip())))
        except:
            return str(valor).strip()
    df["SKU"] = df["SKU"].apply(limpar_sku) if "SKU" in df.columns else ""

    # === AJUSTE VENDA ===
    def formatar_venda(valor):
        if pd.isna(valor):
            return ""
        valor_str = re.sub(r"[^\d]", "", str(valor))
        return valor_str
    df["Venda"] = df["Venda"].apply(formatar_venda)

    # === DATA ===
    df["Data"] = df["Data"].astype(str).str.replace(r"(hs\.?|às)", "", regex=True).str.strip()
    meses_pt = {
        "janeiro": "01", "fevereiro": "02", "março": "03", "abril": "04",
        "maio": "05", "junho": "06", "julho": "07", "agosto": "08",
        "setembro": "09", "outubro": "10", "novembro": "11", "dezembro": "12"
    }

    def parse_data_portugues(texto):
        if not isinstance(texto, str) or not any(m in texto.lower() for m in meses_pt):
            return None
        try:
            partes = texto.lower().split(" de ")
            dia = partes[0].zfill(2)
            mes = meses_pt.get(partes[1], "01")
            ano_e_hora = partes[2].split(" ")
            ano = ano_e_hora[0]
            hora = ano_e_hora[1] if len(ano_e_hora) > 1 else "00:00"
            return datetime.strptime(f"{dia}/{mes}/{ano} {hora}", "%d/%m/%Y %H:%M")
        except Exception:
            return None

    df["Data"] = pd.to_datetime(df["Data"].apply(parse_data_portugues), errors="coerce")

    # === PERÍODO ===
    data_min, data_max = df["Data"].min(), df["Data"].max()
    if pd.notna(data_min) and pd.notna(data_max):
        st.info(f"📅 **Período de vendas:** {data_min.strftime('%d/%m/%Y')} → {data_max.strftime('%d/%m/%Y')}")
        st.markdown(
            """
            <div style='font-size:13px; color:gray;'>
            ⚖️ <b>Critérios e metodologia:</b><br>
            Este relatório calcula automaticamente as margens e o lucro com base em:<br>
            • Tarifas e impostos retidos pelo ML.<br>
            • Custos de envio.<br>
            • Custo fixo de embalagem e custo fiscal configurável.<br>
            • Quantidade total de unidades por venda.<br><br>
            Lucro Real = Valor da venda − Tarifas − Frete − Embalagem − Custo fiscal.<br>
            </div>
            """,
            unsafe_allow_html=True,
        )

    df["Data"] = df["Data"].dt.strftime("%d/%m/%Y %H:%M")

    # === AUDITORIA ===
    df["Verificacao_Cancelamento"] = df["Valor_Venda"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"] + df["Cancelamentos"])
    df["Cancelamento_Correto"] = (df["Valor_Recebido"] == 0) & (abs(df["Verificacao_Cancelamento"]) <= 0.1)
    df["Diferença_R$"] = df["Valor_Venda"] - df["Valor_Recebido"]
    df["%Diferença"] = ((1 - (df["Valor_Recebido"] / df["Valor_Venda"])) * 100).round(2)
    df["Status"] = df.apply(
        lambda x: "🟦 Cancelamento Correto" if x["Cancelamento_Correto"]
        else "⚠️ Acima da Margem" if x["%Diferença"] > margem_limite
        else "✅ Normal", axis=1
    )

    # === FINANCEIRO ===
    df["Custo_Embalagem"] = custo_embalagem
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    df["Lucro_Bruto"] = df["Valor_Venda"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"])
    df["Lucro_Real"] = df["Lucro_Bruto"] - (df["Custo_Embalagem"] + df["Custo_Fiscal"])
    df["Margem_Liquida_%"] = ((df["Lucro_Real"] / df["Valor_Venda"]) * 100).round(2)

    # === PLANILHA DE CUSTOS ===
    custo_carregado = False
    if uploaded_custo:
        try:
            custo_df = pd.read_excel(uploaded_custo)
            custo_df.columns = custo_df.columns.str.strip()
            custo_df["SKU"] = custo_df["SKU"].astype(str).str.strip()
            custo_df.rename(columns={"CUSTO": "Custo_Produto"}, inplace=True)
            df = df.merge(custo_df[["SKU", "Custo_Produto"]], on="SKU", how="left")
            df["Custo_Produto_Total"] = df["Custo_Produto"].fillna(0) * df[coluna_unidades]
            df["Lucro_Liquido"] = df["Lucro_Real"] - df["Custo_Produto_Total"]
            df["Margem_Final_%"] = ((df["Lucro_Liquido"] / df["Valor_Venda"]) * 100).round(2)
            df["Markup_%"] = ((df["Lucro_Liquido"] / df["Custo_Produto_Total"]) * 100).round(2)
            custo_carregado = True
        except Exception as e:
            st.error(f"Erro ao processar planilha de custos: {e}")

    # === EXCLUI CANCELAMENTOS DO CÁLCULO ===
    df_validas = df[df["Status"] != "🟦 Cancelamento Correto"]

    # === RESUMO ===
    total_vendas = len(df)
    fora_margem = (df["Status"] == "⚠️ Acima da Margem").sum()
    cancelamentos = (df["Status"] == "🟦 Cancelamento Correto").sum()

    if custo_carregado:
        lucro_total = df_validas["Lucro_Liquido"].sum()
        prejuizo_total = abs(df_validas.loc[df_validas["Lucro_Liquido"] < 0, "Lucro_Liquido"].sum())
        margem_media = df_validas["Margem_Final_%"].mean()
    else:
        lucro_total = df_validas["Lucro_Real"].sum()
        prejuizo_total = abs(df_validas.loc[df_validas["Lucro_Real"] < 0, "Lucro_Real"].sum())
        margem_media = df_validas["Margem_Liquida_%"].mean()

    receita_total = df_validas["Valor_Venda"].sum()

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Total de Vendas", total_vendas)
    col2.metric("Fora da Margem", fora_margem)
    col3.metric("Cancelamentos Corretos", cancelamentos)
    col4.metric("Lucro Total (R$)", f"{lucro_total:,.2f}")
    col5.metric("Margem Média (%)", f"{margem_media:.2f}%")
    col6.metric("🔻 Prejuízo Total (R$)", f"{prejuizo_total:,.2f}")

    # === ALERTA DE PRODUTO ===
    st.markdown("---")
    st.subheader("🚨 Produtos Fora da Margem")
    df_alerta = df[df["Status"] == "⚠️ Acima da Margem"]
    if not df_alerta.empty:
        produto_critico = (
            df_alerta.groupby(["SKU", "Anuncio", "Produto"])
            .size().reset_index(name="Ocorrências")
            .sort_values("Ocorrências", ascending=False).head(1)
        )
        sku_critico = produto_critico.iloc[0]["SKU"]
        produto_nome = produto_critico.iloc[0]["Produto"]
        anuncio_critico = produto_critico.iloc[0]["Anuncio"]
        ocorrencias = produto_critico.iloc[0]["Ocorrências"]

        st.warning(
            f"🚨 Produto com mais vendas fora da margem: **{produto_nome}** "
            f"(SKU: {sku_critico} | Anúncio: {anuncio_critico} | {ocorrencias} ocorrências)"
        )

        exemplo = df_alerta[df_alerta["SKU"] == sku_critico].head(1)
        if not exemplo.empty:
            st.markdown("**🧾 Exemplo de venda afetada:**")
            st.write(exemplo[[
                "Venda", "Data", "Valor_Venda", "Valor_Recebido", "Tarifa_Venda",
                "Tarifa_Envio", "Lucro_Real", "%Diferença"
            ]])

        vendas_afetadas = df_alerta[df_alerta["SKU"] == sku_critico]
        st.markdown("**📄 Todas as vendas afetadas por esse produto:**")
        st.dataframe(vendas_afetadas, use_container_width=True)

        output_alerta = BytesIO()
        with pd.ExcelWriter(output_alerta, engine="xlsxwriter") as writer:
            vendas_afetadas.to_excel(writer, index=False, sheet_name="Fora_da_Margem")
        output_alerta.seek(0)
        st.download_button(
            label="⬇️ Exportar Vendas Afetadas (Excel)",
            data=output_alerta,
            file_name=f"Vendas_Fora_da_Margem_{sku_critico}_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("✅ Nenhum produto com vendas fora da margem no período.")

    # === CONSULTA SKU ===
    st.markdown("---")
    st.subheader("🔎 Conferência Manual de SKU")
    sku_detalhe = st.text_input("Digite o SKU para detalhar:")
    if sku_detalhe:
        filtro = df[df["SKU"].astype(str) == sku_detalhe.strip()]
        if filtro.empty:
            st.warning("Nenhum registro encontrado para este SKU.")
        else:
            st.write(filtro[[
                "Produto", "Valor_Venda", "Tarifa_Venda", "Tarifa_Envio",
                "Custo_Embalagem", "Custo_Fiscal", "Lucro_Bruto", "Lucro_Real",
                "Unidades", "Margem_Liquida_%"
            ]].dropna(axis=1, how="all"))

# === VISUALIZAÇÃO DOS DADOS ANALISADOS ===
st.markdown("---")
st.subheader("📋 Itens Avaliados")

st.dataframe(
    df[
        [
            "Venda", "Data", "Produto", "SKU", "Tipo_Anuncio",
            coluna_unidades, "Valor_Venda", "Valor_Recebido",
            "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos",
            "Lucro_Real", "Margem_Liquida_%", "Status"
        ]
        if "SKU" in df.columns
        else df.columns
    ],
    use_container_width=True,
    height=450
)

# === EXPORTAÇÃO FINAL ===
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False, sheet_name="Auditoria", freeze_panes=(1, 0))
output.seek(0)
st.download_button(
    label="⬇️ Baixar Relatório XLSX",
    data=output,
    file_name=f"Auditoria_ML_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

else:
    st.info("Envie o arquivo Excel de vendas para iniciar a análise.")
