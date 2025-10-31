import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import re

st.set_page_config(page_title="📊 Auditoria de Vendas ML", layout="wide")
st.title("📦 Auditoria Financeira Mercado Livre")

# === CONFIGURAÇÃO ===
st.sidebar.header("⚙️ Configurações")
margem_limite = st.sidebar.number_input("Margem limite (%)", min_value=0, max_value=100, value=30, step=1)
custo_embalagem = st.sidebar.number_input("Custo fixo de embalagem (R$)", min_value=0.0, value=3.0, step=0.5)
custo_fiscal = st.sidebar.number_input("Custo fiscal (%)", min_value=0.0, value=10.0, step=0.5)

st.sidebar.markdown(
    f"""
💡 **Lógica da análise de margem:**

A diferença é calculada por:

> **Diferença (%) = (1 - (Valor Recebido ÷ Valor da Venda)) × 100**

➡️ Vendas com diferença **acima de {margem_limite}%** serão sinalizadas como **anormais**.
"""
)

# === UPLOAD ===
uploaded_file = st.file_uploader("Envie o arquivo Excel de vendas (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Vendas BR", header=5)
    df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)

    # --- MAPEAMENTO ---
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

    df = df[[c for c in col_map.keys() if c in df.columns]].rename(columns=col_map)

    # === CONVERSÕES NUMÉRICAS ===
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).abs()

    # === AJUSTE DO SKU ===
    def limpar_sku(valor):
        if pd.isna(valor):
            return ""
        try:
            return str(int(float(str(valor).replace(",", ".").strip())))
        except:
            return str(valor).strip()
    df["SKU"] = df["SKU"].apply(limpar_sku)

    # === TRATAR DATA EM PORTUGUÊS ===
    df["Data"] = df["Data"].astype(str).str.strip()
    df["Data"] = df["Data"].str.replace(r"(hs\.?|às)", "", regex=True).str.strip()

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
            if len(partes) < 3:
                return None
            dia = partes[0].zfill(2)
            mes_nome = partes[1]
            ano_e_hora = partes[2].split(" ")
            mes = meses_pt.get(mes_nome, "01")
            ano = ano_e_hora[0]
            hora = ano_e_hora[1] if len(ano_e_hora) > 1 else "00:00"
            return datetime.strptime(f"{dia}/{mes}/{ano} {hora}", "%d/%m/%Y %H:%M")
        except Exception:
            return None

    df["Data"] = df["Data"].apply(parse_data_portugues)
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

    # === PERÍODO AUTOMÁTICO ===
    data_min = df["Data"].min()
    data_max = df["Data"].max()
    periodo_texto = ""
    if pd.notna(data_min) and pd.notna(data_max):
        periodo_texto = f"{data_min.strftime('%d-%m-%Y')}_a_{data_max.strftime('%d-%m-%Y')}"
        st.info(f"📅 **Dados da planilha:** {data_min.strftime('%d/%m/%Y')} até {data_max.strftime('%d/%m/%Y')}")
    df["Data"] = df["Data"].dt.strftime("%d/%m/%Y %H:%M")

    # === CÁLCULOS DE AUDITORIA ===
    df["Verificacao_Cancelamento"] = (
        df["Valor_Venda"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"] + df["Cancelamentos"])
    ).round(2)
    df["Cancelamento_Correto"] = (df["Valor_Recebido"] == 0) & (abs(df["Verificacao_Cancelamento"]) <= 0.1)
    df["Diferença_R$"] = (df["Valor_Venda"] - df["Valor_Recebido"]).round(2)
    df["%Diferença"] = ((1 - (df["Valor_Recebido"] / df["Valor_Venda"])) * 100).round(2)

    def classificar(linha):
        if linha["Cancelamento_Correto"]:
            return "🟦 Cancelamento Correto"
        if linha["%Diferença"] > margem_limite:
            return "⚠️ Acima da Margem"
        return "✅ Normal"

    df["Status"] = df.apply(classificar, axis=1)

    # === CÁLCULO FINANCEIRO REAL ===
    df["Custo_Embalagem"] = custo_embalagem
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    df["Lucro_Bruto"] = (df["Valor_Venda"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"])).round(2)
    df["Lucro_Real"] = (df["Lucro_Bruto"] - (df["Custo_Embalagem"] + df["Custo_Fiscal"])).round(2)
    df["Margem_Liquida_%"] = ((df["Lucro_Real"] / df["Valor_Venda"]) * 100).round(2)

    # === RESUMO ===
    total_vendas = len(df)
    fora_margem = (df["Status"] == "⚠️ Acima da Margem").sum()
    cancelamentos = (df["Status"] == "🟦 Cancelamento Correto").sum()
    media_dif = df.loc[df["Status"] == "✅ Normal", "%Diferença"].mean()
    lucro_total = df["Lucro_Real"].sum()
    receita_total = df["Valor_Venda"].sum()
    margem_media = (lucro_total / receita_total * 100) if receita_total > 0 else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total de Vendas", total_vendas)
    col2.metric("Fora da Margem", fora_margem)
    col3.metric("Cancelamentos Corretos", cancelamentos)
    col4.metric("Lucro Total (R$)", f"{lucro_total:,.2f}")
    col5.metric("Margem Média (%)", f"{margem_media:.2f}%")

    st.markdown("---")
    st.subheader("📋 Itens Avaliados")
    st.dataframe(df, use_container_width=True)

    df_alerta = df[df["Status"] == "⚠️ Acima da Margem"]
    if not df_alerta.empty:
        produto_critico = (
            df_alerta.groupby(["SKU", "Anuncio", "Produto"])
            .size()
            .reset_index(name="Ocorrências")
            .sort_values("Ocorrências", ascending=False)
            .head(1)
        )
        sku_produto = produto_critico.iloc[0]["SKU"]
        anuncio_id = produto_critico.iloc[0]["Anuncio"]
        nome_produto = produto_critico.iloc[0]["Produto"]
        ocorrencias = produto_critico.iloc[0]["Ocorrências"]
        st.warning(
            f"🚨 Produto com mais vendas fora da margem: **{nome_produto}** "
            f"(SKU: {sku_produto} | Anúncio: {anuncio_id} | {ocorrencias} ocorrências)"
        )
    else:
        st.success("✅ Nenhum produto com vendas fora da margem no período.")

    # === EXPORTAÇÃO XLSX ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Auditoria", freeze_panes=(1, 0))
        workbook = writer.book
        worksheet = writer.sheets["Auditoria"]
        text_fmt = workbook.add_format({'num_format': '@'})
        worksheet.set_column(0, 0, 22, text_fmt)
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
    output.seek(0)

    nome_arquivo = f"Auditoria_ML_{periodo_texto or datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx"
    st.download_button(
        label="⬇️ Baixar Relatório XLSX",
        data=output,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Envie o arquivo Excel de vendas para iniciar a análise.")
