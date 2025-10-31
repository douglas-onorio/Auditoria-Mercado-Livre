import streamlit as st
import pandas as pd
import plotly.express as px
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

> **Diferença (%) = (1 - (Valor Recebido ÷ Valor da Venda)) × 100**

Vendas com diferença **acima de {margem_limite}%** são classificadas como **anormais**.
"""
)

# === UPLOAD PRINCIPAL ===
uploaded_file = st.file_uploader("Envie o arquivo Excel de vendas (.xlsx)", type=["xlsx"])

# === UPLOAD FUTURO DE CUSTOS ===
st.sidebar.markdown("📦 **Integração futura de custo interno**")
uploaded_custo = st.sidebar.file_uploader("Planilha de custos (opcional)", type=["xlsx"])

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

        # === DISCLAIMER ABAIXO DO PERÍODO ===
        st.markdown(
            """
            ---
            ⚖️ **Critérios e metodologia dos cálculos**
            
            Todos os valores apresentados são baseados nos dados reais fornecidos pelo Mercado Livre:
            
            - **Tarifa de venda e impostos (BRL)** → custo fixo + comissão do tipo de anúncio  
              ▫️ *Anúncio Clássico:* 12% de comissão  
              ▫️ *Anúncio Premium:* 17% de comissão  
            - **Tarifas de envio (BRL)** → parte do frete paga pelo vendedor conforme peso e faixa de preço  
            - **Cálculos adicionais aplicados:**  
              ▫️ Custo de embalagem fixo → configurável pelo usuário  
              ▫️ Custo fiscal (%) → aplicado sobre o valor de venda  
            - **Lucro Real = Valor da venda − Tarifas ML − Custo de embalagem − Custo fiscal**
            
            🔹 *Etapas futuras:* será possível anexar uma planilha com o **custo real do produto** (SKU, PRODUTO, CUSTO, OBSERVAÇÕES),  
            para calcular automaticamente o **Lucro Líquido** e a **Margem Final** de cada item.
            ---
            """
        )

    df["Data"] = df["Data"].dt.strftime("%d/%m/%Y %H:%M")

    # === CÁLCULOS ===
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

    # === CÁLCULO FINANCEIRO ===
    df["Custo_Embalagem"] = custo_embalagem
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    df["Lucro_Bruto"] = (df["Valor_Venda"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"])).round(2)
    df["Lucro_Real"] = (df["Lucro_Bruto"] - (df["Custo_Embalagem"] + df["Custo_Fiscal"])).round(2)
    df["Margem_Liquida_%"] = ((df["Lucro_Real"] / df["Valor_Venda"]) * 100).round(2)

    # === FUTURA INTEGRAÇÃO DE CUSTO INTERNO ===
    if uploaded_custo:
        try:
            custo_df = pd.read_excel(uploaded_custo)
            custo_df.columns = custo_df.columns.str.strip()
            custo_df["SKU"] = custo_df["SKU"].astype(str).str.strip()
            custo_df.rename(columns={"CUSTO": "Custo_Produto"}, inplace=True)
            df = df.merge(custo_df[["SKU", "Custo_Produto"]], on="SKU", how="left")
            df["Lucro_Liquido"] = (df["Lucro_Real"] - df["Custo_Produto"].fillna(0)).round(2)
            df["Margem_Final_%"] = ((df["Lucro_Liquido"] / df["Valor_Venda"]) * 100).round(2)
        except Exception as e:
            st.error(f"Erro ao processar planilha de custos: {e}")

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

    # === GRÁFICOS ===
    st.markdown("---")
    st.subheader("📊 Análise Gráfica de Lucro e Margem")

    top_lucro = df.groupby("Produto", as_index=False)["Lucro_Real"].sum().sort_values(by="Lucro_Real", ascending=False).head(10)
    top_margem = df.groupby("Produto", as_index=False)["Margem_Liquida_%"].mean().sort_values(by="Margem_Liquida_%").head(10)

    fig1 = px.bar(top_lucro, x="Lucro_Real", y="Produto", orientation="h", title="💰 Top 10 Produtos por Lucro Real (R$)")
    fig2 = px.bar(top_margem, x="Margem_Liquida_%", y="Produto", orientation="h", title="📉 Top 10 Menores Margens (%)")

    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)

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
