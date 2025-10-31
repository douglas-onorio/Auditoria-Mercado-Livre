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

    # === CONVERSÕES ===
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).abs()

    def limpar_sku(valor):
        if pd.isna(valor):
            return ""
        try:
            return str(int(float(str(valor).replace(",", ".").strip())))
        except:
            return str(valor).strip()
    df["SKU"] = df["SKU"].apply(limpar_sku)

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
    periodo_texto = ""
    if pd.notna(data_min) and pd.notna(data_max):
        periodo_texto = f"{data_min.strftime('%d-%m-%Y')}_a_{data_max.strftime('%d-%m-%Y')}"
        st.info(f"📅 **Dados da planilha:** {data_min.strftime('%d/%m/%Y')} até {data_max.strftime('%d/%m/%Y')}")
        st.markdown(
            f"""
            <div style='font-size:13px; color:gray;'>
            ⚖️ <b>Critérios e metodologia dos cálculos</b><br><br>
            Todos os valores apresentados são baseados nos dados reais do Mercado Livre.<br>
            • <b>Tarifa de venda e impostos (BRL):</b> inclui o custo fixo e a comissão do tipo de anúncio.<br>
            • <b>Tarifas de envio (BRL):</b> representam o frete pago pelo vendedor.<br>
            • <b>Custos adicionais:</b> embalagem fixa e custo fiscal (% configurável).<br>
            • <b>Lucro Real = Valor da venda − Tarifas ML − Custo de embalagem − Custo fiscal.</b><br><br>
            🔹 Etapas futuras: será possível anexar uma planilha com o custo real do produto 
            (<i>SKU, PRODUTO, CUSTO, OBSERVAÇÕES</i>), para calcular automaticamente o Lucro Líquido, a Margem Final e o Markup.<br>
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
    def classificar(linha):
        if linha["Cancelamento_Correto"]:
            return "🟦 Cancelamento Correto"
        if linha["%Diferença"] > margem_limite:
            return "⚠️ Acima da Margem"
        return "✅ Normal"
    df["Status"] = df.apply(classificar, axis=1)

    # === FINANCEIRO ===
    df["Custo_Embalagem"] = custo_embalagem
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    df["Lucro_Bruto"] = df["Valor_Venda"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"])
    df["Lucro_Real"] = df["Lucro_Bruto"] - (df["Custo_Embalagem"] + df["Custo_Fiscal"])
    df["Margem_Liquida_%"] = ((df["Lucro_Real"] / df["Valor_Venda"]) * 100).round(2)

    # === PLANILHA DE CUSTOS ===
    if uploaded_custo:
        try:
            custo_df = pd.read_excel(uploaded_custo)
            custo_df.columns = custo_df.columns.str.strip()
            custo_df["SKU"] = custo_df["SKU"].astype(str).str.strip()
            custo_df.rename(columns={"CUSTO": "Custo_Produto"}, inplace=True)
            df = df.merge(custo_df[["SKU", "Custo_Produto"]], on="SKU", how="left")
            df["Lucro_Liquido"] = df["Lucro_Real"] - df["Custo_Produto"].fillna(0)
            df["Margem_Final_%"] = ((df["Lucro_Liquido"] / df["Valor_Venda"]) * 100).round(2)
            df["Markup_%"] = ((df["Lucro_Liquido"] / df["Custo_Produto"]) * 100).round(2)
        except Exception as e:
            st.error(f"Erro ao processar planilha de custos: {e}")

    # === RESUMO ===
    total_vendas = len(df)
    fora_margem = (df["Status"] == "⚠️ Acima da Margem").sum()
    cancelamentos = (df["Status"] == "🟦 Cancelamento Correto").sum()
    lucro_total = df["Lucro_Real"].sum()
    receita_total = df["Valor_Venda"].sum()
    margem_media = (lucro_total / receita_total * 100) if receita_total > 0 else 0
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total de Vendas", total_vendas)
    col2.metric("Fora da Margem", fora_margem)
    col3.metric("Cancelamentos Corretos", cancelamentos)
    col4.metric("Lucro Total (R$)", f"{lucro_total:,.2f}")
    col5.metric("Margem Média (%)", f"{margem_media:.2f}%")

    # === DISCLAIMER COMPLEMENTAR ===
    st.markdown(
        """
        <div style='font-size:13px; color:gray;'>
        ⚙️ <b>Interpretação dos indicadores</b><br>
        • <b>Total de Vendas:</b> quantidade total de registros válidos.<br>
        • <b>Fora da Margem:</b> vendas cuja diferença excede o limite definido.<br>
        • <b>Lucro Total (R$):</b> soma dos lucros reais das vendas analisadas.<br>
        • <b>Margem Média (%):</b> média simples das margens por item.<br><br>
        🧮 <b>Diferença entre Margem e Markup:</b><br>
        • <b>Margem:</b> (Lucro ÷ Valor de Venda) × 100 → mostra quanto do preço é lucro.<br>
        • <b>Markup:</b> (Lucro ÷ Custo do Produto) × 100 → mostra quanto o preço supera o custo.<br>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # === TABELA ===
    st.markdown("---")
    st.subheader("📋 Itens Avaliados")
    st.dataframe(df, use_container_width=True)

    # === ALERTA DE PRODUTO ===
    df_alerta = df[df["Status"] == "⚠️ Acima da Margem"]
    if not df_alerta.empty:
        produto_critico = (
            df_alerta.groupby(["SKU", "Anuncio", "Produto"])
            .size().reset_index(name="Ocorrências")
            .sort_values("Ocorrências", ascending=False).head(1)
        )
        st.warning(
            f"🚨 Produto com mais vendas fora da margem: **{produto_critico.iloc[0]['Produto']}** "
            f"(SKU: {produto_critico.iloc[0]['SKU']} | Anúncio: {produto_critico.iloc[0]['Anuncio']} | "
            f"{produto_critico.iloc[0]['Ocorrências']} ocorrências)"
        )
    else:
        st.success("✅ Nenhum produto com vendas fora da margem no período.")

    # === RESUMO POR TIPO DE ANÚNCIO ===
    st.markdown("---")
    st.subheader("📦 Resumo Financeiro por Tipo de Anúncio")
    resumo = df.groupby("Tipo_Anuncio").agg(
        Vendas=("Venda", "count"),
        Receita_Total=("Valor_Venda", "sum"),
        Lucro_Total=("Lucro_Real", "sum"),
        Margem_Média=("Margem_Liquida_%", "mean"),
    ).reset_index()
    resumo["Receita_Total"] = resumo["Receita_Total"].round(2)
    resumo["Lucro_Total"] = resumo["Lucro_Total"].round(2)
    resumo["Margem_Média"] = resumo["Margem_Média"].round(2)
    st.dataframe(resumo, use_container_width=True)

    # === EXPORTAÇÃO ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Auditoria", freeze_panes=(1, 0))
        writer.sheets["Auditoria"].set_column(0, len(df.columns), 18)
    output.seek(0)
    st.download_button(
        label="⬇️ Baixar Relatório XLSX",
        data=output,
        file_name=f"Auditoria_ML_{periodo_texto or datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Envie o arquivo Excel de vendas para iniciar a análise.")
