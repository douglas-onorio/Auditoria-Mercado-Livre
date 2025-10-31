# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import re
import os
from pathlib import Path
import tempfile

# === CRIAÇÃO SEGURA DO DIRETÓRIO ===
try:
    BASE_DIR = Path("dados")
    BASE_DIR.mkdir(exist_ok=True)
except Exception:
    BASE_DIR = Path(tempfile.gettempdir())

ARQUIVO_CUSTOS_SALVOS = BASE_DIR / "custos_salvos.xlsx"

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

# === GESTÃO DE CUSTOS (NOVO BLOCO) ===
def carregar_custos(uploaded_file=None):
    """Lê custos de upload ou do arquivo salvo."""
    if uploaded_file:
        df_custos = pd.read_excel(uploaded_file)
        df_custos.columns = df_custos.columns.str.strip()
        try:
            df_custos.to_excel(ARQUIVO_CUSTOS_SALVOS, index=False)
            st.success("📥 Nova planilha de custos salva automaticamente!")
        except Exception:
            st.warning("⚠️ Ambiente em modo protegido — custos não foram salvos permanentemente.")
        return df_custos
    elif ARQUIVO_CUSTOS_SALVOS.exists():
        st.info("📂 Custos carregados automaticamente do arquivo salvo.")
        return pd.read_excel(ARQUIVO_CUSTOS_SALVOS)
    else:
        st.warning("⚠️ Nenhum custo encontrado. Envie ou edite manualmente para criar um novo.")
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])

def salvar_custos(df):
    try:
        df.to_excel(ARQUIVO_CUSTOS_SALVOS, index=False)
        st.success("💾 Custos atualizados e salvos com sucesso!")
    except Exception:
        st.warning("⚠️ Ambiente em modo protegido — custos não foram salvos permanentemente.")

uploaded_custo = st.sidebar.file_uploader("📦 Planilha de custos (opcional)", type=["xlsx"])
custo_df = carregar_custos(uploaded_custo)

# === AJUSTE VISUAL DE SKUS EM CUSTOS ===
if not custo_df.empty and "SKU" in custo_df.columns:
    custo_df["SKU"] = custo_df["SKU"].astype(str)
    custo_df["SKU"] = custo_df["SKU"].str.replace(r"[^\d]", "", regex=True)

st.markdown("---")
st.subheader("✏️ Edição de Custos")
custos_editados = st.data_editor(custo_df, num_rows="dynamic")
if st.button("💾 Salvar custos atualizados"):
    salvar_custos(custos_editados)


# === UPLOAD DE VENDAS ===
uploaded_file = st.file_uploader("Envie o arquivo Excel de vendas (.xlsx)", type=["xlsx"])

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

    # Renomeia apenas o que consta no mapeamento
    df.rename(columns={c: col_map[c] for c in col_map if c in df.columns}, inplace=True)

            # Renomeia apenas o que consta no mapeamento
    df.rename(columns={c: col_map[c] for c in col_map if c in df.columns}, inplace=True)

    # === AJUSTE DE PACOTES AGRUPADOS (usando preço unitário real dos itens) ===
    import re

    for i, row in df.iterrows():
        estado = str(row.get("Estado", ""))
        match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
        if not match:
            continue

        qtd = int(match.group(1))
        subset = df.iloc[i + 1 : i + 1 + qtd].copy()
        if subset.empty:
            continue

        # --- Totais da linha do pacote ---
        total_venda = float(row.get("Valor_Venda", 0) or 0)
        total_recebido = float(row.get("Valor_Recebido", 0) or 0)
        total_envio = float(row.get("Receita por envio (BRL)", 0) or 0)
        total_tarifa = float(row.get("Tarifa_Venda", 0) or 0)
        total_acrescimo = float(row.get("Receita por acréscimo no preço (pago pelo comprador)", 0) or 0)

        # --- Preço unitário real de cada item ---
        col_preco_unitario = "Preco_Unitario" if "Preco_Unitario" in subset.columns else "Preço unitário de venda do anúncio (BRL)"
        subset["Preco_Unitario_Item"] = pd.to_numeric(
            subset[col_preco_unitario], errors="coerce"
        ).fillna(0)

        soma_precos = subset["Preco_Unitario_Item"].sum() or qtd

            # --- Redistribuição por item com base no tipo de anúncio e preço unitário ---

    def calcular_custo_fixo(preco_unit):
        """Custo fixo ML conforme faixa de preço."""
        if preco_unit < 12.5:
            return preco_unit * 0.5
        elif preco_unit < 30:
            return 6.25
        elif preco_unit < 50:
            return 6.50
        elif preco_unit < 79:
            return 6.75
        else:
            return 0.0  # acima de 79 entra na regra de frete grátis

    def calcular_percentual(tipo_anuncio):
        """Percentual de tarifa conforme tipo de anúncio."""
        tipo = str(tipo_anuncio).strip().lower()
        if "premium" in tipo:
            return 0.17
        else:
            return 0.12  # Clássico por padrão

    for j in subset.index:
        preco_unit = float(subset.loc[j, "Preco_Unitario_Item"] or 0)
        tipo_anuncio = subset.loc[j, "Tipo_Anuncio"]

        # --- Cálculo detalhado da tarifa ---
        perc = calcular_percentual(tipo_anuncio)
        custo_fixo = calcular_custo_fixo(preco_unit)
        tarifa_individual = round(preco_unit * perc + custo_fixo, 4)

        # --- Redistribuição dos valores financeiros do pacote ---
        proporcao = preco_unit / soma_precos
        df.loc[j, "Valor_Venda"] = total_venda * proporcao
        df.loc[j, "Valor_Recebido"] = total_recebido * proporcao
        df.loc[j, "Tarifa_Venda"] = tarifa_individual
        df.loc[j, "Tarifa_Envio"] = 0.0
        df.loc[j, "Receita por acréscimo no preço (pago pelo comprador)"] = total_acrescimo * proporcao

    # --- Marca o pacote como processado (fora do loop interno) ---
    df.loc[i, "Estado"] = f"{estado} (processado)"
    df.loc[i, ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio"]] = 0


    # --- Marca o pacote como processado (mas mantém a linha) ---
    df.loc[i, "Estado"] = f"{estado} (processado)"
    df.loc[i, ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio"]] = 0

    st.info("📦 Pacotes redistribuídos com base no preço unitário real dos produtos.")
                        
    # === COLUNA DE UNIDADES ===
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
        valor = str(valor).strip()

        # Remove espaços, pontos, vírgulas e símbolos
        valor = re.sub(r"[^\d]", "", valor)

        # Remove zeros à esquerda, mas mantém "0" se for o único dígito
        valor = valor.lstrip("0") or "0"

        # Garante formato limpo (apenas números)
        return valor

    if "SKU" in df.columns:
        df["SKU"] = df["SKU"].apply(limpar_sku)

    # === AJUSTE VENDA ===
    def formatar_venda(valor):
        if pd.isna(valor):
            return ""
        return re.sub(r"[^\d]", "", str(valor))
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

    # Se houver receita de envio, soma ao cálculo (senão, considera 0)
    if "Receita por envio (BRL)" in df.columns:
        df["Receita_Envio"] = pd.to_numeric(df["Receita por envio (BRL)"], errors="coerce").fillna(0)
    else:
        df["Receita_Envio"] = 0

    # Lucro Bruto agora considera a receita de envio
    df["Lucro_Bruto"] = (
        df["Valor_Venda"] + df["Receita_Envio"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"])
    ).round(2)

    df["Lucro_Real"] = (
        df["Lucro_Bruto"] - (df["Custo_Embalagem"] + df["Custo_Fiscal"])
    ).round(2)

    df["Margem_Liquida_%"] = ((df["Lucro_Real"] / df["Valor_Venda"]) * 100).round(2)


    # === PLANILHA DE CUSTOS ===
    custo_carregado = False
    if not custo_df.empty:
        try:
            custo_df["SKU"] = custo_df["SKU"].astype(str).str.strip()
            df = df.merge(custo_df[["SKU", "Custo_Produto"]], on="SKU", how="left")
            df["Custo_Produto_Total"] = df["Custo_Produto"].fillna(0) * df[coluna_unidades]
            df["Lucro_Liquido"] = df["Lucro_Real"] - df["Custo_Produto_Total"]
            df["Margem_Final_%"] = ((df["Lucro_Liquido"] / df["Valor_Venda"]) * 100).round(2)
            df["Markup_%"] = ((df["Lucro_Liquido"] / df["Custo_Produto_Total"]) * 100).round(2)
            custo_carregado = True
        except Exception as e:
            st.error(f"Erro ao aplicar custos: {e}")

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

        # === ANÁLISE DE TIPOS DE ANÚNCIO ===
    st.markdown("---")
    st.subheader("📊 Análise por Tipo de Anúncio (Clássico x Premium)")

    if "Tipo_Anuncio" in df.columns:
        tipo_counts = df["Tipo_Anuncio"].value_counts().reset_index()
        tipo_counts.columns = ["Tipo de Anúncio", "Quantidade"]
        tipo_counts["% Participação"] = (tipo_counts["Quantidade"] / tipo_counts["Quantidade"].sum() * 100).round(2)

        col1, col2 = st.columns(2)
        col1.metric("Anúncios Clássicos", int(tipo_counts.loc[tipo_counts["Tipo de Anúncio"].str.contains("Clássico", case=False), "Quantidade"].sum()))
        col2.metric("Anúncios Premium", int(tipo_counts.loc[tipo_counts["Tipo de Anúncio"].str.contains("Premium", case=False), "Quantidade"].sum()))

        st.dataframe(tipo_counts, use_container_width=True)

        # Exporta o resumo para Excel
        output_tipos = BytesIO()
        with pd.ExcelWriter(output_tipos, engine="xlsxwriter") as writer:
            tipo_counts.to_excel(writer, index=False, sheet_name="Tipos_Anuncio")
        output_tipos.seek(0)
        st.download_button(
            label="⬇️ Exportar Resumo de Tipos (Excel)",
            data=output_tipos,
            file_name=f"Resumo_Tipos_Anuncio_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("⚠️ Nenhuma coluna de tipo de anúncio encontrada no arquivo enviado.")


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


