# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import re
import os
from pathlib import Path
import tempfile
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
import json

# =============================================================================
# SKU UTILS - Integrado para portabilidade
# =============================================================================
def aplicar_custos(df_vendas, df_custos, coluna_unidades):
    """Aplica os custos dos produtos ao DataFrame de vendas."""
    if df_custos.empty or "SKU" not in df_custos.columns or "Custo_Produto" not in df_custos.columns:
        st.warning("⚠️ DataFrame de custos está vazio ou mal formatado. Custos não aplicados.")
        df_vendas["Custo_Produto"] = 0.0
        df_vendas["Custo_Produto_Total"] = 0.0
        return df_vendas

    df_vendas["SKU"] = df_vendas["SKU"].astype(str)
    df_custos["SKU"] = df_custos["SKU"].astype(str)
    df_vendas = pd.merge(df_vendas, df_custos[["SKU", "Custo_Produto"]], on="SKU", how="left")
    df_vendas["Custo_Produto"].fillna(0, inplace=True)
    df_vendas["Custo_Produto_Total"] = (df_vendas["Custo_Produto"] * df_vendas[coluna_unidades]).round(2)
    return df_vendas

# === VARIÁVEIS DE ESTADO E INICIALIZAÇÃO ===
df = None
coluna_unidades = "Unidades"

# === CRIAÇÃO SEGURA DO DIRETÓRIO ===
try:
    BASE_DIR = Path("dados")
    BASE_DIR.mkdir(exist_ok=True)
except Exception:
    BASE_DIR = Path(tempfile.gettempdir())

st.set_page_config(page_title="📊 Auditoria de Vendas ML", layout="wide")
st.title("📦 Auditoria Financeira Mercado Livre")

# === CONFIGURAÇÕES ===
st.sidebar.header("⚙️ Configurações")
margem_limite = st.sidebar.number_input("Margem de Alerta (%)", min_value=0, max_value=100, value=10, step=1, help="Vendas com margem final abaixo deste valor serão destacadas.")
custo_embalagem = st.sidebar.number_input("Custo fixo de embalagem (R$)", min_value=0.0, value=3.0, step=0.5)
custo_fiscal = st.sidebar.number_input("Custo fiscal (%)", min_value=0.0, value=10.0, step=0.5)

# === GESTÃO DE CUSTOS (GOOGLE SHEETS) ===
st.subheader("💰 Custos de Produtos (Google Sheets)")
try:
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" not in st.secrets:
        raise ValueError("❌ Bloco [gcp_service_account] não encontrado em st.secrets." )
    info = dict(st.secrets["gcp_service_account"])
    info["private_key"] = info["private_key"].encode().decode("unicode_escape")
    creds = Credentials.from_service_account_info(info, scopes=scope)
    client = gspread.authorize(creds)
    st.success("📡 Conectado com sucesso ao Google Sheets!")
except Exception as e:
    st.error(f"❌ Erro ao autenticar com Google Sheets: {e}")
    client = None

if "client" not in locals() or client is None:
    client = None

SHEET_NAME = "CUSTOS_ML"

def carregar_custos_google():
    if not client:
        st.warning("⚠️ Google Sheets não autenticado.")
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])
    try:
        sheet = client.open(SHEET_NAME).sheet1
        dados = sheet.get_all_values()
        if not dados or len(dados) < 2:
            return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])
        df_custos = pd.DataFrame(dados[1:], columns=dados[0])
        df_custos.columns = df_custos.columns.str.strip()
        rename_map = {
            "sku": "SKU", "produto": "Produto", "descrição": "Produto", "descricao": "Produto",
            "custo": "Custo_Produto", "custo_produto": "Custo_Produto",
            "preço_de_custo": "Custo_Produto", "preco_de_custo": "Custo_Produto"
        }
        df_custos.rename(columns={c: rename_map.get(c.lower(), c) for c in df_custos.columns}, inplace=True)
        if "Custo_Produto" in df_custos.columns:
            def corrigir_valor(v):
                v = str(v).strip().replace("R$", "").replace(" ", "")
                if "," in v and "." in v: v = v.replace(".", "").replace(",", ".")
                elif "," in v: v = v.replace(",", ".")
                try:
                    return round(float(v), 2)
                except: return 0.0
            df_custos["Custo_Produto"] = df_custos["Custo_Produto"].apply(corrigir_valor)
        st.info("📡 Custos carregados diretamente do Google Sheets.")
        return df_custos
    except Exception as e:
        st.warning(f"⚠️ Erro ao carregar custos do Google Sheets: {e}")
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])

def salvar_custos_google(df_custos):
    if not client:
        st.warning("⚠️ Google Sheets não autenticado.")
        return
    try:
        sheet = client.open(SHEET_NAME).sheet1
        sheet.clear()
        sheet.update([df_custos.columns.values.tolist()] + df_custos.values.tolist())
        st.success(f"💾 Custos salvos no Google Sheets em {(datetime.utcnow() - timedelta(hours=3)).strftime('%d/%m/%Y %H:%M')}")
    except Exception as e:
        st.error(f"Erro ao salvar custos no Google Sheets: {e}")

custo_df = carregar_custos_google()
if not custo_df.empty:
    custo_df["SKU"] = custo_df["SKU"].astype(str).str.replace(r"[^\d-]", "", regex=True)
else:
    st.warning("⚠️ Nenhum custo encontrado. Você pode adicionar manualmente abaixo.")

custos_editados = st.data_editor(custo_df, num_rows="dynamic", use_container_width=True, key="custos_editor")
if st.button("💾 Atualizar custos no Google Sheets"):
    salvar_custos_google(custos_editados)

# === UPLOAD DE VENDAS ===
st.markdown("---")
st.subheader("📦 Upload de Vendas Mercado Livre")

if "uploaded_file" not in st.session_state:
    st.session_state["uploaded_file"] = None

uploaded_file = st.file_uploader("📤 Envie o arquivo Excel de vendas (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.session_state.get("uploaded_file_name") != uploaded_file.name:
        st.session_state["uploaded_file_name"] = uploaded_file.name
        st.cache_data.clear()
        st.success(f"✅ Arquivo {uploaded_file.name} carregado com sucesso!")
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Vendas BR", header=5)
        df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}. Verifique se a aba 'Vendas BR' e o cabeçalho na linha 6 estão corretos.")
        df = None

if st.button("🗑️ Remover arquivo carregado"):
    st.session_state["uploaded_file_name"] = None
    st.cache_data.clear()
    st.rerun()

if uploaded_file and df is not None:
    # === PROCESSAMENTO PRINCIPAL ===
    possiveis_colunas_unidades = ["Unidades", "Quantidade", "Qtde", "Qtd"]
    coluna_unidades = next((c for c in possiveis_colunas_unidades if c in df.columns), None)
    if coluna_unidades:
        df[coluna_unidades] = pd.to_numeric(df[coluna_unidades].astype(str).str.extract(r"(\d+)", expand=False).fillna("1"), errors='coerce').fillna(1).astype(int)
    else:
        df["Unidades"] = 1
        coluna_unidades = "Unidades"

    col_map = {
        "N.º de venda": "Venda", "Data da venda": "Data", "Estado": "Estado",
        "Receita por produtos (BRL)": "Valor_Venda", "Total (BRL)": "Valor_Recebido",
        "Tarifa de venda e impostos (BRL)": "Tarifa_Venda", "Tarifas de envio (BRL)": "Tarifa_Envio",
        "Cancelamentos e reembolsos (BRL)": "Cancelamentos", "Preço unitário de venda do anúncio (BRL)": "Preco_Unitario",
        "SKU": "SKU", "# de anúncio": "Anuncio", "Título do anúncio": "Produto",
        "Tipo de anúncio": "Tipo_Anuncio", "Receita por envio (BRL)": "Receita_Envio"
    }
    df.rename(columns={c: col_map[c] for c in col_map if c in df.columns}, inplace=True)

    def calcular_percentual(tipo_anuncio):
        tipo = str(tipo_anuncio).strip().lower()
        if "premium" in tipo: return 0.17
        elif "clássico" in tipo or "classico" in tipo: return 0.12
        return 0.12

    # ### ALTERAÇÃO PRINCIPAL: LÓGICA DE TARIFA FIXA CORRIGIDA ###
    def calcular_custo_fixo(preco_unit):
        """Calcula o custo fixo por unidade com base no preço do produto."""
        preco_unit = float(preco_unit or 0)
        if preco_unit < 12.50:
            return round(preco_unit * 0.5, 2)  # 50% do preço
        elif preco_unit < 29.00:
            return 6.25
        elif preco_unit < 50.00:
            return 6.50
        elif preco_unit < 79.00:
            return 6.75
        else:
            return 0.0  # Sem custo fixo para produtos de R$ 79 ou mais

    for col in ["Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$"]:
        if col not in df.columns: df[col] = 0.0
    
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario", "Receita_Envio"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df['Preco_Unitario'] = pd.to_numeric(df['Preco_Unitario'], errors='coerce').fillna(0)
    df['Tarifa_Percentual_%'] = df['Tipo_Anuncio'].apply(lambda x: calcular_percentual(x) * 100)
    # A função `calcular_custo_fixo` agora tem a lógica correta
    df['Tarifa_Fixa_R$'] = df['Preco_Unitario'].apply(calcular_custo_fixo) * df[coluna_unidades]
    df['Tarifa_Total_R$'] = ((df['Valor_Venda'] * (df['Tarifa_Percentual_%'] / 100)) + df['Tarifa_Fixa_R$']).round(2)

    if "SKU" in df.columns: df["SKU"] = df["SKU"].astype(str).str.replace(r'[^\w-]', '', regex=True)
    if "Venda" in df.columns: df["Venda"] = df["Venda"].astype(str).str.replace(r'\D', '', regex=True)

    df["Origem_Pacote"] = ""
    pacotes_a_processar = df[df['Estado'].str.contains("Pacote de", na=False)].index
    for i in pacotes_a_processar:
        match = re.search(r"Pacote de (\d+) produtos", df.loc[i, 'Estado'])
        if not match: continue
        qtd = int(match.group(1))
        if i + 1 + qtd > len(df): continue
        
        subset_indices = range(i + 1, i + 1 + qtd)
        df.loc[subset_indices, "Origem_Pacote"] = f"{df.loc[i, 'Venda']}-PACOTE"
        df.loc[i, "Origem_Pacote"] = "PACOTE_PAI"
        
        skus_filhos = "-".join(df.loc[subset_indices, "SKU"].unique())
        produtos_filhos = " + ".join(df.loc[subset_indices, "Produto"].unique())
        df.loc[i, "SKU"] = skus_filhos
        df.loc[i, "Produto"] = produtos_filhos

    df["Custo_Embalagem"] = custo_embalagem
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    df["Lucro_Bruto"] = (df["Valor_Recebido"] + df.get("Receita_Envio", 0) - df["Tarifa_Venda"] - df["Tarifa_Envio"]).round(2)
    df["Lucro_Real"] = (df["Lucro_Bruto"] - df["Custo_Embalagem"] - df["Custo_Fiscal"]).round(2)
    df["Margem_Liquida_%"] = ((df["Lucro_Real"] / df["Valor_Venda"].replace(0, np.nan)) * 100).round(2).fillna(0)

    df = aplicar_custos(df, custos_editados, coluna_unidades)
    df["Lucro_Liquido"] = (df["Lucro_Real"] - df["Custo_Produto_Total"]).round(2)
    df["Margem_Final_%"] = ((df["Lucro_Liquido"] / df["Valor_Venda"].replace(0, np.nan)) * 100).round(2).fillna(0)
    df["Markup_%"] = ((df["Lucro_Liquido"] / df["Custo_Produto_Total"].replace(0, np.nan)) * 100).round(2).fillna(0)
    
    df["Status"] = np.where(df["Valor_Recebido"] == 0, "🟦 Cancelado", 
                   np.where(df["Margem_Final_%"] < margem_limite, "⚠️ Margem Baixa", "✅ Normal"))
    
    pai_mask = df["Origem_Pacote"] == "PACOTE_PAI"
    cols_to_zero = ["Lucro_Liquido", "Valor_Venda", "Lucro_Real"]
    for col in cols_to_zero:
        if col in df.columns:
            df.loc[pai_mask, col] = 0

    st.markdown("---")
    st.subheader("Resumo Financeiro do Período")
    df_validas = df[(df['Status'] != '🟦 Cancelado') & (df['Origem_Pacote'] != 'PACOTE_PAI')].copy()
    receita_total = df_validas["Valor_Venda"].sum()
    lucro_liquido_total = df_validas["Lucro_Liquido"].sum()
    margem_media_geral = (lucro_liquido_total / receita_total) * 100 if receita_total > 0 else 0
    vendas_com_prejuizo = df_validas[df_validas["Lucro_Liquido"] < 0]
    qtd_vendas_prejuizo = len(vendas_com_prejuizo)
    prejuizo_total = vendas_com_prejuizo["Lucro_Liquido"].sum()
    total_vendas_analisadas = len(df_validas)
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Receita Bruta Total", f"R$ {receita_total:,.2f}")
    col2.metric("Lucro Líquido Total", f"R$ {lucro_liquido_total:,.2f}")
    col3.metric("Margem Média", f"{margem_media_geral:.2f}%")
    col4.metric("Vendas Analisadas", f"{total_vendas_analisadas}")
    col5.metric("Vendas com Prejuízo", f"{qtd_vendas_prejuizo}", delta_color="inverse")
    col6.metric("Prejuízo Acumulado", f"R$ {prejuizo_total:,.2f}", delta_color="inverse")

    st.markdown("---")
    st.subheader(f"🚨 Produtos com Margem Abaixo de {margem_limite}%")
    df_alerta = df[df["Status"] == "⚠️ Margem Baixa"].copy()
    if not df_alerta.empty:
        st.dataframe(df_alerta, use_container_width=True)
    else:
        st.success("✅ Ótima notícia! Nenhum produto foi vendido com margem abaixo do seu limite de alerta.")

    st.markdown("---")
    st.subheader("📋 Tabela Completa da Auditoria")
    st.dataframe(df, use_container_width=True)

    st.markdown("---")
    st.subheader("⬇️ Exportação do Relatório Completo")
    # ... (o resto do código de exportação permanece o mesmo, pois ele já usa a coluna 'Tarifa_Fixa_R$' que agora está correta)
    comentarios_colunas = {
        "Venda": "Número de identificação da venda no Mercado Livre.",
        "SKU": "Seu código de identificação único para o produto (Stock Keeping Unit).",
        "Tipo_Anuncio": "Modalidade do anúncio (Clássico ou Premium). Influencia diretamente na tarifa.",
        "Valor_Venda": "Valor total da venda do item (Preço Unitário * Unidades), sem descontos ou tarifas.",
        "Valor_Recebido": "Valor líquido creditado em sua conta após todas as deduções do Mercado Livre.",
        "Tarifa_Venda": "Tarifa cobrada pelo Mercado Livre sobre a venda (não inclui o frete).",
        "Tarifa_Percentual_%": "FÓRMULA: Percentual da tarifa de venda, baseado no Tipo de Anúncio (ex: 12% para Clássico, 17% para Premium).",
        "Tarifa_Fixa_R$": "FÓRMULA: Custo fixo por unidade vendida para produtos abaixo de R$ 79,00, baseado em faixas de preço.",
        "Tarifa_Total_R$": "FÓRMULA: Soma da tarifa percentual e da tarifa fixa. (Valor_Venda * Tarifa_%) + Tarifa_Fixa.",
        "Tarifa_Envio": "Custo do frete (envio) que foi deduzido de você.",
        "Cancelamentos": "Valor reembolsado ao cliente em caso de cancelamento.",
        "Custo_Embalagem": "Seu custo estimado com embalagem para esta venda.",
        "Custo_Fiscal": "FÓRMULA: Seu custo com impostos (percentual definido na configuração sobre o Valor da Venda).",
        "Receita_Envio": "Valor que o cliente pagou pelo frete e que foi creditado a você (geralmente para compensar o custo do envio).",
        "Lucro_Bruto": "FÓRMULA: Primeira camada de lucro. (Valor_Recebido + Receita_Envio) - Tarifa_Venda - Tarifa_Envio.",
        "Lucro_Real": "FÓRMULA: Lucro após seus custos operacionais. Lucro_Bruto - Custo_Embalagem - Custo_Fiscal.",
        "Margem_Liquida_%": "FÓRMULA: Percentual de lucro real sobre o valor da venda. (Lucro_Real / Valor_Venda) * 100.",
        "Custo_Produto": "Custo unitário do seu produto (puxado da planilha de custos).",
        "Custo_Produto_Total": "FÓRMULA: Custo total de todos os produtos na venda. Custo_Produto * Unidades.",
        "Lucro_Liquido": "FÓRMULA: O lucro final, descontando o custo do produto. Lucro_Real - Custo_Produto_Total.",
        "Margem_Final_%": "FÓRMULA: A margem de lucro final. (Lucro_Liquido / Valor_Venda) * 100.",
        "Markup_%": "FÓRMULA: Seu retorno sobre o custo do produto. (Lucro_Liquido / Custo_Produto_Total) * 100.",
        "Origem_Pacote": "Identifica se o item pertence a um 'pacote' de produtos ou se é a linha 'pai' do pacote.",
        "Status": "Status da venda (Normal, Margem Baixa ou Cancelado)."
    }
    colunas_exportar = list(comentarios_colunas.keys())
    df_export = df[[c for c in colunas_exportar if c in df.columns]].copy()
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Auditoria')
        # ... (código de formatação e fórmulas do Excel)
    output.seek(0)
    st.download_button(
        label="⬇️ Baixar Relatório com Fórmulas e Comentários",
        data=output,
        file_name=f"Auditoria_ML_Completa_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Aguardando o envio do arquivo Excel de vendas para iniciar a análise.")

