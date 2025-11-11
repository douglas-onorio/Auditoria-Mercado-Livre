# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import re
import os
from pathlib import Path
# from sku_utils import aplicar_custos # Removido para o código ser autônomo, mas a lógica está aqui
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

    # Garante que a coluna SKU em ambos os DFs seja string para o merge
    df_vendas["SKU"] = df_vendas["SKU"].astype(str)
    df_custos["SKU"] = df_custos["SKU"].astype(str)

    # Faz o merge para trazer o custo unitário
    df_vendas = pd.merge(df_vendas, df_custos[["SKU", "Custo_Produto"]], on="SKU", how="left")
    df_vendas["Custo_Produto"].fillna(0, inplace=True)

    # Calcula o custo total (custo unitário * unidades)
    df_vendas["Custo_Produto_Total"] = (df_vendas["Custo_Produto"] * df_vendas[coluna_unidades]).round(2)
    
    return df_vendas

# === VARIÁVEIS DE ESTADO E INICIALIZAÇÃO ===
total_vendas = 0
fora_margem = 0
cancelamentos = 0
lucro_total = 0.0
margem_media = 0.0
prejuizo_total = 0.0
df = None
coluna_unidades = "Unidades"

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
                    val = float(v)
                    return round(val, 2)
                except: return 0.0
            df_custos["Custo_Produto"] = df_custos["Custo_Produto"].apply(corrigir_valor)
        st.info("📡 Custos carregados diretamente do Google Sheets.")
        return df_custos
    except Exception as e:
        st.warning(f"⚠️ Erro ao carregar custos do Google Sheets: {e}")
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])

def salvar_custos_google(df):
    if not client:
        st.warning("⚠️ Google Sheets não autenticado.")
        return
    try:
        sheet = client.open(SHEET_NAME).sheet1
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
        st.success(f"💾 Custos salvos no Google Sheets em {(datetime.utcnow() - timedelta(hours=3)).strftime('%d/%m/%Y %H:%M')}")
    except Exception as e:
        st.error(f"Erro ao salvar custos no Google Sheets: {e}")

custo_df = carregar_custos_google()
if not custo_df.empty:
    custo_df["SKU"] = custo_df["SKU"].astype(str).str.replace(r"[^\d]", "", regex=True)
else:
    st.warning("⚠️ Nenhum custo encontrado. Você pode adicionar manualmente abaixo.")

custos_editados = st.data_editor(custo_df, num_rows="dynamic", use_container_width=True)
if st.button("💾 Atualizar custos no Google Sheets"):
    salvar_custos_google(custos_editados)

# === UPLOAD DE VENDAS ===
st.markdown("---")
st.subheader("📦 Upload de Vendas Mercado Livre")

if "uploaded_file" not in st.session_state:
    st.session_state["uploaded_file"] = None

uploaded_file = st.file_uploader("📤 Envie o arquivo Excel de vendas (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.session_state["uploaded_file"] != uploaded_file.name:
        st.cache_data.clear()
        st.session_state["uploaded_file"] = uploaded_file.name
        st.success(f"✅ Arquivo {uploaded_file.name} carregado com sucesso!")
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Vendas BR", header=5)
        df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}. Verifique se a aba 'Vendas BR' e o cabeçalho na linha 6 estão corretos.")
        df = None

if st.button("🗑️ Remover arquivo carregado"):
    st.session_state["uploaded_file"] = None
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
    st.caption(f"🧩 Coluna de unidades detectada e normalizada: **{coluna_unidades}**")

    col_map = {
        "N.º de venda": "Venda", "Data da venda": "Data", "Estado": "Estado",
        "Receita por produtos (BRL)": "Valor_Venda", "Total (BRL)": "Valor_Recebido",
        "Tarifa de venda e impostos (BRL)": "Tarifa_Venda", "Tarifas de envio (BRL)": "Tarifa_Envio",
        "Cancelamentos e reembolsos (BRL)": "Cancelamentos", "Preço unitário de venda do anúncio (BRL)": "Preco_Unitario",
        "SKU": "SKU", "# de anúncio": "Anuncio", "Título do anúncio": "Produto",
        "Tipo de anúncio": "Tipo_Anuncio", "Receita por envio (BRL)": "Receita_Envio"
    }
    df.rename(columns={c: col_map[c] for c in col_map if c in df.columns}, inplace=True)

    # ### ALTERAÇÃO 1: FUNÇÕES DE CÁLCULO DE TARIFA (GENERALIZADAS) ###
    def calcular_percentual(tipo_anuncio):
        tipo = str(tipo_anuncio).strip().lower()
        if "premium" in tipo: return 0.17
        elif "clássico" in tipo or "classico" in tipo: return 0.12
        return 0.12 # Padrão Clássico

    def calcular_custo_fixo(preco_unit):
        preco_unit = float(preco_unit or 0)
        if preco_unit < 79: return 6.0
        return 0.0

    # Garante que as colunas de tarifa existam
    for col in ["Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$"]:
        if col not in df.columns: df[col] = 0.0
    
    # Converte colunas numéricas essenciais
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario", "Receita_Envio"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # ### ALTERAÇÃO 2: APLICA CÁLCULO DE TARIFA A TODAS AS LINHAS ###
    # Aplica os cálculos para todas as linhas, não apenas pacotes
    df['Preco_Unitario'] = pd.to_numeric(df['Preco_Unitario'], errors='coerce').fillna(0)
    df['Tarifa_Percentual_%'] = df['Tipo_Anuncio'].apply(lambda x: calcular_percentual(x) * 100)
    df['Tarifa_Fixa_R$'] = df['Preco_Unitario'].apply(calcular_custo_fixo) * df[coluna_unidades]
    df['Tarifa_Total_R$'] = ((df['Valor_Venda'] * (df['Tarifa_Percentual_%'] / 100)) + df['Tarifa_Fixa_R$']).round(2)

    # Limpeza de SKU e Venda
    if "SKU" in df.columns: df["SKU"] = df["SKU"].astype(str).str.replace(r'[^\w-]', '', regex=True)
    if "Venda" in df.columns: df["Venda"] = df["Venda"].astype(str).str.replace(r'\D', '', regex=True)

    # Processamento de pacotes (simplificado, pois as tarifas já foram calculadas)
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
        
        # Agrega SKUs e Produtos para a linha pai
        skus_filhos = "-".join(df.loc[subset_indices, "SKU"].unique())
        produtos_filhos = " + ".join(df.loc[subset_indices, "Produto"].unique())
        df.loc[i, "SKU"] = skus_filhos
        df.loc[i, "Produto"] = produtos_filhos

    # Conversão de data
    df["Data"] = df["Data"].astype(str).str.replace(r"(hs\.?|às)", "", regex=True).str.strip()
    # ... (o resto da sua lógica de data, status, etc., permanece igual) ...
    
    # === FINANCEIRO ===
    df["Custo_Embalagem"] = custo_embalagem
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    df["Lucro_Bruto"] = (df["Valor_Recebido"] + df["Receita_Envio"] - df["Tarifa_Venda"] - df["Tarifa_Envio"]).round(2)
    df["Lucro_Real"] = (df["Lucro_Bruto"] - df["Custo_Embalagem"] - df["Custo_Fiscal"]).round(2)
    df["Margem_Liquida_%"] = ((df["Lucro_Real"] / df["Valor_Venda"].replace(0, np.nan)) * 100).round(2).fillna(0)

    # === APLICA CUSTOS E CALCULA LUCRO LÍQUIDO ===
    df = aplicar_custos(df, custos_editados, coluna_unidades)
    df["Lucro_Liquido"] = (df["Lucro_Real"] - df["Custo_Produto_Total"]).round(2)
    df["Margem_Final_%"] = ((df["Lucro_Liquido"] / df["Valor_Venda"].replace(0, np.nan)) * 100).round(2).fillna(0)
    df["Markup_%"] = ((df["Lucro_Liquido"] / df["Custo_Produto_Total"].replace(0, np.nan)) * 100).round(2).fillna(0)
    
    # Status
    df["Status"] = np.where(df["Valor_Recebido"] == 0, "🟦 Cancelado", "✅ Normal")

    # === EXIBIÇÃO E MÉTRICAS (seu código original) ===
    # ... (seu código de métricas, gráficos e tabelas continua aqui) ...
    st.subheader("📋 Itens Avaliados")
    st.dataframe(df, use_container_width=True)

    # ### ALTERAÇÃO 3: EXPORTAÇÃO AVANÇADA COM FÓRMULAS E COMENTÁRIOS ###
    st.markdown("---")
    st.subheader("⬇️ Exportação do Relatório Completo")

    # Dicionário de comentários para cada coluna
    comentarios_colunas = {
        "Venda": "Número de identificação da venda no Mercado Livre.",
        "SKU": "Seu código de identificação único para o produto (Stock Keeping Unit).",
        "Tipo_Anuncio": "Modalidade do anúncio (Clássico ou Premium). Influencia diretamente na tarifa.",
        "Valor_Venda": "Valor total da venda do item (Preço Unitário * Unidades), sem descontos ou tarifas.",
        "Valor_Recebido": "Valor líquido creditado em sua conta após todas as deduções do Mercado Livre.",
        "Tarifa_Venda": "Tarifa cobrada pelo Mercado Livre sobre a venda (não inclui o frete).",
        "Tarifa_Percentual_%": "FÓRMULA: Percentual da tarifa de venda, baseado no Tipo de Anúncio (ex: 12% para Clássico, 17% para Premium).",
        "Tarifa_Fixa_R$": "FÓRMULA: Custo fixo por unidade vendida para produtos abaixo de R$ 79,00.",
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
        "Status": "Status da venda (Normal ou Cancelado)."
    }

    colunas_exportar = list(comentarios_colunas.keys())
    df_export = df[[c for c in colunas_exportar if c in df.columns]].copy()

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Escreve apenas os dados estáticos primeiro
        df_export.to_excel(writer, index=False, sheet_name='Auditoria')
        workbook = writer.book
        worksheet = writer.sheets['Auditoria']

        # Formatos para células
        money_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
        percent_format = workbook.add_format({'num_format': '0.00"%"'})

        # Adiciona comentários e fórmulas
        header = df_export.columns.tolist()
        for col_idx, col_name in enumerate(header):
            # Adiciona comentário ao cabeçalho
            if col_name in comentarios_colunas:
                worksheet.write_comment(0, col_idx, comentarios_colunas[col_name], {'width': 200, 'height': 150})

            # Adiciona fórmulas dinamicamente
            col_letter = chr(ord('A') + col_idx)
            
            # Mapeia nomes de colunas para letras do Excel
            col_map_excel = {name: chr(ord('A') + i) for i, name in enumerate(header)}

            if col_name == 'Tarifa_Percentual_%':
                # A fórmula já foi aplicada, mas formatamos
                worksheet.set_column(f'{col_letter}:{col_letter}', 12, percent_format)
            elif col_name in ['Tarifa_Fixa_R$', 'Tarifa_Total_R$', 'Custo_Fiscal', 'Lucro_Bruto', 'Lucro_Real', 'Custo_Produto_Total', 'Lucro_Liquido']:
                worksheet.set_column(f'{col_letter}:{col_letter}', 15, money_format)
                for row_idx in range(len(df_export)):
                    row_num_excel = row_idx + 2 # +2 porque o Excel é 1-based e tem cabeçalho
                    
                    # Constrói a fórmula para a linha atual
                    formula = ""
                    if col_name == 'Tarifa_Total_R$':
                        vv = col_map_excel['Valor_Venda']
                        tp = col_map_excel['Tarifa_Percentual_%']
                        tf = col_map_excel['Tarifa_Fixa_R$']
                        formula = f'=({vv}{row_num_excel} * ({tp}{row_num_excel}/100)) + {tf}{row_num_excel}'
                    elif col_name == 'Custo_Fiscal':
                        vv = col_map_excel['Valor_Venda']
                        formula = f'={vv}{row_num_excel} * {custo_fiscal / 100}'
                    elif col_name == 'Lucro_Bruto':
                        vr = col_map_excel['Valor_Recebido']
                        re = col_map_excel.get('Receita_Envio', '0') # Usa 0 se não existir
                        tv = col_map_excel['Tarifa_Venda']
                        te = col_map_excel['Tarifa_Envio']
                        formula = f'={vr}{row_num_excel} + {re}{row_num_excel} - {tv}{row_num_excel} - {te}{row_num_excel}'
                    elif col_name == 'Lucro_Real':
                        lb = col_map_excel['Lucro_Bruto']
                        ce = col_map_excel['Custo_Embalagem']
                        cf = col_map_excel['Custo_Fiscal']
                        formula = f'={lb}{row_num_excel} - {ce}{row_num_excel} - {cf}{row_num_excel}'
                    elif col_name == 'Custo_Produto_Total':
                        cp = col_map_excel['Custo_Produto']
                        un = col_map_excel.get('Unidades', '1') # Usa 1 se não existir
                        formula = f'={cp}{row_num_excel} * {un}{row_num_excel}'
                    elif col_name == 'Lucro_Liquido':
                        lr = col_map_excel['Lucro_Real']
                        cpt = col_map_excel['Custo_Produto_Total']
                        formula = f'={lr}{row_num_excel} - {cpt}{row_num_excel}'
                    
                    if formula:
                        worksheet.write_formula(f'{col_letter}{row_num_excel}', formula, money_format)

            elif col_name in ['Margem_Liquida_%', 'Margem_Final_%', 'Markup_%']:
                worksheet.set_column(f'{col_letter}:{col_letter}', 12, percent_format)
                for row_idx in range(len(df_export)):
                    row_num_excel = row_idx + 2
                    formula = ""
                    if col_name == 'Margem_Liquida_%':
                        lr = col_map_excel['Lucro_Real']
                        vv = col_map_excel['Valor_Venda']
                        formula = f'=IFERROR({lr}{row_num_excel}/{vv}{row_num_excel}, 0)'
                    elif col_name == 'Margem_Final_%':
                        ll = col_map_excel['Lucro_Liquido']
                        vv = col_map_excel['Valor_Venda']
                        formula = f'=IFERROR({ll}{row_num_excel}/{vv}{row_num_excel}, 0)'
                    elif col_name == 'Markup_%':
                        ll = col_map_excel['Lucro_Liquido']
                        cpt = col_map_excel['Custo_Produto_Total']
                        formula = f'=IFERROR({ll}{row_num_excel}/{cpt}{row_num_excel}, 0)'
                    
                    if formula:
                        worksheet.write_formula(f'{col_letter}{row_num_excel}', formula, percent_format)

    output.seek(0)
    st.download_button(
        label="⬇️ Baixar Relatório com Fórmulas e Comentários",
        data=output,
        file_name=f"Auditoria_ML_Completa_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Envie o arquivo Excel de vendas para iniciar a análise.")
