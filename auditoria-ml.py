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
from sku_utils import aplicar_custos

# === VARIÁVEIS DE ESTADO ===
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
import gspread
from google.oauth2.service_account import Credentials
import json

st.subheader("💰 Custos de Produtos (Google Sheets)")

try:
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    if "gcp_service_account" not in st.secrets:
        raise ValueError("❌ Bloco [gcp_service_account] não encontrado em st.secrets.")

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
    """Lê custos diretamente do Google Sheets e corrige formato pt-BR."""
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
            "sku": "SKU", "produto": "Produto", "descrição": "Produto",
            "descricao": "Produto", "custo": "Custo_Produto",
            "custo_produto": "Custo_Produto", "preço_de_custo": "Custo_Produto",
            "preco_de_custo": "Custo_Produto"
        }
        df_custos.rename(columns={c: rename_map.get(c.lower(), c) for c in df_custos.columns}, inplace=True)

        # Converte valores para float
        if "Custo_Produto" in df_custos.columns:
            def corrigir_valor(v):
                v = str(v).strip().replace("R$", "").replace(" ", "")
                if v in ["", "-", "nan", "N/A", "None"]:
                    return 0.0
                if "," in v and "." in v:
                    v = v.replace(".", "").replace(",", ".")
                elif "," in v:
                    v = v.replace(",", ".")
                try:
                    val = float(v)
                    if val > 999:
                        val = val / 100
                    return round(val, 2)
                except:
                    return 0.0
            df_custos["Custo_Produto"] = df_custos["Custo_Produto"].apply(corrigir_valor)

        st.info("📡 Custos carregados do Google Sheets.")
        return df_custos
    except Exception as e:
        st.warning(f"⚠️ Erro ao carregar custos: {e}")
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])

# === BLOCO VISUAL ===
st.markdown("---")
st.subheader("💰 Custos de Produtos (Google Sheets)")

custo_df = carregar_custos_google()
if not custo_df.empty:
    custo_df["SKU"] = custo_df["SKU"].astype(str).str.replace(r"[^\d]", "", regex=True)
else:
    st.warning("⚠️ Nenhum custo encontrado.")

st.dataframe(custo_df, use_container_width=True, height=200)

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
        st.dataframe(df.head(15), use_container_width=True)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        df = None

if st.button("🗑️ Remover arquivo carregado"):
    st.session_state["uploaded_file"] = None
    st.cache_data.clear()
    st.rerun()
# === PROCESSAMENTO PRINCIPAL ===
if uploaded_file and df is not None:

    # === DETECÇÃO E NORMALIZAÇÃO DE UNIDADES ===
    possiveis_colunas_unidades = ["Unidades", "Quantidade", "Qtde", "Qtd"]
    coluna_unidades = next((c for c in possiveis_colunas_unidades if c in df.columns), None)
    if coluna_unidades:
        df[coluna_unidades] = (
            df[coluna_unidades].astype(str).str.strip()
            .replace({"": "1", "-": "1", "–": "1", "—": "1", "nan": "1"}, regex=True)
            .str.extract(r"(\d+)", expand=False)
            .fillna("1")
            .astype(int)
        )
    else:
        df["Unidades"] = 1
        coluna_unidades = "Unidades"

    st.caption(f"🧩 Coluna de unidades detectada: **{coluna_unidades}**")

    # === MAPEAMENTO PRINCIPAL ===
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
    df.rename(columns={c: col_map[c] for c in col_map if c in df.columns}, inplace=True)

    # === FUNÇÕES DE CÁLCULO ===
    def calcular_custo_fixo(preco_unit):
        if preco_unit < 12.5:
            return round(preco_unit * 0.5, 2)
        elif preco_unit < 30:
            return 6.25
        elif preco_unit < 50:
            return 6.50
        elif preco_unit < 79:
            return 6.75
        else:
            return 0.0

    def calcular_percentual(tipo_anuncio):
        tipo = str(tipo_anuncio).strip().lower()
        if "premium" in tipo:
            return 0.17
        elif "clássico" in tipo or "classico" in tipo:
            return 0.12
        return 0.12

    # === GARANTE COLUNAS OBRIGATÓRIAS ===
    for col in ["Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$", "Origem_Pacote", "Tarifa_Envio"]:
        if col not in df.columns:
            df[col] = None

    # === PROCESSA PACOTES AGRUPADOS ===
    for i, row in df.iterrows():
        estado = str(row.get("Estado", ""))
        match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
        if not match:
            df.loc[i, "Origem_Pacote"] = None
            continue

        qtd = int(match.group(1))
        if i + 1 + qtd > len(df):
            st.warning(f"⚠️ Pacote da venda {row.get('Venda', 'N/A')} está incompleto.")
            continue

        subset = df.iloc[i + 1: i + 1 + qtd].copy()
        if subset.empty:
            continue

        total_venda = float(row.get("Valor_Venda", 0) or 0)
        total_recebido = float(row.get("Valor_Recebido", 0) or 0)
        frete_total = float(row.get("Tarifa_Envio", 0) or 0)
        soma_precos = subset["Preco_Unitario"].fillna(0).sum()
        total_unidades = subset[coluna_unidades].sum() or 1

        for j in subset.index:
            preco_unit = float(subset.loc[j, "Preco_Unitario"] or 0)
            tipo_anuncio = subset.loc[j, "Tipo_Anuncio"]
            perc = calcular_percentual(tipo_anuncio)
            custo_fixo = calcular_custo_fixo(preco_unit)
            unidades_item = subset.loc[j, coluna_unidades]

            valor_item_total = preco_unit * unidades_item
            tarifa_total = round(valor_item_total * perc + (custo_fixo * unidades_item), 2)
            proporcao_venda = preco_unit / soma_precos if soma_precos else 0
            valor_recebido_item = round(total_recebido * proporcao_venda, 2)
            proporcao_unidades = unidades_item / total_unidades
            frete_item = round(frete_total * proporcao_unidades, 2)

            df.loc[j, "Valor_Venda"] = valor_item_total
            df.loc[j, "Valor_Recebido"] = valor_recebido_item
            df.loc[j, "Tarifa_Venda"] = tarifa_total
            df.loc[j, "Tarifa_Percentual_%"] = perc * 100
            df.loc[j, "Tarifa_Fixa_R$"] = custo_fixo * unidades_item
            df.loc[j, "Tarifa_Total_R$"] = tarifa_total
            df.loc[j, "Tarifa_Envio"] = frete_item
            df.loc[j, "Origem_Pacote"] = f"{row['Venda']}-PACOTE"

        df.loc[i, "Origem_Pacote"] = "PACOTE"
        df.loc[i, "Tarifa_Venda"] = df.iloc[i + 1: i + 1 + qtd]["Tarifa_Venda"].sum()
        df.loc[i, "Tarifa_Envio"] = frete_total
        df.loc[i, "Valor_Recebido"] = total_recebido

    # === AJUSTE TARIFAS E FRETE PARA VENDAS UNITÁRIAS ===
    if "Origem_Pacote" not in df.columns:
        df["Origem_Pacote"] = None

    mask_unitarios = df["Origem_Pacote"].isna() & df["Tipo_Anuncio"].notna()

    for i, row in df.loc[mask_unitarios].iterrows():
        preco_unit = float(row.get("Preco_Unitario", 0) or 0)
        tipo_anuncio = str(row.get("Tipo_Anuncio", "")).lower().strip()
        unidades = int(row.get(coluna_unidades, 1))
        valor_total_item = preco_unit * unidades

        if "premium" in tipo_anuncio:
            perc = 0.17
        elif "clássico" in tipo_anuncio or "classico" in tipo_anuncio:
            perc = 0.12
        else:
            perc = 0.12

        custo_fixo = calcular_custo_fixo(preco_unit)
        tarifa_total = round(valor_total_item * perc + (custo_fixo * unidades), 2)

        frete_total = float(row.get("Tarifa_Envio", 0) or 0)
        if frete_total > 0 and unidades > 1:
            df.loc[i, "Tarifa_Envio"] = round(frete_total / unidades, 2)

        df.loc[i, "Tarifa_Percentual_%"] = perc * 100
        df.loc[i, "Tarifa_Fixa_R$"] = custo_fixo * unidades
        df.loc[i, "Tarifa_Total_R$"] = tarifa_total
        df.loc[i, "Tarifa_Venda"] = tarifa_total
    # === LIMPEZA E AJUSTES FINAIS DE DADOS ===
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).abs()

    # === AJUSTE DE SKU ===
    def limpar_sku(valor):
        if pd.isna(valor):
            return ""
        valor = str(valor).strip()
        valor = re.sub(r"[^\d\-]", "", valor)
        if "-" not in valor:
            valor = valor.lstrip("0") or "0"
        return valor

    if "SKU" in df.columns:
        df["SKU"] = df["SKU"].apply(limpar_sku)

    # === COMPLETA DADOS DE PACOTES (SKU + PRODUTO) ===
    for i, row in df.iterrows():
        estado = str(row.get("Estado", ""))
        match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
        if not match:
            continue
        qtd = int(match.group(1))
        if i + 1 + qtd > len(df):
            continue
        subset = df.iloc[i + 1 : i + 1 + qtd].copy()
        if subset.empty:
            continue
        skus = subset["SKU"].astype(str).replace("nan", "").unique().tolist()
        produtos = subset["Produto"].astype(str).replace("nan", "").unique().tolist()
        sku_concat = "-".join([s for s in skus if s and s != "0"])
        produto_concat = " + ".join(produtos[:2]) + (f" + {len(produtos)-2} outros" if len(produtos) > 2 else "")
        if sku_concat:
            df.loc[i, "SKU"] = sku_concat
        if produto_concat:
            df.loc[i, "Produto"] = produto_concat

# === EXIBE PACOTES PROCESSADOS ===
if "Estado" in df.columns:
    pacotes_processados = df[df["Estado"].astype(str).str.contains("Pacote", case=False, na=False)][["Venda", "SKU", "Produto"]]
    if not pacotes_processados.empty:
        st.success("✅ Pacotes processados (SKU e Produto combinados):")
        st.dataframe(
            pacotes_processados.drop_duplicates(subset=["Venda", "SKU"]),
            use_container_width=True,
            height=250
        )
else:
    st.info("Nenhuma coluna 'Estado' encontrada no arquivo. Pulando exibição de pacotes.")

    # === AJUSTES DE DATA E VENDA ===
    df["Venda"] = df["Venda"].astype(str).str.replace(r"[^\d]", "", regex=True)
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

    df["Data"] = pd.to_datetime(df["Data"].astype(str).apply(parse_data_portugues), errors="coerce")
    df["Data"] = df["Data"].dt.strftime("%d/%m/%Y %H:%M")

    # === CÁLCULO DE CUSTOS E MARGENS ===
    df["Custo_Embalagem"] = custo_embalagem
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    df = aplicar_custos(df, custo_df, coluna_unidades)

    df["Lucro_Bruto"] = (df["Valor_Venda"] - df["Tarifa_Venda"] - df["Tarifa_Envio"]).round(2)
    df["Lucro_Real"] = (df["Lucro_Bruto"] - df["Custo_Embalagem"] - df["Custo_Fiscal"]).round(2)
    df["Lucro_Liquido"] = (df["Lucro_Real"] - df["Custo_Produto_Total"]).round(2)
    df["Margem_Final_%"] = (df["Lucro_Liquido"] / df["Valor_Venda"].replace(0, np.nan) * 100).round(2)
    df["Markup_%"] = (df["Lucro_Liquido"] / df["Custo_Produto_Total"].replace(0, np.nan) * 100).round(2)
    df["Margem_Liquida_%"] = (df["Lucro_Real"] / df["Valor_Venda"].replace(0, np.nan) * 100).round(2)

    # === EXPORTAÇÃO FINAL ===
    colunas_exportar = [
        "Venda", "SKU", "Tipo_Anuncio", "Valor_Venda", "Valor_Recebido", "Tarifa_Venda",
        "Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$", "Tarifa_Envio",
        "Cancelamentos", "Custo_Embalagem", "Custo_Fiscal", "Receita_Envio",
        "Lucro_Bruto", "Lucro_Real", "Margem_Liquida_%", "Custo_Produto",
        "Custo_Produto_Total", "Lucro_Liquido", "Margem_Final_%", "Markup_%",
        "Origem_Pacote", "Status"
    ]
    df_export = df[[c for c in colunas_exportar if c in df.columns]].copy()

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Auditoria", freeze_panes=(1, 0))
        workbook = writer.book
        worksheet = writer.sheets["Auditoria"]

        # === COMENTÁRIOS ===
        comentarios = {
            "Venda": "Número identificador da venda no Mercado Livre.",
            "SKU": "Código interno ou agrupamento de produtos.",
            "Tipo_Anuncio": "Tipo de anúncio: Clássico ou Premium.",
            "Valor_Venda": "Valor total bruto da venda.",
            "Valor_Recebido": "Valor líquido recebido após tarifas e descontos.",
            "Tarifa_Venda": "Soma das tarifas cobradas pelo ML.",
            "Tarifa_Percentual_%": "Percentual aplicado conforme o tipo de anúncio.",
            "Tarifa_Fixa_R$": "Tarifa fixa cobrada por unidade, conforme faixa de preço.",
            "Tarifa_Total_R$": "Tarifa total do ML = Valor_Venda*% + Tarifa_Fixa.",
            "Tarifa_Envio": "Parte proporcional do frete atribuída ao item.",
            "Cancelamentos": "Valor de reembolso/cancelamento, se houver.",
            "Custo_Embalagem": "Custo fixo configurado no painel lateral.",
            "Custo_Fiscal": "Percentual de custo fiscal sobre o valor da venda.",
            "Receita_Envio": "Receita obtida com o envio (quando aplicável).",
            "Lucro_Bruto": "Valor_Venda - Tarifa_Venda - Tarifa_Envio.",
            "Lucro_Real": "Lucro_Bruto - Custo_Embalagem - Custo_Fiscal.",
            "Margem_Liquida_%": "Lucro_Real ÷ Valor_Venda × 100.",
            "Custo_Produto": "Custo unitário conforme planilha de custos.",
            "Custo_Produto_Total": "Custo total conforme unidades vendidas.",
            "Lucro_Liquido": "Lucro_Real - Custo_Produto_Total.",
            "Margem_Final_%": "Lucro_Liquido ÷ Valor_Venda × 100.",
            "Markup_%": "Lucro_Liquido ÷ Custo_Produto_Total × 100.",
            "Origem_Pacote": "Indica se pertence a pacote ou é item individual.",
            "Status": "Classificação automática: Normal, Fora da Margem, etc."
        }

        for idx, col in enumerate(df_export.columns):
            if col in comentarios:
                worksheet.write_comment(0, idx, comentarios[col])

        # === FÓRMULAS E FORMATAÇÃO ===
        linhas = len(df_export)
        for row in range(1, linhas + 1):
            worksheet.write_formula(f"N{row+1}", f"=D{row+1}-F{row+1}-J{row+1}")  # Lucro_Bruto
            worksheet.write_formula(f"O{row+1}", f"=N{row+1}-L{row+1}-M{row+1}")  # Lucro_Real
            worksheet.write_formula(f"P{row+1}", f"=O{row+1}/D{row+1}*100")       # Margem_Liquida_%
            worksheet.write_formula(f"T{row+1}", f"=O{row+1}-S{row+1}")           # Lucro_Liquido
            worksheet.write_formula(f"U{row+1}", f"=T{row+1}/D{row+1}*100")       # Margem_Final_%
            worksheet.write_formula(f"V{row+1}", f"=T{row+1}/S{row+1}*100")       # Markup_%

        formato_real = workbook.add_format({'num_format': 'R$ #,##0.00'})
        formato_pct = workbook.add_format({'num_format': '0.00%'})
        for col_num, col_name in enumerate(df_export.columns):
            if "R$" in col_name or "Valor" in col_name or "Lucro" in col_name or "Custo" in col_name:
                worksheet.set_column(col_num, col_num, 15, formato_real)
            elif "%" in col_name:
                worksheet.set_column(col_num, col_num, 12, formato_pct)
            else:
                worksheet.set_column(col_num, col_num, 20)

    output.seek(0)
    st.download_button(
        label="⬇️ Baixar Relatório XLSX (com fórmulas e comentários)",
        data=output,
        file_name=f"Auditoria_ML_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Envie o arquivo Excel de vendas para iniciar a análise.")
