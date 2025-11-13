# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import re
import os
from pathlib import Path
from sku_utils import aplicar_custos
import tempfile
import numpy as np

# === VARIÁVEIS DE ESTADO E INICIALIZAÇÃO PARA EVITAR NAMEERROR ===
# Inicializando as variáveis que seriam usadas no bloco de métricas,
# garantindo que elas existam mesmo sem um arquivo carregado.
total_vendas = 0
fora_margem = 0
cancelamentos = 0
lucro_total = 0.0
margem_media = 0.0
prejuizo_total = 0.0
df = None # Inicializa o DataFrame principal como None
coluna_unidades = "Unidades" # Inicializa a coluna de unidades

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

# Margem Limite com balão de informação
margem_limite = st.sidebar.number_input(
    "Margem limite (%)",
    min_value=0,
    max_value=100,
    value=30,
    step=1,
    help="Define o percentual máximo de taxas (Tarifas ML + Frete) que você considera aceitável por venda. Vendas com taxas acima deste limite serão marcadas como 'Fora da Margem' para análise."
)

# Custo de Embalagem com balão de informação
custo_embalagem = st.sidebar.number_input(
    "Custo fixo de embalagem (R$)",
    min_value=0.0,
    value=3.0,
    step=0.5,
    help="Custo fixo em R$ para a embalagem de cada venda. Em vendas de pacotes, este valor é rateado entre os itens."
)

# Custo Fiscal com balão de informação
custo_fiscal = st.sidebar.number_input(
    "Custo fiscal (%)",
    min_value=0.0,
    value=10.0,
    step=0.5,
    help="Percentual de imposto (Ex: Simples Nacional) que incide sobre o 'Valor da Venda'. O valor é calculado individualmente para cada item vendido."
)


st.sidebar.markdown(
    f"""
💡 **Lógica da análise de margem:**

> **Diferença (%) = (1 - (Valor Recebido ÷ Valor da Venda)) × 100**

Vendas com diferença **acima de {margem_limite}%** são classificadas como **anormais**.
"""
)

# === GESTÃO DE CUSTOS (INTEGRAÇÃO GOOGLE SHEETS) ===
import gspread
from google.oauth2.service_account import Credentials
import json

st.subheader("💰 Custos de Produtos (Google Sheets)")

client = None
try:
    # Escopos obrigatórios do Google Sheets e Drive
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    if "gcp_service_account" not in st.secrets:
        # Se estiver rodando localmente sem secrets, pode ser um problema.
        raise ValueError("❌ Bloco [gcp_service_account] não encontrado em st.secrets.")

    info = dict(st.secrets["gcp_service_account"])

    # Corrige quebras de linha na private_key
    info["private_key"] = info["private_key"].encode().decode("unicode_escape")

    # Autentica e conecta
    creds = Credentials.from_service_account_info(info, scopes=scope)
    client = gspread.authorize(creds)
    st.success("📡 Conectado com sucesso ao Google Sheets!")

except Exception as e:
    st.error(f"❌ Erro ao autenticar com Google Sheets: {e}")
    client = None

SHEET_NAME = "CUSTOS_ML"  # nome da planilha no Google Sheets

def carregar_custos_google():
    """Lê custos diretamente do Google Sheets e corrige formato pt-BR."""
    if not client:
        st.warning("⚠️ Google Sheets não autenticado.")
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])
    try:
        sheet = client.open(SHEET_NAME).sheet1
        dados = sheet.get_all_values()  # pega TUDO como texto (não tenta converter)
        if not dados or len(dados) < 2:
            return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])

        # Constrói DataFrame manualmente
        df_custos = pd.DataFrame(dados[1:], columns=dados[0])
        df_custos.columns = df_custos.columns.str.strip()

        # 🔧 Normaliza nomes de colunas
        rename_map = {
            "sku": "SKU",
            "produto": "Produto",
            "descrição": "Produto",
            "descricao": "Produto",
            "custo": "Custo_Produto",
            "custo_produto": "Custo_Produto",
            "preço_de_custo": "Custo_Produto",
            "preco_de_custo": "Custo_Produto"
        }
        df_custos.rename(columns={c: rename_map.get(c.lower(), c) for c in df_custos.columns}, inplace=True)

        # 🔢 Converte custos respeitando o formato BR e ajusta escala corretamente
        if "Custo_Produto" in df_custos.columns:
            def corrigir_valor(v):
                v = str(v).strip()
                if v in ["", "-", "nan", "N/A", "None"]:
                    return 0.0

                v = v.replace("R$", "").replace(" ", "")
                # Detecta o padrão de separadores
                if "," in v and "." in v:
                    # Ex: 1.234,56 → 1234.56
                    v = v.replace(".", "").replace(",", ".")
                elif "," in v and "." not in v:
                    # Ex: 162,49 → 162.49
                    v = v.replace(",", ".")
                elif "." in v and "," not in v:
                    # Ex: 162.49 → 162.49 (mantém)
                    pass

                try:
                    val = float(v)
                    # Corrige apenas valores absurdos (erro de escala)
                    # A lógica de correção de escala foi mantida como estava no script original
                    if val > 999:
                        val = val / 100
                    return round(val, 2)
                except:
                    return 0.0

            df_custos["Custo_Produto"] = df_custos["Custo_Produto"].apply(corrigir_valor)

        st.info("📡 Custos carregados diretamente do Google Sheets.")
        return df_custos

    except Exception as e:
        st.warning(f"⚠️ Erro ao carregar custos do Google Sheets: {e}")
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])

def salvar_custos_google(df):
    """Atualiza custos diretamente no Google Sheets."""
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

# === BLOCO VISUAL ===
st.markdown("---")
st.subheader("💰 Custos de Produtos (Google Sheets)")

custo_df = carregar_custos_google()
if not custo_df.empty:
    # A limpeza aqui foi mantida, mas a junção de dados será feita mais tarde.
    custo_df["SKU"] = custo_df["SKU"].astype(str).str.replace(r"[^\d]", "", regex=True)
else:
    st.warning("⚠️ Nenhum custo encontrado. Você pode adicionar manualmente abaixo.")

custos_editados = st.data_editor(custo_df, num_rows="dynamic", use_container_width=True)
if st.button("💾 Atualizar custos no Google Sheets"):
    salvar_custos_google(custos_editados)

# === UPLOAD DE VENDAS ===
st.markdown("---")
st.subheader("📦 Upload de Vendas Mercado Livre")

# === CONTROLE DE UPLOAD / REINICIALIZAÇÃO ===
if "uploaded_file" not in st.session_state:
    st.session_state["uploaded_file"] = None

uploaded_file = st.file_uploader("📤 Envie o arquivo Excel de vendas (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Se o arquivo mudou, limpa cache e atualiza
    if st.session_state["uploaded_file"] != uploaded_file.name:
        st.cache_data.clear()
        st.session_state["uploaded_file"] = uploaded_file.name
        st.success(f"✅ Arquivo {uploaded_file.name} carregado com sucesso!")

    # --- LEITURA COMPLETA ---
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Vendas BR", header=5)
        df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)
        st.dataframe(df.head(20), use_container_width=True)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}. Verifique se a aba 'Vendas BR' e o cabeçalho na linha 6 estão corretos.")
        df = None # Define df como None se houver erro

# Botão para limpar o arquivo e forçar reload
if st.button("🗑️ Remover arquivo carregado"):
    st.session_state["uploaded_file"] = None
    st.cache_data.clear()
    st.rerun()

# Inicia o processamento principal se o arquivo foi carregado com sucesso
if uploaded_file and df is not None:
        # A segunda leitura completa (existente no original) foi removida aqui para evitar redundância.
    
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
    
        # === Funções de Cálculo de Tarifa (Mantidas do original) ===
        # A Tarifa Fixa original está complexa, mas mantida para replicar a regra do usuário
        def calcular_tarifa_fixa_unit(preco_unit):
            """Calcula a Tarifa Fixa unitária (R$) com base na lógica fornecida no script original."""
            if preco_unit < 12.5:
                # Replicando a lógica original
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
            """Calcula o percentual de tarifa com base no tipo de anúncio."""
            tipo = str(tipo_anuncio).strip().lower()
            if "premium" in tipo:
                return 0.17
            elif "clássico" in tipo or "classico" in tipo:
                return 0.12
            return 0.12 # Padrão para casos não identificados
    
        # Garante que todas as colunas necessárias existam
        for col in ["Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$", 
                    "Origem_Pacote", "Valor_Item_Total", "Custo_Embalagem", "Tarifa_Venda_Calculada"]:
            if col not in df.columns:
                df[col] = None
    
        # --- Conversões iniciais de valores para processamento
        for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).abs().round(2)
    
        # === PROCESSA PACOTES AGRUPADOS (com cálculo de tarifas e rateio automático) ===
        df_pacotes = df[df["Estado"].astype(str).str.contains("Pacote de", case=False, na=False)].copy()
        
        indices_pacotes_filhos = []
        
        for i, row in df_pacotes.iterrows():
            estado = str(row.get("Estado", ""))
            match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
            if not match:
                df.loc[i, "Origem_Pacote"] = None
                continue
    
            qtd = int(match.group(1))
            
            # Encontra o índice inicial dos itens do pacote (assumindo que estão na sequência)
            idx_inicio = i + 1
            idx_fim = i + 1 + qtd
            
            if idx_fim > len(df):
                st.warning(f"⚠️ Pacote da venda {row.get('Venda', 'N/A')} na linha {i+6} está incompleto e foi ignorado.")
                continue
    
            subset = df.iloc[idx_inicio : idx_fim].copy()
            if subset.empty:
                continue
    
            total_venda_pacote = float(row.get("Valor_Venda", 0) or 0)
            total_recebido_pacote = float(row.get("Valor_Recebido", 0) or 0)
            frete_total_pacote = abs(float(row.get("Tarifa_Envio", 0) or 0))
    
            col_preco_unitario = "Preco_Unitario"
            subset["Preco_Unitario_Item"] = pd.to_numeric(subset[col_preco_unitario], errors="coerce").fillna(0)
    
            soma_precos = subset["Preco_Unitario_Item"].sum() # Soma dos preços unitários dos itens no pacote
            total_unidades_pacote = subset[coluna_unidades].sum() or 1
            
            total_tarifa_percentual_acumulada = 0
            total_tarifa_fixa_acumulada = 0
            
            custo_embalagem_unit = round(float(custo_embalagem) / qtd, 2)
    
            # --- Cálculo e atribuição individual para ITENS FILHOS ---
            for j in subset.index:
                preco_unit = float(subset.loc[j, "Preco_Unitario_Item"] or 0)
                tipo_anuncio = str(subset.loc[j, "Tipo_Anuncio"]).lower()
                unidades_item = subset.loc[j, coluna_unidades]
                
                valor_item_total = preco_unit * unidades_item
                
                perc = calcular_percentual(tipo_anuncio)
                tarifa_fixa = calcular_tarifa_fixa_unit(preco_unit)
    
                tarifa_percentual = round(valor_item_total * perc, 2)
                tarifa_fixa_total_item = round(tarifa_fixa * unidades_item, 2)
                tarifa_total_calculada = round(tarifa_percentual + tarifa_fixa_total_item, 2)
    
                # Rateio do Valor Recebido e Frete (mantido por proporção/unidades)
                proporcao_venda = (preco_unit / soma_precos) if soma_precos else 0
                valor_recebido_item = round(total_recebido_pacote * proporcao_venda, 2)
                proporcao_unidades = unidades_item / total_unidades_pacote
                frete_item = round(frete_total_pacote * proporcao_unidades, 2)
    
                # Atribuição dos valores ao DataFrame principal
                df.loc[j, "Valor_Venda"] = valor_item_total
                df.loc[j, "Valor_Recebido"] = valor_recebido_item
                df.loc[j, "Tarifa_Percentual_%"] = perc * 100
                df.loc[j, "Tarifa_Fixa_R$"] = tarifa_fixa
                # Tarifa_Venda (coluna do ML, agora contendo a tarifa percentual calculada para o rateio)
                df.loc[j, "Tarifa_Venda"] = tarifa_percentual
                df.loc[j, "Tarifa_Venda_Calculada"] = tarifa_percentual
                df.loc[j, "Tarifa_Total_R$"] = tarifa_total_calculada # Tarifa Total (percentual + fixa)
                df.loc[j, "Tarifa_Envio"] = frete_item
                df.loc[j, "Custo_Embalagem"] = custo_embalagem_unit
                df.loc[j, "Origem_Pacote"] = f"{row['Venda']}-PACOTE"
                df.loc[j, "Tipo_Anuncio"] = "Agrupado (Item)"
                
                indices_pacotes_filhos.append(j)
    
                total_tarifa_percentual_acumulada += tarifa_percentual
                total_tarifa_fixa_acumulada += tarifa_fixa_total_item
            
            # Linha mãe (pacote) — mostra totais calculados
            df.loc[i, "Tipo_Anuncio"] = "Agrupado (Pacotes)"
            df.loc[i, "Tarifa_Venda"] = round(total_tarifa_percentual_acumulada, 2) # Tarifa percentual total (pode ser usado para conferência)
            df.loc[i, "Tarifa_Total_R$"] = round(total_tarifa_percentual_acumulada + total_tarifa_fixa_acumulada, 2)
            df.loc[i, "Custo_Embalagem"] = round(float(custo_embalagem), 2)
            df.loc[i, "Tarifa_Percentual_%"] = None
            df.loc[i, "Tarifa_Fixa_R$"] = None
            df.loc[i, "Origem_Pacote"] = "PACOTE"
    
        
        # === CORREÇÃO 1: APLICA TARIFA E TAXA FIXA EM VENDAS NÃO AGRUPADAS (Unitárias) ===
        # Máscara para itens que não são pais e não são filhos (vendas simples)
        mask_unitario = df.index.difference(df_pacotes.index).difference(indices_pacotes_filhos)
    
        for i in mask_unitario:
            row = df.loc[i]
            
            # Garante que Preco_Unitario existe
            preco_unit = float(row.get("Preco_Unitario", 0) or 0)
            tipo_anuncio = str(row.get("Tipo_Anuncio", "")).lower()
            unidades_item = row.get(coluna_unidades, 1)
    
            # O Valor_Venda (Receita por produtos) já é o valor total para esta linha unitária
            valor_item_total = row["Valor_Venda"]
            
            perc = calcular_percentual(tipo_anuncio)
            tarifa_fixa = calcular_tarifa_fixa_unit(preco_unit)
    
            tarifa_percentual = round(valor_item_total * perc, 2)
            tarifa_fixa_total_item = round(tarifa_fixa * unidades_item, 2)
            tarifa_total_calculada = round(tarifa_percentual + tarifa_fixa_total_item, 2)
            
            # A Tarifa_Venda (coluna original do ML) *deve* conter a tarifa total (percentual + fixa).
            # Usamos Tarifa_Total_R$ para conferência e Tarifa_Venda_Calculada para o valor percentual puro.
            df.loc[i, "Tarifa_Percentual_%"] = perc * 100
            df.loc[i, "Tarifa_Fixa_R$"] = tarifa_fixa
            df.loc[i, "Tarifa_Venda_Calculada"] = tarifa_percentual
            df.loc[i, "Tarifa_Total_R$"] = tarifa_total_calculada
            # Custo de Embalagem: aplica o valor cheio
            df.loc[i, "Custo_Embalagem"] = round(float(custo_embalagem), 2)
    
        # === NORMALIZA CAMPOS NUMÉRICOS (Tarifas) ===
        for col_fix in ["Tarifa_Venda", "Tarifa_Fixa_R$", "Tarifa_Total_R$", "Tarifa_Envio", "Custo_Embalagem", "Tarifa_Venda_Calculada"]:
            if col_fix in df.columns:
                df[col_fix] = pd.to_numeric(df[col_fix], errors="coerce").fillna(0).abs().round(2)
    
        # === CORREÇÃO 2: REFORÇA O RATEIO DO CUSTO DE EMBALAGEM ===
        # Este bloco garante que o rateio de embalagem seja aplicado de forma consistente
        mask_mae = df["Estado"].astype(str).str.contains("Pacote de", case=False, na=False)
        mask_filho = df["Origem_Pacote"].astype(str).str.endswith("-PACOTE", na=False)
        
        # 1. Recalcula e aplica custo de embalagem para pacotes e filhos (garantindo correção)
        for idx in df.loc[mask_mae].index:
            venda_pai = df.loc[idx, "Venda"]
            filhos = df[df["Origem_Pacote"] == f"{venda_pai}-PACOTE"]
            if not filhos.empty:
                qtd = len(filhos)
                custo_unit = round(float(custo_embalagem) / qtd, 2)
                df.loc[filhos.index, "Custo_Embalagem"] = custo_unit
                df.loc[idx, "Custo_Embalagem"] = round(custo_unit * qtd, 2)
            else:
                 # Se for mãe de pacote sem filhos válidos, assume custo total
                 df.loc[idx, "Custo_Embalagem"] = round(float(custo_embalagem), 2)
    
        # 2. Aplica custo de embalagem total para vendas unitárias/simples
        df.loc[~mask_mae & ~mask_filho, "Custo_Embalagem"] = round(float(custo_embalagem), 2)
    
        # === VALIDAÇÃO DOS PACOTES (Melhorada para usar Tarifa Total Calculada) ===
        df["Tarifa_Validada_ML"] = ""
        for pacote in df.loc[mask_filho, "Origem_Pacote"].unique():
            if not isinstance(pacote, str):
                continue
                
            venda_pai_id = pacote.split("-")[0]
            pai = df[df["Venda"].astype(str).eq(venda_pai_id)]
            filhos = df[df["Origem_Pacote"] == pacote]
                
            if not pai.empty:
                # Soma das tarifas totais calculadas (percentual + fixa) + frete das filhas
                soma_filhas_tarifas = filhos["Tarifa_Total_R$"].sum() + filhos["Tarifa_Envio"].sum()
                
                # Tarifa ML reportada (Tarifa de Venda + Tarifa de Envio do PAI)
                tarifa_pai_ml_reportada = pai["Tarifa_Venda"].iloc[0] + abs(pai["Tarifa_Envio"].iloc[0])
                
                # Usa a Tarifa Total reportada pelo ML como referência para a validação
                df.loc[df["Origem_Pacote"] == pacote, "Tarifa_Validada_ML"] = "✔️" if abs(soma_filhas_tarifas - tarifa_pai_ml_reportada) < 1.01 else "❌"
    
    
        # === AJUSTE SKU ===
        def limpar_sku(valor):
            if pd.isna(valor):
                return ""
            valor = str(valor).strip()
            # Mantém hífens (para pacotes) e remove apenas outros caracteres
            valor = re.sub(r"[^\d\-]", "", valor)
            # Evita limpar SKUs compostos (como 3888-3937)
            if "-" not in valor:
                valor = valor.lstrip("0") or "0"
            return valor
    
        if "SKU" in df.columns:
            df["SKU"] = df["SKU"].apply(limpar_sku)
    
        # === COMPLETA DADOS DE PACOTES COM SKUs E TÍTULOS AGRUPADOS ===
        for i, row in df.loc[mask_mae].iterrows():
            estado = str(row.get("Estado", ""))
            match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
            if not match:
                continue
    
            qtd = int(match.group(1))
    
            idx_inicio = i + 1
            idx_fim = i + 1 + qtd
    
            if idx_fim > len(df):
                continue
    
            subset = df.iloc[idx_inicio : idx_fim].copy()
            if subset.empty:
                continue
    
            # Concatena SKUs e títulos dos filhos
            skus = subset["SKU"].astype(str).replace("nan", "").unique().tolist()
            produtos = subset["Produto"].astype(str).replace("nan", "").unique().tolist()
    
            # Formata SKUs concatenando com hífens, sem duplicar zeros ou nulos
            skus_formatados = [s for s in skus if s and s != "0"]
            sku_concat = "-".join(skus_formatados)
    
            # Se houver mais de dois produtos, simplifica o nome
            if len(produtos) > 2:
                produto_concat = f"{produtos[0]} + {len(produtos)-1} outros"
            else:
                produto_concat = " + ".join([p for p in produtos if p])
    
            # Atualiza apenas se houver algo válido
            if sku_concat:
                df.loc[i, "SKU"] = sku_concat
            if produto_concat:
                df.loc[i, "Produto"] = produto_concat
    
        # Exibe resumo de conferência
        st.write("✅ Pacotes processados (SKU e Produto combinados):")
        st.dataframe(
            df[df["Estado"].str.contains("Pacote", case=False, na=False)][["Venda", "SKU", "Produto"]],
            use_container_width=True,
            height=200
        )
    
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
                Lucro Real = Valor da venda − Tarifas (ML reportadas/calculadas) − Frete − Embalagem − Custo fiscal.<br>
                </div>
                """,
                unsafe_allow_html=True,
            )
    
        df["Data"] = df["Data"].dt.strftime("%d/%m/%Y %H:%M")
    
        # === AUDITORIA E CUSTOS INICIAIS ===
        # A Tarifa_Venda é a tarifa PERCENTUAL calculada no loop de pacotes/unitários.
        # O Valor_Recebido é o Total (BRL) do ML, que já é líquido das taxas.
        df["Verificacao_Cancelamento"] = df["Valor_Venda"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"] + df["Cancelamentos"])
        df["Cancelamento_Correto"] = (df["Valor_Recebido"] == 0) & (abs(df["Verificacao_Cancelamento"]) <= 0.1)
        df["Diferença_R$"] = df["Valor_Venda"] - df["Valor_Recebido"]
        
        # Adiciona tratamento de divisão por zero
        df["%Diferença"] = ((1 - (df["Valor_Recebido"] / df["Valor_Venda"].replace(0, np.nan))) * 100).round(2).fillna(0)
        
        df["Status"] = df.apply(
            lambda x: "🟦 Cancelamento Correto" if x["Cancelamento_Correto"]
            else "⚠️ Acima da Margem" if x["%Diferença"] > margem_limite
            else "✅ Normal", axis=1
        )
    
        df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)
    
        # Se houver receita de envio, soma ao cálculo (senão, considera 0)
        if "Receita por envio (BRL)" in df.columns:
            df["Receita_Envio"] = pd.to_numeric(df["Receita por envio (BRL)"], errors="coerce").fillna(0)
        else:
            df["Receita_Envio"] = 0
    
        # Lucro Bruto agora considera a Receita_Envio e as tarifas TOTAL (Tarifa_Venda original + Taxa Fixa, que é a Tarifa_Total_R$)
        # Para ser coerente, usaremos a coluna Tarifa_Total_R$ que foi calculada/ajustada (Tarifa % + Taxa Fixa) para o Lucro Bruto.
        # Se a Tarifa_Total_R$ for 0 (caso o cálculo falhe), usaremos a Tarifa_Venda original do ML (Valor líquido).
        
        # Cria uma coluna de tarifa ML Líquida: usa Tarifa_Total_R$ se for calculada, senão usa a Tarifa_Venda do ML (que é líquida)
        df["Tarifa_Total_Liquida"] = df.apply(
            lambda row: row["Tarifa_Total_R$"] if row["Origem_Pacote"] is not None or row["Tarifa_Total_R$"] > 0 else row["Tarifa_Venda"],
            axis=1
        )
        df["Tarifa_Total_Liquida"] = df["Tarifa_Total_Liquida"].abs().round(2)
        
        df["Lucro_Bruto"] = (
            df["Valor_Venda"] + df["Receita_Envio"] - (df["Tarifa_Total_Liquida"] + df["Tarifa_Envio"])
        ).round(2)
    
        df["Lucro_Real"] = (
            df["Lucro_Bruto"] - (df["Custo_Embalagem"] + df["Custo_Fiscal"])
        ).round(2)
    
        # === PLANILHA DE CUSTOS (SEGUNDO BLOCO DE CÁLCULO) ===
        custo_carregado = False
        if not custo_df.empty:
            try:
                custo_df["SKU"] = custo_df["SKU"].astype(str).str.strip()
    
                df = aplicar_custos(df, custo_df, coluna_unidades)
    
                # --- Custo Fiscal e Embalagem ---
                # Garante que as colunas existam após o merge
                if "Custo_Fiscal" not in df.columns:
                     df["Custo_Fiscal"] = 0.0
                if "Custo_Embalagem" not in df.columns:
                     df["Custo_Embalagem"] = 0.0
                else:
                     df["Custo_Embalagem"] = pd.to_numeric(df["Custo_Embalagem"], errors="coerce").fillna(0)
                
                # Garante que Custo_Produto_Total exista
                if "Custo_Produto_Total" not in df.columns:
                    df["Custo_Produto_Total"] = 0.0
    
                # --- Lucro e Margens completas ---
                # Lucro Líquido = Lucro Real (já com fiscal/embalagem) - Custo do Produto Total
                df["Lucro_Liquido"] = (df["Lucro_Real"] - df["Custo_Produto_Total"]).round(2)
    
                df["Margem_Final_%"] = (
                    (df["Lucro_Liquido"] / df["Valor_Venda"].replace(0, np.nan)) * 100
                ).round(2)
    
                df["Markup_%"] = (
                    (df["Lucro_Liquido"] / df["Custo_Produto_Total"].replace(0, np.nan)) * 100
                ).round(2)
    
                custo_carregado = True
            except Exception as e:
                st.error(f"Erro ao aplicar custos: {e}")
    
        # Garante que as colunas existam para o bloco de métricas, mesmo que o merge de custo falhe
        if "Margem_Final_%" not in df.columns:
            df["Margem_Final_%"] = 0.0
        if "Lucro_Liquido" not in df.columns:
            df["Lucro_Liquido"] = df["Lucro_Real"].copy()
        
        # Define Margem_Liquida_% (baseada em Lucro_Real para o caso sem custos de produto)
        df["Margem_Liquida_%"] = (
            (df["Lucro_Real"] / df["Valor_Venda"].replace(0, np.nan)) * 100
        ).round(2).fillna(0)
    
    
        # === AJUSTE FINAL: ZERA PACOTES APÓS REDISTRIBUIÇÃO ===
        if "Estado" in df.columns:
            mask_pacotes = df["Estado"].str.contains("Pacote de", case=False, na=False)
            campos_financeiros = [
                "Lucro_Real", "Lucro_Liquido", "Margem_Liquida_%",
                "Margem_Final_%", "Markup_%", "Lucro_Bruto",
                "Custo_Produto_Total", "Tarifa_Total_Liquida", "Tarifa_Total_R$" # Zera as colunas de custo/lucro da linha mãe
            ]
            for campo in campos_financeiros:
                if campo in df.columns:
                    df.loc[mask_pacotes, campo] = 0.0
            df.loc[mask_pacotes, "Status"] = "🔹 Pacote Agrupado (Somente Controle)"
    
        # === EXCLUI CANCELAMENTOS DO CÁLCULO ===
        df_validas = df[df["Status"] != "🟦 Cancelamento Correto"].copy() # Cria uma cópia para evitar SettingWithCopyWarning
        
        # Exclui também os pais de pacotes
        mask_validas = ~df_validas["Estado"].astype(str).str.contains("Pacote de", case=False, na=False, regex=False)
        df_validas = df_validas[mask_validas]   
    
    # === MÉTRICAS FINAIS (CÁLCULO) ===
        if custo_carregado:
            lucro_total = df_validas["Lucro_Liquido"].sum()
            prejuizo_total = abs(df_validas.loc[df_validas["Lucro_Liquido"] < 0, "Lucro_Liquido"].sum())
            margem_media = df_validas["Margem_Final_%"].replace([np.inf, -np.inf], np.nan).mean()
        else:
            lucro_total = df_validas["Lucro_Real"].sum()
            prejuizo_total = abs(df_validas.loc[df_validas["Lucro_Real"] < 0, "Lucro_Real"].sum())
            margem_media = df_validas["Margem_Liquida_%"].replace([np.inf, -np.inf], np.nan).mean()
    
        receita_total = df_validas["Valor_Venda"].sum()
        total_vendas = len(df[~df["Estado"].astype(str).str.contains("Pacote de", case=False, na=False, regex=False)]) - cancelamentos
        fora_margem = (df["Status"] == "⚠️ Acima da Margem").sum()
        cancelamentos = (df["Status"] == "🟦 Cancelamento Correto").sum()
    
        # === MÉTRICAS FINAIS (EXIBIÇÃO) ===
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.metric("Total de Vendas", total_vendas)
        col2.metric("Fora da Margem", fora_margem)
        col3.metric("Cancelamentos Corretos", cancelamentos)
        col4.metric("Lucro Total (R$)", f"{lucro_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        col5.metric("Margem Média (%)", f"{margem_media:.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        col6.metric("🔻 Prejuízo Total (R$)", f"{prejuizo_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    # Ajusta o formato de números para o padrão BR
        if uploaded_file and df is not None:
            # === ANÁLISE DE TIPOS DE ANÚNCIO ===
            st.markdown("---")
            st.subheader("📊 Análise por Tipo de Anúncio (Clássico x Premium)")
        
            if "Tipo_Anuncio" in df.columns:
                # Corrige campos vazios e preenche pacotes
                df["Tipo_Anuncio"] = (
                    df["Tipo_Anuncio"]
                    .astype(str)
                    .str.strip()
                    .replace(["nan", "None", ""], "Unitário/Simples") # Ajustado para refletir o que é um item não agrupado
                )
                
                # Filtra as linhas 'mãe' de pacotes para o resumo estatístico
                mask_nao_mae = ~df["Estado"].astype(str).str.contains("Pacote de", case=False, na=False, regex=False)
                df_tipos = df[mask_nao_mae].copy()
        
                tipo_counts = df_tipos["Tipo_Anuncio"].value_counts(dropna=False).reset_index()
                tipo_counts.columns = ["Tipo de Anúncio", "Quantidade"]
                tipo_counts["% Participação"] = (
                    tipo_counts["Quantidade"] / tipo_counts["Quantidade"].sum() * 100
                ).round(2)
        
                col1, col2 = st.columns(2)
                col1.metric(
                    "Anúncios Clássicos",
                    int(
                        tipo_counts.loc[
                            tipo_counts["Tipo de Anúncio"].str.contains("Clássico", case=False, na=False),
                            "Quantidade"
                        ].sum()
                    ),
                )
                col2.metric(
                    "Anúncios Premium",
                    int(
                        tipo_counts.loc[
                            tipo_counts["Tipo de Anúncio"].str.contains("Premium", case=False, na=False),
                            "Quantidade"
                        ].sum()
                    ),
                )
        
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
        df_alerta = df[df["Status"] == "⚠️ Acima da Margem"].copy()
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
                cols_to_display = [
                    "Venda", "Data", "Valor_Venda", "Valor_Recebido", "Tarifa_Venda",
                    "Tarifa_Envio", "Lucro_Real", "%Diferença"
                ]
                exemplo_display = exemplo[[c for c in cols_to_display if c in exemplo.columns]]
                st.write(exemplo_display)
    
    
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
                cols_to_display = [
                    "Produto", "Valor_Venda", "Tarifa_Total_R$", "Tarifa_Envio",
                    "Custo_Embalagem", "Custo_Fiscal", "Lucro_Bruto", "Lucro_Real",
                    coluna_unidades, "Margem_Liquida_%"
                ]
                # Usa Tarifa_Total_R$ para mostrar a tarifa calculada (percentual + fixa)
                filtro_display = filtro[[c for c in cols_to_display if c in filtro.columns]]
                st.write(filtro_display.dropna(axis=1, how="all"))
    
        # === VISUALIZAÇÃO DOS DADOS ANALISADOS ===
        st.markdown("---")
        st.subheader("📋 Itens Avaliados")
    
        colunas_vis = [
            "Venda", "Data", "Produto", "SKU", "Tipo_Anuncio", "Status",
            coluna_unidades, "Valor_Venda", "Valor_Recebido",
            "Tarifa_Venda", "Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$",
            "Tarifa_Envio", "Cancelamentos", "Custo_Embalagem", "Custo_Fiscal",
            "Custo_Produto_Total", "Lucro_Liquido", "Margem_Final_%", "Markup_%"
        ]
        
        # Filtra as colunas existentes
        colunas_finais = [c for c in colunas_vis if c in df.columns]
    
        st.dataframe(df[colunas_finais].sort_values("Data", ascending=False), use_container_width=True)
    
        # Exportar DataFrame Completo
        output_df = BytesIO()
        with pd.ExcelWriter(output_df, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Auditoria_Completa")
        output_df.seek(0)
        # === CORREÇÃO PONTUAL: MARGENS ERRADAS EM PACOTES AGRUPADOS ===
        # Identifica linhas-mãe de pacotes (ex: "Pacote de X produtos")
        mask_pacote_mae = df["Estado"].astype(str).str.contains("Pacote de", case=False, na=False)
        
        # Nessas linhas, zera margens e markups, pois não fazem sentido financeiro direto
        df.loc[mask_pacote_mae, ["Margem_Liquida_%", "Margem_Final_%", "Markup_%"]] = 0.0
    
        # Para itens filhos de pacotes, recalcula margem apenas se o Valor_Venda for válido
        mask_pacote_filho = df["Origem_Pacote"].astype(str).str.endswith("-PACOTE", na=False)
        if "Lucro_Liquido" in df.columns and "Valor_Venda" in df.columns:
            df.loc[mask_pacote_filho, "Margem_Final_%"] = (
                df.loc[mask_pacote_filho, "Lucro_Liquido"] /
                df.loc[mask_pacote_filho, "Valor_Venda"].replace(0, np.nan)
            ).clip(-500, 500).round(4)
    
        if "Lucro_Real" in df.columns and "Valor_Venda" in df.columns:
            df.loc[mask_pacote_filho, "Margem_Liquida_%"] = (
                df.loc[mask_pacote_filho, "Lucro_Real"] /
                df.loc[mask_pacote_filho, "Valor_Venda"].replace(0, np.nan)
            ).clip(-500, 500).round(4)
    
        if "Lucro_Liquido" in df.columns and "Custo_Produto_Total" in df.columns:
            df.loc[mask_pacote_filho, "Markup_%"] = (
                df.loc[mask_pacote_filho, "Lucro_Liquido"] /
                df.loc[mask_pacote_filho, "Custo_Produto_Total"].replace(0, np.nan)
            ).clip(-500, 500).round(4)
    
    # === EXPORTAÇÃO FINAL COMPLETA COM FÓRMULAS E CORES (VERSÃO FINAL CORRIGIDA) ===
        st.markdown("---")
        st.subheader("📤 Exportar Relatório de Auditoria Completo")
    
        colunas_exportar = [
            "Venda", "SKU", "Unidades", "Tipo_Anuncio",
            "Valor_Venda", "Valor_Recebido",
            "Tarifa_Venda", "Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$",
            "Tarifa_Envio", "Cancelamentos",
            "Custo_Embalagem", "Custo_Fiscal", "Receita_Envio",
            "Lucro_Bruto", "Lucro_Real", "Margem_Liquida_%",
            "Custo_Produto_Unitario", "Custo_Produto_Total",
            "Lucro_Liquido", "Margem_Final_%", "Markup_%",
            "Origem_Pacote", "Status"
        ]
        df_export = df[[c for c in colunas_exportar if c in df.columns]].copy()
    
    # Converte % para fração ANTES de exportar
        for col in ["Tarifa_Percentual_%", "Margem_Liquida_%", "Margem_Final_%", "Markup_%"]:
            if col in df_export.columns:
                df_export[col] = pd.to_numeric(df_export[col], errors='coerce').apply(lambda x: x / 100 if pd.notna(x) and abs(x) > 1 else x).fillna(0)
        
        output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Auditoria", header=False, startrow=1)
        wb = writer.book
        ws = writer.sheets["Auditoria"]
    
        # === FORMATOS ===
        # ✅ CABEÇALHO COM FUNDO BRANCO
        fmt_header = wb.add_format({"bold": True, "bg_color": "#FFFFFF", "align": "center", "valign": "vcenter", "border": 1})
    
        # Formatos para linhas normais (fundo branco)
        fmt_money = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_pct = wb.add_format({'num_format': '0.00%', 'border': 1})
        fmt_int = wb.add_format({'num_format': '0', 'border': 1})
        fmt_txt = wb.add_format({'border': 1})
    
        # Formatos para linhas de PACOTE (azul)
        fmt_pacote_money = wb.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#D9E1F2', 'border': 1})
        fmt_pacote_pct = wb.add_format({'num_format': '0.00%', 'bg_color': '#D9E1F2', 'border': 1})
        fmt_pacote_int = wb.add_format({'num_format': '0', 'bg_color': '#D9E1F2', 'border': 1})
        fmt_pacote_txt = wb.add_format({'bg_color': '#D9E1F2', 'border': 1})
    
        # Formatos para linhas de ITEM de pacote (laranja)
        fmt_item_money = wb.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#FCE4D6', 'border': 1})
        fmt_item_pct = wb.add_format({'num_format': '0.00%', 'bg_color': '#FCE4D6', 'border': 1})
        fmt_item_int = wb.add_format({'num_format': '0', 'bg_color': '#FCE4D6', 'border': 1})
        fmt_item_txt = wb.add_format({'bg_color': '#FCE4D6', 'border': 1})
    
        # === APLICA CABEÇALHO E LARGURA DAS COLUNAS ===
        headers = list(df_export.columns)
        ws.set_row(0, 22)
        for j, col_name in enumerate(headers):
            ws.write(0, j, col_name, fmt_header) # Usa o novo formato de cabeçalho
            if col_name in ["Unidades"]: ws.set_column(j, j, 10)
            elif "%" in col_name: ws.set_column(j, j, 12)
            elif any(x in col_name for x in ["Valor", "Lucro", "Custo", "Tarifa", "Receita"]) and "%" not in col_name: ws.set_column(j, j, 16)
            else: ws.set_column(j, j, 20)
    
        # === FUNÇÕES DE AJUDA PARA FÓRMULAS ===
        col_idx = {name: i for i, name in enumerate(headers)}
        def C(name):
            idx = col_idx.get(name, -1)
            if idx == -1: return ""
            s = ""
            while idx >= 0:
                s = chr(idx % 26 + 65) + s
                idx = idx // 26 - 1
            return s
    
        # === LOOP PRINCIPAL PARA APLICAR FORMATOS E FÓRMULAS ===
        for i, (idx, row_data) in enumerate(df_export.iterrows(), start=2):
            tipo_anuncio = str(row_data.get("Tipo_Anuncio", "")).lower()
            is_mae_pacote = "agrupado (pacotes" in tipo_anuncio
            is_item_pacote = "agrupado (item" in tipo_anuncio
    
            # Escolhe o conjunto de formatos correto para a linha
            if is_mae_pacote:
                formats = {'money': fmt_pacote_money, 'pct': fmt_pacote_pct, 'int': fmt_pacote_int, 'txt': fmt_pacote_txt}
            elif is_item_pacote:
                formats = {'money': fmt_item_money, 'pct': fmt_item_pct, 'int': fmt_item_int, 'txt': fmt_item_txt}
            else:
                formats = {'money': fmt_money, 'pct': fmt_pct, 'int': fmt_int, 'txt': fmt_txt}
    
            # Itera sobre as colunas para aplicar o formato correto a cada célula
            for j, col_name in enumerate(headers):
                cell_value = row_data[col_name]
                fmt = formats['txt'] # Formato padrão
                if col_name in ["Unidades"]: fmt = formats['int']
                elif "%" in col_name: fmt = formats['pct']
                elif any(x in col_name for x in ["Valor", "Lucro", "Custo", "Tarifa", "Receita"]) and "%" not in col_name: fmt = formats['money']
    
                # Escreve o valor com o formato correto (sem fórmulas por enquanto)
                if pd.isna(cell_value):
                    ws.write_blank(i - 1, j, None, fmt)
                elif isinstance(cell_value, (int, float)):
                    ws.write_number(i - 1, j, cell_value, fmt)
                else:
                    ws.write_string(i - 1, j, str(cell_value), fmt)
    
            # Se não for linha-mãe, sobrescreve as células necessárias com FÓRMULAS
            if not is_mae_pacote:
                if all(k in col_idx for k in ["Lucro_Bruto","Valor_Venda","Receita_Envio","Tarifa_Total_R$","Tarifa_Envio"]):
                    ws.write_formula(i-1, col_idx["Lucro_Bruto"], f"=IFERROR({C('Valor_Venda')}{i}+{C('Receita_Envio')}{i}-{C('Tarifa_Total_R$')}{i}-{C('Tarifa_Envio')}{i},0)", formats['money'])
                if all(k in col_idx for k in ["Lucro_Real","Lucro_Bruto","Custo_Embalagem","Custo_Fiscal"]):
                    ws.write_formula(i-1, col_idx["Lucro_Real"], f"=IFERROR({C('Lucro_Bruto')}{i}-{C('Custo_Embalagem')}{i}-{C('Custo_Fiscal')}{i},0)", formats['money'])
                if all(k in col_idx for k in ["Margem_Liquida_%","Lucro_Real","Valor_Venda"]):
                    ws.write_formula(i-1, col_idx["Margem_Liquida_%"], f"=IFERROR({C('Lucro_Real')}{i}/{C('Valor_Venda')}{i},0)", formats['pct'])
                if all(k in col_idx for k in ["Lucro_Liquido","Lucro_Real","Custo_Produto_Total"]):
                    ws.write_formula(i-1, col_idx["Lucro_Liquido"], f"=IFERROR({C('Lucro_Real')}{i}-{C('Custo_Produto_Total')}{i},0)", formats['money'])
                if all(k in col_idx for k in ["Margem_Final_%","Lucro_Liquido","Valor_Venda"]):
                    ws.write_formula(i-1, col_idx["Margem_Final_%"], f"=IFERROR({C('Lucro_Liquido')}{i}/{C('Valor_Venda')}{i},0)", formats['pct'])
                if all(k in col_idx for k in ["Markup_%","Lucro_Liquido","Custo_Produto_Total"]):
                    ws.write_formula(i-1, col_idx["Markup_%"], f"=IFERROR({C('Lucro_Liquido')}{i}/{C('Custo_Produto_Total')}{i},0)", formats['pct'])
    
        # === ABA DE AJUDA (mantida como estava) ===
        ajuda_data = [
            ["Coluna","Descrição","Exemplo"],
            ["Venda","Número da venda no Mercado Livre.","200009741628937"],
            ["SKU","Código interno ou SKU composto (pacote).","3888-3937"],
            ["Unidades","Quantidade vendida.","2"],
            ["Tipo_Anuncio","Clássico (12%), Premium (17%) ou Agrupado.","Premium"],
            ["Valor_Venda","Preço total da venda.","162,49"],
            ["Valor_Recebido","Valor líquido após tarifas.","140,00"],
            ["Tarifa_Venda","Tarifa percentual do ML.","19,49"],
            ["Tarifa_Percentual_%","Percentual da tarifa ML.","12%"],
            ["Tarifa_Fixa_R$","Tarifa fixa cobrada por unidade.","6,75"],
            ["Tarifa_Total_R$","Soma da tarifa percentual + fixa.","26,24"],
            ["Tarifa_Envio","Custo de envio pago.","15,71"],
            ["Cancelamentos","Valores reembolsados.","0,00"],
            ["Custo_Embalagem","Custo fixo ou rateado por pacote.","2,50"],
            ["Custo_Fiscal","% fiscal sobre venda.","16,25"],
            ["Receita_Envio","Valor recebido do comprador (frete).","10,00"],
            ["Lucro_Bruto","Valor_Venda + Receita_Envio − Tarifas − Frete.","135,25"],
            ["Lucro_Real","Lucro_Bruto − Custo_Embalagem − Custo_Fiscal.","116,50"],
            ["Margem_Liquida_%","Lucro_Real ÷ Valor_Venda.","28%"],
            ["Custo_Produto_Unitario","Custo de aquisição unitário.","95,00"],
            ["Custo_Produto_Total","Custo total do item.","190,00"],
            ["Lucro_Liquido","Lucro_Real − Custo_Produto_Total.","55,00"],
            ["Margem_Final_%","Lucro_Liquido ÷ Valor_Venda.","25%"],
            ["Markup_%","Lucro_Liquido ÷ Custo_Produto_Total.","29%"],
            ["Origem_Pacote","Identificador do pacote (mãe/filho).","200009741628937-PACOTE"],
            ["Status","Normal, Fora da Margem ou Cancelamento.","✅ Normal"]
        ]
        df_ajuda = pd.DataFrame(ajuda_data[1:], columns=ajuda_data[0])
        df_ajuda.to_excel(writer, index=False, sheet_name="AJUDA")
        ws_ajuda = writer.sheets["AJUDA"]
        fmt_header_ajuda = wb.add_format({"bold": True, "bg_color": "#92D050", "align": "center", "valign": "vcenter", "border": 1})
        fmt_text_ajuda = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
        fmt_exemplo = wb.add_format({"italic": True, "color": "#666666", "border": 1})
        ws_ajuda.set_row(0, 28, fmt_header_ajuda)
        ws_ajuda.set_column("A:A", 25, fmt_text_ajuda)
        ws_ajuda.set_column("B:B", 80, fmt_text_ajuda)
        ws_ajuda.set_column("C:C", 25, fmt_exemplo)
    
    output.seek(0)
    st.download_button(
        label="⬇️ Baixar Relatório XLSX (com fórmulas, cores e aba AJUDA explicativa)",
        data=output,
        file_name=f"Auditoria_ML_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
