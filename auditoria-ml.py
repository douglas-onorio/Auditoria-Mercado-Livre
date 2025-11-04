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
import json
import gspread
from google.oauth2.service_account import Credentials

# ==============================================================================
# === 1. VARIÁVEIS DE ESTADO E INICIALIZAÇÃO PARA EVITAR NAMEERROR ===
# ==============================================================================
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
custo_carregado = False # Flag para rastrear se o custo foi aplicado

# === CRIAÇÃO SEGURA DO DIRETÓRIO (usado para referência, não para firestore) ===
try:
    BASE_DIR = Path("dados")
    BASE_DIR.mkdir(exist_ok=True)
except Exception:
    BASE_DIR = Path(tempfile.gettempdir())

st.set_page_config(page_title="📊 Auditoria de Vendas ML", layout="wide")
st.title("📦 Auditoria Financeira Mercado Livre")

# ==============================================================================
# === 2. CONFIGURAÇÕES ===
# ==============================================================================
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

# ==============================================================================
# === 3. GESTÃO DE CUSTOS (INTEGRAÇÃO GOOGLE SHEETS) ===
# ==============================================================================
st.subheader("💰 Custos de Produtos (Google Sheets)")
client = None
SHEET_NAME = "CUSTOS_ML"  # nome da planilha no Google Sheets

try:
    # Escopos obrigatórios do Google Sheets e Drive
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    # Verifica se a secret gcp_service_account está disponível
    if "gcp_service_account" not in st.secrets:
        st.warning("⚠️ Bloco [gcp_service_account] não encontrado em st.secrets. A gestão de custos via Sheets está desabilitada.")
        client = None
    else:
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

def carregar_custos_google():
    """Lê custos diretamente do Google Sheets e corrige formato pt-BR."""
    if not client:
        return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])
    try:
        sheet = client.open(SHEET_NAME).sheet1
        dados = sheet.get_all_values()  # pega TUDO como texto
        if not dados or len(dados) < 2:
            return pd.DataFrame(columns=["SKU", "Produto", "Custo_Produto"])

        # Constrói DataFrame manualmente
        df_custos = pd.DataFrame(dados[1:], columns=dados[0])
        df_custos.columns = df_custos.columns.str.strip()

        # 🔧 Normaliza nomes de colunas
        rename_map = {
            "sku": "SKU", "produto": "Produto", "descrição": "Produto",
            "descricao": "Produto", "custo": "Custo_Produto", "custo_produto": "Custo_Produto",
            "preço_de_custo": "Custo_Produto", "preco_de_custo": "Custo_Produto"
        }
        df_custos.rename(columns={c: rename_map.get(c.lower(), c) for c in df_custos.columns}, inplace=True)

        # 🔢 Converte custos respeitando o formato BR e ajusta escala
        if "Custo_Produto" in df_custos.columns:
            def corrigir_valor(v):
                v = str(v).strip()
                if v in ["", "-", "nan", "N/A", "None", "0", "0,00", "0.00"]:
                    return 0.0

                v = v.replace("R$", "").replace(" ", "")
                # Detecta o padrão de separadores
                if "," in v and "." in v:
                    # Ex: 1.234,56 → 1234.56
                    v = v.replace(".", "").replace(",", ".")
                elif "," in v and "." not in v:
                    # Ex: 162,49 → 162.49
                    v = v.replace(",", ".")
                
                try:
                    val = float(v)
                    # Adiciona uma heurística para valores absurdos (erro de escala)
                    if val > 999 and "." in v and v.split('.')[-1] not in ["00", "0", ""]:
                         val = val / 100
                    elif val > 9999 and not ('.' in v or ',' in v):
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
        # Converte para strings para garantir que o Sheets entenda.
        df_str = df.astype(str) 
        sheet.update([df_str.columns.values.tolist()] + df_str.values.tolist())
        st.success(f"💾 Custos salvos no Google Sheets em {(datetime.utcnow() - timedelta(hours=3)).strftime('%d/%m/%Y %H:%M')}")
    except Exception as e:
        st.error(f"Erro ao salvar custos no Google Sheets: {e}")

# --- Bloco Visual de Custos ---
st.markdown("---")
custo_df = carregar_custos_google()

if not custo_df.empty:
    custo_df["SKU"] = custo_df["SKU"].astype(str).str.replace(r"[^\d\-]", "", regex=True)
else:
    st.warning("⚠️ Nenhum custo encontrado. Você pode adicionar manualmente abaixo.")

# Garante que o DataFrame a ser editado tenha a coluna SKU e Custo_Produto (limpas para edição)
if "SKU" not in custo_df.columns: custo_df["SKU"] = ""
if "Custo_Produto" not in custo_df.columns: custo_df["Custo_Produto"] = 0.0
if "Produto" not in custo_df.columns: custo_df["Produto"] = ""

custos_editados = st.data_editor(
    custo_df[["SKU", "Produto", "Custo_Produto"]], 
    num_rows="dynamic", 
    use_container_width=True,
    column_config={
        "Custo_Produto": st.column_config.NumberColumn(
            "Custo do Produto (R$)", format="R$ %.2f", min_value=0.0
        )
    }
)

if st.button("💾 Atualizar custos no Google Sheets"):
    # Limpa os dados de edição antes de salvar
    custos_para_salvar = custos_editados.copy()
    custos_para_salvar["SKU"] = custos_para_salvar["SKU"].astype(str).str.replace(r"[^\d\-]", "", regex=True)
    salvar_custos_google(custos_para_salvar)

# ==============================================================================
# === 4. UPLOAD E PROCESSAMENTO DE VENDAS ===
# ==============================================================================
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
        # Tenta ler a aba correta com o cabeçalho correto
        df = pd.read_excel(uploaded_file, sheet_name="Vendas BR", header=5)
        df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)
        st.caption("Primeiras 20 linhas do arquivo carregado:")
        st.dataframe(df.head(20), use_container_width=True)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}. Verifique se a aba 'Vendas BR' e o cabeçalho na linha 6 estão corretos.")
        df = None # Define df como None se houver erro

# Botão para limpar o arquivo e forçar reload
if st.button("🗑️ Remover arquivo carregado"):
    st.session_state["uploaded_file"] = None
    st.cache_data.clear()
    st.rerun()

# ==============================================================================
# === 5. PROCESSAMENTO PRINCIPAL E CÁLCULOS ===
# ==============================================================================
if uploaded_file and df is not None:
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
        "N.º de venda": "Venda", "Data da venda": "Data", "Estado": "Estado", 
        "Receita por produtos (BRL)": "Valor_Venda", "Total (BRL)": "Valor_Recebido", 
        "Tarifa de venda e impostos (BRL)": "Tarifa_Venda", "Tarifas de envio (BRL)": "Tarifa_Envio", 
        "Cancelamentos e reembolsos (BRL)": "Cancelamentos", 
        "Preço unitário de venda do anúncio (BRL)": "Preco_Unitario",
        "SKU": "SKU", "# de anúncio": "Anuncio", "Título do anúncio": "Produto", 
        "Tipo de anúncio": "Tipo_Anuncio", "Receita por envio (BRL)": "Receita_Envio_ML" # Adicionado para uso posterior
    }

    df.rename(columns={c: col_map[c] for c in col_map if c in df.columns}, inplace=True)

    # === AJUSTE VENDA (ESSENCIAL PARA IDENTIFICAÇÃO DE PACOTES) ===
    def formatar_venda(valor):
        if pd.isna(valor):
            return ""
        return re.sub(r"[^\d]", "", str(valor))
    df["Venda"] = df["Venda"].apply(formatar_venda)
    
    # === AJUSTE SKU (ESSENCIAL PARA DADOS LIMPOS NOS ITENS FILHOS) ===
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

    # === REDISTRIBUI PACOTES (COM DETALHAMENTO DE TARIFAS E FRETE POR UNIDADE) ===
    # Funções de cálculo de tarifas do ML (estimativa)
    def calcular_custo_fixo(preco_unit):
        if preco_unit < 12.5: return round(preco_unit * 0.5, 2)
        elif preco_unit < 30: return 6.25
        elif preco_unit < 50: return 6.50
        elif preco_unit < 79: return 6.75
        else: return 0.0

    def calcular_percentual(tipo_anuncio):
        tipo = str(tipo_anuncio).strip().lower()
        if "premium" in tipo: return 0.17
        elif "clássico" in tipo or "classico" in tipo: return 0.12
        return 0.12 # Padrão para Clássico ou indefinido

    # Garante que todas as colunas necessárias existam
    for col in ["Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$", "Origem_Pacote", "Valor_Item_Total"]:
        if col not in df.columns:
            df[col] = None
    
    # Prepara Tarifa_Envio e Valor_Recebido para uso no loop (pode estar como None)
    if "Tarifa_Envio" not in df.columns: df["Tarifa_Envio"] = 0.0
    if "Valor_Recebido" not in df.columns: df["Valor_Recebido"] = 0.0

    # === PROCESSA PACOTES AGRUPADOS ===
    for i, row in df.iterrows():
        estado = str(row.get("Estado", ""))
        match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
        if not match:
            df.loc[i, "Origem_Pacote"] = None # Linhas fora do pacote
            continue

        qtd = int(match.group(1))

        if i + 1 + qtd > len(df):
             # st.warning(f"⚠️ Aviso: Pacote da venda {row.get('Venda', 'N/A')} na linha {i+6} está incompleto no final do arquivo e foi ignorado.")
             continue

        subset = df.iloc[i + 1 : i + 1 + qtd].copy()
        if subset.empty:
            continue
            
        venda_pai_id = str(row['Venda']).strip()
        if not venda_pai_id:
             continue

        total_recebido = float(row.get("Valor_Recebido", 0) or 0)
        frete_total = float(row.get("Tarifa_Envio", 0) or 0)

        col_preco_unitario = "Preco_Unitario" if "Preco_Unitario" in subset.columns else None
        if col_preco_unitario is None: continue # Não processa se não tiver preço unitário

        subset["Preco_Unitario_Item"] = pd.to_numeric(subset[col_preco_unitario], errors="coerce").fillna(0)
        
        # Soma dos preços unitários dos itens do pacote para proporção de Valor_Recebido
        soma_precos = subset["Preco_Unitario_Item"].sum() 
        # Soma das unidades para proporção de frete
        total_unidades = subset[coluna_unidades].sum() or 1 

        total_tarifas_calc = total_recebido_calc = total_frete_calc = 0

        # Loop para itens filhos
        for j in subset.index:
            preco_unit = float(subset.loc[j, "Preco_Unitario_Item"] or 0)
            tipo_anuncio = subset.loc[j, "Tipo_Anuncio"]
            perc = calcular_percentual(tipo_anuncio)
            custo_fixo = calcular_custo_fixo(preco_unit)

            unidades_item = subset.loc[j, coluna_unidades]

            valor_item_total = preco_unit * unidades_item
            tarifa_total = round(valor_item_total * perc + (custo_fixo * unidades_item), 2)

            # Distribuição da Receita Baseada no Valor do Item em relação à Soma dos Preços Unitários
            proporcao_venda = (preco_unit / soma_precos) if soma_precos else 0
            valor_recebido_item = round(total_recebido * proporcao_venda, 2)
            
            # Distribui frete baseado na proporção de unidades do item
            proporcao_unidades = unidades_item / total_unidades if total_unidades else 0
            frete_item = round(frete_total * proporcao_unidades, 2)

            # atualiza DataFrame na linha do item (j)
            df.loc[j, "Valor_Venda"] = valor_item_total 
            df.loc[j, "Valor_Recebido"] = valor_recebido_item
            df.loc[j, "Tarifa_Venda"] = tarifa_total
            df.loc[j, "Tarifa_Envio"] = frete_item
            df.loc[j, "Tarifa_Percentual_%"] = perc * 100
            df.loc[j, "Tarifa_Fixa_R$"] = custo_fixo * unidades_item
            df.loc[j, "Tarifa_Total_R$"] = tarifa_total
            df.loc[j, "Origem_Pacote"] = f"{venda_pai_id}-PACOTE" 
            df.loc[j, "Valor_Item_Total"] = valor_item_total

            total_tarifas_calc += tarifa_total
            total_recebido_calc += valor_recebido_item
            total_frete_calc += frete_item

        # Atualiza a linha principal do pacote (i)
        df.loc[i, "Estado"] = f"{estado} (processado)"
        df.loc[i, "Tarifa_Venda"] = round(total_tarifas_calc, 2)
        df.loc[i, "Tarifa_Envio"] = round(frete_total, 2)
        df.loc[i, "Valor_Recebido"] = total_recebido
        df.loc[i, "Origem_Pacote"] = "PACOTE"
        
        # Zera métricas de lucro para a linha mãe
        df.loc[i, "Lucro_Real"] = 0
        df.loc[i, "Lucro_Liquido"] = 0
        df.loc[i, "Margem_Final_%"] = 0
        df.loc[i, "Markup_%"] = 0
        df.loc[i, "Margem_Liquida_%"] = 0
        df.loc[i, "Custo_Produto_Total"] = 0

    # === VALIDAÇÃO DOS PACOTES ===
    df["Tarifa_Validada_ML"] = ""
    mask_pacotes_filhos = df["Origem_Pacote"].apply(lambda x: isinstance(x, str) and x.endswith("-PACOTE"))
    
    for pacote in df.loc[mask_pacotes_filhos, "Origem_Pacote"].unique():
        if not isinstance(pacote, str): continue
            
        if pacote.endswith("-PACOTE"):
            filhos = df[df["Origem_Pacote"] == pacote]
            
            venda_pai_id = pacote.split("-")[0]
            # Usa .iloc[0] para pegar a primeira (e única) linha pai
            pai = df[df["Venda"].astype(str).eq(venda_pai_id)]
            
            if not pai.empty:
                soma_filhas = filhos["Tarifa_Venda"].sum() + filhos["Tarifa_Envio"].sum()
                tarifa_pai = pai["Tarifa_Venda"].iloc[0].sum() + pai["Tarifa_Envio"].iloc[0].sum()
                
                # Aplica o resultado da validação nas linhas filhas
                df.loc[df["Origem_Pacote"] == pacote, "Tarifa_Validada_ML"] = "✔️" if abs(soma_filhas - tarifa_pai) < 1 else "❌"

    # === CONVERSÕES FINAIS ===
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario", "Receita_Envio_ML"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).abs()
            
    # === COMPLETA DADOS DE PACOTES COM SKUs E TÍTULOS AGRUPADOS ===
    for i, row in df.iterrows():
        if row.get("Origem_Pacote") == "PACOTE":
            venda_pai_id = str(row['Venda']).strip()
            subset_mask = df["Origem_Pacote"] == f"{venda_pai_id}-PACOTE"
            subset = df[subset_mask].copy()

            if subset.empty:
                continue

            # Concatena SKUs e títulos dos filhos
            skus = subset["SKU"].astype(str).replace("nan", "").unique().tolist()
            produtos = subset["Produto"].astype(str).replace("nan", "").unique().tolist()

            skus_formatados = [s for s in skus if s and s != "0"]
            sku_concat = "-".join(skus_formatados)

            if len(produtos) > 2:
                produto_concat = f"{produtos[0]} + {len(produtos)-1} outros"
            else:
                produto_concat = " + ".join([p for p in produtos if p])

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

    # === DATA ===
    df["Data"] = df["Data"].astype(str).str.replace(r"(hs\.?|às)", "", regex=True).str.strip()
    meses_pt = {
        "janeiro": "01", "fevereiro": "02", "março": "03", "abril": "04",
        "maio": "05", "junho": "06", "julho": "07", "agosto": "08",
        "setembro": "09", "outubro": "10", "novembro": "11", "dezembro": "12"
    }

    def parse_data_portugues(texto):
        if not isinstance(texto, str) or not any(m in texto.lower() for m in meses_pt):
            # Tenta conversão direta para lidar com formatos datetime já limpos
            try:
                 return pd.to_datetime(texto)
            except:
                 return None

        try:
            texto = texto.lower().replace(",", "")
            partes = texto.split(" de ")
            if len(partes) < 3: return None
            
            dia = partes[0].split(" ")[-1].zfill(2) # Garante que pegue o dia, mesmo com texto antes
            mes = meses_pt.get(partes[1].strip(), "01")
            ano_e_hora = partes[2].split(" ")
            ano = ano_e_hora[0]
            hora = " ".join(ano_e_hora[1:]).strip() if len(ano_e_hora) > 1 else "00:00"
            
            # Limpa a hora, caso contenha "h" ou "minutos"
            hora = re.sub(r"[^\d:]", "", hora)
            if len(hora.split(':')) == 1: hora = hora + ":00" # Adiciona minutos se só tiver hora
                
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
            • Custos de envio (Tarifas de Envio).<br>
            • Custo fixo de embalagem e custo fiscal configurável.<br>
            • Quantidade total de unidades por venda.<br><br>
            Lucro Real = Valor da venda + Receita Envio ML − Tarifas ML − Frete − Embalagem − Custo fiscal.<br>
            </div>
            """,
            unsafe_allow_html=True,
        )

    df["Data"] = df["Data"].dt.strftime("%d/%m/%Y %H:%M")

    # === AUDITORIA ===
    # Verificação: (Valor_Venda + Receita_Envio_ML) - (Tarifa_Venda + Tarifa_Envio + Cancelamentos)
    df["Verificacao_Cancelamento"] = df["Valor_Venda"] + df["Receita_Envio_ML"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"] + df["Cancelamentos"])
    
    # Cancelamento é considerado correto se o valor recebido for próximo de zero E o cálculo de tarifas bater com a Venda (indicando que foi 100% estornado)
    df["Cancelamento_Correto"] = (df["Valor_Recebido"] <= 0.1) & (abs(df["Verificacao_Cancelamento"]) <= 0.1)
    
    df["Diferença_R$"] = df["Valor_Venda"] - df["Valor_Recebido"]
    
    # Adiciona tratamento de divisão por zero
    df["%Diferença"] = ((1 - (df["Valor_Recebido"] / df["Valor_Venda"].replace(0, np.nan))) * 100).round(2).fillna(0)
    
    df["Status"] = df.apply(
        lambda x: "🟦 Cancelamento Correto" if x["Cancelamento_Correto"]
        else "⚠️ Acima da Margem" if x["%Diferença"] > margem_limite
        else "✅ Normal", axis=1
    )

    # === FINANCEIRO (PRIMEIRA FASE: SEM CUSTO DO PRODUTO) ===
    df["Custo_Embalagem"] = custo_embalagem * df[coluna_unidades]
    df["Custo_Fiscal"] = (df["Valor_Venda"] * (custo_fiscal / 100)).round(2)

    # Lucro Bruto (Receita Total - Custos ML)
    df["Lucro_Bruto"] = (
        df["Valor_Venda"] + df["Receita_Envio_ML"] - (df["Tarifa_Venda"] + df["Tarifa_Envio"])
    ).round(2)

    # Lucro Real (Lucro Bruto - Fiscal - Embalagem)
    df["Lucro_Real"] = (
        df["Lucro_Bruto"] - (df["Custo_Embalagem"] + df["Custo_Fiscal"])
    ).round(2)

    # === PLANILHA DE CUSTOS (SEGUNDA FASE: COM CUSTO DO PRODUTO) ===
    if not custos_editados.empty and "Custo_Produto" in custos_editados.columns:
        try:
            # Filtra e limpa a tabela de custos editada pelo usuário
            custo_df_final = custos_editados[["SKU", "Custo_Produto"]].copy()
            custo_df_final["SKU"] = custo_df_final["SKU"].astype(str).str.strip()

            # Remove colunas de custo temporárias do merge anterior, se houver
            if "Custo_Produto" in df.columns: df.drop(columns=["Custo_Produto"], inplace=True)
            
            df = df.merge(custo_df_final, on="SKU", how="left")
            
            # Custo_Produto é o custo unitário do merge. Custo_Produto_Total é o custo total da venda.
            df["Custo_Produto_Total"] = df["Custo_Produto"].fillna(0) * df[coluna_unidades]

            # --- Lucro e Margens completas ---
            # Lucro Líquido = Lucro Real - Custo do Produto Total
            df["Lucro_Liquido"] = (df["Lucro_Real"] - df["Custo_Produto_Total"]).round(2)

            # Margem Final = Lucro Líquido / Valor da Venda
            df["Margem_Final_%"] = (
                (df["Lucro_Liquido"] / df["Valor_Venda"].replace(0, np.nan)) * 100
            ).round(2)

            # Markup = Lucro Líquido / Custo do Produto Total
            df["Markup_%"] = (
                (df["Lucro_Liquido"] / df["Custo_Produto_Total"].replace(0, np.nan)) * 100
            ).round(2)

            custo_carregado = True
            st.success("✅ Custos aplicados com sucesso para cálculo de Lucro Líquido e Margens Finais.")

        except Exception as e:
            st.error(f"Erro ao aplicar custos: {e}")
            
    # Garante que as colunas existam para o bloco de métricas, caso o merge de custo falhe
    if "Margem_Final_%" not in df.columns:
        df["Margem_Final_%"] = np.nan
    if "Lucro_Liquido" not in df.columns:
        df["Lucro_Liquido"] = df["Lucro_Real"].copy()
    if "Custo_Produto_Total" not in df.columns:
        df["Custo_Produto_Total"] = 0.0
    
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
            "Custo_Produto_Total", "%Diferença", "Diferença_R$"
        ]
        for campo in campos_financeiros:
            if campo in df.columns:
                df.loc[mask_pacotes, campo] = 0.0
        df.loc[mask_pacotes, "Status"] = "🔹 Pacote Agrupado (Somente Controle)"
        df.loc[mask_pacotes, coluna_unidades] = 0.0 # Zera unidades na linha mãe
        df.loc[mask_pacotes, "Valor_Venda"] = 0.0 # Zera Valor_Venda na linha mãe (já foi redistribuído)


    # === EXCLUI CANCELAMENTOS DO CÁLCULO ===
    df_validas = df[df["Status"].isin(["✅ Normal", "⚠️ Acima da Margem"])].copy()

    # === MÉTRICAS FINAIS (CÁLCULO) ===
    receita_total = df_validas["Valor_Venda"].sum()
    total_vendas = df[df["Estado"].str.contains("Pacote", case=False, na=False) == False].shape[0] # Conta todas as linhas exceto as mães de pacote
    
    if custo_carregado:
        lucro_total = df_validas["Lucro_Liquido"].sum()
        # Prejuízo: soma dos valores absolutos dos lucros líquidos negativos
        prejuizo_total = abs(df_validas.loc[df_validas["Lucro_Liquido"] < 0, "Lucro_Liquido"].sum())
        margem_media = df_validas["Margem_Final_%"].replace([np.inf, -np.inf], np.nan).mean()
    else:
        lucro_total = df_validas["Lucro_Real"].sum()
        # Prejuízo: soma dos valores absolutos dos lucros reais negativos
        prejuizo_total = abs(df_validas.loc[df_validas["Lucro_Real"] < 0, "Lucro_Real"].sum())
        margem_media = df_validas["Margem_Liquida_%"].replace([np.inf, -np.inf], np.nan).mean()

    fora_margem = (df["Status"] == "⚠️ Acima da Margem").sum()
    cancelamentos = (df["Status"] == "🟦 Cancelamento Correto").sum()

# ==============================================================================
# === 6. MÉTRICAS FINAIS (EXIBIÇÃO) ===
# ==============================================================================
col1, col2, col3, col4, col5, col6 = st.columns(6)
# Função auxiliar para formatar em padrão BR
def format_br(value, is_currency=True, decimals=2):
    if pd.isna(value): return "-"
    format_str = f"{{:,.{decimals}f}}"
    formatted = format_str.format(value).replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {formatted}" if is_currency else f"{formatted}%"

col1.metric("Total de Vendas Processadas", total_vendas)
col2.metric("Fora da Margem (ML)", fora_margem)
col3.metric("Cancelamentos Corretos", cancelamentos)
col4.metric("Lucro Total Estimado", format_br(lucro_total))
col5.metric("Margem Média Estimada", format_br(margem_media, is_currency=False))
col6.metric("🔻 Prejuízo Total", format_br(prejuizo_total))

if uploaded_file and df is not None:
    # ==============================================================================
    # === 7. ANÁLISE POR TIPO DE ANÚNCIO ===
    # ==============================================================================
    st.markdown("---")
    st.subheader("📊 Análise por Tipo de Anúncio (Clássico x Premium)")

    if "Tipo_Anuncio" in df.columns:
        # Corrige campos vazios e preenche pacotes
        df_tipo = df.copy()
        df_tipo["Tipo_Anuncio"] = (
            df_tipo["Tipo_Anuncio"]
            .astype(str)
            .str.strip()
            .replace(["nan", "None", ""], "Agrupado (Pacotes)")
        )
        # Exclui linhas mães de pacote da contagem
        df_tipo = df_tipo[df_tipo["Origem_Pacote"] != "PACOTE"]

        tipo_counts = df_tipo["Tipo_Anuncio"].value_counts(dropna=False).reset_index()
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

    # ==============================================================================
    # === 8. ANÁLISE ANALÍTICA DE MARGEM POR ITEM DE PACOTE (COMPLETADO AQUI) ===
    # ==============================================================================
    st.markdown("---")
    st.subheader("📦 Margem Analítica por Item de Pacote")

    mask_filhos = df["Origem_Pacote"].apply(lambda x: isinstance(x, str) and "-PACOTE" in x)
    df_pacotes_itens = df[mask_filhos].copy()

    if custo_carregado and not df_pacotes_itens.empty:
        analitico = []

        # Itera por cada pacote identificado (ex: "2000009496621409-PACOTE")
        for pacote_id in df_pacotes_itens["Origem_Pacote"].unique():
            # Filtra os filhos para o cálculo
            grupo = df[df["Origem_Pacote"] == pacote_id]
            if grupo.empty: continue

            # Encontra a linha principal (pai)
            venda_pai = pacote_id.replace("-PACOTE", "")
            linha_pai = df[df["Venda"].astype(str) == venda_pai]
            if linha_pai.empty: continue

            # Valores totais da venda principal (que serão rateados)
            total_venda_pai = float(linha_pai["Valor_Venda"].iloc[0] or 0)
            total_frete = float(linha_pai["Tarifa_Envio"].iloc[0] or 0)
            total_tarifa = float(linha_pai["Tarifa_Venda"].iloc[0] or 0)
            total_custofiscal = float(linha_pai.get("Custo_Fiscal", pd.Series([0.0])).iloc[0] or 0)
            total_embalagem = float(linha_pai.get("Custo_Embalagem", pd.Series([0.0])).iloc[0] or 0)

            # Soma de Valores_Venda (dos itens, que é o Preço Unitário * Unidades)
            soma_valores_itens = grupo["Valor_Venda"].sum()
            if soma_valores_itens == 0: continue

            num_itens_distintos = len(grupo) # Número de linhas filhas (itens distintos)

            for _, item in grupo.iterrows():
                # Proporção da participação no valor total do pacote
                proporcao = item["Valor_Venda"] / soma_valores_itens
                
                # Rateio dos custos totais do PACOTE (linha pai)
                tarifa_prop = round(total_tarifa * proporcao, 2)
                frete_prop = round(total_frete * proporcao, 2)
                fiscal_prop = round(total_custofiscal * proporcao, 2)
                embalagem_prop = round(total_embalagem * proporcao, 2) # Rateado por proporção de valor

                # Custo do produto (já embutido na linha filha)
                custo_prod = float(item.get("Custo_Produto_Total", 0) or 0) 
                
                lucro_liquido = (
                    item["Valor_Venda"] - tarifa_prop - frete_prop - fiscal_prop - embalagem_prop - custo_prod
                )
                margem_item = round((lucro_liquido / item["Valor_Venda"]) * 100, 2) if item["Valor_Venda"] > 0 else 0

                analitico.append({
                    "Pacote": pacote_id,
                    "Venda_Pai": venda_pai,
                    "Produto": item["Produto"],
                    "SKU": item["SKU"],
                    "Unidades": item[coluna_unidades],
                    "Valor_Venda_Item": item["Valor_Venda"],
                    "Tarifa_ML_Prop": tarifa_prop,
                    "Frete_Prop": frete_prop,
                    "Fiscal_Prop": fiscal_prop,
                    "Embalagem_Prop": embalagem_prop,
                    "Custo_Produto_Total": round(custo_prod, 2),
                    "Lucro_Liquido_Item": round(lucro_liquido, 2),
                    "Margem_Item_%": margem_item
                })

        df_analitico = pd.DataFrame(analitico)
        
        # Formata o DataFrame para exibição (apenas R$ e %)
        df_analitico_display = df_analitico.copy()
        for col in ["Valor_Venda_Item", "Tarifa_ML_Prop", "Frete_Prop", "Fiscal_Prop", "Embalagem_Prop", "Custo_Produto_Total", "Lucro_Liquido_Item"]:
             df_analitico_display[col] = df_analitico_display[col].apply(lambda x: format_br(x, is_currency=True))
        df_analitico_display["Margem_Item_%"] = df_analitico_display["Margem_Item_%"].apply(lambda x: format_br(x, is_currency=False))
        
        st.dataframe(df_analitico_display, use_container_width=True, hide_index=True)

        # Exporta o resumo analítico
        output_analitico = BytesIO()
        with pd.ExcelWriter(output_analitico, engine="xlsxwriter") as writer:
            df_analitico.to_excel(writer, index=False, sheet_name="Margem_Analitica_Pacotes")
        output_analitico.seek(0)
        st.download_button(
            label="⬇️ Exportar Margem Analítica por Item (Excel)",
            data=output_analitico,
            file_name=f"Margem_Analitica_Pacotes_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Nenhuma venda em pacote agrupado encontrada ou custos não foram aplicados.")

    # ==============================================================================
    # === 9. VENDAS ANORMAIS / FORA DA MARGEM ===
    # ==============================================================================
    st.markdown("---")
    st.subheader(f"🚨 Vendas Fora da Margem (Acima de {margem_limite}%)")

    # Filtra apenas itens que NÃO são pacotes e que estão fora da margem
    df_anormais = df[
        (df["Status"] == "⚠️ Acima da Margem") & 
        (df["Origem_Pacote"] != "PACOTE") & 
        (df["Status"] != "🟦 Cancelamento Correto")
    ].copy()

    if not df_anormais.empty:
        # Prepara a sugestão de ajuste
        colunas_anormais = [
            "Data", "Venda", "SKU", "Produto", coluna_unidades, 
            "Valor_Venda", "Valor_Recebido", "Diferença_R$", 
            "%Diferença", "Status", "Tarifa_Validada_ML"
        ]
        
        df_anormais_display = df_anormais[colunas_anormais].copy()
        for col in ["Valor_Venda", "Valor_Recebido", "Diferença_R$"]:
             df_anormais_display[col] = df_anormais_display[col].apply(lambda x: format_br(x, is_currency=True))
        df_anormais_display["%Diferença"] = df_anormais_display["%Diferença"].apply(lambda x: format_br(x, is_currency=False))
        
        st.dataframe(
            df_anormais_display, 
            use_container_width=True, 
            hide_index=True
        )

        # Exporta vendas anormais
        output_anormais = BytesIO()
        with pd.ExcelWriter(output_anormais, engine="xlsxwriter") as writer:
            df_anormais.to_excel(writer, index=False, sheet_name="Vendas_Anormais")
        output_anormais.seek(0)
        st.download_button(
            label="⬇️ Exportar Vendas Anormais (Excel)",
            data=output_anormais,
            file_name=f"Vendas_Anormais_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("🎉 Não há vendas acima do limite de margem definido!")

    # ==============================================================================
    # === 10. TABELA DETALHADA DE VENDAS ===
    # ==============================================================================
    st.markdown("---")
    st.subheader("📚 Tabela Detalhada de Vendas (Completa)")
    st.caption("Use os filtros e a busca na tabela abaixo para explorar os dados.")
    
    colunas_finais = [
        "Data", "Venda", "Estado", "Status", "Origem_Pacote", 
        "SKU", "Produto", coluna_unidades, 
        "Valor_Venda", "Valor_Recebido", "Diferença_R$", 
        "%Diferença", "Tarifa_Venda", "Tarifa_Envio", 
        "Cancelamentos", "Custo_Embalagem", "Custo_Fiscal", 
        "Custo_Produto_Total" if custo_carregado else None,
        "Lucro_Bruto", "Lucro_Real", 
        "Lucro_Liquido" if custo_carregado else "Lucro_Real (Sem Custo Prod)", 
        "Margem_Final_%" if custo_carregado else "Margem_Liquida_% (Sem Custo Prod)", 
        "Markup_%" if custo_carregado else None,
        "Tarifa_Validada_ML"
    ]
    
    # Remove colunas None
    colunas_finais = [c for c in colunas_finais if c is not None and c in df.columns]

    df_display = df[colunas_finais].copy()

    # Formata colunas de valores e porcentagem para exibição
    for col in ["Valor_Venda", "Valor_Recebido", "Diferença_R$", "Tarifa_Venda", "Tarifa_Envio", 
                "Cancelamentos", "Custo_Embalagem", "Custo_Fiscal", "Lucro_Bruto", "Lucro_Real"]:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(lambda x: format_br(x, is_currency=True))

    for col in ["%Diferença", "Margem_Final_%", "Margem_Liquida_%", "Markup_%"]:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(lambda x: format_br(x, is_currency=False))
            
    if "Custo_Produto_Total" in df_display.columns:
        df_display["Custo_Produto_Total"] = df_display["Custo_Produto_Total"].apply(lambda x: format_br(x, is_currency=True))
    
    if "Lucro_Liquido" in df.columns:
        col_lucro_liquido = "Lucro_Liquido" if custo_carregado else "Lucro_Real (Sem Custo Prod)"
        df_display[col_lucro_liquido] = df[col_lucro_liquido].apply(lambda x: format_br(x, is_currency=True))


    # Usa o st.data_editor para permitir a filtragem nativa
    st.data_editor(
        df_display,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
    )

    # Exporta o DataFrame final
    output_total = BytesIO()
    with pd.ExcelWriter(output_total, engine="xlsxwriter") as writer:
        # Usa o df ORIGINAL (sem formatação de string) para o Excel
        df.to_excel(writer, index=False, sheet_name="Auditoria_Completa")
    output_total.seek(0)

    st.download_button(
        label="⬇️ Exportar Tabela Completa (Excel)",
        data=output_total,
        file_name=f"Auditoria_ML_Completa_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
