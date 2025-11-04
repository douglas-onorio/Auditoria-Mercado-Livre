# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import re
import os
from pathlib import Path
import tempfile
import numpy as np # Adicionado: necessário para usar np.nan no cálculo das margens

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

# === GESTÃO DE CUSTOS (INTEGRAÇÃO GOOGLE SHEETS) ===
import gspread
from google.oauth2.service_account import Credentials
# import pandas as pd # Já importado
# from datetime import datetime # Já importado
# import streamlit as st # Já importado
import json

st.subheader("💰 Custos de Produtos (Google Sheets)")

try:
    # Escopos obrigatórios do Google Sheets e Drive
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    if "gcp_service_account" not in st.secrets:
        # Se estiver rodando localmente sem secrets, pode ser um problema.
        # Mas mantemos a lógica original de levantar a exceção.
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

# --- Garante que o client exista ---
if "client" not in locals() or client is None:
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

    # === AJUSTE VENDA (MOVIDO PARA CÁ: ESSENCIAL PARA IDENTIFICAÇÃO DE PACOTES) ===
    def formatar_venda(valor):
        if pd.isna(valor):
            return ""
        return re.sub(r"[^\d]", "", str(valor))
    df["Venda"] = df["Venda"].apply(formatar_venda)
    
    # === AJUSTE SKU (MOVIDO PARA CÁ: ESSENCIAL PARA DADOS LIMPOS NOS ITENS FILHOS) ===
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
    # import re # Já importado

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

    # Garante que todas as colunas necessárias existam
    for col in ["Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$", "Origem_Pacote", "Tarifa_Envio", "Valor_Item_Total"]:
        if col not in df.columns:
            df[col] = None

    # === PROCESSA PACOTES AGRUPADOS ===
    for i, row in df.iterrows():
        estado = str(row.get("Estado", ""))
        match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
        if not match:
            df.loc[i, "Origem_Pacote"] = None # Linhas fora do pacote
            continue

        qtd = int(match.group(1))

        # Verifica se o subset está dentro dos limites do DataFrame
        if i + 1 + qtd > len(df):
             st.warning(f"⚠️ Aviso: Pacote da venda {row.get('Venda', 'N/A')} na linha {i+6} está incompleto no final do arquivo e foi ignorado.")
             continue

        subset = df.iloc[i + 1 : i + 1 + qtd].copy()
        if subset.empty:
            continue
            
        # Garante que o ID da Venda Pai é válido antes de prosseguir
        venda_pai_id = str(row['Venda']).strip()
        if not venda_pai_id:
             st.warning(f"⚠️ Aviso: Pacote da venda na linha {i+6} tem N.º de venda inválido/vazio e foi ignorado.")
             continue

        total_venda = float(row.get("Valor_Venda", 0) or 0)
        total_recebido = float(row.get("Valor_Recebido", 0) or 0)

        # Captura o valor total do frete (Tarifa_Envio da linha principal)
        frete_total = float(row.get("Tarifa_Envio", 0) or 0)

        # Define a coluna de preço unitário
        col_preco_unitario = "Preco_Unitario" if "Preco_Unitario" in subset.columns else "Preço unitário de venda do anúncio (BRL)"
        subset["Preco_Unitario_Item"] = pd.to_numeric(subset[col_preco_unitario], errors="coerce").fillna(0)
        
        # Calcula a soma dos preços unitários dos itens do pacote para proporção
        soma_precos = subset["Preco_Unitario_Item"].sum()  
        # Calcula a soma das unidades para proporção de frete
        total_unidades = subset[coluna_unidades].sum() or 1  

        total_tarifas_calc = total_recebido_calc = total_frete_calc = 0

        # Loop corrigido: estava indentado incorretamente no código original
        for j in subset.index:
            preco_unit = float(subset.loc[j, "Preco_Unitario_Item"] or 0)
            tipo_anuncio = subset.loc[j, "Tipo_Anuncio"]
            perc = calcular_percentual(tipo_anuncio)
            custo_fixo = calcular_custo_fixo(preco_unit)

            # 🧮 quantidade comprada do item
            unidades_item = subset.loc[j, coluna_unidades]

            # 💰 calcula tarifa com base no valor total do item (preço unitário × unidades)
            valor_item_total = preco_unit * unidades_item
            tarifa_total = round(valor_item_total * perc + (custo_fixo * unidades_item), 2)

            # Proporção da receita recebida (Valor_Recebido) baseada no preço unitário
            proporcao_venda = (preco_unit / soma_precos) if soma_precos else 0
            valor_recebido_item = round(total_recebido * proporcao_venda, 2)
            
            # Distribui frete baseado na proporção de unidades do item
            proporcao_unidades = unidades_item / total_unidades if total_unidades else 0
            frete_item = round(frete_total * proporcao_unidades, 2)

            # atualiza DataFrame na linha do item (j)
            df.loc[j, "Valor_Venda"] = valor_item_total # Valor total do item (Preço unitário * Unidades)
            df.loc[j, "Valor_Recebido"] = valor_recebido_item
            df.loc[j, "Tarifa_Venda"] = tarifa_total
            df.loc[j, "Tarifa_Percentual_%"] = perc * 100
            df.loc[j, "Tarifa_Fixa_R$"] = custo_fixo * unidades_item
            df.loc[j, "Tarifa_Total_R$"] = tarifa_total
            df.loc[j, "Tarifa_Envio"] = frete_item
            df.loc[j, "Origem_Pacote"] = f"{venda_pai_id}-PACOTE" # ID de venda pai garantido como string e limpo
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

    # === VALIDAÇÃO DOS PACOTES ===
    df["Tarifa_Validada_ML"] = ""
    # Filtra apenas linhas que são strings e terminam com -PACOTE
    mask_pacotes_filhos = df["Origem_Pacote"].apply(lambda x: isinstance(x, str) and x.endswith("-PACOTE"))
    
    for pacote in df.loc[mask_pacotes_filhos, "Origem_Pacote"].unique():
        if not isinstance(pacote, str):
            continue
            
        # Garante que estamos pegando apenas os pacotes filhos
        if pacote.endswith("-PACOTE"):
            filhos = df[df["Origem_Pacote"] == pacote]
            
            # Tenta encontrar a linha pai (Venda original)
            venda_pai_id = pacote.split("-")[0]
            pai = df[df["Venda"].astype(str).eq(venda_pai_id)]
            
            if not pai.empty:
                soma_filhas = filhos["Tarifa_Venda"].sum() + filhos["Tarifa_Envio"].sum()
                tarifa_pai = pai["Tarifa_Venda"].sum() + pai["Tarifa_Envio"].sum()
                
                # Aplica o resultado da validação nas linhas filhas
                df.loc[df["Origem_Pacote"] == pacote, "Tarifa_Validada_ML"] = "✔️" if abs(soma_filhas - tarifa_pai) < 1 else "❌"

    # === CONVERSÕES ===
    for c in ["Valor_Venda", "Valor_Recebido", "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos", "Preco_Unitario"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).abs()

    # O bloco de limpeza de SKU e Venda foi movido para antes
    
    # === COMPLETA DADOS DE PACOTES COM SKUs E TÍTULOS AGRUPADOS ===
    for i, row in df.iterrows():
        estado = str(row.get("Estado", ""))
        match = re.search(r"Pacote de (\d+) produtos", estado, flags=re.IGNORECASE)
        if not match:
            continue

        qtd = int(match.group(1))

        # Garante que o subset esteja dentro dos limites novamente
        if i + 1 + qtd > len(df):
            continue

        subset = df.iloc[i + 1 : i + 1 + qtd].copy()
        if subset.empty:
            continue

        # Concatena SKUs e títulos dos filhos
        # OBS: Como o SKU foi limpo ANTES, ele já deve estar ok aqui.
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
    
    # Adiciona tratamento de divisão por zero
    df["%Diferença"] = ((1 - (df["Valor_Recebido"] / df["Valor_Venda"].replace(0, np.nan))) * 100).round(2).fillna(0)
    
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

    # === PLANILHA DE CUSTOS (SEGUNDO BLOCO DE CÁLCULO) ===
    custo_carregado = False
    if not custo_df.empty:
        try:
            custo_df["SKU"] = custo_df["SKU"].astype(str).str.strip()
            df = df.merge(custo_df[["SKU", "Custo_Produto"]], on="SKU", how="left")
            df["Custo_Produto_Total"] = df["Custo_Produto"].fillna(0) * df[coluna_unidades]

            # --- Custo Fiscal e Embalagem ---
            # O custo fiscal já foi calculado sobre Valor_Venda (total), mantendo assim.
            if "Custo_Fiscal" not in df.columns:
                 df["Custo_Fiscal"] = 0.0
            
            if "Custo_Embalagem" not in df.columns:
                 df["Custo_Embalagem"] = 0.0
            else:
                 df["Custo_Embalagem"] = pd.to_numeric(df["Custo_Embalagem"], errors="coerce").fillna(0)

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
            "Custo_Produto_Total"
        ]
        for campo in campos_financeiros:
            if campo in df.columns:
                df.loc[mask_pacotes, campo] = 0.0
        df.loc[mask_pacotes, "Status"] = "🔹 Pacote Agrupado (Somente Controle)"

    # === EXCLUI CANCELAMENTOS DO CÁLCULO ===
    df_validas = df[df["Status"] != "🟦 Cancelamento Correto"].copy() # Cria uma cópia para evitar SettingWithCopyWarning

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
    total_vendas = len(df)
    fora_margem = (df["Status"] == "⚠️ Acima da Margem").sum()
    cancelamentos = (df["Status"] == "🟦 Cancelamento Correto").sum()

# === MÉTRICAS FINAIS (EXIBIÇÃO) ===
# Este bloco usa as variáveis que agora estão inicializadas (0.0) ou calculadas
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
            .replace(["nan", "None", ""], "Agrupado (Pacotes)")
        )

        tipo_counts = df["Tipo_Anuncio"].value_counts(dropna=False).reset_index()
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

# === ANÁLISE ANALÍTICA DE MARGEM POR ITEM DE PACOTE ===
st.markdown("---")
st.subheader("📦 Margem Analítica por Item de Pacote")

if df is None or "Origem_Pacote" not in df.columns:
    st.info("Nenhum pacote com múltiplos produtos encontrado para análise detalhada.")
else:
    try:
        mask_filhos = df["Origem_Pacote"].apply(lambda x: isinstance(x, str) and "-PACOTE" in x)
        df_pacotes_itens = df[mask_filhos].copy()

        if df_pacotes_itens.empty:
            st.info("Nenhum pacote com múltiplos produtos encontrado para análise detalhada.")
        else:
            st.write(f"🔍 {len(df_pacotes_itens['Origem_Pacote'].unique())} pacotes identificados para análise...")
            analitico = []

            for idx, pacote_id in enumerate(df_pacotes_itens["Origem_Pacote"].unique(), start=1):
                st.caption(f"📦 Processando pacote {idx}/{len(df_pacotes_itens['Origem_Pacote'].unique())}: {pacote_id}")
                
                grupo = df[df["Origem_Pacote"] == pacote_id]
                if grupo.empty:
                    continue

                venda_pai = pacote_id.replace("-PACOTE", "")
                linha_pai = df[df["Venda"].astype(str) == venda_pai]
                if linha_pai.empty:
                    st.warning(f"⚠️ Linha pai não encontrada para venda {venda_pai}")
                    continue

                total_venda = float(linha_pai["Valor_Venda"].iloc[0] or 0)
                total_frete = float(linha_pai["Tarifa_Envio"].iloc[0] or 0)
                total_tarifa = float(linha_pai["Tarifa_Venda"].iloc[0] or 0)
                total_custofiscal = float(linha_pai.get("Custo_Fiscal", pd.Series([0.0])).iloc[0])
                total_embalagem = float(linha_pai.get("Custo_Embalagem", pd.Series([0.0])).iloc[0])

                soma_valores_itens = grupo["Valor_Venda"].sum()
                if soma_valores_itens <= 0:
                    st.warning(f"⚠️ Pacote {venda_pai} ignorado (Valor_Venda total zero)")
                    continue

                num_itens = max(len(grupo), 1)

                for _, item in grupo.iterrows():
                    valor_venda = float(item.get("Valor_Venda", 0) or 0)
                    proporcao = valor_venda / soma_valores_itens
                    tarifa_prop = round(total_tarifa * proporcao, 2)
                    frete_prop = round(total_frete * proporcao, 2)
                    fiscal_prop = round(total_custofiscal * proporcao, 2)
                    embalagem_prop = round(total_embalagem / num_itens, 2)

                    custo_prod = float(item.get("Custo_Produto", 0) or 0) * float(item.get(coluna_unidades, 1) or 1)
                    lucro_liquido = valor_venda - tarifa_prop - frete_prop - fiscal_prop - embalagem_prop - custo_prod
                    margem_item = round((lucro_liquido / valor_venda) * 100, 2) if valor_venda > 0 else 0

                    analitico.append({
                        "Pacote": pacote_id,
                        "Venda_Pai": venda_pai,
                        "Produto": item["Produto"],
                        "SKU": item["SKU"],
                        "Unidades": item[coluna_unidades],
                        "Valor_Venda_Item": valor_venda,
                        "Tarifa_Prop": tarifa_prop,
                        "Frete_Prop": frete_prop,
                        "Fiscal_Prop": fiscal_prop,
                        "Embalagem_Prop": embalagem_prop,
                        "Custo_Produto_Total": round(custo_prod, 2),
                        "Lucro_Liquido_Item": round(lucro_liquido, 2),
                        "Margem_Item_%": margem_item
                    })

            st.success("✅ Análise de pacotes concluída.")

            if analitico:
                df_analitico = pd.DataFrame(analitico)
                st.dataframe(df_analitico.head(20), use_container_width=True)
            else:
                st.info("Nenhum pacote processado com sucesso.")
                
    except Exception as e:
        st.error(f"❌ Erro ao processar análise de pacotes: {e}")
       
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
                "Produto", "Valor_Venda", "Tarifa_Venda", "Tarifa_Envio",
                "Custo_Embalagem", "Custo_Fiscal", "Lucro_Bruto", "Lucro_Real",
                coluna_unidades, "Margem_Liquida_%"
            ]
            filtro_display = filtro[[c for c in cols_to_display if c in filtro.columns]]
            st.write(filtro_display.dropna(axis=1, how="all"))

    # === VISUALIZAÇÃO DOS DADOS ANALISADOS ===
    st.markdown("---")
    st.subheader("📋 Itens Avaliados")

    colunas_vis = [
        "Venda", "Data", "Produto", "SKU", "Tipo_Anuncio",
        coluna_unidades, "Valor_Venda", "Valor_Recebido",
        "Tarifa_Venda", "Tarifa_Percentual_%", "Tarifa_Fixa_R$", "Tarifa_Total_R$",
        "Tarifa_Envio", "Cancelamentos",
        "Lucro_Real", "Margem_Liquida_%", "Status", "Origem_Pacote"
    ]

    # Filtra colunas que realmente existem no df
    cols_existentes = [c for c in colunas_vis if c in df.columns]

    st.dataframe(
        df[cols_existentes],
        use_container_width=True,
        height=450
    )

    # === EXPORTAÇÃO FINAL (colunas principais e financeiras) ===
    colunas_principais = [
        "Venda", "Data", "Produto", "SKU", "Tipo_Anuncio",
        coluna_unidades, "Valor_Venda", "Valor_Recebido",
        "Tarifa_Venda", "Tarifa_Envio", "Cancelamentos",
        "Custo_Fiscal", "Custo_Embalagem",
        "Custo_Produto_Total", "Lucro_Real", "Lucro_Liquido",
        "Margem_Liquida_%", "Margem_Final_%", "Markup_%",
        "Status", "Origem_Pacote", "Tarifa_Validada_ML"
    ]

    output_final = BytesIO()
    with pd.ExcelWriter(output_final, engine="xlsxwriter") as writer:
        df_export = df[[c for c in colunas_principais if c in df.columns]].copy()
        
        # Renomeia coluna de unidades para o export final
        if coluna_unidades != "Unidades":
             df_export.rename(columns={coluna_unidades: "Unidades"}, inplace=True)
             
        df_export.to_excel(writer, index=False, sheet_name="Auditoria_Completa")
    output_final.seek(0)

    st.download_button(
        label="⬇️ Exportar Auditoria Completa (Excel)",
        data=output_final,
        file_name=f"Auditoria_Vendas_ML_Completa_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
