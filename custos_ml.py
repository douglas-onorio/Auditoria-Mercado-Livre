import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

st.set_page_config(page_title="📦 Custos ML", layout="wide")
st.title("💰 Gerenciador de Custos Mercado Livre")

# === AUTENTICAÇÃO GOOGLE SHEETS ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# ✅ Usa as credenciais armazenadas em SECRETS do Streamlit Cloud
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
client = gspread.authorize(creds)

# === ABRIR PLANILHA ===
SHEET_NAME = "CUSTOS_ML"  # nome exato da planilha no Google Sheets
sheet = client.open(SHEET_NAME).sheet1   # ✅ estava faltando aspas aqui
dados = sheet.get_all_records()
df = pd.DataFrame(dados)

st.info("✅ Conectado à planilha de custos do Google Sheets.")

# === MOSTRAR E PERMITIR EDIÇÃO ===
st.subheader("📋 Editar Custos")
edit_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

# === BOTÃO SALVAR ===
if st.button("💾 Salvar alterações"):
    sheet.clear()
    sheet.update([edit_df.columns.values.tolist()] + edit_df.values.tolist())
    st.success(f"Alterações salvas com sucesso em {datetime.now().strftime('%d/%m/%Y %H:%M')}!")
