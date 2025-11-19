# 📊 Auditoria Financeira Mercado Livre

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://auditoria-mercadolivre.streamlit.app/)

## 🚀 Sobre o Projeto
Esta ferramenta foi desenvolvida para solucionar um problema crítico de vendedores do Mercado Livre: a **conferência financeira de vendas em lote e pacotes**. 

Diferente de planilhas manuais, este sistema processa o relatório de vendas (`.xlsx`), cruza com custos de produtos (integrado via **Google Sheets API**) e audita automaticamente se as taxas cobradas, custos de envio e impostos estão dentro da margem de lucro esperada.

## ✨ Funcionalidades Poderosas

* **Auditoria de "Pacotes" (Bundles):** Algoritmo inteligente que identifica vendas agrupadas ("Pacote de X produtos"), realiza o rateio proporcional de descontos, fretes e taxas entre os itens e valida se a cobrança do Mercado Livre está correta.
* **Integração com Google Sheets:** Busca e atualiza a base de custos dos produtos em tempo real, sem necessidade de re-upload de planilhas de custo.
* **Cálculo de Lucro Real:** Considera comissões (Clássico/Premium), Tarifa Fixa, Frete, Impostos (Simples Nacional) e Custo de Embalagem.
* **Exportação Avançada (XlsxWriter):** Gera um relatório Excel final não apenas com valores estáticos, mas com **fórmulas ativas** e formatação condicional (cores), facilitando a análise posterior pelo time financeiro.
* **Alertas Automáticos:** Identifica visualmente vendas que ficaram abaixo da margem mínima estipulada ou com prejuízo.

## 🛠 Tecnologias Utilizadas

* **Python 3.9+**
* **Streamlit:** Interface web interativa.
* **Pandas & NumPy:** Processamento de dados e cálculos financeiros.
* **Gspread (Google API):** Conexão com banco de dados de custos em nuvem.
* **XlsxWriter:** Engine para gerar Excels complexos com fórmulas e estilos.

## ⚙️ Como Rodar Localmente

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/douglas-onorio/Auditoria-Mercado-Livre.git](https://github.com/douglas-onorio/Auditoria-Mercado-Livre.git)
    ```
2.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Configure as Credenciais:**
    * É necessário configurar as credenciais do Google Cloud (`secrets.toml`) para a integração com o Sheets.
4.  **Execute a aplicação:**
    ```bash
    streamlit run app.py
    ```

---
**Desenvolvido por Douglas Onorio**
