# 🛡️ Painel de Inteligência e Apoio ao Controle Interno — PNCP

Sistema analítico desenvolvido em **Python (Streamlit, Pandas e Scikit-Learn)** para o monitoramento preventivo, triagem de riscos e auditoria de contratações públicas municipais, consumindo dados abertos da API do **PNCP (Portal Nacional de Contratações Públicas)**.

---

## 🚀 Principais Funcionalidades
- **Consulta Dinâmica e Multithread:** Extração de contratos, atas de registro de preços e editais/compras diretamente do PNCP com alta performance e tratamento automático de falhas (*Retry / Exponential Backoff*).
- **Motor de Classificação de Risco (Rule-Based + Statistical):** Avaliação automatizada de anomalias com base em desvios estatísticos (IQR), valores limiar, completude cadastral e concentração de fornecedores.
- **Clusterização Semântica (NLP):** Uso de **TF-IDF** e **Similaridade de Cosseno** para encontrar contratações historicamente semelhantes na mesma amostra, auxiliando na verificação de preços e padrões de objetos.
- **Segmentação Estratégica:** Filtros globais por modalidade de contratação (Pregão Eletrônico, Dispensa, Inexigibilidade, etc.).
- **Exportação de Relatórios:** Geração de planilhas formatadas em Excel (`.xlsx`) com abas separadas para a Matriz de Risco e Dados Brutos.

---

## ⚙️ Instalação e Execução Local

1. **Clone o repositório ou salve o script principal como `app.py`**.
2. **Instale as dependências necessárias**:
   ```bash
   pip install streamlit pandas requests urllib3 openpyxl scikit-learn pyarrow
