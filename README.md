# 🛡️ Painel de Inteligência e Apoio ao Controle Interno - PNCP

Aplicação web desenvolvida em **Streamlit** projetada para auditoria, triagem de riscos, monitoramento de vigências, busca ativa de aditivos, detecção de possível fracionamento de despesa e análise semântica de contratações públicas integrada à API oficial do **PNCP** (Portal Nacional de Contratações Públicas). Configurado por padrão para a **Prefeitura Municipal de Rio das Pedras/SP** (`CNPJ 44.826.840/0001-83`), permitindo também consultar qualquer outro órgão público.

---

## ⚡ Principais Melhorias de Desempenho e Arquitetura (v7.0)

1. **Motor de Risco em Complexidade Linear $O(N)$ (`preparar_estatisticas_contexto`):**
   - **Antes:** `calcular_risco` recalculava quantis (`Q25`, `Q50`, `Q75`, `Q90`), limites de `IQR` e contagens de `CNPJ` para todo o DataFrame a cada linha processada ($O(N^2)$).
   - **Agora:** Todas as métricas estatísticas e mapas de frequência de fornecedores são pré-computados uma única vez por filtragem ($O(N)$), tornando a renderização da matriz de risco quase instantânea.
2. **Varredura Paralela de Editais e Modalidades (`consultar_editais`):**
   - **Antes:** As 15 modalidades do PNCP eram consultadas em loop sequencial bloqueante.
   - **Agora:** As modalidades são consultadas simultaneamente via `ThreadPoolExecutor` com pool de conexões HTTP dimensionado (`pool_maxsize = MAX_WORKERS * 4`), reduzindo o tempo de espera em até **8x**.
3. **Cache de NLP/TF-IDF e Pré-processamento Léxico com `str.translate`:**
   - **Antes:** A função `calcular_modelo_tfidf` estava sem cache e reconstruía a matriz TF-IDF do zero a cada troca de contrato na Aba de Auditoria, usando 6 expressões regulares seguidas para remover acentos.
   - **Agora:** A remoção de acentos utiliza tabela de tradução em memória (`TABELA_SEM_ACENTOS`) em $O(1)$ por caractere e a matriz TF-IDF é cacheada via `@st.cache_data`.
4. **Indexação em Memória para Pesquisa Livre Instantânea:**
   - Uma chave pré-normalizada (`_busca`) é montada na carga dos registros (incluindo Objeto, Fornecedor, CNPJ, Número, Processo e Controle PNCP), evitando concatenações repetidas a cada tecla digitada.
5. **Novas Capacidades de Auditoria (Lei nº 14.133/2021):**
   - **Aba 3 — Inteligência de Fornecedores e Triagem de Fracionamento:** Ranking consolidado por CNPJ/Fornecedor (Volume Acumulado, % do Volume, Score Médio e Contratos de Alto Risco) e detector semântico de contratações com objetos similares (apoio à verificação do art. 75, § 1º, da Lei 14.133/2021).
   - **Exportação de Dossiê Individual em Word (`.docx`):** Além do relatório executivo geral, o auditor pode baixar o dossiê completo de uma contratação específica com um clique.
   - **Excel Multicamadas:** Planilha `.xlsx` agora exporta 4 abas formatadas (`Matriz de Risco`, `Prazos e Vigências`, `Ranking Fornecedores` e `Dados Brutos`) com hiperlinks ativos.

---

## 🛠️ Tecnologias e Dependências

* **Python 3.9+**
* [`Streamlit`](https://streamlit.io/) — Interface executiva e reativa
* [`Pandas`](https://pandas.pydata.org/) — Manipulação e normalização tabular
* [`Requests`](https://requests.readthedocs.io/) & [`Urllib3`](https://urllib3.readthedocs.io/) — Pool de conexões HTTP com *retry* exponencial
* [`Scikit-Learn`](https://scikit-learn.org/) — Vetorização TF-IDF e similaridade de cosseno (opcional)
* [`OpenPyXL`](https://openpyxl.readthedocs.io/) — Exportação de planilhas Excel multicamadas
* [`python-docx`](https://python-docx.readthedocs.io/) — Geração de relatórios executivos e dossiês em Word

---

## 📦 Instalação e Execução

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

2. Execute o painel:
   ```bash
   streamlit run app.py
   ```
