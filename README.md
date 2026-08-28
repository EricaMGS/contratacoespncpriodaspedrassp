# Painel de Inteligência e Apoio ao Controle Interno - PNCP

Aplicação web desenvolvida em **Streamlit** projetada para auditoria, triagem de riscos e análise semântica de contratações públicas integrada à API do **PNCP** (Portal Nacional de Contratações Públicas). Foco específico no monitoramento municipal (configurado por padrão para Rio das Pedras/SP).

---

## 🚀 Funcionalidades Principais

* **Integração Multithread com o PNCP:** Consulta otimizada via `ThreadPoolExecutor` com retransmissão robusta (`urllib3` Retry Adapter) para contratos, atas de registro de preços e editais.
* **Motor Automático de Análise de Risco:**
* Classificação em níveis (**🔴 ALTO**, **🟡 MÉDIO**, **🟢 BAIXO**) baseada em completude cadastral, exceções legais (dispensas/inexigibilidades) e desvios estatísticos financeiros.
* Detecção de outliers financeiros utilizando a técnica de Amplitude Interquartil (**IQR**).
* Identificação de concentração de fornecedores no escopo.


* **Processamento de Linguagem Natural (NLP):** Clusterização semântica de objetos contratuais com **Scikit-Learn** (`TfidfVectorizer`) para encontrar licitações e contratos similares no mesmo segmento.
* **Segmentação Estratégica:** Filtros dinâmicos por modalidade de contratação com geração de dossiês de auditoria individualizados.
* **Exportação Avançada:** Relatórios formatados para Excel (`openpyxl`) contendo matrizes de risco e dados brutos formatados com congelamento de painéis e ajuste automático de colunas.

---

## 🛠️ Tecnologias e Dependências

* **Python 3.8+**
* [`Streamlit`](https://streamlit.io/) — Interface de usuário reativa
* [`Pandas`](https://pandas.pydata.org/) — Manipulação e normalização de dados tabulares
* [`Requests`](https://requests.readthedocs.io/) & [`Urllib3`](https://urllib3.readthedocs.io/) — Requisições HTTP tolerantes a falhas
* [`Scikit-Learn`](https://scikit-learn.org/) — Vetorização TF-IDF para análise de similaridade textual (opcional)
* [`Openpyxl`](https://openpyxl.readthedocs.io/) — Exportação de planilhas Excel estruturadas

---

## ⚙️ Configuração e Instalação

1. **Clone o repositório ou salve o script principal** (ex: `app.py`).
2. **Instale as dependências necessárias**:
```bash
pip install streamlit pandas requests urllib3 openpyxl scikit-learn

```


3. **Configure variáveis de ambiente opcionais (se necessário)**:
* `PNCP_MAX_WORKERS`: Define o limite de threads simultâneas para requisições (padrão: `6`).
* `PNCP_TIMEOUT_READ`: Tempo limite de leitura da API em segundos (padrão: `45`).



---

## 📦 Como Executar

Execute o comando abaixo no terminal na pasta onde o arquivo do sistema está salvo:

```bash
streamlit run app.py

```

---

## 🗂️ Estrutura do Código

* **Integração HTTP:** Sessão HTTP centralizada com políticas automáticas de recuperação de erros (`status_forcelist`).
* **Normalização:** Conversão automática de dados aninhados da API do PNCP em dataframes limpos.
* **Motor de Risco:** Atribuição ponderada de pontos com base em regras de controle interno e limites estatísticos (IQR).
* **Interface (Frontend):** Layout dividido em abas para visão geral da matriz de risco e detalhamento de dossiês individuais de auditoria.
