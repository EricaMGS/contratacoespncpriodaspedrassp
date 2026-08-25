import io
import datetime
import time
import re
import asyncio
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import httpx
import streamlit as st
import plotly.express as px

from docx import Document
from docx.shared import Pt, RGBColor
from fpdf import FPDF


# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Radar PNCP - Rio das Pedras/SP",
    page_icon="🏛️",
    layout="wide",
)


# ============================================================
# DADOS DO MUNICÍPIO E CONSTANTES
# ============================================================

CNPJ_RIO_DAS_PEDRAS = "44826840000183"
CODIGO_IBGE_RIO_DAS_PEDRAS = "3544004"
UF = "SP"

BASE_URL = "https://pncp.gov.br/api/consulta/v1"
NOME_MUNICIPIO = "Rio das Pedras/SP"
NOME_PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"

TAMANHO_PAGINA_CONTRATOS = 100
TAMANHO_PAGINA_ATAS = 100
TAMANHO_PAGINA_EDITAIS = 50

MAX_PAGINAS = 100
TIMEOUT_HTTP = 25.0

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}


# ============================================================
# FUNÇÕES GERAIS DE FORMATAÇÃO E PARSING
# ============================================================

def texto_valido(valor: Any, padrao: str = "N/D") -> str:
    """Converte qualquer tipo de dado em string tratada de forma segura."""
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass

    if isinstance(valor, dict):
        for chave in ("nome", "razaoSocial", "descricao", "valor", "nomeUnidade", "municipioNome"):
            if chave in valor:
                res = texto_valido(valor.get(chave), "")
                if res:
                    return res
        return str(valor)

    if isinstance(valor, list):
        if not valor:
            return padrao
        partes = [texto_valido(item, "") for item in valor if texto_valido(item, "")]
        return ", ".join(partes) if partes else padrao

    texto = str(valor).strip()
    if texto.lower() in {"", "none", "nan", "null", "n/d", "nd", "nat"}:
        return padrao
    return texto


def valor_numerico(valor: Any) -> Optional[float]:
    """Converte valores monetários do PNCP para float."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        try:
            return None if pd.isna(valor) else float(valor)
        except Exception:
            return None

    if isinstance(valor, dict):
        for chave in ("valor", "value", "valorTotal", "valorGlobal", "valorInicial"):
            if chave in valor:
                res = valor_numerico(valor[chave])
                if res is not None:
                    return res
        return None

    texto = str(valor).strip().replace("R$", "").replace("r$", "").replace(" ", "")
    if not texto:
        return None

    try:
        return float(texto)
    except Exception:
        pass

    try:
        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        return float(texto)
    except Exception:
        return None


def formatar_moeda_br(valor: Any) -> str:
    """Formata valor numérico para padrão monetário Real (R$)."""
    num = valor_numerico(valor)
    if num is None:
        return "N/D"
    return f"R$ {num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_data(valor: Any) -> str:
    """Converte ISO datas do PNCP para DD/MM/AAAA."""
    if valor is None:
        return "N/D"
    texto = str(valor).strip()
    if not texto or texto.lower() in {"none", "nan", "null", "n/d", "nat"}:
        return "N/D"
    try:
        data = pd.to_datetime(valor, errors="coerce")
        return texto if pd.isna(data) else data.strftime("%d/%m/%Y")
    except Exception:
        return texto


def normalizar_cnpj(valor: Any) -> str:
    """Remove pontuação e caracteres não numéricos de CPF/CNPJ."""
    return re.sub(r"\D", "", texto_valido(valor, ""))


def extrair_valor_recursivo(valor: Any, profundidade: int = 0) -> Any:
    """Extrai informações aninhadas em dicionários retornados pela API."""
    if profundidade > 4 or valor is None:
        return valor
    if isinstance(valor, dict):
        chaves = ["nome", "razaoSocial", "descricao", "descricaoObjeto", "valor", "numero", "cnpj", "numeroDocumento", "nomeUnidade", "municipioNome"]
        for chave in chaves:
            if chave in valor:
                res = extrair_valor_recursivo(valor[chave], profundidade + 1)
                if res not in (None, "", "N/D"):
                    return res
        return valor
    if isinstance(valor, list):
        res_lista = [str(extrair_valor_recursivo(item, profundidade + 1)) for item in valor if extrair_valor_recursivo(item, profundidade + 1) not in (None, "", "N/D")]
        return ", ".join(res_lista) if res_lista else None
    return valor


def obter_primeiro_valor(row: Any, campos: List[str], padrao: Any = "N/D") -> Any:
    """Busca o primeiro campo com dado válido dentro de múltiplos aliases."""
    if row is None:
        return padrao
    for campo in campos:
        try:
            if isinstance(row, pd.Series) and campo in row.index:
                val = row.get(campo)
            elif isinstance(row, dict) and campo in row:
                val = row.get(campo)
            else:
                continue
        except Exception:
            continue

        if val is None:
            continue
        try:
            if pd.isna(val):
                continue
        except Exception:
            pass

        val_limpo = extrair_valor_recursivo(val)
        if val_limpo is not None and texto_valido(val_limpo, ""):
            return val_limpo
    return padrao


# ============================================================
# EXTRAÇÃO ESTRUTURADA DE REGISTROS
# ============================================================

def obter_dados_registro(row: Any, tipo: str) -> Dict[str, Any]:
    id_pncp = obter_primeiro_valor(row, ["numeroControlePNCP", "numeroControlePNCPAta", "numeroControlePNCPCompra", "idContratoPNCP"])
    processo = obter_primeiro_valor(row, ["processo", "numeroProcesso", "processoAdministrativo", "numeroProcessoCompra"])
    objeto = obter_primeiro_valor(row, ["objetoCompra", "objetoContrato", "objetoAta", "objeto", "descricaoObjeto", "descricao"])
    fornecedor = obter_primeiro_valor(row, ["nomeRazaoSocialFornecedor", "razaoSocialFornecedor", "nomeFornecedor", "razaoSocial", "fornecedor"])
    cnpj_fornecedor = obter_primeiro_valor(row, ["niFornecedor", "cnpjFornecedor", "numeroDocumentoFornecedor", "cpfCnpjFornecedor", "ni"])
    situacao = obter_primeiro_valor(row, ["situacao", "situacaoCompra", "situacaoAta", "status"])
    orgao = obter_primeiro_valor(row, ["orgaoEntidade", "nomeOrgao", "razaoSocialOrgao", "orgao"])
    unidade = obter_primeiro_valor(row, ["unidadeOrgao", "nomeUnidade", "unidadeAdministrativa", "unidade"])

    if tipo == "Atas de Registro de Preços":
        numero = obter_primeiro_valor(row, ["numeroAtaRegistroPreco", "numeroAta", "numeroRegistroPreco", "numero"])
        valor = obter_primeiro_valor(row, ["valorTotal", "valorGlobal", "valorTotalAta", "valorAta", "valorTotalEstimado", "valorTotalHomologado", "valor"], padrao=None)
        data_ref = obter_primeiro_valor(row, ["dataAssinatura", "dataAssinaturaAta", "dataCelebracao"])
        vig_ini = obter_primeiro_valor(row, ["vigenciaInicio", "dataInicioVigencia", "dataVigenciaInicio"])
        vig_fim = obter_primeiro_valor(row, ["vigenciaFim", "dataFimVigencia", "dataVigenciaFim"])
    elif tipo == "Contratos":
        numero = obter_primeiro_valor(row, ["numeroContratoEmpenho", "numeroContrato", "numeroContratoPncp", "numero"])
        valor = obter_primeiro_valor(row, ["valorGlobal", "valorInicial", "valorTotal", "valorContrato"], padrao=None)
        data_ref = obter_primeiro_valor(row, ["dataAssinatura", "dataCelebracao"])
        vig_ini = obter_primeiro_valor(row, ["dataVigenciaInicio", "vigenciaInicio"])
        vig_fim = obter_primeiro_valor(row, ["dataVigenciaFim", "vigenciaFim"])
    else:  # Editais
        numero = obter_primeiro_valor(row, ["numeroCompra", "numeroEdital", "numero"])
        valor = obter_primeiro_valor(row, ["valorTotalHomologado", "valorTotalEstimado", "valorEstimado", "valorTotal", "valor"], padrao=None)
        data_ref = obter_primeiro_valor(row, ["dataPublicacao", "dataPublicacaoPncp", "dataAberturaProposta", "dataInclusao"])
        vig_ini = "N/D"
        vig_fim = "N/D"

    return {
        "id_pncp": texto_valido(id_pncp),
        "numero": texto_valido(numero),
        "processo": texto_valido(processo),
        "objeto": texto_valido(objeto),
        "fornecedor": texto_valido(fornecedor),
        "cnpj_fornecedor": texto_valido(cnpj_fornecedor),
        "valor_num": valor_numerico(valor) or 0.0,
        "valor": formatar_moeda_br(valor),
        "data_referencia": formatar_data(data_ref),
        "vigencia_inicio": formatar_data(vig_ini),
        "vigencia_fim": formatar_data(vig_fim),
        "situacao": texto_valido(situacao),
        "orgao": texto_valido(orgao),
        "unidade": texto_valido(unidade),
    }


def obter_identificador_contrato(row: Any) -> Tuple[Optional[str], Optional[int], Optional[int]]:
    cnpj = normalizar_cnpj(obter_primeiro_valor(row, ["cnpjOrgao", "cnpj", "cnpjCompra", "cnpjOrgaoEntidade"], padrao=None))
    ano = obter_primeiro_valor(row, ["anoContrato", "anoContratoEmpenho", "ano"], padrao=None)
    sequencial = obter_primeiro_valor(row, ["sequencialContrato", "sequencialContratoEmpenho", "sequencial"], padrao=None)

    try:
        ano_int = int(ano)
    except Exception:
        ano_int = None
    try:
        seq_int = int(sequencial)
    except Exception:
        seq_int = None

    return (cnpj if len(cnpj) == 14 else None, ano_int, seq_int)


# ============================================================
# CLIENTE HTTP ASSÍNCRONO ULTRA RÁPIDO
# ============================================================

async def fetch_page(client: httpx.AsyncClient, url: str, params: Dict[str, Any], page_num: int) -> List[Dict[str, Any]]:
    """Busca uma página individual de forma assíncrona com retry rápido."""
    p = params.copy()
    p["pagina"] = page_num

    for tentativa in range(1, 4):
        try:
            resp = await client.get(url, params=p, timeout=TIMEOUT_HTTP)
            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, dict):
                    for k in ("data", "items", "content", "dados", "registros"):
                        if isinstance(data.get(k), list):
                            return data[k]
                elif isinstance(data, list):
                    return data
                return []
            if resp.status_code in (204, 404):
                return []
        except Exception:
            if tentativa < 3:
                await asyncio.sleep(0.2 * tentativa)
                continue
    return []


async def consultar_pncp_turbo(url: str, params: Dict[str, Any], max_paginas: int = MAX_PAGINAS) -> Tuple[List[Dict[str, Any]], int]:
    """
    1. Busca a página 1 para descobrir a quantidade total de páginas.
    2. Dispara TODAS as páginas restantes simultaneamente em paralelo.
    """
    limits = httpx.Limits(max_keepalive_connections=20, max_connections=30)
    async with httpx.AsyncClient(headers=HEADERS, limits=limits, timeout=TIMEOUT_HTTP) as client:
        # Passo 1: Busca a página inicial
        p1_params = params.copy()
        p1_params["pagina"] = 1

        try:
            resp = await client.get(url, params=p1_params)
            if resp.status_code == 204 or resp.status_code != 200:
                return [], 0
            dados_p1 = resp.json()
        except Exception:
            return [], 0

        registros_totais = []
        if isinstance(dados_p1, dict):
            for k in ("data", "items", "content", "dados", "registros"):
                if isinstance(dados_p1.get(k), list):
                    registros_totais.extend(dados_p1[k])
                    break
            total_paginas = dados_p1.get("totalPaginas", 1)
        elif isinstance(dados_p1, list):
            registros_totais.extend(dados_p1)
            total_paginas = 1
        else:
            total_paginas = 1

        total_paginas = min(int(total_paginas or 1), max_paginas)

        # Passo 2: Disparo paralelo das páginas restantes
        if total_paginas > 1:
            tasks = [fetch_page(client, url, params, p) for p in range(2, total_paginas + 1)]
            resultados = await asyncio.gather(*tasks)
            for lote in resultados:
                registros_totais.extend(lote)

        return registros_totais, total_paginas


@st.cache_data(ttl=1800, show_spinner=False)
def executar_consulta_cacheada(url: str, params_tuple: Tuple[Tuple[str, str], ...], max_paginas: int) -> Tuple[List[Dict[str, Any]], int]:
    """Envoltório síncrono compatível com o cache nativo do Streamlit."""
    params = dict(params_tuple)
    return asyncio.run(consultar_pncp_turbo(url, params, max_paginas))


def executar_consulta(url: str, params: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], int]:
    params_serializados = tuple(sorted((str(k), str(v)) for k, v in params.items()))
    return executar_consulta_cacheada(url, params_serializados, MAX_PAGINAS)


def consultar_detalhes_contrato(row: Any) -> List[Dict[str, Any]]:
    cnpj, ano, sequencial = obter_identificador_contrato(row)
    if not all([cnpj, ano, sequencial]):
        return []
    url = f"{BASE_URL}/orgaos/{cnpj}/contratos/{ano}/{sequencial}"
    try:
        with httpx.Client(headers=HEADERS, timeout=15.0) as client:
            resp = client.get(url)
            if resp.status_code != 200:
                return []
            contrato = resp.json()
            if not isinstance(contrato, dict):
                return []

            documentos = []
            for chave, nome_padrao in [
                ("termos", "Termo de Contrato"),
                ("termosAditivos", "Termo Aditivo"),
                ("termosApostilamentos", "Termo de Apostilamento"),
                ("historico", "Histórico")
            ]:
                lista = contrato.get(chave, [])
                if isinstance(lista, list):
                    for item in lista:
                        it = dict(item)
                        it["tipoDocumentoNome"] = it.get("tipoDocumentoNome") or it.get("tipoTermoNome") or it.get("tipoEventoNome") or nome_padrao
                        documentos.append(it)
            return documentos
    except Exception:
        return []


# ============================================================
# EXPORTADORES (WORD & PDF)
# ============================================================

def gerar_word(df: pd.DataFrame, tipo_consulta: str, data_inicio, data_fim, modalidade_nome: Optional[str] = None) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Pt(36)

    p_head = doc.add_paragraph()
    p_head.alignment = 1
    r_head = p_head.add_run("PORTAL NACIONAL DE CONTRATAÇÕES PÚBLICAS – PNCP\nRELATÓRIO EXECUTIVO DE AUDITORIA E CONSULTA")
    r_head.bold = True
    r_head.font.size = Pt(14)
    r_head.font.color.rgb = RGBColor(0, 51, 102)

    p_sub = doc.add_paragraph()
    p_sub.alignment = 1
    r_sub = p_sub.add_run(f"{NOME_PREFEITURA} | {tipo_consulta.upper()}" + (f" - {modalidade_nome}" if modalidade_nome else ""))
    r_sub.font.size = Pt(10)

    doc.add_paragraph(f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')} | Registros: {len(df)}")
    doc.add_heading("Registros Selecionados", level=2)

    for idx, (_, row) in enumerate(df.head(150).iterrows(), start=1):
        p = doc.add_paragraph()
        r = p.add_run(f"Item #{idx:03d} - {row['objeto'][:120]}...")
        r.bold = True
        r.font.size = Pt(10)

        tabela = doc.add_table(rows=0, cols=2)
        tabela.style = "Table Grid"
        itens = [
            ("Número / ID PNCP", f"{row['numero']} ({row['id_pncp']})"),
            ("Processo / Situação", f"{row['processo']} | Status: {row['situacao']}"),
            ("Fornecedor / Documento", f"{row['fornecedor']} ({row['cnpj_fornecedor']})"),
            ("Valor / Data", f"{row['valor']} | Data: {row['data_referencia']}"),
        ]
        for k, v in itens:
            row_cells = tabela.add_row().cells
            row_cells[0].text = k
            row_cells[1].text = texto_valido(v)
        doc.add_paragraph()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def gerar_pdf(df: pd.DataFrame, tipo_consulta: str, data_inicio, data_fim, modalidade_nome: Optional[str] = None) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 8, "PORTAL NACIONAL DE CONTRATACOES PUBLICAS - PNCP", ln=True, align="C")
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, f"{NOME_PREFEITURA} - {tipo_consulta.upper()}", ln=True, align="C")
    pdf.set_font("Arial", size=8)
    pdf.cell(0, 5, f"Periodo: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')} | Registros: {len(df)}", ln=True, align="C")
    pdf.ln(5)

    for idx, (_, row) in enumerate(df.head(100).iterrows(), start=1):
        pdf.set_font("Arial", "B", 9)
        pdf.cell(0, 6, f"#{idx:03d} | Doc: {row['numero']} | Valor: {row['valor']}".encode("latin-1", "replace").decode("latin-1"), ln=True)
        pdf.set_font("Arial", size=8)
        pdf.multi_cell(0, 4, f"Objeto: {row['objeto']}\nFornecedor: {row['fornecedor']} ({row['cnpj_fornecedor']}) | Data: {row['data_referencia']}".encode("latin-1", "replace").decode("latin-1"), border=1)
        pdf.ln(2)

    res = pdf.output(dest="S")
    return res.encode("latin-1") if isinstance(res, str) else bytes(res)


# ============================================================
# INTERFACE PRINCIPAL & BARRA LATERAL
# ============================================================

st.title("🏛️ Radar PNCP Analytics - Rio das Pedras/SP")
st.caption("Painel de Inteligência, Transparência e Auditoria de Contratações Públicas (Lei nº 14.133/2021)")

with st.sidebar:
    st.header("⚙️ Parâmetros da Consulta")
    tipo_consulta = st.selectbox(
        "Tipo de Registro:",
        ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratações"]
    )

    modalidade_codigo = None
    modalidade_nome_exportacao = None

    if tipo_consulta == "Editais e Avisos de Contratações":
        modalidade_opcoes = {
            "Pregão - Eletrônico (6)": 6,
            "Dispensa de Licitação (8)": 8,
            "Inexigibilidade (9)": 9,
            "Concorrência - Eletrônica (2)": 2,
        }
        mod_escolhida = st.selectbox("Modalidade:", list(modalidade_opcoes.keys()))
        modalidade_codigo = modalidade_opcoes[mod_escolhida]
        modalidade_nome_exportacao = mod_escolhida

    col_d1, col_d2 = st.columns(2)
    data_inicio = col_d1.date_input("Data Inicial", value=datetime.date(2026, 1, 1))
    data_fim = col_d2.date_input("Data Final", value=datetime.date.today())

    if data_fim < data_inicio:
        st.error("Data final anterior à inicial.")
        st.stop()

    btn_executar = st.button("🔎 Buscar no PNCP", type="primary", use_container_width=True)


# ============================================================
# EXECUÇÃO DA CONSULTA
# ============================================================

if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = None
if "tipo_anterior" not in st.session_state:
    st.session_state.tipo_anterior = tipo_consulta

if st.session_state.tipo_anterior != tipo_consulta:
    st.session_state.df_resultado = None
    st.session_state.tipo_anterior = tipo_consulta

if btn_executar:
    endpoints = {
        "Contratos": f"{BASE_URL}/contratos",
        "Atas de Registro de Preços": f"{BASE_URL}/atas",
        "Editais e Avisos de Contratações": f"{BASE_URL}/contratacoes/publicacao",
    }
    tamanhos = {
        "Contratos": TAMANHO_PAGINA_CONTRATOS,
        "Atas de Registro de Preços": TAMANHO_PAGINA_ATAS,
        "Editais e Avisos de Contratações": TAMANHO_PAGINA_EDITAIS,
    }

    params = {
        "dataInicial": data_inicio.strftime("%Y%m%d"),
        "dataFinal": data_fim.strftime("%Y%m%d"),
        "pagina": 1,
        "tamanhoPagina": tamanhos[tipo_consulta],
    }

    if tipo_consulta == "Editais e Avisos de Contratações":
        params.update({
            "codigoModalidadeContratacao": modalidade_codigo,
            "uf": UF,
            "codigoMunicipioIbge": CODIGO_IBGE_RIO_DAS_PEDRAS,
            "cnpj": CNPJ_RIO_DAS_PEDRAS,
        })
    elif tipo_consulta == "Contratos":
        params["cnpjOrgao"] = CNPJ_RIO_DAS_PEDRAS
    else:
        params["cnpj"] = CNPJ_RIO_DAS_PEDRAS

    try:
        t0 = time.time()
        with st.spinner("Consultando dados oficiais no PNCP em paralelo..."):
            registros, paginas = executar_consulta(endpoints[tipo_consulta], params)
            df_bruto = pd.DataFrame([obter_dados_registro(r, tipo_consulta) for r in registros])
            if not df_bruto.empty:
                df_bruto = df_bruto.drop_duplicates(subset=["id_pncp", "numero"]).reset_index(drop=True)
            st.session_state.df_resultado = df_bruto
            st.session_state.tempo_consulta = time.time() - t0
            st.session_state.paginas_consultadas = paginas
    except Exception as e:
        st.error(f"Erro na conexão com o PNCP: {e}")
        st.session_state.df_resultado = None


# ============================================================
# EXIBIÇÃO ANALÍTICA E INTERATIVA
# ============================================================

df = st.session_state.df_resultado

if df is not None and not df.empty:
    # ------------------ KPIs GERENCIAIS ------------------
    valor_total = df["valor_num"].sum()
    ticket_medio = df["valor_num"].mean() if len(df) > 0 else 0.0
    fornecedores_unicos = df["cnpj_fornecedor"].replace("N/D", None).dropna().nunique()

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total de Registros", f"{len(df):,}")
    kpi2.metric("Montante Total", formatar_moeda_br(valor_total))
    kpi3.metric("Ticket Médio", formatar_moeda_br(ticket_medio))
    tempo_exec = st.session_state.get("tempo_consulta", 0.0)
    kpi4.metric("Tempo de Consulta", f"{tempo_exec:.2f} s")

    # ------------------ AUDITORIA: FRACIONAMENTO / DISPENSAS ------------------
    if tipo_consulta == "Editais e Avisos de Contratações" and modalidade_codigo == 8:
        st.markdown("---")
        st.subheader("🚩 Trilha de Auditoria: Monitoramento de Dispensas")
        dispensas_fornecedor = df.groupby(["fornecedor", "cnpj_fornecedor"])["valor_num"].agg(["count", "sum"]).reset_index()
        dispensas_alerta = dispensas_fornecedor[(dispensas_fornecedor["count"] > 1) & (dispensas_fornecedor["fornecedor"] != "N/D")]

        if not dispensas_alerta.empty:
            st.warning(f"⚠️ Atenção: Detectadas {len(dispensas_alerta)} empresas com múltiplas dispensas diretas no período selecionado.")
            st.dataframe(
                dispensas_alerta.rename(columns={"count": "Qtd. Contratações", "sum": "Soma Total (R$)"}),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("✅ Nenhuma concentração excessiva de dispensas diretas por fornecedor identificada.")

    # ------------------ VISUALIZAÇÃO GRÁFICA (PLOTLY) ------------------
    st.markdown("---")
    tab_graf1, tab_graf2 = st.tabs(["📊 Evolução Temporal e Status", "🏢 Top Fornecedores / Valores"])

    with tab_graf1:
        df_plot = df.copy()
        df_plot["data_dt"] = pd.to_datetime(df_plot["data_referencia"], format="%d/%m/%Y", errors="coerce")
        df_plot = df_plot.dropna(subset=["data_dt"])
        if not df_plot.empty:
            df_evolucao = df_plot.groupby(df_plot["data_dt"].dt.to_period("M").dt.to_timestamp())["valor_num"].sum().reset_index()
            fig1 = px.bar(df_evolucao, x="data_dt", y="valor_num", labels={"data_dt": "Mês/Ano", "valor_num": "Volume Total (R$)"}, title="Volume Financeiro por Mês")
            fig1.update_layout(yaxis_tickprefix="R$ ", hovermode="x unified")
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("Sem datas estruturadas suficientes para evolução temporal.")

    with tab_graf2:
        df_forn = df[df["fornecedor"] != "N/D"].groupby("fornecedor")["valor_num"].sum().nlargest(10).reset_index()
        if not df_forn.empty:
            fig2 = px.bar(df_forn, x="valor_num", y="fornecedor", orientation="h", labels={"valor_num": "Total Homologado/Contratado (R$)", "fornecedor": "Fornecedor"}, title="Top 10 Fornecedores por Volume Financeiro")
            fig2.update_layout(xaxis_tickprefix="R$ ", yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Sem identificação de fornecedores para ranking financeiro.")

    # ------------------ FILTROS EM MEMÓRIA & TABELA ------------------
    st.markdown("---")
    st.subheader("📋 Tabela Analítica de Contratações")

    col_busca, _ = st.columns([2, 1])
    termo_busca = col_busca.text_input("🔍 Filtrar tabela por palavra-chave (Objeto, Fornecedor ou CNPJ):", placeholder="Digite para filtrar instantaneamente...")

    df_exibicao = df.copy()
    if termo_busca:
        mascara = (
            df_exibicao["objeto"].str.contains(termo_busca, case=False, na=False) |
            df_exibicao["fornecedor"].str.contains(termo_busca, case=False, na=False) |
            df_exibicao["cnpj_fornecedor"].str.contains(termo_busca, case=False, na=False)
        )
        df_exibicao = df_exibicao[mascara]

    st.dataframe(
        df_exibicao[["numero", "objeto", "fornecedor", "cnpj_fornecedor", "valor", "data_referencia", "situacao", "id_pncp"]],
        column_config={
            "numero": "Documento / Edital",
            "objeto": "Objeto da Contratação",
            "fornecedor": "Fornecedor",
            "cnpj_fornecedor": "CNPJ/CPF",
            "valor": "Valor Total",
            "data_referencia": "Data",
            "situacao": "Status",
            "id_pncp": "Controle PNCP"
        },
        use_container_width=True,
        hide_index=True
    )

    # ------------------ EXPORTAÇÃO DE RELATÓRIOS ------------------
    st.markdown("---")
    st.subheader("📥 Central de Exportação")
    exp1, exp2, exp3, exp4 = st.columns(4)

    buf_xlsx = io.BytesIO()
    df.to_excel(buf_xlsx, index=False)
    exp1.download_button("📊 Planilha Excel", buf_xlsx.getvalue(), f"PNCP_{tipo_consulta}_RioDasPedras.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    csv_data = df.to_csv(index=False, encoding="utf-8-sig")
    exp2.download_button("📄 Dados em CSV", csv_data, f"PNCP_{tipo_consulta}_RioDasPedras.csv", mime="text/csv", use_container_width=True)

    try:
        w_bytes = gerar_word(df, tipo_consulta, data_inicio, data_fim, modalidade_nome_exportacao)
        exp3.download_button("📝 Relatório Word (.docx)", w_bytes, f"Relatorio_PNCP_{tipo_consulta}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    except Exception as ew:
        exp3.error(f"Erro Word: {ew}")

    try:
        p_bytes = gerar_pdf(df, tipo_consulta, data_inicio, data_fim, modalidade_nome_exportacao)
        exp4.download_button("📕 Relatório PDF", p_bytes, f"Relatorio_PNCP_{tipo_consulta}.pdf", mime="application/pdf", use_container_width=True)
    except Exception as ep:
        exp4.error(f"Erro PDF: {ep}")

    # ------------------ DETALHAMENTO DE CONTRATOS & ADITIVOS ------------------
    if tipo_consulta == "Contratos":
        st.markdown("---")
        st.subheader("🔍 Investigação de Aditivos e Histórico de Contrato")
        opcoes = [f"{r['numero']} | {r['fornecedor']} | PNCP: {r['id_pncp']}" for _, r in df.iterrows()]
        contrato_escolhido = st.selectbox("Selecione o contrato para auditar aditivos:", opcoes)

        if st.button("Buscar Documentos e Termos Aditivos", use_container_width=True):
            idx_sel = opcoes.index(contrato_escolhido)
            row_sel = df.iloc[idx_sel]
            with st.spinner("Buscando termos aditivos e histórico na API..."):
                docs = consultar_detalhes_contrato(row_sel)
                if docs:
                    st.success(f"Foram encontrados {len(docs)} termo(s) associado(s).")
                    for d in docs:
                        st.info(f"**Tipo:** {d.get('tipoDocumentoNome')}\n\n**Número:** {d.get('numero', 'S/N')}\n\n**Data:** {formatar_data(d.get('dataPublicacaoPncp') or d.get('dataAssinatura'))}\n\n**Objeto:** {d.get('objeto', 'N/D')}")
                else:
                    st.warning("Nenhum termo aditivo localizado para este contrato no PNCP.")

elif df is not None and df.empty:
    st.warning("⚠️ Nenhum registro localizado no PNCP para os parâmetros informados.")
