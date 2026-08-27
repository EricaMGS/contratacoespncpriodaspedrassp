import io
import zipfile
import datetime
import time
import re
import os
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import numpy as np
import requests
import streamlit as st

from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler

from docx import Document
from docx.shared import Pt, RGBColor
from fpdf import FPDF


# ============================================================
# 🏛️ CONFIGURAÇÃO DA PÁGINA E DO SISTEMA
# ============================================================

st.set_page_config(
    page_title="SACI — Sistema de Apoio ao Controle Interno",
    page_icon="🛡️",
    layout="wide",
)

CNPJ_RIO_DAS_PEDRAS = "44826840000183"
CODIGO_IBGE_RIO_DAS_PEDRAS = "3544004"
UF = "SP"

BASE_URL = "https://pncp.gov.br/api/consulta/v1"
BASE_DOCUMENTOS_URL = "https://pncp.gov.br/api/pncp/v1"

NOME_MUNICIPIO = "Rio das Pedras/SP"
NOME_PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"

TAMANHO_PAGINA_CONTRATOS = 100
TAMANHO_PAGINA_ATAS = 100
TAMANHO_PAGINA_EDITAIS = 50

# Parâmetros de segurança e limite de paginação para evitar instabilidade
MAX_PAGINAS_SEGURO = 10 
MAX_TENTATIVAS = 4
TIMEOUT_CONEXAO = 15
TIMEOUT_LEITURA = 60

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
    "Connection": "keep-alive",
}


# ============================================================
# 🛠️ UTILITÁRIOS DE DADOS E FORMATAÇÃO
# ============================================================

def texto_valido(valor: Any, padrao: str = "N/D") -> str:
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
        partes = [texto_valido(item, "") for item in valor]
        partes = [i for i in partes if i]
        return ", ".join(partes) if partes else padrao
    texto = str(valor).strip()
    if texto.lower() in {"", "none", "nan", "null", "n/d", "nd", "nat"}:
        return padrao
    return texto


def valor_numerico(valor: Any) -> Optional[float]:
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
    numero = valor_numerico(valor)
    if numero is None:
        return "N/D"
    texto = f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def formatar_data(valor: Any) -> str:
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
    return re.sub(r"\D", "", texto_valido(valor, ""))


def extrair_valor_recursivo(valor: Any, profundidade: int = 0) -> Any:
    if profundidade > 4 or valor is None:
        return valor
    if isinstance(valor, dict):
        for chave in ("nome", "razaoSocial", "descricao", "descricaoObjeto", "valor", "numero", "cnpj", "ni"):
            if chave in valor:
                res = extrair_valor_recursivo(valor[chave], profundidade + 1)
                if res not in (None, "", "N/D"):
                    return res
        return valor
    if isinstance(valor, list):
        resultados = [str(extrair_valor_recursivo(item, profundidade + 1)) for item in valor if item is not None]
        return ", ".join([r for r in resultados if r not in ("", "N/D")]) if resultados else None
    return valor


def obter_primeiro_valor(row: Any, campos: List[str], padrao: Any = "N/D") -> Any:
    if row is None:
        return padrao
    for campo in campos:
        try:
            val = row.get(campo) if isinstance(row, (dict, pd.Series)) else None
            if val is not None and not (isinstance(val, float) and pd.isna(val)):
                res = extrair_valor_recursivo(val)
                if res not in (None, "", "N/D"):
                    return res
        except Exception:
            continue
    return padrao


# ============================================================
# 🌐 CLIENTE HTTP ROBUSTO COM CONTROLE DE ESTABILIDADE
# ============================================================

@st.cache_resource
def criar_sessao_http():
    sessao = requests.Session()
    sessao.headers.update(HEADERS)
    return sessao


def consultar_pncp(url: str, params: Dict[str, Any], max_tentativas: int = MAX_TENTATIVAS) -> Any:
    sessao = criar_sessao_http()
    ultimo_erro = None
    for tentativa in range(1, max_tentativas + 1):
        try:
            response = sessao.get(url, params=params, timeout=(TIMEOUT_CONEXAO, TIMEOUT_LEITURA))
            if response.status_code == 200:
                try:
                    return response.json()
                except ValueError as e:
                    raise RuntimeError("API retornou HTTP 200, mas o conteúdo não é um JSON válido.") from e
            if response.status_code == 204:
                return []
            if response.status_code in (400, 422):
                raise RuntimeError(f"Parâmetros rejeitados pelo PNCP (HTTP {response.status_code}).")
            if response.status_code == 429:
                time.sleep(5 * tentativa)
                continue
            if response.status_code in (500, 502, 503, 504):
                ultimo_erro = RuntimeError(f"Erro no servidor PNCP (HTTP {response.status_code}).")
                time.sleep(2 ** tentativa)
                continue
            raise RuntimeError(f"PNCP retornou erro HTTP {response.status_code}.")
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            ultimo_erro = e
            if tentativa < max_tentativas:
                time.sleep(2 ** tentativa)
                continue
            raise RuntimeError("Tempo limite esgotado ao consultar o PNCP. O portal pode estar instável.") from e
    raise RuntimeError("Falha de comunicação com o PNCP após várias tentativas.") from ultimo_erro


def consultar_paginas_seguro(url: str, params: Dict[str, Any], max_paginas: int = MAX_PAGINAS_SEGURO) -> List[Dict[str, Any]]:
    todos_registros = []
    for pagina in range(1, max_paginas + 1):
        p_params = params.copy()
        p_params["pagina"] = pagina
        try:
            dados = consultar_pncp(url, p_params)
            registros = dados.get("data") or dados.get("items") or dados.get("content") or (dados if isinstance(dados, list) else [])
            if not registros:
                break
            todos_registros.extend(registros)
            if len(registros) < int(params.get("tamanhoPagina", 50)):
                break
            time.sleep(0.2)
        except Exception:
            break
    return todos_registros


# ============================================================
# 🗂️ PROCESSAMENTO E NORMALIZAÇÃO DE REGISTROS
# ============================================================

def extrair_dados_padrao(row: Any, tipo: str) -> Dict[str, str]:
    id_pncp = obter_primeiro_valor(row, ["numeroControlePNCP", "numeroControlePNCPAta", "idContratoPNCP"])
    
    if tipo == "Contratos":
        numero = obter_primeiro_valor(row, ["numeroContratoEmpenho", "numeroContrato", "numero"])
        processo = obter_primeiro_valor(row, ["processo", "numeroProcesso"])
        objeto = obter_primeiro_valor(row, ["objetoContrato", "objetoCompra", "objeto"])
        fornecedor = obter_primeiro_valor(row, ["nomeRazaoSocialFornecedor", "razaoSocialFornecedor", "nomeFornecedor"])
        cnpj_forn = obter_primeiro_valor(row, ["niFornecedor", "cnpjFornecedor"])
        valor = obter_primeiro_valor(row, ["valorGlobal", "valorInicial", "valorTotal", "valorContrato"], padrao=None)
        data_ass = obter_primeiro_valor(row, ["dataAssinatura", "dataCelebracao"])
        situacao = obter_primeiro_valor(row, ["situacao", "status"])
    elif tipo == "Atas de Registro de Preços":
        numero = obter_primeiro_valor(row, ["numeroAtaRegistroPreco", "numeroAta", "numero"])
        processo = obter_primeiro_valor(row, ["processo", "numeroProcesso"])
        objeto = obter_primeiro_valor(row, ["objetoCompra", "objeto"])
        fornecedor = obter_primeiro_valor(row, ["nomeRazaoSocialFornecedor", "razaoSocialFornecedor"])
        cnpj_forn = obter_primeiro_valor(row, ["niFornecedor", "cnpjFornecedor"])
        valor = obter_primeiro_valor(row, ["valorTotal", "valorGlobal", "valorAta"], padrao=None)
        data_ass = obter_primeiro_valor(row, ["dataAssinatura"])
        situacao = obter_primeiro_valor(row, ["situacao"])
    else:
        numero = obter_primeiro_valor(row, ["numeroCompra", "numeroEdital", "numero"])
        processo = obter_primeiro_valor(row, ["processo", "numeroProcesso"])
        objeto = obter_primeiro_valor(row, ["objetoCompra", "objeto", "descricaoObjeto"])
        fornecedor = obter_primeiro_valor(row, ["nomeRazaoSocialFornecedor", "razaoSocialFornecedor"])
        cnpj_forn = obter_primeiro_valor(row, ["niFornecedor", "cnpjFornecedor"])
        valor = obter_primeiro_valor(row, ["valorTotalHomologado", "valorTotalEstimado", "valorTotal"], padrao=None)
        data_ass = obter_primeiro_valor(row, ["dataPublicacao", "dataInclusao"])
        situacao = obter_primeiro_valor(row, ["situacaoCompra", "situacao"])

    return {
        "id_pncp": texto_valido(id_pncp),
        "numero": texto_valido(numero),
        "processo": texto_valido(processo),
        "objeto": texto_valido(objeto),
        "fornecedor": texto_valido(fornecedor),
        "cnpj_fornecedor": texto_valido(cnpj_forn),
        "valor_numerico": valor_numerico(valor) or 0.0,
        "valor": formatar_moeda_br(valor),
        "data": formatar_data(data_ass),
        "situacao": texto_valido(situacao),
    }


# ============================================================
# 🚦 MATRIZ DE RISCO AVANÇADA & CRITÉRIOS DE AUDITORIA
# ============================================================

def avaliar_risco_contratacao(row: Any, df_historico: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """
    Avaliação multifatorial de risco para auditoria interna.
    Separa Risco Inerente, Critérios de Alerta, Evidências e Prioridade.
    """
    pontos = 0
    criterios_alerta = []
    
    val = valor_numerico(obter_primeiro_valor(row, ["valorGlobal", "valorInicial", "valorTotal", "valorTotalHomologado"], None)) or 0.0
    objeto = str(obter_primeiro_valor(row, ["objetoContrato", "objetoCompra", "objeto"], "")).lower()
    modalidade = str(obter_primeiro_valor(row, ["modalidadeNome", "modalidadeContratacaoNome"], "")).lower()

    # 1. Análise de Risco Inerente (Materialidade Financeira Contextual)
    if val > 1000000:
        pontos += 30
        criterios_alerta.append("Materialidade elevada (> R$ 1 milhão)")
    elif val > 300000:
        pontos += 15
        criterios_alerta.append("Materialidade intermediária relevante")

    # 2. Critérios de Alerta de Procedimento (Dispensa / Emergência)
    if "dispensa" in modalidade or "inexigibilidade" in modalidade:
        pontos += 25
        criterios_alerta.append("Contratação direta (Dispensa/Inexigibilidade)")
    
    if "emergencial" in objeto or "emergência" in objeto or "calamidade" in objeto:
        pontos += 35
        criterios_alerta.append("Alegação de situação emergencial")

    # 3. Análise estatística de desvio se houver histórico carregado
    if df_historico is not None and not df_historico.empty and "valor_numerico" in df_historico.columns:
        media_historica = df_historico["valor_numerico"].mean()
        if media_historica > 0 and val > (media_historica * 3):
            pontos += 25
            criterios_alerta.append(f"Valor 3x superior à média histórica da amostra ({formatar_moeda_br(media_historica)})")

    pontos = min(pontos, 100)

    if pontos >= 60:
        prioridade = "🔴 ALTA PRIORIDADE"
        risco_inerente = "Alto"
    elif pontos >= 30:
        prioridade = "🟡 MÉDIA PRIORIDADE"
        risco_inerente = "Moderado"
    else:
        prioridade = "🟢 BAIXA PRIORIDADE"
        risco_inerente = "Baixo"

    if not criterios_alerta:
        criterios_alerta.append("Nenhum desvio crítico automatizado nos parâmetros básicos")

    return {
        "pontos": pontos,
        "prioridade": prioridade,
        "risco_inerente": risco_inerente,
        "criterios": criterios_alerta,
        "evidencia_amostra": f"Objeto analisado: {objeto[:80]}..."
    }


# ============================================================
# 🤖 MÓDULO DE MACHINE LEARNING (DETECÇÃO DE ANOMALIAS)
# ============================================================

def executar_modelo_ml_anomalias(df: pd.DataFrame) -> pd.DataFrame:
    """
    Usa Isolation Forest para identificar contratos/compras estatisticamente anômalos
    com base no valor financeiro e complexidade do texto.
    """
    if df.empty or len(df) < 5:
        df["anomalia_ml"] = "Normal"
        return df

    df_ml = df.copy()
    df_ml["tamanho_objeto"] = df_ml["objeto"].astype(str).apply(len)
    
    features = df_ml[["valor_numerico", "tamanho_objeto"]].fillna(0)
    
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(features)

    model = IsolationForest(contamination=0.15, random_state=42)
    preds = model.fit_predict(X_scaled)
    
    df_ml["anomalia_ml"] = ["Anômalo (Outlier)" if p == -1 else "Normal" for p in preds]
    return df_ml


# ============================================================
# 📋 GERAÇÃO DE PAPEL DE TRABALHO DE AUDITORIA (WORD)
# ============================================================

def gerar_papel_trabalho_auditoria_word(row: Any, tipo_consulta: str, avaliacao_risco: Dict[str, Any]) -> bytes:
    doc = Document()
    dados = obter_dados_padrao(row, tipo_consulta)

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = 1
    run = p_titulo.add_run("PAPEL DE TRABALHO DE AUDITORIA / CONTROLE INTERNO")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_paragraph(f"Órgão Fiscalizado: {NOME_PREFEITURA}")
    doc.add_paragraph(f"Data da Emissão: {datetime.date.today().strftime('%d/%m/%Y')}")
    
    doc.add_heading("1. Identificação da Contratação", level=2)
    doc.add_paragraph(f"• Processo: {dados['processo']}")
    doc.add_paragraph(f"• Número/Edital: {dados['numero']}")
    doc.add_paragraph(f"• Identificador PNCP: {dados['id_pncp']}")
    doc.add_paragraph(f"• Objeto: {dados['objeto']}")
    doc.add_paragraph(f"• Fornecedor: {dados['fornecedor']} (CNPJ: {dados['cnpj_fornecedor']})")
    doc.add_paragraph(f"• Valor Global Registrado: {dados['valor']}")

    doc.add_heading("2. Matriz de Risco e Alertas", level=2)
    doc.add_paragraph(f"• Nível de Prioridade: {avaliacao_risco['prioridade']}")
    doc.add_paragraph(f"• Índice de Risco Atribuído: {avaliacao_risco['pontos']}/100")
    doc.add_paragraph("• Critérios e Alertas Identificados:")
    for crit in avaliacao_risco["criterios"]:
        doc.add_paragraph(f"   - {crit}", style='List Bullet')

    doc.add_heading("3. Procedimentos de Controle e Conclusão", level=2)
    doc.add_paragraph("• Objetivo do Teste: Verificar a conformidade documental e materialidade da contratação pública.")
    doc.add_paragraph("• Evidência Encontrada: Verificação automatizada via cruzamento de dados do PNCP.")
    doc.add_paragraph("• Conclusão do Controlador: [ ] Homologado sem ressalvas   [X] Requer diligência presencial/documental.")
    
    doc.add_paragraph("\n\n__________________________________________________")
    doc.add_paragraph("Assinatura do Servidor / Controlador Interno")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# 🖥️ INTERFACE PRINCIPAL DO APLICATIVO (STREAMLIT)
# ============================================================

st.title("🛡️ Sistema de Apoio ao Controle Interno (SACI)")
st.markdown("Plataforma de inteligência analítica, matriz de risco e auditoria preventiva das contratações públicas de **Rio das Pedras/SP**.")
st.markdown("Dados coletados do Portal Nacional de Contratações Públicas - PNCP.")

st.sidebar.header("⚙️ Parâmetros de Consulta")
tipo_consulta = st.sidebar.selectbox(
    "Módulo de Análise:",
    ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratações"]
)

data_inicio = st.sidebar.date_input("📅 Data Inicial", value=pd.to_datetime("2026-01-01").date())
data_fim = st.sidebar.date_input("📅 Data Final", value=datetime.date.today())

if data_fim < data_inicio:
    st.sidebar.error("⚠️ A Data Final não pode ser anterior à Data Inicial.")
    st.stop()

if st.sidebar.button("🔎 Executar Varredura e Análise", type="primary", use_container_width=True):
    endpoints = {
        "Contratos": f"{BASE_URL}/contratos",
        "Atas de Registro de Preços": f"{BASE_URL}/atas",
        "Editais e Avisos de Contratações": f"{BASE_URL}/contratacoes/publicacao",
    }
    endpoint = endpoints[tipo_consulta]
    tamanho_pag = TAMANHO_PAGINA_CONTRATOS if tipo_consulta == "Contratos" else TAMANHO_PAGINA_ATAS

    params = {
        "dataInicial": data_inicio.strftime("%Y%m%d"),
        "dataFinal": data_fim.strftime("%Y%m%d"),
        "tamanhoPagina": tamanho_pag,
    }
    if tipo_consulta == "Contratos":
        params["cnpjOrgao"] = CNPJ_RIO_DAS_PEDRAS
    elif tipo_consulta == "Atas de Registro de Preços":
        params["cnpj"] = CNPJ_RIO_DAS_PEDRAS

    try:
        with st.spinner("Extraindo dados do PNCP e aplicando motores de IA..."):
            registros = consultar_paginas_seguro(endpoint, params)
            df_temp = pd.DataFrame(registros)
            
            if not df_temp.empty:
                # Normalização e limpeza
                lista_proc = [extrair_dados_padrao(row, tipo_consulta) for _, row in df_temp.iterrows()]
                df_processado = pd.DataFrame(lista_proc)
                
                # Executa modelo de Machine Learning (Isolation Forest)
                df_processado = executar_modelo_ml_anomalias(df_processado)
                st.session_state.df_resultado = df_processado
            else:
                st.session_state.df_resultado = pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erro na execução: {e}")

df = st.session_state.get("df_resultado", None)

if df is not None and not df.empty:
    st.markdown("---")
    st.subheader("📊 Painel Executivo de Controle Interno")

    # Avaliação de risco em lote para o painel
    riscos_lote = [avaliar_risco_contratacao(row, df) for _, row in df.iterrows()]
    altos = sum(1 for r in riscos_lote if "ALTA" in r["prioridade"])
    medios = sum(1 for r in riscos_lote if "MÉDIA" in r["prioridade"])
    anomalias_ia = len(df[df.get("anomalia_ml", "") == "Anômalo (Outlier)"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Analisado", len(df))
    c2.metric("🔴 Alta Prioridade", altos)
    c3.metric("🟡 Média Prioridade", medios)
    c4.metric("🤖 Alertas de Anomalia (ML)", anomalias_ia)

    st.markdown("---")
    st.subheader("🚨 Matriz de Risco Analítica e Auditoria")
    st.caption("Nota: Classificação baseada em critérios objetivos de materialidade, procedimentos e detecção de outliers por IA.")

    # Tabela consolidada com alertas explicados
    tabela_audit = []
    for i, row in df.iterrows():
        risco_item = avaliar_risco_contratacao(row, df)
        tabela_audit.append({
            "Prioridade": risco_item["prioridade"],
            "Risco Inerente": risco_item["risco_inerente"],
            "Anomalia IA": row.get("anomalia_ml", "Normal"),
            "Número": row["numero"],
            "Objeto": row["objeto"],
            "Fornecedor": row["fornecedor"],
            "Valor": row["valor"],
            "Motivos do Alerta": "; ".join(risco_item["criterios"])
        })
    
    df_audit_view = pd.DataFrame(tabela_audit)
    st.dataframe(df_audit_view, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("📋 Geração de Papel de Trabalho de Auditoria")
    
    opcoes_sel = [f"[{r['Prioridade']}] Proc: {r['Número']} | Forn: {r['Fornecedor']} ({r['Valor']})" for r in tabela_audit]
    escolha_processo = st.selectbox("Selecione o processo para emitir o Papel de Trabalho oficial:", opcoes_sel)

    if escolha_processo:
        idx_selecionado = opcoes_sel.index(escolha_processo)
        row_alvo = df.iloc[idx_selecionado]
        risco_alvo = avaliar_risco_contratacao(row_alvo, df)

        if st.button("📥 Baixar Papel de Trabalho Oficial (.docx)", type="primary"):
            docx_bytes = gerar_papel_trabalho_auditoria_word(row_alvo, tipo_consulta, risco_alvo)
            st.download_button(
                "⬇️ Clique aqui para baixar o arquivo Word",
                data=docx_bytes,
                file_name=f"Papel_Trabalho_Auditoria_{row_alvo['numero'].replace('/', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

elif df is not None and df.empty:
    st.warning("⚠️ Nenhum registro encontrado para o período especificado.")
else:
    st.info("💡 Configure o período na barra lateral e clique em **Executar Varredura e Análise** para carregar os dados e acionar a inteligência analítica.")

st.markdown("---")
st.caption("Dados do Portal Nacional de Contratações Públicas - PNCP.")
