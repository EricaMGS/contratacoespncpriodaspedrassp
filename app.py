import datetime
import io
import math
import re
import time
from dataclasses import dataclass
from threading import local
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from requests.adapters import HTTPAdapter
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import RobustScaler
from urllib3.util.retry import Retry

# ============================================================
# 1. CONFIGURAÇÕES E TIPOS
# ============================================================

st.set_page_config(
    page_title="SACI - Sistema de Apoio ao Controle Interno",
    page_icon="🛡️",
    layout="wide",
)

CNPJ_ORGAO = "44826840000183"
NOME_MUNICIPIO = "Rio das Pedras/SP"
NOME_PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"

BASE_URL = "https://pncp.gov.br/api/consulta/v1"
TIMEOUT = (10, 45)
TAMANHO_PAGINA = 50
MAX_PAGINAS = 100
MAX_TRABALHADORES = 4

HEADERS = {
    "User-Agent": "SACI-Controle-Interno/2.0",
    "Accept": "application/json",
}

_thread_local = local()

@dataclass(frozen=True)
class ConfiguracaoConsulta:
    endpoint: str
    parametros_extras: Dict[str, Any]


# ============================================================
# 2. FUNÇÕES DE CONVERSÃO E SEGURANÇA DE DADOS
# ============================================================

VALORES_AUSENTES = {"", "none", "nan", "null", "n/d", "nd", "nat"}

def valor_ausente(valor: Any) -> bool:
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(valor, str) and valor.strip().lower() in VALORES_AUSENTES

def texto_valido(valor: Any, padrao: str = "N/D") -> str:
    if valor_ausente(valor):
        return padrao
    if isinstance(valor, dict):
        campos_preferidos = (
            "nome",
            "razaoSocial",
            "descricao",
            "valor",
            "nomeUnidade",
            "municipioNome",
        )
        for campo in campos_preferidos:
            if campo in valor:
                resultado = texto_valido(valor[campo], "")
                if resultado:
                    return resultado
        return padrao
    if isinstance(valor, list):
        partes = [texto_valido(item, "") for item in valor]
        partes = [parte for parte in partes if parte]
        return ", ".join(partes) if partes else padrao
    texto = str(valor).strip()
    return texto if texto.lower() not in VALORES_AUSENTES else padrao

def valor_numerico(valor: Any) -> Optional[float]:
    if valor_ausente(valor) or isinstance(valor, bool):
        return None
    if isinstance(valor, dict):
        for campo in (
            "valor",
            "value",
            "valorTotal",
            "valorGlobal",
            "valorInicial",
        ):
            if campo in valor:
                numero = valor_numerico(valor[campo])
                if numero is not None:
                    return numero
        return None
    if isinstance(valor, (int, float)):
        numero = float(valor)
        return numero if math.isfinite(numero) else None
    texto = re.sub(r"[^\d,.\-]", "", str(valor))
    if not texto:
        return None
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        numero = float(texto)
        return numero if math.isfinite(numero) else None
    except ValueError:
        return None

def formatar_moeda_br(valor: Any) -> str:
    numero = valor_numerico(valor)
    if numero is None:
        return "N/D"
    texto = f"{numero:,.2f}"
    texto = texto.replace(",", "#").replace(".", ",").replace("#", ".")
    return f"R$ {texto}"

def formatar_data(valor: Any) -> str:
    if valor_ausente(valor):
        return "N/D"
    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return texto_valido(valor)
    return data.strftime("%d/%m/%Y")

def normalizar_documento(valor: Any) -> str:
    return re.sub(r"\D", "", texto_valido(valor, ""))

def obter_primeiro_valor(
    registro: Any,
    campos: Sequence[str],
    padrao: Any = None,
) -> Any:
    if registro is None:
        return padrao
    for campo in campos:
        try:
            valor = registro.get(campo)
        except (AttributeError, TypeError):
            continue
        if not valor_ausente(valor):
            return valor
    return padrao


# ============================================================
# 3. CLIENTE HTTP COM RETENTATIVAS (THREAD-SAFE)
# ============================================================

def criar_sessao_http() -> requests.Session:
    sessao = requests.Session()
    sessao.headers.update(HEADERS)
    estrategia = Retry(
        total=3,
        connect=3,
        read=3,
        status=3,
        backoff_factor=1.0,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
        respect_retry_after_header=True,
        raise_on_status=False,
    )
    adaptador = HTTPAdapter(
        max_retries=estrategia,
        pool_connections=MAX_TRABALHADORES,
        pool_maxsize=MAX_TRABALHADORES,
    )
    sessao.mount("https://", adaptador)
    return sessao

def obter_sessao_thread() -> requests.Session:
    if not hasattr(_thread_local, "sessao"):
        _thread_local.sessao = criar_sessao_http()
    return _thread_local.sessao

def consultar_pncp(
    url: str,
    parametros: Dict[str, Any],
) -> Dict[str, Any]:
    sessao = obter_sessao_thread()
    try:
        resposta = sessao.get(
            url,
            params=parametros,
            timeout=TIMEOUT,
        )
    except requests.Timeout as erro:
        raise RuntimeError("Tempo limite excedido ao consultar o PNCP.") from erro
    except requests.ConnectionError as erro:
        raise RuntimeError("Não foi possível conectar ao PNCP.") from erro
    except requests.RequestException as erro:
        # Cobre qualquer outra falha de rede/HTTP não prevista explicitamente
        # acima (ex.: erro de SSL, redirecionamento inválido, etc.), para que
        # a página nunca quebre com um traceback bruto para o usuário.
        raise RuntimeError(
            f"Falha inesperada de comunicação com o PNCP: {erro}"
        ) from erro
    if resposta.status_code == 204:
        return {"data": [], "totalPaginas": 0}
    if resposta.status_code != 200:
        detalhe = resposta.text[:300].strip()
        mensagem = (
            f"PNCP retornou HTTP {resposta.status_code}. "
            f"Endpoint: {url}. Parâmetros: {parametros}."
        )
        if detalhe:
            mensagem += f" Resposta: {detalhe}"
        raise RuntimeError(mensagem)
    try:
        conteudo = resposta.json()
    except ValueError as erro:
        raise RuntimeError(
            "O PNCP respondeu com conteúdo que não é JSON válido."
        ) from erro
    if isinstance(conteudo, list):
        return {
            "data": conteudo,
            "totalPaginas": 1,
        }
    if not isinstance(conteudo, dict):
        raise RuntimeError("Formato de resposta inesperado do PNCP.")
    return conteudo


# ============================================================
# 4. PAGINAÇÃO COMPLETA E VERIFICÁVEL
# ============================================================

def extrair_lista_resposta(conteudo: Dict[str, Any]) -> List[Dict[str, Any]]:
    for campo in ("data", "items", "content"):
        registros = conteudo.get(campo)
        if isinstance(registros, list):
            return registros
    return []

def extrair_total_paginas(
    conteudo: Dict[str, Any],
    tamanho_primeira_pagina: int,
    tamanho_pagina: int,
) -> int:
    for campo in ("totalPaginas", "totalPages"):
        total = conteudo.get(campo)
        if isinstance(total, int) and total >= 0:
            return total
    total_registros = conteudo.get("totalRegistros")
    if isinstance(total_registros, int) and total_registros >= 0:
        return math.ceil(total_registros / tamanho_pagina)
    return 1 if tamanho_primeira_pagina > 0 else 0

def buscar_pagina(
    url: str,
    parametros_base: Dict[str, Any],
    pagina: int,
) -> Tuple[int, List[Dict[str, Any]]]:
    parametros = dict(parametros_base)
    parametros["pagina"] = pagina
    resposta = consultar_pncp(url, parametros)
    return pagina, extrair_lista_resposta(resposta)

@st.cache_data(ttl=900, show_spinner=False)
def consultar_todas_paginas(
    url: str,
    parametros: Dict[str, Any],
    max_paginas: int = MAX_PAGINAS,
) -> List[Dict[str, Any]]:
    parametros_base = dict(parametros)
    parametros_base["pagina"] = 1
    resposta_inicial = consultar_pncp(url, parametros_base)
    primeira_pagina = extrair_lista_resposta(resposta_inicial)
    if not primeira_pagina:
        return []
    tamanho_pagina = int(
        parametros_base.get("tamanhoPagina", TAMANHO_PAGINA)
    )
    total_paginas = extrair_total_paginas(
        resposta_inicial,
        len(primeira_pagina),
        tamanho_pagina,
    )
    total_paginas = min(max(total_paginas, 1), max_paginas)
    if total_paginas == 1:
        return primeira_pagina
    paginas: Dict[int, List[Dict[str, Any]]] = {
        1: primeira_pagina
    }
    erros: List[str] = []
    from concurrent.futures import ThreadPoolExecutor, as_completed
    with ThreadPoolExecutor(
        max_workers=min(MAX_TRABALHADORES, total_paginas - 1)
    ) as executor:
        tarefas = {
            executor.submit(
                buscar_pagina,
                url,
                parametros_base,
                pagina,
            ): pagina
            for pagina in range(2, total_paginas + 1)
        }
        for tarefa in as_completed(tarefas):
            pagina = tarefas[tarefa]
            try:
                numero, registros = tarefa.result()
                paginas[numero] = registros
            except Exception as erro:
                erros.append(f"Página {pagina}: {erro}")
    if erros:
        raise RuntimeError(
            "A consulta ficou incompleta. "
            + " | ".join(erros[:5])
        )
    resultado: List[Dict[str, Any]] = []
    for pagina in sorted(paginas):
        resultado.extend(paginas[pagina])
    return resultado


# ============================================================
# 5. FILTRO DE SEGURANÇA POR ÓRGÃO (CNPJ)
# ============================================================
#
# CORREÇÃO IMPORTANTE: o endpoint /contratacoes/publicacao retorna
# registros de TODOS os órgãos do Brasil para a modalidade/período
# pesquisados — ele não filtra por CNPJ do órgão. A versão anterior
# deste arquivo tinha um laço que aparentava filtrar por CNPJ, mas
# não fazia nenhuma comparação de fato (copiava tudo sem checar),
# o que podia expor contratações de OUTROS municípios como se
# fossem de Rio das Pedras/SP. As funções abaixo implementam o
# filtro de verdade e são aplicadas a todos os módulos, inclusive
# Contratos e Atas, como camada extra de segurança caso o filtro
# do lado do servidor falhe silenciosamente.

def extrair_cnpj_orgao_registro(registro: Dict[str, Any]) -> str:
    candidatos: List[Any] = [
        registro.get("cnpjOrgao"),
        registro.get("cnpj"),
    ]
    orgao_entidade = registro.get("orgaoEntidade")
    if isinstance(orgao_entidade, dict):
        candidatos.append(orgao_entidade.get("cnpj"))
    unidade_orgao = registro.get("unidadeOrgao")
    if isinstance(unidade_orgao, dict):
        candidatos.append(unidade_orgao.get("cnpj"))
        orgao_aninhado = unidade_orgao.get("orgaoEntidade")
        if isinstance(orgao_aninhado, dict):
            candidatos.append(orgao_aninhado.get("cnpj"))
    for candidato in candidatos:
        documento = normalizar_documento(candidato)
        if documento:
            return documento
    return ""

def filtrar_registros_por_orgao(
    registros: List[Dict[str, Any]],
    cnpj_esperado: str,
) -> Tuple[List[Dict[str, Any]], int]:
    """Mantém apenas registros cujo CNPJ do órgão bate com o esperado.

    Registros em que não foi possível identificar o CNPJ do órgão são
    mantidos (para não perder dados silenciosamente), mas contados à
    parte para que a interface avise o usuário sobre a incerteza.
    """
    filtrados: List[Dict[str, Any]] = []
    sem_identificacao = 0
    for registro in registros:
        cnpj_registro = extrair_cnpj_orgao_registro(registro)
        if not cnpj_registro:
            sem_identificacao += 1
            filtrados.append(registro)
            continue
        if cnpj_registro == cnpj_esperado:
            filtrados.append(registro)
    return filtrados, sem_identificacao


# ============================================================
# 6. NORMALIZAÇÃO DE REGISTROS POR MÓDULO
# ============================================================

CAMPOS_POR_TIPO = {
    "Contratos": {
        "id_pncp": (
            "numeroControlePNCP",
            "idContratoPNCP",
        ),
        "numero": (
            "numeroContratoEmpenho",
            "numeroContrato",
            "numero",
        ),
        "processo": (
            "processo",
            "numeroProcesso",
        ),
        "objeto": (
            "objetoContrato",
            "objetoCompra",
            "objeto",
        ),
        "fornecedor": (
            "nomeRazaoSocialFornecedor",
            "razaoSocialFornecedor",
            "nomeFornecedor",
        ),
        "cnpj_fornecedor": (
            "niFornecedor",
            "cnpjFornecedor",
        ),
        "valor": (
            "valorGlobal",
            "valorInicial",
            "valorTotal",
            "valorContrato",
        ),
        "data": (
            "dataAssinatura",
            "dataCelebracao",
            "dataPublicacaoPncp",
        ),
        "situacao": (
            "situacao",
            "status",
        ),
        "modalidade": (
            "modalidadeNome",
            "modalidadeContratacaoNome",
        ),
    },
    "Atas de Registro de Preços": {
        "id_pncp": (
            "numeroControlePNCPAta",
            "numeroControlePNCP",
        ),
        "numero": (
            "numeroAtaRegistroPreco",
            "numeroAta",
            "numero",
        ),
        "processo": (
            "processo",
            "numeroProcesso",
        ),
        "objeto": (
            "objetoCompra",
            "objeto",
        ),
        "fornecedor": (
            "nomeRazaoSocialFornecedor",
            "razaoSocialFornecedor",
        ),
        "cnpj_fornecedor": (
            "niFornecedor",
            "cnpjFornecedor",
        ),
        "valor": (
            "valorTotal",
            "valorGlobal",
            "valorAta",
        ),
        "data": (
            "dataAssinatura",
            "dataPublicacaoPncp",
        ),
        "situacao": (
            "situacao",
            "status",
        ),
        "modalidade": (
            "modalidadeNome",
            "modalidadeContratacaoNome",
        ),
    },
    "Editais e Avisos de Contratações": {
        "id_pncp": (
            "numeroControlePNCP",
        ),
        "numero": (
            "numeroCompra",
            "numeroEdital",
            "numero",
        ),
        "processo": (
            "processo",
            "numeroProcesso",
        ),
        "objeto": (
            "objetoCompra",
            "objeto",
            "descricaoObjeto",
        ),
        "fornecedor": (
            "nomeRazaoSocialFornecedor",
            "razaoSocialFornecedor",
        ),
        "cnpj_fornecedor": (
            "niFornecedor",
            "cnpjFornecedor",
        ),
        "valor": (
            "valorTotalHomologado",
            "valorTotalEstimado",
            "valorTotal",
        ),
        "data": (
            "dataPublicacaoPncp",
            "dataPublicacao",
            "dataInclusao",
        ),
        "situacao": (
            "situacaoCompraNome",
            "situacaoCompra",
            "situacao",
        ),
        "modalidade": (
            "modalidadeNome",
            "modalidadeContratacaoNome",
        ),
    },
}

def extrair_dados_padrao(
    registro: Any,
    tipo: str,
) -> Dict[str, Any]:
    campos = CAMPOS_POR_TIPO[tipo]
    valor_bruto = obter_primeiro_valor(
        registro,
        campos["valor"],
    )
    documento_fornecedor = normalizar_documento(
        obter_primeiro_valor(
            registro,
            campos["cnpj_fornecedor"],
        )
    )
    return {
        "id_pncp": texto_valido(
            obter_primeiro_valor(registro, campos["id_pncp"])
        ),
        "numero": texto_valido(
            obter_primeiro_valor(registro, campos["numero"])
        ),
        "processo": texto_valido(
            obter_primeiro_valor(registro, campos["processo"])
        ),
        "objeto": texto_valido(
            obter_primeiro_valor(registro, campos["objeto"])
        ),
        "fornecedor": texto_valido(
            obter_primeiro_valor(registro, campos["fornecedor"])
        ),
        "cnpj_fornecedor": documento_fornecedor or "N/D",
        "valor_numerico": valor_numerico(valor_bruto),
        "valor": formatar_moeda_br(valor_bruto),
        "data": formatar_data(
            obter_primeiro_valor(registro, campos["data"])
        ),
        "situacao": texto_valido(
            obter_primeiro_valor(registro, campos["situacao"])
        ),
        "modalidade": texto_valido(
            obter_primeiro_valor(registro, campos["modalidade"])
        ),
    }

def normalizar_registros(
    registros: Iterable[Dict[str, Any]],
    tipo: str,
) -> pd.DataFrame:
    processados = [
        extrair_dados_padrao(registro, tipo)
        for registro in registros
    ]
    df = pd.DataFrame(processados)
    if df.empty:
        return df
    df = df.drop_duplicates(
        subset=["id_pncp", "numero"],
        keep="first",
    ).reset_index(drop=True)
    return df


# ============================================================
# 7. AVALIAÇÃO DE RISCO COERENTE
# ============================================================

def calcular_referencia_historica(
    df_historico: Optional[pd.DataFrame],
    id_pncp_atual: str,
) -> Optional[float]:
    if df_historico is None or df_historico.empty:
        return None
    if "valor_numerico" not in df_historico.columns:
        return None
    historico = df_historico.copy()
    if "id_pncp" in historico.columns:
        historico = historico[
            historico["id_pncp"] != id_pncp_atual
        ]
    valores = pd.to_numeric(
        historico["valor_numerico"],
        errors="coerce",
    )
    valores = valores[
        valores.notna()
        & (valores >= 0)
    ]
    if len(valores) < 4:
        return None
    return float(valores.median())

def avaliar_risco_contratacao(
    row: Any,
    df_historico: Optional[pd.DataFrame] = None,
) -> Dict[str, Any]:
    pontos = 0
    criterios: List[str] = []
    valor = valor_numerico(row.get("valor_numerico")) or 0.0
    objeto = texto_valido(row.get("objeto"), "").lower()
    modalidade = texto_valido(row.get("modalidade"), "").lower()
    id_pncp = texto_valido(row.get("id_pncp"), "")
    if valor > 1_000_000:
        pontos += 30
        criterios.append(
            "Materialidade elevada, superior a R$ 1 milhão"
        )
    elif valor > 300_000:
        pontos += 15
        criterios.append(
            "Materialidade intermediária relevante"
        )
    if any(
        termo in modalidade
        for termo in ("dispensa", "inexigibilidade")
    ):
        pontos += 25
        criterios.append(
            "Contratação direta por dispensa ou inexigibilidade"
        )
    if any(
        termo in objeto
        for termo in (
            "emergencial",
            "emergência",
            "calamidade",
        )
    ):
        pontos += 35
        criterios.append(
            "Objeto contém indicação de situação emergencial"
        )
    referencia = calcular_referencia_historica(
        df_historico,
        id_pncp,
    )
    if (
        referencia is not None
        and referencia > 0
        and valor > referencia * 3
    ):
        pontos += 25
        criterios.append(
            "Valor superior a três vezes a mediana dos demais registros "
            f"da amostra ({formatar_moeda_br(referencia)})"
        )
    if texto_valido(row.get("anomalia_ml"), "Normal") == "Anômalo":
        pontos += 15
        criterios.append(
            "Registro classificado como anômalo pelo modelo estatístico"
        )
    pontos = min(pontos, 100)
    if pontos >= 60:
        prioridade = "🔴 Alta"
        risco = "Alto"
    elif pontos >= 30:
        prioridade = "🟡 Média"
        risco = "Moderado"
    else:
        prioridade = "🟢 Baixa"
        risco = "Baixo"
    if not criterios:
        criterios.append(
            "Nenhum alerta automatizado identificado"
        )
    return {
        "pontos": pontos,
        "prioridade": prioridade,
        "risco_inerente": risco,
        "criterios": criterios,
    }


# ============================================================
# 8. MODELO DE ANOMALIAS ROBUSTO (ISOLATION FOREST)
# ============================================================

def executar_modelo_ml_anomalias(
    df: pd.DataFrame,
) -> pd.DataFrame:
    resultado = df.copy()
    resultado["anomalia_ml"] = "Não avaliado"
    resultado["score_anomalia"] = pd.NA
    if resultado.empty or len(resultado) < 20:
        return resultado
    resultado["tamanho_objeto"] = (
        resultado["objeto"]
        .fillna("")
        .astype(str)
        .str.len()
    )
    resultado["valor_ml"] = (
        pd.to_numeric(
            resultado["valor_numerico"],
            errors="coerce",
        )
        .fillna(0)
        .clip(lower=0)
        .map(lambda valor: math.log1p(valor))
    )
    features = resultado[
        ["valor_ml", "tamanho_objeto"]
    ]
    if features.nunique().max() <= 1:
        return resultado
    escalador = RobustScaler()
    dados_escalados = escalador.fit_transform(features)
    modelo = IsolationForest(
        n_estimators=300,
        contamination="auto",
        random_state=42,
        n_jobs=-1,
    )
    previsoes = modelo.fit_predict(dados_escalados)
    scores = -modelo.decision_function(dados_escalados)
    resultado["anomalia_ml"] = [
        "Anômalo" if previsao == -1 else "Normal"
        for previsao in previsoes
    ]
    resultado["score_anomalia"] = scores.round(4)
    return resultado.drop(
        columns=["valor_ml", "tamanho_objeto"],
        errors="ignore",
    )


# ============================================================
# 9. GERAÇÃO DE PAPEL DE TRABALHO DE AUDITORIA (WORD)
# ============================================================

def adicionar_item_documento(
    doc: Document,
    rotulo: str,
    valor: Any,
) -> None:
    paragrafo = doc.add_paragraph(style="List Bullet")
    rotulo_run = paragrafo.add_run(f"{rotulo}: ")
    rotulo_run.bold = True
    paragrafo.add_run(texto_valido(valor))

def gerar_papel_trabalho_auditoria_word(
    row: Any,
    tipo_consulta: str,
    avaliacao_risco: Dict[str, Any],
) -> bytes:
    doc = Document()
    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run(
        "PAPEL DE TRABALHO DE AUDITORIA E CONTROLE INTERNO"
    )
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 51, 102)
    doc.add_paragraph(f"Órgão fiscalizado: {NOME_PREFEITURA}")
    doc.add_paragraph(
        "Data da emissão: "
        f"{datetime.date.today().strftime('%d/%m/%Y')}"
    )
    doc.add_paragraph(f"Módulo: {tipo_consulta}")
    doc.add_heading("1. Identificação da contratação", level=2)
    adicionar_item_documento(doc, "Processo", row.get("processo"))
    adicionar_item_documento(doc, "Número", row.get("numero"))
    adicionar_item_documento(
        doc,
        "Identificador PNCP",
        row.get("id_pncp"),
    )
    adicionar_item_documento(doc, "Objeto", row.get("objeto"))
    adicionar_item_documento(
        doc,
        "Modalidade",
        row.get("modalidade"),
    )
    adicionar_item_documento(
        doc,
        "Fornecedor",
        row.get("fornecedor"),
    )
    adicionar_item_documento(
        doc,
        "CNPJ/CPF do fornecedor",
        row.get("cnpj_fornecedor"),
    )
    adicionar_item_documento(doc, "Valor", row.get("valor"))
    adicionar_item_documento(doc, "Data", row.get("data"))
    adicionar_item_documento(
        doc,
        "Situação",
        row.get("situacao"),
    )
    doc.add_heading("2. Matriz de risco", level=2)
    adicionar_item_documento(
        doc,
        "Prioridade",
        avaliacao_risco["prioridade"],
    )
    adicionar_item_documento(
        doc,
        "Pontuação",
        f"{avaliacao_risco['pontos']}/100",
    )
    adicionar_item_documento(
        doc,
        "Risco inerente",
        avaliacao_risco["risco_inerente"],
    )
    adicionar_item_documento(
        doc,
        "Anomalia estatística",
        row.get("anomalia_ml"),
    )
    doc.add_paragraph("Critérios identificados:")
    for criterio in avaliacao_risco["criterios"]:
        doc.add_paragraph(
            criterio,
            style="List Bullet 2",
        )
    doc.add_heading("3. Procedimentos sugeridos", level=2)
    procedimentos = (
        "Conferir o processo administrativo e a publicação no PNCP.",
        "Validar justificativas, pareceres e documentos de habilitação.",
        "Confirmar a compatibilidade do objeto, do valor e do fornecedor.",
        "Registrar as evidências documentais examinadas.",
    )
    for procedimento in procedimentos:
        doc.add_paragraph(
            procedimento,
            style="List Bullet",
        )
    doc.add_heading("4. Conclusão", level=2)
    doc.add_paragraph(
        "[ ] Sem ressalvas\n"
        "[ ] Com ressalvas\n"
        "[ ] Requer diligência complementar\n"
        "[ ] Encaminhar para apuração específica"
    )
    doc.add_paragraph(
        "\nObservação: a classificação automatizada constitui mecanismo "
        "de priorização e não representa, isoladamente, constatação de "
        "irregularidade."
    )
    doc.add_paragraph("\n\n________________________________________")
    doc.add_paragraph("Servidor responsável / Controlador interno")
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ============================================================
# 10. CONSULTA DE MODALIDADES SEM OMISSÃO
# ============================================================

MODALIDADES = {
    1: "Leilão eletrônico",
    2: "Diálogo competitivo",
    3: "Concurso",
    4: "Concorrência eletrônica",
    5: "Concorrência presencial",
    6: "Pregão eletrônico",
    7: "Pregão presencial",
    8: "Dispensa de licitação",
    9: "Inexigibilidade",
    10: "Manifestação de interesse",
    11: "Pré-qualificação",
    12: "Credenciamento",
    13: "Leilão presencial",
}

def consultar_registros(
    tipo: str,
    data_inicio: datetime.date,
    data_fim: datetime.date,
    modalidades: Optional[Sequence[int]] = None,
) -> Tuple[List[Dict[str, Any]], int]:
    """Retorna (registros_do_orgao, quantidade_sem_cnpj_identificado)."""
    parametros_base = {
        "dataInicial": data_inicio.strftime("%Y%m%d"),
        "dataFinal": data_fim.strftime("%Y%m%d"),
        "pagina": 1,
        "tamanhoPagina": TAMANHO_PAGINA,
    }
    if tipo == "Contratos":
        parametros = {
            **parametros_base,
            "cnpjOrgao": CNPJ_ORGAO,
        }
        brutos = consultar_todas_paginas(
            f"{BASE_URL}/contratos",
            parametros,
        )
        return filtrar_registros_por_orgao(brutos, CNPJ_ORGAO)

    if tipo == "Atas de Registro de Preços":
        parametros = {
            **parametros_base,
            "cnpj": CNPJ_ORGAO,
        }
        brutos = consultar_todas_paginas(
            f"{BASE_URL}/atas",
            parametros,
        )
        return filtrar_registros_por_orgao(brutos, CNPJ_ORGAO)

    # "Editais e Avisos de Contratações": o endpoint de publicação NÃO
    # aceita filtro por CNPJ do órgão, então a filtragem abaixo é a
    # única barreira contra misturar contratações de outros municípios.
    resultados: List[Dict[str, Any]] = []
    for codigo_modalidade in modalidades or (6, 8, 9):
        parametros = {
            **parametros_base,
            "codigoModalidadeContratacao": codigo_modalidade,
        }
        resultados.extend(
            consultar_todas_paginas(
                f"{BASE_URL}/contratacoes/publicacao",
                parametros,
            )
        )
        time.sleep(0.15)

    resultados_filtrados, sem_identificacao = filtrar_registros_por_orgao(
        resultados,
        CNPJ_ORGAO,
    )

    unicos: Dict[str, Dict[str, Any]] = {}
    for registro in resultados_filtrados:
        chave = texto_valido(
            registro.get("numeroControlePNCP"),
            "",
        )
        if not chave:
            chave = repr(
                (
                    registro.get("numeroCompra"),
                    registro.get("processo"),
                    registro.get("dataPublicacaoPncp"),
                )
            )
        unicos[chave] = registro
    return list(unicos.values()), sem_identificacao


# ============================================================
# 11. INTERFACE STREAMLIT
# ============================================================

st.title("🛡️ Sistema de Apoio ao Controle Interno - SACI")
st.markdown(
    "Plataforma de triagem analítica e auditoria preventiva das "
    f"contratações públicas de **{NOME_MUNICIPIO}**."
)
st.caption(
    "Fonte: Portal Nacional de Contratações Públicas - PNCP. "
    "Os alertas automatizados indicam prioridade de análise e não "
    "constituem prova de irregularidade."
)

st.sidebar.header("⚙️ Parâmetros de consulta")
tipo_consulta = st.sidebar.selectbox(
    "Módulo de análise",
    (
        "Contratos",
        "Atas de Registro de Preços",
        "Editais e Avisos de Contratações",
    ),
)
data_inicio = st.sidebar.date_input(
    "Data inicial",
    value=datetime.date(datetime.date.today().year, 1, 1),
)
data_fim = st.sidebar.date_input(
    "Data final",
    value=datetime.date.today(),
)

modalidades_selecionadas: List[int] = []
if tipo_consulta == "Editais e Avisos de Contratações":
    codigos_selecionados = st.sidebar.multiselect(
        "Modalidades",
        options=list(MODALIDADES.keys()),
        default=[6, 8, 9],
        format_func=lambda codigo: MODALIDADES[codigo],
    )
    modalidades_selecionadas = list(codigos_selecionados)

if data_fim < data_inicio:
    st.sidebar.error(
        "A data final não pode ser anterior à data inicial."
    )
    st.stop()

executar = st.sidebar.button(
    "🔎 Executar varredura",
    type="primary",
    use_container_width=True,
)

if executar:
    if (
        tipo_consulta == "Editais e Avisos de Contratações"
        and not modalidades_selecionadas
    ):
        st.error("Selecione ao menos uma modalidade.")
        st.stop()
    try:
        with st.spinner(
            "Consultando o PNCP e processando os registros..."
        ):
            registros, sem_identificacao = consultar_registros(
                tipo=tipo_consulta,
                data_inicio=data_inicio,
                data_fim=data_fim,
                modalidades=modalidades_selecionadas,
            )
            df_resultado = normalizar_registros(
                registros,
                tipo_consulta,
            )
            df_resultado = executar_modelo_ml_anomalias(
                df_resultado
            )
            st.session_state["df_resultado"] = df_resultado
            st.session_state["tipo_resultado"] = tipo_consulta
            st.session_state["sem_identificacao"] = sem_identificacao
    except Exception as erro:
        st.session_state.pop("df_resultado", None)
        st.error(f"Erro durante a consulta: {erro}")

df = st.session_state.get("df_resultado")

if df is None:
    st.info(
        "Configure os parâmetros e execute a varredura para carregar "
        "os dados."
    )
    st.stop()

sem_identificacao = st.session_state.get("sem_identificacao", 0)
if sem_identificacao:
    st.warning(
        f"⚠️ {sem_identificacao} registro(s) retornado(s) pelo PNCP não "
        "traziam o CNPJ do órgão de forma identificável na resposta da "
        "API e foram mantidos na lista sem confirmação de que pertencem "
        f"a {NOME_MUNICIPIO}. Recomenda-se conferência manual desses "
        "itens antes de qualquer conclusão."
    )

if df.empty:
    st.warning(
        "Nenhum registro foi encontrado para os parâmetros informados."
    )
    st.stop()

avaliacoes = [
    avaliar_risco_contratacao(row, df)
    for _, row in df.iterrows()
]

df_exibicao = df.copy()
df_exibicao["prioridade"] = [
    avaliacao["prioridade"]
    for avaliacao in avaliacoes
]
df_exibicao["risco_inerente"] = [
    avaliacao["risco_inerente"]
    for avaliacao in avaliacoes
]
df_exibicao["pontuacao"] = [
    avaliacao["pontos"]
    for avaliacao in avaliacoes
]
df_exibicao["motivos"] = [
    "; ".join(avaliacao["criterios"])
    for avaliacao in avaliacoes
]

st.divider()
st.subheader("📊 Painel executivo")

alta = int(
    df_exibicao["prioridade"]
    .str.contains("Alta", na=False)
    .sum()
)
media = int(
    df_exibicao["prioridade"]
    .str.contains("Média", na=False)
    .sum()
)
anomalias = int(
    (df_exibicao["anomalia_ml"] == "Anômalo").sum()
)

coluna_1, coluna_2, coluna_3, coluna_4 = st.columns(4)
coluna_1.metric("Total analisado", len(df_exibicao))
coluna_2.metric("🔴 Alta prioridade", alta)
coluna_3.metric("🟡 Média prioridade", media)
coluna_4.metric("🤖 Anomalias estatísticas", anomalias)

st.subheader("🚨 Matriz de risco")
colunas_tabela = {
    "prioridade": "Prioridade",
    "pontuacao": "Pontuação",
    "risco_inerente": "Risco",
    "anomalia_ml": "Anomalia",
    "numero": "Número",
    "processo": "Processo",
    "modalidade": "Modalidade",
    "objeto": "Objeto",
    "fornecedor": "Fornecedor",
    "valor": "Valor",
    "motivos": "Motivos",
}
st.dataframe(
    df_exibicao[list(colunas_tabela)].rename(
        columns=colunas_tabela
    ),
    use_container_width=True,
    hide_index=True,
)

st.download_button(
    "⬇️ Baixar resultado em CSV",
    data=df_exibicao.to_csv(
        index=False,
        sep=";",
    ).encode("utf-8-sig"),
    file_name="resultado_saci.csv",
    mime="text/csv",
    use_container_width=True,
)

st.divider()
st.subheader("📋 Papel de trabalho de auditoria")

opcoes = {
    indice: (
        f"[{row['prioridade']}] "
        f"{row['numero']} - {row['fornecedor']} - {row['valor']}"
    )
    for indice, row in df_exibicao.iterrows()
}
indice_selecionado = st.selectbox(
    "Selecione o registro",
    options=list(opcoes),
    format_func=lambda indice: opcoes[indice],
)

row_selecionada = df_exibicao.loc[indice_selecionado]
avaliacao_selecionada = avaliar_risco_contratacao(
    row_selecionada,
    df,
)

documento = gerar_papel_trabalho_auditoria_word(
    row_selecionada,
    st.session_state.get("tipo_resultado", tipo_consulta),
    avaliacao_selecionada,
)

numero_arquivo = re.sub(
    r"[^\w.-]+",
    "_",
    texto_valido(row_selecionada["numero"], "sem_numero"),
)

st.download_button(
    "⬇️ Baixar papel de trabalho (.docx)",
    data=documento,
    file_name=f"Papel_Trabalho_{numero_arquivo}.docx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document"
    ),
    use_container_width=True,
)
