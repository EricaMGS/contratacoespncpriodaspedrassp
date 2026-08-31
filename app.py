"""
Painel de Inteligência e Apoio ao Controle Interno - PNCP
Prefeitura Municipal de Rio das Pedras/SP

Tecnologias:
- Streamlit
- Pandas
- Requests
- Scikit-Learn (opcional)
- OpenPyXL

Objetivo:
Consulta, consolidação, segmentação, análise preliminar de risco,
similaridade semântica e exportação de dados públicos do PNCP.

IMPORTANTE:
Os alertas produzidos pelo sistema são indicadores automatizados.
Toda conclusão de auditoria deve ser validada por servidor responsável.
"""

from __future__ import annotations

import io
import logging
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ============================================================
# CONFIGURAÇÃO DE LOG
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

LOGGER = logging.getLogger(__name__)


# ============================================================
# DEPENDÊNCIAS OPCIONAIS
# ============================================================

try:
    from sklearn.feature_extraction.text import TfidfVectorizer

    SKLEARN_OK = True
except ImportError:
    TfidfVectorizer = None
    SKLEARN_OK = False

    LOGGER.warning(
        "scikit-learn não encontrado. "
        "A funcionalidade de similaridade semântica será desativada."
    )


# ============================================================
# CONFIGURAÇÕES GLOBAIS
# ============================================================

CNPJ = "44826840000183"
IBGE = "3544004"

MUNICIPIO = "Rio das Pedras/SP"
PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"

BASE_CONSULTA = "https://pncp.gov.br/api/consulta/v1"

MAX_WORKERS = max(
    1,
    min(
        16,
        int(os.getenv("PNCP_MAX_WORKERS", "6")),
    ),
)

TIMEOUT_CONNECT = max(
    3,
    int(os.getenv("PNCP_TIMEOUT_CONNECT", "10")),
)

TIMEOUT_READ = max(
    10,
    int(os.getenv("PNCP_TIMEOUT_READ", "45")),
)

TIMEOUT = (TIMEOUT_CONNECT, TIMEOUT_READ)

MAX_TENTATIVAS = max(
    1,
    int(os.getenv("PNCP_MAX_RETRIES", "4")),
)

PAGE_CONTRATOS = 100
PAGE_ATAS = 100
PAGE_EDITAIS = 50

# Limite de segurança para não disparar uma quantidade
# excessiva de páginas contra a API.
MAX_PAGINAS = max(
    1,
    int(os.getenv("PNCP_MAX_PAGES", "15")),
)

CACHE_TTL = 900

TIPOS = [
    "Contratos",
    "Atas de Registro de Preços",
    "Editais e Avisos de Contratações",
]

MODALIDADES_PNCP = tuple(range(1, 16))

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "Chrome/131.0 Safari/537.36 "
        "MonitoramentoControleInterno/5.0"
    ),
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "keep-alive",
}


# ============================================================
# ESTILOS
# ============================================================

def aplicar_estilo() -> None:
    """Aplica pequenos ajustes visuais sem alterar a identidade do app."""

    st.markdown(
        """
        <style>
        .stMetric {
            border-radius: 10px;
        }

        div[data-testid="stMetricValue"] {
            font-size: 1.55rem;
        }

        .risk-high {
            color: #b91c1c;
            font-weight: 700;
        }

        .risk-medium {
            color: #b45309;
            font-weight: 700;
        }

        .risk-low {
            color: #15803d;
            font-weight: 700;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# HTTP
# ============================================================

@st.cache_resource(show_spinner=False)
def sessao_http() -> requests.Session:
    """
    Cria uma sessão HTTP reutilizável.

    O retry foi limitado a situações transitórias.
    Erros como 400/404/422 não devem ser repetidos automaticamente.
    """

    session = requests.Session()
    session.headers.update(HEADERS)

    retry = Retry(
        total=MAX_TENTATIVAS,
        connect=MAX_TENTATIVAS,
        read=MAX_TENTATIVAS,
        status=MAX_TENTATIVAS,
        backoff_factor=1.5,
        status_forcelist=(408, 425, 429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
        respect_retry_after_header=True,
        raise_on_status=False,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=MAX_WORKERS,
        pool_maxsize=MAX_WORKERS,
    )

    session.mount("http://", adapter)
    session.mount("https://", adapter)

    return session


def detalhe_http(response: requests.Response) -> str:
    """Extrai uma mensagem útil da resposta HTTP."""

    try:
        payload = response.json()

        if isinstance(payload, dict):
            for key in (
                "message",
                "error",
                "detail",
                "mensagem",
                "erro",
            ):
                value = payload.get(key)

                if value is not None:
                    return str(value)[:500]

    except (ValueError, TypeError):
        pass

    text = (response.text or "").strip()

    return text[:500] if text else "Resposta sem detalhes."


def get_json(
    url: str,
    params: Optional[Dict[str, Any]] = None,
) -> Any:
    """
    Executa GET e retorna JSON.

    Erros são convertidos para RuntimeError com mensagens
    mais amigáveis para a camada Streamlit.
    """

    session = sessao_http()

    try:
        response = session.get(
            url,
            params=params,
            timeout=TIMEOUT,
        )

    except requests.exceptions.Timeout as exc:
        raise RuntimeError(
            "Tempo limite excedido ao consultar o PNCP."
        ) from exc

    except requests.exceptions.ConnectionError as exc:
        raise RuntimeError(
            "Não foi possível estabelecer conexão com o PNCP."
        ) from exc

    except requests.exceptions.RequestException as exc:
        raise RuntimeError(
            f"Erro de comunicação com o PNCP: {exc}"
        ) from exc

    if response.status_code == 204:
        return []

    if response.status_code == 200:
        try:
            return response.json()
        except ValueError as exc:
            raise RuntimeError(
                "O PNCP respondeu com conteúdo que não pôde ser interpretado como JSON."
            ) from exc

    detalhe = detalhe_http(response)

    if response.status_code == 429:
        raise RuntimeError(
            "O PNCP informou limite de requisições (HTTP 429). "
            "Aguarde alguns instantes e tente novamente."
        )

    if response.status_code in (500, 502, 503, 504):
        raise RuntimeError(
            f"O PNCP está temporariamente indisponível "
            f"(HTTP {response.status_code}). Detalhe: {detalhe}"
        )

    if response.status_code == 404:
        raise RuntimeError(
            f"Endpoint ou recurso não encontrado no PNCP (HTTP 404). "
            f"Detalhe: {detalhe}"
        )

    if response.status_code in (400, 422):
        raise RuntimeError(
            f"Parâmetros rejeitados pelo PNCP "
            f"(HTTP {response.status_code}): {detalhe}"
        )

    if response.status_code == 403:
        raise RuntimeError(
            "O PNCP recusou o acesso à consulta (HTTP 403)."
        )

    raise RuntimeError(
        f"PNCP respondeu HTTP {response.status_code}: {detalhe}"
    )


# ============================================================
# UTILITÁRIOS
# ============================================================

INVALID_TEXTS = {
    "",
    "none",
    "nan",
    "null",
    "n/d",
    "nat",
}


def valor_vazio(valor: Any) -> bool:
    """
    Determina se um valor pode ser considerado vazio.

    Importante:
    pd.isna(dict/list) pode retornar estruturas booleanas,
    então não devemos fazer simplesmente `if pd.isna(valor)`.
    """

    if valor is None:
        return True

    if isinstance(valor, (dict, list, tuple, set)):
        return len(valor) == 0

    try:
        resultado = pd.isna(valor)

        if isinstance(resultado, bool):
            return resultado

    except (TypeError, ValueError):
        pass

    return False


def texto(
    valor: Any,
    padrao: str = "N/D",
) -> str:
    """Converte um valor em texto seguro."""

    if valor_vazio(valor):
        return padrao

    if isinstance(valor, (dict, list, tuple, set)):
        return str(valor)

    resultado = str(valor).strip()

    if resultado.lower() in INVALID_TEXTS:
        return padrao

    return resultado


def numero(valor: Any) -> Optional[float]:
    """
    Converte números brasileiros e internacionais para float.

    Exemplos:
    1.234,56 -> 1234.56
    1234,56  -> 1234.56
    1234.56  -> 1234.56
    R$ 1.234,56 -> 1234.56
    """

    if valor_vazio(valor) or isinstance(valor, bool):
        return None

    if isinstance(valor, (int, float)):
        try:
            resultado = float(valor)

            if pd.isna(resultado):
                return None

            return resultado

        except (TypeError, ValueError):
            return None

    s = str(valor).strip()

    if not s:
        return None

    s = (
        s.replace("R$", "")
        .replace("r$", "")
        .replace(" ", "")
        .replace("\xa0", "")
    )

    # Trata formatos brasileiros:
    # 1.234,56
    # 1234,56
    # e formatos internacionais:
    # 1,234.56
    if "," in s and "." in s:

        ultima_virgula = s.rfind(",")
        ultimo_ponto = s.rfind(".")

        if ultima_virgula > ultimo_ponto:
            # Brasileiro
            s = s.replace(".", "")
            s = s.replace(",", ".")

        else:
            # Internacional
            s = s.replace(",", "")

    elif "," in s:
        partes = s.split(",")

        if len(partes) == 2 and len(partes[1]) <= 2:
            s = s.replace(",", ".")

        else:
            s = s.replace(",", "")

    try:
        resultado = float(s)

        if pd.isna(resultado):
            return None

        return resultado

    except (TypeError, ValueError):
        return None


def moeda(valor: Any) -> str:
    """Formata um número como moeda brasileira."""

    n = numero(valor)

    if n is None:
        return "N/D"

    return (
        f"R$ {n:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def data_br(valor: Any) -> str:
    """Converte data para dd/mm/aaaa."""

    if valor_vazio(valor):
        return "N/D"

    data = pd.to_datetime(
        valor,
        errors="coerce",
    )

    if pd.isna(data):
        return "N/D"

    return data.strftime("%d/%m/%Y")


def cnpj_limpo(valor: Any) -> str:
    """Mantém somente números de um CNPJ."""

    if valor_vazio(valor):
        return ""

    return re.sub(r"\D", "", str(valor))


def slug(valor: Any) -> str:
    """Gera nome seguro para arquivos."""

    s = re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        texto(valor, "registro"),
    )

    s = s.strip("_")

    return s[:100] or "registro"


def primeiro(
    row_dict: Dict[str, Any],
    campos: Sequence[str],
    padrao: Any = "N/D",
) -> Any:
    """Retorna o primeiro campo válido."""

    for campo in campos:
        if campo not in row_dict:
            continue

        valor = row_dict.get(campo)

        if valor_vazio(valor):
            continue

        if isinstance(valor, str):
            valor_limpo = valor.strip()

            if valor_limpo.lower() in INVALID_TEXTS:
                continue

            return valor_limpo

        return valor

    return padrao


def fuzzy_match(
    row_dict: Dict[str, Any],
    palavras: Sequence[str],
) -> Any:
    """
    Busca aproximada pelo nome do campo.

    Usada apenas como fallback quando os campos oficiais
    não foram encontrados.
    """

    for chave, valor in row_dict.items():

        if valor_vazio(valor):
            continue

        chave_lower = str(chave).lower()

        if not any(
            palavra.lower() in chave_lower
            for palavra in palavras
        ):
            continue

        if isinstance(valor, (dict, list)):
            continue

        valor_str = str(valor).strip()

        if valor_str.lower() in INVALID_TEXTS:
            continue

        return valor

    return None


# ============================================================
# EXTRAÇÃO DOS REGISTROS
# ============================================================

def dados_registro(
    row: Any,
    tipo: str,
) -> Dict[str, Any]:
    """Extrai os principais campos de um registro PNCP."""

    if hasattr(row, "to_dict"):
        r_dict = row.to_dict()
    else:
        r_dict = dict(row)

    controle = primeiro(
        r_dict,
        [
            "numeroControlePNCP",
            "numeroControlePNCPAta",
            "numeroControlePNCPCompra",
            "idContratoPNCP",
        ],
    )

    if tipo == "Contratos":

        num = primeiro(
            r_dict,
            [
                "numeroContratoEmpenho",
                "numeroContrato",
                "numero",
            ],
        )

        proc = primeiro(
            r_dict,
            [
                "processo",
                "numeroProcesso",
                "compra.processo",
            ],
        )

        obj = primeiro(
            r_dict,
            [
                "objetoContrato",
                "objetoCompra",
                "compra.objetoCompra",
                "objeto",
                "descricaoObjeto",
            ],
        )

        forn = primeiro(
            r_dict,
            [
                "nomeRazaoSocialFornecedor",
                "fornecedor.nomeRazaoSocial",
                "razaoSocialFornecedor",
                "nomeFornecedor",
            ],
        )

        cnpj = primeiro(
            r_dict,
            [
                "niFornecedor",
                "fornecedor.niFornecedor",
                "cnpjFornecedor",
                "cnpj",
            ],
        )

        val = primeiro(
            r_dict,
            [
                "valorGlobal",
                "valorInicial",
                "valorTotal",
                "valorContrato",
                "compra.valorTotalHomologado",
            ],
        )

        dt_ = primeiro(
            r_dict,
            [
                "dataAssinatura",
                "dataCelebracao",
                "dataPublicacaoPncp",
            ],
        )

        sit = primeiro(
            r_dict,
            [
                "situacao",
                "status",
            ],
        )

    elif tipo == "Atas de Registro de Preços":

        num = primeiro(
            r_dict,
            [
                "numeroAtaRegistroPreco",
                "numeroAta",
                "numero",
            ],
        )

        proc = primeiro(
            r_dict,
            [
                "processo",
                "numeroProcesso",
                "processoAdministrativo",
                "compra.processo",
            ],
        )

        obj = primeiro(
            r_dict,
            [
                "objetoAta",
                "objetoAtaRegistroPreco",
                "objetoCompra",
                "compra.objetoCompra",
                "objeto",
                "descricaoObjeto",
            ],
        )

        forn = primeiro(
            r_dict,
            [
                "nomeRazaoSocialFornecedor",
                "fornecedor.nomeRazaoSocial",
                "razaoSocialFornecedor",
                "nomeFornecedor",
            ],
        )

        cnpj = primeiro(
            r_dict,
            [
                "niFornecedor",
                "fornecedor.niFornecedor",
                "cnpjFornecedor",
                "ni",
            ],
        )

        val = primeiro(
            r_dict,
            [
                "valorTotalAta",
                "valorTotal",
                "valorGlobal",
                "valorAta",
                "compra.valorTotalHomologado",
            ],
        )

        dt_ = primeiro(
            r_dict,
            [
                "dataAssinatura",
                "dataPublicacaoPncp",
                "dataCelebracao",
            ],
        )

        sit = primeiro(
            r_dict,
            [
                "situacao",
                "status",
            ],
        )

    else:

        num = primeiro(
            r_dict,
            [
                "numeroCompra",
                "compra.numeroCompra",
                "numeroEdital",
                "numero",
            ],
        )

        proc = primeiro(
            r_dict,
            [
                "processo",
                "compra.processo",
                "numeroProcesso",
            ],
        )

        obj = primeiro(
            r_dict,
            [
                "objetoCompra",
                "compra.objetoCompra",
                "objeto",
                "descricaoObjeto",
            ],
        )

        forn = primeiro(
            r_dict,
            [
                "nomeRazaoSocialFornecedor",
                "fornecedor.nomeRazaoSocial",
                "razaoSocialFornecedor",
                "nomeFornecedor",
            ],
        )

        cnpj = primeiro(
            r_dict,
            [
                "niFornecedor",
                "fornecedor.niFornecedor",
                "cnpjFornecedor",
                "cnpjOrgao",
            ],
        )

        val = primeiro(
            r_dict,
            [
                "valorTotalHomologado",
                "compra.valorTotalHomologado",
                "valorTotalEstimado",
                "compra.valorTotalEstimado",
                "valorTotal",
            ],
        )

        dt_ = primeiro(
            r_dict,
            [
                "dataPublicacao",
                "dataPublicacaoPncp",
                "dataInclusao",
            ],
        )

        sit = primeiro(
            r_dict,
            [
                "situacaoCompra",
                "compra.situacaoCompraNome",
                "situacao",
                "status",
            ],
        )

    mod = primeiro(
        r_dict,
        [
            "modalidadeNome",
            "modalidadeContratacaoNome",
            "compra.modalidadeNome",
            "modalidade",
        ],
    )

    # --------------------------------------------------------
    # FALLBACKS
    # --------------------------------------------------------

    if texto(obj, "") == "":
        obj = fuzzy_match(
            r_dict,
            ["objeto", "descricao"],
        )

    if texto(forn, "") == "":
        forn = fuzzy_match(
            r_dict,
            [
                "razaosocial",
                "fornec",
                "fornecedor.nome",
            ],
        )

    if texto(val, "") == "":
        val = fuzzy_match(
            r_dict,
            [
                "valortotal",
                "valorglobal",
                "valorata",
                "valor",
            ],
        )

    if texto(mod, "") == "":
        mod = fuzzy_match(
            r_dict,
            [
                "modalidade",
                "tipoata",
            ],
        )

    if texto(cnpj, "") == "":
        cnpj = fuzzy_match(
            r_dict,
            [
                "cnpj",
                "ni",
            ],
        )

    return {
        "controle": texto(controle),
        "numero": texto(num),
        "processo": texto(proc),
        "objeto": texto(obj),
        "fornecedor": texto(forn),
        "cnpj_fornecedor": cnpj_limpo(cnpj),
        "valor_num": numero(val),
        "valor": moeda(val),
        "data": data_br(dt_),
        "situacao": texto(sit),
        "modalidade": texto(mod),
    }


# ============================================================
# PAGINAÇÃO
# ============================================================

def registros_api(data: Any) -> List[Dict[str, Any]]:
    """Extrai a lista de registros de diferentes formatos possíveis."""

    if isinstance(data, list):
        return [
            item
            for item in data
            if isinstance(item, dict)
        ]

    if not isinstance(data, dict):
        return []

    for chave in (
        "data",
        "items",
        "content",
        "dados",
        "registros",
    ):
        valor = data.get(chave)

        if isinstance(valor, list):
            return [
                item
                for item in valor
                if isinstance(item, dict)
            ]

    return []


def inteiro_seguro(valor: Any) -> Optional[int]:
    """Converte um valor para inteiro sem lançar exceção."""

    if valor_vazio(valor):
        return None

    try:
        return int(valor)
    except (TypeError, ValueError):
        return None


def paginacao(
    data: Any,
) -> Dict[str, Optional[int]]:
    """Extrai informações de paginação."""

    if not isinstance(data, dict):
        return {
            "totalPaginas": None,
            "totalRegistros": None,
            "numeroPagina": None,
        }

    return {
        "totalPaginas": inteiro_seguro(
            data.get("totalPaginas")
        ),
        "totalRegistros": inteiro_seguro(
            data.get("totalRegistros")
        ),
        "numeroPagina": inteiro_seguro(
            data.get("numeroPagina")
        ),
    }


def calcular_total_paginas(
    info: Dict[str, Optional[int]],
    tamanho_pagina: int,
) -> Optional[int]:
    """Determina o total de páginas."""

    total_paginas = info.get("totalPaginas")

    if total_paginas is not None:
        return max(1, total_paginas)

    total_registros = info.get("totalRegistros")

    if total_registros is None:
        return None

    tamanho = max(1, tamanho_pagina)

    return max(
        1,
        (total_registros + tamanho - 1) // tamanho,
    )


@st.cache_data(
    ttl=CACHE_TTL,
    show_spinner=False,
)
def consultar_cache(
    url: str,
    params_tuple: Tuple[Tuple[str, str], ...],
    max_paginas: int,
) -> Tuple[
    List[Dict[str, Any]],
    int,
    Optional[int],
    List[str],
]:
    """
    Consulta paginada com paralelismo.

    Retorna:
    registros,
    páginas consultadas,
    total de páginas informado pela API,
    erros encontrados.
    """

    base = dict(params_tuple)

    base["pagina"] = 1

    primeira = get_json(
        url,
        base,
    )

    registros_primeira = registros_api(
        primeira
    )

    info = paginacao(primeira)

    tamanho_pagina = inteiro_seguro(
        base.get("tamanhoPagina")
    ) or 50

    total = calcular_total_paginas(
        info,
        tamanho_pagina,
    )

    if not registros_primeira:
        return (
            [],
            1,
            total,
            [],
        )

    limite = min(
        max(1, max_paginas),
        total if total is not None else max_paginas,
    )

    todos = list(registros_primeira)

    if limite <= 1:
        return (
            todos,
            1,
            total,
            [],
        )

    resultados: Dict[
        int,
        List[Dict[str, Any]]
    ] = {}

    erros: List[str] = []

    workers = min(
        MAX_WORKERS,
        max(1, limite - 1),
    )

    def buscar_pagina(
        pagina: int,
    ) -> Tuple[
        int,
        List[Dict[str, Any]]
    ]:
        params = dict(base)
        params["pagina"] = pagina

        resposta = get_json(
            url,
            params,
        )

        return pagina, registros_api(
            resposta
        )

    with ThreadPoolExecutor(
        max_workers=workers
    ) as executor:

        futuros = {
            executor.submit(
                buscar_pagina,
                pagina,
            ): pagina
            for pagina in range(2, limite + 1)
        }

        for futuro in as_completed(futuros):

            pagina = futuros[futuro]

            try:
                pagina_resultado, registros = futuro.result()

                resultados[pagina_resultado] = registros

            except Exception as exc:
                mensagem = (
                    f"Página {pagina}: {exc}"
                )

                LOGGER.error(
                    "Erro na consulta: %s",
                    mensagem,
                )

                erros.append(mensagem)

    for pagina in range(2, limite + 1):
        todos.extend(
            resultados.get(
                pagina,
                [],
            )
        )

    return (
        todos,
        limite,
        total,
        erros,
    )


def consultar(
    url: str,
    params: Dict[str, Any],
    max_paginas: int,
) -> Tuple[
    List[Dict[str, Any]],
    int,
    Optional[int],
    List[str],
]:
    """
    Wrapper da consulta.

    A tupla é usada para tornar os parâmetros compatíveis
    com o cache do Streamlit.
    """

    serial = tuple(
        sorted(
            (
                str(chave),
                str(valor),
            )
            for chave, valor in params.items()
        )
    )

    return consultar_cache(
        url,
        serial,
        max_paginas,
    )


# ============================================================
# DATAFRAME
# ============================================================

def normalizar_pncp(
    registros: List[Dict[str, Any]],
) -> pd.DataFrame:
    """Normaliza JSON PNCP para DataFrame."""

    if not registros:
        return pd.DataFrame()

    df = pd.json_normalize(
        registros,
        sep=".",
    )

    if df.empty:
        return df

    for coluna in df.columns:

        def converter(valor: Any) -> Any:

            if isinstance(valor, (dict, list)):
                return str(valor)

            return valor

        df[coluna] = df[coluna].map(
            converter
        )

    return df


def chaves_identificacao(
    df: pd.DataFrame,
) -> List[str]:
    """Retorna chaves possíveis para identificação única."""

    return [
        coluna
        for coluna in (
            "numeroControlePNCP",
            "numeroControlePNCPAta",
            "numeroControlePNCPCompra",
            "idContratoPNCP",
        )
        if coluna in df.columns
    ]


def deduplicar(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Deduplicação progressiva.

    O código original retornava assim que encontrava uma chave.
    Isso podia deixar duplicidades baseadas em outra chave.
    """

    if df.empty:
        return df

    resultado = df.copy()

    chaves = chaves_identificacao(
        resultado
    )

    if chaves:

        mascara_manter = pd.Series(
            True,
            index=resultado.index,
        )

        identificados = pd.Series(
            False,
            index=resultado.index,
        )

        for coluna in chaves:

            serie = (
                resultado[coluna]
                .astype("string")
                .str.strip()
            )

            validos = (
                serie.notna()
                & ~serie.str.lower().isin(
                    {
                        "",
                        "nan",
                        "none",
                        "null",
                        "n/d",
                    }
                )
            )

            duplicados = (
                validos
                & serie.duplicated(
                    keep="first"
                )
            )

            mascara_manter &= ~(
                duplicados
                & ~identificados
            )

            identificados |= validos

        resultado = resultado.loc[
            mascara_manter
        ]

    else:
        resultado = resultado.drop_duplicates()

    return resultado.reset_index(
        drop=True
    )


def filtrar_cnpj(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Filtra o órgão pelo CNPJ.

    Tenta todas as colunas conhecidas antes de desistir.
    """

    if df.empty:
        return df

    alvo = cnpj_limpo(CNPJ)

    colunas = (
        "cnpjOrgao",
        "cnpj",
        "cnpjCompra",
        "orgaoEntidade.cnpj",
        "orgao.cnpj",
    )

    mascaras: List[pd.Series] = []

    for coluna in colunas:

        if coluna not in df.columns:
            continue

        serie = df[coluna].map(
            cnpj_limpo
        )

        mascara = serie == alvo

        if mascara.any():
            mascaras.append(
                mascara
            )

    if not mascaras:
        return df

    mascara_final = mascaras[0].copy()

    for mascara in mascaras[1:]:
        mascara_final |= mascara

    return df.loc[
        mascara_final
    ].reset_index(drop=True)


# ============================================================
# FILTRO DE DATA
# ============================================================

COLUNAS_DATA = (
    "dataPublicacaoPncp",
    "dataPublicacao",
    "dataInclusao",
    "dataAssinatura",
    "dataCelebracao",
)


def filtrar_periodo(
    df: pd.DataFrame,
    inicio: pd.Timestamp,
    fim: pd.Timestamp,
) -> pd.DataFrame:
    """
    Aplica filtro temporal.

    Tenta todas as colunas de data disponíveis.
    """

    if df.empty:
        return df

    colunas_disponiveis = [
        coluna
        for coluna in COLUNAS_DATA
        if coluna in df.columns
    ]

    if not colunas_disponiveis:
        return df

    mascaras = []

    for coluna in colunas_disponiveis:

        datas = pd.to_datetime(
            df[coluna],
            errors="coerce",
        )

        datas = datas.dt.normalize()

        mascara = (
            datas.notna()
            & (datas >= inicio)
            & (datas <= fim)
        )

        mascaras.append(
            mascara
        )

    mascara_final = mascaras[0].copy()

    for mascara in mascaras[1:]:
        mascara_final |= mascara

    return df.loc[
        mascara_final
    ].reset_index(drop=True)


# ============================================================
# MOTOR DE RISCO
# ============================================================

def limites_iqr(
    serie: pd.Series,
) -> Tuple[
    Optional[float],
    Optional[float],
    Optional[float],
]:
    """
    Calcula limites IQR.

    Retorna:
    limite inferior,
    limite superior,
    mediana.
    """

    valores = pd.to_numeric(
        serie,
        errors="coerce",
    ).dropna()

    if len(valores) < 5:
        return None, None, None

    q1 = valores.quantile(0.25)
    q3 = valores.quantile(0.75)

    iqr = q3 - q1

    mediana = valores.median()

    if iqr <= 0:
        return (
            float(q1),
            None,
            float(mediana),
        )

    limite_inferior = q1 - 1.5 * iqr
    limite_superior = q3 + 1.5 * iqr

    return (
        float(limite_inferior),
        float(limite_superior),
        float(mediana),
    )


def calcular_risco(
    registro: Dict[str, Any],
    contexto: pd.DataFrame,
    tipo: str,
) -> Dict[str, Any]:
    """
    Calcula um score heurístico de risco.

    O score não representa uma conclusão de auditoria.
    """

    pontos = 0

    motivos: List[str] = []
    testes: List[str] = []

    valor = registro.get(
        "valor_num"
    )

    modalidade = texto(
        registro.get("modalidade"),
        "",
    ).lower()

    objeto = texto(
        registro.get("objeto"),
        "",
    ).lower()

    # --------------------------------------------------------
    # COMPLETUDE
    # --------------------------------------------------------

    campos_obrigatorios = (
        ("objeto", "objeto"),
        ("controle", "controle PNCP"),
        ("data", "data"),
    )

    faltantes = [
        label
        for campo, label in campos_obrigatorios
        if texto(
            registro.get(campo),
            "",
        )
        in {"", "N/D"}
    ]

    if faltantes:

        pontos += min(
            12,
            4 * len(faltantes),
        )

        motivos.append(
            "Campos incompletos: "
            + ", ".join(faltantes)
        )

        testes.append(
            "Atenção na completude cadastral"
        )

    # --------------------------------------------------------
    # MODALIDADE / EXCEÇÃO
    # --------------------------------------------------------

    if (
        "dispensa" in modalidade
        or "inexig" in modalidade
    ):
        pontos += 20

        motivos.append(
            "Contratação por exceção"
        )

        testes.append(
            "Revisar fundamento legal"
        )

    elif (
        "emerg" in objeto
        or "emerg" in modalidade
    ):
        pontos += 25

        motivos.append(
            "Indício emergencial"
        )

        testes.append(
            "Revisar caracterização e prazo emergencial"
        )

    # --------------------------------------------------------
    # VALOR
    # --------------------------------------------------------

    if valor is not None:

        if valor >= 500_000:
            pontos += 10

            motivos.append(
                "Valor elevado (> R$ 500 mil)"
            )

        elif valor >= 150_000:
            pontos += 5

            motivos.append(
                "Valor relevante (> R$ 150 mil)"
            )

    # --------------------------------------------------------
    # OUTLIER
    # --------------------------------------------------------

    serie_valores = (
        contexto["valor_num"]
        if "valor_num" in contexto.columns
        else pd.Series(
            dtype="float64"
        )
    )

    _, limite_superior, _ = limites_iqr(
        serie_valores
    )

    outlier = False

    if (
        valor is not None
        and limite_superior is not None
        and valor > limite_superior
    ):
        outlier = True

        pontos += 25

        motivos.append(
            "Valor anômalo acima do limite IQR"
        )

        testes.append(
            "Avaliar formação de preços e pesquisa de mercado"
        )

    # --------------------------------------------------------
    # OBJETO CONTINUADO / ADITIVO
    # --------------------------------------------------------

    palavras_continuadas = (
        "prorroga",
        "aditivo",
        "continuado",
    )

    if any(
        palavra in objeto
        for palavra in palavras_continuadas
    ):
        pontos += 10

        motivos.append(
            "Indício de contrato continuado/aditivo"
        )

        testes.append(
            "Revisar vantajosidade e justificativa"
        )

    # --------------------------------------------------------
    # CONCENTRAÇÃO DE FORNECEDOR
    # --------------------------------------------------------

    fornecedor = texto(
        registro.get("fornecedor"),
        "",
    ).strip()

    if (
        fornecedor
        and not contexto.empty
        and "fornecedor" in contexto.columns
    ):

        fornecedores = (
            contexto["fornecedor"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
        )

        qtd = int(
            (
                fornecedores
                == fornecedor.casefold()
            ).sum()
        )

        limite_concentracao = max(
            3,
            int(
                len(contexto) * 0.10
            ),
        )

        if qtd >= limite_concentracao:

            pontos += 10

            motivos.append(
                f"Fornecedor concentrado ({qtd} ocorrências)"
            )

            testes.append(
                "Avaliar concentração e eventual direcionamento"
            )

    # --------------------------------------------------------
    # RESULTADO
    # --------------------------------------------------------

    pontos = max(
        0,
        min(
            100,
            int(pontos),
        ),
    )

    if pontos >= 60:
        nivel = "🔴 ALTO"

    elif pontos >= 30:
        nivel = "🟡 MÉDIO"

    else:
        nivel = "🟢 BAIXO"

    if not motivos:
        motivos = [
            "Sem alertas automáticos."
        ]

    return {
        "pontos": pontos,
        "nivel": nivel,
        "motivos": motivos,
        "testes": testes,
        "outlier": outlier,
    }


# ============================================================
# NLP / SIMILARIDADE
# ============================================================

def texto_nlp(valor: Any) -> str:
    """
    Normaliza texto para o TF-IDF.
    """

    s = texto(
        valor,
        "",
    ).lower()

    s = re.sub(
        r"\s+",
        " ",
        s,
    ).strip()

    if len(
        re.sub(
            r"[^a-zA-ZÀ-ÿ0-9]",
            "",
            s,
        )
    ) < 2:
        return "objeto ausente"

    return s


@st.cache_data(
    ttl=CACHE_TTL,
    show_spinner=False,
)
def calcular_modelo_tfidf(
    textos: Tuple[str, ...],
):
    """
    Cria a matriz TF-IDF.

    O cache evita recalcular o modelo a cada seleção.
    """

    if (
        not SKLEARN_OK
        or TfidfVectorizer is None
        or len(textos) < 2
    ):
        return None

    textos_seguros = tuple(
        texto_nlp(texto)
        for texto in textos
    )

    if len(
        set(textos_seguros)
    ) < 2:
        return None

    vetor = TfidfVectorizer(
        lowercase=True,
        strip_accents="unicode",
        ngram_range=(1, 2),
        min_df=1,
        max_features=8000,
        sublinear_tf=True,
    )

    try:
        return vetor.fit_transform(
            textos_seguros
        )

    except ValueError as exc:
        LOGGER.warning(
            "Não foi possível criar modelo TF-IDF: %s",
            exc,
        )
        return None


def similares(
    dados_processados: List[Dict[str, Any]],
    idx: int,
    limite: int = 5,
) -> pd.DataFrame:
    """
    Retorna registros semanticamente semelhantes.
    """

    if (
        not SKLEARN_OK
        or len(dados_processados) < 2
        or idx < 0
        or idx >= len(dados_processados)
    ):
        return pd.DataFrame()

    objetos = tuple(
        texto_nlp(
            registro.get("objeto")
        )
        for registro in dados_processados
    )

    matriz = calcular_modelo_tfidf(
        objetos
    )

    if matriz is None:
        return pd.DataFrame()

    try:
        scores = (
            matriz
            @ matriz[idx].T
        ).toarray().ravel()

    except (
        IndexError,
        ValueError,
    ) as exc:
        LOGGER.warning(
            "Erro no cálculo de similaridade: %s",
            exc,
        )
        return pd.DataFrame()

    ordem = scores.argsort()[::-1]

    saida = []

    for j in ordem:

        j = int(j)

        if j == idx:
            continue

        registro = dados_processados[j]

        saida.append(
            {
                "Similaridade": (
                    f"{float(scores[j]) * 100:.1f}%"
                ),
                "Número": texto(
                    registro.get("numero")
                ),
                "Objeto": texto(
                    registro.get("objeto")
                ),
                "Fornecedor": texto(
                    registro.get("fornecedor")
                ),
                "Valor": texto(
                    registro.get("valor")
                ),
            }
        )

        if len(saida) >= limite:
            break

    return pd.DataFrame(saida)


# ============================================================
# EXPORTAÇÃO EXCEL
# ============================================================

def ajustar_planilha_excel(
    worksheet: Any,
) -> None:
    """Formatação básica da planilha."""

    worksheet.freeze_panes = "A2"

    if worksheet.max_row >= 1:
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

    for coluna in worksheet.columns:

        if not coluna:
            continue

        letra = coluna[0].column_letter

        maior = 0

        for celula in coluna:

            valor = celula.value

            if valor is None:
                tamanho = 0
            else:
                tamanho = len(
                    str(valor)
                )

            maior = max(
                maior,
                tamanho,
            )

        worksheet.column_dimensions[
            letra
        ].width = min(
            55,
            max(
                12,
                maior + 2,
            ),
        )


def gerar_excel(
    df: pd.DataFrame,
    tipo: str,
    df_risco: pd.DataFrame,
) -> bytes:
    """Gera arquivo Excel para download."""

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl",
    ) as writer:

        df_risco.to_excel(
            writer,
            sheet_name="Matriz de Risco",
            index=False,
        )

        df.to_excel(
            writer,
            sheet_name="Dados Brutos",
            index=False,
        )

        for worksheet in writer.book.worksheets:
            ajustar_planilha_excel(
                worksheet
            )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# CONSULTAS ESPECÍFICAS
# ============================================================

def consultar_contratos(
    inicio: Any,
    fim: Any,
) -> Tuple[
    List[Dict[str, Any]],
    List[str],
]:
    """Consulta contratos."""

    params = {
        "dataInicial": inicio.strftime(
            "%Y%m%d"
        ),
        "dataFinal": fim.strftime(
            "%Y%m%d"
        ),
        "cnpjOrgao": CNPJ,
        "tamanhoPagina": PAGE_CONTRATOS,
    }

    registros, _, _, erros = consultar(
        f"{BASE_CONSULTA}/contratos",
        params,
        MAX_PAGINAS,
    )

    return registros, erros


def consultar_atas(
    inicio: Any,
    fim: Any,
) -> Tuple[
    List[Dict[str, Any]],
    List[str],
]:
    """Consulta atas de registro de preços."""

    params = {
        "dataInicial": inicio.strftime(
            "%Y%m%d"
        ),
        "dataFinal": fim.strftime(
            "%Y%m%d"
        ),
        "cnpj": CNPJ,
        "tamanhoPagina": PAGE_ATAS,
    }

    registros, _, _, erros = consultar(
        f"{BASE_CONSULTA}/atas",
        params,
        MAX_PAGINAS,
    )

    return registros, erros


def consultar_editais(
    inicio: Any,
    fim: Any,
    progress_callback: Any = None,
) -> Tuple[
    List[Dict[str, Any]],
    List[str],
]:
    """
    Consulta editais/contratações por modalidade.

    Mantém a estratégia do código original,
    mas centraliza o tratamento dos erros.
    """

    params_base = {
        "dataInicial": inicio.strftime(
            "%Y%m%d"
        ),
        "dataFinal": fim.strftime(
            "%Y%m%d"
        ),
        "cnpj": CNPJ,
        "tamanhoPagina": PAGE_EDITAIS,
    }

    todos_registros: List[
        Dict[str, Any]
    ] = []

    erros: List[str] = []

    total_modalidades = len(
        MODALIDADES_PNCP
    )

    for posicao, modalidade in enumerate(
        MODALIDADES_PNCP,
        start=1,
    ):

        if progress_callback:
            progress_callback(
                posicao / total_modalidades,
                modalidade,
            )

        params = dict(
            params_base
        )

        params[
            "codigoModalidadeContratacao"
        ] = modalidade

        try:

            registros, _, _, erros_pagina = consultar(
                f"{BASE_CONSULTA}/contratacoes/publicacao",
                params,
                MAX_PAGINAS,
            )

            if registros:
                todos_registros.extend(
                    registros
                )

            if erros_pagina:
                erros.extend(
                    [
                        f"Modalidade {modalidade}: {erro}"
                        for erro in erros_pagina
                    ]
                )

        except Exception as exc:

            mensagem = (
                f"Modalidade {modalidade}: {exc}"
            )

            LOGGER.warning(
                mensagem
            )

            erros.append(
                mensagem
            )

    return (
        todos_registros,
        erros,
    )


# ============================================================
# ESTADO DO STREAMLIT
# ============================================================

def inicializar_estado() -> None:
    """Inicializa session_state."""

    hoje = pd.Timestamp.today().date()

    defaults = {
        "df": None,
        "tipo": TIPOS[0],
        "inicio": dt_date(
            2026,
            1,
            1,
        ),
        "fim": hoje,
    }

    for chave, valor in defaults.items():

        if chave not in st.session_state:
            st.session_state[
                chave
            ] = valor


def dt_date(
    ano: int,
    mes: int,
    dia: int,
):
    """Pequeno helper para criação de date."""

    import datetime as dt

    return dt.date(
        ano,
        mes,
        dia,
    )


# ============================================================
# VALIDAÇÃO DE PERÍODO
# ============================================================

def validar_periodo(
    inicio: Any,
    fim: Any,
) -> Optional[str]:
    """Valida o período de consulta."""

    if inicio is None or fim is None:
        return "Informe as datas inicial e final."

    if inicio > fim:
        return (
            "A data inicial não pode ser posterior "
            "à data final."
        )

    return None


# ============================================================
# INTERFACE
# ============================================================

def main() -> None:

    st.set_page_config(
        page_title=(
            "Controle Interno — PNCP Rio das Pedras"
        ),
        page_icon="🛡️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    aplicar_estilo()
    inicializar_estado()

    # ========================================================
    # CABEÇALHO
    # ========================================================

    st.title(
        "🏛️ Inteligência de Controle Interno PNCP"
    )

    st.info(
        "ℹ️ Sistema de alertas automáticos. "
        "Os resultados são indicadores de apoio e "
        "devem ser validados por análise e auditoria humana."
    )

    # ========================================================
    # SIDEBAR
    # ========================================================

    with st.sidebar:

        st.header("⚙️ Parâmetros")

        tipo_selecionado = st.selectbox(
            "Escopo",
            TIPOS,
            index=TIPOS.index(
                st.session_state.tipo
            ),
        )

        if (
            tipo_selecionado
            != st.session_state.tipo
        ):

            st.session_state.tipo = (
                tipo_selecionado
            )

            st.session_state.df = None

        inicio = st.date_input(
            "📅 Data inicial",
            value=st.session_state.inicio,
        )

        fim = st.date_input(
            "📅 Data final",
            value=st.session_state.fim,
        )

        st.session_state.inicio = inicio
        st.session_state.fim = fim

        erro_periodo = validar_periodo(
            inicio,
            fim,
        )

        if erro_periodo:
            st.error(erro_periodo)

        buscar_btn = st.button(
            "🔎 Carregar dados",
            type="primary",
            use_container_width=True,
            disabled=erro_periodo is not None,
        )

        st.markdown("---")

        st.caption(
            f"Município: {MUNICIPIO}"
        )

        st.caption(
            f"CNPJ: {CNPJ}"
        )

        st.caption(
            f"IBGE: {IBGE}"
        )

        st.caption(
            f"Máx. páginas por consulta: {MAX_PAGINAS}"
        )

        if not SKLEARN_OK:
            st.warning(
                "Scikit-Learn não instalado. "
                "A análise semântica ficará desativada."
            )

    # ========================================================
    # CONSULTA
    # ========================================================

    if buscar_btn:

        try:

            todas_regs: List[
                Dict[str, Any]
            ] = []

            erros_consulta: List[
                str
            ] = []

            with st.spinner(
                "Consultando o PNCP..."
            ):

                tipo_consulta = (
                    tipo_selecionado
                )

                if tipo_consulta == "Contratos":

                    registros, erros = consultar_contratos(
                        inicio,
                        fim,
                    )

                    todas_regs.extend(
                        registros
                    )

                    erros_consulta.extend(
                        erros
                    )

                elif (
                    tipo_consulta
                    == "Atas de Registro de Preços"
                ):

                    registros, erros = consultar_atas(
                        inicio,
                        fim,
                    )

                    todas_regs.extend(
                        registros
                    )

                    erros_consulta.extend(
                        erros
                    )

                else:

                    barra = st.progress(
                        0,
                        text=(
                            "🔎 Preparando varredura "
                            "das modalidades..."
                        ),
                    )

                    def atualizar_progress(
                        progresso: float,
                        modalidade: int,
                    ) -> None:

                        barra.progress(
                            progresso,
                            text=(
                                "🔎 Varrendo modalidade "
                                f"PNCP cód. {modalidade}..."
                            ),
                        )

                    registros, erros = consultar_editais(
                        inicio,
                        fim,
                        atualizar_progress,
                    )

                    todas_regs.extend(
                        registros
                    )

                    erros_consulta.extend(
                        erros
                    )

                    barra.empty()

            # ------------------------------------------------
            # NORMALIZAÇÃO
            # ------------------------------------------------

            df = normalizar_pncp(
                todas_regs
            )

            df = deduplicar(
                df
            )

            if tipo_consulta != (
                "Editais e Avisos de Contratações"
            ):
                df = filtrar_cnpj(
                    df
                )

            # ------------------------------------------------
            # DATA
            # ------------------------------------------------

            inicio_ts = pd.Timestamp(
                inicio
            ).normalize()

            fim_ts = pd.Timestamp(
                fim
            ).normalize()

            df = filtrar_periodo(
                df,
                inicio_ts,
                fim_ts,
            )

            # ------------------------------------------------
            # ESTADO
            # ------------------------------------------------

            st.session_state.df = df
            st.session_state.tipo = (
                tipo_consulta
            )

            # ------------------------------------------------
            # RESULTADO DA CONSULTA
            # ------------------------------------------------

            if erros_consulta:

                st.warning(
                    "A consulta terminou com alguns avisos. "
                    "Verifique os detalhes abaixo."
                )

                with st.expander(
                    "⚠️ Detalhes técnicos da consulta"
                ):

                    for erro in erros_consulta[
                        :30
                    ]:
                        st.write(
                            f"- {erro}"
                        )

                    if len(erros_consulta) > 30:
                        st.caption(
                            f"{len(erros_consulta) - 30} "
                            "avisos adicionais foram omitidos."
                        )

            st.rerun()

        except Exception as exc:

            LOGGER.exception(
                "Erro durante consulta PNCP"
            )

            st.error(
                f"❌ Erro na consulta ao PNCP: {exc}"
            )

            st.stop()

    # ========================================================
    # DADOS
    # ========================================================

    df_bruto = st.session_state.df
    tipo_atual = st.session_state.tipo

    if df_bruto is None:

        st.info(
            "👈 Defina os parâmetros na barra lateral "
            "e clique em **Carregar dados**."
        )

        st.stop()

    if df_bruto.empty:

        st.warning(
            "Nenhum dado retornado para o filtro aplicado."
        )

        st.stop()

    # ========================================================
    # CONVERSÃO DOS REGISTROS
    # ========================================================

    dados_totais = [
        dados_registro(
            row,
            tipo_atual,
        )
        for row in df_bruto.to_dict(
            "records"
        )
    ]

    if not dados_totais:

        st.warning(
            "Não foi possível estruturar os registros retornados."
        )

        st.stop()

    # ========================================================
    # SEGMENTAÇÃO
    # ========================================================

    st.markdown("---")

    st.subheader(
        "🗂️ Segmentação Estratégica"
    )

    todas_modalidades = sorted(
        {
            d["modalidade"]
            for d in dados_totais
            if d["modalidade"]
            and d["modalidade"] != "N/D"
        }
    )

    opcao_geral = (
        "Visão Geral (Todas as Modalidades)"
    )

    if todas_modalidades:

        mod_escolhida = st.selectbox(
            "Filtrar e Segmentar por Modalidade:",
            [
                opcao_geral,
                *todas_modalidades,
            ],
        )

    else:

        mod_escolhida = opcao_geral

    if mod_escolhida != opcao_geral:

        indices_ativos = [
            i
            for i, registro in enumerate(
                dados_totais
            )
            if registro["modalidade"]
            == mod_escolhida
        ]

    else:

        indices_ativos = list(
            range(
                len(dados_totais)
            )
        )

    if not indices_ativos:

        st.warning(
            f"Sem registros para a modalidade: "
            f"{mod_escolhida}"
        )

        st.stop()

    df_filtrado = (
        df_bruto
        .iloc[indices_ativos]
        .reset_index(drop=True)
    )

    dados_processados = [
        dados_totais[i]
        for i in indices_ativos
    ]

    contexto = pd.DataFrame(
        dados_processados
    )

    riscos = [
        calcular_risco(
            registro,
            contexto,
            tipo_atual,
        )
        for registro in dados_processados
    ]

    # ========================================================
    # INDICADORES
    # ========================================================

    altos = sum(
        risco["nivel"] == "🔴 ALTO"
        for risco in riscos
    )

    medios = sum(
        risco["nivel"] == "🟡 MÉDIO"
        for risco in riscos
    )

    baixos = sum(
        risco["nivel"] == "🟢 BAIXO"
        for risco in riscos
    )

    if (
        "valor_num" in contexto.columns
        and not contexto["valor_num"].dropna().empty
    ):

        valor_total = contexto[
            "valor_num"
        ].sum()

    else:

        valor_total = None

    k1, k2, k3, k4 = st.columns(4)

    k1.metric(
        "Registros Segmentados",
        len(dados_processados),
    )

    k2.metric(
        "🔴 Alto risco",
        altos,
    )

    k3.metric(
        "🟡 Médio risco",
        medios,
    )

    k4.metric(
        "Valor Analisado",
        moeda(valor_total),
    )

    # ========================================================
    # MATRIZ DE RISCO
    # ========================================================

    rows = []

    for i_local, risco in enumerate(
        riscos
    ):

        registro = dados_processados[
            i_local
        ]

        rows.append(
            {
                "Índice": i_local,
                "Risco": risco["nivel"],
                "Score": risco["pontos"],
                "Número": registro["numero"],
                "Modalidade": registro[
                    "modalidade"
                ],
                "Fornecedor": registro[
                    "fornecedor"
                ],
                "Valor": registro[
                    "valor"
                ],
                "Objeto": registro[
                    "objeto"
                ],
                "Gatilhos": "; ".join(
                    risco["motivos"]
                ),
            }
        )

    df_risco = pd.DataFrame(
        rows
    )

    if not df_risco.empty:

        df_risco = df_risco.sort_values(
            [
                "Score",
                "Índice",
            ],
            ascending=[
                False,
                True,
            ],
        ).reset_index(
            drop=True
        )

    # ========================================================
    # ABAS
    # ========================================================

    st.markdown("---")

    tab1, tab2 = st.tabs(
        [
            "🚦 Matriz de Risco Segmentada",
            "📋 Auditoria e Semântica Individual",
        ]
    )

    # ========================================================
    # ABA 1
    # ========================================================

    with tab1:

        colunas_visualizacao = [
            coluna
            for coluna in (
                "Risco",
                "Score",
                "Número",
                "Modalidade",
                "Fornecedor",
                "Valor",
                "Objeto",
                "Gatilhos",
            )
            if coluna in df_risco.columns
        ]

        st.dataframe(
            df_risco[
                colunas_visualizacao
            ],
            use_container_width=True,
            hide_index=True,
        )

        try:

            excel = gerar_excel(
                df_filtrado,
                tipo_atual,
                df_risco,
            )

            nome_arquivo = (
                f"Auditoria_"
                f"{slug(mod_escolhida)}.xlsx"
            )

            st.download_button(
                "📊 Baixar Matriz Deste Segmento (Excel)",
                data=excel,
                file_name=nome_arquivo,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=False,
            )

        except Exception as exc:

            LOGGER.exception(
                "Erro ao gerar Excel"
            )

            st.error(
                f"Não foi possível gerar o Excel: {exc}"
            )

    # ========================================================
    # ABA 2
    # ========================================================

    with tab2:

        f1, f2 = st.columns(2)

        filtro_risco = f1.multiselect(
            "Nível de Risco Alvo",
            [
                "🔴 ALTO",
                "🟡 MÉDIO",
                "🟢 BAIXO",
            ],
            default=[
                "🔴 ALTO",
                "🟡 MÉDIO",
            ],
        )

        mostrar_outlier = f2.checkbox(
            "Apenas anomalias financeiras"
        )

        idx_auditoria = [
            i
            for i, risco in enumerate(
                riscos
            )
            if (
                risco["nivel"]
                in filtro_risco
                and (
                    not mostrar_outlier
                    or risco["outlier"]
                )
            )
        ]

        if not idx_auditoria:

            st.warning(
                "Ajuste os filtros de priorização "
                "acima para investigar um dossiê."
            )

        else:

            mapa = {}

            for i in idx_auditoria:

                risco = riscos[i]
                registro = dados_processados[i]

                chave = (
                    f"[{risco['nivel']} | "
                    f"Score {risco['pontos']}] "
                    f"{registro['numero']} — "
                    f"{registro['valor']}"
                )

                # Evita colisão quando dois registros
                # possuem exatamente a mesma descrição.
                if chave in mapa:

                    contador = 2

                    chave_base = chave

                    while chave in mapa:

                        chave = (
                            f"{chave_base} "
                            f"(#{contador})"
                        )

                        contador += 1

                mapa[chave] = i

            escolhido = st.selectbox(
                "Selecione o dossiê da contratação:",
                list(mapa.keys()),
            )

            i_alvo = mapa[
                escolhido
            ]

            registro = dados_processados[
                i_alvo
            ]

            risco = riscos[
                i_alvo
            ]

            # ------------------------------------------------
            # RESUMO
            # ------------------------------------------------

            with st.container():

                col_a, col_b, col_c = st.columns(
                    3
                )

                col_a.metric(
                    "Score de Risco",
                    f"{risco['pontos']} pts",
                )

                col_b.metric(
                    "Volume",
                    registro["valor"],
                )

                col_c.metric(
                    "Enquadramento Legal",
                    registro["modalidade"],
                )

                st.info(
                    f"**Objeto Descrito:** "
                    f"{registro['objeto']}"
                )

            # ------------------------------------------------
            # DETALHES
            # ------------------------------------------------

            with st.expander(
                "🔎 Justificativa Analítica e Testes",
                expanded=True,
            ):

                st.markdown(
                    "**Variáveis de Acionamento:**"
                )

                for motivo in risco[
                    "motivos"
                ]:
                    st.write(
                        f"- {motivo}"
                    )

                if risco["testes"]:

                    st.markdown(
                        "**Roteiro de Auditoria Sugerido:**"
                    )

                    for teste in risco[
                        "testes"
                    ]:
                        st.write(
                            f"- {teste}"
                        )

                else:

                    st.caption(
                        "Nenhum teste específico foi "
                        "sugerido automaticamente."
                    )

            # ------------------------------------------------
            # INFORMAÇÕES DO REGISTRO
            # ------------------------------------------------

            with st.expander(
                "📄 Informações do Registro",
                expanded=False,
            ):

                info_col1, info_col2 = st.columns(
                    2
                )

                info_col1.write(
                    f"**Controle PNCP:** "
                    f"{registro['controle']}"
                )

                info_col1.write(
                    f"**Número:** "
                    f"{registro['numero']}"
                )

                info_col1.write(
                    f"**Processo:** "
                    f"{registro['processo']}"
                )

                info_col1.write(
                    f"**Data:** "
                    f"{registro['data']}"
                )

                info_col2.write(
                    f"**Fornecedor:** "
                    f"{registro['fornecedor']}"
                )

                info_col2.write(
                    f"**CNPJ Fornecedor:** "
                    f"{registro['cnpj_fornecedor'] or 'N/D'}"
                )

                info_col2.write(
                    f"**Situação:** "
                    f"{registro['situacao']}"
                )

                info_col2.write(
                    f"**Modalidade:** "
                    f"{registro['modalidade']}"
                )

            # ------------------------------------------------
            # SEMÂNTICA
            # ------------------------------------------------

            st.markdown(
                "### 🤖 Clusterização Semântica "
                "(Apenas neste Segmento)"
            )

            if not SKLEARN_OK:

                st.warning(
                    "Scikit-Learn não está instalado. "
                    "Instale a dependência para habilitar "
                    "a análise semântica."
                )

            else:

                sim = similares(
                    dados_processados,
                    i_alvo,
                    limite=5,
                )

                if sim.empty:

                    st.info(
                        "Não foi possível identificar "
                        "vizinhança semântica relevante "
                        "para este item."
                    )

                else:

                    st.dataframe(
                        sim,
                        use_container_width=True,
                        hide_index=True,
                    )

    # ========================================================
    # RODAPÉ
    # ========================================================

    st.markdown("---")

    st.caption(
        "Painel de apoio ao Controle Interno — "
        f"{PREFEITURA}. "
        "Dados públicos consultados no PNCP. "
        "Alertas automatizados não substituem análise técnica, "
        "jurídica ou auditoria humana."
    )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    main()
