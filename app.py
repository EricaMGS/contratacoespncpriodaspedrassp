"""
Painel de Inteligência e Apoio ao Controle Interno - PNCP
Prefeitura Municipal de Rio das Pedras/SP

Tecnologias:
- Streamlit
- Pandas
- Requests / Urllib3
- Scikit-Learn (opcional, para clusterização semântica TF-IDF)
- OpenPyXL
- python-docx (opcional, para exportação executiva e dossiês em Word)

Objetivo:
Consulta multithread otimizada, consolidação, segmentação, análise preliminar de risco
com cálculo estatístico pré-computado (O(N)), monitoramento de vigência, busca ativa de
termos aditivos, auditoria de concentração de fornecedores, detecção de possível
fracionamento de despesa e exportação (Excel e Word).

IMPORTANTE:
Os alertas produzidos pelo sistema são indicadores automatizados.
Toda conclusão de auditoria deve ser validada por servidor responsável.
"""

from __future__ import annotations

import datetime as dt
import io
import logging
import os
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    DOCX_OK = True
except ImportError:
    Document = None
    WD_ALIGN_PARAGRAPH = None
    WD_TABLE_ALIGNMENT = None
    Pt = None
    DOCX_OK = False

# ============================================================
# CONFIGURAÇÃO DE LOG
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

LOGGER = logging.getLogger(__name__)

# ============================================================
# DEPENDÊNCIAS OPCIONAIS (NLP)
# ============================================================

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity

    SKLEARN_OK = True
except ImportError:
    TfidfVectorizer = None
    cosine_similarity = None
    SKLEARN_OK = False
    LOGGER.warning(
        "scikit-learn não encontrado. "
        "A funcionalidade de similaridade semântica será desativada."
    )

# ============================================================
# CONFIGURAÇÕES GLOBAIS
# ============================================================

CNPJ_PADRAO = "44826840000183"
IBGE_PADRAO = "3544004"

MUNICIPIO = "Rio das Pedras/SP"
PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"

BASE_CONSULTA = "https://pncp.gov.br/api/consulta/v1"
BASE_SERVICOS = "https://pncp.gov.br/api/pncp/v1"

MAX_WORKERS = max(
    2,
    min(
        16,
        int(os.getenv("PNCP_MAX_WORKERS", "8")),
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

MAX_PAGINAS = max(
    1,
    int(os.getenv("PNCP_MAX_PAGES", "15")),
)

CACHE_TTL = 900
CACHE_TERMOS_TTL = 3600

TIPOS = [
    "Contratos",
    "Atas de Registro de Preços",
    "Editais e Avisos de Contratações",
]

MODALIDADES_PNCP_NOMES: Dict[int, str] = {
    1: "1 - Leilão Eletrônico",
    2: "2 - Diálogo Competitivo",
    3: "3 - Concurso",
    4: "4 - Concorrência Eletrônica",
    5: "5 - Concorrência Presencial",
    6: "6 - Pregão Eletrônico",
    7: "7 - Pregão Presencial",
    8: "8 - Dispensa de Licitação",
    9: "9 - Inexigibilidade",
    10: "10 - Manifestação de Interesse",
    11: "11 - Pré-qualificação",
    12: "12 - Credenciamento",
    13: "13 - Leilão Presencial",
    14: "14 - Inaplicabilidade da Licitação",
    15: "15 - Chamada Pública",
}

MODALIDADES_PNCP = tuple(MODALIDADES_PNCP_NOMES.keys())

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "Chrome/131.0 Safari/537.36 "
        "MonitoramentoControleInterno/7.0"
    ),
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "keep-alive",
}

# Tabela pré-computada O(1) para remoção de acentos em português
TABELA_SEM_ACENTOS = str.maketrans(
    "áàãâäéèêëíìîïóòõôöúùûüçÁÀÃÂÄÉÈÊËÍÌÎÏÓÒÕÔÖÚÙÛÜÇ",
    "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC",
)


# ============================================================
# ESTILOS
# ============================================================

def aplicar_estilo() -> None:
    """Aplica ajustes visuais corporativos para o painel."""
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
# HTTP E COMUNICAÇÃO COM O PNCP
# ============================================================

@st.cache_resource(show_spinner=False)
def sessao_http() -> requests.Session:
    """Cria uma sessão HTTP reutilizável com pool ampliado para multithreading."""
    session = requests.Session()
    session.headers.update(HEADERS)

    retry = Retry(
        total=MAX_TENTATIVAS,
        connect=MAX_TENTATIVAS,
        read=MAX_TENTATIVAS,
        status=MAX_TENTATIVAS,
        backoff_factor=1.2,
        status_forcelist=(408, 425, 429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
        respect_retry_after_header=True,
        raise_on_status=False,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=MAX_WORKERS * 2,
        pool_maxsize=MAX_WORKERS * 4,
    )

    session.mount("http://", adapter)
    session.mount("https://", adapter)

    return session


def detalhe_http(response: requests.Response) -> str:
    """Extrai uma mensagem útil da resposta HTTP."""
    try:
        payload = response.json()
        if isinstance(payload, dict):
            for key in ("message", "error", "detail", "mensagem", "erro"):
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
    """Executa GET e retorna JSON."""
    session = sessao_http()

    try:
        response = session.get(
            url,
            params=params,
            timeout=TIMEOUT,
        )
    except requests.exceptions.Timeout as exc:
        raise RuntimeError("Tempo limite excedido ao consultar o PNCP.") from exc
    except requests.exceptions.ConnectionError as exc:
        raise RuntimeError("Não foi possível estabelecer conexão com o PNCP.") from exc
    except requests.exceptions.RequestException as exc:
        raise RuntimeError(f"Erro de comunicação com o PNCP: {exc}") from exc

    if response.status_code in (204, 404):
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

    if response.status_code in (400, 422):
        raise RuntimeError(
            f"Parâmetros rejeitados pelo PNCP "
            f"(HTTP {response.status_code}): {detalhe}"
        )

    if response.status_code == 403:
        raise RuntimeError("O PNCP recusou o acesso à consulta (HTTP 403).")

    raise RuntimeError(f"PNCP respondeu HTTP {response.status_code}: {detalhe}")


# ============================================================
# UTILITÁRIOS DE DADOS E FORMATAÇÃO
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
    """Determina se um valor pode ser considerado vazio."""
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
    """Converte números brasileiros e internacionais para float."""
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

    if "," in s and "." in s:
        ultima_virgula = s.rfind(",")
        ultimo_ponto = s.rfind(".")
        if ultima_virgula > ultimo_ponto:
            s = s.replace(".", "").replace(",", ".")
        else:
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

    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return "N/D"

    return data.strftime("%d/%m/%Y")


def data_timestamp(valor: Any) -> Optional[pd.Timestamp]:
    """Converte valor para pd.Timestamp seguro."""
    if valor_vazio(valor):
        return None
    ts = pd.to_datetime(valor, errors="coerce")
    return ts.normalize() if pd.notna(ts) else None


def cnpj_limpo(valor: Any) -> str:
    """Mantém somente números de um CNPJ/CPF."""
    if valor_vazio(valor):
        return ""
    return re.sub(r"\D", "", str(valor))


def cnpj_formatado(valor: Any) -> str:
    """Formata um CNPJ (14 dígitos) ou CPF (11 dígitos) para leitura amigável."""
    limpo = cnpj_limpo(valor)
    if len(limpo) == 14:
        return f"{limpo[:2]}.{limpo[2:5]}.{limpo[5:8]}/{limpo[8:12]}-{limpo[12:]}"
    if len(limpo) == 11:
        return f"{limpo[:3]}.{limpo[3:6]}.{limpo[6:9]}-{limpo[9:]}"
    return limpo or "N/D"


def slug(valor: Any) -> str:
    """Gera nome seguro para arquivos."""
    s = unicodedata.normalize("NFKD", texto(valor, "registro"))
    s = s.encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Za-z0-9_-]+", "_", s).strip("_")
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
    padrao: Any = None,
) -> Any:
    """Busca aproximada pelo nome do campo como fallback."""
    palavras_lower = [p.lower() for p in palavras]
    for chave, valor in row_dict.items():
        if valor_vazio(valor) or isinstance(valor, (dict, list)):
            continue

        chave_lower = str(chave).lower()
        if not any(p in chave_lower for p in palavras_lower):
            continue

        valor_str = str(valor).strip()
        if valor_str.lower() in INVALID_TEXTS:
            continue

        return valor

    return padrao


# ============================================================
# EXTRAÇÃO DOS REGISTROS COM BUSCA ATIVA DE TERMOS ADITIVOS
# ============================================================

def registros_api(data: Any) -> List[Dict[str, Any]]:
    """Extrai a lista de registros de diferentes formatos possíveis da API."""
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]

    if not isinstance(data, dict):
        return []

    for chave in ("data", "items", "content", "dados", "registros"):
        valor = data.get(chave)
        if isinstance(valor, list):
            return [item for item in valor if isinstance(item, dict)]

    return []


def _consultar_termos_http(
    cnpj: str,
    ano_int: int,
    seq_int: int,
) -> Dict[str, Any]:
    """Executa a chamada HTTP real ao sub-recurso de termos aditivos (thread-safe)."""
    url = f"{BASE_SERVICOS}/orgaos/{cnpj}/contratos/{ano_int}/{seq_int}/termos"
    resultado: Dict[str, Any] = {
        "status": "ERRO_CONSULTA",
        "termos": [],
        "http_status": None,
        "mensagem": "",
        "url": url,
    }

    session = sessao_http()

    try:
        response = session.get(url, timeout=TIMEOUT)
    except requests.exceptions.Timeout:
        resultado["mensagem"] = "Tempo limite excedido ao consultar os termos no PNCP."
        return resultado
    except requests.exceptions.ConnectionError:
        resultado["mensagem"] = "Não foi possível estabelecer conexão com o PNCP."
        return resultado
    except requests.exceptions.RequestException as exc:
        resultado["mensagem"] = f"Erro de comunicação com o PNCP: {exc}"
        return resultado

    resultado["http_status"] = response.status_code

    if response.status_code == 200:
        try:
            payload = response.json()
        except ValueError:
            resultado["status"] = "RESPOSTA_INVALIDA"
            resultado["mensagem"] = (
                "O PNCP respondeu HTTP 200, mas o conteúdo não pôde ser "
                "interpretado como JSON."
            )
            return resultado

        termos = registros_api(payload)
        if not termos:
            resultado["status"] = "SEM_REGISTROS"
            resultado["mensagem"] = (
                "Consulta realizada com sucesso. O PNCP não retornou "
                "termos vinculados a este contrato."
            )
            return resultado

        resultado["status"] = "REGISTROS_ENCONTRADOS"
        resultado["termos"] = termos
        resultado["mensagem"] = f"{len(termos)} termo(s) localizado(s) no PNCP."
        return resultado

    if response.status_code == 204:
        resultado["status"] = "SEM_REGISTROS"
        resultado["mensagem"] = (
            "O PNCP respondeu HTTP 204 (sem conteúdo). Nenhum termo foi retornado."
        )
        return resultado

    if response.status_code == 404:
        resultado["status"] = "HTTP_404"
        resultado["mensagem"] = (
            "O PNCP respondeu HTTP 404 para o recurso de termos. "
            "Isso não deve ser interpretado automaticamente como ausência de aditivos."
        )
        return resultado

    if response.status_code == 429:
        resultado["status"] = "ERRO_CONSULTA"
        resultado["mensagem"] = (
            "O PNCP informou limite de requisições (HTTP 429). "
            "A consulta dos termos não pôde ser concluída."
        )
        return resultado

    if response.status_code in (500, 502, 503, 504):
        resultado["status"] = "ERRO_CONSULTA"
        resultado["mensagem"] = (
            f"O PNCP apresentou indisponibilidade temporária "
            f"(HTTP {response.status_code})."
        )
        return resultado

    resultado["status"] = "ERRO_CONSULTA"
    resultado["mensagem"] = (
        f"O PNCP respondeu HTTP {response.status_code}: {detalhe_http(response)}"
    )
    return resultado


@st.cache_data(ttl=CACHE_TERMOS_TTL, show_spinner=False)
def buscar_termos_contrato_cache(
    cnpj: str,
    ano_int: int,
    seq_int: int,
) -> Dict[str, Any]:
    """Cacheia resultados conclusivos da consulta de termos."""
    return _consultar_termos_http(cnpj, ano_int, seq_int)


def buscar_termos_contrato(
    cnpj_orgao: str,
    ano: Any,
    sequencial: Any,
    consultar_api: bool = True,
) -> Dict[str, Any]:
    """Consulta os termos vinculados a um contrato e preserva o diagnóstico real."""
    resultado: Dict[str, Any] = {
        "status": "SEM_IDENTIFICADORES",
        "termos": [],
        "http_status": None,
        "mensagem": "",
        "url": "",
    }

    if not consultar_api:
        resultado["status"] = "NAO_CONSULTADO"
        resultado["mensagem"] = (
            "Busca ativa de termos na API desativada pelo usuário (modo rápido)."
        )
        return resultado

    if valor_vazio(ano) or valor_vazio(sequencial):
        resultado["mensagem"] = (
            "O contrato não possui anoContrato e/ou sequencialContrato "
            "disponíveis para consulta dos termos."
        )
        return resultado

    cnpj = cnpj_limpo(cnpj_orgao)
    if not cnpj:
        resultado["mensagem"] = "CNPJ do órgão não disponível para consulta."
        return resultado

    try:
        ano_int = int(float(str(ano).strip()))
        seq_int = int(float(str(sequencial).strip()))
    except (TypeError, ValueError):
        resultado["mensagem"] = (
            f"Identificadores inválidos: ano={ano!r}, sequencial={sequencial!r}."
        )
        return resultado

    return buscar_termos_contrato_cache(cnpj, ano_int, seq_int)


def dados_registro(
    row: Any,
    tipo: str,
    cnpj_orgao: str = CNPJ_PADRAO,
    consultar_termos_api: bool = True,
    hoje_ts: Optional[pd.Timestamp] = None,
) -> Dict[str, Any]:
    """Extrai os campos padronizados de um registro do PNCP."""
    if hasattr(row, "to_dict"):
        r_dict = row.to_dict()
    else:
        r_dict = dict(row)

    hoje = hoje_ts if hoje_ts is not None else pd.Timestamp.today().normalize()

    controle = primeiro(
        r_dict,
        [
            "numeroControlePNCP",
            "numeroControlePncp",
            "numeroControlePNCPAta",
            "numeroControlePNCPCompra",
            "idContratoPNCP",
        ],
        padrao="",
    )
    if not controle or controle == "N/D":
        controle = fuzzy_match(r_dict, ["controlepncp", "numerocontrole"], padrao="")

    ano_contrato = primeiro(
        r_dict,
        ["anoContrato", "ano", "compra.anoCompra"],
        padrao=None,
    )
    seq_contrato = primeiro(
        r_dict,
        ["sequencialContrato", "sequencial", "numeroSequencial"],
        padrao=None,
    )

    # ========================================================
    # EXTRAÇÃO ROBUSTA DO LINK PÚBLICO DO PNCP
    # ========================================================
    ano_url = None
    seq_url = None
    seq_ata_url = None
    controle_str = texto(controle, "").strip()

    if tipo == "Atas de Registro de Preços":
        match_ata = re.search(
            r"^(\d{14})-1-(\d+)/(\d{4})-(\d+)$",
            controle_str,
        )
        if match_ata:
            seq_url = str(int(match_ata.group(2)))
            ano_url = str(int(match_ata.group(3)))
            seq_ata_url = str(int(match_ata.group(4)))

    ano_fb = primeiro(
        r_dict,
        [
            "anoCompra",
            "anoAta",
            "anoContrato",
            "ano",
            "compra.anoCompra",
            "ata.anoCompra",
        ],
        padrao=None,
    )
    if valor_vazio(ano_fb):
        ano_fb = fuzzy_match(r_dict, ["anocompra", "anoata", "ano"], padrao=None)

    if tipo == "Atas de Registro de Preços":
        seq_fb = primeiro(
            r_dict,
            [
                "sequencialCompra",
                "compra.sequencialCompra",
                "sequencialAtaCompra",
                "sequencial",
            ],
            padrao=None,
        )
        seq_ata_fb = primeiro(
            r_dict,
            [
                "sequencialAta",
                "ata.sequencialAta",
                "numeroSequencialAta",
                "sequencialAtaRegistroPreco",
            ],
            padrao=None,
        )

        if valor_vazio(seq_fb):
            seq_fb = fuzzy_match(r_dict, ["sequencialcompra", "sequencial"], padrao=None)
        if valor_vazio(seq_ata_fb):
            seq_ata_fb = fuzzy_match(
                r_dict,
                ["sequencialata", "sequencialataregistro"],
                padrao=None,
            )

        try:
            if valor_vazio(ano_url) and not valor_vazio(ano_fb):
                ano_url = str(int(float(str(ano_fb).strip())))
            if valor_vazio(seq_url) and not valor_vazio(seq_fb):
                seq_url = str(int(float(str(seq_fb).strip())))
            if valor_vazio(seq_ata_url) and not valor_vazio(seq_ata_fb):
                seq_ata_url = str(int(float(str(seq_ata_fb).strip())))
        except (ValueError, TypeError):
            pass
    else:
        seq_fb = primeiro(
            r_dict,
            [
                "sequencialContrato" if tipo == "Contratos" else "sequencialCompra",
                "sequencial",
                "numeroSequencial",
                "compra.sequencialCompra",
            ],
            padrao=None,
        )
        if valor_vazio(seq_fb):
            seq_fb = fuzzy_match(r_dict, ["sequencial"], padrao=None)

        match_controle = re.search(r"-(\d+)/(\d{4})$", controle_str)
        if match_controle:
            seq_url = str(int(match_controle.group(1)))
            ano_url = str(int(match_controle.group(2)))
        else:
            try:
                if not valor_vazio(ano_fb):
                    ano_url = str(int(float(str(ano_fb).strip())))
                if not valor_vazio(seq_fb):
                    seq_url = str(int(float(str(seq_fb).strip())))
            except (ValueError, TypeError):
                pass

    link_pncp = "https://pncp.gov.br"
    cnpj_url = cnpj_limpo(cnpj_orgao)

    if tipo == "Contratos" and ano_url and seq_url:
        link_pncp = f"https://pncp.gov.br/app/contratos/{cnpj_url}/{ano_url}/{seq_url}"
    elif tipo == "Atas de Registro de Preços" and ano_url and seq_url and seq_ata_url:
        link_pncp = (
            f"https://pncp.gov.br/app/atas/{cnpj_url}/{ano_url}/{seq_url}/{seq_ata_url}"
        )
    elif tipo == "Editais e Avisos de Contratações" and ano_url and seq_url:
        link_pncp = f"https://pncp.gov.br/app/editais/{cnpj_url}/{ano_url}/{seq_url}"

    tipo_instrumento = primeiro(
        r_dict,
        [
            "tipoContrato.nome",
            "tipoContratoNome",
            "tipoTermoContrato.nome",
            "tipoTermoContratoNome",
            "tipoDocumento.nome",
        ],
        padrao="Contrato" if tipo == "Contratos" else tipo,
    )

    tipo_contrato_id = primeiro(
        r_dict,
        [
            "tipoContrato.id",
            "tipoContratoId",
            "tipoTermoContrato.id",
            "tipoTermoContratoId",
        ],
        padrao=None,
    )

    dt_fim_vig = None
    val_inicial = None
    val_global = None

    if tipo == "Contratos":
        num = primeiro(r_dict, ["numeroContratoEmpenho", "numeroContrato", "numero"])
        proc = primeiro(r_dict, ["processo", "numeroProcesso", "compra.processo"])
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
        val_global = primeiro(
            r_dict,
            [
                "valorGlobal",
                "valorTotal",
                "valorContrato",
                "compra.valorTotalHomologado",
            ],
        )
        val_inicial = primeiro(
            r_dict,
            ["valorInicial", "valorTotalEstimado", "valorContrato"],
        )
        dt_ = primeiro(
            r_dict,
            ["dataAssinatura", "dataCelebracao", "dataPublicacaoPncp"],
        )
        dt_fim_vig = primeiro(
            r_dict,
            [
                "dataVigenciaFim",
                "dataFimVigencia",
                "dataTerminoVigencia",
                "dataEncerramento",
            ],
        )
        sit = primeiro(r_dict, ["situacao", "status"])

    elif tipo == "Atas de Registro de Preços":
        num = primeiro(r_dict, ["numeroAtaRegistroPreco", "numeroAta", "numero"])
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
        val_global = primeiro(
            r_dict,
            [
                "valorTotalAta",
                "valorTotal",
                "valorGlobal",
                "valorAta",
                "compra.valorTotalHomologado",
            ],
        )
        val_inicial = primeiro(r_dict, ["valorInicial", "valorTotalEstimado"])
        dt_ = primeiro(
            r_dict,
            ["dataAssinatura", "dataPublicacaoPncp", "dataCelebracao"],
        )
        dt_fim_vig = primeiro(
            r_dict,
            ["dataVigenciaFim", "dataFimVigencia", "dataValidadeFim"],
        )
        sit = primeiro(r_dict, ["situacao", "status"])

    else:
        num = primeiro(
            r_dict,
            ["numeroCompra", "compra.numeroCompra", "numeroEdital", "numero"],
        )
        proc = primeiro(r_dict, ["processo", "compra.processo", "numeroProcesso"])
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
            ],
        )
        val_global = primeiro(
            r_dict,
            [
                "valorTotalHomologado",
                "compra.valorTotalHomologado",
                "valorTotal",
                "valorGlobal",
                "valorTotalEstimado",
                "compra.valorTotalEstimado",
            ],
        )
        val_inicial = primeiro(
            r_dict,
            ["valorTotalEstimado", "compra.valorTotalEstimado"],
        )
        dt_ = primeiro(
            r_dict,
            ["dataPublicacao", "dataPublicacaoPncp", "dataInclusao"],
        )
        dt_fim_vig = primeiro(
            r_dict,
            ["dataVigenciaFim", "dataFimVigencia", "dataEncerramentoProposta"],
        )
        sit = primeiro(
            r_dict,
            [
                "situacaoCompra",
                "situacaoCompraNome",
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

    if texto(obj, "") == "":
        obj = fuzzy_match(r_dict, ["objeto", "descricao"])
    if texto(forn, "") == "":
        forn = fuzzy_match(r_dict, ["razaosocial", "fornec", "fornecedor.nome"])
    if texto(val_global, "") == "":
        val_global = fuzzy_match(
            r_dict,
            ["valortotal", "valorglobal", "valorata", "valor"],
        )
    if dt_fim_vig == "N/D" or dt_fim_vig is None:
        dt_fim_vig = fuzzy_match(
            r_dict,
            ["vigenciafim", "fimvigencia", "datatermino", "datavalidadefim"],
        )
    if texto(mod, "") == "":
        mod = fuzzy_match(r_dict, ["modalidade", "tipoata"])
    if texto(cnpj, "") == "" and tipo != "Editais e Avisos de Contratações":
        cnpj = fuzzy_match(r_dict, ["nifornecedor", "cnpjfornecedor"])

    v_num = numero(val_global)
    v_ini_num = numero(val_inicial)

    eh_termo_aditivo = False
    if str(tipo_contrato_id) == "2" or "aditivo" in str(tipo_instrumento).lower():
        eh_termo_aditivo = True
    elif any(
        palavra in str(obj).lower() for palavra in ("termo aditivo", "aditamento")
    ):
        eh_termo_aditivo = True

    termos_vinculados: List[Dict[str, Any]] = []
    diagnostico_aditivo: Dict[str, Any] = {
        "status": "NAO_CONSULTADO",
        "termos": [],
        "http_status": None,
        "mensagem": "Consulta de termos não aplicável a este escopo.",
        "url": "",
    }

    if tipo == "Contratos":
        diagnostico_aditivo = buscar_termos_contrato(
            cnpj_orgao,
            ano_contrato if not valor_vazio(ano_contrato) else ano_url,
            seq_contrato if not valor_vazio(seq_contrato) else seq_url,
            consultar_api=consultar_termos_api,
        )
        termos_vinculados = diagnostico_aditivo.get("termos", [])
        if diagnostico_aditivo.get("status") == "REGISTROS_ENCONTRADOS":
            eh_termo_aditivo = True

    perc_aditivo = None
    if (
        v_num is not None
        and v_ini_num is not None
        and v_ini_num > 0
        and v_num > v_ini_num
    ):
        perc_aditivo = ((v_num - v_ini_num) / v_ini_num) * 100.0

    ts_fim_vig = data_timestamp(dt_fim_vig)
    dias_para_vencer = None
    if ts_fim_vig is not None:
        dias_para_vencer = int((ts_fim_vig - hoje).days)

    obj_txt = texto(obj)
    forn_txt = texto(forn)
    num_txt = texto(num)
    proc_txt = texto(proc)
    cnpj_forn_limpo = cnpj_limpo(cnpj)

    # Pré-computa string unificada de busca para filtragem instantânea na UI
    busca_unificada = (
        f"{obj_txt} {forn_txt} {cnpj_forn_limpo} {num_txt} {proc_txt} {controle_str}"
    ).casefold()

    return {
        "controle": texto(controle),
        "numero": num_txt,
        "processo": proc_txt,
        "ano_contrato": ano_contrato if not valor_vazio(ano_contrato) else ano_url,
        "sequencial_contrato": seq_contrato if not valor_vazio(seq_contrato) else seq_url,
        "ano_url": ano_url,
        "sequencial_url": seq_url,
        "sequencial_ata_url": seq_ata_url,
        "tipo_instrumento": texto(tipo_instrumento),
        "eh_aditivo": eh_termo_aditivo,
        "qtd_termos_aditivos": len(termos_vinculados),
        "termos_detalhes": termos_vinculados,
        "diagnostico_aditivo": diagnostico_aditivo,
        "objeto": obj_txt,
        "fornecedor": forn_txt,
        "cnpj_fornecedor": cnpj_forn_limpo,
        "valor_num": v_num,
        "valor_inicial_num": v_ini_num,
        "perc_aditivo": perc_aditivo,
        "valor": moeda(val_global),
        "data": data_br(dt_),
        "data_fim_vigencia": data_br(dt_fim_vig),
        "dias_para_vencer": dias_para_vencer,
        "situacao": texto(sit),
        "modalidade": texto(mod),
        "link_pncp": link_pncp,
        "_busca": busca_unificada,
    }


def estruturar_dados_registros(
    df: pd.DataFrame,
    tipo: str,
    cnpj_orgao: str = CNPJ_PADRAO,
    consultar_termos_api: bool = True,
    progress_callback: Optional[Callable[[float, str], None]] = None,
) -> List[Dict[str, Any]]:
    """Estrutura os registros em paralelo uma única vez por carga."""
    if df.empty:
        return []

    registros = df.to_dict("records")
    hoje_ts = pd.Timestamp.today().normalize()
    total = len(registros)

    if tipo != "Contratos" or not consultar_termos_api or total <= 1:
        saida = []
        for idx, row in enumerate(registros, start=1):
            saida.append(
                dados_registro(
                    row,
                    tipo,
                    cnpj_orgao=cnpj_orgao,
                    consultar_termos_api=consultar_termos_api,
                    hoje_ts=hoje_ts,
                )
            )
            if progress_callback and (idx % 20 == 0 or idx == total):
                progress_callback(idx / total, f"Estruturando registro {idx}/{total}...")
        return saida

    workers = min(MAX_WORKERS, total)
    resultado: List[Optional[Dict[str, Any]]] = [None] * total

    def processar(indice: int, row: Dict[str, Any]) -> Tuple[int, Dict[str, Any]]:
        return (
            indice,
            dados_registro(
                row,
                tipo,
                cnpj_orgao=cnpj_orgao,
                consultar_termos_api=consultar_termos_api,
                hoje_ts=hoje_ts,
            ),
        )

    concluidos = 0
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futuros = {
            executor.submit(processar, indice, row): indice
            for indice, row in enumerate(registros)
        }
        for futuro in as_completed(futuros):
            indice = futuros[futuro]
            concluidos += 1
            try:
                indice_resultado, dado = futuro.result()
                resultado[indice_resultado] = dado
            except Exception:
                LOGGER.exception("Falha ao estruturar contrato na posição %s", indice)
                resultado[indice] = dados_registro(
                    registros[indice],
                    tipo,
                    cnpj_orgao=cnpj_orgao,
                    consultar_termos_api=False,
                    hoje_ts=hoje_ts,
                )
            if progress_callback:
                progress_callback(
                    concluidos / total,
                    f"Verificando termos aditivos ({concluidos}/{total})...",
                )

    return [item for item in resultado if item is not None]


# ============================================================
# PAGINAÇÃO E CONSULTAS PARALELAS
# ============================================================

def inteiro_seguro(valor: Any) -> Optional[int]:
    """Converte um valor para inteiro sem lançar exceção."""
    if valor_vazio(valor):
        return None
    try:
        return int(valor)
    except (TypeError, ValueError):
        return None


def paginacao(data: Any) -> Dict[str, Optional[int]]:
    """Extrai informações de paginação."""
    if not isinstance(data, dict):
        return {"totalPaginas": None, "totalRegistros": None, "numeroPagina": None}

    return {
        "totalPaginas": inteiro_seguro(data.get("totalPaginas")),
        "totalRegistros": inteiro_seguro(data.get("totalRegistros")),
        "numeroPagina": inteiro_seguro(data.get("numeroPagina")),
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
    return max(1, (total_registros + tamanho - 1) // tamanho)


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def consultar_cache(
    url: str,
    params_tuple: Tuple[Tuple[str, str], ...],
    max_paginas: int,
) -> Tuple[List[Dict[str, Any]], int, Optional[int], List[str]]:
    """Consulta paginada com paralelismo."""
    base = dict(params_tuple)
    base["pagina"] = 1

    primeira = get_json(url, base)
    registros_primeira = registros_api(primeira)
    info = paginacao(primeira)

    tamanho_pagina = inteiro_seguro(base.get("tamanhoPagina")) or 50
    total = calcular_total_paginas(info, tamanho_pagina)

    if not registros_primeira:
        return ([], 1, total, [])

    limite = min(max(1, max_paginas), total if total is not None else max_paginas)
    todos = list(registros_primeira)

    if limite <= 1:
        return (todos, 1, total, [])

    resultados: Dict[int, List[Dict[str, Any]]] = {}
    erros: List[str] = []
    workers = min(MAX_WORKERS, max(1, limite - 1))

    def buscar_pagina(pagina: int) -> Tuple[int, List[Dict[str, Any]]]:
        params = dict(base)
        params["pagina"] = pagina
        resposta = get_json(url, params)
        return pagina, registros_api(resposta)

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futuros = {executor.submit(buscar_pagina, p): p for p in range(2, limite + 1)}
        for futuro in as_completed(futuros):
            pagina = futuros[futuro]
            try:
                pagina_resultado, registros = futuro.result()
                resultados[pagina_resultado] = registros
            except Exception as exc:
                mensagem = f"Página {pagina}: {exc}"
                LOGGER.error("Erro na consulta: %s", mensagem)
                erros.append(mensagem)

    for pagina in range(2, limite + 1):
        todos.extend(resultados.get(pagina, []))

    return (todos, limite, total, erros)


def consultar(
    url: str,
    params: Dict[str, Any],
    max_paginas: int,
) -> Tuple[List[Dict[str, Any]], int, Optional[int], List[str]]:
    """Wrapper da consulta com tupla serializada para o cache."""
    serial = tuple(sorted((str(k), str(v)) for k, v in params.items()))
    return consultar_cache(url, serial, max_paginas)


# ============================================================
# PROCESSAMENTO DE DATAFRAMES
# ============================================================

def normalizar_pncp(registros: List[Dict[str, Any]]) -> pd.DataFrame:
    """Normaliza JSON PNCP para DataFrame de forma otimizada."""
    if not registros:
        return pd.DataFrame()

    df = pd.json_normalize(registros, sep=".")
    if df.empty:
        return df

    # Converte apenas colunas do tipo 'object' que possuam dicts/lists
    obj_cols = df.select_dtypes(include=["object"]).columns
    for coluna in obj_cols:
        df[coluna] = df[coluna].map(
            lambda v: str(v) if isinstance(v, (dict, list)) else v
        )

    return df


def deduplicar(df: pd.DataFrame) -> pd.DataFrame:
    """Deduplica mantendo sequenciais diferentes de contratos/aditivos."""
    if df.empty:
        return df

    colunas_compostas = [
        c
        for c in (
            "numeroControlePNCP",
            "numeroControlePNCPAta",
            "numeroControlePNCPCompra",
            "idContratoPNCP",
            "sequencialContrato",
            "sequencialTermoContrato",
            "numeroTermoContrato",
            "tipoContrato.id",
        )
        if c in df.columns
    ]

    if colunas_compostas:
        resultado = df.drop_duplicates(subset=colunas_compostas, keep="first")
    else:
        resultado = df.drop_duplicates()

    return resultado.reset_index(drop=True)


def filtrar_cnpj(df: pd.DataFrame, cnpj_orgao: str = CNPJ_PADRAO) -> pd.DataFrame:
    """Filtra o órgão pelo CNPJ informado."""
    if df.empty:
        return df

    alvo = cnpj_limpo(cnpj_orgao)
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
        serie = df[coluna].map(cnpj_limpo)
        mascara = serie == alvo
        if mascara.any():
            mascaras.append(mascara)

    if not mascaras:
        return df

    mascara_final = mascaras[0].copy()
    for mascara in mascaras[1:]:
        mascara_final |= mascara

    return df.loc[mascara_final].reset_index(drop=True)


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
    """Aplica filtro temporal."""
    if df.empty:
        return df

    colunas_disponiveis = [c for c in COLUNAS_DATA if c in df.columns]
    if not colunas_disponiveis:
        return df

    mascaras = []
    for coluna in colunas_disponiveis:
        datas = pd.to_datetime(df[coluna], errors="coerce").dt.normalize()
        mascaras.append(datas.notna() & (datas >= inicio) & (datas <= fim))

    mascara_final = mascaras[0].copy()
    for mascara in mascaras[1:]:
        mascara_final |= mascara

    return df.loc[mascara_final].reset_index(drop=True)


# ============================================================
# MOTOR DE RISCO OTIMIZADO (COMPLEXIDADE O(N))
# ============================================================

def preparar_estatisticas_contexto(
    dados_processados: Sequence[Dict[str, Any]],
) -> Dict[str, Any]:
    """Pré-calcula estatísticas da amostra uma única vez (evita O(N²) no loop)."""
    valores = pd.Series(
        [
            r["valor_num"]
            for r in dados_processados
            if r.get("valor_num") is not None and r["valor_num"] > 0
        ],
        dtype="float64",
    ).dropna()

    qtd_valores = len(valores)
    q50 = q75 = q90 = limite_iqr = None

    if qtd_valores >= 5:
        q25 = float(valores.quantile(0.25))
        q50 = float(valores.median())
        q75 = float(valores.quantile(0.75))
        q90 = float(valores.quantile(0.90))
        iqr = q75 - q25
        if iqr > 0:
            limite_iqr = float(q75 + 1.5 * iqr)

    contagem_cnpj: Dict[str, int] = {}
    total_com_cnpj = 0
    for r in dados_processados:
        c = cnpj_limpo(r.get("cnpj_fornecedor"))
        if c:
            total_com_cnpj += 1
            contagem_cnpj[c] = contagem_cnpj.get(c, 0) + 1

    return {
        "amostra_total": len(dados_processados),
        "qtd_valores": qtd_valores,
        "q50": q50,
        "q75": q75,
        "q90": q90,
        "limite_iqr": limite_iqr,
        "contagem_cnpj": contagem_cnpj,
        "total_com_cnpj": total_com_cnpj,
    }


def calcular_risco(
    registro: Dict[str, Any],
    stats: Dict[str, Any],
    tipo: str,
) -> Dict[str, Any]:
    """Classificação de risco orientada a auditoria (O(1) por registro)."""
    pontos = 0
    motivos: List[str] = []
    testes: List[str] = []
    dimensoes_ativas = 0

    valor = registro.get("valor_num")
    modalidade = texto(registro.get("modalidade"), "").lower()
    objeto = texto(registro.get("objeto"), "").lower()

    # 1. Materialidade
    if valor is not None and valor > 0:
        dimensoes_ativas += 1
        qtd_valores = stats.get("qtd_valores", 0)
        q75 = stats.get("q75")
        q90 = stats.get("q90")

        if qtd_valores >= 5:
            if q90 is not None and q90 > 0 and valor >= q90:
                pontos += 15
                motivos.append("Alta materialidade relativa no conjunto analisado (≥ P90)")
                testes.append("Priorizar conferência do valor, pesquisa de preços e execução")
            elif q75 is not None and q75 > 0 and valor >= q75:
                pontos += 8
                motivos.append("Materialidade acima do padrão do conjunto (≥ P75)")

            if valor >= 5_000_000:
                pontos += 10
                motivos.append("Materialidade absoluta muito elevada (> R$ 5 milhões)")
            elif valor >= 1_000_000:
                pontos += 5
                motivos.append("Materialidade absoluta elevada (> R$ 1 milhão)")
        else:
            if valor >= 5_000_000:
                pontos += 15
                motivos.append("Materialidade absoluta muito elevada (> R$ 5 milhões)")
            elif valor >= 1_000_000:
                pontos += 8
                motivos.append("Materialidade absoluta elevada (> R$ 1 milhão)")

    # 2. Anomalia financeira (IQR)
    limite_superior = stats.get("limite_iqr")
    outlier = False
    if valor is not None and limite_superior is not None and valor > limite_superior:
        outlier = True
        dimensoes_ativas += 1
        pontos += 20
        motivos.append("Valor anômalo acima do limite estatístico IQR")
        testes.append("Avaliar formação de preços, pesquisa de mercado e justificativa do valor")

    # 3. Modalidade / contratação direta
    if "dispensa" in modalidade or "inexig" in modalidade:
        dimensoes_ativas += 1
        pontos += 12
        motivos.append("Contratação direta — requer verificação do fundamento legal")
        testes.append("Conferir hipótese legal, justificativa, pesquisa de preços e documentos")

    if "emerg" in objeto or "emerg" in modalidade:
        dimensoes_ativas += 1
        pontos += 18
        motivos.append("Indício de contratação emergencial")
        testes.append("Verificar caracterização da emergência, prazo, objeto e justificativa")

    # 4. Vigência
    dias_venc = registro.get("dias_para_vencer")
    if dias_venc is not None:
        dimensoes_ativas += 1
        if dias_venc < 0:
            pontos += 20
            motivos.append(f"Instrumento com vigência encerrada há {abs(dias_venc)} dia(s)")
            testes.append("Verificar situação atual, eventual prorrogação e continuidade da execução")
        elif dias_venc <= 15:
            pontos += 15
            motivos.append(f"Vencimento iminente ({dias_venc} dias restantes)")
            testes.append("Verificar providências para encerramento ou prorrogação tempestiva")
        elif dias_venc <= 30:
            pontos += 10
            motivos.append(f"Vencimento próximo ({dias_venc} dias restantes)")
            testes.append("Acompanhar planejamento do encerramento ou eventual prorrogação")
        elif dias_venc <= 90:
            pontos += 4
            motivos.append(f"Vigência próxima do encerramento ({dias_venc} dias restantes)")

    # 5. Aditamentos — somente contratos
    if tipo == "Contratos":
        perc_adit = registro.get("perc_aditivo")
        eh_adit = registro.get("eh_aditivo", False)
        qtd_termos = registro.get("qtd_termos_aditivos", 0)

        if perc_adit is not None and perc_adit > 0:
            dimensoes_ativas += 1
            if perc_adit > 50.0:
                pontos += 25
                motivos.append(f"Variação financeira superior a 50% (+{perc_adit:.1f}%)")
                testes.append("Verificar natureza do objeto, hipótese legal, justificativa e limite aplicável")
            elif perc_adit > 25.0:
                pontos += 16
                motivos.append(f"Variação financeira superior a 25% (+{perc_adit:.1f}%)")
                testes.append("Verificar enquadramento jurídico, justificativa e limite aplicável")
            elif perc_adit > 10.0:
                pontos += 7
                motivos.append(f"Aditamento financeiro relevante (+{perc_adit:.1f}%)")
                testes.append("Conferir justificativa, memória de cálculo e documentação do aditivo")
            else:
                pontos += 3
                motivos.append(f"Aditamento financeiro identificado (+{perc_adit:.1f}%)")

        if eh_adit or qtd_termos > 0:
            dimensoes_ativas += 1
            pontos += 5
            if qtd_termos > 0:
                motivos.append(f"Termo(s) aditivo(s) identificado(s) no PNCP ({qtd_termos})")
            else:
                motivos.append("Registro identificado como possível termo aditivo")
            testes.append("Revisar motivação, formalização, vigência e anexos do termo aditivo")

    # 6. Concentração de fornecedor
    cnpj_forn = cnpj_limpo(registro.get("cnpj_fornecedor"))
    total_com_cnpj = stats.get("total_com_cnpj", 0)
    contagem_cnpj = stats.get("contagem_cnpj", {})
    if cnpj_forn and total_com_cnpj >= 5:
        qtd = contagem_cnpj.get(cnpj_forn, 0)
        if qtd >= 3:
            proporcao = qtd / total_com_cnpj
            dimensoes_ativas += 1
            if proporcao >= 0.50:
                pontos += 12
                motivos.append(f"Alta concentração de fornecedor ({qtd}/{total_com_cnpj})")
                testes.append("Avaliar concentração, competição e justificativa da contratação")
            elif proporcao >= 0.25:
                pontos += 7
                motivos.append(f"Concentração relevante de fornecedor ({qtd}/{total_com_cnpj})")
                testes.append("Avaliar distribuição das contratações e competição")
            elif qtd >= 5:
                pontos += 4
                motivos.append(f"Fornecedor recorrente no conjunto ({qtd} ocorrências)")

    # 7. Qualidade cadastral (não altera o score)
    campos_obrigatorios = ("objeto", "controle", "data")
    faltantes = [
        campo
        for campo in campos_obrigatorios
        if texto(registro.get(campo), "") in {"", "N/D"}
    ]
    if faltantes:
        motivos.append("Qualidade dos dados: campos ausentes — " + ", ".join(faltantes))
        testes.append("Conferir o cadastro original no PNCP antes da análise")

    pontos = max(0, min(100, int(pontos)))

    if pontos >= 60:
        nivel = "🔴 ALTO"
    elif pontos >= 35:
        nivel = "🟡 MÉDIO"
    else:
        nivel = "🟢 BAIXO"

    if not motivos:
        motivos = ["Sem sinais relevantes de risco identificados automaticamente."]

    if stats.get("amostra_total", 0) < 5:
        confianca = "BAIXA — amostra pequena"
    elif dimensoes_ativas <= 1:
        confianca = "MODERADA — poucos sinais independentes"
    else:
        confianca = "BOA — múltiplos sinais analisados"

    return {
        "pontos": pontos,
        "nivel": nivel,
        "motivos": motivos,
        "testes": testes,
        "outlier": outlier,
        "confianca": confianca,
    }


def calcular_riscos_lote(
    dados_processados: Sequence[Dict[str, Any]],
    tipo: str,
) -> List[Dict[str, Any]]:
    """Calcula a matriz de risco para todos os registros em O(N)."""
    if not dados_processados:
        return []
    stats = preparar_estatisticas_contexto(dados_processados)
    return [calcular_risco(reg, stats, tipo) for reg in dados_processados]


# ============================================================
# NLP / SIMILARIDADE SEMÂNTICA E DETECÇÃO DE FRACIONAMENTO
# ============================================================

STOPWORDS_CONTRATOS = {
    "a", "o", "os", "as", "um", "uma", "uns", "umas", "de", "do", "da", "dos", "das",
    "em", "no", "na", "nos", "nas", "por", "pelo", "pela", "pelos", "pelas", "com",
    "para", "que", "e", "ou", "ao", "aos", "sob", "sobre", "entre", "se",
    "contrato", "contratacao", "contratada", "contratante", "celebrado", "celebram",
    "termo", "referencia", "especificacoes", "contidas", "anexo", "anexos", "edital",
    "dispensa", "eletronica", "processo", "administrativo", "objeto", "presente",
    "visa", "visando", "atender", "atendimento", "prestacao", "servico", "servicos",
    "empresa", "especializada", "prefeitura", "municipio", "municipal", "secretaria",
    "rio", "pedras", "sp", "conforme", "observadas", "condicoes", "exigencias",
    "estabelecidas", "leis", "lei", "artigo", "art", "vigor", "fornecimento",
    "aquisicao", "locacao", "realizacao", "execucao", "ramo", "me", "epp", "ltda",
    "sa", "cnpj", "cpf", "compromisso", "governo", "hora", "horas", "minutos",
    "dia", "dias", "inicio", "duracao", "valor", "total", "itens", "item",
}


def texto_nlp(valor: Any) -> str:
    """Normaliza o texto do objeto e remove jargões burocráticos rapidamente."""
    s = texto(valor, "").translate(TABELA_SEM_ACENTOS).lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    palavras = [
        p for p in s.split() if len(p) > 2 and p not in STOPWORDS_CONTRATOS
    ]
    if not palavras:
        return "objeto_sem_termos_chave"
    return " ".join(palavras)


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def calcular_modelo_tfidf(textos: Tuple[str, ...]):
    """Cria e cacheia a matriz TF-IDF baseada no núcleo semântico dos objetos."""
    if not SKLEARN_OK or TfidfVectorizer is None or len(textos) < 2:
        return None

    textos_seguros = [texto_nlp(t) for t in textos]
    textos_validos = {t for t in textos_seguros if t != "objeto_sem_termos_chave"}
    if len(textos_validos) < 2:
        return None

    vetor = TfidfVectorizer(
        lowercase=True,
        ngram_range=(1, 2),
        min_df=1,
        max_features=8000,
        sublinear_tf=True,
    )
    try:
        return vetor.fit_transform(textos_seguros)
    except ValueError as exc:
        LOGGER.warning("Não foi possível criar modelo TF-IDF: %s", exc)
        return None


def similares(
    dados_processados: List[Dict[str, Any]],
    idx: int,
    limite: int = 5,
    corte_minimo: float = 0.15,
) -> pd.DataFrame:
    """Retorna registros semanticamente semelhantes filtrando ruídos burocráticos."""
    if (
        not SKLEARN_OK
        or len(dados_processados) < 2
        or idx < 0
        or idx >= len(dados_processados)
    ):
        return pd.DataFrame()

    objetos = tuple(r.get("objeto", "") for r in dados_processados)
    matriz = calcular_modelo_tfidf(objetos)
    if matriz is None:
        return pd.DataFrame()

    try:
        scores = cosine_similarity(matriz[idx], matriz).ravel()
    except (IndexError, ValueError) as exc:
        LOGGER.warning("Erro no cálculo de similaridade: %s", exc)
        return pd.DataFrame()

    ordem = scores.argsort()[::-1]
    saida = []

    for j in ordem:
        j = int(j)
        if j == idx:
            continue

        score_val = float(scores[j])
        if score_val < corte_minimo:
            continue

        registro = dados_processados[j]
        saida.append(
            {
                "Similaridade": f"{score_val * 100:.1f}%",
                "Número": texto(registro.get("numero")),
                "Modalidade": texto(registro.get("modalidade")),
                "Objeto": texto(registro.get("objeto")),
                "Fornecedor": texto(registro.get("fornecedor")),
                "Valor": texto(registro.get("valor")),
                "Link PNCP": registro.get("link_pncp", "https://pncp.gov.br"),
            }
        )
        if len(saida) >= limite:
            break

    return pd.DataFrame(saida)


def detectar_possiveis_fracionamentos(
    dados_processados: List[Dict[str, Any]],
    similaridade_minima: float = 0.38,
) -> pd.DataFrame:
    """Identifica pares ou grupos de contratações diretas com objetos muito semelhantes (Lei 14.133/21, art. 75, § 1º)."""
    if not SKLEARN_OK or len(dados_processados) < 2:
        return pd.DataFrame()

    # Filtra apenas contratações diretas ou analisa todas caso não haja dispensas rotuladas
    dispensas_idx = [
        i
        for i, r in enumerate(dados_processados)
        if any(
            termo in texto(r.get("modalidade"), "").lower()
            for termo in ("dispensa", "inexig", "direta")
        )
    ]

    indices_alvo = dispensas_idx if len(dispensas_idx) >= 2 else list(range(len(dados_processados)))
    if len(indices_alvo) < 2:
        return pd.DataFrame()

    objetos = tuple(dados_processados[i].get("objeto", "") for i in indices_alvo)
    matriz = calcular_modelo_tfidf(objetos)
    if matriz is None:
        return pd.DataFrame()

    try:
        sim_matrix = cosine_similarity(matriz)
    except Exception:
        return pd.DataFrame()

    pares = []
    n = len(indices_alvo)
    for a in range(n):
        for b in range(a + 1, n):
            score = float(sim_matrix[a, b])
            if score < similaridade_minima:
                continue

            reg_a = dados_processados[indices_alvo[a]]
            reg_b = dados_processados[indices_alvo[b]]
            soma = (reg_a.get("valor_num") or 0.0) + (reg_b.get("valor_num") or 0.0)

            pares.append(
                {
                    "Similaridade": f"{score * 100:.1f}%",
                    "_score": score,
                    "Registro 1": reg_a.get("numero", "N/D"),
                    "Data 1": reg_a.get("data", "N/D"),
                    "Valor 1": reg_a.get("valor", "N/D"),
                    "Registro 2": reg_b.get("numero", "N/D"),
                    "Data 2": reg_b.get("data", "N/D"),
                    "Valor 2": reg_b.get("valor", "N/D"),
                    "Soma Combinada": moeda(soma) if soma > 0 else "N/D",
                    "Mesmo Fornecedor?": (
                        "Sim"
                        if reg_a.get("cnpj_fornecedor")
                        and reg_a.get("cnpj_fornecedor") == reg_b.get("cnpj_fornecedor")
                        else "Não"
                    ),
                    "Objeto 1": reg_a.get("objeto", "N/D"),
                    "Objeto 2": reg_b.get("objeto", "N/D"),
                }
            )

    if not pares:
        return pd.DataFrame()

    df_pares = (
        pd.DataFrame(pares)
        .sort_values("_score", ascending=False)
        .drop(columns=["_score"])
        .head(50)
        .reset_index(drop=True)
    )
    return df_pares


def consolidar_fornecedores(
    dados_processados: List[Dict[str, Any]],
    riscos: List[Dict[str, Any]],
) -> pd.DataFrame:
    """Consolida o ranking de fornecedores por volume financeiro, quantidade e score médio."""
    agrupado: Dict[str, Dict[str, Any]] = {}
    valor_total_amostra = sum(
        r["valor_num"]
        for r in dados_processados
        if r.get("valor_num") is not None and r["valor_num"] > 0
    )

    for reg, risco in zip(dados_processados, riscos):
        forn_nome = texto(reg.get("fornecedor"), "Fornecedor Não Informado")
        cnpj_forn = cnpj_limpo(reg.get("cnpj_fornecedor"))
        chave = cnpj_forn if cnpj_forn else forn_nome

        if chave not in agrupado:
            agrupado[chave] = {
                "Fornecedor": forn_nome,
                "CNPJ/CPF": cnpj_formatado(cnpj_forn) if cnpj_forn else "N/D",
                "Qtd. Registros": 0,
                "Valor Acumulado (Num)": 0.0,
                "Soma Score": 0,
                "Alto Risco": 0,
            }

        item = agrupado[chave]
        item["Qtd. Registros"] += 1
        if reg.get("valor_num") is not None and reg["valor_num"] > 0:
            item["Valor Acumulado (Num)"] += float(reg["valor_num"])
        item["Soma Score"] += int(risco.get("pontos", 0))
        if risco.get("nivel") == "🔴 ALTO":
            item["Alto Risco"] += 1

    if not agrupado:
        return pd.DataFrame()

    linhas = []
    for item in agrupado.values():
        val_acum = item["Valor Acumulado (Num)"]
        qtd = item["Qtd. Registros"]
        part_perc = (
            (val_acum / valor_total_amostra * 100.0)
            if valor_total_amostra > 0
            else 0.0
        )
        linhas.append(
            {
                "Fornecedor": item["Fornecedor"],
                "CNPJ/CPF": item["CNPJ/CPF"],
                "Qtd. Registros": qtd,
                "Valor Acumulado": moeda(val_acum),
                "% do Volume": f"{part_perc:.1f}%",
                "Score Médio": round(item["Soma Score"] / max(1, qtd), 1),
                "Registros Alto Risco": item["Alto Risco"],
                "_val": val_acum,
            }
        )

    df_forn = (
        pd.DataFrame(linhas)
        .sort_values(["_val", "Qtd. Registros"], ascending=[False, False])
        .drop(columns=["_val"])
        .reset_index(drop=True)
    )
    return df_forn


# ============================================================
# EXPORTAÇÃO WORD - DASHBOARD E DOSSIÊ INDIVIDUAL
# ============================================================

def _word_text(valor: Any, padrao: str = "N/D") -> str:
    """Converte valores para texto adequado ao Word."""
    if valor_vazio(valor):
        return padrao

    if isinstance(valor, float):
        if pd.isna(valor):
            return padrao
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return texto(valor, padrao)


def _adicionar_hyperlink_word(
    paragrafo: Any, url: str, texto_link: str = "🔗 Acessar"
) -> None:
    """Adiciona um hyperlink externo clicável em um parágrafo do Word."""
    if not url or not str(url).startswith(("http://", "https://")):
        paragrafo.add_run(texto_link)
        return

    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        parte = paragrafo.part
        rel_id = parte.relate_to(
            str(url),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )

        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), rel_id)

        run = OxmlElement("w:r")
        r_pr = OxmlElement("w:rPr")
        r_style = OxmlElement("w:rStyle")
        r_style.set(qn("w:val"), "Hyperlink")
        r_pr.append(r_style)
        run.append(r_pr)

        text_element = OxmlElement("w:t")
        text_element.text = texto_link
        run.append(text_element)
        hyperlink.append(run)
        paragrafo._p.append(hyperlink)
    except Exception:
        paragrafo.add_run(str(url))


def _adicionar_tabela_word(
    doc: Any,
    df: pd.DataFrame,
    titulo: Optional[str] = None,
    max_linhas: int = 500,
) -> None:
    """Adiciona um DataFrame ao Word em formato de tabela."""
    if titulo:
        p = doc.add_paragraph()
        run = p.add_run(titulo)
        run.bold = True
        run.font.size = Pt(12)

    if df is None or df.empty:
        doc.add_paragraph("Nenhum dado disponível.")
        return

    tabela_df = df.copy().head(max_linhas)

    tabela = doc.add_table(
        rows=1,
        cols=len(tabela_df.columns),
    )
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.style = "Table Grid"

    for j, coluna in enumerate(tabela_df.columns):
        celula = tabela.rows[0].cells[j]
        celula.text = str(coluna)
        for paragrafo in celula.paragraphs:
            for run in paragrafo.runs:
                run.bold = True
                run.font.size = Pt(8)

    for _, row in tabela_df.iterrows():
        celulas = tabela.add_row().cells
        for j, valor in enumerate(row.tolist()):
            coluna = str(tabela_df.columns[j])
            if coluna == "Link PNCP" and str(valor).startswith(("http://", "https://")):
                celulas[j].text = ""
                paragrafo = celulas[j].paragraphs[0]
                _adicionar_hyperlink_word(paragrafo, str(valor), "🔗 Acessar")
                for run in paragrafo.runs:
                    run.font.size = Pt(7)
            else:
                celulas[j].text = _word_text(valor)
                for paragrafo in celulas[j].paragraphs:
                    for run in paragrafo.runs:
                        run.font.size = Pt(7)

    if len(df) > max_linhas:
        doc.add_paragraph(
            f"Observação: foram exibidas {max_linhas} linhas de "
            f"{len(df)} registros para manter o documento em tamanho adequado."
        )


@st.cache_data(show_spinner=False)
def gerar_word_dashboard_principal(
    df_risco: pd.DataFrame,
    tipo: str,
    inicio: Any,
    fim: Any,
    modalidade: str,
    total_segmentados: int,
    altos: int,
    medios: int,
    valor_total: Any,
    cnpj_orgao: str = CNPJ_PADRAO,
) -> bytes:
    """Gera o Word com o resumo executivo: identificação, indicadores, matriz de risco e aviso."""
    if not DOCX_OK:
        raise RuntimeError(
            "A exportação para Word requer a biblioteca 'python-docx'."
        )

    doc = Document()

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run("PAINEL DE INTELIGÊNCIA E APOIO AO CONTROLE INTERNO - PNCP")
    run.bold = True
    run.font.size = Pt(16)

    subtitulo = doc.add_paragraph()
    subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitulo.add_run(PREFEITURA)
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph(
        f"Escopo: {tipo}\n"
        f"Período: {data_br(inicio)} a {data_br(fim)}\n"
        f"Segmentação: {modalidade}\n"
        f"CNPJ do órgão: {cnpj_formatado(cnpj_orgao)}\n"
        f"IBGE: {IBGE_PADRAO}"
    )

    doc.add_heading("1. Indicadores Principais", level=1)
    tabela_indicadores = doc.add_table(rows=4, cols=2)
    tabela_indicadores.style = "Table Grid"
    tabela_indicadores.alignment = WD_TABLE_ALIGNMENT.CENTER

    indicadores = [
        ("Registros Segmentados", str(total_segmentados)),
        ("🔴 Alto risco", str(altos)),
        ("🟡 Médio risco", str(medios)),
        ("Valor Analisado", moeda(valor_total)),
    ]

    for i, (rotulo, valor) in enumerate(indicadores):
        tabela_indicadores.rows[i].cells[0].text = rotulo
        tabela_indicadores.rows[i].cells[1].text = valor

    doc.add_heading("2. Matriz de Risco Segmentada", level=1)
    colunas_visualizacao = [
        c
        for c in (
            "Risco",
            "Score",
            "Número",
            "Tipo",
            "Modalidade",
            "Fornecedor",
            "Valor",
            "Fim Vigência",
            "Aditivo",
            "Link PNCP",
            "Objeto",
            "Gatilhos",
        )
        if c in df_risco.columns and (c != "Aditivo" or tipo == "Contratos")
    ]

    _adicionar_tabela_word(
        doc,
        df_risco[colunas_visualizacao] if colunas_visualizacao else df_risco,
        max_linhas=500,
    )

    doc.add_paragraph(
        "Aviso: os alertas e indicadores apresentados neste relatório "
        "são ferramentas automatizadas de apoio ao Controle Interno. "
        "Os resultados devem ser validados por análise técnica, "
        "jurídica e/ou auditoria humana antes de qualquer conclusão."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def gerar_word_dossie_individual(
    registro: Dict[str, Any],
    risco: Dict[str, Any],
    df_similares: pd.DataFrame,
    tipo: str,
    cnpj_orgao: str = CNPJ_PADRAO,
) -> bytes:
    """Gera um Dossiê Individual de Auditoria em Word para o registro selecionado."""
    if not DOCX_OK:
        raise RuntimeError(
            "A exportação para Word requer a biblioteca 'python-docx'."
        )

    doc = Document()

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tit = titulo.add_run("DOSSIÊ INDIVIDUAL DE AUDITORIA INTERNA - PNCP")
    r_tit.bold = True
    r_tit.font.size = Pt(15)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = sub.add_run(f"{PREFEITURA} (CNPJ {cnpj_formatado(cnpj_orgao)})")
    r_sub.bold = True
    r_sub.font.size = Pt(11)

    doc.add_heading("1. Identificação do Registro", level=1)
    campos = [
        ("Escopo", tipo),
        ("Número do Instrumento", registro.get("numero", "N/D")),
        ("Controle PNCP", registro.get("controle", "N/D")),
        ("Processo Administrativo", registro.get("processo", "N/D")),
        ("Tipo de Instrumento", registro.get("tipo_instrumento", "N/D")),
        ("Modalidade", registro.get("modalidade", "N/D")),
        ("Fornecedor", registro.get("fornecedor", "N/D")),
        ("CNPJ/CPF Fornecedor", cnpj_formatado(registro.get("cnpj_fornecedor"))),
        ("Valor Global/Homologado", registro.get("valor", "N/D")),
        ("Data de Assinatura/Publicação", registro.get("data", "N/D")),
        ("Fim da Vigência", registro.get("data_fim_vigencia", "N/D")),
        ("Situação", registro.get("situacao", "N/D")),
        ("Classificação de Risco", f"{risco.get('nivel')} ({risco.get('pontos')} pts)"),
        ("Confiabilidade da Amostra", risco.get("confianca", "N/D")),
    ]

    tab = doc.add_table(rows=len(campos), cols=2)
    tab.style = "Table Grid"
    for i, (k, v) in enumerate(campos):
        tab.rows[i].cells[0].text = k
        tab.rows[i].cells[1].text = str(v)

    p_obj = doc.add_paragraph()
    r_obj = p_obj.add_run("\nObjeto: ")
    r_obj.bold = True
    p_obj.add_run(registro.get("objeto", "N/D"))

    if registro.get("link_pncp"):
        p_link = doc.add_paragraph("Link Oficial no PNCP: ")
        _adicionar_hyperlink_word(p_link, registro["link_pncp"], registro["link_pncp"])

    doc.add_heading("2. Gatilhos e Sinais de Alerta Identificados", level=1)
    for m in risco.get("motivos", []):
        doc.add_paragraph(str(m), style="List Bullet")

    doc.add_heading("3. Roteiro de Testes de Auditoria Sugerido", level=1)
    testes = risco.get("testes", [])
    if testes:
        for t in testes:
            doc.add_paragraph(str(t), style="List Bullet")
    else:
        doc.add_paragraph("Nenhum procedimento extraordinário requerido automaticamente.")

    if df_similares is not None and not df_similares.empty:
        doc.add_heading("4. Contratações Similares no Mesmo Segmento (Clusterização NLP)", level=1)
        _adicionar_tabela_word(doc, df_similares, max_linhas=15)

    doc.add_paragraph(
        "\nNota Técnica: Este dossiê é gerado de forma automatizada para subsidiar "
        "os trabalhos do Controle Interno, devendo suas conclusões ser confrontadas "
        "com os autos do processo administrativo."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# EXPORTAÇÃO EXCEL MULTICAMADAS
# ============================================================

def ajustar_planilha_excel(worksheet: Any) -> None:
    """Formatação básica da planilha e criação de links clicáveis para o PNCP."""
    worksheet.freeze_panes = "A2"

    if worksheet.max_row >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    cabecalhos = {
        str(celula.value).strip(): celula.column
        for celula in worksheet[1]
        if celula.value is not None
    }
    coluna_link = cabecalhos.get("Link PNCP")
    if coluna_link:
        for linha in range(2, worksheet.max_row + 1):
            celula = worksheet.cell(row=linha, column=coluna_link)
            url = str(celula.value or "").strip()
            if url.startswith(("http://", "https://")):
                celula.hyperlink = url
                celula.value = "🔗 Acessar"
                celula.style = "Hyperlink"

    for coluna in worksheet.columns:
        if not coluna:
            continue
        letra = coluna[0].column_letter
        maior = 0
        for celula in coluna:
            valor = celula.value
            tamanho = 0 if valor is None else len(str(valor))
            maior = max(maior, tamanho)

        if coluna_link and coluna[0].column == coluna_link:
            worksheet.column_dimensions[letra].width = 16
        else:
            worksheet.column_dimensions[letra].width = min(55, max(12, maior + 2))


@st.cache_data(show_spinner=False)
def gerar_excel(
    df: pd.DataFrame,
    tipo: str,
    df_risco: pd.DataFrame,
    df_vig: pd.DataFrame,
    df_forn: pd.DataFrame,
    links_pncp: Tuple[str, ...],
) -> bytes:
    """Gera Excel completo com Matriz de Risco, Prazos, Fornecedores e Dados Brutos."""
    buffer = io.BytesIO()
    df_brutos_export = df.copy()

    if links_pncp and len(links_pncp) == len(df_brutos_export):
        df_brutos_export["Link PNCP"] = list(links_pncp)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_risco.to_excel(writer, sheet_name="Matriz de Risco", index=False)
        if df_vig is not None and not df_vig.empty:
            df_vig.to_excel(writer, sheet_name="Prazos e Vigências", index=False)
        if df_forn is not None and not df_forn.empty:
            df_forn.to_excel(writer, sheet_name="Ranking Fornecedores", index=False)
        df_brutos_export.to_excel(writer, sheet_name="Dados Brutos", index=False)
        for worksheet in writer.book.worksheets:
            ajustar_planilha_excel(worksheet)

    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# CONSULTAS ESPECÍFICAS OTIMIZADAS
# ============================================================

def _registrar_limite_paginacao(
    erros: List[str],
    paginas_consultadas: int,
    total_paginas: Optional[int],
    contexto: str,
) -> None:
    """Registra quando o limite local impede a coleta integral do resultado."""
    if total_paginas is not None and paginas_consultadas < total_paginas:
        erros.append(
            f"{contexto}: o PNCP informou {total_paginas} página(s), mas o app "
            f"consultou {paginas_consultadas} por causa do limite configurado ({paginas_consultadas}). "
            "O conjunto exibido é parcial."
        )


def consultar_contratos(
    inicio: Any,
    fim: Any,
    cnpj_orgao: str = CNPJ_PADRAO,
    max_paginas: int = MAX_PAGINAS,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Consulta contratos."""
    params = {
        "dataInicial": inicio.strftime("%Y%m%d"),
        "dataFinal": fim.strftime("%Y%m%d"),
        "cnpjOrgao": cnpj_limpo(cnpj_orgao),
        "tamanhoPagina": PAGE_CONTRATOS,
    }
    registros, paginas, total, erros = consultar(
        f"{BASE_CONSULTA}/contratos", params, max_paginas
    )
    _registrar_limite_paginacao(erros, paginas, total, "Contratos")
    return registros, erros


def consultar_atas(
    inicio: Any,
    fim: Any,
    cnpj_orgao: str = CNPJ_PADRAO,
    max_paginas: int = MAX_PAGINAS,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Consulta atas de registro de preços."""
    params = {
        "dataInicial": inicio.strftime("%Y%m%d"),
        "dataFinal": fim.strftime("%Y%m%d"),
        "cnpj": cnpj_limpo(cnpj_orgao),
        "tamanhoPagina": PAGE_ATAS,
    }
    registros, paginas, total, erros = consultar(
        f"{BASE_CONSULTA}/atas", params, max_paginas
    )
    _registrar_limite_paginacao(erros, paginas, total, "Atas")
    return registros, erros


def consultar_editais(
    inicio: Any,
    fim: Any,
    cnpj_orgao: str = CNPJ_PADRAO,
    modalidades_codigos: Sequence[int] = MODALIDADES_PNCP,
    max_paginas: int = MAX_PAGINAS,
    progress_callback: Any = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Consulta editais/contratações por modalidade em PARALELO (até 8x mais rápido)."""
    params_base = {
        "dataInicial": inicio.strftime("%Y%m%d"),
        "dataFinal": fim.strftime("%Y%m%d"),
        "cnpj": cnpj_limpo(cnpj_orgao),
        "tamanhoPagina": PAGE_EDITAIS,
    }

    todos_registros: List[Dict[str, Any]] = []
    erros: List[str] = []
    lista_mods = list(modalidades_codigos) or list(MODALIDADES_PNCP)
    total_modalidades = len(lista_mods)

    resultados_por_mod: Dict[int, List[Dict[str, Any]]] = {}
    workers = min(MAX_WORKERS, max(1, total_modalidades))

    def buscar_modalidade(
        modalidade: int,
    ) -> Tuple[int, List[Dict[str, Any]], List[str]]:
        params = dict(params_base)
        params["codigoModalidadeContratacao"] = modalidade
        registros, paginas, total, erros_pagina = consultar(
            f"{BASE_CONSULTA}/contratacoes/publicacao",
            params,
            max_paginas,
        )
        _registrar_limite_paginacao(
            erros_pagina, paginas, total, f"Modalidade {modalidade}"
        )
        return modalidade, registros, erros_pagina

    concluidas = 0
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futuros = {executor.submit(buscar_modalidade, m): m for m in lista_mods}
        for futuro in as_completed(futuros):
            mod = futuros[futuro]
            concluidas += 1
            if progress_callback:
                progress_callback(concluidas / total_modalidades, mod)
            try:
                _, regs, errs = futuro.result()
                if regs:
                    resultados_por_mod[mod] = regs
                if errs:
                    erros.extend([f"Modalidade {mod}: {e}" for e in errs])
            except Exception as exc:
                mensagem = f"Modalidade {mod}: {exc}"
                LOGGER.warning(mensagem)
                erros.append(mensagem)

    # Mantém a ordem determinística por código de modalidade
    for mod in lista_mods:
        todos_registros.extend(resultados_por_mod.get(mod, []))

    return (todos_registros, erros)


# ============================================================
# ESTADO DO STREAMLIT E VALIDAÇÕES
# ============================================================

def dt_date(ano: int, mes: int, dia: int) -> dt.date:
    """Helper para criação de date."""
    return dt.date(ano, mes, dia)


def inicializar_estado() -> None:
    """Inicializa session_state."""
    hoje = pd.Timestamp.today().date()

    defaults = {
        "df": None,
        "dados_totais": None,
        "consulta_meta": None,
        "tipo": TIPOS[0],
        "inicio": dt_date(hoje.year, 1, 1),
        "fim": hoje,
        "cnpj_orgao": CNPJ_PADRAO,
        "consultar_termos_api": True,
        "max_paginas_ui": MAX_PAGINAS,
    }

    for chave, valor in defaults.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def validar_periodo(inicio: Any, fim: Any) -> Optional[str]:
    """Valida o período de consulta."""
    if inicio is None or fim is None:
        return "Informe as datas inicial e final."
    if inicio > fim:
        return "A data inicial não pode ser posterior à data final."
    return None


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

def main() -> None:
    st.set_page_config(
        page_title="Controle Interno — PNCP Rio das Pedras",
        page_icon="🛡️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    aplicar_estilo()
    inicializar_estado()

    # ==========================================
    # CABEÇALHO COM LOGOTIPO
    # ==========================================
    col_logo, col_titulo = st.columns([1, 10], vertical_alignment="center")

    caminho_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_controle.jpg")
    if not os.path.exists(caminho_logo):
        caminho_logo = "logo_controle.jpg"

    with col_logo:
        if os.path.exists(caminho_logo):
            st.image(caminho_logo, use_container_width=True)
        else:
            st.markdown(
                "<h1 style='text-align: center; margin:0;'>🛡️</h1>",
                unsafe_allow_html=True,
            )

    with col_titulo:
        st.title("Inteligência de Controle Interno PNCP")

    st.info(
        "ℹ️ **Aviso:** Os alertas produzidos por este sistema são indicadores automatizados "
        "para apoio ao Controle Interno. Toda conclusão deve ser validada por análise e auditoria humana."
    )

    # ==========================================
    # SIDEBAR (PARÂMETROS E OTIMIZAÇÕES)
    # ==========================================
    with st.sidebar:
        if os.path.exists(caminho_logo):
            st.image(caminho_logo, use_container_width=True)
            st.markdown("---")

        st.header("⚙️ Parâmetros")

        tipo_selecionado = st.selectbox(
            "Escopo da Consulta",
            TIPOS,
            index=TIPOS.index(st.session_state.tipo),
            help="Selecione o tipo de registro que deseja buscar na base do PNCP.",
        )

        if tipo_selecionado != st.session_state.tipo:
            st.session_state.tipo = tipo_selecionado
            st.session_state.df = None
            st.session_state.dados_totais = None
            st.session_state.consulta_meta = None

        inicio = st.date_input(
            "📅 Data inicial",
            value=st.session_state.inicio,
            help="Data de início para o filtro de publicação/assinatura.",
        )
        fim = st.date_input(
            "📅 Data final",
            value=st.session_state.fim,
            help="Data final para o filtro da consulta.",
        )

        st.session_state.inicio = inicio
        st.session_state.fim = fim

        modalidades_edital_sel: List[int] = list(MODALIDADES_PNCP)

        with st.expander("⚡ Desempenho e Opções Avançadas", expanded=False):
            cnpj_input = st.text_input(
                "CNPJ do Órgão",
                value=st.session_state.cnpj_orgao,
                help="Padrão: Prefeitura Municipal de Rio das Pedras/SP (44826840000183).",
            )
            st.session_state.cnpj_orgao = cnpj_limpo(cnpj_input) or CNPJ_PADRAO

            max_pag_input = st.slider(
                "Máximo de páginas por consulta",
                min_value=1,
                max_value=50,
                value=int(st.session_state.max_paginas_ui),
                help="Define o teto de paginação na API do PNCP.",
            )
            st.session_state.max_paginas_ui = max_pag_input

            if tipo_selecionado == "Contratos":
                st.session_state.consultar_termos_api = st.checkbox(
                    "Buscar termos aditivos na API (sub-recurso)",
                    value=st.session_state.consultar_termos_api,
                    help=(
                        "Quando ativo, consulta o endpoint individual de termos de cada contrato. "
                        "Desmarque para um carregamento instantâneo."
                    ),
                )
            elif tipo_selecionado == "Editais e Avisos de Contratações":
                opcoes_mod = [MODALIDADES_PNCP_NOMES[k] for k in MODALIDADES_PNCP]
                mods_escolhidas_str = st.multiselect(
                    "Modalidades PNCP a varrer:",
                    options=opcoes_mod,
                    default=opcoes_mod,
                    help="Reduzir as modalidades acelera ainda mais a busca de editais.",
                )
                mapa_inv = {v: k for k, v in MODALIDADES_PNCP_NOMES.items()}
                modalidades_edital_sel = [mapa_inv[s] for s in mods_escolhidas_str]

        erro_periodo = validar_periodo(inicio, fim)
        if erro_periodo:
            st.error(erro_periodo)

        col_btn1, col_btn2 = st.columns([3, 1])
        with col_btn1:
            buscar_btn = st.button(
                "🔎 Carregar dados",
                type="primary",
                use_container_width=True,
                disabled=erro_periodo is not None,
                help="Clique para buscar os dados direto da base oficial do PNCP.",
            )
        with col_btn2:
            if st.button("🧹", use_container_width=True, help="Limpar cache local e forçar nova consulta à API."):
                st.cache_data.clear()
                st.toast("Cache limpo com sucesso!", icon="🧹")

        st.markdown("---")
        st.caption(f"Município: {MUNICIPIO}")
        st.caption(f"CNPJ Ativo: {cnpj_formatado(st.session_state.cnpj_orgao)}")
        st.caption(f"IBGE: {IBGE_PADRAO}")

        if not SKLEARN_OK:
            st.warning(
                "Scikit-Learn não instalado. "
                "A análise semântica ficará desativada."
            )

    # ==========================================
    # EXECUÇÃO DA CONSULTA
    # ==========================================
    if buscar_btn:
        try:
            todas_regs: List[Dict[str, Any]] = []
            erros_consulta: List[str] = []
            cnpj_ativo = st.session_state.cnpj_orgao
            max_pags = st.session_state.max_paginas_ui

            with st.spinner("Consultando o PNCP em alta velocidade..."):
                tipo_consulta = tipo_selecionado

                if tipo_consulta == "Contratos":
                    registros, erros = consultar_contratos(
                        inicio, fim, cnpj_orgao=cnpj_ativo, max_paginas=max_pags
                    )
                    todas_regs.extend(registros)
                    erros_consulta.extend(erros)

                elif tipo_consulta == "Atas de Registro de Preços":
                    registros, erros = consultar_atas(
                        inicio, fim, cnpj_orgao=cnpj_ativo, max_paginas=max_pags
                    )
                    todas_regs.extend(registros)
                    erros_consulta.extend(erros)

                else:
                    barra = st.progress(
                        0, text="🔎 Varrendo modalidades em paralelo no PNCP..."
                    )

                    def atualizar_progress(progresso: float, modalidade: int) -> None:
                        barra.progress(
                            progresso,
                            text=f"🔎 Modalidade {modalidade} concluída ({int(progresso * 100)}%)...",
                        )

                    registros, erros = consultar_editais(
                        inicio,
                        fim,
                        cnpj_orgao=cnpj_ativo,
                        modalidades_codigos=modalidades_edital_sel,
                        max_paginas=max_pags,
                        progress_callback=atualizar_progress,
                    )
                    todas_regs.extend(registros)
                    erros_consulta.extend(erros)
                    barra.empty()

            df = normalizar_pncp(todas_regs)
            df = deduplicar(df)

            if tipo_consulta != "Editais e Avisos de Contratações":
                df = filtrar_cnpj(df, cnpj_orgao=cnpj_ativo)

            inicio_ts = pd.Timestamp(inicio).normalize()
            fim_ts = pd.Timestamp(fim).normalize()
            df = filtrar_periodo(df, inicio_ts, fim_ts)

            barra_est = st.progress(0, text="⚙️ Estruturando registros...")

            def prog_estrutura(p: float, msg: str) -> None:
                barra_est.progress(p, text=f"⚙️ {msg}")

            dados_estruturados = estruturar_dados_registros(
                df,
                tipo_consulta,
                cnpj_orgao=cnpj_ativo,
                consultar_termos_api=st.session_state.consultar_termos_api,
                progress_callback=prog_estrutura,
            )
            barra_est.empty()

            st.session_state.df = df
            st.session_state.dados_totais = dados_estruturados
            st.session_state.tipo = tipo_consulta
            st.session_state.consulta_meta = {
                "inicio": inicio,
                "fim": fim,
                "cnpj": cnpj_ativo,
                "registros": len(df),
                "avisos": len(erros_consulta),
                "erros": list(erros_consulta),
                "carregado_em": dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }

            st.rerun()

        except Exception as exc:
            LOGGER.exception("Erro durante consulta PNCP")
            st.error(f"❌ Erro na consulta ao PNCP: {exc}")
            st.stop()

    # ==========================================
    # VERIFICAÇÃO DOS DADOS CARREGADOS
    # ==========================================
    df_bruto = st.session_state.df
    tipo_atual = st.session_state.tipo

    if df_bruto is None:
        st.markdown("### Bem-vindo ao Painel de Inteligência de Dados")
        st.markdown(
            "Esta ferramenta automatiza a coleta e a análise de dados públicos do Portal Nacional de Contratações "
            "Públicas (PNCP), estruturando as informações para facilitar a detecção de riscos e anomalias."
        )
        st.info(
            "👈 **Como começar sua análise:**\n\n"
            "1. **Escolha o Escopo:** No menu à esquerda, selecione se deseja visualizar Contratos, Atas ou Editais.\n"
            "2. **Defina o Período:** Ajuste as datas para refletir o intervalo de publicação que deseja investigar.\n"
            "3. **Gere o Dossiê:** Clique em **Carregar dados** e aguarde o processamento automático da matriz de risco."
        )
        st.stop()

    meta_consulta = st.session_state.get("consulta_meta") or {}
    erros_persistidos = meta_consulta.get("erros") or []

    if df_bruto.empty:
        if erros_persistidos:
            st.error(
                "❌ A consulta não produziu registros e apresentou falhas técnicas. "
                "O resultado é inconclusivo: não é correto afirmar que não existem registros no PNCP."
            )
            with st.expander("⚠️ Diagnóstico técnico da consulta", expanded=True):
                for erro in erros_persistidos[:30]:
                    st.write(f"- {erro}")
                if len(erros_persistidos) > 30:
                    st.caption(
                        f"{len(erros_persistidos) - 30} aviso(s) adicional(is) omitido(s)."
                    )
        else:
            st.info(
                f"ℹ️ Consulta concluída sem erro, mas nenhum registro foi retornado "
                f"para o escopo '{tipo_atual}' no período selecionado."
            )
        st.stop()

    dados_totais = st.session_state.get("dados_totais")
    if dados_totais is None or len(dados_totais) != len(df_bruto):
        with st.spinner("Estruturando os registros carregados..."):
            dados_totais = estruturar_dados_registros(
                df_bruto,
                tipo_atual,
                cnpj_orgao=st.session_state.cnpj_orgao,
                consultar_termos_api=st.session_state.consultar_termos_api,
            )
        st.session_state.dados_totais = dados_totais

    if not dados_totais:
        st.warning(
            "Não foi possível estruturar os registros retornados. Verifique se a base do PNCP está instável."
        )
        st.stop()

    if meta_consulta:
        st.caption(
            f"Última carga: {meta_consulta.get('carregado_em', 'N/D')} • "
            f"{meta_consulta.get('registros', len(df_bruto))} registro(s) • "
            f"{meta_consulta.get('avisos', 0)} aviso(s) de consulta"
        )

        if erros_persistidos:
            st.warning(
                "⚠️ A consulta retornou dados, mas houve falha em parte da coleta. "
                "Analise os avisos antes de considerar o conjunto como completo."
            )
            with st.expander("⚠️ Diagnóstico técnico da última consulta", expanded=False):
                for erro in erros_persistidos[:30]:
                    st.write(f"- {erro}")
                if len(erros_persistidos) > 30:
                    st.caption(
                        f"{len(erros_persistidos) - 30} aviso(s) adicional(is) omitido(s)."
                    )

        if (
            meta_consulta.get("inicio") != inicio
            or meta_consulta.get("fim") != fim
            or tipo_atual != tipo_selecionado
        ):
            st.info(
                "ℹ️ Os parâmetros da barra lateral foram alterados após a última carga. "
                "Clique em **Carregar dados** para atualizar a análise."
            )

    # ==========================================
    # SEGMENTAÇÃO E PESQUISA INSTANTÂNEA
    # ==========================================
    st.markdown("---")
    st.subheader("🗂️ Segmentação Estratégica e Pesquisa")

    col_seg1, col_seg2 = st.columns(2)

    todas_modalidades = sorted(
        {
            d["modalidade"]
            for d in dados_totais
            if d["modalidade"] and d["modalidade"] != "N/D"
        }
    )
    opcao_geral = "Visão Geral (Todas as Modalidades)"

    with col_seg1:
        mod_escolhida = st.selectbox(
            "Filtrar e Segmentar por Modalidade de Contratação:",
            [opcao_geral, *todas_modalidades],
            help="Filtre os resultados para isolar processos como Dispensas, Inexigibilidades, Pregões, etc.",
        )

    with col_seg2:
        termo_pesquisa = (
            st.text_input(
                "🔍 Pesquisa Livre (Objeto, Fornecedor, CNPJ, Número, Processo...):",
                "",
                help="Digite para buscar instantaneamente em todos os registros retornados.",
            )
            .strip()
            .casefold()
        )

    indices_ativos = []
    for i, registro in enumerate(dados_totais):
        if mod_escolhida != opcao_geral and registro["modalidade"] != mod_escolhida:
            continue
        if termo_pesquisa and termo_pesquisa not in registro.get("_busca", ""):
            continue
        indices_ativos.append(i)

    if not indices_ativos:
        mensagem_vazio = f"Nenhum registro encontrado para a modalidade '{mod_escolhida}'"
        if termo_pesquisa:
            mensagem_vazio += f" contendo o termo '{termo_pesquisa}'."
        st.warning(mensagem_vazio)
        st.stop()

    df_filtrado = df_bruto.iloc[indices_ativos].reset_index(drop=True)
    dados_processados = [dados_totais[i] for i in indices_ativos]

    # Cálculo O(N) de todos os riscos
    riscos = calcular_riscos_lote(dados_processados, tipo_atual)

    # Indicadores Gerais
    altos = sum(r["nivel"] == "🔴 ALTO" for r in riscos)
    medios = sum(r["nivel"] == "🟡 MÉDIO" for r in riscos)
    baixos = sum(r["nivel"] == "🟢 BAIXO" for r in riscos)

    valores_validos = [
        r["valor_num"]
        for r in dados_processados
        if r.get("valor_num") is not None and r["valor_num"] > 0
    ]
    valor_total = sum(valores_validos) if valores_validos else None

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Registros Segmentados", len(dados_processados))
    k2.metric("🔴 Alto risco", altos)
    k3.metric("🟡 Médio risco", medios)
    k4.metric("🟢 Baixo risco", baixos)
    k5.metric("Valor Analisado", moeda(valor_total))

    # Construção das tabelas principais
    rows = []
    tabela_vig = []

    for i_local, risco in enumerate(riscos):
        registro = dados_processados[i_local]

        fim_vig = registro["data_fim_vigencia"]
        dias_v = registro["dias_para_vencer"]
        if dias_v is not None:
            if dias_v < 0:
                vig_status = f"{fim_vig} (Vencido há {abs(dias_v)}d)"
                status_prazo = f"⚫ Encerrado ({abs(dias_v)}d atrás)"
            elif dias_v <= 30:
                vig_status = f"{fim_vig} ({dias_v}d restantes)"
                status_prazo = f"🔴 Crítico ({dias_v} dias)"
            elif dias_v <= 90:
                vig_status = f"{fim_vig} ({dias_v}d restantes)"
                status_prazo = f"🟡 Atenção ({dias_v} dias)"
            else:
                vig_status = f"{fim_vig} ({dias_v}d restantes)"
                status_prazo = f"🟢 Regular ({dias_v} dias)"
        else:
            vig_status = fim_vig
            status_prazo = "⚪ Não informado"

        perc_a = registro["perc_aditivo"]
        qtd_t = registro.get("qtd_termos_aditivos", 0)

        if perc_a is not None:
            adit_status = f"+{perc_a:.1f}%"
        elif qtd_t > 0:
            adit_status = f"{qtd_t} termo(s) no PNCP"
        elif registro["eh_aditivo"]:
            adit_status = "Aditivo formal"
        else:
            adit_status = "Inicial (0%)"

        rows.append(
            {
                "Índice": i_local,
                "Risco": risco["nivel"],
                "Score": risco["pontos"],
                "Número": registro["numero"],
                "Tipo": registro["tipo_instrumento"],
                "Modalidade": registro["modalidade"],
                "Fornecedor": registro["fornecedor"],
                "Valor": registro["valor"],
                "Fim Vigência": vig_status,
                "Aditivo": adit_status,
                "Objeto": registro["objeto"],
                "Gatilhos": "; ".join(risco["motivos"]),
                "Confiança": risco.get("confianca", "N/D"),
                "Link PNCP": registro.get("link_pncp", "https://pncp.gov.br"),
            }
        )

        tabela_vig.append(
            {
                "Número": registro["numero"],
                "Tipo": registro["tipo_instrumento"],
                "Fornecedor": registro["fornecedor"],
                "Data Assinatura": registro["data"],
                "Término Vigência": registro["data_fim_vigencia"],
                "Status do Prazo": status_prazo,
                "Valor": registro["valor"],
                "Link PNCP": registro.get("link_pncp", "https://pncp.gov.br"),
                "Objeto": registro["objeto"],
            }
        )

    df_risco = pd.DataFrame(rows)
    if not df_risco.empty:
        df_risco = df_risco.sort_values(
            ["Score", "Índice"],
            ascending=[False, True],
        ).reset_index(drop=True)

    df_vig_show = pd.DataFrame(tabela_vig)
    df_fornecedores = consolidar_fornecedores(dados_processados, riscos)

    # ==========================================
    # ABAS DA APLICAÇÃO
    # ==========================================
    st.markdown("---")
    tab1, tab_vigencia, tab_fornecedores, tab2 = st.tabs(
        [
            "🚦 1. Matriz de Risco Segmentada",
            "⏳ 2. Prazos e Aditamentos",
            "🏢 3. Fornecedores e Fracionamento",
            "🔎 4. Auditoria de Risco Individual",
        ]
    )

    config_coluna_link = {
        "Link PNCP": st.column_config.LinkColumn(
            "Portal PNCP",
            help="Clique para abrir a contratação na página oficial do PNCP",
            display_text="🔗 Acessar",
        )
    }

    # ========================================================
    # ABA 1 - MATRIZ DE RISCO PRINCIPAL
    # ========================================================
    with tab1:
        colunas_visualizacao = [
            c
            for c in (
                "Risco",
                "Score",
                "Confiança",
                "Número",
                "Tipo",
                "Modalidade",
                "Fornecedor",
                "Valor",
                "Fim Vigência",
                "Aditivo",
                "Link PNCP",
                "Objeto",
                "Gatilhos",
            )
            if c in df_risco.columns and (c != "Aditivo" or tipo_atual == "Contratos")
        ]

        st.dataframe(
            df_risco[colunas_visualizacao],
            use_container_width=True,
            hide_index=True,
            column_config=config_coluna_link,
        )

        st.markdown("#### 📥 Opções de Exportação")

        try:
            links_tupla = tuple(
                r.get("link_pncp", "https://pncp.gov.br") for r in dados_processados
            )
            excel = gerar_excel(
                df_filtrado,
                tipo_atual,
                df_risco,
                df_vig_show,
                df_fornecedores,
                links_tupla,
            )
            nome_arquivo = f"Auditoria_{slug(mod_escolhida)}.xlsx"

            col_download1, col_download2 = st.columns(2)

            with col_download1:
                st.download_button(
                    "📊 Baixar Planilha Completa de Auditoria (Excel)",
                    data=excel,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Gera um Excel com 4 abas: Matriz de Risco, Prazos, Ranking de Fornecedores e Dados Brutos.",
                )

            if DOCX_OK:
                try:
                    word_aba1 = gerar_word_dashboard_principal(
                        df_risco=df_risco,
                        tipo=tipo_atual,
                        inicio=inicio,
                        fim=fim,
                        modalidade=mod_escolhida,
                        total_segmentados=len(dados_processados),
                        altos=altos,
                        medios=medios,
                        valor_total=valor_total,
                        cnpj_orgao=st.session_state.cnpj_orgao,
                    )
                    nome_word_aba1 = f"Dashboard_Matriz_Risco_{slug(mod_escolhida)}.docx"

                    with col_download2:
                        st.download_button(
                            "📝 Baixar Relatório Executivo em Word",
                            data=word_aba1,
                            file_name=nome_word_aba1,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            help="Gera um relatório executivo formatado no Word para arquivamento ou apresentação.",
                        )
                except Exception as exc:
                    LOGGER.exception("Erro ao gerar Word do dashboard principal")
                    with col_download2:
                        st.error(f"Não foi possível gerar o Word: {exc}")
            else:
                with col_download2:
                    st.warning("python-docx não instalado para exportação Word.")

        except Exception as exc:
            LOGGER.exception("Erro ao gerar Excel")
            st.error(f"Não foi possível gerar o Excel: {exc}")

    # ========================================================
    # ABA 2 - VIGÊNCIA E ADITIVOS LEGAIS
    # ========================================================
    with tab_vigencia:
        if tipo_atual == "Contratos":
            st.subheader(
                "⏱️ Monitoramento Preditivo de Vigência & Limites da Lei 14.133/21"
            )
            st.markdown(
                "Rastreamento de contratos com prazo expirando para planejamento de novas contratações "
                "e verificação de aditamentos formais e financeiros (+25% em geral / +50% reformas)."
            )
        else:
            st.subheader("⏱️ Monitoramento Preditivo de Vigência")
            st.markdown(
                "Rastreamento de prazos expirando para planejamento tempestivo de novas licitações."
            )

        vencidos = [
            r
            for r in dados_processados
            if r.get("dias_para_vencer") is not None and r["dias_para_vencer"] < 0
        ]
        vencendo_30 = [
            r
            for r in dados_processados
            if r.get("dias_para_vencer") is not None
            and 0 <= r["dias_para_vencer"] <= 30
        ]
        vencendo_90 = [
            r
            for r in dados_processados
            if r.get("dias_para_vencer") is not None
            and 30 < r["dias_para_vencer"] <= 90
        ]

        if tipo_atual == "Contratos":
            total_aditivos = [
                r
                for r in dados_processados
                if r.get("eh_aditivo", False)
                or (r.get("qtd_termos_aditivos", 0) > 0)
                or (r.get("perc_aditivo") is not None and r["perc_aditivo"] > 0)
            ]

            m0, m1, m2, m3 = st.columns(4)
            m0.metric("⚫ Vigência Encerrada", len(vencidos))
            m1.metric("🚨 Vencem em até 30 dias", len(vencendo_30))
            m2.metric("⚠️ Vencem entre 31 e 90 dias", len(vencendo_90))
            m3.metric("📈 Termos Aditivos Identificados", len(total_aditivos))

            diagnosticos = [r.get("diagnostico_aditivo", {}) for r in dados_processados]

            total_contratos = len(diagnosticos)
            qtd_encontrados = sum(
                d.get("status") == "REGISTROS_ENCONTRADOS" for d in diagnosticos
            )
            qtd_sem_registros = sum(
                d.get("status") == "SEM_REGISTROS" for d in diagnosticos
            )
            qtd_sem_identificadores = sum(
                d.get("status") in {"SEM_IDENTIFICADORES", "NAO_CONSULTADO"}
                for d in diagnosticos
            )
            qtd_404 = sum(d.get("status") == "HTTP_404" for d in diagnosticos)
            qtd_erros = sum(
                d.get("status") in {"ERRO_CONSULTA", "RESPOSTA_INVALIDA"}
                for d in diagnosticos
            )

            st.markdown("### 🔎 Diagnóstico da Consulta de Aditivos")

            d1, d2, d3, d4, d5 = st.columns(5)
            d1.metric("Contratos analisados", total_contratos)
            d2.metric("🟢 Com termos", qtd_encontrados)
            d3.metric("⚪ Sem termos", qtd_sem_registros)
            d4.metric("🟡 Não consultados/Sem ID", qtd_sem_identificadores)
            d5.metric("🔴 Falhas/404", qtd_erros + qtd_404)

            if qtd_encontrados > 0:
                st.success(
                    f"✅ O PNCP retornou termos para {qtd_encontrados} contrato(s). "
                    "Esses registros foram considerados na análise de aditamentos."
                )

            if qtd_sem_registros > 0:
                st.info(
                    f"ℹ️ {qtd_sem_registros} contrato(s) foram consultados com sucesso "
                    "e não apresentaram termos retornados pelo PNCP."
                )
                st.markdown(
                    "ℹ️ **A ausência de termo no PNCP não comprova, isoladamente, a inexistência de "
                    "aditamento administrativo. Recomenda-se confrontar o resultado com os processos, "
                    "contratos e documentos mantidos pela Prefeitura.**"
                )

            if qtd_404 > 0:
                st.warning(
                    f"🟠 {qtd_404} consulta(s) retornaram HTTP 404. Isso significa que o "
                    "recurso de termos não foi localizado pelo endpoint consultado; não "
                    "é correto concluir apenas por isso que não existem aditivos."
                )

            if qtd_erros > 0:
                st.error(
                    f"🔴 {qtd_erros} consulta(s) apresentaram erro ou resposta inválida. "
                    "Não é possível usar esses casos para afirmar ausência de aditivos."
                )

            with st.expander("📋 Detalhes técnicos das consultas de termos", expanded=False):
                detalhes_diagnostico = []
                rotulos = {
                    "REGISTROS_ENCONTRADOS": "🟢 Termo(s) encontrado(s)",
                    "SEM_REGISTROS": "⚪ Consulta realizada — nenhum termo",
                    "SEM_IDENTIFICADORES": "🟡 Identificadores insuficientes",
                    "HTTP_404": "🟠 HTTP 404",
                    "ERRO_CONSULTA": "🔴 Erro na consulta",
                    "RESPOSTA_INVALIDA": "🔴 Resposta inválida",
                    "NAO_CONSULTADO": "⚪ Modo rápido (não consultado)",
                }
                for r in dados_processados:
                    diag = r.get("diagnostico_aditivo", {})
                    status = diag.get("status", "NAO_CONSULTADO")
                    detalhes_diagnostico.append(
                        {
                            "Contrato": r.get("numero", "N/D"),
                            "Ano": r.get("ano_contrato", "N/D"),
                            "Sequencial": r.get("sequencial_contrato", "N/D"),
                            "Situação": rotulos.get(status, status),
                            "HTTP": diag.get("http_status", ""),
                            "Termos": len(diag.get("termos", [])),
                            "Mensagem": diag.get("mensagem", ""),
                        }
                    )

                st.dataframe(
                    pd.DataFrame(detalhes_diagnostico),
                    use_container_width=True,
                    hide_index=True,
                )

            st.markdown("---")

            sub_tab_vig, sub_tab_adit = st.tabs(
                ["📅 Linha do Tempo de Vigências", "📊 Auditoria de Aditamentos"]
            )

            with sub_tab_vig:
                st.dataframe(
                    df_vig_show,
                    use_container_width=True,
                    hide_index=True,
                    column_config=config_coluna_link,
                )

            with sub_tab_adit:
                tabela_adit = []
                for r in dados_processados:
                    perc_a = r.get("perc_aditivo")
                    eh_adit = r.get("eh_aditivo", False)
                    qtd_t = r.get("qtd_termos_aditivos", 0)
                    v_ini = r.get("valor_inicial_num")
                    v_fim = r.get("valor_num")

                    if eh_adit or qtd_t > 0 or (perc_a is not None and perc_a > 0):
                        if perc_a is not None and perc_a > 50:
                            status_legal = "🔴 > 50% (Risco Alto)"
                        elif perc_a is not None and perc_a > 25:
                            status_legal = "🟡 > 25% (Risco Médio)"
                        elif perc_a is not None and perc_a > 0:
                            status_legal = "🟢 Aditivo Financeiro Regular"
                        elif qtd_t > 0:
                            status_legal = f"ℹ️ {qtd_t} Termo(s) Aditivo(s) no PNCP"
                        else:
                            status_legal = "ℹ️ Aditivo de Prazo / Qualitativo"

                        tabela_adit.append(
                            {
                                "Número": r["numero"],
                                "Tipo de Instrumento": r["tipo_instrumento"],
                                "Termos no PNCP": qtd_t,
                                "Fornecedor": r["fornecedor"],
                                "Valor Inicial": moeda(v_ini) if v_ini else "N/D",
                                "Valor Atualizado": moeda(v_fim),
                                "Variação (%)": (
                                    f"+{perc_a:.2f}%" if perc_a is not None else "0.00%"
                                ),
                                "Avaliação de Conformidade": status_legal,
                                "Link PNCP": r.get("link_pncp", "https://pncp.gov.br"),
                                "Objeto": r["objeto"],
                            }
                        )

                if tabela_adit:
                    df_adit_show = pd.DataFrame(tabela_adit)
                    st.dataframe(
                        df_adit_show,
                        use_container_width=True,
                        hide_index=True,
                        column_config=config_coluna_link,
                    )
                else:
                    st.info(
                        "ℹ️ Nenhum aditivo formal ou variação financeira positiva foi identificado nos contratos exibidos."
                    )

        else:
            m0, m1, m2 = st.columns(3)
            m0.metric("⚫ Vigência Encerrada", len(vencidos))
            m1.metric("🚨 Vencem em até 30 dias", len(vencendo_30))
            m2.metric("⚠️ Vencem entre 31 e 90 dias", len(vencendo_90))

            st.markdown("---")
            st.markdown("### 📅 Linha do Tempo de Vigências")
            st.dataframe(
                df_vig_show,
                use_container_width=True,
                hide_index=True,
                column_config=config_coluna_link,
            )

    # ========================================================
    # ABA 3 - CONCENTRAÇÃO DE FORNECEDORES E FRACIONAMENTO
    # ========================================================
    with tab_fornecedores:
        st.subheader("🏢 Inteligência de Fornecedores & Triagem de Fracionamento")
        st.markdown(
            "Visão consolidada da participação econômica de fornecedores e detecção automatizada "
            "de contratações com objetos semelhantes que possam requerer verificação sob a ótica do "
            "**art. 75, § 1º, da Lei nº 14.133/2021** (vedação ao fracionamento indevido de despesa)."
        )

        sub_forn, sub_frac = st.tabs(
            [
                "📊 Ranking e Concentração de Fornecedores",
                "🔍 Triagem Semântica de Possível Fracionamento",
            ]
        )

        with sub_forn:
            if df_fornecedores.empty:
                st.info("Não há dados suficientes de fornecedores neste segmento.")
            else:
                st.dataframe(
                    df_fornecedores,
                    use_container_width=True,
                    hide_index=True,
                )

        with sub_frac:
            if not SKLEARN_OK:
                st.warning(
                    "Instale `scikit-learn` para habilitar a detecção semântica de possível fracionamento."
                )
            else:
                limiar_frac = st.slider(
                    "Sensibilidade da similaridade textual entre objetos (%):",
                    min_value=20,
                    max_value=85,
                    value=38,
                    step=2,
                    help="Valores menores capturam mais pares; valores maiores exigem alta semelhança entre os objetos.",
                )
                df_frac = detectar_possiveis_fracionamentos(
                    dados_processados, similaridade_minima=limiar_frac / 100.0
                )
                if df_frac.empty:
                    st.success(
                        f"✅ Nenhum par de registros superou o limiar de {limiar_frac}% de similaridade semântica neste segmento."
                    )
                else:
                    st.warning(
                        f"⚠️ Foram localizados **{len(df_frac)} par(es)** de registros com alta proximidade "
                        "textual de objeto. Verifique o somatório no exercício financeiro e o ramo de atividade."
                    )
                    st.dataframe(df_frac, use_container_width=True, hide_index=True)

    # ========================================================
    # ABA 4 - AUDITORIA E SEMÂNTICA INDIVIDUAL
    # ========================================================
    with tab2:
        st.info(
            "💡 **Entenda os Alertas de Risco:**\n\n"
            "**• Valor anômalo acima do limite IQR:** Identifica discrepâncias financeiras quando o "
            "valor foge significativamente do padrão dos demais itens do mesmo segmento (outlier).\n\n"
            "**• Fornecedor concentrado:** Considera fornecedores com CNPJ identificado em pelo menos "
            "5 registros do conjunto, graduando quando atinge 25% ou 50% da amostra."
        )

        st.markdown("### Aprofundamento por Registro")

        f1, f2 = st.columns(2)

        filtro_risco = f1.multiselect(
            "Filtrar por Nível de Risco Alvo:",
            ["🔴 ALTO", "🟡 MÉDIO", "🟢 BAIXO"],
            default=["🔴 ALTO", "🟡 MÉDIO"],
            help="Selecione os níveis de risco que deseja analisar no detalhe.",
        )

        mostrar_outlier = f2.checkbox(
            "Isolar apenas anomalias financeiras (Outliers)",
            help="Se marcado, mostrará apenas os registros que ativaram o alerta de limite IQR financeiro.",
        )

        idx_auditoria = [
            i
            for i, risco in enumerate(riscos)
            if risco["nivel"] in filtro_risco
            and (not mostrar_outlier or risco["outlier"])
        ]

        if not idx_auditoria:
            st.warning(
                "Nenhum dossiê atende aos filtros de priorização ajustados acima."
            )
        else:
            mapa = {}
            for i in idx_auditoria:
                risco = riscos[i]
                registro = dados_processados[i]
                chave = f"[{risco['nivel']} | Score {risco['pontos']}] {registro['numero']} — {registro['valor']}"

                if chave in mapa:
                    contador = 2
                    chave_base = chave
                    while chave in mapa:
                        chave = f"{chave_base} (#{contador})"
                        contador += 1

                mapa[chave] = i

            escolhido = st.selectbox(
                "Selecione o dossiê da contratação para expandir:",
                list(mapa.keys()),
                help="Escolha um item da lista para visualizar a justificativa dos alertas.",
            )

            i_alvo = mapa[escolhido]
            registro = dados_processados[i_alvo]
            risco = riscos[i_alvo]

            with st.container():
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("Score de Risco", f"{risco['pontos']} pts")
                col_b.metric("Volume", registro["valor"])
                col_c.metric("Enquadramento Legal", registro["modalidade"])

                st.info(
                    f"**Objeto Descrito no Edital/Contrato:** {registro['objeto']}"
                )

                col_link_pncp, col_btn_dossie = st.columns([2, 1])
                with col_link_pncp:
                    if registro.get("link_pncp"):
                        st.markdown(
                            f"👉 [**Acessar dados oficiais e anexos no Portal PNCP**]({registro['link_pncp']})"
                        )

            with st.expander(
                "🔎 Justificativa Analítica e Roteiro de Auditoria", expanded=True
            ):
                st.markdown("📍 **Variáveis que acionaram o alerta no sistema:**")
                for motivo in risco["motivos"]:
                    st.write(f"- {motivo}")

                st.markdown("---")

                if risco["testes"]:
                    st.markdown("📋 **Roteiro de Auditoria Sugerido pelo Sistema:**")
                    for teste in risco["testes"]:
                        st.write(f"- {teste}")
                else:
                    st.caption(
                        "Nenhum teste específico foi sugerido automaticamente para este cenário."
                    )

            with st.expander(
                "📄 Informações de Registro e Partes Envolvidas", expanded=False
            ):
                info_col1, info_col2 = st.columns(2)
                info_col1.write(f"**Controle PNCP:** {registro['controle']}")
                info_col1.write(f"**Número:** {registro['numero']}")
                info_col1.write(
                    f"**Tipo de Instrumento:** {registro['tipo_instrumento']}"
                )
                info_col1.write(f"**Processo:** {registro['processo']}")
                info_col1.write(f"**Data Assinatura:** {registro['data']}")
                info_col1.write(
                    f"**Fim da Vigência:** {registro['data_fim_vigencia']}"
                )

                info_col2.write(f"**Fornecedor:** {registro['fornecedor']}")
                info_col2.write(
                    f"**CNPJ/CPF Fornecedor:** {cnpj_formatado(registro['cnpj_fornecedor'])}"
                )
                info_col2.write(f"**Situação:** {registro['situacao']}")
                info_col2.write(f"**Modalidade:** {registro['modalidade']}")

                if registro.get("termos_detalhes"):
                    st.markdown("---")
                    st.markdown("**Termos Aditivos / Apostilamentos Retornados pelo PNCP:**")
                    st.dataframe(
                        normalizar_pncp(registro["termos_detalhes"]),
                        use_container_width=True,
                        hide_index=True,
                    )

            st.markdown("### 🤖 Clusterização Semântica (Apenas neste Segmento)")

            sim_df = pd.DataFrame()
            if not SKLEARN_OK:
                st.warning(
                    "Scikit-Learn não está instalado. "
                    "Instale a dependência para habilitar a análise semântica automática de objetos similares."
                )
            else:
                sim_df = similares(dados_processados, i_alvo, limite=5, corte_minimo=0.15)
                if sim_df.empty:
                    st.info(
                        "Não foi possível identificar vizinhança semântica relevante (contratos com objetos similares) para este item."
                    )
                else:
                    st.dataframe(
                        sim_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config=config_coluna_link,
                    )

            if DOCX_OK:
                try:
                    bytes_dossie = gerar_word_dossie_individual(
                        registro=registro,
                        risco=risco,
                        df_similares=sim_df,
                        tipo=tipo_atual,
                        cnpj_orgao=st.session_state.cnpj_orgao,
                    )
                    with col_btn_dossie:
                        st.download_button(
                            "📄 Baixar Dossiê Individual (Word)",
                            data=bytes_dossie,
                            file_name=f"Dossie_{slug(registro['numero'])}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
                except Exception as exc:
                    LOGGER.warning("Falha ao gerar Word do dossiê individual: %s", exc)

    # ========================================================
    # METODOLOGIA DA CLASSIFICAÇÃO DE RISCO
    # ========================================================
    st.markdown("---")
    st.subheader("📘 Metodologia da Classificação de Risco")
    st.markdown(
        "A classificação abaixo foi desenvolvida como **instrumento de priorização do Controle Interno**. "
        "Ela combina sinais objetivos dos registros carregados e não representa probabilidade de fraude, "
        "irregularidade ou culpabilidade. Um nível de risco maior significa apenas que o registro merece "
        "maior prioridade para conferência humana."
    )

    met_col1, met_col2 = st.columns(2)

    with met_col1:
        st.markdown("**1. Materialidade — até 25 pontos**")
        st.markdown(
            "O valor é comparado com o conjunto analisado. Registros no quartil superior (Q75) recebem "
            "**8 pontos** e registros no percentil 90 (Q90) recebem **15 pontos**. Há ainda um complemento "
            "conservador para valores absolutos muito elevados: **+5 pontos acima de R$ 1 milhão** ou "
            "**+10 pontos acima de R$ 5 milhões**."
        )

        st.markdown("**2. Anomalia financeira — 20 pontos**")
        st.markdown(
            "É aplicado o **IQR (Intervalo Interquartil)**. Quando o valor supera o limite superior estatístico "
            "do conjunto, o registro é marcado como outlier e recebe **20 pontos**."
        )

        st.markdown("**3. Modalidade e situação emergencial — até 30 pontos**")
        st.markdown(
            "Dispensa ou inexigibilidade geram **12 pontos** como fator de atenção, pois exigem conferência "
            "do fundamento legal. Um indício de emergência acrescenta **18 pontos**."
        )

        st.markdown("**4. Vigência — até 20 pontos**")
        st.markdown(
            "Contrato vencido: **20 pontos**; vencimento em até 15 dias: **15 pontos**; até 30 dias: "
            "**10 pontos**; entre 31 e 90 dias: **4 pontos**."
        )

    with met_col2:
        st.markdown("**5. Aditamentos em contratos — até 30 pontos**")
        st.markdown(
            "A variação financeira identificada pontua: **+3 pontos** até 10%; **+7 pontos** acima de 10%; "
            "**+16 pontos** acima de 25%; **+25 pontos** acima de 50%. A identificação de termo aditivo "
            "acrescenta **5 pontos**."
        )

        st.markdown("**6. Concentração de fornecedor — até 12 pontos**")
        st.markdown(
            "O sistema exige pelo menos **5 registros com CNPJ identificado** e presença mínima de 3 ocorrências "
            "do fornecedor. Participação de **25% ou mais** gera **7 pontos**; **50% ou mais**, **12 pontos**."
        )

        st.markdown("**7. Qualidade dos dados — não altera o score**")
        st.markdown(
            "Campos ausentes, como objeto, controle PNCP ou data, aparecem como alerta de qualidade cadastral, "
            "mas **não aumentam a pontuação de risco**."
        )

        st.markdown("**8. Faixas finais e confiança**")
        st.markdown(
            "**🔴 ALTO:** 60 pontos ou mais · **🟡 MÉDIO:** 35 a 59 pontos · **🟢 BAIXO:** abaixo de 35 pontos."
        )

    st.info(
        "⚠️ **Como interpretar:** o score é uma regra heurística de priorização. Ele não substitui exame do "
        "processo administrativo, documentos, pesquisa de preços, justificativas, execução contratual ou análise "
        "jurídica."
    )

    # Rodapé
    st.markdown("---")
    st.caption(
        f"Painel de apoio ao Controle Interno — {PREFEITURA}. "
        "Dados públicos consultados de forma ativa via Integração API PNCP. "
        "Alertas automatizados não substituem análise técnica, jurídica ou auditoria humana."
    )


if __name__ == "__main__":
    main()
