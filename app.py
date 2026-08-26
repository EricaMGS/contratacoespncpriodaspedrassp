import io
import datetime
import time
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor
from fpdf import FPDF


# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Portal PNCP - Rio das Pedras/SP",
    page_icon="🏛️",
    layout="wide",
)


# ============================================================
# DADOS DO MUNICÍPIO
# ============================================================

CNPJ_RIO_DAS_PEDRAS = "44826840000183"
CODIGO_IBGE_RIO_DAS_PEDRAS = "3544004"
UF = "SP"

BASE_URL = "https://pncp.gov.br/api/consulta/v1"

NOME_MUNICIPIO = "Rio das Pedras/SP"
NOME_PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"


# ============================================================
# CONFIGURAÇÕES DA API
# ============================================================

# O PNCP trabalha com limites diferentes conforme o endpoint.
# Mantemos valores seguros para evitar HTTP 400.
TAMANHO_PAGINA_CONTRATOS = 100
TAMANHO_PAGINA_ATAS = 100
TAMANHO_PAGINA_EDITAIS = 50

# Segurança contra consultas excessivamente grandes.
MAX_PAGINAS = 100

# Número de tentativas para erros transitórios.
MAX_TENTATIVAS = 4

# Timeouts:
# conexão relativamente curta + leitura mais longa.
TIMEOUT_CONEXAO = 20
TIMEOUT_LEITURA = 90


# ============================================================
# HEADERS
# ============================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Connection": "keep-alive",
}


# ============================================================
# FUNÇÕES GERAIS DE FORMATAÇÃO
# ============================================================

def texto_valido(valor: Any, padrao: str = "N/D") -> str:
    """
    Converte qualquer valor para texto seguro.
    """
    if valor is None:
        return padrao

    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass

    if isinstance(valor, dict):
        for chave in (
            "nome",
            "razaoSocial",
            "descricao",
            "valor",
            "nomeUnidade",
            "municipioNome",
        ):
            if chave in valor:
                resultado = texto_valido(valor.get(chave), "")
                if resultado:
                    return resultado

        return str(valor)

    if isinstance(valor, list):
        if not valor:
            return padrao

        partes = [
            texto_valido(item, "")
            for item in valor
        ]

        partes = [
            item for item in partes
            if item
        ]

        return ", ".join(partes) if partes else padrao

    texto = str(valor).strip()

    if texto.lower() in {
        "",
        "none",
        "nan",
        "null",
        "n/d",
        "nd",
        "nat",
    }:
        return padrao

    return texto


def valor_numerico(valor: Any) -> Optional[float]:
    """
    Converte valores monetários/númericos vindos do PNCP.
    Aceita:
      1234.56
      "1234.56"
      "1.234,56"
      "R$ 1.234,56"
    """

    if valor is None:
        return None

    if isinstance(valor, bool):
        return None

    if isinstance(valor, (int, float)):
        try:
            if pd.isna(valor):
                return None
            return float(valor)
        except Exception:
            return None

    if isinstance(valor, dict):
        for chave in (
            "valor",
            "value",
            "valorTotal",
            "valorGlobal",
            "valorInicial",
        ):
            if chave in valor:
                resultado = valor_numerico(valor[chave])
                if resultado is not None:
                    return resultado

        return None

    texto = str(valor).strip()

    if not texto:
        return None

    texto = (
        texto
        .replace("R$", "")
        .replace("r$", "")
        .replace(" ", "")
    )

    # Primeiro tenta como número normal.
    try:
        return float(texto)
    except Exception:
        pass

    # Formato brasileiro.
    try:
        if "," in texto and "." in texto:
            texto = texto.replace(".", "")
            texto = texto.replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")

        return float(texto)
    except Exception:
        return None


def formatar_moeda_br(valor: Any) -> str:
    """
    Formata valor no padrão brasileiro.
    """

    numero = valor_numerico(valor)

    if numero is None:
        return "N/D"

    texto = f"{numero:,.2f}"

    texto = (
        texto
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    return f"R$ {texto}"


def formatar_data(valor: Any) -> str:
    """
    Converte datas do PNCP para DD/MM/AAAA.
    """

    if valor is None:
        return "N/D"

    texto = str(valor).strip()

    if not texto:
        return "N/D"

    if texto.lower() in {
        "none",
        "nan",
        "null",
        "n/d",
        "nat",
    }:
        return "N/D"

    try:
        data = pd.to_datetime(
            valor,
            errors="coerce",
        )

        if pd.isna(data):
            return texto

        return data.strftime("%d/%m/%Y")

    except Exception:
        return texto


def normalizar_cnpj(valor: Any) -> str:
    """
    Remove máscara de CNPJ/CPF.
    """

    return re.sub(
        r"\D",
        "",
        texto_valido(valor, ""),
    )


# ============================================================
# EXTRAÇÃO ROBUSTA DE CAMPOS
# ============================================================

def extrair_valor_recursivo(
    valor: Any,
    profundidade: int = 0,
) -> Any:
    """
    Tenta extrair texto útil de estruturas aninhadas.
    """

    if profundidade > 4:
        return valor

    if valor is None:
        return None

    if isinstance(valor, dict):

        # Chaves prioritárias.
        chaves_prioritarias = [
            "nome",
            "razaoSocial",
            "descricao",
            "descricaoObjeto",
            "valor",
            "numero",
            "cnpj",
            "ni",
            "numeroDocumento",
            "nomeUnidade",
            "municipioNome",
            "ufSigla",
        ]

        for chave in chaves_prioritarias:

            if chave in valor:

                resultado = extrair_valor_recursivo(
                    valor[chave],
                    profundidade + 1,
                )

                if resultado not in (
                    None,
                    "",
                    "N/D",
                ):
                    return resultado

        return valor

    if isinstance(valor, list):

        resultados = []

        for item in valor:

            resultado = extrair_valor_recursivo(
                item,
                profundidade + 1,
            )

            if resultado not in (
                None,
                "",
                "N/D",
            ):
                resultados.append(
                    str(resultado)
                )

        if resultados:
            return ", ".join(resultados)

        return None

    return valor


def obter_primeiro_valor(
    row: Any,
    campos: List[str],
    padrao: Any = "N/D",
) -> Any:
    """
    Procura o primeiro campo realmente preenchido.

    Permite trabalhar com diferentes versões/
    estruturas retornadas pelo PNCP.
    """

    if row is None:
        return padrao

    for campo in campos:

        try:

            if isinstance(row, pd.Series):
                if campo not in row.index:
                    continue

                valor = row.get(campo)

            elif isinstance(row, dict):
                if campo not in row:
                    continue

                valor = row.get(campo)

            else:
                continue

        except Exception:
            continue

        if valor is None:
            continue

        try:
            if pd.isna(valor):
                continue
        except Exception:
            pass

        valor = extrair_valor_recursivo(valor)

        if valor is None:
            continue

        texto = texto_valido(
            valor,
            "",
        )

        if texto:
            return valor

    return padrao


# ============================================================
# EXTRAÇÃO DE DADOS DO REGISTRO
# ============================================================

def obter_dados_registro(
    row: Any,
    tipo: str,
) -> Dict[str, str]:
    """
    Extrai os principais campos de Contratos,
    Atas e Editais.

    A função utiliza vários nomes alternativos
    para reduzir a ocorrência de N/D.
    """

    id_pncp = obter_primeiro_valor(
        row,
        [
            "numeroControlePNCP",
            "numeroControlePNCPAta",
            "numeroControlePNCPCompra",
            "numeroControlePncp",
            "idContratoPNCP",
            "idContratacaoPNCP",
        ],
    )

    # ========================================================
    # ATAS
    # ========================================================

    if tipo == "Atas de Registro de Preços":

        numero = obter_primeiro_valor(
            row,
            [
                "numeroAtaRegistroPreco",
                "numeroAtaRegistroPrecos",
                "numeroAta",
                "numeroRegistroPreco",
                "numero",
            ],
        )

        processo = obter_primeiro_valor(
            row,
            [
                "processo",
                "numeroProcesso",
                "processoAdministrativo",
                "numeroProcessoAdministrativo",
                "numeroProcessoCompra",
            ],
        )

        objeto = obter_primeiro_valor(
            row,
            [
                "objetoCompra",
                "objetoAta",
                "objeto",
                "descricaoObjeto",
                "descricao",
                "objetoContratacao",
            ],
        )

        fornecedor = obter_primeiro_valor(
            row,
            [
                "nomeRazaoSocialFornecedor",
                "razaoSocialFornecedor",
                "nomeFornecedor",
                "fornecedorNome",
                "razaoSocial",
                "nomeRazaoSocial",
                "fornecedor",
            ],
        )

        cnpj_fornecedor = obter_primeiro_valor(
            row,
            [
                "niFornecedor",
                "cnpjFornecedor",
                "numeroDocumentoFornecedor",
                "documentoFornecedor",
                "cpfCnpjFornecedor",
                "ni",
            ],
        )

        valor = obter_primeiro_valor(
            row,
            [
                "valorTotal",
                "valorGlobal",
                "valorTotalAta",
                "valorAta",
                "valorTotalEstimado",
                "valorTotalHomologado",
                "valor",
            ],
            padrao=None,
        )

        data_assinatura = obter_primeiro_valor(
            row,
            [
                "dataAssinatura",
                "dataAssinaturaAta",
                "dataCelebracao",
                "dataFormalizacao",
            ],
        )

        vigencia_inicio = obter_primeiro_valor(
            row,
            [
                "vigenciaInicio",
                "dataInicioVigencia",
                "dataVigenciaInicio",
                "inicioVigencia",
            ],
        )

        vigencia_fim = obter_primeiro_valor(
            row,
            [
                "vigenciaFim",
                "dataFimVigencia",
                "dataVigenciaFim",
                "fimVigencia",
            ],
        )

        situacao = obter_primeiro_valor(
            row,
            [
                "situacao",
                "situacaoAta",
                "status",
                "situacaoRegistro",
                "cancelado",
            ],
        )

        orgao = obter_primeiro_valor(
            row,
            [
                "orgaoEntidade",
                "nomeOrgao",
                "razaoSocialOrgao",
                "orgao",
                "orgaoNome",
            ],
        )

        unidade = obter_primeiro_valor(
            row,
            [
                "unidadeOrgao",
                "nomeUnidade",
                "unidade",
                "unidadeAdministrativa",
            ],
        )

        return {
            "id_pncp": texto_valido(id_pncp),
            "numero": texto_valido(numero),
            "processo": texto_valido(processo),
            "objeto": texto_valido(objeto),
            "fornecedor": texto_valido(fornecedor),
            "cnpj_fornecedor": texto_valido(
                cnpj_fornecedor
            ),
            "valor": formatar_moeda_br(valor),
            "data_assinatura": formatar_data(
                data_assinatura
            ),
            "vigencia_inicio": formatar_data(
                vigencia_inicio
            ),
            "vigencia_fim": formatar_data(
                vigencia_fim
            ),
            "situacao": texto_valido(situacao),
            "orgao": texto_valido(orgao),
            "unidade": texto_valido(unidade),
        }

    # ========================================================
    # CONTRATOS
    # ========================================================

    if tipo == "Contratos":

        numero = obter_primeiro_valor(
            row,
            [
                "numeroContratoEmpenho",
                "numeroContrato",
                "numeroContratoPncp",
                "numero",
            ],
        )

        processo = obter_primeiro_valor(
            row,
            [
                "processo",
                "numeroProcesso",
                "processoAdministrativo",
            ],
        )

        objeto = obter_primeiro_valor(
            row,
            [
                "objetoContrato",
                "objetoCompra",
                "objeto",
                "descricaoObjeto",
            ],
        )

        fornecedor = obter_primeiro_valor(
            row,
            [
                "nomeRazaoSocialFornecedor",
                "razaoSocialFornecedor",
                "nomeFornecedor",
                "fornecedorNome",
                "razaoSocial",
                "nomeRazaoSocial",
            ],
        )

        cnpj_fornecedor = obter_primeiro_valor(
            row,
            [
                "niFornecedor",
                "cnpjFornecedor",
                "numeroDocumentoFornecedor",
                "documentoFornecedor",
                "cpfCnpjFornecedor",
            ],
        )

        valor = obter_primeiro_valor(
            row,
            [
                "valorGlobal",
                "valorInicial",
                "valorTotal",
                "valorContrato",
            ],
            padrao=None,
        )

        data_assinatura = obter_primeiro_valor(
            row,
            [
                "dataAssinatura",
                "dataCelebracao",
            ],
        )

        vigencia_inicio = obter_primeiro_valor(
            row,
            [
                "dataVigenciaInicio",
                "vigenciaInicio",
            ],
        )

        vigencia_fim = obter_primeiro_valor(
            row,
            [
                "dataVigenciaFim",
                "vigenciaFim",
            ],
        )

        situacao = obter_primeiro_valor(
            row,
            [
                "situacao",
                "status",
            ],
        )

        orgao = obter_primeiro_valor(
            row,
            [
                "orgaoEntidade",
                "nomeOrgao",
                "orgao",
            ],
        )

        unidade = obter_primeiro_valor(
            row,
            [
                "unidadeOrgao",
                "nomeUnidade",
                "unidade",
            ],
        )

        return {
            "id_pncp": texto_valido(id_pncp),
            "numero": texto_valido(numero),
            "processo": texto_valido(processo),
            "objeto": texto_valido(objeto),
            "fornecedor": texto_valido(fornecedor),
            "cnpj_fornecedor": texto_valido(
                cnpj_fornecedor
            ),
            "valor": formatar_moeda_br(valor),
            "data_assinatura": formatar_data(
                data_assinatura
            ),
            "vigencia_inicio": formatar_data(
                vigencia_inicio
            ),
            "vigencia_fim": formatar_data(
                vigencia_fim
            ),
            "situacao": texto_valido(situacao),
            "orgao": texto_valido(orgao),
            "unidade": texto_valido(unidade),
        }

    # ========================================================
    # EDITAIS / CONTRATAÇÕES
    # ========================================================

    processo = obter_primeiro_valor(
        row,
        [
            "processo",
            "numeroProcesso",
            "processoAdministrativo",
        ],
    )

    objeto = obter_primeiro_valor(
        row,
        [
            "objetoCompra",
            "objeto",
            "descricaoObjeto",
            "descricao",
        ],
    )

    valor = obter_primeiro_valor(
        row,
        [
            "valorTotalHomologado",
            "valorTotalEstimado",
            "valorEstimado",
            "valorTotal",
            "valor",
        ],
        padrao=None,
    )

    numero = obter_primeiro_valor(
        row,
        [
            "numeroCompra",
            "numeroEdital",
            "numero",
        ],
    )

    fornecedor = obter_primeiro_valor(
        row,
        [
            "nomeRazaoSocialFornecedor",
            "razaoSocialFornecedor",
            "nomeFornecedor",
        ],
    )

    cnpj_fornecedor = obter_primeiro_valor(
        row,
        [
            "niFornecedor",
            "cnpjFornecedor",
        ],
    )

    data_publicacao = obter_primeiro_valor(
        row,
        [
            "dataPublicacao",
            "dataPublicacaoPncp",
            "dataAberturaProposta",
            "dataInclusao",
        ],
    )

    situacao = obter_primeiro_valor(
        row,
        [
            "situacaoCompra",
            "situacao",
            "status",
        ],
    )

    orgao = obter_primeiro_valor(
        row,
        [
            "orgaoEntidade",
            "nomeOrgao",
            "orgao",
        ],
    )

    unidade = obter_primeiro_valor(
        row,
        [
            "unidadeOrgao",
            "nomeUnidade",
            "unidade",
        ],
    )

    return {
        "id_pncp": texto_valido(id_pncp),
        "numero": texto_valido(numero),
        "processo": texto_valido(processo),
        "objeto": texto_valido(objeto),
        "fornecedor": texto_valido(fornecedor),
        "cnpj_fornecedor": texto_valido(
            cnpj_fornecedor
        ),
        "valor": formatar_moeda_br(valor),
        "data_assinatura": formatar_data(
            data_publicacao
        ),
        "vigencia_inicio": "N/D",
        "vigencia_fim": "N/D",
        "situacao": texto_valido(situacao),
        "orgao": texto_valido(orgao),
        "unidade": texto_valido(unidade),
    }


# ============================================================
# IDENTIFICAÇÃO DE CAMPOS PNCP
# ============================================================

def obter_identificador_contrato(
    row: Any,
) -> Tuple[Optional[str], Optional[int], Optional[int]]:
    """
    Tenta recuperar:
      CNPJ
      ano
      sequencial

    necessários para consultar um contrato específico.
    """

    cnpj = obter_primeiro_valor(
        row,
        [
            "cnpjOrgao",
            "cnpj",
            "cnpjCompra",
            "cnpjOrgaoEntidade",
        ],
        padrao=None,
    )

    ano = obter_primeiro_valor(
        row,
        [
            "anoContrato",
            "anoContratoEmpenho",
            "ano",
        ],
        padrao=None,
    )

    sequencial = obter_primeiro_valor(
        row,
        [
            "sequencialContrato",
            "sequencialContratoEmpenho",
            "sequencial",
        ],
        padrao=None,
    )

    cnpj_limpo = normalizar_cnpj(cnpj)

    try:
        ano_int = int(ano)
    except Exception:
        ano_int = None

    try:
        seq_int = int(sequencial)
    except Exception:
        seq_int = None

    if len(cnpj_limpo) != 14:
        cnpj_limpo = None

    return (
        cnpj_limpo,
        ano_int,
        seq_int,
    )


# ============================================================
# EXTRAÇÃO DO ENVELOPE DA API
# ============================================================

def extrair_registros(
    data: Any,
) -> List[Dict[str, Any]]:
    """
    Extrai a lista de registros de diferentes formatos
    possíveis de resposta do PNCP.
    """

    if data is None:
        return []

    if isinstance(data, list):
        return data

    if not isinstance(data, dict):
        return []

    # Formato paginado padrão.
    for chave in (
        "data",
        "items",
        "content",
        "dados",
        "registros",
    ):

        valor = data.get(chave)

        if isinstance(valor, list):
            return valor

    # Alguns retornos podem trazer uma lista em agrupadores.
    for valor in data.values():

        if isinstance(valor, list):
            return valor

    return []


def obter_informacoes_paginacao(
    data: Any,
) -> Dict[str, Any]:
    """
    Recupera metadados de paginação.
    """

    if not isinstance(data, dict):
        return {
            "totalRegistros": None,
            "totalPaginas": None,
            "numeroPagina": None,
            "paginasRestantes": None,
            "empty": False,
        }

    def numero(chave):
        valor = data.get(chave)

        try:
            return int(valor)
        except Exception:
            return None

    return {
        "totalRegistros": numero(
            "totalRegistros"
        ),
        "totalPaginas": numero(
            "totalPaginas"
        ),
        "numeroPagina": numero(
            "numeroPagina"
        ),
        "paginasRestantes": numero(
            "paginasRestantes"
        ),
        "empty": bool(
            data.get("empty", False)
        ),
    }


# ============================================================
# TRATAMENTO DO DATAFRAME
# ============================================================

def tratar_dataframe(
    df: pd.DataFrame,
) -> pd.DataFrame:

    if df is None or df.empty:
        return pd.DataFrame()

    df_tratado = df.copy()

    for coluna in df_tratado.columns:

        def converter(valor):

            if isinstance(valor, dict):

                resultado = extrair_valor_recursivo(
                    valor
                )

                if isinstance(
                    resultado,
                    (dict, list),
                ):
                    return str(resultado)

                return texto_valido(
                    resultado
                )

            if isinstance(valor, list):

                resultado = extrair_valor_recursivo(
                    valor
                )

                return texto_valido(
                    resultado
                )

            return valor

        try:
            df_tratado[coluna] = (
                df_tratado[coluna]
                .map(converter)
            )
        except Exception:
            pass

    return df_tratado


# ============================================================
# CLIENTE HTTP PNCP
# ============================================================

@st.cache_resource
def criar_sessao_http():

    sessao = requests.Session()

    sessao.headers.update(
        HEADERS
    )

    return sessao


def mensagem_erro_http(
    response: requests.Response,
) -> str:

    try:
        data = response.json()

        if isinstance(data, dict):

            mensagem = (
                data.get("message")
                or data.get("error")
                or data.get("detail")
            )

            if mensagem:
                return str(mensagem)

    except Exception:
        pass

    texto = response.text.strip()

    if not texto:
        texto = "Sem detalhes fornecidos pela API."

    return texto[:500]


def consultar_pncp(
    url: str,
    params: Dict[str, Any],
    max_tentativas: int = MAX_TENTATIVAS,
) -> Any:
    """
    Consulta o PNCP com:
      - timeout controlado
      - retry
      - backoff
      - tratamento de 204
      - tratamento diferenciado de 400/422
    """

    sessao = criar_sessao_http()

    ultimo_erro = None

    for tentativa in range(
        1,
        max_tentativas + 1,
    ):

        try:

            response = sessao.get(
                url,
                params=params,
                timeout=(
                    TIMEOUT_CONEXAO,
                    TIMEOUT_LEITURA,
                ),
            )

            # ------------------------------------------------
            # SUCESSO
            # ------------------------------------------------

            if response.status_code == 200:

                try:
                    return response.json()

                except ValueError as erro:
                    raise RuntimeError(
                        "O PNCP respondeu HTTP 200, "
                        "mas o conteúdo não é um JSON válido."
                    ) from erro

            # ------------------------------------------------
            # SEM RESULTADOS
            # ------------------------------------------------

            if response.status_code == 204:
                return []

            # ------------------------------------------------
            # ERROS DE PARÂMETROS
            # Não adianta repetir a requisição.
            # ------------------------------------------------

            if response.status_code in (
                400,
                422,
            ):

                detalhe = mensagem_erro_http(
                    response
                )

                raise RuntimeError(
                    f"PNCP rejeitou os parâmetros "
                    f"(HTTP {response.status_code}). "
                    f"{detalhe}"
                )

            # ------------------------------------------------
            # NÃO ENCONTRADO
            # ------------------------------------------------

            if response.status_code == 404:

                raise RuntimeError(
                    "Endpoint ou recurso não encontrado "
                    "no PNCP (HTTP 404)."
                )

            # ------------------------------------------------
            # RATE LIMIT / SERVIDOR
            # ------------------------------------------------

            if response.status_code in (
                408,
                429,
                500,
                502,
                503,
                504,
            ):

                detalhe = mensagem_erro_http(
                    response
                )

                ultimo_erro = RuntimeError(
                    f"HTTP {response.status_code}: "
                    f"{detalhe}"
                )

                if tentativa < max_tentativas:

                    espera = min(
                        2 ** tentativa,
                        20,
                    )

                    time.sleep(
                        espera
                    )

                    continue

                raise ultimo_erro

            # ------------------------------------------------
            # OUTROS
            # ------------------------------------------------

            detalhe = mensagem_erro_http(
                response
            )

            raise RuntimeError(
                f"PNCP retornou HTTP "
                f"{response.status_code}: "
                f"{detalhe}"
            )

        except (
            requests.exceptions.Timeout,
            requests.exceptions.ConnectionError,
        ) as erro:

            ultimo_erro = erro

            if tentativa < max_tentativas:

                espera = min(
                    2 ** tentativa,
                    20,
                )

                time.sleep(
                    espera
                )

                continue

            raise RuntimeError(
                "Não foi possível obter resposta do PNCP "
                "após várias tentativas. "
                "O portal pode estar lento ou temporariamente "
                "indisponível."
            ) from erro

    raise RuntimeError(
        "Falha inesperada na comunicação com o PNCP."
    )


# ============================================================
# PAGINAÇÃO CONTROLADA
# ============================================================

def consultar_paginas(
    url: str,
    params: Dict[str, Any],
    max_paginas: int = MAX_PAGINAS,
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Paginação sequencial e controlada.

    Usa totalPaginas/paginasRestantes quando disponíveis.
    Não dispara várias requisições simultaneamente.
    """

    todos_registros = []

    pagina = 1
    paginas_consultadas = 0

    while pagina <= max_paginas:

        params_pagina = params.copy()

        params_pagina["pagina"] = pagina

        dados = consultar_pncp(
            url,
            params_pagina,
        )

        paginas_consultadas += 1

        registros = extrair_registros(
            dados
        )

        if not registros:
            break

        todos_registros.extend(
            registros
        )

        info = obter_informacoes_paginacao(
            dados
        )

        total_paginas = info.get(
            "totalPaginas"
        )

        paginas_restantes = info.get(
            "paginasRestantes"
        )

        numero_pagina = info.get(
            "numeroPagina"
        )

        # ----------------------------------------------------
        # Melhor situação:
        # API informa total de páginas.
        # ----------------------------------------------------

        if total_paginas is not None:

            if pagina >= total_paginas:
                break

        # ----------------------------------------------------
        # Algumas respostas informam páginas restantes.
        # ----------------------------------------------------

        elif paginas_restantes is not None:

            if paginas_restantes <= 0:
                break

        # ----------------------------------------------------
        # Se a API informa número da página atual.
        # ----------------------------------------------------

        elif (
            numero_pagina is not None
            and numero_pagina < pagina
        ):
            break

        # ----------------------------------------------------
        # Fallback:
        # se veio menos que o tamanho solicitado,
        # normalmente acabou a paginação.
        # ----------------------------------------------------

        else:

            tamanho = int(
                params.get(
                    "tamanhoPagina",
                    50,
                )
            )

            if len(registros) < tamanho:
                break

        pagina += 1

        # Pequeno intervalo para evitar rajadas.
        if pagina <= max_paginas:
            time.sleep(0.15)

    return (
        todos_registros,
        paginas_consultadas,
    )


# ============================================================
# CACHE DE CONSULTAS
# ============================================================

@st.cache_data(
    ttl=600,
    show_spinner=False,
)
def executar_consulta_cacheada(
    url: str,
    params_tuple: Tuple[Tuple[str, str], ...],
    max_paginas: int,
) -> Tuple[List[Dict[str, Any]], int]:

    params = dict(
        params_tuple
    )

    return consultar_paginas(
        url,
        params,
        max_paginas,
    )


def executar_consulta(
    url: str,
    params: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], int]:

    params_serializados = tuple(
        sorted(
            (
                str(chave),
                str(valor),
            )
            for chave, valor
            in params.items()
        )
    )

    return executar_consulta_cacheada(
        url,
        params_serializados,
        MAX_PAGINAS,
    )


# ============================================================
# REMOVER DUPLICIDADES
# ============================================================

def remover_duplicidades(
    df: pd.DataFrame,
) -> pd.DataFrame:

    if df.empty:
        return df

    resultado = df.copy()

    colunas_identificacao = [
        "numeroControlePNCP",
        "numeroControlePNCPAta",
        "numeroControlePNCPCompra",
        "numeroContratoEmpenho",
        "numeroContrato",
        "numeroAtaRegistroPreco",
        "idContratoPNCP",
    ]

    for coluna in colunas_identificacao:

        if coluna in resultado.columns:

            serie = (
                resultado[coluna]
                .astype(str)
                .str.strip()
            )

            validos = (
                serie
                .replace(
                    {
                        "nan": "",
                        "None": "",
                        "N/D": "",
                    }
                )
                != ""
            )

            if validos.any():

                resultado = (
                    resultado
                    .loc[
                        ~validos
                        | ~serie.duplicated(
                            keep="first"
                        )
                    ]
                )

                return resultado.reset_index(
                    drop=True
                )

    return (
        resultado
        .drop_duplicates()
        .reset_index(drop=True)
    )


# ============================================================
# FILTRO EXTRA DE SEGURANÇA
# ============================================================

def filtrar_por_cnpj(
    df: pd.DataFrame,
    cnpj: str,
) -> pd.DataFrame:

    if df.empty:
        return df

    cnpj_limpo = normalizar_cnpj(
        cnpj
    )

    if not cnpj_limpo:
        return df

    possiveis_colunas = [
        "cnpjOrgao",
        "cnpj",
        "cnpjCompra",
        "cnpjOrgaoEntidade",
        "orgaoEntidade",
        "orgao",
        "unidadeOrgao",
    ]

    encontrou_coluna = False

    for coluna in possiveis_colunas:

        if coluna not in df.columns:
            continue

        encontrou_coluna = True

        serie = (
            df[coluna]
            .astype(str)
            .map(normalizar_cnpj)
        )

        mascara = (
            serie == cnpj_limpo
        )

        if mascara.any():

            return (
                df.loc[mascara]
                .reset_index(drop=True)
            )

    # Se não foi possível identificar
    # uma coluna de CNPJ, mantém os dados.
    if not encontrou_coluna:
        return df

    return df


# ============================================================
# DETALHES DO CONTRATO
# ============================================================

def consultar_detalhes_contrato(
    row: Any,
) -> List[Dict[str, Any]]:

    cnpj, ano, sequencial = (
        obter_identificador_contrato(
            row
        )
    )

    if not all(
        [
            cnpj,
            ano,
            sequencial,
        ]
    ):
        return []

    url = (
        f"{BASE_URL}/orgaos/"
        f"{cnpj}/contratos/"
        f"{ano}/"
        f"{sequencial}"
    )

    try:

        contrato = consultar_pncp(
            url,
            {},
            max_tentativas=3,
        )

        if not isinstance(
            contrato,
            dict,
        ):
            return []

        documentos = []

        # ----------------------------------------------------
        # Termos
        # ----------------------------------------------------

        termos = contrato.get(
            "termos",
            [],
        )

        if isinstance(
            termos,
            list,
        ):

            for item in termos:

                item = dict(item)

                item[
                    "tipoDocumentoNome"
                ] = (
                    item.get(
                        "tipoDocumentoNome"
                    )
                    or item.get(
                        "tipoTermoNome"
                    )
                    or "Termo de Contrato"
                )

                documentos.append(
                    item
                )

        # ----------------------------------------------------
        # Compatibilidade com estruturas antigas.
        # ----------------------------------------------------

        aditivos = contrato.get(
            "termosAditivos",
            [],
        )

        if isinstance(
            aditivos,
            list,
        ):

            for item in aditivos:

                item = dict(item)

                item[
                    "tipoDocumentoNome"
                ] = "Termo Aditivo"

                documentos.append(
                    item
                )

        apostilamentos = contrato.get(
            "termosApostilamentos",
            [],
        )

        if isinstance(
            apostilamentos,
            list,
        ):

            for item in apostilamentos:

                item = dict(item)

                item[
                    "tipoDocumentoNome"
                ] = (
                    "Termo de Apostilamento"
                )

                documentos.append(
                    item
                )

        # ----------------------------------------------------
        # Histórico, quando disponível.
        # ----------------------------------------------------

        historico = contrato.get(
            "historico",
            [],
        )

        if isinstance(
            historico,
            list,
        ):

            for item in historico:

                item = dict(item)

                item[
                    "tipoDocumentoNome"
                ] = (
                    item.get(
                        "tipoEventoNome"
                    )
                    or "Histórico"
                )

                documentos.append(
                    item
                )

        return documentos

    except Exception:
        return []


# ============================================================
# EXPORTAÇÃO - WORD
# ============================================================

def gerar_word(
    df: pd.DataFrame,
    tipo_consulta: str,
    data_inicio,
    data_fim,
) -> bytes:

    doc = Document()

    section = doc.sections[0]

    section.top_margin = Pt(40)
    section.bottom_margin = Pt(40)
    section.left_margin = Pt(45)
    section.right_margin = Pt(45)

    # --------------------------------------------------------
    # TÍTULO
    # --------------------------------------------------------

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = 1

    run = p_titulo.add_run(
        f"RELATÓRIO DE {tipo_consulta.upper()}"
    )

    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(
        0,
        51,
        102,
    )

    p_subtitulo = doc.add_paragraph()
    p_subtitulo.alignment = 1

    run = p_subtitulo.add_run(
        NOME_PREFEITURA
    )

    run.bold = True
    run.font.size = Pt(11)

    doc.add_paragraph(
        "Portal Nacional de Contratações Públicas – PNCP"
    )

    doc.add_paragraph(
        f"Período da consulta: "
        f"{data_inicio.strftime('%d/%m/%Y')} "
        f"a "
        f"{data_fim.strftime('%d/%m/%Y')}"
    )

    doc.add_paragraph(
        f"Total de registros encontrados: "
        f"{len(df)}"
    )

    doc.add_paragraph("")

    doc.add_heading(
        "Principais Registros",
        level=2,
    )

    # --------------------------------------------------------
    # REGISTROS
    # --------------------------------------------------------

    for pos, (_, row) in enumerate(
        df.head(100).iterrows(),
        start=1,
    ):

        dados = obter_dados_registro(
            row,
            tipo_consulta,
        )

        tabela = doc.add_table(
            rows=0,
            cols=2,
        )

        tabela.style = "Table Grid"

        campos = [
            ("Registro", str(pos)),
            (
                "Número",
                dados["numero"],
            ),
            (
                "Controle PNCP",
                dados["id_pncp"],
            ),
            (
                "Processo",
                dados["processo"],
            ),
            (
                "Objeto",
                dados["objeto"],
            ),
            (
                "Fornecedor",
                dados["fornecedor"],
            ),
            (
                "CNPJ/CPF Fornecedor",
                dados["cnpj_fornecedor"],
            ),
            (
                "Valor",
                dados["valor"],
            ),
            (
                "Data de Assinatura/Publicação",
                dados["data_assinatura"],
            ),
            (
                "Início da Vigência",
                dados["vigencia_inicio"],
            ),
            (
                "Fim da Vigência",
                dados["vigencia_fim"],
            ),
            (
                "Situação",
                dados["situacao"],
            ),
            (
                "Órgão",
                dados["orgao"],
            ),
            (
                "Unidade",
                dados["unidade"],
            ),
        ]

        for nome_campo, valor_campo in campos:

            cells = tabela.add_row().cells

            cells[0].text = str(
                nome_campo
            )

            cells[1].text = texto_valido(
                valor_campo
            )

            for run in (
                cells[0]
                .paragraphs[0]
                .runs
            ):
                run.bold = True
                run.font.size = Pt(9)

            for run in (
                cells[1]
                .paragraphs[0]
                .runs
            ):
                run.font.size = Pt(9)

        doc.add_paragraph("")

    footer = section.footer

    p_footer = footer.paragraphs[0]
    p_footer.alignment = 1

    run = p_footer.add_run(
        "Consulta realizada no Portal Nacional "
        "de Contratações Públicas – PNCP"
    )

    run.font.size = Pt(8)

    buffer = io.BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# PDF
# ============================================================

def texto_pdf(valor: Any) -> str:
    """
    Converte texto para latin-1 sem quebrar o PDF.
    """

    texto = texto_valido(
        valor
    )

    return (
        texto
        .encode(
            "latin-1",
            "replace",
        )
        .decode(
            "latin-1"
        )
    )


def gerar_pdf(
    df: pd.DataFrame,
    tipo_consulta: str,
    data_inicio,
    data_fim,
) -> bytes:

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15,
    )

    pdf.add_page()

    # --------------------------------------------------------
    # TÍTULO
    # --------------------------------------------------------

    pdf.set_font(
        "Arial",
        "B",
        15,
    )

    titulo = texto_pdf(
        f"RELATORIO DE {tipo_consulta.upper()}"
    )

    pdf.cell(
        0,
        10,
        txt=titulo,
        ln=True,
        align="C",
    )

    pdf.set_font(
        "Arial",
        "B",
        10,
    )

    pdf.cell(
        0,
        7,
        txt=texto_pdf(
            NOME_PREFEITURA
        ),
        ln=True,
        align="C",
    )

    pdf.ln(3)

    pdf.set_font(
        "Arial",
        size=9,
    )

    pdf.cell(
        0,
        6,
        txt=texto_pdf(
            "Portal Nacional de "
            "Contratacoes Publicas - PNCP"
        ),
        ln=True,
        align="C",
    )

    pdf.cell(
        0,
        6,
        txt=texto_pdf(
            f"Periodo: "
            f"{data_inicio.strftime('%d/%m/%Y')} "
            f"a "
            f"{data_fim.strftime('%d/%m/%Y')}"
        ),
        ln=True,
    )

    pdf.cell(
        0,
        6,
        txt=texto_pdf(
            f"Total de registros: {len(df)}"
        ),
        ln=True,
    )

    pdf.ln(5)

    # --------------------------------------------------------
    # REGISTROS
    # --------------------------------------------------------

    for pos, (_, row) in enumerate(
        df.head(100).iterrows(),
        start=1,
    ):

        dados = obter_dados_registro(
            row,
            tipo_consulta,
        )

        pdf.set_font(
            "Arial",
            "B",
            10,
        )

        pdf.cell(
            0,
            7,
            txt=texto_pdf(
                f"Registro {pos}"
            ),
            ln=True,
        )

        pdf.set_font(
            "Arial",
            size=8,
        )

        campos = [
            (
                "Numero",
                dados["numero"],
            ),
            (
                "Controle PNCP",
                dados["id_pncp"],
            ),
            (
                "Processo",
                dados["processo"],
            ),
            (
                "Objeto",
                dados["objeto"],
            ),
            (
                "Fornecedor",
                dados["fornecedor"],
            ),
            (
                "CNPJ/CPF Fornecedor",
                dados["cnpj_fornecedor"],
            ),
            (
                "Valor",
                dados["valor"],
            ),
            (
                "Data de Assinatura/Publicacao",
                dados["data_assinatura"],
            ),
            (
                "Inicio da Vigencia",
                dados["vigencia_inicio"],
            ),
            (
                "Fim da Vigencia",
                dados["vigencia_fim"],
            ),
            (
                "Situacao",
                dados["situacao"],
            ),
            (
                "Orgao",
                dados["orgao"],
            ),
            (
                "Unidade",
                dados["unidade"],
            ),
        ]

        for nome_campo, valor_campo in campos:

            texto = (
                f"{nome_campo}: "
                f"{texto_valido(valor_campo)}"
            )

            pdf.multi_cell(
                0,
                5,
                txt=texto_pdf(
                    texto
                ),
            )

        pdf.ln(4)

    # --------------------------------------------------------
    # RODAPÉ
    # --------------------------------------------------------

    pdf.set_font(
        "Arial",
        "I",
        7,
    )

    pdf.cell(
        0,
        5,
        txt=texto_pdf(
            "Consulta realizada no Portal Nacional "
            "de Contratacoes Publicas - PNCP"
        ),
        ln=True,
        align="C",
    )

    resultado = pdf.output(
        dest="S"
    )

    if isinstance(
        resultado,
        str,
    ):
        resultado = resultado.encode(
            "latin-1"
        )

    return bytes(resultado)



# ============================================================
# DOCUMENTOS / ARQUIVOS DO PNCP
# ============================================================

def extrair_ano_sequencial_controle(controle: str):
    """Tenta extrair ano e sequencial de um número de controle PNCP."""
    controle = texto_valido(controle, "").strip()
    if not controle:
        return None, None

    match = re.search(r"-(\d+)/(?P<ano>\d{4})$", controle)
    if match:
        return int(match.group("ano")), int(match.group(1))

    match = re.search(r"(\d+)/(?P<ano>\d{4})$", controle)
    if match:
        return int(match.group("ano")), int(match.group(1))

    return None, None


def consultar_documentos_pncp(
    cnpj: str,
    ano: int,
    sequencial: int,
    tipo_recurso: str = "contratacao",
    sequencial_ata: Optional[int] = None,
):
    """Consulta os documentos publicados para contratação, contrato ou ata."""
    cnpj = normalizar_cnpj(cnpj)
    if len(cnpj) != 14:
        raise ValueError("CNPJ inválido. Informe 14 dígitos.")
    if not ano or not sequencial:
        raise ValueError("Informe ano e sequencial válidos.")

    if tipo_recurso == "contrato":
        url = (
            f"{BASE_URL}/orgaos/{cnpj}/contratos/"
            f"{int(ano)}/{int(sequencial)}/arquivos"
        )
    elif tipo_recurso == "ata":
        if not sequencial_ata:
            raise ValueError("Informe o sequencial da Ata.")
        url = (
            f"{BASE_URL}/orgaos/{cnpj}/compras/"
            f"{int(ano)}/{int(sequencial)}/atas/"
            f"{int(sequencial_ata)}/arquivos"
        )
    else:
        url = (
            f"{BASE_URL}/orgaos/{cnpj}/compras/"
            f"{int(ano)}/{int(sequencial)}/arquivos"
        )

    resposta = consultar_pncp(url, {}, max_tentativas=MAX_TENTATIVAS)

    if isinstance(resposta, dict):
        documentos = resposta.get("documentos")
        if isinstance(documentos, list):
            return documentos

    return extrair_registros(resposta)


def obter_url_documento(documento: Dict[str, Any]) -> str:
    """Retorna a URL/URI de download informada pelo PNCP."""
    for campo in ("url", "uri", "URL", "URI"):
        valor = documento.get(campo)
        if valor:
            return str(valor).strip()
    return ""


def baixar_documento_pncp(url: str) -> Tuple[bytes, str]:
    """Baixa o arquivo binário publicado pelo PNCP."""
    if not url:
        raise ValueError("O PNCP não forneceu uma URL para este documento.")

    sessao = criar_sessao_http()
    response = sessao.get(
        url,
        timeout=(TIMEOUT_CONEXAO, TIMEOUT_LEITURA),
        allow_redirects=True,
        headers={**HEADERS, "Accept": "*/*"},
    )

    if response.status_code != 200:
        detalhe = mensagem_erro_http(response)
        raise RuntimeError(
            f"Não foi possível baixar o arquivo. HTTP {response.status_code}: {detalhe}"
        )

    nome = "documento_pncp"
    content_disposition = response.headers.get("Content-Disposition", "")
    match = re.search(
        r"filename\*?=(?:UTF-8'' )?[\"']?([^\"';]+)",
        content_disposition,
        flags=re.IGNORECASE,
    )
    if match:
        nome = match.group(1).strip()

    content_type = response.headers.get("Content-Type", "").lower().split(";")[0].strip()
    if "." not in nome:
        extensoes = {
            "application/pdf": ".pdf",
            "application/zip": ".zip",
            "application/msword": ".doc",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
            "application/vnd.ms-excel": ".xls",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
        }
        nome += extensoes.get(content_type, ".bin")

    return response.content, nome


def nome_arquivo_seguro(nome: str) -> str:
    nome = re.sub(r"[^A-Za-z0-9._-]+", "_", nome)
    return nome.strip("._") or "documento_pncp"


def renderizar_aba_arquivos_pncp():
    """Interface da área dedicada aos arquivos/documentos do PNCP."""
    st.title("📥 Arquivos do PNCP")
    st.markdown(
        "Localize e baixe os **documentos publicados no PNCP** "
        "associados a uma contratação, contrato ou Ata de Registro de Preços."
    )
    st.info(
        "💡 A API oficial do PNCP disponibiliza, quando cadastrados, "
        "título, tipo de documento, data de publicação e URL/URI para download."
    )

    col1, col2 = st.columns(2)
    with col1:
        tipo_arquivo = st.selectbox(
            "O que você deseja consultar?",
            [
                "📄 Edital / Aviso / Contratação",
                "📑 Contrato / Empenho",
                "📋 Ata de Registro de Preços",
            ],
            key="tipo_arquivo_pncp",
        )
    with col2:
        controle = st.text_input(
            "Número de Controle PNCP (opcional)",
            placeholder="Ex.: 44826840000183-1-000123/2026",
            help="Se informado no formato completo, ano e sequencial serão identificados automaticamente.",
            key="controle_arquivo_pncp",
        )

    ano_auto, seq_auto = extrair_ano_sequencial_controle(controle)
    col1, col2, col3 = st.columns(3)
    with col1:
        ano = st.number_input(
            "Ano",
            min_value=2000,
            max_value=datetime.date.today().year + 2,
            value=ano_auto or datetime.date.today().year,
            step=1,
            key="ano_arquivo_pncp",
        )
    with col2:
        sequencial = st.number_input(
            "Sequencial da contratação/contrato",
            min_value=1,
            value=seq_auto or 1,
            step=1,
            key="sequencial_arquivo_pncp",
        )
    with col3:
        if tipo_arquivo == "📋 Ata de Registro de Preços":
            sequencial_ata = st.number_input(
                "Sequencial da Ata",
                min_value=1,
                value=1,
                step=1,
                key="sequencial_ata_arquivo_pncp",
            )
        else:
            sequencial_ata = None
            st.markdown("**Órgão contratante**")
            st.code(CNPJ_RIO_DAS_PEDRAS, language="text")

    tipo_recurso = (
        "contrato"
        if tipo_arquivo == "📑 Contrato / Empenho"
        else "ata"
        if tipo_arquivo == "📋 Ata de Registro de Preços"
        else "contratacao"
    )

    if st.button(
        "🔎 Localizar arquivos no PNCP",
        type="primary",
        use_container_width=True,
        key="buscar_arquivos_pncp",
    ):
        try:
            with st.spinner("Consultando os documentos publicados no PNCP..."):
                documentos = consultar_documentos_pncp(
                    CNPJ_RIO_DAS_PEDRAS,
                    int(ano),
                    int(sequencial),
                    tipo_recurso=tipo_recurso,
                    sequencial_ata=int(sequencial_ata) if sequencial_ata else None,
                )
            st.session_state["documentos_pncp"] = documentos
            st.session_state["documentos_pncp_contexto"] = {
                "tipo": tipo_arquivo,
                "ano": int(ano),
                "sequencial": int(sequencial),
                "sequencial_ata": int(sequencial_ata) if sequencial_ata else None,
            }
        except Exception as erro:
            st.session_state["documentos_pncp"] = None
            st.error(f"❌ Não foi possível consultar os arquivos: {erro}")

    documentos = st.session_state.get("documentos_pncp")
    contexto = st.session_state.get("documentos_pncp_contexto", {})

    if documentos is None:
        st.caption("Informe os dados e clique em **Localizar arquivos no PNCP**.")
        return

    if not documentos:
        st.warning(
            "ℹ️ Nenhum arquivo/documento foi retornado para os dados informados. "
            "Isso pode significar que não há anexos publicados ou que o registro "
            "informado não possui documentos disponíveis na API."
        )
        return

    st.success(f"✅ {len(documentos)} documento(s) encontrado(s).")

    for i, documento in enumerate(documentos, start=1):
        if not isinstance(documento, dict):
            continue

        titulo = texto_valido(
            documento.get("titulo") or documento.get("nome"),
            f"Documento {i}",
        )
        tipo_doc = texto_valido(
            documento.get("tipoDocumentoNome") or documento.get("tipoDocumento"),
            "Documento",
        )
        data_doc = formatar_data(
            documento.get("dataPublicacaoPncp")
            or documento.get("dataPublicacao")
        )
        sequencial_doc = texto_valido(
            documento.get("sequencialDocumento"),
            str(i),
        )
        url_doc = obter_url_documento(documento)

        with st.container(border=True):
            c1, c2 = st.columns([4, 1])
            with c1:
                st.markdown(f"### 📄 {titulo}")
                st.write(f"**Tipo:** {tipo_doc}")
                st.write(f"**Data de publicação no PNCP:** {data_doc}")
                st.write(f"**Sequencial do documento:** {sequencial_doc}")

                if url_doc:
                    st.caption("Origem: documento publicado no PNCP")
                else:
                    st.warning("O PNCP retornou o documento, mas não informou uma URL/URI de download.")

            with c2:
                if url_doc:
                    try:
                        with st.spinner("Preparando arquivo..."):
                            arquivo_bytes, nome_original = baixar_documento_pncp(url_doc)
                        nome_download = nome_arquivo_seguro(
                            nome_original or f"PNCP_documento_{sequencial_doc}"
                        )
                        st.download_button(
                            "📥 Baixar arquivo",
                            data=arquivo_bytes,
                            file_name=nome_download,
                            mime="application/octet-stream",
                            use_container_width=True,
                            key=f"download_pncp_{contexto.get('ano')}_{contexto.get('sequencial')}_{sequencial_doc}_{i}",
                        )
                    except Exception as erro:
                        st.error(f"Erro no download: {erro}")
                        st.link_button(
                            "🔗 Abrir no PNCP",
                            url_doc,
                            use_container_width=True,
                        )

# ============================================================
# TELA PRINCIPAL
# ============================================================

st.title(
    "🏛️ Contratações de Rio das Pedras/SP"
)

st.markdown(
    "Consulta integrada de **Contratos, Atas de Registro "
    "de Preços e Editais/Avisos de Contratações** "
    "diretamente do Portal Nacional de Contratações "
    "Públicas – PNCP."
)

st.warning(
    "⚠️ **Atenção:** o Portal Nacional de Contratações "
    "Públicas (PNCP) pode apresentar lentidão, timeout "
    "ou indisponibilidade temporária. Consultas com "
    "períodos maiores podem levar alguns minutos. "
    "O sistema possui tentativas automáticas para erros "
    "temporários."
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header(
    "⚙️ Parâmetros da Consulta"
)

tipo_consulta = st.sidebar.selectbox(
    "Selecione o tipo:",
    [
        "Contratos",
        "Atas de Registro de Preços",
        "Editais e Avisos de Contratações",
        "📥 Arquivos do PNCP",
    ],
)

if tipo_consulta == "📥 Arquivos do PNCP":
    renderizar_aba_arquivos_pncp()
    st.stop()


# ============================================================
# MODALIDADES
# ============================================================

modalidade_codigo = None

if (
    tipo_consulta
    == "Editais e Avisos de Contratações"
):

    modalidade_opcoes = {
        "Pregão - Eletrônico (6)": 6,
        "Dispensa de Licitação (8)": 8,
        "Inexigibilidade (9)": 9,
        "Concorrência - Eletrônica (2)": 2,
    }

    mod_escolhida = st.sidebar.selectbox(
        "Modalidade:",
        list(
            modalidade_opcoes.keys()
        ),
    )

    modalidade_codigo = (
        modalidade_opcoes[
            mod_escolhida
        ]
    )


# ============================================================
# DATAS
# ============================================================

data_inicio = st.sidebar.date_input(
    "📅 Data Inicial",
    value=pd.to_datetime(
        "2026-01-01"
    ).date(),
)

data_fim = st.sidebar.date_input(
    "📅 Data Final",
    value=datetime.date.today(),
)


# ============================================================
# VALIDAÇÃO
# ============================================================

if data_fim < data_inicio:

    st.sidebar.error(
        "⚠️ A Data Final não pode ser anterior "
        "à Data Inicial."
    )

    st.stop()


quantidade_dias = (
    data_fim - data_inicio
).days


if quantidade_dias > 365:

    st.sidebar.error(
        "⚠️ O período não pode ser superior "
        "a 365 dias."
    )

    st.stop()


# ============================================================
# SESSION STATE
# ============================================================

if (
    "df_resultado"
    not in st.session_state
):
    st.session_state.df_resultado = None


if (
    "tipo_anterior"
    not in st.session_state
):
    st.session_state.tipo_anterior = (
        tipo_consulta
    )


if (
    st.session_state.tipo_anterior
    != tipo_consulta
):

    st.session_state.df_resultado = None

    st.session_state.tipo_anterior = (
        tipo_consulta
    )


if (
    "paginas_consultadas"
    not in st.session_state
):
    st.session_state.paginas_consultadas = 0


if (
    "ultima_consulta"
    not in st.session_state
):
    st.session_state.ultima_consulta = None


# ============================================================
# BOTÃO DE CONSULTA
# ============================================================

if st.sidebar.button(
    "🔎 Gerar Consulta",
    type="primary",
    use_container_width=True,
):

    # --------------------------------------------------------
    # ENDPOINT
    # --------------------------------------------------------

    endpoints = {

        "Contratos":
            f"{BASE_URL}/contratos",

        "Atas de Registro de Preços":
            f"{BASE_URL}/atas",

        "Editais e Avisos de Contratações":
            f"{BASE_URL}/contratacoes/publicacao",
    }

    endpoint = endpoints[
        tipo_consulta
    ]

    # --------------------------------------------------------
    # TAMANHO DA PÁGINA
    # --------------------------------------------------------

    if tipo_consulta == "Contratos":

        tamanho_pagina = (
            TAMANHO_PAGINA_CONTRATOS
        )

    elif (
        tipo_consulta
        == "Atas de Registro de Preços"
    ):

        tamanho_pagina = (
            TAMANHO_PAGINA_ATAS
        )

    else:

        tamanho_pagina = (
            TAMANHO_PAGINA_EDITAIS
        )

    # --------------------------------------------------------
    # PARÂMETROS BASE
    # --------------------------------------------------------

    params = {

        "dataInicial":
            data_inicio.strftime(
                "%Y%m%d"
            ),

        "dataFinal":
            data_fim.strftime(
                "%Y%m%d"
            ),

        "pagina": 1,

        "tamanhoPagina":
            tamanho_pagina,
    }

    # --------------------------------------------------------
    # EDITAIS
    # --------------------------------------------------------

    if (
        tipo_consulta
        == "Editais e Avisos de Contratações"
    ):

        params.update(
            {
                "codigoModalidadeContratacao":
                    modalidade_codigo,

                "uf":
                    UF,

                "codigoMunicipioIbge":
                    CODIGO_IBGE_RIO_DAS_PEDRAS,

                "cnpj":
                    CNPJ_RIO_DAS_PEDRAS,
            }
        )

    # --------------------------------------------------------
    # CONTRATOS
    # --------------------------------------------------------

    elif tipo_consulta == "Contratos":

        params[
            "cnpjOrgao"
        ] = CNPJ_RIO_DAS_PEDRAS

    # --------------------------------------------------------
    # ATAS
    # --------------------------------------------------------

    elif (
        tipo_consulta
        == "Atas de Registro de Preços"
    ):

        params[
            "cnpj"
        ] = CNPJ_RIO_DAS_PEDRAS

    # --------------------------------------------------------
    # EXECUÇÃO
    # --------------------------------------------------------

    st.session_state.df_resultado = None

    try:

        inicio_execucao = time.time()

        with st.spinner(
            "🔄 Consultando o PNCP. "
            "Isso pode levar alguns segundos..."
        ):

            registros, paginas = (
                executar_consulta(
                    endpoint,
                    params,
                )
            )

        # ----------------------------------------------------
        # DATAFRAME
        # ----------------------------------------------------

        df_temp = pd.DataFrame(
            registros
        )

        df_temp = tratar_dataframe(
            df_temp
        )

        # ----------------------------------------------------
        # FILTRO DE SEGURANÇA
        # ----------------------------------------------------

        if (
            tipo_consulta
            == "Contratos"
            and not df_temp.empty
        ):

            df_filtrado = (
                filtrar_por_cnpj(
                    df_temp,
                    CNPJ_RIO_DAS_PEDRAS,
                )
            )

            # Só substitui se realmente encontrou.
            if not df_filtrado.empty:
                df_temp = df_filtrado

        # ----------------------------------------------------
        # DUPLICIDADES
        # ----------------------------------------------------

        df_temp = remover_duplicidades(
            df_temp
        )

        # ----------------------------------------------------
        # SESSION
        # ----------------------------------------------------

        st.session_state.df_resultado = (
            df_temp
        )

        st.session_state.paginas_consultadas = (
            paginas
        )

        st.session_state.ultima_consulta = (
            time.time() - inicio_execucao
        )

        # ----------------------------------------------------
        # MENSAGEM
        # ----------------------------------------------------

        if df_temp.empty:

            st.info(
                "ℹ️ O PNCP não retornou registros "
                "para os parâmetros selecionados."
            )

        else:

            st.success(
                f"✅ Consulta concluída. "
                f"{len(df_temp)} registros encontrados "
                f"em {paginas} página(s)."
            )

    except Exception as erro:

        st.session_state.df_resultado = None

        mensagem = str(
            erro
        )

        if (
            "Tamanho de página inválido"
            in mensagem
        ):

            st.error(
                "❌ O PNCP rejeitou o tamanho da página. "
                "O código já utiliza limites específicos "
                "por endpoint. Verifique se o PNCP alterou "
                "as regras da API."
            )

        elif (
            "timeout"
            in mensagem.lower()
        ):

            st.error(
                "❌ O PNCP demorou demais para responder."
            )

            st.info(
                "💡 Tente novamente com um período menor. "
                "O PNCP pode estar temporariamente lento."
            )

        else:

            st.error(
                f"❌ Erro na consulta ao PNCP:\n\n"
                f"{mensagem}"
            )


# ============================================================
# RESULTADO
# ============================================================

df = (
    st.session_state.df_resultado
)


if (
    df is not None
    and not df.empty
):

    # ========================================================
    # RESUMO
    # ========================================================

    st.markdown("---")

    st.subheader(
        "📊 Resumo da Consulta"
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Total de Registros",
        len(df),
    )

    col2.metric(
        "Páginas Consultadas",
        st.session_state.paginas_consultadas,
    )

    if (
        st.session_state.ultima_consulta
        is not None
    ):

        tempo = (
            st.session_state.ultima_consulta
        )

        col3.metric(
            "Tempo da Consulta",
            f"{tempo:.1f} s",
        )

    # ========================================================
    # VALOR TOTAL
    # ========================================================

    coluna_valor = next(
        (
            coluna
            for coluna in [
                "valorGlobal",
                "valorInicial",
                "valorTotal",
                "valorTotalHomologado",
                "valorTotalEstimado",
                "valorAta",
                "valorContrato",
            ]
            if coluna in df.columns
        ),
        None,
    )

    if coluna_valor:

        valores = (
            df[coluna_valor]
            .map(valor_numerico)
        )

        total_valor = (
            valores
            .fillna(0)
            .sum()
        )

        st.metric(
            "💰 Valor Total Encontrado",
            formatar_moeda_br(
                total_valor
            ),
        )

    st.markdown("---")


    # ========================================================
    # EXPORTAÇÕES
    # ========================================================

    st.subheader(
        "📥 Exportação dos Dados"
    )

    cols = st.columns(4)

    nome_arquivo = (
        tipo_consulta
        .replace(
            " ",
            "_",
        )
        .replace(
            "/",
            "_",
        )
        .replace(
            "ç",
            "c",
        )
        .replace(
            "ã",
            "a",
        )
    )

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    buffer_xlsx = io.BytesIO()

    df.to_excel(
        buffer_xlsx,
        index=False,
    )

    cols[0].download_button(
        "📊 Excel",
        buffer_xlsx.getvalue(),
        f"{nome_arquivo}_Rio_Das_Pedras.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # CSV
    # --------------------------------------------------------

    csv_data = df.to_csv(
        index=False,
        encoding="utf-8-sig",
    )

    cols[1].download_button(
        "📄 CSV",
        csv_data,
        f"{nome_arquivo}_Rio_Das_Pedras.csv",
        mime="text/csv",
        use_container_width=True,
    )

    # --------------------------------------------------------
    # WORD
    # --------------------------------------------------------

    try:

        word_bytes = gerar_word(
            df,
            tipo_consulta,
            data_inicio,
            data_fim,
        )

        cols[2].download_button(
            "📝 Word",
            word_bytes,
            f"Relatorio_{nome_arquivo}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    except Exception as erro_word:

        cols[2].error(
            f"Erro Word: {erro_word}"
        )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    try:

        pdf_bytes = gerar_pdf(
            df,
            tipo_consulta,
            data_inicio,
            data_fim,
        )

        cols[3].download_button(
            "📕 PDF",
            pdf_bytes,
            f"Relatorio_{nome_arquivo}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    except Exception as erro_pdf:

        cols[3].error(
            f"Erro PDF: {erro_pdf}"
        )


    # ========================================================
    # CONSULTA DE ADITIVOS
    # ========================================================

    if tipo_consulta == "Contratos":

        st.markdown("---")

        st.subheader(
            "🔍 Aditivos / Termos do Contrato"
        )

        opcoes_contratos = []

        for indice, (_, row) in enumerate(
            df.iterrows()
        ):

            dados = obter_dados_registro(
                row,
                "Contratos",
            )

            identificacao = (
                f"{dados['numero']} "
                f"| Processo: "
                f"{dados['processo']} "
                f"| PNCP: "
                f"{dados['id_pncp']}"
            )

            opcoes_contratos.append(
                (
                    indice,
                    identificacao,
                )
            )

        if opcoes_contratos:

            mapa_contratos = {
                texto: indice
                for indice, texto
                in opcoes_contratos
            }

            selecionado = st.selectbox(
                "Selecione um contrato:",
                list(
                    mapa_contratos.keys()
                ),
            )

            numero_aditivo = st.text_input(
                "Número do aditivo (opcional):",
                placeholder="Ex.: 01/2026",
            )

            if st.button(
                "🔎 Buscar Termos / Aditivos",
                use_container_width=True,
            ):

                indice = (
                    mapa_contratos[
                        selecionado
                    ]
                )

                row = df.iloc[
                    indice
                ]

                with st.spinner(
                    "Consultando detalhes do contrato..."
                ):

                    documentos = (
                        consultar_detalhes_contrato(
                            row
                        )
                    )

                if documentos:

                    if numero_aditivo.strip():

                        termo_busca = (
                            numero_aditivo
                            .strip()
                            .lower()
                        )

                        documentos = [
                            documento
                            for documento
                            in documentos
                            if termo_busca
                            in str(
                                documento
                            ).lower()
                        ]

                    if documentos:

                        st.success(
                            f"Encontrados "
                            f"{len(documentos)} "
                            f"documento(s)."
                        )

                        for documento in documentos:

                            tipo_doc = texto_valido(
                                documento.get(
                                    "tipoDocumentoNome"
                                ),
                                "Documento",
                            )

                            numero_doc = texto_valido(
                                documento.get(
                                    "numero"
                                ),
                                "S/N",
                            )

                            data_doc = formatar_data(
                                documento.get(
                                    "dataPublicacao"
                                )
                                or documento.get(
                                    "dataPublicacaoPncp"
                                )
                                or documento.get(
                                    "dataAssinatura"
                                )
                            )

                            objeto_doc = texto_valido(
                                documento.get(
                                    "objeto"
                                ),
                                "",
                            )

                            st.info(
                                f"**Tipo:** {tipo_doc}\n\n"
                                f"**Número:** {numero_doc}\n\n"
                                f"**Data:** {data_doc}\n\n"
                                f"**Objeto:** {objeto_doc}"
                            )

                    else:

                        st.warning(
                            "Nenhum documento encontrado "
                            "com o número informado."
                        )

                else:

                    st.warning(
                        "Nenhum termo/aditivo foi encontrado "
                        "ou os dados detalhados deste contrato "
                        "não estão disponíveis no PNCP."
                    )


    # ========================================================
    # GRÁFICOS
    # ========================================================

    st.markdown("---")

    st.subheader(
        "📈 Análise Gráfica"
    )

    coluna_data = next(
        (
            coluna
            for coluna in [
                "dataPublicacao",
                "dataAssinatura",
                "dataInclusao",
                "dataPublicacaoPncp",
                "dataAssinaturaAta",
                "dataCelebracao",
            ]
            if coluna in df.columns
        ),
        None,
    )

    coluna_valor_grafico = next(
        (
            coluna
            for coluna in [
                "valorGlobal",
                "valorInicial",
                "valorTotal",
                "valorTotalHomologado",
                "valorTotalEstimado",
                "valorAta",
                "valorContrato",
            ]
            if coluna in df.columns
        ),
        None,
    )

    if coluna_data:

        try:

            df_grafico = df.copy()

            df_grafico[
                "data_grafico"
            ] = pd.to_datetime(
                df_grafico[
                    coluna_data
                ],
                errors="coerce",
            )

            df_grafico = (
                df_grafico
                .dropna(
                    subset=[
                        "data_grafico"
                    ]
                )
            )

            if not df_grafico.empty:

                df_grafico[
                    "mes_ano"
                ] = (
                    df_grafico[
                        "data_grafico"
                    ]
                    .dt
                    .to_period("M")
                    .astype(str)
                )

                aba1, aba2 = st.tabs(
                    [
                        "🔢 Quantidade",
                        "💰 Volume Financeiro",
                    ]
                )

                with aba1:

                    quantidade = (
                        df_grafico[
                            "mes_ano"
                        ]
                        .value_counts()
                        .sort_index()
                    )

                    st.bar_chart(
                        quantidade
                    )

                with aba2:

                    if coluna_valor_grafico:

                        df_grafico[
                            "_valor_grafico"
                        ] = (
                            df_grafico[
                                coluna_valor_grafico
                            ]
                            .map(
                                valor_numerico
                            )
                            .fillna(0)
                        )

                        financeiro = (
                            df_grafico
                            .groupby(
                                "mes_ano"
                            )[
                                "_valor_grafico"
                            ]
                            .sum()
                            .sort_index()
                        )

                        st.bar_chart(
                            financeiro
                        )

                    else:

                        st.info(
                            "ℹ️ Não há coluna financeira "
                            "compatível para gerar este gráfico."
                        )

            else:

                st.info(
                    "ℹ️ Não foram encontradas datas "
                    "válidas para gerar o gráfico."
                )

        except Exception as erro:

            st.info(
                f"ℹ️ Não foi possível gerar os gráficos: "
                f"{erro}"
            )

    else:

        st.info(
            "ℹ️ Não foi encontrada uma coluna de data "
            "compatível com a análise gráfica."
        )


    # ========================================================
    # TABELA
    # ========================================================

    st.markdown("---")

    st.subheader(
        "📋 Tabela de Dados Detalhada"
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# RESULTADO VAZIO
# ============================================================

elif (
    st.session_state.df_resultado
    is not None
    and st.session_state.df_resultado.empty
):

    st.warning(
        "⚠️ Nenhum registro encontrado para "
        f"{NOME_MUNICIPIO} no período selecionado."
    )

    st.info(
        "💡 Verifique o período, a modalidade selecionada "
        "e tente novamente. O PNCP pode eventualmente "
        "retornar 204 quando não há registros."
    )


# ============================================================
# RODAPÉ
# ============================================================

st.markdown("---")

st.caption(
    "Fonte: Portal Nacional de Contratações Públicas (PNCP) "
    "— Consulta pública da API."
)
