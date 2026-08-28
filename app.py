import io
import re
import time
import zipfile
import datetime as dt
import json
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor
from fpdf import FPDF

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    SKLEARN_OK = True
except Exception:
    SKLEARN_OK = False


# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Controle Interno — PNCP Rio das Pedras/SP",
    page_icon="🛡️",
    layout="wide",
)

CNPJ = "44826840000183"
IBGE = "3544004"
MUNICIPIO = "Rio das Pedras/SP"
PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"

BASE_CONSULTA = "https://pncp.gov.br/api/consulta/v1"
BASE_DOCUMENTOS = "https://pncp.gov.br/api/pncp/v1"

# Limites conservadores para reduzir timeout/instabilidade do PNCP.
PAGE_CONTRATOS = 100
PAGE_ATAS = 100
PAGE_EDITAIS = 50
MAX_PAGINAS_PADRAO = 15
MAX_TENTATIVAS = 3
TIMEOUT = (10, 45)
MAX_WORKERS = 6
CACHE_TTL = 900
CACHE_DIR = Path("dados")
HISTORICO_PATH = CACHE_DIR / "contratacoes_controle_interno.parquet"
META_PATH = CACHE_DIR / "controle_incremental.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151 Safari/537.36",
    "Accept": "application/json",
    "Connection": "keep-alive",
}

TIPOS = ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratações"]


# ============================================================
# UTILITÁRIOS
# ============================================================

def texto(valor: Any, padrao: str = "N/D") -> str:
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass
    if isinstance(valor, dict):
        for k in ("nome", "razaoSocial", "descricao", "valor", "numero", "municipioNome"):
            if k in valor:
                v = texto(valor[k], "")
                if v:
                    return v
        return str(valor)
    if isinstance(valor, list):
        vals = [texto(x, "") for x in valor]
        vals = [x for x in vals if x]
        return ", ".join(vals) if vals else padrao
    s = str(valor).strip()
    return padrao if s.lower() in {"", "none", "nan", "null", "n/d", "nat"} else s


def numero(valor: Any) -> Optional[float]:
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        try:
            return None if pd.isna(valor) else float(valor)
        except Exception:
            return None
    if isinstance(valor, dict):
        for k in ("valor", "value", "valorTotal", "valorGlobal", "valorInicial"):
            if k in valor:
                n = numero(valor[k])
                if n is not None:
                    return n
        return None
    s = str(valor).strip().replace("R$", "").replace(" ", "")
    try:
        return float(s)
    except Exception:
        pass
    try:
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except Exception:
        return None


def moeda(valor: Any) -> str:
    n = numero(valor)
    if n is None:
        return "N/D"
    return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def data_br(valor: Any) -> str:
    if valor is None:
        return "N/D"
    d = pd.to_datetime(valor, errors="coerce")
    return "N/D" if pd.isna(d) else d.strftime("%d/%m/%Y")


def cnpj_limpo(valor: Any) -> str:
    return re.sub(r"\D", "", texto(valor, ""))


def slug(valor: Any) -> str:
    s = re.sub(r"[^A-Za-z0-9_-]+", "_", texto(valor, "registro"))
    return s.strip("_")[:100] or "registro"


def recursivo(valor: Any, nivel: int = 0) -> Any:
    if nivel > 4 or valor is None:
        return valor
    if isinstance(valor, dict):
        for k in ("nome", "razaoSocial", "descricao", "descricaoObjeto", "valor", "numero", "cnpj", "ni", "numeroDocumento"):
            if k in valor:
                r = recursivo(valor[k], nivel + 1)
                if r not in (None, "", "N/D"):
                    return r
        return valor
    if isinstance(valor, list):
        r = [str(recursivo(x, nivel + 1)) for x in valor]
        return ", ".join(x for x in r if x) or None
    return valor


def primeiro(row: Any, campos: List[str], padrao: Any = "N/D") -> Any:
    for campo in campos:
        try:
            valor = row.get(campo) if isinstance(row, (dict, pd.Series)) else None
        except Exception:
            valor = None
        if valor is None:
            continue
        try:
            if pd.isna(valor):
                continue
        except Exception:
            pass
        valor = recursivo(valor)
        if texto(valor, ""):
            return valor
    return padrao


# ============================================================
# HTTP / PNCP
# ============================================================

@st.cache_resource
def sessao_http():
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


def detalhe_http(r: requests.Response) -> str:
    try:
        j = r.json()
        if isinstance(j, dict):
            return texto(j.get("message") or j.get("error") or j.get("detail"), "")
    except Exception:
        pass
    return r.text[:500].strip()


def get_json(url: str, params: Optional[Dict[str, Any]] = None) -> Any:
    ultimo = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            r = sessao_http().get(url, params=params, timeout=TIMEOUT)
            if r.status_code == 200:
                try:
                    return r.json()
                except ValueError as e:
                    raise RuntimeError("O PNCP respondeu 200, mas o conteúdo não é JSON válido.") from e
            if r.status_code == 204:
                return []
            if r.status_code in (400, 422):
                raise RuntimeError(f"PNCP rejeitou os parâmetros (HTTP {r.status_code}). {detalhe_http(r)}")
            if r.status_code == 404:
                raise RuntimeError(f"Recurso não encontrado no PNCP (HTTP 404). Endpoint: {r.url}")
            if r.status_code in (408, 429, 500, 502, 503, 504):
                ultimo = RuntimeError(f"HTTP {r.status_code}: {detalhe_http(r)}")
                if tentativa < MAX_TENTATIVAS:
                    time.sleep(min(2 ** tentativa, 12))
                    continue
                raise ultimo
            raise RuntimeError(f"PNCP retornou HTTP {r.status_code}: {detalhe_http(r)}")
        except (requests.Timeout, requests.ConnectionError) as e:
            ultimo = e
            if tentativa < MAX_TENTATIVAS:
                time.sleep(min(2 ** tentativa, 12))
                continue
            raise RuntimeError("O PNCP não respondeu após várias tentativas. O portal pode estar instável ou sobrecarregado.") from e
    raise RuntimeError("Falha inesperada na comunicação com o PNCP.") from ultimo


def get_json_com_sessao(sessao: requests.Session, url: str, params: Optional[Dict[str, Any]] = None) -> Any:
    """Versão da função get_json que usa uma sessão específica (para uso em threads)."""
    ultimo = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            r = sessao.get(url, params=params, timeout=TIMEOUT)
            if r.status_code == 200:
                try:
                    return r.json()
                except ValueError as e:
                    raise RuntimeError("O PNCP respondeu 200, mas o conteúdo não é JSON válido.") from e
            if r.status_code == 204:
                return []
            if r.status_code in (400, 422):
                raise RuntimeError(f"PNCP rejeitou os parâmetros (HTTP {r.status_code}). {detalhe_http(r)}")
            if r.status_code == 404:
                raise RuntimeError(f"Recurso não encontrado no PNCP (HTTP 404). Endpoint: {r.url}")
            if r.status_code in (408, 429, 500, 502, 503, 504):
                ultimo = RuntimeError(f"HTTP {r.status_code}: {detalhe_http(r)}")
                if tentativa < MAX_TENTATIVAS:
                    time.sleep(min(2 ** tentativa, 12))
                    continue
                raise ultimo
            raise RuntimeError(f"PNCP retornou HTTP {r.status_code}: {detalhe_http(r)}")
        except (requests.Timeout, requests.ConnectionError) as e:
            ultimo = e
            if tentativa < MAX_TENTATIVAS:
                time.sleep(min(2 ** tentativa, 12))
                continue
            raise RuntimeError("O PNCP não respondeu após várias tentativas. O portal pode estar instável ou sobrecarregado.") from e
    raise RuntimeError("Falha inesperada na comunicação com o PNCP.") from ultimo


def registros_api(data: Any) -> List[Dict[str, Any]]:
    if isinstance(data, list):
        return data
    if not isinstance(data, dict):
        return []
    for k in ("data", "items", "content", "dados", "registros"):
        if isinstance(data.get(k), list):
            return data[k]
    return []


def paginacao(data: Any) -> Dict[str, Optional[int]]:
    if not isinstance(data, dict):
        return {"totalPaginas": None, "totalRegistros": None}
    out = {}
    for k in ("totalPaginas", "totalRegistros", "numeroPagina"):
        try:
            out[k] = int(data[k]) if data.get(k) is not None else None
        except Exception:
            out[k] = None
    return out


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def consultar_cache(url: str, params_tuple: Tuple[Tuple[str, str], ...], max_paginas: int) -> Tuple[List[Dict[str, Any]], int, Optional[int]]:
    base = dict(params_tuple)
    base["pagina"] = 1
    primeira = get_json(url, base)
    regs1 = registros_api(primeira)
    if not regs1:
        return [], 0, paginacao(primeira).get("totalPaginas")
    info = paginacao(primeira)
    total = info.get("totalPaginas")
    if total is None and info.get("totalRegistros") is not None:
        tamanho = max(1, int(base.get("tamanhoPagina", 50)))
        total = (info["totalRegistros"] + tamanho - 1) // tamanho
    limite = max(1, min(max_paginas, total or max_paginas))
    todos = list(regs1)
    if limite == 1:
        return todos, 1, total

    def buscar(pagina: int):
        p = dict(base)
        p["pagina"] = pagina
        # Cada worker usa sua própria Session para evitar concorrência sobre a mesma Session.
        r = requests.Session()
        r.headers.update(HEADERS)
        data = get_json_com_sessao(r, url, p)
        return pagina, registros_api(data)

    resultados = {}
    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, limite - 1)) as pool:
        futuros = {pool.submit(buscar, pagina): pagina for pagina in range(2, limite + 1)}
        for f in as_completed(futuros):
            pagina = futuros[f]
            try:
                resultados[pagina] = f.result()[1]
            except Exception:
                resultados[pagina] = []
    for pagina in range(2, limite + 1):
        todos.extend(resultados.get(pagina, []))
    return todos, limite, total


def consultar(url: str, params: Dict[str, Any], max_paginas: int) -> Tuple[List[Dict[str, Any]], int, Optional[int]]:
    serial = tuple(sorted((str(k), str(v)) for k, v in params.items()))
    return consultar_cache(url, serial, max_paginas)


# ============================================================
# BASE HISTÓRICA / CONSULTA INCREMENTAL
# ============================================================

def garantir_cache():
    CACHE_DIR.mkdir(parents=True, exist_ok=True)


def ler_meta():
    garantir_cache()
    if not META_PATH.exists():
        return {}
    try:
        return json.loads(META_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_meta(meta):
    garantir_cache()
    META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def carregar_historico(tipo: str) -> pd.DataFrame:
    garantir_cache()
    if not HISTORICO_PATH.exists():
        return pd.DataFrame()
    try:
        df = pd.read_parquet(HISTORICO_PATH)
        if "__tipo_controle" in df.columns:
            df = df[df["__tipo_controle"] == tipo].drop(columns=["__tipo_controle"])
        return df.reset_index(drop=True)
    except Exception:
        return pd.DataFrame()


def chave_historico(row):
    for c in ("numeroControlePNCP", "numeroControlePNCPAta", "numeroControlePNCPCompra", "idContratoPNCP"):
        if c in row.index:
            v = texto(row[c], "")
            if v:
                return f"{c}:{v}"
    return "|".join(texto(row.get(c), "") for c in ("numero", "numeroCompra", "numeroContrato", "processo"))


def mesclar_historico(df_antigo: pd.DataFrame, df_novo: pd.DataFrame, tipo: str) -> pd.DataFrame:
    partes = [x for x in (df_antigo, df_novo) if x is not None and not x.empty]
    if not partes:
        return pd.DataFrame()
    out = pd.concat(partes, ignore_index=True, sort=False)
    out["__chave"] = out.apply(chave_historico, axis=1)
    out = out.drop_duplicates("__chave", keep="last").drop(columns=["__chave"])
    out["__tipo_controle"] = tipo
    garantir_cache()
    try:
        out.to_parquet(HISTORICO_PATH, index=False)
    except Exception:
        # Parquet é preferencial; CSV é fallback caso pyarrow/fastparquet não esteja disponível.
        out.to_csv(CACHE_DIR / "contratacoes_controle_interno.csv", index=False, encoding="utf-8-sig")
    return out.drop(columns=["__tipo_controle"], errors="ignore").reset_index(drop=True)


# ============================================================
# DADOS / NORMALIZAÇÃO
# ============================================================

def tratar_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    for col in out.columns:
        try:
            out[col] = out[col].map(lambda v: texto(recursivo(v)) if isinstance(v, (dict, list)) else v)
        except Exception:
            pass
    return out


def deduplicar(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    chaves = ["numeroControlePNCP", "numeroControlePNCPAta", "numeroControlePNCPCompra", "idContratoPNCP"]
    for c in chaves:
        if c in df.columns:
            s = df[c].astype(str).str.strip()
            valid = s.notna() & ~s.isin({"", "nan", "None", "N/D"})
            return df.loc[~valid | ~s.duplicated(keep="first")].reset_index(drop=True)
    return df.drop_duplicates().reset_index(drop=True)


def filtrar_cnpj(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    alvo = cnpj_limpo(CNPJ)
    for col in ("cnpjOrgao", "cnpj", "cnpjCompra", "cnpjOrgaoEntidade"):
        if col in df.columns:
            s = df[col].map(cnpj_limpo)
            m = s == alvo
            if m.any():
                return df.loc[m].reset_index(drop=True)
    return df


def dados_registro(row: Any, tipo: str) -> Dict[str, Any]:
    controle = primeiro(row, ["numeroControlePNCP", "numeroControlePNCPAta", "numeroControlePNCPCompra", "idContratoPNCP"])
    if tipo == "Contratos":
        campos = {
            "numero": ["numeroContratoEmpenho", "numeroContrato", "numero"],
            "processo": ["processo", "numeroProcesso"],
            "objeto": ["objetoContrato", "objetoCompra", "objeto"],
            "fornecedor": ["nomeRazaoSocialFornecedor", "razaoSocialFornecedor", "nomeFornecedor"],
            "cnpj_fornecedor": ["niFornecedor", "cnpjFornecedor"],
            "valor": ["valorGlobal", "valorInicial", "valorTotal", "valorContrato"],
            "data": ["dataAssinatura", "dataCelebracao", "dataPublicacaoPncp"],
            "situacao": ["situacao", "status"],
        }
    elif tipo == "Atas de Registro de Preços":
        campos = {
            "numero": ["numeroAtaRegistroPreco", "numeroAta", "numero"],
            "processo": ["processo", "numeroProcesso", "processoAdministrativo"],
            "objeto": ["objetoCompra", "objeto", "descricaoObjeto"],
            "fornecedor": ["nomeRazaoSocialFornecedor", "razaoSocialFornecedor", "nomeFornecedor"],
            "cnpj_fornecedor": ["niFornecedor", "cnpjFornecedor", "ni"],
            "valor": ["valorTotal", "valorGlobal", "valorAta"],
            "data": ["dataAssinatura", "dataPublicacaoPncp", "dataCelebracao"],
            "situacao": ["situacao", "status"],
        }
    else:
        campos = {
            "numero": ["numeroCompra", "numeroEdital", "numero"],
            "processo": ["processo", "numeroProcesso"],
            "objeto": ["objetoCompra", "objeto", "descricaoObjeto"],
            "fornecedor": ["nomeRazaoSocialFornecedor", "razaoSocialFornecedor", "nomeFornecedor"],
            "cnpj_fornecedor": ["niFornecedor", "cnpjFornecedor"],
            "valor": ["valorTotalHomologado", "valorTotalEstimado", "valorTotal"],
            "data": ["dataPublicacao", "dataPublicacaoPncp", "dataInclusao"],
            "situacao": ["situacaoCompra", "situacao", "status"],
        }
    return {
        "controle": texto(controle),
        "numero": texto(primeiro(row, campos["numero"])),
        "processo": texto(primeiro(row, campos["processo"])),
        "objeto": texto(primeiro(row, campos["objeto"])),
        "fornecedor": texto(primeiro(row, campos["fornecedor"])),
        "cnpj_fornecedor": texto(primeiro(row, campos["cnpj_fornecedor"])),
        "valor_num": numero(primeiro(row, campos["valor"], None)),
        "valor": moeda(primeiro(row, campos["valor"], None)),
        "data": data_br(primeiro(row, campos["data"], None)),
        "situacao": texto(primeiro(row, campos["situacao"])),
        "modalidade": texto(primeiro(row, ["modalidadeNome", "modalidadeContratacaoNome", "modalidade"], "N/D")),
    }


# ============================================================
# RISCO — REGRAS TRANSPARENTES + ANOMALIA RELATIVA
# ============================================================

def construir_contexto_risco(df: pd.DataFrame, tipo: str) -> pd.DataFrame:
    rows = []
    for _, row in df.iterrows():
        d = dados_registro(row, tipo)
        rows.append(d)
    return pd.DataFrame(rows)


def limites_iqr(serie: pd.Series) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    s = pd.to_numeric(serie, errors="coerce").dropna()
    if len(s) < 5:
        return None, None, None
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    if iqr <= 0:
        return None, None, float(s.median())
    return float(q1 - 1.5 * iqr), float(q3 + 1.5 * iqr), float(s.median())


def calcular_risco(d: Dict[str, Any], contexto: pd.DataFrame, tipo: str) -> Dict[str, Any]:
    pontos = 0
    motivos = []
    testes = []
    valor = d.get("valor_num")
    modalidade = d.get("modalidade", "").lower()
    objeto = d.get("objeto", "").lower()

    # 1) Completude — alerta de controle, não juízo de irregularidade.
    faltantes = []
    for campo, label in (("objeto", "objeto"), ("controle", "controle PNCP"), ("data", "data")):
        if texto(d.get(campo), "") in {"", "N/D"}:
            faltantes.append(label)
    if faltantes:
        pontos += min(12, 4 * len(faltantes))
        motivos.append("Campos relevantes não identificados: " + ", ".join(faltantes))
        testes.append("Completude cadastral: atenção")
    else:
        testes.append("Completude cadastral: OK")

    # 2) Modalidades/situações que merecem análise específica.
    if "dispensa" in modalidade or "inexig" in modalidade:
        pontos += 20
        motivos.append("Modalidade que requer análise específica de fundamento e justificativas")
        testes.append("Modalidade excepcional: revisar fundamento/documentação")
    elif "emerg" in objeto or "emerg" in modalidade:
        pontos += 25
        motivos.append("Indício textual de situação emergencial")
        testes.append("Emergência: revisar motivação, prazo e documentação")
    else:
        testes.append("Modalidade: sem alerta automático desta regra")

    # 3) Valor — somente como fator de priorização.
    if valor is not None:
        if valor >= 500_000:
            pontos += 10
            motivos.append("Valor elevado: priorização por materialidade")
        elif valor >= 150_000:
            pontos += 5
            motivos.append("Valor relevante: priorização por materialidade")
        testes.append("Materialidade: avaliada")
    else:
        testes.append("Materialidade: valor não identificado")

    # 4) Outlier relativo ao conjunto consultado.
    baixo, alto, mediana = limites_iqr(contexto.get("valor_num", pd.Series(dtype=float)))
    outlier = False
    if valor is not None and alto is not None and valor > alto:
        outlier = True
        pontos += 25
        motivos.append("Valor acima do limite superior do IQR da amostra consultada")
        testes.append("Anomalia de valor: ALERTA")
    else:
        testes.append("Anomalia de valor: sem alerta pelo IQR")

    # 5) Texto que sugere continuidade/aditivo/prorrogação.
    if any(k in objeto for k in ("prorroga", "aditivo", "continuado", "continuidade")):
        pontos += 10
        motivos.append("Objeto contém termos associados a continuidade/aditamento")
        testes.append("Continuidade/aditivo: revisar histórico contratual")

    # 6) Fornecedor — concentração na própria amostra.
    fornecedor = texto(d.get("fornecedor"), "N/D")
    if fornecedor != "N/D" and not contexto.empty and "fornecedor" in contexto.columns:
        qtd = int((contexto["fornecedor"].astype(str).str.strip() == fornecedor.strip()).sum())
        if qtd >= max(3, len(contexto) * 0.10):
            pontos += 10
            motivos.append(f"Fornecedor aparece {qtd} vezes na amostra consultada")
            testes.append("Concentração de fornecedor: ALERTA para análise")
        else:
            testes.append("Concentração de fornecedor: sem alerta pela regra")

    pontos = min(100, int(pontos))
    if pontos >= 60:
        nivel = "🔴 ALTO"
    elif pontos >= 30:
        nivel = "🟡 MÉDIO"
    else:
        nivel = "🟢 BAIXO"
    if not motivos:
        motivos = ["Nenhum alerta automático identificado pelas regras aplicadas."]

    return {
        "pontos": pontos,
        "nivel": nivel,
        "motivos": motivos,
        "testes": testes,
        "outlier": outlier,
        "mediana_amostra": mediana,
        "limite_iqr": alto,
    }


# ============================================================
# MACHINE LEARNING — SIMILARIDADE DE OBJETOS
# ============================================================

@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def calcular_modelo_tfidf(textos: Tuple[str, ...]):
    if not SKLEARN_OK or len(textos) < 2:
        return None
    vec = TfidfVectorizer(lowercase=True, strip_accents="unicode", ngram_range=(1, 2), min_df=1, max_features=8000, sublinear_tf=True)
    return vec.fit_transform([t if t.strip() else "sem objeto" for t in textos])


def similares(df: pd.DataFrame, idx: int, tipo: str, limite: int = 5) -> pd.DataFrame:
    if not SKLEARN_OK or len(df) < 2:
        return pd.DataFrame()
    objetos = [dados_registro(row, tipo)["objeto"] for _, row in df.iterrows()]
    matriz = calcular_modelo_tfidf(tuple(objetos))
    if matriz is None:
        return pd.DataFrame()
    # TF-IDF normalizado: produto entre a linha selecionada e as demais = cosseno.
    scores = (matriz @ matriz[idx].T).toarray().ravel()
    ordem = scores.argsort()[::-1]
    saida = []
    for j in ordem:
        if j == idx:
            continue
        d = dados_registro(df.iloc[int(j)], tipo)
        saida.append({"Similaridade": f"{float(scores[j])*100:.1f}%", "Número": d["numero"], "Objeto": d["objeto"], "Fornecedor": d["fornecedor"], "Valor": d["valor"]})
        if len(saida) >= limite:
            break
    return pd.DataFrame(saida)


# ============================================================
# DOCUMENTOS PNCP
# ============================================================

def extrair_ano_seq_controle(controle: str) -> Tuple[Optional[int], Optional[int]]:
    s = texto(controle, "")
    m = re.search(r"-(\d+)-(\d{4})$", s)
    if not m:
        return None, None
    return int(m.group(2)), int(m.group(1))


def identificador_compra(row: Any) -> Tuple[Optional[str], Optional[int], Optional[int]]:
    cnpj = cnpj_limpo(primeiro(row, ["cnpjOrgao", "cnpj", "cnpjCompra"], CNPJ))
    ano = primeiro(row, ["anoCompra", "ano", "anoContratacao"], None)
    seq = primeiro(row, ["sequencialCompra", "sequencialContratacao", "sequencial"], None)
    try: 
        ano = int(ano)
    except Exception: 
        ano = None
    try: 
        seq = int(seq)
    except Exception: 
        seq = None
    if ano is None or seq is None:
        a, s = extrair_ano_seq_controle(primeiro(row, ["numeroControlePNCP", "numeroControlePNCPCompra"], ""))
        ano, seq = ano or a, seq or s
    return (cnpj if len(cnpj) == 14 else CNPJ), ano, seq


def identificador_contrato(row: Any) -> Tuple[Optional[str], Optional[int], Optional[int]]:
    cnpj = cnpj_limpo(primeiro(row, ["cnpjOrgao", "cnpj"], CNPJ))
    ano = primeiro(row, ["anoContrato", "anoContratoEmpenho", "ano"], None)
    seq = primeiro(row, ["sequencialContrato", "sequencialContratoEmpenho", "sequencial"], None)
    try: 
        ano = int(ano)
    except Exception: 
        ano = None
    try: 
        seq = int(seq)
    except Exception: 
        seq = None
    return (cnpj if len(cnpj) == 14 else CNPJ), ano, seq


def consultar_documentos(url: str) -> List[Dict[str, Any]]:
    data = get_json(url)
    return registros_api(data)


def listar_documentos(row: Any, tipo: str) -> List[Dict[str, Any]]:
    if tipo == "Contratos":
        c, a, s = identificador_contrato(row)
        if a is None or s is None:
            raise RuntimeError("Não foi possível identificar ano e sequencial do contrato no registro retornado pelo PNCP.")
        url = f"{BASE_DOCUMENTOS}/orgaos/{c}/contratos/{a}/{s}/arquivos"
        return consultar_documentos(url)
    c, a, s = identificador_compra(row)
    if a is None or s is None:
        raise RuntimeError("Não foi possível identificar ano e sequencial da contratação no registro retornado pelo PNCP.")
    url = f"{BASE_DOCUMENTOS}/orgaos/{c}/compras/{a}/{s}/arquivos"
    return consultar_documentos(url)


def url_documento(row: Any, tipo: str, doc: Dict[str, Any]) -> Optional[str]:
    for k in ("url", "uri", "link"):
        if texto(doc.get(k), "") not in {"", "N/D"}:
            return texto(doc[k])
    seq_doc = primeiro(doc, ["sequencialDocumento", "sequencial"], None)
    if seq_doc is None:
        return None
    if tipo == "Contratos":
        c, a, s = identificador_contrato(row)
        return f"{BASE_DOCUMENTOS}/orgaos/{c}/contratos/{a}/{s}/arquivos/{seq_doc}"
    c, a, s = identificador_compra(row)
    return f"{BASE_DOCUMENTOS}/orgaos/{c}/compras/{a}/{s}/arquivos/{seq_doc}"


def baixar_bytes(url: str) -> bytes:
    r = sessao_http().get(url, timeout=TIMEOUT, allow_redirects=True)
    if r.status_code != 200:
        raise RuntimeError(f"Download HTTP {r.status_code}: {detalhe_http(r)}")
    return r.content


def nome_documento(doc: Dict[str, Any], pos: int) -> str:
    nome = texto(doc.get("titulo") or doc.get("nomeArquivo") or doc.get("nome") or doc.get("tipoDocumentoNome"), f"documento_{pos}")
    nome = re.sub(r"[^\w\-. ]+", "_", nome, flags=re.UNICODE).strip() or f"documento_{pos}"
    return nome[:120]


# ============================================================
# EXPORTAÇÕES
# ============================================================

def gerar_excel(df: pd.DataFrame, tipo: str, inicio, fim, df_risco: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        resumo = pd.DataFrame({
            "Informação": ["Órgão", "Consulta", "Período", "Registros", "Fonte"],
            "Valor": [PREFEITURA, tipo, f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y}", len(df), "PNCP"]
        })
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        df_risco.to_excel(writer, sheet_name="Matriz de Risco", index=False)
        df.to_excel(writer, sheet_name="Dados PNCP", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                letra = col[0].column_letter
                ws.column_dimensions[letra].width = min(55, max(12, max(len(str(x.value or "")) for x in col) + 2))
    buf.seek(0)
    return buf.getvalue()


def gerar_word(row: Any, tipo: str, risco: Dict[str, Any]) -> bytes:
    d = dados_registro(row, tipo)
    doc = Document()
    sec = doc.sections[0
