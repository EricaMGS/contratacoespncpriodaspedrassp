"""
Painel de Inteligência e Apoio ao Controle Interno - PNCP
Desenvolvido com Streamlit, Pandas e Scikit-Learn.
Arquitetura otimizada para resiliência de rede e I/O.
"""

import io
import re
import os
import time
import zipfile
import datetime as dt
import json
import logging
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import streamlit as st

# Configuração de log
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Verificação de dependência opcional (Machine Learning)
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    SKLEARN_OK = True
except ImportError:
    SKLEARN_OK = False
    logging.warning("scikit-learn não encontrado. Funcionalidade de similaridade desativada.")

# ============================================================
# CONFIGURAÇÕES GLOBAIS (Preparadas para Env Vars)
# ============================================================

CNPJ = "44826840000183"
IBGE = "3544004"
MUNICIPIO = "Rio das Pedras/SP"
PREFEITURA = "Prefeitura Municipal de Rio das Pedras/SP"

BASE_CONSULTA = "https://pncp.gov.br/api/consulta/v1"
BASE_DOCUMENTOS = "https://pncp.gov.br/api/pncp/v1"

# Ajuste dinâmico de workers baseado no ambiente
MAX_WORKERS = int(os.getenv("PNCP_MAX_WORKERS", 6))
TIMEOUT = (10, int(os.getenv("PNCP_TIMEOUT_READ", 45)))
MAX_TENTATIVAS = 4

PAGE_CONTRATOS = 100
PAGE_ATAS = 100
PAGE_EDITAIS = 50
MAX_PAGINAS_PADRAO = 15

CACHE_TTL = 900
CACHE_DIR = Path("dados")
HISTORICO_PATH = CACHE_DIR / "contratacoes_controle_interno.parquet"
META_PATH = CACHE_DIR / "controle_incremental.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) MonitoramentoControleInterno/1.0",
    "Accept": "application/json",
    "Connection": "keep-alive",
}

TIPOS = ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratações"]

# ============================================================
# INTEGRAÇÃO HTTP ROBUSTA (Engenharia de Ingestão)
# ============================================================

@st.cache_resource
def sessao_http() -> requests.Session:
    """
    Cria uma sessão HTTP persistente com política de Exponential Backoff nativa.
    Isso é vital para Data Engineering ao raspar APIs instáveis.
    """
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # Configura tentativas automáticas para erros comuns de sobrecarga (429) e falha de servidor (5xx)
    retries = Retry(
        total=MAX_TENTATIVAS,
        backoff_factor=1.5, # 1.5s, 3s, 6s...
        status_forcelist=[408, 429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"]
    )
    
    adapter = HTTPAdapter(max_retries=retries, pool_connections=MAX_WORKERS, pool_maxsize=MAX_WORKERS)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    
    return session

def detalhe_http(r: requests.Response) -> str:
    try:
        j = r.json()
        if isinstance(j, dict):
            return str(j.get("message") or j.get("error") or j.get("detail") or r.text[:100])
    except ValueError:
        pass
    return r.text[:250].strip()

def get_json(url: str, params: Optional[Dict[str, Any]] = None) -> Any:
    """Faz a requisição HTTP com tratamento de erro simplificado, delegando retry ao urllib3."""
    session = sessao_http()
    try:
        r = session.get(url, params=params, timeout=TIMEOUT)
        
        if r.status_code == 200:
            return r.json()
        if r.status_code == 204:
            return []
            
        r.raise_for_status() # Lança erro se não for 200/204, capturado abaixo
        
    except requests.exceptions.RetryError as e:
        raise RuntimeError(f"Falha ao conectar no PNCP após múltiplas tentativas: Limite de taxa ou API indisponível.") from e
    except requests.exceptions.HTTPError as e:
        raise RuntimeError(f"PNCP rejeitou a requisição (HTTP {e.response.status_code}): {detalhe_http(e.response)}") from e
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Erro de conexão ao acessar o PNCP: {e}") from e

# ============================================================
# UTILITÁRIOS DE LIMPEZA E FORMATAÇÃO
# ============================================================

def texto(valor: Any, padrao: str = "N/D") -> str:
    if pd.isna(valor) or valor is None:
        return padrao
    s = str(valor).strip()
    return padrao if s.lower() in {"", "none", "nan", "null", "n/d", "nat"} else s

def numero(valor: Any) -> Optional[float]:
    if pd.isna(valor) or valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace(" ", "")
    try:
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except ValueError:
        return None

def moeda(valor: Any) -> str:
    n = numero(valor)
    if n is None:
        return "N/D"
    return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def data_br(valor: Any) -> str:
    if pd.isna(valor) or valor is None:
        return "N/D"
    d = pd.to_datetime(valor, errors="coerce")
    return "N/D" if pd.isna(d) else d.strftime("%d/%m/%Y")

def cnpj_limpo(valor: Any) -> str:
    return re.sub(r"\D", "", texto(valor, ""))

def slug(valor: Any) -> str:
    s = re.sub(r"[^A-Za-z0-9_-]+", "_", texto(valor, "registro"))
    return s.strip("_")[:100] or "registro"

def primeiro(row_dict: Dict[str, Any], campos: List[str], padrao: Any = "N/D") -> Any:
    for campo in campos:
        valor = row_dict.get(campo)
        if pd.isna(valor) or valor is None:
            continue
        if isinstance(valor, str):
            v_limpo = valor.strip()
            if v_limpo.lower() in {"", "nan", "none", "null", "n/d", "nat"}:
                continue
            return v_limpo
        return valor
    return padrao

def fuzzy_match(r_dict: Dict[str, Any], palavras: List[str]) -> Any:
    for k, v in r_dict.items():
        if pd.isna(v) or v is None:
            continue
        v_str = str(v).strip()
        if v_str.lower() in ("", "nan", "none", "n/d"):
            continue
        k_lower = str(k).lower()
        if any(p in k_lower for p in palavras):
            return v
    return None

def dados_registro(row: Any, tipo: str) -> Dict[str, Any]:
    """Padroniza as chaves do dicionário de acordo com o tipo de contratação."""
    r_dict = row.to_dict() if hasattr(row, 'to_dict') else dict(row)
    
    controle = primeiro(r_dict, ["numeroControlePNCP", "numeroControlePNCPAta", "numeroControlePNCPCompra", "idContratoPNCP"])
    
    if tipo == "Contratos":
        num = primeiro(r_dict, ["numeroContratoEmpenho", "numeroContrato", "numero"])
        proc = primeiro(r_dict, ["processo", "numeroProcesso", "compra.processo"])
        obj = primeiro(r_dict, ["objetoContrato", "objetoCompra", "compra.objetoCompra", "objeto", "descricaoObjeto"])
        forn = primeiro(r_dict, ["nomeRazaoSocialFornecedor", "fornecedor.nomeRazaoSocial", "razaoSocialFornecedor", "nomeFornecedor"])
        cnpj = primeiro(r_dict, ["niFornecedor", "fornecedor.niFornecedor", "cnpjFornecedor", "cnpj"])
        val = primeiro(r_dict, ["valorGlobal", "valorInicial", "valorTotal", "valorContrato", "compra.valorTotalHomologado"])
        dt_ = primeiro(r_dict, ["dataAssinatura", "dataCelebracao", "dataPublicacaoPncp"])
        sit = primeiro(r_dict, ["situacao", "status"])
        
    elif tipo == "Atas de Registro de Preços":
        num = primeiro(r_dict, ["numeroAtaRegistroPreco", "numeroAta", "numero"])
        proc = primeiro(r_dict, ["processo", "numeroProcesso", "processoAdministrativo", "compra.processo"])
        obj = primeiro(r_dict, ["objetoAta", "objetoAtaRegistroPreco", "objetoCompra", "compra.objetoCompra", "objeto", "descricaoObjeto"])
        forn = primeiro(r_dict, ["nomeRazaoSocialFornecedor", "fornecedor.nomeRazaoSocial", "razaoSocialFornecedor", "nomeFornecedor"])
        cnpj = primeiro(r_dict, ["niFornecedor", "fornecedor.niFornecedor", "cnpjFornecedor", "ni"])
        val = primeiro(r_dict, ["valorTotalAta", "valorTotal", "valorGlobal", "valorAta", "compra.valorTotalHomologado"])
        dt_ = primeiro(r_dict, ["dataAssinatura", "dataPublicacaoPncp", "dataCelebracao"])
        sit = primeiro(r_dict, ["situacao", "status"])
        
    else: 
        num = primeiro(r_dict, ["numeroCompra", "compra.numeroCompra", "numeroEdital", "numero"])
        proc = primeiro(r_dict, ["processo", "compra.processo", "numeroProcesso"])
        obj = primeiro(r_dict, ["objetoCompra", "compra.objetoCompra", "objeto", "descricaoObjeto"])
        forn = primeiro(r_dict, ["nomeRazaoSocialFornecedor", "fornecedor.nomeRazaoSocial", "razaoSocialFornecedor", "nomeFornecedor"])
        cnpj = primeiro(r_dict, ["niFornecedor", "fornecedor.niFornecedor", "cnpjFornecedor", "cnpjOrgao"])
        val = primeiro(r_dict, ["valorTotalHomologado", "compra.valorTotalHomologado", "valorTotalEstimado", "compra.valorTotalEstimado", "valorTotal"])
        dt_ = primeiro(r_dict, ["dataPublicacao", "dataPublicacaoPncp", "dataInclusao"])
        sit = primeiro(r_dict, ["situacaoCompra", "compra.situacaoCompraNome", "situacao", "status"])
        
    mod = primeiro(r_dict, ["modalidadeNome", "modalidadeContratacaoNome", "compra.modalidadeNome", "modalidade"])

    if texto(obj, "") == "": obj = fuzzy_match(r_dict, ["objeto", "descricao"])
    if texto(forn, "") == "": forn = fuzzy_match(r_dict, ["razaosocial", "fornecedornome", "fornecedor.nome"])
    if texto(val, "") == "": val = fuzzy_match(r_dict, ["valortotal", "valorglobal", "valorata", "valor"])
    if texto(mod, "") == "": mod = fuzzy_match(r_dict, ["modalidade", "tipoata"])
    if texto(cnpj, "") == "": cnpj = fuzzy_match(r_dict, ["cnpj", "ni"])
    
    return {
        "controle": texto(controle),
        "numero": texto(num),
        "processo": texto(proc),
        "objeto": texto(obj),
        "fornecedor": texto(forn),
        "cnpj_fornecedor": texto(cnpj),
        "valor_num": numero(val),
        "valor": moeda(val),
        "data": data_br(dt_),
        "situacao": texto(sit),
        "modalidade": texto(mod),
    }

# ============================================================
# PAGINAÇÃO E API HANDLERS
# ============================================================

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
        except (ValueError, TypeError):
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
        data = get_json(url, p)
        return pagina, registros_api(data)

    resultados = {}
    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, limite - 1)) as pool:
        futuros = {pool.submit(buscar, pagina): pagina for pagina in range(2, limite + 1)}
        for f in as_completed(futuros):
            pagina = futuros[f]
            try:
                resultados[pagina] = f.result()[1]
            except Exception as e:
                logging.error(f"Erro na página {pagina}: {e}")
                resultados[pagina] = []
                
    for pagina in range(2, limite + 1):
        todos.extend(resultados.get(pagina, []))
        
    return todos, limite, total

def consultar(url: str, params: Dict[str, Any], max_paginas: int) -> Tuple[List[Dict[str, Any]], int, Optional[int]]:
    serial = tuple(sorted((str(k), str(v)) for k, v in params.items()))
    return consultar_cache(url, serial, max_paginas)

# ============================================================
# GERENCIAMENTO DE STORAGE LOCAL (Parquet IO)
# ============================================================

def garantir_cache():
    CACHE_DIR.mkdir(parents=True, exist_ok=True)

def ler_meta() -> dict:
    garantir_cache()
    if not META_PATH.exists():
        return {}
    try:
        return json.loads(META_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}

def salvar_meta(meta: dict):
    garantir_cache()
    META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

def carregar_historico(tipo: str) -> pd.DataFrame:
    garantir_cache()
    if not HISTORICO_PATH.exists():
        return pd.DataFrame()
    try:
        # Engine pyarrow garantido para evitar falhas de cast de tipos
        df = pd.read_parquet(HISTORICO_PATH, engine="pyarrow")
        if "__tipo_controle" in df.columns:
            df = df[df["__tipo_controle"] == tipo].drop(columns=["__tipo_controle"])
        return df.reset_index(drop=True)
    except Exception as e:
        logging.warning(f"Aviso ao ler parquet histórico: {e}")
        return pd.DataFrame()

def chave_historico(row: pd.Series) -> str:
    for c in ("numeroControlePNCP", "numeroControlePNCPAta", "numeroControlePNCPCompra", "idContratoPNCP"):
        if c in row.index:
            v = texto(row[c], "")
            if v:
                return f"{c}:{v}"
    return "|".join(texto(row.get(c), "") for c in ("numero", "numeroCompra", "numeroContrato", "processo"))

def mesclar_historico(df_antigo: pd.DataFrame, df_novo: pd.DataFrame, tipo: str) -> pd.DataFrame:
    """Mescla dados com segurança, usando engine explícito e gerenciando memória."""
    partes = [x for x in (df_antigo, df_novo) if x is not None and not x.empty]
    if not partes:
        return pd.DataFrame()
        
    out = pd.concat(partes, ignore_index=True, sort=False)
    out["__chave"] = out.apply(chave_historico, axis=1)
    out = out.drop_duplicates("__chave", keep="last").drop(columns=["__chave"])
    out["__tipo_controle"] = tipo
    
    garantir_cache()
    try:
        if HISTORICO_PATH.exists():
            df_completo = pd.read_parquet(HISTORICO_PATH, engine="pyarrow")
            if not df_completo.empty and "__tipo_controle" in df_completo.columns:
                df_outros_tipos = df_completo[df_completo["__tipo_controle"] != tipo]
                df_salvar = pd.concat([df_outros_tipos, out], ignore_index=True, sort=False)
            else:
                df_salvar = out
        else:
            df_salvar = out
            
        # Forçar string nas colunas mistas antes de salvar no Parquet
        df_salvar = df_salvar.astype(str)
        df_salvar.to_parquet(HISTORICO_PATH, index=False, engine="pyarrow", compression="snappy")
    except Exception as e:
        logging.error(f"Erro crítico ao salvar histórico parquet: {e}")
        # Fallback de emergência CSV
        out.to_csv(CACHE_DIR / "contratacoes_controle_interno.csv", index=False, encoding="utf-8-sig")
        
    return out.drop(columns=["__tipo_controle"], errors="ignore").reset_index(drop=True)

def normalizar_pncp(regs: List[Dict[str, Any]]) -> pd.DataFrame:
    if not regs:
        return pd.DataFrame()
    df = pd.json_normalize(regs)
    for col in df.columns:
        if any(isinstance(v, (dict, list)) for v in df[col]):
            df[col] = df[col].astype(str)
    return df

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
    for col in ("cnpjOrgao", "cnpj", "cnpjCompra", "orgaoEntidade.cnpj", "orgao.cnpj"):
        if col in df.columns:
            s = df[col].map(cnpj_limpo)
            m = s == alvo
            if m.any():
                return df.loc[m].reset_index(drop=True)
    return df

# ============================================================
# MOTOR DE RISCO (ESTATÍSTICA)
# ============================================================

def construir_contexto_risco(df: pd.DataFrame, tipo: str) -> pd.DataFrame:
    registros = [dados_registro(row, tipo) for row in df.to_dict('records')]
    return pd.DataFrame(registros)

def limites_iqr(serie: pd.Series) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    s = pd.to_numeric(serie, errors="coerce").dropna()
    if len(s) < 5:
        return None, None, None
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    if iqr <= 0:
        return None, None, float(s.median())
    return float(q1 - 1.5 * iqr), float(q3 + 1.5 * iqr), float(s.median())

def avaliar_modalidade(modalidade: str, objeto: str) -> Tuple[int, List[str], List[str]]:
    pontos, motivos, testes = 0, [], []
    if "dispensa" in modalidade or "inexig" in modalidade:
        pontos += 20
        motivos.append("Modalidade de exceção direta")
        testes.append("Revisar fundamento legal da dispensa/inexigibilidade")
    elif "emerg" in objeto or "emerg" in modalidade:
        pontos += 25
        motivos.append("Indício de contratação emergencial")
        testes.append("Revisar motivação e prazo emergencial")
    return pontos, motivos, testes

def calcular_risco(d: Dict[str, Any], contexto: pd.DataFrame, tipo: str) -> Dict[str, Any]:
    pontos = 0
    motivos = []
    testes = []
    valor = d.get("valor_num")
    modalidade = d.get("modalidade", "").lower()
    objeto = d.get("objeto", "").lower()

    faltantes = []
    for campo, label in (("objeto", "objeto"), ("controle", "controle PNCP"), ("data", "data")):
        if texto(d.get(campo), "") in {"", "N/D"}:
            faltantes.append(label)
            
    if faltantes:
        pontos += min(12, 4 * len(faltantes))
        motivos.append("Campos não identificados: " + ", ".join(faltantes))
        testes.append("Atenção na completude cadastral")

    p_mod, m_mod, t_mod = avaliar_modalidade(modalidade, objeto)
    pontos += p_mod
    motivos.extend(m_mod)
    testes.extend(t_mod)

    if valor is not None:
        if valor >= 500_000:
            pontos += 10
            motivos.append("Valor elevado (>500k)")
        elif valor >= 150_000:
            pontos += 5
            motivos.append("Valor relevante (>150k)")

    baixo, alto, mediana = limites_iqr(contexto.get("valor_num", pd.Series(dtype=float)))
    outlier = False
    if valor is not None and alto is not None and valor > alto:
        outlier = True
        pontos += 25
        motivos.append("Valor anômalo (Acima do limite superior estatístico IQR)")
        testes.append("ALERTA: Avaliar composição de custos frente ao mercado")

    if any(k in objeto for k in ("prorroga", "aditivo", "continuado", "continuidade")):
        pontos += 10
        motivos.append("Contrato continuado/aditivo")
        testes.append("Revisar histórico e vantajosidade da prorrogação")

    fornecedor = texto(d.get("fornecedor"), "N/D")
    if fornecedor != "N/D" and not contexto.empty and "fornecedor" in contexto.columns:
        qtd = int((contexto["fornecedor"].astype(str).str.strip() == fornecedor.strip()).sum())
        if qtd >= max(3, len(contexto) * 0.10):
            pontos += 10
            motivos.append(f"Fornecedor concentrado ({qtd} ocorrências na amostra)")
            testes.append("ALERTA: Avaliar possível dependência/direcionamento")

    pontos = min(100, int(pontos))
    if pontos >= 60:
        nivel = "🔴 ALTO"
    elif pontos >= 30:
        nivel = "🟡 MÉDIO"
    else:
        nivel = "🟢 BAIXO"
        
    if not motivos:
        motivos = ["Sem alertas automáticos."]

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
# NLP (PROCESSAMENTO SEMÂNTICO)
# ============================================================

@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def calcular_modelo_tfidf(textos: Tuple[str, ...]):
    if not SKLEARN_OK or len(textos) < 2:
        return None
    
    vec = TfidfVectorizer(lowercase=True, strip_accents="unicode", ngram_range=(1, 2), min_df=1, max_features=8000, sublinear_tf=True)
    textos_seguros = []
    for t in textos:
        t_str = str(t)
        if len(re.sub(r'[^a-zA-Z0-9]', '', t_str)) > 1:
            textos_seguros.append(t_str)
        else:
            textos_seguros.append("objeto_nao_informado_pelo_orgao")
            
    try:
        return vec.fit_transform(textos_seguros)
    except ValueError as e:
        logging.error(f"Erro no pipeline NLP (TfIdf): {e}")
        return None

def similares(df: pd.DataFrame, idx: int, tipo: str, dados_processados: list = None, limite: int = 5) -> pd.DataFrame:
    if not SKLEARN_OK or len(df) < 2:
        return pd.DataFrame()
        
    if dados_processados:
        objetos = [d["objeto"] for d in dados_processados]
    else:
        objetos = [dados_registro(row, tipo)["objeto"] for row in df.to_dict('records')]
        
    matriz = calcular_modelo_tfidf(tuple(objetos))
    if matriz is None:
        return pd.DataFrame()
        
    scores = (matriz @ matriz[idx].T).toarray().ravel()
    ordem = scores.argsort()[::-1]
    saida = []
    
    for j in ordem:
        if j == idx:
            continue
        d = dados_processados[int(j)] if dados_processados else dados_registro(df.iloc[int(j)].to_dict(), tipo)
        saida.append({
            "Similaridade": f"{float(scores[j])*100:.1f}%", 
            "Número": d["numero"], 
            "Objeto": d["objeto"], 
            "Fornecedor": d["fornecedor"], 
            "Valor": d["valor"]
        })
        if len(saida) >= limite:
            break
    return pd.DataFrame(saida)

# ============================================================
# EXPORTAÇÕES
# ============================================================

def gerar_excel(df: pd.DataFrame, tipo: str, inicio: dt.date, fim: dt.date, df_risco: pd.DataFrame) -> bytes:
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


# ============================================================
# APPLICATION ENTRYPOINT (STREAMLIT APP)
# ============================================================

def main():
    st.set_page_config(
        page_title="Controle Interno — PNCP Rio das Pedras/SP",
        page_icon="🛡️",
        layout="wide",
    )

    defaults = {
        "df": None,
        "tipo": TIPOS[0],
        "paginas": 0,
        "total_paginas_api": 0,
        "inicio": dt.date(2026, 1, 1),
        "fim": dt.date.today(),
        "modo": "⚡ Consulta por período"
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    st.title("🏛️ Painel de Inteligência e Apoio ao Controle Interno")
    st.caption("Monitoramento preventivo de contratações públicas de Rio das Pedras/SP com dados do PNCP.")
    st.info("ℹ️ Os alertas são instrumentos de priorização. Um alerta não significa irregularidade. A conclusão depende da análise do controlador.")

    with st.sidebar:
        st.header("⚙️ Parâmetros de Pesquisa")
        tipo = st.selectbox("Escopo", TIPOS, index=TIPOS.index(st.session_state.tipo))
        
        if tipo != st.session_state.tipo and st.session_state.df is not None:
            st.warning("⚠️ Você alterou o escopo. Clique em **Carregar dados** para atualizar.")

        inicio = st.date_input("📅 Data inicial", st.session_state.inicio)
        fim = st.date_input("📅 Data final", st.session_state.fim)
        max_paginas = st.slider("📄 Limite de páginas da API", 1, 30, MAX_PAGINAS_PADRAO)
        modo = st.radio("Modo de Consulta", ["⚡ Consulta por período", "🔄 Atualização incremental"], index=0)

        meta = ler_meta()
        historico = carregar_historico(tipo)

        if meta.get("ultima_atualizacao"):
            st.caption(f"Última atualização local: {meta['ultima_atualizacao']}")
        else:
            st.caption("Nenhuma atualização incremental registrada.")

        if fim < inicio:
            st.error("Data final anterior à data inicial.")
            st.stop()
            
        buscar_btn = st.button("🔎 Carregar dados", type="primary", use_container_width=True)

    if buscar_btn:
        endpoints = {
            "Contratos": f"{BASE_CONSULTA}/contratos",
            "Atas de Registro de Preços": f"{BASE_CONSULTA}/atas",
            "Editais e Avisos de Contratações": f"{BASE_DOCUMENTOS}/orgaos/{CNPJ}/compras", 
        }
        tamanhos = {"Contratos": PAGE_CONTRATOS, "Atas de Registro de Preços": PAGE_ATAS, "Editais e Avisos de Contratações": PAGE_EDITAIS}
        
        try:
            with st.spinner("Consultando o PNCP... aplicando backoff exponencial via socket para prevenir timeouts."):
                if modo == "🔄 Atualização incremental":
                    ultima = meta.get("ultima_data_consultada")
                    try:
                        base_date = dt.date.fromisoformat(ultima) if ultima else inicio
                    except ValueError:
                        base_date = inicio
                    data_ini = max(inicio, base_date - dt.timedelta(days=1))
                    data_fim = min(fim, dt.date.today())
                    if data_fim < data_ini:
                        data_ini = data_fim
                else:
                    data_ini, data_fim = inicio, fim

                params = {
                    "dataInicial": data_ini.strftime("%Y%m%d"), 
                    "dataFinal": data_fim.strftime("%Y%m%d"), 
                    "tamanhoPagina": tamanhos[tipo], 
                    "pagina": 1
                }
                
                if tipo == "Contratos": 
                    params["cnpjOrgao"] = CNPJ
                elif tipo == "Atas de Registro de Preços": 
                    params["cnpj"] = CNPJ
                
                regs, pags, total_paginas = consultar(endpoints[tipo], params, max_paginas)
                novo = normalizar_pncp(regs)
                novo = deduplicar(novo)
                
                if tipo != "Editais e Avisos de Contratações":
                    novo = filtrar_cnpj(novo)

                combinado = mesclar_historico(historico, novo, tipo)
                df = combinado.copy()
                
                col_data = next((c for c in ("dataPublicacaoPncp", "dataPublicacao", "dataInclusao", "dataAssinatura", "dataCelebracao") if c in df.columns), None)
                if col_data and not df.empty:
                    ds = pd.to_datetime(df[col_data], errors="coerce").dt.date
                    df = df[(ds >= inicio) & (ds <= fim)].reset_index(drop=True)

                meta["ultima_atualizacao"] = dt.datetime.now().isoformat(timespec="seconds")
                meta["ultima_data_consultada"] = data_fim.isoformat()
                meta["tipo"] = tipo
                salvar_meta(meta)
                
                st.session_state.df = df
                st.session_state.tipo = tipo
                st.session_state.paginas = pags
                st.session_state.total_paginas_api = total_paginas
                st.session_state.inicio = inicio
                st.session_state.fim = fim
                st.session_state.modo = modo
                st.success(f"Consulta concluída com sucesso: {len(df)} registro(s) obtidos.")
        
        except Exception as e:
            st.error(f"❌ Erro na consulta de rede: {e}")

    # ============================================================
    # VISUALIZAÇÃO DE DADOS (FRONTEND)
    # ============================================================
    
    df = st.session_state.get("df")
    tipo_atual = st.session_state.get("tipo", tipo)

    if df is None:
        st.info("👈 Defina os parâmetros na barra lateral e clique em **Carregar dados**.")
        st.stop()
    if df.empty:
        st.warning("Nenhum dado retornado para o filtro aplicado.")
        st.stop()

    df_records = df.to_dict('records')
    dados_processados = [dados_registro(row, tipo_atual) for row in df_records]
    contexto = pd.DataFrame(dados_processados)
    riscos = [calcular_risco(d, contexto, tipo_atual) for d in dados_processados]

    altos = sum(r["nivel"] == "🔴 ALTO" for r in riscos)
    medios = sum(r["nivel"] == "🟡 MÉDIO" for r in riscos)
    outliers = sum(r["outlier"] for r in riscos)
    valor_total = contexto["valor_num"].sum(min_count=1) if "valor_num" in contexto else None

    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Registros Consolidados", len(df))
    kpi2.metric("🔴 Alto risco", altos)
    kpi3.metric("🟡 Médio risco", medios)
    kpi4.metric("🚨 Outliers Financeiros", outliers)
    kpi5.metric("Valor Total Analisado", moeda(valor_total))

    st.markdown("---")
    
    tab1, tab2 = st.tabs(["🚦 Matriz de Risco", "📋 Auditoria Detalhada"])

    with tab1:
        st.subheader("Visão Consolidada")
        rows = []
        for i, r in enumerate(riscos):
            d = dados_processados[i]
            rows.append({
                "Índice": i, 
                "Risco": r["nivel"], 
                "Score": r["pontos"], 
                "Número": d["numero"], 
                "Modalidade": d["modalidade"], 
                "Fornecedor": d["fornecedor"], 
                "Valor": d["valor"], 
                "Objeto": d["objeto"], 
                "Gatilhos": "; ".join(r["motivos"])
            })
        df_risco = pd.DataFrame(rows).sort_values(["Score", "Índice"], ascending=[False, True])
        st.dataframe(df_risco.drop(columns=["Índice"]), use_container_width=True, hide_index=True)

        st.subheader("📤 Exportação")
        excel = gerar_excel(df, tipo_atual, st.session_state.get("inicio", inicio), st.session_state.get("fim", fim), df_risco)
        st.download_button(
            "📊 Baixar Relatório (Excel)", 
            excel, 
            file_name=f"Relatorio_Auditoria_{slug(tipo_atual)}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with tab2:
        st.subheader("Investigação e Triagem")
        f1, f2 = st.columns(2)
        filtro_risco = f1.multiselect("Nível de Risco Alvo", ["🔴 ALTO", "🟡 MÉDIO", "🟢 BAIXO"], default=["🔴 ALTO", "🟡 MÉDIO"])
        mostrar_outlier = f2.checkbox("Somente anomalias financeiras de IQR", False)

        indices = []
        for i, r in enumerate(riscos):
            if r["nivel"] not in filtro_risco: continue
            if mostrar_outlier and not r["outlier"]: continue
            indices.append(i)

        st.caption(f"{len(indices)} registro(s) qualificado(s) pelo filtro.")

        if not indices:
            st.warning("Ajuste os filtros de priorização para carregar registros.")
        else:
            opcoes = []
            mapa = {}
            for i in indices:
                d = dados_processados[i]
                label = f"[{riscos[i]['nivel']} Score {riscos[i]['pontos']}] {d['numero']} — {d['fornecedor']} — {d['valor']}"
                opcoes.append(label)
                mapa[label] = i
                
            escolhido = st.selectbox("Selecione o dossiê da contratação:", opcoes)
            i = mapa[escolhido]
            d = dados_processados[i]
            r = riscos[i]

            with st.container():
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("Score Calculado", f"{r['pontos']} pts")
                col_b.metric("Volume Financeiro", d["valor"])
                col_c.metric("Enquadramento", d["modalidade"])
                st.info(f"**Escopo da Contratação:** {d['objeto']}")

            with st.expander("🔎 Matriz de Justificativa Analítica", expanded=True):
                st.markdown("**Variáveis de Acionamento (Features):**")
                for m in r["motivos"]: 
                    st.write(f"- {m}")
                st.markdown("**Recomendação de Auditoria (Ações):**")
                for t in r["testes"]: 
                    st.write(f"- {t}")

            st.markdown("### 🤖 Clusterização Semântica")
            if not SKLEARN_OK:
                st.warning("O pacote `scikit-learn` não está instalado no ambiente.")
            elif len(df) < 2:
                st.info("Insuficiência de dados (N < 2) para modelo de vizinhos mais próximos (KNN/Cosseno).")
            else:
                sim = similares(df, i, tipo_atual, dados_processados)
                if sim.empty:
                    st.info("O modelo NLP não encontrou vizinhança semântica acima da nota de corte.")
                else:
                    st.dataframe(sim, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
